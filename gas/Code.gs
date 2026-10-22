/**
 * リヴァイ：Googleカレンダーの予定をアプリに渡し、アプリから入れた予定を登録する
 * 予定を渡すときに、やる日のない締切（3週間以内）を逆算して「▶ 〜」の作業日を置く
 *
 * 使い方
 * 1. この内容をApps Scriptに貼り、setup() を1回実行する（ログに合言葉が出る）
 * 2. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *    実行するユーザー：自分　／　アクセスできるユーザー：全員
 * 3. 出てきたURLと合言葉を、アプリの設定に入れる
 *
 * 合言葉が合わないアクセスには何も返さない。
 * コードを変えたら：デプロイ → デプロイを管理 → 鉛筆 → バージョン「新バージョン」→ デプロイ（URLは変わらない）
 */

// カテゴリのルール（google-calendar-categorize と同じ）
const CATS = {
  family: { prefix: '【家族】', color: '10' }, // Basil
  work:   { prefix: '【仕事】', color: '7' },  // Peacock
  union:  { prefix: '【組合】', color: '11' }  // Tomato
};

const DAYS_AHEAD = 190; // 月の暦のために半年先（約6か月）まで読む

function doGet(e) {
  const key = PropertiesService.getScriptProperties().getProperty('LEVI_KEY');
  const given = e && e.parameter ? e.parameter.k : '';
  if (!key || given !== key) return json_({ error: 'unauthorized' });

  const now = new Date();
  const start = new Date(now);
  start.setHours(0, 0, 0, 0);
  const end = new Date(start.getTime() + DAYS_AHEAD * 24 * 60 * 60 * 1000);

  const cal = CalendarApp.getDefaultCalendar();
  const raw = cal.getEvents(start, end);
  let events = raw.map(toObj_);

  // 締切の逆算：まだやる日のない締切に「▶ 〜」を置く（置いたものは返す予定にも足す）
  let planned = [], planError = '';
  try {
    planned = autoPlan_(cal, raw, events, now);
    events = events.concat(planned);
  } catch (err) { planError = String(err); }

  return json_({ now: now.toISOString(), events: events, planned: planned.length, planError: planError });
}

function toObj_(ev) {
  return {
    id: ev.getId(),
    title: ev.getTitle(),
    start: ev.getStartTime().toISOString(),
    end: ev.getEndTime().toISOString(),
    allDay: ev.isAllDayEvent(),
    color: String(ev.getColor() || ''),
    location: ev.getLocation() || '',
    desc: (ev.getDescription() || '').slice(0, 500)
  };
}

/**
 * アプリからの登録と取り消し
 * add:    { k, action:'add', items:[{ title, cat, allDay, date:'YYYY-MM-DD' | start, end (ISO), desc }] }
 * update: { k, action:'update', item:{ id, title, cat, allDay, date | start, end } }
 * delete: { k, action:'delete', ids:[...], any }  ※anyがなければリヴァイが入れた予定だけ消す
 */
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ error: 'bad request' }); }
  const key = PropertiesService.getScriptProperties().getProperty('LEVI_KEY');
  if (!key || body.k !== key) return json_({ error: 'unauthorized' });
  const cal = CalendarApp.getDefaultCalendar();

  if (body.action === 'add') {
    const ids = (body.items || []).map(function (it) {
      const c = CATS[it.cat];
      const title = titled_(it);
      const opts = { description: (it.desc ? it.desc + '\n' : '') + '（リヴァイが登録）' };
      let ev;
      if (it.allDay) {
        const p = String(it.date).split('-').map(Number);
        ev = cal.createAllDayEvent(title, new Date(p[0], p[1] - 1, p[2]), opts);
      } else {
        ev = cal.createEvent(title, new Date(it.start), new Date(it.end), opts);
      }
      if (c) ev.setColor(c.color);
      return ev.getId();
    });
    return json_({ ok: true, ids: ids });
  }

  if (body.action === 'update') {
    const it = body.item || {};
    const ev = cal.getEventById(it.id);
    if (!ev) return json_({ error: '予定が見つかりません' });
    ev.setTitle(titled_(it));
    if (it.allDay) {
      const p = String(it.date).split('-').map(Number);
      ev.setAllDayDate(new Date(p[0], p[1] - 1, p[2]));
    } else {
      ev.setTime(new Date(it.start), new Date(it.end));
    }
    if (CATS[it.cat]) ev.setColor(CATS[it.cat].color);
    return json_({ ok: true, id: ev.getId() });
  }

  if (body.action === 'delete') {
    let n = 0;
    (body.ids || []).forEach(function (id) {
      const ev = cal.getEventById(id);
      // any: 長押しから本人が消すとき。それ以外はリヴァイが入れた予定だけ
      if (ev && (body.any || ev.getDescription().indexOf('（リヴァイが登録）') !== -1)) { ev.deleteEvent(); n++; }
    });
    return json_({ ok: true, deleted: n });
  }

  return json_({ error: 'unknown action' });
}

// ---------- 締切の逆算 ----------
// 締切の何日前から逆算するか（それより先の締切は、近づいてから置く）
const PLAN_DAYS = 21;
const STRONG_DL = /締切|締め切り|〆切|しめきり|期限|〆/;
const REST_RE = /休業|休み|年休|振替|休日/;
const PLANNED_MARK = '（逆算済み）'; // 置いたら締切の説明に付ける。▶を消しても置き直さない
const NO_PLAN = '（逆算しない）';    // 締切の説明にこれを書けば逆算しない
// 作業してよい時間帯（分）。仕事・組合は平日の放課後、家族は平日夜か休日の朝
const WINDOWS = {
  work:    { weekday: [15 * 60 + 30, 17 * 60 + 30], holiday: null },
  family:  { weekday: [20 * 60, 21 * 60],           holiday: [9 * 60, 11 * 60] }
};
const STAGES = {
  1: ['仕上げて出せる形に'],
  2: ['とりかかって半分まで', '仕上げて出せる形に'],
  3: ['とりかかる（材料を集めて全体を見る）', '半分まで進める', '仕上げて出せる形に']
};
// 祝日（カレンダーから読めないときの予備）
const HOLIDAYS = ['2026-11-03', '2026-11-23', '2027-01-01', '2027-01-11', '2027-02-11', '2027-02-23', '2027-03-22',
  '2027-04-29', '2027-05-03', '2027-05-04', '2027-05-05', '2027-07-19', '2027-08-11', '2027-09-20', '2027-09-23',
  '2027-10-11', '2027-11-03', '2027-11-23'];

/** 計画を立ててカレンダーに書く。置いた予定（アプリに返す形）を返す */
function autoPlan_(cal, raw, events, now) {
  const plan = planDeadlines_(events, holidays_(now), now);
  if (!plan.create.length && !plan.mark.length) return [];
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) return []; // ほかの端末が置いている最中
  try {
    const byId = {};
    raw.forEach(function (ev) { byId[ev.getId()] = ev; });
    // ロック待ちの間に別の端末が置いていたら、何もしない
    const fresh = plan.mark.filter(function (id) { return byId[id] && byId[id].getDescription().indexOf(PLANNED_MARK) === -1; });
    if (!fresh.length) return [];
    const out = [];
    plan.create.forEach(function (it) {
      if (fresh.indexOf(it.deadlineId) === -1) return;
      const ev = cal.createEvent(it.title, it.start, it.end, { description: it.desc });
      if (it.color) ev.setColor(it.color);
      out.push(toObj_(ev));
    });
    fresh.forEach(function (id) {
      const ev = byId[id], d = ev.getDescription() || '';
      ev.setDescription((d ? d + '\n' : '') + PLANNED_MARK);
    });
    return out;
  } finally { lock.releaseLock(); }
}

/**
 * 逆算だけを行う（カレンダーには触らない）。evs はアプリに返す形の予定
 * 戻り値：create（置く予定）, mark（逆算済みにする締切のid）
 */
function planDeadlines_(evs, holidays, now) {
  const today = sod_(now), out = { create: [], mark: [] };
  evs.forEach(function (dl) {
    const t = stripCat_(dl.title), desc = dl.desc || '';
    if (t.indexOf('▶') === 0 || !STRONG_DL.test(t)) return;
    if (desc.indexOf(PLANNED_MARK) !== -1 || desc.indexOf(NO_PLAN) !== -1) return;
    const D = sod_(new Date(dl.start)), days = Math.round((D - today) / 864e5);
    if (days < 1 || days > PLAN_DAYS) return;
    // すでに作業日が置かれている締切（手で置いたもの・＋から選んだもの）は印だけ付ける
    if (/作業日|逆算|計画/.test(desc) || hasPlan_(dl, evs)) { out.mark.push(dl.id); return; }

    const cat = catOf_(dl), mins = workMins_(desc), want = days >= 8 ? 3 : days >= 3 ? 2 : 1;
    const finish = addD_(D, days >= 3 ? -2 : -1);
    const first = new Date(Math.max(today, addD_(finish, -10)));
    const mid = new Date(first.getTime() + Math.round((finish - first) / 864e5 / 2) * 864e5);
    // 探す順番：仕上げは2日前→前日→さらに前、とりかかりは早い日から、真ん中は近い日から
    const order = [
      [0, 1, 2, 3].map(function (n) { return addD_(first, n); }),
      [0, 1, -1, 2, -2].map(function (n) { return addD_(mid, n); }),
      [0, 1, -1, -2, -3].map(function (n) { return addD_(finish, n); })
    ];
    const targets = want === 3 ? order : want === 2 ? [order[0], order[2]] : [order[2]];
    const used = {}, slots = [];
    targets.forEach(function (cands) {
      for (let i = 0; i < cands.length; i++) {
        const d = cands[i], k = ymd_(d);
        if (d < today || d >= D || used[k]) continue;
        const s = slotOn_(d, cat, mins, evs, holidays, now);
        if (s) { used[k] = true; slots.push(s); return; }
      }
    });
    if (!slots.length) return; // 空きがない。次に開いたときにもう一度考える
    slots.sort(function (a, b) { return a.start - b.start; });
    const stages = STAGES[slots.length], md = (D.getMonth() + 1) + '/' + D.getDate();
    const prefix = (dl.title.match(/^【(家族|仕事|組合)】/) || [''])[0];
    slots.forEach(function (s, i) {
      out.create.push({
        deadlineId: dl.id,
        title: prefix + '▶ ' + baseTitle_(t),
        start: s.start, end: s.end, color: dl.color || '',
        desc: 'ここまで：' + stages[i] + '（' + (i + 1) + '/' + slots.length + '）\n締切' + md + 'の逆算（' + mins + '分）\n（リヴァイが登録）'
      });
    });
    out.mark.push(dl.id);
  });
  return out;
}

/** その日の空き時間から、作業時間ぶんの枠を探す */
function slotOn_(day, cat, mins, evs, holidays, now) {
  const k = ymd_(day), dow = day.getDay();
  const rest = evs.some(function (e) { return e.allDay && REST_RE.test(e.title) && sod_(new Date(e.start)) <= day && new Date(e.end) > day; });
  const off = dow === 0 || dow === 6 || holidays.indexOf(k) !== -1 || rest;
  const w = (WINDOWS[cat] || WINDOWS.work)[off ? 'holiday' : 'weekday'];
  if (!w) return null; // 仕事は休みの日に置かない
  let from = new Date(day.getTime() + w[0] * 6e4);
  const until = new Date(day.getTime() + w[1] * 6e4);
  if (from < now) { from = new Date(now); from.setMinutes(Math.ceil(from.getMinutes() / 15) * 15, 0, 0); }
  const busy = evs.filter(function (e) { return !e.allDay && new Date(e.end) > from && new Date(e.start) < until; })
    .sort(function (a, b) { return new Date(a.start) - new Date(b.start); });
  for (let i = 0; i < busy.length; i++) {
    const bs = new Date(busy[i].start), be = new Date(busy[i].end);
    if ((bs - from) / 6e4 >= mins) break;
    if (be > from) from = be;
  }
  if ((until - from) / 6e4 < mins) return null;
  return { start: from, end: new Date(from.getTime() + mins * 6e4) };
}

/** 締切の前に、その締切のための作業がもう置かれているか */
function hasPlan_(dl, evs) {
  const D = sod_(new Date(dl.start)), md = (D.getMonth() + 1) + '/' + D.getDate();
  const want = '▶ ' + baseTitle_(stripCat_(dl.title));
  return evs.some(function (e) {
    if (e.id === dl.id || new Date(e.start) >= addD_(D, 1)) return false;
    const d = e.desc || '';
    return stripCat_(e.title).indexOf(want) === 0 || d.indexOf('締切' + md) !== -1 || d.indexOf(md + '締切') !== -1;
  });
}

function catOf_(e) {
  const m = e.title.match(/^【(家族|仕事|組合)】/);
  if (m) return { 家族: 'family', 仕事: 'work', 組合: 'union' }[m[1]];
  return { '10': 'family', '7': 'work', '11': 'union' }[e.color] || 'other';
}
function stripCat_(t) { return String(t || '').replace(/^【(家族|仕事|組合)】\s*/, ''); }
function baseTitle_(t) {
  return t.replace(/^(締切|〆切|期限)[:：]\s*|^〆\s*/, '').replace(/\s*(締切|締め切り|〆切|しめきり|期限)$/, '').trim();
}
/** 締切の説明に「作業60分」「1時間」などがあればその長さ、なければ30分 */
function workMins_(desc) {
  const m = String(desc).match(/作業\s*(\d+)\s*(分|時間)/);
  return m ? +m[1] * (m[2] === '分' ? 1 : 60) : 30;
}
function holidays_(now) {
  const list = HOLIDAYS.slice();
  try {
    const hc = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
    if (hc) hc.getEvents(sod_(now), addD_(sod_(now), PLAN_DAYS + 1)).forEach(function (e) { list.push(ymd_(e.getStartTime())); });
  } catch (err) {}
  return list;
}
function sod_(d) { const x = new Date(d); x.setHours(0, 0, 0, 0); return x; }
function addD_(d, n) { const x = sod_(d); x.setDate(x.getDate() + n); return x; }
function ymd_(d) { return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }

/** タイトルの先頭に【家族】などを付け直す */
function titled_(it) {
  const t = String(it.title || '').trim().replace(/^【(家族|仕事|組合)】/, '');
  return (CATS[it.cat] ? CATS[it.cat].prefix : '') + t;
}

/** 最初に1回だけ実行する。合言葉を作ってログに出す */
function setup() {
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('LEVI_KEY');
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '');
    props.setProperty('LEVI_KEY', key);
  }
  CalendarApp.getDefaultCalendar(); // カレンダーの許可を求めるため
  Logger.log('合言葉: ' + key);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
