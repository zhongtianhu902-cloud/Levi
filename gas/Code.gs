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
  if (e.parameter.a) return json_(pushGet_(e.parameter.a)); // 通知の鍵・受け箱

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
 * push:   { k, action:'push', sub:{endpoint}, off, name, cfg:{morning,noon,evening,lead,leadFamily}, test }
 */
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ error: 'bad request' }); }
  const key = PropertiesService.getScriptProperties().getProperty('LEVI_KEY');
  if (!key || body.k !== key) return json_({ error: 'unauthorized' });
  if (body.action === 'push') return json_(pushPost_(body)); // 通知の端末登録・時刻・ためし送信
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

// ---------- 声かけ（通知） ----------
// アプリ（ホーム画面のWebアプリ）に Web Push で届ける。押すと A「いま」が開く
// 送るのは「中身なし」の合図だけ。合図を受けたアプリが a=inbox で文面を取りに来る（暗号化が要らない）
// 使い方：コードを貼ったら setupPush() を1回実行する（鍵を作り、5分ごとの見回りを置く）
const PUSH_DEFAULT = { morning: '06:30', noon: '12:30', evening: '16:30', lead: 10, leadFamily: 15 };
const NO_BEFORE = /\d限|校時/;      // 授業の前は知らせない（毎時間鳴るため）
const WRAP_MIN = 5;                // 何分前に切り上げるか（アプリと同じ）
const PUSH_SUB = 'https://zhongtianhu902-cloud.github.io/Levi/';
const OUT_RE_ = /迎え|送り|運動会|遠足|えんそく|サッカー|スイミング|試合|校外|見学|プール|公園|散歩|外出|出張|会館|センター|いもほり|ハロウィン|子ども会|買い物|病院|帰省|旅行|訪問/;

/** 最初に1回だけ実行する。通知の鍵を作り、5分ごとの見回り（tick）を置く */
function setupPush() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('PUSH_PRIV')) {
    const k = p256NewKey_();
    props.setProperties({ PUSH_PRIV: k.priv, PUSH_PUB: k.pub });
  }
  es256_(utf8_('check'), props.getProperty('PUSH_PRIV')); // 署名できるか確かめる（できなければここで止まる）
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'tick') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('tick').timeBased().everyMinutes(5).create();
  UrlFetchApp.fetch('https://api.open-meteo.com/v1/forecast?latitude=33.24&longitude=131.61&hourly=temperature_2m&forecast_days=1'); // 外へつなぐ許可を求めるため
  Logger.log('通知の準備ができました。アプリの⚙︎から「通知をオンにする」を押してください');
}

function pushCfg_() {
  const c = JSON.parse(PropertiesService.getScriptProperties().getProperty('PUSH_CFG') || '{}');
  const out = {};
  Object.keys(PUSH_DEFAULT).forEach(function (k) { out[k] = c[k] != null && c[k] !== '' ? c[k] : PUSH_DEFAULT[k]; });
  return out;
}
function subs_() { return JSON.parse(PropertiesService.getScriptProperties().getProperty('PUSH_SUBS') || '[]'); }
function saveSubs_(list) { PropertiesService.getScriptProperties().setProperty('PUSH_SUBS', JSON.stringify(list)); }

/** doGet / doPost から呼ぶ（a=vapid, a=inbox / action=push） */
function pushGet_(a) {
  const props = PropertiesService.getScriptProperties();
  if (a === 'vapid') return { pub: props.getProperty('PUSH_PUB') || '', cfg: pushCfg_(), devices: subs_().length };
  if (a === 'inbox') return { inbox: JSON.parse(props.getProperty('PUSH_INBOX') || '[]') };
  return { error: 'unknown' };
}
function pushPost_(body) {
  const props = PropertiesService.getScriptProperties();
  if (body.sub && body.sub.endpoint) {
    const list = subs_().filter(function (s) { return s.endpoint !== body.sub.endpoint; });
    if (!body.off) list.push({ endpoint: body.sub.endpoint, name: body.name || '', at: Date.now() });
    saveSubs_(list.slice(-5));
  }
  if (body.cfg) {
    const c = {};
    Object.keys(PUSH_DEFAULT).forEach(function (k) { if (body.cfg[k] != null) c[k] = body.cfg[k]; });
    props.setProperty('PUSH_CFG', JSON.stringify(c));
  }
  if (body.test) {
    const now = new Date(), evs = dayEvents_(now);
    const n = beforeNudge_(evs.filter(function (e) { return !e.allDay && new Date(e.start) > now; })[0], evs, now, pushCfg_())
      || { title: 'リヴァイ', body: '通知はこう届きます。押すと「いま」が開きます' };
    deliver_([{ key: 'test' + now.getTime(), title: n.title, body: n.body }]);
  }
  return { ok: true, devices: subs_().length, cfg: pushCfg_() };
}

/** 5分ごとに動く。この5分のあいだに来た声かけを送る */
function tick() {
  if (!subs_().length) return;
  const props = PropertiesService.getScriptProperties(), now = new Date();
  let last = +props.getProperty('PUSH_LAST') || now.getTime() - 5 * 6e4;
  if (now - last > 30 * 6e4) last = now.getTime() - 6 * 6e4; // 止まっていた間の分はまとめて送らない
  props.setProperty('PUSH_LAST', String(now.getTime()));
  const evs = dayEvents_(now);
  const due = nudgesBetween_(evs, new Date(last), now, pushCfg_(), function () { return wxOita_(); });
  // 同じ声かけを2回送らない（今日の分だけ覚える）
  const sent = JSON.parse(props.getProperty('PUSH_SENT') || '{}'), today = ymd_(now);
  const keys = sent.day === today ? sent.keys : [];
  const fresh = due.filter(function (n) { return keys.indexOf(n.key) === -1; });
  if (!fresh.length) return;
  props.setProperty('PUSH_SENT', JSON.stringify({ day: today, keys: keys.concat(fresh.map(function (n) { return n.key; })).slice(-80) }));
  deliver_(fresh);
}

/** 今日と明日の予定（アプリに渡す形＋cat） */
function dayEvents_(now) {
  return CalendarApp.getDefaultCalendar().getEvents(sod_(now), addD_(now, 2)).map(function (ev) {
    const o = toObj_(ev); o.cat = catOf_(o); o.name = shortTitle_(o.title); return o;
  });
}
function shortTitle_(t) { return stripCat_(t).replace(/^▶\s*/, '').replace(/[（(][^）)]*[）)]$/, '').trim(); }

/** 文面を受け箱に入れてから、各端末に合図を送る */
function deliver_(list) {
  const props = PropertiesService.getScriptProperties();
  const inbox = JSON.parse(props.getProperty('PUSH_INBOX') || '[]');
  let id = Date.now();
  list.forEach(function (n) { inbox.push({ id: id++, at: Date.now(), title: n.title, body: n.body, tag: n.key }); });
  props.setProperty('PUSH_INBOX', JSON.stringify(inbox.slice(-10)));
  const priv = props.getProperty('PUSH_PRIV'), pub = props.getProperty('PUSH_PUB');
  if (!priv) return;
  const keep = [];
  subs_().forEach(function (s) {
    const aud = s.endpoint.match(/^https:\/\/[^/]+/)[0];
    const jwt = vapidJwt_(aud, priv);
    let code = 0;
    try {
      code = UrlFetchApp.fetch(s.endpoint, {
        method: 'post', muteHttpExceptions: true, payload: '',
        headers: { TTL: '1800', Urgency: 'high', Authorization: 'vapid t=' + jwt + ', k=' + pub }
      }).getResponseCode();
    } catch (err) { code = 0; }
    if (code !== 404 && code !== 410) keep.push(s); // 404/410 はもう届かない端末
  });
  saveSubs_(keep);
}

/** from〜to のあいだに来る声かけ。wx は天気が要るときだけ呼ぶ */
function nudgesBetween_(evs, from, to, cfg, wx) {
  const out = [], day = sod_(to), at = function (hm) { const p = String(hm).split(':'); return new Date(day.getTime() + (+p[0] * 60 + +p[1]) * 6e4); };
  const inWin = function (t) { return t > from && t <= to; };
  if (inWin(at(cfg.morning))) out.push(morningNudge_(evs, to, wx()));
  if (inWin(at(cfg.noon))) out.push(noonNudge_(evs, to));
  if (inWin(at(cfg.evening))) out.push(eveningNudge_(evs, to));
  evs.forEach(function (e) {
    if (e.allDay || NO_BEFORE.test(e.title)) return;
    const s = new Date(e.start), lead = e.cat === 'family' ? +cfg.leadFamily : +cfg.lead;
    if (s.getHours() < 5 || s.getHours() >= 22) return;
    if (inWin(new Date(s.getTime() - lead * 6e4))) {
      const n = beforeNudge_(e, evs, to, cfg, OUT_RE_.test(e.title) || e.cat === 'family' ? wx : null);
      if (n) out.push(n);
    }
  });
  return out.filter(Boolean);
}

const hm_ = function (d) { d = new Date(d); return d.getHours() + ':' + ('0' + d.getMinutes()).slice(-2); };
const isDo_ = function (e) { return stripCat_(e.title).indexOf('▶') === 0 || !!stage_(e) || /締切\d+\/\d+の逆算|\d+\/\d+締切/.test(e.desc || ''); };
function stage_(e) { const m = (e.desc || '').match(/ここまで[」』]?[:：]\s*([^\n。]+?)(?:（\d+\/\d+）)?(?:\n|。|$)/); return m ? m[1] : ''; }
function todays_(evs, now) {
  return evs.filter(function (e) { return sameDay_(new Date(e.start), now); })
    .sort(function (a, b) { return new Date(a.start) - new Date(b.start); });
}
function sameDay_(a, b) { return ymd_(a) === ymd_(b); }
function isDl_(e) { return !isDo_(e) && /締切|締め切り|〆切|しめきり|期限|提出|〆/.test(e.title); }
/** 近い締切（3日以内）を一言で */
function dlLine_(evs, now) {
  const near = evs.filter(function (e) { return isDl_(e) && e.allDay && new Date(e.end) > now; })
    .sort(function (a, b) { return new Date(a.start) - new Date(b.start); })[0];
  if (!near) return '';
  const d = Math.round((sod_(new Date(near.start)) - sod_(now)) / 864e5);
  return '締切：' + baseTitle_(near.name) + '（' + (d <= 0 ? '今日' : d === 1 ? '明日' : d + '日後') + '）';
}

function morningNudge_(evs, now, wx) {
  const t = todays_(evs, now), timed = t.filter(function (e) { return !e.allDay; });
  const lines = [];
  const fam = timed.filter(function (e) { return e.cat === 'family'; });
  const first = timed.filter(function (e) { return e.cat !== 'family' && !isDo_(e); })[0];
  if (first) lines.push('最初：' + hm_(first.start) + ' ' + first.name);
  fam.forEach(function (e) { lines.push('家族：' + hm_(e.start) + ' ' + e.name); });
  t.filter(function (e) { return e.allDay && !isDl_(e); }).slice(0, 2).forEach(function (e) { lines.push('今日：' + e.name); });
  const d = t.filter(isDo_)[0];
  if (d) lines.push('やる：' + hm_(d.start) + '〜 ' + d.name + (stage_(d) ? '（' + stage_(d) + 'まで）' : ''));
  const dl = dlLine_(evs, now); if (dl) lines.push(dl);
  (weatherLines_(t, now, wx) || []).forEach(function (w) { lines.push(w); });
  return { key: 'morning', title: 'おはよう。今日は' + (timed.length ? timed.length + '件' : '決まった予定なし'), body: lines.join('\n') || '自分のための時間にしてよい日' };
}

function noonNudge_(evs, now) {
  const rest = todays_(evs, now).filter(function (e) { return !e.allDay && new Date(e.start) > now; });
  const lines = [];
  const d = rest.filter(isDo_)[0];
  if (d) lines.push('午後のやる：' + hm_(d.start) + '〜 ' + d.name + '。とりあえず ' + (stage_(d) || '手をつける') + ' まで');
  const nx = rest.filter(function (e) { return !isDo_(e) && !NO_BEFORE.test(e.title); })[0];
  if (nx) lines.push('次の約束：' + hm_(nx.start) + ' ' + nx.name);
  const fam = rest.filter(function (e) { return e.cat === 'family'; })[0];
  if (fam) lines.push(hm_(fam.start) + ' ' + fam.name + 'があるので、仕事は' + hm_(new Date(new Date(fam.start) - WRAP_MIN * 6e4)) + 'までに切り上げ');
  const dl = dlLine_(evs, now); if (dl) lines.push(dl);
  return { key: 'noon', title: '昼のひと押し', body: lines.join('\n') || '午後は決まった予定なし。締切の少し先を進めるなら今' };
}

function eveningNudge_(evs, now) {
  const rest = todays_(evs, now).filter(function (e) { return !e.allDay && new Date(e.start) > now; });
  const lines = [];
  const fam = rest.filter(function (e) { return e.cat === 'family'; })[0];
  if (fam) lines.push(hm_(fam.start) + ' ' + fam.name + '。' + hm_(new Date(new Date(fam.start) - WRAP_MIN * 6e4)) + 'までに切り上げ');
  const d = rest.filter(isDo_)[0];
  if (d) lines.push(hm_(d.start) + '〜 ' + d.name + '。' + (stage_(d) || '手をつける') + ' まででよい');
  const tm = addD_(now, 1);
  const first = todays_(evs, tm).filter(function (e) { return !e.allDay && !isDo_(e); })[0];
  if (first) lines.push('明日の最初：' + hm_(first.start) + ' ' + first.name);
  const dl = dlLine_(evs, now); if (dl) lines.push(dl);
  return { key: 'evening', title: fam ? 'そろそろ切り上げ' : '夕方のひと押し', body: lines.join('\n') || '今日はここまでで十分' };
}

/** 予定の直前のひと押し。「18:10 スイミング迎え。18:05までに切り上げ、今の作業はここで区切る」 */
function beforeNudge_(e, evs, now, cfg, wx) {
  if (!e) return null;
  const s = new Date(e.start), left = Math.max(0, Math.round((s - now) / 6e4));
  if (isDo_(e)) {
    return { key: 'before:' + e.id + ':' + s.getTime(), title: hm_(s) + 'から「' + e.name + '」', body: 'とりあえずここまで：' + (stage_(e) || '手をつける') + '\n' + Math.round((new Date(e.end) - s) / 6e4) + '分。終わらなくても区切ってよい' };
  }
  const lines = [hm_(new Date(s - WRAP_MIN * 6e4)) + 'までに切り上げ'];
  const cur = evs.filter(function (x) { return !x.allDay && x.id !== e.id && new Date(x.start) <= now && new Date(x.end) > now; })[0];
  if (cur && isDo_(cur)) lines[0] += '。「' + cur.name + '」' + (stage_(cur) ? 'は ' + stage_(cur) + ' まで' : 'はここで区切る');
  if (e.location) lines.push('場所：' + e.location);
  const rain = wx ? rainAt_(wx(), s) : null;
  if (rain) lines.push('雨' + rain + '%。' + (e.cat === 'family' ? '傘を車に' : '傘を持って'));
  return { key: 'before:' + e.id + ':' + s.getTime(), title: hm_(s) + ' ' + e.name + ' まであと' + left + '分', body: lines.join('\n') };
}

// ---------- 天気（大分市、Open-Meteo）。予定にかかわるときだけ ----------
function wxOita_() {
  const cache = CacheService.getScriptCache(), c = cache.get('wx');
  if (c) return JSON.parse(c);
  try {
    const r = UrlFetchApp.fetch('https://api.open-meteo.com/v1/forecast?latitude=33.2396&longitude=131.6093'
      + '&hourly=precipitation_probability,precipitation&daily=temperature_2m_max,temperature_2m_min&timezone=Asia%2FTokyo&forecast_days=2', { muteHttpExceptions: true });
    const d = JSON.parse(r.getContentText());
    if (!d.hourly) return null;
    cache.put('wx', JSON.stringify(d), 1800);
    return d;
  } catch (err) { return null; }
}
/** その時刻から2時間の降水確率（50%以上のときだけ返す） */
function rainAt_(wx, t) {
  if (!wx || !wx.hourly) return 0;
  let pop = 0;
  for (let h = 0; h < 2; h++) {
    const x = new Date(t.getTime() + h * 36e5);
    const i = wx.hourly.time.indexOf(ymd_(x) + 'T' + ('0' + x.getHours()).slice(-2) + ':00');
    if (i >= 0) pop = Math.max(pop, wx.hourly.precipitation_probability[i] || 0, (wx.hourly.precipitation[i] || 0) >= 1 ? 60 : 0);
  }
  return pop >= 50 ? pop : 0;
}
function weatherLines_(today, now, wx) {
  const out = [];
  if (!wx || !wx.daily) return out;
  today.filter(function (e) { return e.cat === 'family' || OUT_RE_.test(e.title) || e.location; }).forEach(function (e) {
    const s = e.allDay ? new Date(sod_(now).getTime() + 8 * 36e5) : new Date(e.start);
    const p = rainAt_(wx, s);
    if (p && out.length < 2) out.push((e.allDay ? '' : hm_(s) + ' ') + e.name + 'のころ雨' + p + '%。' + (e.cat === 'family' ? '傘を車に' : '傘を持って'));
  });
  const i = wx.daily.time.indexOf(ymd_(now));
  if (i >= 0) {
    const lo = Math.round(wx.daily.temperature_2m_min[i]), hi = Math.round(wx.daily.temperature_2m_max[i]);
    if (lo <= 10) out.push('今朝は' + lo + '°。上着を');
    else if (hi - lo >= 10) out.push('寒暖差' + (hi - lo) + '°（' + lo + '°→' + hi + '°）。脱ぎ着しやすく');
  }
  return out;
}

// ---------- VAPID（ES256 = P-256のECDSA）。Apps Scriptに無いので自前で ----------
// BigIntは文字で書かずに作る（古い実行環境でもファイル全体は読めるように）
const BI_ = typeof BigInt === 'function' ? BigInt : Number;
const B0 = BI_(0), B1 = BI_(1), B2 = BI_(2), B3 = BI_(3), B4 = BI_(4), B8 = BI_(8);
const P256 = {
  p: BI_('0xffffffff00000001000000000000000000000000ffffffffffffffffffffffff'),
  n: BI_('0xffffffff00000000ffffffffffffffffbce6faada7179e84f3b9cac2fc632551'),
  gx: BI_('0x6b17d1f2e12c4247f8bce6e563a440f277037d812deb33a0f4a13945d898c296'),
  gy: BI_('0x4fe342e2fe1a7f9b8ee7eb4a7c0f9e162bce33576b315ececbb6406837bf51f5')
};
function mod_(a, m) { const r = a % m; return r < B0 ? r + m : r; }
function powmod_(b, e, m) { let r = B1; b = mod_(b, m); while (e > B0) { if (e & B1) r = r * b % m; b = b * b % m; e >>= B1; } return r; }
function inv_(a, m) { return powmod_(a, m - B2, m); }
// ヤコビアン座標 [X, Y, Z]
function pDbl_(P) {
  const p = P256.p, X = P[0], Y = P[1], Z = P[2];
  if (Y === B0 || Z === B0) return [B0, B1, B0];
  const d = Y * Y % p, a = B4 * X * d % p, zz = Z * Z % p, b = B3 * mod_(X - zz, p) * (X + zz) % p;
  const x3 = mod_(b * b - B2 * a, p), y3 = mod_(b * (a - x3) - B8 * d * d, p), z3 = B2 * Y * Z % p;
  return [x3, y3, z3];
}
function pAdd_(P, Q) {
  const p = P256.p;
  if (P[2] === B0) return Q; if (Q[2] === B0) return P;
  const z1 = P[2] * P[2] % p, z2 = Q[2] * Q[2] % p;
  const u1 = P[0] * z2 % p, u2 = Q[0] * z1 % p, s1 = P[1] * z2 % p * Q[2] % p, s2 = Q[1] * z1 % p * P[2] % p;
  if (u1 === u2) return s1 === s2 ? pDbl_(P) : [B0, B1, B0];
  const h = mod_(u2 - u1, p), r = mod_(s2 - s1, p), hh = h * h % p, hhh = hh * h % p, v = u1 * hh % p;
  const x3 = mod_(r * r - hhh - B2 * v, p), y3 = mod_(r * (v - x3) - s1 * hhh, p), z3 = h * P[2] % p * Q[2] % p;
  return [x3, y3, z3];
}
function pMul_(k) {
  let R = [B0, B1, B0], A = [P256.gx, P256.gy, B1];
  while (k > B0) { if (k & B1) R = pAdd_(R, A); A = pDbl_(A); k >>= B1; }
  const zi = inv_(R[2], P256.p), zi2 = zi * zi % P256.p;
  return [R[0] * zi2 % P256.p, R[1] * zi2 % P256.p * zi % P256.p];
}
function u8_(bytes) { return bytes.map(function (b) { return b & 255; }); }
function s8_(bytes) { return bytes.map(function (b) { return b > 127 ? b - 256 : b; }); }
function big2b_(x) { const h = x.toString(16).padStart(64, '0'), out = []; for (let i = 0; i < 64; i += 2) out.push(parseInt(h.substr(i, 2), 16)); return out; }
function b2big_(b) { return BigInt('0x' + (u8_(b).map(function (x) { return ('0' + x.toString(16)).slice(-2); }).join('') || '0')); }
function b64u_(bytes) { return Utilities.base64EncodeWebSafe(s8_(bytes)).replace(/=+$/, ''); }
function sha256_(bytes) { return u8_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s8_(bytes))); }
function hmac_(key, msg) { return u8_(Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, s8_(msg), s8_(key))); }
function utf8_(s) { return u8_(Utilities.newBlob(s).getBytes()); }

function p256NewKey_() {
  let seed = [];
  for (let i = 0; i < 4; i++) seed = seed.concat(utf8_(Utilities.getUuid()));
  const d = mod_(b2big_(sha256_(seed.concat(utf8_(String(Date.now()))))), P256.n - B1) + B1;
  const Q = pMul_(d);
  return { priv: d.toString(16), pub: b64u_([4].concat(big2b_(Q[0]), big2b_(Q[1]))) };
}
/** ECDSA署名（kはRFC 6979で決める＝乱数に頼らない） */
function es256_(msgBytes, privHex) {
  const n = P256.n, d = BigInt('0x' + privHex), h = sha256_(msgBytes), e = mod_(b2big_(h), n);
  const x = big2b_(d), hb = big2b_(e);
  let V = new Array(32).fill(1), K = new Array(32).fill(0);
  K = hmac_(K, V.concat([0], x, hb)); V = hmac_(K, V);
  K = hmac_(K, V.concat([1], x, hb)); V = hmac_(K, V);
  for (;;) {
    V = hmac_(K, V);
    const k = b2big_(V);
    if (k > B0 && k < n) {
      const r = mod_(pMul_(k)[0], n);
      const s = mod_(inv_(k, n) * (e + r * d), n);
      if (r !== B0 && s !== B0) return big2b_(r).concat(big2b_(s));
    }
    K = hmac_(K, V.concat([0])); V = hmac_(K, V);
  }
}
function vapidJwt_(aud, privHex) {
  const cache = CacheService.getScriptCache(), ck = 'jwt:' + aud, c = cache.get(ck);
  if (c) return c;
  const enc = function (o) { return b64u_(utf8_(JSON.stringify(o))); };
  const head = enc({ typ: 'JWT', alg: 'ES256' }) + '.' + enc({ aud: aud, exp: Math.floor(Date.now() / 1000) + 12 * 3600, sub: PUSH_SUB });
  const jwt = head + '.' + b64u_(es256_(utf8_(head), privHex));
  cache.put(ck, jwt, 6 * 3600);
  return jwt;
}
