/**
 * リヴァイ：Googleカレンダーの予定をアプリに渡し、アプリから入れた予定を登録する
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

const DAYS_AHEAD = 60; // 締切の地平線のために60日先まで読む

function doGet(e) {
  const key = PropertiesService.getScriptProperties().getProperty('LEVI_KEY');
  const given = e && e.parameter ? e.parameter.k : '';
  if (!key || given !== key) return json_({ error: 'unauthorized' });

  const now = new Date();
  const start = new Date(now);
  start.setHours(0, 0, 0, 0);
  const end = new Date(start.getTime() + DAYS_AHEAD * 24 * 60 * 60 * 1000);

  const events = CalendarApp.getDefaultCalendar().getEvents(start, end).map(function (ev) {
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
  });

  return json_({ now: now.toISOString(), events: events });
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
