/**
 * リヴァイ ステップ1：Googleカレンダーの予定をアプリに渡す
 *
 * 使い方
 * 1. この内容をApps Scriptに貼り、setup() を1回実行する（ログに合言葉が出る）
 * 2. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *    実行するユーザー：自分　／　アクセスできるユーザー：全員
 * 3. 出てきたURLと合言葉を、アプリの設定に入れる
 *
 * 合言葉が合わないアクセスには何も返さない。
 */

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
