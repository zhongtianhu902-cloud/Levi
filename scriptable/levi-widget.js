// リヴァイ：今日のToDoを待受（ロック画面・常時表示）に出す。7時〜18時だけ
// 使い方：Scriptableに新しいスクリプトを作ってこれを貼る → ▶で1回実行し、URLと合言葉を入れる
//         待受を長押し →カスタマイズ→ロック画面→ウィジェット追加→Scriptable（長方形）→ 長押しで「Script」をこのスクリプトに
// 押すとリヴァイ（今日の部屋）が開く

const KC = 'levi_cfg', FROM = 7, UNTIL = 18;
const APP = 'https://zhongtianhu902-cloud.github.io/Levi/';
let cfg = Keychain.contains(KC) ? JSON.parse(Keychain.get(KC)) : null;

if (config.runsInApp) {
  const a = new Alert();
  a.title = 'リヴァイの待受';
  a.message = 'Apps ScriptのURLと合言葉（リヴァイの⚙︎と同じもの）';
  a.addTextField('https://script.google.com/macros/s/…/exec', cfg ? cfg.u : '');
  a.addTextField('合言葉', cfg ? cfg.k : '');
  a.addAction('保存');
  a.addCancelAction('そのまま');
  if (await a.present() === 0) {
    cfg = { u: a.textFieldValue(0).trim(), k: a.textFieldValue(1).trim() };
    Keychain.set(KC, JSON.stringify(cfg));
  }
}

const fm = FileManager.local(), cache = fm.joinPath(fm.documentsDirectory(), 'levi_todo.json');
async function load() {
  try {
    const r = new Request(`${cfg.u}?k=${encodeURIComponent(cfg.k)}&a=todo&t=${Date.now()}`);
    r.timeoutInterval = 15;
    const d = await r.loadJSON();
    if (d.error) throw new Error(d.error);
    fm.writeString(cache, JSON.stringify(d));
    return d;
  } catch (e) {
    return fm.fileExists(cache) ? JSON.parse(fm.readString(cache)) : { items: [], error: String(e) };
  }
}

const now = new Date(), h = now.getHours(), on = h >= FROM && h < UNTIL;
const fam = config.widgetFamily || 'accessoryRectangular';
const w = new ListWidget();
w.url = APP;

if (!cfg) {
  w.addText('リヴァイ：Scriptableで1回実行して設定');
} else if (on || config.runsInApp) {
  const d = await load();
  const today = (d.items || []).filter(t => t.today).sort((a, b) => a.order - b.order);
  const hot = t => t.due && (new Date(t.due + 'T00:00') - new Date(now.getFullYear(), now.getMonth(), now.getDate())) / 864e5 <= 3;
  if (fam === 'accessoryInline') {
    w.addText(today.length ? `${today[0].title}（あと${today.length}）` : '今日の分はまだ選んでいない');
  } else if (fam === 'accessoryCircular') {
    const t = w.addText(String(today.length)); t.font = Font.boldSystemFont(22); t.centerAlignText();
    const s = w.addText('あと'); s.font = Font.systemFont(10); s.centerAlignText();
  } else {
    const head = w.addText(today.length ? `今日 あと${today.length}` : '今日の分はまだ');
    head.font = Font.boldSystemFont(12);
    const max = fam === 'accessoryRectangular' ? 3 : 8;
    today.slice(0, max).forEach((t, i) => {
      const x = w.addText(`${i + 1} ${hot(t) ? '!' : ''}${t.title}`);
      x.font = Font.systemFont(12); x.lineLimit = 1;
    });
    if (!today.length) { const x = w.addText('リヴァイの「積む」で☆'); x.font = Font.systemFont(12); }
  }
}

// 次に描き直す時刻：時間外なら次の7時、時間内なら15分後（実際の間隔はiPhoneが決める）
const next = new Date(now);
if (h >= UNTIL) { next.setDate(next.getDate() + 1); next.setHours(FROM, 0, 0, 0); }
else if (h < FROM) next.setHours(FROM, 0, 0, 0);
else next.setMinutes(next.getMinutes() + 15);
w.refreshAfterDate = next;

if (config.runsInApp) await w.presentAccessoryRectangular();
Script.setWidget(w);
Script.complete();
