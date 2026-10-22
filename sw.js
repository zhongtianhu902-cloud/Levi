// リヴァイの通知係。Apps Scriptからの合図（中身なし）を受けたら、受け箱から文面を取って出す
// 画面の保存（オフライン）はしない。いつも最新のアプリを開くため
const CFG = 'levi-cfg';
const getJ = async k => { try { const r = await (await caches.open(CFG)).match(k); return r ? await r.json() : null; } catch (e) { return null; } };
const putJ = async (k, v) => (await caches.open(CFG)).put(k, new Response(JSON.stringify(v)));

self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));

self.addEventListener('push', e => e.waitUntil((async () => {
  let list = [];
  try {
    const c = await getJ('cfg'), seen = (await getJ('seen')) || 0;
    const d = await (await fetch(`${c.u}?k=${encodeURIComponent(c.k)}&a=inbox&t=${Date.now()}`)).json();
    list = (d.inbox || []).filter(m => m.id > seen && Date.now() - m.at < 3 * 36e5);
    if (list.length) await putJ('seen', Math.max(...list.map(m => m.id)));
  } catch (err) {}
  // iPhoneは合図のたびに必ず通知を出す決まり。文面が取れなかったときも一言出す
  if (!list.length) list = [{ title: 'リヴァイ', body: '開いて、いまやることを確かめて', tag: 'levi' }];
  for (const m of list.slice(-3)) {
    await self.registration.showNotification(m.title, { body: m.body, tag: m.tag || String(m.id), icon: 'icon.png', badge: 'icon.png', data: { url: './' } });
  }
})()));

// 通知を押したら A「いま」を開く
self.addEventListener('notificationclick', e => {
  e.notification.close();
  e.waitUntil((async () => {
    const all = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
    for (const c of all) { c.postMessage('goA'); if ('focus' in c) return c.focus(); }
    return self.clients.openWindow('./');
  })());
});
