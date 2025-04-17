// -----------------------------------------------------------------------------
//  poll.js – long‑poll loop that talks to DesktopAutomator’s bridge
// -----------------------------------------------------------------------------
const BASE = 'http://127.0.0.1:9234';

export function startPolling() { loop(); }

async function loop() {
  while (true) {
    try {
      const res = await fetch(`${BASE}/pending`, { cache: 'no-store' });
      if (!res.ok) { await sleep(300); continue; }

      const cmd = await res.json();
      if (!cmd.reqId) {                        // queue was empty
        await sleep(300);
        continue;
      }

      switch (cmd.action) {
        // ----------------------------------------------------------- list tabs
        case 'getTabs': {
          const tabs = await chrome.tabs.query({});
          await deliver(cmd.reqId, tabs.map(t => ({
            windowId: t.windowId,
            tabId:    t.id,
            title:    t.title,
            url:      t.url,
            active:   t.active
          })));
          break;
        }

        // ------------------------------------------------------- activate tab
        case 'activate': {
            await chrome.windows.update(cmd.args.windowId, { focused: true });
            await chrome.tabs.update  (cmd.args.tabId,    { active:  true });
          
            const info = await chrome.tabs.get(cmd.args.tabId);   // grab title
            await deliver(cmd.reqId, { ok: true, title: info.title });
            break;
          }

        // --------------------------------------------------------- close tab
        case 'close':
          await chrome.tabs.remove(cmd.args.tabId);
          await deliver(cmd.reqId, 'ok');
          break;

        // ---------------------------------------------------------- open tab
        case 'open': {
            const nt = await chrome.tabs.create({ url: cmd.args.url });
            await chrome.windows.update(nt.windowId, { focused: true });          // bring inside Chrome
            await deliver(cmd.reqId, { windowId: nt.windowId, tabId: nt.id });    // no title guessing
            break;
          }
          

        default:
          await deliver(cmd.reqId, { error: 'unknown action' });
      }
    } catch (e) {
      // network error etc. – wait a bit, then try again
      await sleep(750);
    }
  }
}

async function deliver(id, data) {
  await fetch(`${BASE}/deliver`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ reqId: id, data })
  });
}

const sleep = ms => new Promise(r => setTimeout(r, ms));
