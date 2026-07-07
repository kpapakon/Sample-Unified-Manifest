/* eslint-disable @typescript-eslint/no-unused-expressions */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/prefer-for-of */
// addin-monitor.js
// Paste this file very early in your shared runtime (index.html) *before* your app bundle.
// It patches some Office/Word write APIs, snapshots OOXML before/after initialization,
// and prints a diagnostics report to the console. It also exposes window.__addin_monitor_markReady()
// so you can invoke the "after" snapshot from your Angular bootstrap when ready.

(function () {
  const CONFIG = {
    afterSnapshotDelayMs: 5000, // default automatic delay (ms) before taking "after" snapshot
    enableAutoAfterSnapshot: true,
    maxStackLength: 2000
  };

  const state = {
    beforeOoxml: null,
    afterOoxml: null,
    events: [],
    patched: false,
    original: {}
  };

  // small helpers
  function nowIso() { return (new Date()).toISOString(); }
  function safeSlice(s, n = 200) { return typeof s === 'string' ? (s.length > n ? s.slice(0, n) + '…' : s) : s; }
  function getStack() {
    const e = new Error();
    if (!e.stack) return '(no-stack)';
    // remove first two frames for clarity (getStack + wrapper)
    const lines = e.stack.split('\n').slice(2).join('\n');
    return lines.slice(0, CONFIG.maxStackLength);
  }
  async function sha256Hex(str) {
    try {
      const enc = new TextEncoder();
      const buf = enc.encode(str);
      const hash = await crypto.subtle.digest('SHA-256', buf);
      return Array.from(new Uint8Array(hash)).map(b => b.toString(16).padStart(2, '0')).join('');
    } catch (e) {
      return 'hash-failed';
    }
  }
  function downloadFile(filename, content, mime='application/xml') {
    try {
      const blob = new Blob([content], { type: mime });
      const url = URL.createObjectURL(blob);
      console.log(`[Monitor] Download ${filename}: ${url} (click to open)`);
      // also append a hidden link to the DOM so it is clickable:
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      a.textContent = `Download ${filename}`;
      a.style = 'display:block;font-family:monospace;color:#06c';
      const holderId = '__addin_monitor_downloads';
      let holder = document.getElementById(holderId);
      if (!holder) {
        holder = document.createElement('div');
        holder.id = holderId;
        holder.style = 'position:fixed;right:10px;bottom:10px;z-index:999999;background:#fff;border:1px solid #ccc;padding:8px;max-width:320px;max-height:40vh;overflow:auto;font-size:12px';
        document.body && document.body.appendChild(holder);
      }
      holder.appendChild(a);
      return url;
    } catch (e) {
      console.warn('[Monitor] downloadFile failed', e);
    }
  }

  function recordEvent(type, details) {
    const ev = { time: nowIso(), type, details, stack: getStack() };
    state.events.push(ev);
    console.warn(`[Monitor][${ev.time}] ${type}`, details);
  }

  // Patch APIs to detect write calls
  function safePatch(pathArray, methodName, wrapper) {
    try {
      let obj = window;
      for (let i = 0; i < pathArray.length; i++) {
        if (!obj[pathArray[i]]) return null;
        obj = obj[pathArray[i]];
      }
      const orig = obj[methodName];
      if (!orig || typeof orig !== 'function') return null;
      // store original reference
      state.original[`${pathArray.join('.')}.${methodName}`] = orig;
      obj[methodName] = wrapper(orig);
      return true;
    } catch (e) {
      console.warn('[Monitor] safePatch error', e);
      return null;
    }
  }

  // Patching logic when Office is available
  function applyPatches() {
    if (!window.Office) return;
    if (state.patched) return;
    state.patched = true;

    // Patch Office.context.document.settings.set
    try {
      if (window.Office.context && window.Office.context.document && window.Office.context.document.settings) {
        const settings = window.Office.context.document.settings;
        if (typeof settings.set === 'function') {
          const orig = settings.set.bind(settings);
          settings.set = function (key, value) {
            recordEvent('settings.set', { key, value: safeSlice(String(value)) });
            return orig(key, value);
          };
          state.original['Office.context.document.settings.set'] = orig;
        }
        if (typeof settings.saveAsync === 'function') {
          const origSave = settings.saveAsync.bind(settings);
          settings.saveAsync = function (callback) {
            recordEvent('settings.saveAsync', {});
            return origSave(function (...args) {
              recordEvent('settings.saveAsync.callback', {});
              if (typeof callback === 'function') callback.apply(this, args);
            });
          };
          state.original['Office.context.document.settings.saveAsync'] = origSave;
        }
      }
    } catch (e) { console.warn('[Monitor] patch settings error', e); }

    // Patch Office.context.document.setSelectedDataAsync
    try {
      const doc = window.Office.context && window.Office.context.document;
      if (doc && typeof doc.setSelectedDataAsync === 'function') {
        const origSetSel = doc.setSelectedDataAsync.bind(doc);
        doc.setSelectedDataAsync = function (...args) {
          // normalize args
          const data = args[0];
          const options = args[1];
          const details = { sample: safeSlice(typeof data === 'string' ? data : JSON.stringify(data)), options };
          recordEvent('setSelectedDataAsync', details);
          return origSetSel.apply(this, args);
        };
        state.original['Office.context.document.setSelectedDataAsync'] = origSetSel;
      }
    } catch (e) { console.warn('[Monitor] patch setSelectedDataAsync error', e); }

    // Patch customXmlParts.addAsync and getByIdAsync (if present)
    try {
      const parts = window.Office.context && window.Office.context.document && window.Office.context.document.customXmlParts;
      if (parts) {
        if (typeof parts.addAsync === 'function') {
          const origAdd = parts.addAsync.bind(parts);
          parts.addAsync = function (...args) {
            recordEvent('customXmlParts.addAsync', { sample: safeSlice(String(args[0])) });
            return origAdd.apply(this, args);
          };
          state.original['Office.context.document.customXmlParts.addAsync'] = origAdd;
        }
        if (typeof parts.getByIdAsync === 'function') {
          const origGet = parts.getByIdAsync.bind(parts);
          parts.getByIdAsync = function (...args) {
            const id = args[0];
            recordEvent('customXmlParts.getByIdAsync', { id });
            return origGet.apply(this, args);
          };
          state.original['Office.context.document.customXmlParts.getByIdAsync'] = origGet;
        }
      }
    } catch (e) { console.warn('[Monitor] patch customXmlParts error', e); }

    // Patch Word.run (if present)
    try {
      if (window.Word && typeof window.Word.run === 'function') {
        const origWordRun = window.Word.run.bind(Word);
        state.original['Word.run'] = origWordRun;
        window.Word.run = async function (callback) {
          recordEvent('Word.run.start', {});
          try {
            const r = await origWordRun(async (context) => {
              // inside the run: we can't easily intercept all write methods here,
              // but logging that a run started is useful to correlate actions.
              return callback(context);
            });
            recordEvent('Word.run.complete', {});
            return r;
          } catch (err) {
            recordEvent('Word.run.error', { message: String(err) });
            throw err;
          }
        };
      }
    } catch (e) { console.warn('[Monitor] patch Word.run error', e); }

    // Patch other common Office.js async patterns by name if present
    // (This is conservative; extend as needed)
    console.info('[Monitor] patches applied; monitoring Office write APIs.');
  }

  // OOXML snapshot helpers using Word.run
  async function takeOoxmlSnapshot(label) {
    if (!window.Word || typeof Word.run !== 'function') {
      console.warn('[Monitor] Word.run not available for OOXML snapshot');
      return null;
    }
    try {
      const res = await Word.run(async (context) => {
        const ooxml = context.document.body.getOoxml();
        await context.sync();
        return ooxml.value;
      });
      const hash = await sha256Hex(res || '');
      const rec = { label, time: nowIso(), length: (res || '').length, sha256: hash };
      if (label === 'before') state.beforeOoxml = { text: res, meta: rec };
      else state.afterOoxml = { text: res, meta: rec };
      console.info(`[Monitor] OOXML snapshot ${label}: len=${rec.length} sha256=${rec.sha256}`);
      return rec;
    } catch (e) {
      console.warn('[Monitor] takeOoxmlSnapshot failed', e);
      return null;
    }
  }

  async function produceReport() {
    const before = state.beforeOoxml && state.beforeOoxml.meta;
    const after = state.afterOoxml && state.afterOoxml.meta;
    const didChange = before && after && before.sha256 !== after.sha256;
    console.group('[Monitor] Report');
    console.log('events (most recent at bottom):', state.events);
    console.log('before meta:', before);
    console.log('after meta:', after);
    console.log('documentChanged:', !!didChange);
    if (state.beforeOoxml && state.afterOoxml) {
      // offer downloads for both OOXML snapshots
      downloadFile('ooxml_before.xml', state.beforeOoxml.text || '', 'application/xml');
      downloadFile('ooxml_after.xml', state.afterOoxml.text || '', 'application/xml');
      const report = {
        generated: nowIso(),
        before: before || null,
        after: after || null,
        events: state.events
      };
      const repStr = JSON.stringify(report, null, 2);
      downloadFile('addin_monitor_report.json', repStr, 'application/json');
      console.log('Downloaded before/after OOXML and report links are appended to the page (bottom-right).');
    } else {
      console.log('Snapshots missing. Ensure Office was ready and snapshots completed.');
    }
    console.groupEnd();
  }

  // Public hook to mark the add-in "ready" (use this from your Angular bootstrap)
  window.__addin_monitor_markReady = async function markReadyNow() {
    try {
      console.info('[Monitor] markReady called — taking after snapshot now.');
      await takeOoxmlSnapshot('after');
      await produceReport();
    } catch (e) {
      console.warn('[Monitor] markReady error', e);
    }
  };

  // Initialize logic: wait for Office.onReady and then apply patches and take "before" snapshot
  function initMonitor() {
    // try to apply patches immediately if Office already present
    applyPatches();
    // take before snapshot if possible
    (async () => {
      await takeOoxmlSnapshot('before');
      // schedule after snapshot if auto enabled
      if (CONFIG.enableAutoAfterSnapshot) {
        setTimeout(async () => {
          console.info('[Monitor] auto after-snapshot timer fired');
          await takeOoxmlSnapshot('after');
          await produceReport();
        }, CONFIG.afterSnapshotDelayMs);
      }
    })();
  }

  // If Office provides onReady, wait for it; otherwise poll for Office availability
  if (window.Office && typeof window.Office.onReady === 'function') {
    try {
      window.Office.onReady().then(() => {
        console.info('[Monitor] Office.onReady resolved');
        initMonitor();
      }).catch((e) => {
        console.warn('[Monitor] Office.onReady rejected', e);
        // still attempt immediate init
        initMonitor();
      });
    } catch (e) {
      console.warn('[Monitor] Office.onReady call error', e);
      initMonitor();
    }
  } else {
    // fallback: poll for Word.run/Office
    let tries = 0;
    const poll = setInterval(() => {
      tries++;
      if (window.Office) {
        clearInterval(poll);
        console.info('[Monitor] Office became available via poll');
        initMonitor();
      } else if (tries > 20) {
        clearInterval(poll);
        console.warn('[Monitor] Office never became available (poll ended)');
      }
    }, 300);
  }

  // Safety: if the developer wants to disable auto-after snapshot, expose config toggles:
  window.__addin_monitor_config = CONFIG;

})();