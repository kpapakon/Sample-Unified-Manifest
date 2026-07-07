import { bootstrapApplication } from '@angular/platform-browser';
import { appConfig } from './app/app.config';
import { App } from './app/app';
import './addin-monitor';

declare global {
  interface Window {
    __addin_monitor_markReady?: () => void;
  }
}

Office.onReady().then(appInit);

function appInit() {
  Office.addin.setStartupBehavior(Office.StartupBehavior.load);

  if (!window.history.replaceState) {
    window.history.replaceState = function () {};
  }
  if (!window.history.pushState) {
    window.history.pushState = function () {};
  }
  bootstrapApplication(App, appConfig).catch((err) => console.error(err));
  if (window.__addin_monitor_markReady) window.__addin_monitor_markReady();

  initMode();
}

async function initMode() {
  const isReadOnly =
    Office.context.document.mode === Office.DocumentMode.ReadOnly;
  console.log('Document is in read-only mode:', isReadOnly);
}
