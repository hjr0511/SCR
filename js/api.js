/**
 * 透過隱藏 iframe 呼叫 Apps Script（google.script.run），避開 GitHub Pages 的 CORS / POST 轉址問題。
 * 並提供 google.script.run 相容介面，讓原有前端程式幾乎不用改呼叫方式。
 */
(function (global) {
  var pending = {};
  var callId = 0;
  var iframe = null;
  var iframeReady = false;
  var readyWaiters = [];

  function getUrl() {
    var cfg = global.APP_CONFIG || {};
    return String(cfg.GAS_WEB_APP_URL || '').trim().replace(/\/+$/, '');
  }

  function isConfigured() {
    var url = getUrl();
    return !!url && url.indexOf('YOUR_DEPLOYMENT_ID') === -1 && /^https:\/\/script\.google\.com\//.test(url);
  }

  function showConfigError() {
    if (!document.body) return;
    if (document.getElementById('gas-config-warning')) return;
    var bar = document.createElement('div');
    bar.id = 'gas-config-warning';
    bar.style.cssText = 'position:sticky;top:0;z-index:9999;background:#c62828;color:#fff;padding:12px 16px;text-align:center;font-family:"Microsoft JhengHei",Arial,sans-serif;';
    bar.textContent = '尚未設定後端網址：請開啟 js/config.js，填入 Google Apps Script Web App 網址。';
    document.body.insertBefore(bar, document.body.firstChild);
  }

  function rejectReadyWaiters(err) {
    var waiters = readyWaiters.slice();
    readyWaiters = [];
    waiters.forEach(function (item) {
      clearTimeout(item.timer);
      item.reject(err);
    });
  }

  function resolveReadyWaiters() {
    var waiters = readyWaiters.slice();
    readyWaiters = [];
    waiters.forEach(function (item) {
      clearTimeout(item.timer);
      item.resolve();
    });
  }

  function ensureIframe() {
    if (iframe) return iframe;
    iframe = document.createElement('iframe');
    iframe.id = 'gas-api-bridge';
    iframe.title = 'Google Apps Script bridge';
    iframe.setAttribute('aria-hidden', 'true');
    iframe.style.cssText = 'position:absolute;width:0;height:0;border:0;opacity:0;pointer-events:none;';
    iframe.src = getUrl() + '?page=bridge';
    (document.body || document.documentElement).appendChild(iframe);
    return iframe;
  }

  function waitForBridge() {
    if (iframeReady) return Promise.resolve();
    if (!isConfigured()) {
      if (document.body) showConfigError();
      else document.addEventListener('DOMContentLoaded', showConfigError);
      return Promise.reject(new Error('尚未設定 Apps Script Web App 網址（js/config.js）'));
    }

    ensureIframe();
    return new Promise(function (resolve, reject) {
      var timer = setTimeout(function () {
        reject(new Error('後端橋接頁載入逾時，請確認 Web App 已部署，且存取權為「任何人」。'));
      }, 20000);
      readyWaiters.push({ resolve: resolve, reject: reject, timer: timer });
    });
  }

  function callApi(action, args) {
    return waitForBridge().then(function () {
      return new Promise(function (resolve, reject) {
        var id = 'c' + (++callId);
        pending[id] = { resolve: resolve, reject: reject };
        var frameWindow = iframe && iframe.contentWindow;
        if (!frameWindow) {
          delete pending[id];
          reject(new Error('後端橋接 iframe 尚未就緒'));
          return;
        }
        frameWindow.postMessage({
          type: 'gas-call',
          id: id,
          action: action,
          args: args || [],
          authPassword: (function () {
            try { return sessionStorage.getItem('SCHOOL_SCORE_AUTH') || ''; }
            catch (err) { return ''; }
          })()
        }, '*');
        setTimeout(function () {
          if (!pending[id]) return;
          delete pending[id];
          reject(new Error('呼叫後端逾時：' + action));
        }, 180000);
      });
    });
  }

  global.addEventListener('message', function (e) {
    var msg = e.data || {};
    if (!msg || typeof msg !== 'object') return;
    if (msg.type === 'gas-ready') {
      iframeReady = true;
      resolveReadyWaiters();
      return;
    }
    if (msg.type !== 'gas-result') return;
    var item = pending[msg.id];
    if (!item) return;
    delete pending[msg.id];
    if (msg.ok) item.resolve(msg.result);
    else item.reject(new Error(msg.message || '後端發生錯誤'));
  });

  function createRunner(handlers) {
    handlers = handlers || {};
    return new Proxy({}, {
      get: function (target, prop) {
        if (prop === 'withSuccessHandler') {
          return function (fn) {
            return createRunner({ success: fn, fail: handlers.fail });
          };
        }
        if (prop === 'withFailureHandler') {
          return function (fn) {
            return createRunner({ success: handlers.success, fail: fn });
          };
        }
        return function () {
          var args = Array.prototype.slice.call(arguments);
          callApi(String(prop), args).then(function (result) {
            if (typeof handlers.success === 'function') handlers.success(result);
          }).catch(function (err) {
            if (typeof handlers.fail === 'function') handlers.fail(err);
            else console.error(err);
          });
        };
      }
    });
  }

  global.google = global.google || {};
  global.google.script = global.google.script || {};
  global.google.script.run = createRunner({});
  global.callSchoolApi = callApi;

  function boot() {
    if (isConfigured()) ensureIframe();
    else showConfigError();
  }
  if (document.body) boot();
  else document.addEventListener('DOMContentLoaded', boot);
})(window);
