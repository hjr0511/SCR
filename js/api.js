/**
 * 一般資料用 JSONP（一進頁就能點按鈕）。
 * 照片等大資料才在需要時載入隱藏 iframe + google.script.run（一次傳完）。
 */
(function (global) {
  var callId = 0;
  var iframe = null;
  var iframeReady = false;
  var iframeFailed = false;
  var readyWaiters = [];
  var pending = {};
  var MUTATING = {
    saveScore: true,
    uploadSinglePhoto: true,
    uploadPhotoChunk: true,
    finalizePhotoUpload: true
  };

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
    var cfgFile = (global.APP_CONFIG && global.APP_CONFIG.CONFIG_FILE) || 'js/config.js';
    bar.textContent = '尚未設定後端網址：請開啟 ' + cfgFile + '，填入 Google Apps Script Web App 網址。';
    document.body.insertBefore(bar, document.body.firstChild);
  }

  function getAuthPassword() {
    try {
      var cfg = global.APP_CONFIG || {};
      var key = cfg.AUTH_STORAGE_KEY || 'SCHOOL_SCORE_AUTH';
      return sessionStorage.getItem(key) || '';
    } catch (err) {
      return '';
    }
  }

  function buildPayload(action, args) {
    var payload = { action: action, args: args || [] };
    if (MUTATING[action]) payload.authPassword = getAuthPassword();
    return payload;
  }

  function unwrap(data) {
    if (data && data.__exception) throw new Error(data.message || '後端發生錯誤');
    return data;
  }

  function ensureIframe() {
    if (iframe) return iframe;
    iframe = document.createElement('iframe');
    iframe.id = 'gas-api-bridge';
    iframe.title = 'Google Apps Script bridge';
    iframe.tabIndex = -1;
    iframe.setAttribute('aria-hidden', 'true');
    iframe.setAttribute('tabindex', '-1');
    iframe.style.cssText = 'position:fixed;left:-10000px;top:0;width:1px;height:1px;border:0;opacity:0;pointer-events:none;visibility:hidden;z-index:-1;';
    iframe.addEventListener('focus', function () {
      try { iframe.blur(); } catch (err) {}
    });
    iframe.src = getUrl() + '?page=bridge';
    (document.body || document.documentElement).appendChild(iframe);
    return iframe;
  }

  function postToBridge(data) {
    if (!iframe || !iframe.contentWindow) return;
    var win = iframe.contentWindow;
    try { win.postMessage(data, '*'); } catch (err) {}
    try {
      var n = win.length || 0;
      var i;
      for (i = 0; i < n; i++) {
        try { win[i].postMessage(data, '*'); } catch (innerErr) {}
      }
    } catch (err) {}
  }

  function waitForBridge(timeoutMs) {
    if (iframeReady) return Promise.resolve(true);
    if (iframeFailed || !isConfigured()) return Promise.resolve(false);
    ensureIframe();
    return new Promise(function (resolve) {
      var settled = false;
      var timer = setTimeout(function () {
        if (settled) return;
        settled = true;
        iframeFailed = true;
        resolve(false);
      }, timeoutMs || 2500);
      readyWaiters.push({
        resolve: function () {
          if (settled) return;
          settled = true;
          clearTimeout(timer);
          resolve(true);
        }
      });
    });
  }

  function callViaIframe(action, args) {
    return waitForBridge(2500).then(function (ok) {
      if (!ok || !iframe || !iframe.contentWindow) {
        throw new Error('NO_IFRAME');
      }
      return new Promise(function (resolve, reject) {
        var id = 'c' + (++callId);
        pending[id] = { resolve: resolve, reject: reject };
        postToBridge({
          type: 'gas-call',
          id: id,
          action: action,
          args: args || [],
          authPassword: MUTATING[action] ? getAuthPassword() : ''
        });
        setTimeout(function () {
          if (!pending[id]) return;
          delete pending[id];
          reject(new Error('呼叫後端逾時：' + action));
        }, 120000);
      });
    });
  }

  function jsonpGet(payload) {
    return new Promise(function (resolve, reject) {
      var cb = 'schoolScoreCb' + (++callId) + '_' + Date.now();
      var script = document.createElement('script');
      var timer = setTimeout(function () {
        cleanup();
        reject(new Error('後端沒有回應。請把專案裡的 Code.gs 貼到 Apps Script，再部署「新版本」。'));
      }, 30000);

      function cleanup() {
        clearTimeout(timer);
        try { delete global[cb]; } catch (err) { global[cb] = undefined; }
        if (script.parentNode) script.parentNode.removeChild(script);
      }

      global[cb] = function (data) {
        cleanup();
        try { resolve(unwrap(data)); } catch (err) { reject(err); }
      };
      script.onerror = function () {
        cleanup();
        reject(new Error('無法連到後端，請確認 Web App 存取權為「任何人」。'));
      };
      var qs = [
        'action=' + encodeURIComponent(payload.action || ''),
        'args=' + encodeURIComponent(JSON.stringify(payload.args || [])),
        'callback=' + encodeURIComponent(cb)
      ];
      if (payload.authPassword) qs.push('authPassword=' + encodeURIComponent(payload.authPassword));
      script.src = getUrl() + '?' + qs.join('&');
      document.head.appendChild(script);
    });
  }

  function runPool(tasks, limit) {
    var i = 0;
    var active = 0;
    return new Promise(function (resolve, reject) {
      var done = 0;
      var failed = false;
      function next() {
        if (failed) return;
        if (done >= tasks.length) {
          resolve();
          return;
        }
        while (active < limit && i < tasks.length) {
          active++;
          tasks[i++]().then(function () {
            active--;
            done++;
            next();
          }).catch(function (err) {
            failed = true;
            reject(err);
          });
        }
      }
      if (!tasks.length) resolve();
      else next();
    });
  }

  function uploadPhotoByChunks(base64, filename) {
    var chunkSize = 2400;
    var data = String(base64 || '');
    var total = Math.ceil(data.length / chunkSize) || 1;
    var tasks = [];
    var n;
    for (n = 0; n < total; n++) {
      tasks.push((function (index) {
        return function () {
          var part = data.substr(index * chunkSize, chunkSize);
          return jsonpGet(buildPayload('uploadPhotoChunk', [filename, index, total, part]));
        };
      })(n));
    }
    return runPool(tasks, 8).then(function () {
      return jsonpGet(buildPayload('finalizePhotoUpload', [filename]));
    });
  }

  function callApi(action, args) {
    if (!isConfigured()) {
      showConfigError();
      return Promise.reject(new Error('尚未設定 Apps Script Web App 網址（js/config.js）'));
    }
    args = args || [];
    if (action === 'uploadSinglePhoto') {
      return callViaIframe(action, args).catch(function () {
        return uploadPhotoByChunks(args[0], args[1]);
      });
    }
    return jsonpGet(buildPayload(action, args));
  }

  global.addEventListener('message', function (e) {
    var msg = e.data || {};
    if (!msg || typeof msg !== 'object') return;
    if (msg.type === 'gas-ready') {
      iframeReady = true;
      readyWaiters.splice(0).forEach(function (item) {
        if (item.resolve) item.resolve();
      });
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
          var fnArgs = Array.prototype.slice.call(arguments);
          callApi(String(prop), fnArgs).then(function (result) {
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
    if (!isConfigured()) showConfigError();
  }
  if (document.body) boot();
  else document.addEventListener('DOMContentLoaded', boot);
})(window);
