/**
 * 以 JSONP 呼叫 Apps Script（script 標籤不受 CORS 限制）。
 * 並提供 google.script.run 相容介面。
 */
(function (global) {
  var callId = 0;
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
    bar.textContent = '尚未設定後端網址：請開啟 js/config.js，填入 Google Apps Script Web App 網址。';
    document.body.insertBefore(bar, document.body.firstChild);
  }

  function getAuthPassword() {
    try {
      return sessionStorage.getItem('SCHOOL_SCORE_AUTH') || '';
    } catch (err) {
      return '';
    }
  }

  function buildPayload(action, args) {
    var payload = {
      action: action,
      args: args || []
    };
    if (MUTATING[action]) {
      payload.authPassword = getAuthPassword();
    }
    return payload;
  }

  function unwrap(data) {
    if (data && data.__exception) {
      throw new Error(data.message || '後端發生錯誤');
    }
    return data;
  }

  function jsonpGet(payload) {
    return new Promise(function (resolve, reject) {
      var cb = 'schoolScoreCb' + (++callId) + '_' + Date.now();
      var script = document.createElement('script');
      var timer = setTimeout(function () {
        cleanup();
        reject(new Error('後端沒有回應。請把專案裡的 Code.gs 貼到 Apps Script，再部署「新版本」，存取權選「任何人」。'));
      }, 30000);

      function cleanup() {
        clearTimeout(timer);
        try { delete global[cb]; } catch (err) { global[cb] = undefined; }
        if (script.parentNode) script.parentNode.removeChild(script);
      }

      global[cb] = function (data) {
        cleanup();
        try {
          resolve(unwrap(data));
        } catch (err) {
          reject(err);
        }
      };

      script.onerror = function () {
        cleanup();
        reject(new Error('無法連到後端，請確認 Web App 已部署最新 Code.gs，且存取權為「任何人」。'));
      };

      script.src = getUrl()
        + '?payload=' + encodeURIComponent(JSON.stringify(payload))
        + '&callback=' + encodeURIComponent(cb);
      document.head.appendChild(script);
    });
  }

  function uploadPhotoByChunks(base64, filename) {
    var chunkSize = 800;
    var data = String(base64 || '');
    var total = Math.ceil(data.length / chunkSize) || 1;
    var chain = Promise.resolve();
    var i;
    for (i = 0; i < total; i++) {
      chain = (function (index, next) {
        return next.then(function () {
          var part = data.substr(index * chunkSize, chunkSize);
          return jsonpGet(buildPayload('uploadPhotoChunk', [filename, index, total, part]));
        });
      })(i, chain);
    }
    return chain.then(function () {
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
      return uploadPhotoByChunks(args[0], args[1]);
    }
    return jsonpGet(buildPayload(action, args));
  }

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
})(window);
