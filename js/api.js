/**
 * 呼叫 Apps Script Web App（GET / JSONP / POST）。
 * 並提供 google.script.run 相容介面，讓原有前端程式幾乎不用改呼叫方式。
 */
(function (global) {
  var callId = 0;

  function getUrl() {
    var cfg = global.APP_CONFIG || {};
    return String(cfg.GAS_WEB_APP_URL || '').trim().replace(/\/+$/, '');
  }

  function isConfigured() {
    var url = getUrl();
    return !!url && url.indexOf('YOUR_DEPLOYMENT_ID') === -1 && /^https:\/\/script\.google\.com\//.test(url);
  }

  function showConfigError(text) {
    if (!document.body) return;
    var bar = document.getElementById('gas-config-warning');
    if (!bar) {
      bar = document.createElement('div');
      bar.id = 'gas-config-warning';
      bar.style.cssText = 'position:sticky;top:0;z-index:9999;background:#c62828;color:#fff;padding:12px 16px;text-align:center;font-family:"Microsoft JhengHei",Arial,sans-serif;';
      document.body.insertBefore(bar, document.body.firstChild);
    }
    bar.textContent = text || '尚未設定後端網址：請開啟 js/config.js，填入 Google Apps Script Web App 網址。';
  }

  function getAuthPassword() {
    try {
      return sessionStorage.getItem('SCHOOL_SCORE_AUTH') || '';
    } catch (err) {
      return '';
    }
  }

  function buildPayload(action, args) {
    return {
      action: action,
      args: args || [],
      authPassword: getAuthPassword()
    };
  }

  function unwrap(data) {
    if (data && data.__exception) {
      throw new Error(data.message || '後端發生錯誤');
    }
    return data;
  }

  function looksLikeHtml(text) {
    var s = String(text || '').replace(/^\s+/, '').slice(0, 200).toLowerCase();
    return s.indexOf('<!doctype') === 0 || s.indexOf('<html') === 0 || s.indexOf('評分結果查詢') >= 0;
  }

  function parseResponseText(text) {
    if (looksLikeHtml(text)) {
      throw new Error('後端仍是舊版網頁，不是 API。請把專案裡的 Code.gs 貼到 Apps Script，再「部署 → 管理部署 → 編輯 → 新版本」。');
    }
    var data;
    try {
      data = JSON.parse(text);
    } catch (err) {
      throw new Error('後端回傳不是有效 JSON，請確認已部署最新 Code.gs');
    }
    return unwrap(data);
  }

  function fetchGet(payload) {
    var url = getUrl() + '?payload=' + encodeURIComponent(JSON.stringify(payload));
    return fetch(url, { method: 'GET', redirect: 'follow', credentials: 'omit' })
      .then(function (res) { return res.text(); })
      .then(parseResponseText);
  }

  function fetchPost(payload) {
    return fetch(getUrl(), {
      method: 'POST',
      redirect: 'follow',
      credentials: 'omit',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload)
    })
      .then(function (res) { return res.text(); })
      .then(parseResponseText);
  }

  function jsonpGet(payload) {
    return new Promise(function (resolve, reject) {
      var cb = 'schoolScoreCb' + (++callId) + '_' + Date.now();
      var script = document.createElement('script');
      var timer = setTimeout(function () {
        cleanup();
        reject(new Error('後端沒有回應。請把專案裡的 Code.gs 貼到 Apps Script，再部署「新版本」，存取權選「任何人」。'));
      }, 25000);

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
        reject(new Error('無法連到後端，請確認 Web App 網址與存取權為「任何人」。'));
      };

      var url = getUrl()
        + '?payload=' + encodeURIComponent(JSON.stringify(payload))
        + '&callback=' + encodeURIComponent(cb);
      script.src = url;
      document.head.appendChild(script);
    });
  }

  function callApi(action, args) {
    if (!isConfigured()) {
      showConfigError();
      return Promise.reject(new Error('尚未設定 Apps Script Web App 網址（js/config.js）'));
    }

    var payload = buildPayload(action, args);
    var body = JSON.stringify(payload);
    var useGet = body.length < 1800;

    var start = useGet ? fetchGet(payload) : fetchPost(payload);
    return start.catch(function (err) {
      if (err && String(err.message).indexOf('舊版網頁') >= 0) throw err;
      if (useGet) return jsonpGet(payload);
      throw err;
    });
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
})(window);
