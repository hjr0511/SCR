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
    finalizePhotoUpload: true,
    exportWeeklyStatisticsCsv: true,
    exportWeeklyStatisticsPdf: true,
    exportWeeklyStatisticsToSheet: true
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

  function hasPendingCalls() {
    var id;
    for (id in pending) {
      if (Object.prototype.hasOwnProperty.call(pending, id)) return true;
    }
    return false;
  }

  var lastRecreateAt = 0;

  function recreateBridge() {
    var now = Date.now();
    if (iframe && now - lastRecreateAt < 800) return;
    lastRecreateAt = now;
    iframeReady = false;
    iframeFailed = false;
    iframeWarmed = false;
    if (iframe && iframe.parentNode) iframe.parentNode.removeChild(iframe);
    iframe = null;
    ensureIframe();
  }

  function reviveBridge() {
    iframeFailed = false;
    if (!isConfigured()) return;
    if (iframeReady) return;
    if (hasPendingCalls()) return;
    recreateBridge();
  }

  function prepareBridgeForCamera() {
    return;
  }

  function waitForBridge(timeoutMs, markFailed) {
    if (iframeReady) return Promise.resolve(true);
    if (!isConfigured()) return Promise.resolve(false);
    if (iframeFailed) {
      iframeFailed = false;
      if (!hasPendingCalls()) recreateBridge();
    } else {
      ensureIframe();
    }
    return new Promise(function (resolve) {
      var settled = false;
      var timer = setTimeout(function () {
        if (settled) return;
        settled = true;
        if (markFailed === true) iframeFailed = true;
        resolve(false);
      }, timeoutMs || 8000);
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

  function postCall(action, args, timeoutMs) {
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
      }, timeoutMs || 120000);
    });
  }

  function callViaIframe(action, args) {
    return waitForBridge(2500).then(function (ok) {
      if (!ok || !iframe || !iframe.contentWindow) {
        throw new Error('NO_IFRAME');
      }
      return postCall(action, args);
    });
  }

  function jsonpGet(payload, timeoutMs) {
    return new Promise(function (resolve, reject) {
      var cb = 'schoolScoreCb' + (++callId) + '_' + Date.now();
      var script = document.createElement('script');
      var waitMs = timeoutMs || 60000;
      var settled = false;
      var timer = setTimeout(function () {
        fail(new Error('後端沒有回應。請把專案裡的 Code.gs 貼到 Apps Script，再部署「新版本」。'));
      }, waitMs);

      function cleanup() {
        clearTimeout(timer);
        global[cb] = function () {};
        if (script.parentNode) script.parentNode.removeChild(script);
      }

      function fail(err) {
        if (settled) return;
        settled = true;
        cleanup();
        reject(err);
      }

      global[cb] = function (data) {
        if (settled) return;
        settled = true;
        cleanup();
        try { resolve(unwrap(data)); } catch (err) { reject(err); }
      };
      script.onerror = function () {};
      var qs = [
        'action=' + encodeURIComponent(payload.action || ''),
        'args=' + encodeURIComponent(JSON.stringify(payload.args || [])),
        'callback=' + encodeURIComponent(cb)
      ];
      if (payload.authPassword) qs.push('authPassword=' + encodeURIComponent(payload.authPassword));
      script.async = true;
      script.src = getUrl() + '?' + qs.join('&') + '&_=' + Date.now();
      document.head.appendChild(script);
    });
  }

  function delay(ms) {
    return new Promise(function (resolve) {
      setTimeout(resolve, ms);
    });
  }

  function jsonpGetRetry(payload, timeoutMs, retries) {
    retries = retries == null ? 2 : retries;
    return jsonpGet(payload, timeoutMs).catch(function (err) {
      if (retries <= 0) {
        throw new Error('連線暫時失敗，請再試一次。');
      }
      return delay(retries === 2 ? 250 : 700).then(function () {
        return jsonpGetRetry(payload, timeoutMs, retries - 1);
      });
    });
  }

  function postTextJson(payload, timeoutMs) {
    return new Promise(function (resolve, reject) {
      if (typeof fetch !== 'function') {
        reject(new Error('NO_FETCH'));
        return;
      }
      var settled = false;
      var timer = setTimeout(function () {
        if (settled) return;
        settled = true;
        reject(new Error('後端沒有回應。請把專案裡的 Code.gs 貼到 Apps Script，再部署「新版本」。'));
      }, timeoutMs || 40000);
      fetch(getUrl(), {
        method: 'POST',
        mode: 'cors',
        redirect: 'follow',
        credentials: 'omit',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(payload)
      }).then(function (res) {
        if (!res.ok) throw new Error('HTTP ' + res.status);
        return res.json();
      }).then(function (data) {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        resolve(unwrap(data));
      }).catch(function (err) {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        reject(err);
      });
    });
  }

  function apiTimeoutFor(action) {
    if (action === 'verifyScoreSystemPassword' || action === 'ping') return 12000;
    if (action === 'getAllClassrooms' || action === 'getScoreRecords') return 90000;
    if (action === 'uploadPhotoChunk') return 20000;
    if (action === 'finalizePhotoUpload' || action === 'uploadSinglePhoto') return 40000;
    return 60000;
  }

  function apiRequest(payload, timeoutMs) {
    var waitMs = timeoutMs || apiTimeoutFor(payload && payload.action);
    return jsonpGetRetry(payload, waitMs, 1);
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

  var iframeWarmed = false;

  function uploadPhotoByChunks(base64, filename) {
    var chunkSize = 1000;
    var data = String(base64 || '');
    var comma = data.indexOf(',');
    if (comma >= 0) data = data.substring(comma + 1);
    var total = Math.ceil(data.length / chunkSize) || 1;
    var index = 0;

    function sendChunk(chunkIndex, attempt) {
      var part = data.substr(chunkIndex * chunkSize, chunkSize);
      return jsonpGet(buildPayload('uploadPhotoChunk', [filename, chunkIndex, total, part]), 10000).catch(function (err) {
        if (attempt >= 1) throw err;
        return delay(250).then(function () {
          return sendChunk(chunkIndex, attempt + 1);
        });
      });
    }

    function sendNext() {
      if (index >= total) {
        return jsonpGet(buildPayload('finalizePhotoUpload', [filename]), 20000).catch(function () {
          return delay(400).then(function () {
            return jsonpGet(buildPayload('finalizePhotoUpload', [filename]), 20000);
          });
        });
      }
      var current = index;
      index += 1;
      return sendChunk(current, 0).then(sendNext);
    }

    return sendNext();
  }

  function formPost(payload, timeoutMs) {
    return new Promise(function (resolve, reject) {
      var id = 'f' + (++callId) + '_' + Date.now();
      var iframe = document.createElement('iframe');
      iframe.name = 'gasForm_' + id;
      iframe.setAttribute('aria-hidden', 'true');
      iframe.style.cssText = 'position:fixed;left:-9999px;width:1px;height:1px;opacity:0;border:0;';
      var form = document.createElement('form');
      form.method = 'POST';
      form.action = getUrl();
      form.target = iframe.name;
      function field(name, value) {
        var input = document.createElement('input');
        input.type = 'hidden';
        input.name = name;
        input.value = value == null ? '' : String(value);
        form.appendChild(input);
      }
      field('action', payload.action || '');
      field('args', JSON.stringify(payload.args || []));
      if (payload.authPassword) field('authPassword', payload.authPassword);
      field('embed', '1');
      field('msgId', id);
      var settled = false;
      var timer = setTimeout(function () {
        finish(new Error('後端沒有回應。請把專案裡的 Code.gs 貼到 Apps Script，再部署「新版本」。'), true);
      }, timeoutMs || 25000);
      function finish(err, isReject) {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        window.removeEventListener('message', onMsg);
        if (iframe.parentNode) iframe.parentNode.removeChild(iframe);
        if (form.parentNode) form.parentNode.removeChild(form);
        if (isReject) reject(err);
      }
      function onMsg(e) {
        var msg = e && e.data;
        if (!msg || msg.type !== 'gas-form-result' || String(msg.id) !== id) return;
        if (msg.ok) {
          try {
            finish(null, false);
            resolve(unwrap(msg.result));
          } catch (err) {
            finish(err, true);
          }
        } else {
          finish(new Error(msg.message || '後端發生錯誤'), true);
        }
      }
      window.addEventListener('message', onMsg);
      (document.body || document.documentElement).appendChild(iframe);
      (document.body || document.documentElement).appendChild(form);
      form.submit();
    });
  }
  var uploadChain = Promise.resolve();
  var photoUploadCount = 0;

  function uploadSinglePhotoFast(base64, filename) {
    if (!getAuthPassword()) {
      return Promise.reject(new Error('未授權：請先從查詢頁輸入密碼進入評分系統'));
    }
    var run = uploadChain.then(function () {
      var pause = photoUploadCount === 0 ? Promise.resolve() : delay(400);
      photoUploadCount += 1;
      return pause.then(function () {
        return formPost(buildPayload('uploadSinglePhoto', [base64, filename]), 20000);
      });
    });
    uploadChain = run.then(function () {}, function () {});
    return run;
  }

  function classroomCacheKey() {
    var cfg = global.APP_CONFIG || {};
    return (cfg.AUTH_STORAGE_KEY || 'SCHOOL_SCORE_AUTH') + '_CLASSROOMS';
  }

  function readClassroomCache() {
    try {
      var parsed = JSON.parse(localStorage.getItem(classroomCacheKey()) || 'null');
      if (!parsed || !Array.isArray(parsed.list) || !parsed.list.length) return null;
      return parsed.list;
    } catch (err) {
      return null;
    }
  }

  function writeClassroomCache(list) {
    try {
      localStorage.setItem(classroomCacheKey(), JSON.stringify({ ts: Date.now(), list: list || [] }));
    } catch (err) {}
  }

  function classroomListsEqual(a, b) {
    if (!Array.isArray(a) || !Array.isArray(b) || a.length !== b.length) return false;
    for (var i = 0; i < a.length; i++) {
      var x = a[i] || {};
      var y = b[i] || {};
      if (String(x.id || '') !== String(y.id || '')) return false;
      if (String(x.name || '') !== String(y.name || '')) return false;
      if (String(x.grade || '') !== String(y.grade || '')) return false;
    }
    return true;
  }

  function shouldPreloadBridge() {
    return false;
  }

  function callApi(action, args) {
    if (!isConfigured()) {
      showConfigError();
      return Promise.reject(new Error('尚未設定 Apps Script Web App 網址（js/config.js）'));
    }
    args = args || [];
    if (action === 'uploadSinglePhoto') {
      return uploadSinglePhotoFast(args[0], args[1]);
    }
    var payload = buildPayload(action, args);
    return apiRequest(payload, apiTimeoutFor(action));
  }

  global.addEventListener('message', function (e) {
    var msg = e.data || {};
    if (!msg || typeof msg !== 'object') return;
    if (msg.type === 'gas-ready') {
      iframeReady = true;
      readyWaiters.splice(0).forEach(function (item) {
        if (item.resolve) item.resolve();
      });
      if (!iframeWarmed) {
        iframeWarmed = true;
        postCall('ping', [], 15000).catch(function () {});
      }
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
          var action = String(prop);
          var cachedClassrooms = action === 'getAllClassrooms' ? readClassroomCache() : null;
          if (cachedClassrooms && typeof handlers.success === 'function') {
            try { handlers.success(cachedClassrooms); } catch (err) {}
          }
          callApi(action, fnArgs).then(function (result) {
            if (action === 'getAllClassrooms' && Array.isArray(result) && result.length) {
              writeClassroomCache(result);
            }
            if (typeof handlers.success === 'function') {
              if (cachedClassrooms && classroomListsEqual(cachedClassrooms, result)) return;
              handlers.success(result);
            }
          }).catch(function (err) {
            if (cachedClassrooms) return;
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
  global.reviveSchoolApiBridge = reviveBridge;
  global.prepareSchoolApiBridge = prepareBridgeForCamera;

  var warmPromise = null;
  function warmBackend() {
    if (!isConfigured()) return Promise.resolve(false);
    if (warmPromise) return warmPromise;
    warmPromise = apiRequest(buildPayload('ping', []), 20000).then(function () {
      return true;
    }).catch(function () {
      return false;
    });
    return warmPromise;
  }
  global.warmSchoolApi = warmBackend;

  function onPageHidden() {
    iframeFailed = false;
  }

  function onPageVisible() {
    iframeFailed = false;
  }
  global.addEventListener('pageshow', onPageVisible);
  global.addEventListener('pagehide', onPageHidden);
  document.addEventListener('visibilitychange', function () {
    if (document.visibilityState === 'hidden') onPageHidden();
    else onPageVisible();
  });

  function boot() {
    if (!isConfigured()) showConfigError();
  }
  if (document.body) boot();
  else document.addEventListener('DOMContentLoaded', boot);
})(window);
