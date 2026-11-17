/* VisualTeX for WPS Presentation: the installed bootstrap.
 *
 * Keep this file small and stable. WPS asks the user to trust an add-in again
 * whenever its files change, so the add-in's logic (module-wpp.js) is not
 * installed: the running VisualTeX app serves it on 127.0.0.1 and it updates
 * with the app. config.js, written at install time, holds the port and key.
 */
var VisualTeXWps = {
  config: window.VISUALTEX_WPS_CONFIG || {},
  ribbon: null,
  module: null,
  moduleVersion: "",
  loading: null
};

function vtxUrl(path) {
  var config = VisualTeXWps.config;
  return "http://127.0.0.1:" + config.port + "/wps/v1" + path +
    (path.indexOf("?") >= 0 ? "&" : "?") + "host=wpp&token=" + encodeURIComponent(config.token || "");
}

function vtxGet(path, timeout, done) {
  var xhr = new XMLHttpRequest();
  xhr.open("GET", vtxUrl(path), true);
  xhr.timeout = timeout;
  xhr.onload = function () { done(xhr.status === 200 ? xhr.responseText : null); };
  xhr.onerror = xhr.ontimeout = function () { done(null); };
  xhr.send();
}

function vtxModuleVersion(hello) {
  try { return JSON.parse(hello).module || ""; } catch (error) { return ""; }
}

/* Loads the module from VisualTeX once; callers wait for the same attempt. */
function vtxLoadModule(callback) {
  if (VisualTeXWps.module) {
    callback(true);
    return;
  }
  if (VisualTeXWps.loading) {
    VisualTeXWps.loading.push(callback);
    return;
  }
  VisualTeXWps.loading = [callback];
  var finish = function (ok) {
    var waiting = VisualTeXWps.loading || [];
    VisualTeXWps.loading = null;
    for (var i = 0; i < waiting.length; i++) waiting[i](ok);
  };
  vtxGet("/hello", 4000, function (hello) {
    if (hello === null) {
      finish(false);
      return;
    }
    vtxGet("/module.js", 4000, function (source) {
      if (source === null) {
        finish(false);
        return;
      }
      try {
        (0, eval)(source);
        VisualTeXWps.module = window.VisualTeXWpsModule;
        VisualTeXWps.moduleVersion = vtxModuleVersion(hello);
        VisualTeXWps.module.init(VisualTeXWps);
        /* Lists and labels the module provides replace the placeholders. */
        try { if (VisualTeXWps.ribbon) VisualTeXWps.ribbon.Invalidate(); } catch (error) { /* older WPS */ }
        finish(true);
      } catch (error) {
        VisualTeXWps.module = null;
        finish(false);
      }
    });
  });
}

/* The logic comes from the running VisualTeX, so an updated VisualTeX serves a
   newer module while WPS stays open: take it before a command, unless a
   formula is being edited with the current one. */
function vtxRefreshModule(callback) {
  if (!VisualTeXWps.module) {
    vtxLoadModule(callback);
    return;
  }
  vtxGet("/hello", 2000, function (hello) {
    var module = VisualTeXWps.module;
    var version = hello === null ? "" : vtxModuleVersion(hello);
    if (module && version && version !== VisualTeXWps.moduleVersion &&
        (!module.isIdle || module.isIdle())) {
      VisualTeXWps.module = null;
    }
    vtxLoadModule(callback);
  });
}

/* Double-clicks run the module already loaded: look for a newer one now and
   then while the user works. */
var vtxLastRefresh = 0;

function vtxRefreshSoon() {
  var now = Date.now();
  if (now - vtxLastRefresh < 30000) return;
  vtxLastRefresh = now;
  vtxRefreshModule(function () {});
}

/* VisualTeX may start after WPS: keep trying quietly in the background. */
function vtxKeepLoading() {
  vtxLoadModule(function (ok) {
    if (!ok) setTimeout(vtxKeepLoading, 5000);
  });
}

function vtxForward(name) {
  var module = VisualTeXWps.module;
  if (module && typeof module[name] === "function") module[name]();
}

function vtxRegisterEvents() {
  try {
    wps.ApiEvent.AddApiEventListener("WindowBeforeDoubleClick", function () {
      var module = VisualTeXWps.module;
      if (module && module.onBeforeDoubleClick && module.onBeforeDoubleClick()) {
        try { wps.ApiEvent.Cancel = true; } catch (error) { /* older WPS */ }
      }
    });
  } catch (error) { /* events unavailable in this WPS version */ }
  var forwarded = {
    WindowSelectionChange: "onSelectionChange",
    WindowActivate: "onWindowActivate",
    WindowDeactivate: "onWindowDeactivate"
  };
  for (var name in forwarded) {
    try {
      wps.ApiEvent.AddApiEventListener(name, (function (handler) {
        return function () {
          vtxForward(handler);
          vtxRefreshSoon();
        };
      })(forwarded[name]));
    } catch (error) { /* events unavailable in this WPS version */ }
  }
}

function OnAddinLoad(ribbonUI) {
  VisualTeXWps.ribbon = ribbonUI;
  vtxRegisterEvents();
  vtxKeepLoading();
  return true;
}

function OnVisualTeXAction(control) {
  var id = control && control.Id;
  vtxRefreshModule(function (ok) {
    if (ok) {
      VisualTeXWps.module.onAction(id);
    } else {
      alert("VisualTeX 没有运行，或 WPS 插件需要修复。请先打开 VisualTeX，然后再试。");
    }
  });
}

function OnVisualTeXImage(control) {
  return "images/" + control.Id + ".svg";
}

var VTX_NOT_RUNNING = "VisualTeX 没有运行";

/* Ribbon queries must answer at once: the loaded module answers them, and a
   placeholder stands in until VisualTeX is running. */
function vtxAsk(name, args, fallback) {
  var module = VisualTeXWps.module;
  if (!module || typeof module[name] !== "function") return fallback;
  try {
    var value = module[name].apply(module, args);
    return value === undefined || value === null ? fallback : value;
  } catch (error) {
    return fallback;
  }
}

function OnVisualTeXMenuContent(control) {
  return vtxAsk("menuContent", [control && control.Id],
    '<menu xmlns="http://schemas.microsoft.com/office/2006/01/customui">' +
    '<button id="vtxNotRunning" label="' + VTX_NOT_RUNNING + '" enabled="false" /></menu>');
}

function OnVisualTeXItemCount(control) {
  return vtxAsk("itemCount", [control && control.Id], 1);
}

function OnVisualTeXItemLabel(control, index) {
  return vtxAsk("itemLabel", [control && control.Id, index], VTX_NOT_RUNNING);
}

function OnVisualTeXSelectedIndex(control) {
  return vtxAsk("selectedIndex", [control && control.Id], 0);
}

function OnVisualTeXSelect(control, selectedId, selectedIndex) {
  var id = control && control.Id;
  vtxRefreshModule(function (ok) {
    if (ok && VisualTeXWps.module.onSelect) {
      VisualTeXWps.module.onSelect(id, Number(selectedIndex));
    } else if (!ok) {
      alert("VisualTeX 没有运行，或 WPS 插件需要修复。请先打开 VisualTeX，然后再试。");
    }
  });
}
