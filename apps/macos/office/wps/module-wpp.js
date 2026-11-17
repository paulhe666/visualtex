/* VisualTeX for WPS Presentation: the add-in's logic, served by the running
 * VisualTeX app (the installed bootstrap addin-wpp/main.js loads it).
 *
 * It does for WPS Presentation what the PowerPoint-for-Mac add-in does: picture
 * formulas (VTPowerPointAdapter.bas) with the same shape state (name,
 * AlternativeText, tags), so a presentation stays editable when it moves
 * between WPS and PowerPoint, and native equations in text
 * (VTPowerPointNativeMath.bas) the way VisualTeX for Windows writes them in
 * WPS (WpsNativeEquations). WPS cannot set a shape's Title, so the formula
 * reference the Title holds in PowerPoint is also kept in a tag.
 */
(function () {
  "use strict";

  var METADATA_PREFIX = "visualtex:v1:deflate:";
  var FORMULA_REF_PREFIX = "visualtex:formula-ref:v1:";
  var PENDING_PREFIX = "visualtex:pending:v1:";
  var SHAPE_PREFIX = "VisualTeX_";
  var TAG_FORMULA_ID = "VisualTeXFormulaId";
  var TAG_DOCUMENT = "VisualTeXDocumentId";
  var TAG_OBJECT = "VisualTeXObjectId";
  var TAG_SESSION = "VisualTeXSessionId";
  var TAG_PENDING = "VisualTeXPending";
  var TAG_METADATA = "VisualTeXMetadata";
  var TAG_REFERENCE = "VisualTeXFormulaReference";
  var TAG_FONT_SIZE = "VisualTeXFontSizePt";
  var TAG_REFERENCE_WIDTH = "VisualTeXReferenceWidthPt";
  var TAG_REFERENCE_HEIGHT = "VisualTeXReferenceHeightPt";
  var PLACEHOLDER_WIDTH = 180;
  var PLACEHOLDER_HEIGHT = 42;
  var REFERENCE_FONT_SIZE = 14;
  var DEFAULT_FONT_SIZE = 18;
  var POLL_SECONDS = 20;
  var EVENT_WAIT_SECONDS = 20;
  var PP_SELECTION_SHAPES = 2;
  var PP_SELECTION_TEXT = 3;
  var PP_ALIGN_CENTER = 2;
  var MSO_SHAPE_ROUNDED_RECTANGLE = 5;
  var MSO_TRUE = -1;
  var MSO_FALSE = 0;
  var MSO_BRING_TO_FRONT = 0;
  var MSO_SEND_TO_BACK = 1;
  var MSO_BRING_FORWARD = 2;
  var MSO_SEND_BACKWARD = 3;
  var GRAY = 128 + 128 * 256 + 128 * 65536;
  var NATIVE_TEXT_PREFIX = "visualtex-ppt-native-text:";
  var NATIVE_SLIDE_PREFIX = "visualtex-ppt-native-slide:";
  var STRUCTURE_START = "\u001a";
  var ARGUMENT_END = "\u001b";
  var MSO_TEXT_ORIENTATION_HORIZONTAL = 1;
  var PP_SAVE_AS_OPEN_XML_PRESENTATION = 24;
  var PP_AUTO_SIZE_SHAPE_TO_FIT_TEXT = 1;
  var STANDALONE_WIDTH = 240;
  var STANDALONE_HEIGHT = 90;
  var TABLE_CELLS_UNSUPPORTED = "WPS 演示的表格单元格中暂不支持 VisualTeX 公式。";
  var UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;

  var host = null;
  var api = null;
  var active = null;
  var lastDoubleClick = 0;
  /* WPS shows documents of every kind in one window: a double-click counts
     only while a presentation is the document in front. */
  var presentationInFront = true;
  var sizeRefresh = null;

  function app() {
    return wps.WppApplication();
  }

  function url(path) {
    var config = host.config;
    return "http://127.0.0.1:" + config.port + "/wps/v1" + path +
      (path.indexOf("?") >= 0 ? "&" : "?") + "token=" + encodeURIComponent(config.token || "");
  }

  function call(method, path, body, timeoutMs) {
    return new Promise(function (resolve, reject) {
      var xhr = new XMLHttpRequest();
      xhr.open(method, url(path), true);
      xhr.timeout = timeoutMs || 10000;
      if (body !== undefined) xhr.setRequestHeader("Content-Type", "application/json");
      xhr.onload = function () {
        var json = null;
        try { json = xhr.responseText ? JSON.parse(xhr.responseText) : null; } catch (error) { json = null; }
        if (xhr.status >= 200 && xhr.status < 300) resolve(json || {});
        else reject(new Error((json && json.error) || ("VisualTeX HTTP " + xhr.status)));
      };
      xhr.onerror = function () { reject(new Error("VisualTeX is not reachable")); };
      xhr.ontimeout = function () { reject(new Error("VisualTeX did not answer in time")); };
      xhr.send(body === undefined ? null : JSON.stringify(body));
    });
  }

  function log(message) {
    call("POST", "/log", { message: "wpp " + String(message) }, 3000).catch(function () {});
  }

  /* An alert, or the host's own notifier (a WPS alert is modal and stops
     every add-in until it is answered). */
  function notify(message) {
    /* An agent's command (or a formula it started) says it in its answer instead. */
    if (agentRun || (active && active.agent)) {
      if (agentRun) agentRun.notes.push(String(message));
      else log("agent notice: " + message);
      return;
    }
    if (host && typeof host.notify === "function") {
      host.notify(String(message));
      return;
    }
    try { alert(message); } catch (error) { /* no UI */ }
  }

  function safe(read, fallback) {
    try { return read(); } catch (error) { return fallback; }
  }

  function newUuid() {
    if (window.crypto && crypto.randomUUID) return crypto.randomUUID();
    var bytes = new Uint8Array(16);
    crypto.getRandomValues(bytes);
    bytes[6] = (bytes[6] & 0x0f) | 0x40;
    bytes[8] = (bytes[8] & 0x3f) | 0x80;
    var hex = Array.prototype.map.call(bytes, function (b) { return (b + 0x100).toString(16).slice(1); }).join("");
    return hex.slice(0, 8) + "-" + hex.slice(8, 12) + "-" + hex.slice(12, 16) + "-" + hex.slice(16, 20) + "-" + hex.slice(20);
  }

  function isUuid(value) {
    return UUID.test(String(value || ""));
  }

  /* ---- presentations, slides and shapes ---- */

  /* VTPresentationIdentityFor and VTBoundedIdentity. */
  function presentationIdentity(presentation) {
    var identity = String(safe(function () { return presentation.FullName; }, "") || "") ||
      String(safe(function () { return presentation.Name; }, "") || "");
    identity = identity.replace(/[\r\n]/g, " ").replace(/\u0000/g, "");
    return identity.length > 2048 ? identity.slice(0, 2048) : identity;
  }

  function findPresentation(identity) {
    var presentations = app().Presentations;
    var match = null;
    var count = 0;
    for (var index = 1; index <= presentations.Count; index++) {
      var candidate = presentations.Item(index);
      if (presentationIdentity(candidate) === identity) {
        match = candidate;
        count++;
      }
    }
    if (count !== 1) throw new Error("这个公式所在的演示文稿已经关闭，或者打开了多个同名演示文稿");
    return match;
  }

  /* VTRequireWritablePowerPointPresentationObject. */
  function requireWritable(presentation) {
    if (!presentation) throw new Error("请先打开一个 WPS 演示文稿。");
    if (Number(safe(function () { return presentation.ReadOnly; }, MSO_FALSE)) === MSO_TRUE) {
      throw new Error("这个演示文稿是只读的。");
    }
    if (Number(safe(function () { return presentation.Windows.Count; }, 1)) === 0) {
      throw new Error("这个演示文稿没有可编辑的窗口。");
    }
  }

  function activePresentation() {
    var presentation = safe(function () { return app().ActivePresentation; }, null);
    if (!presentation || safe(function () { return app().Presentations.Count; }, 0) === 0) {
      throw new Error("请先打开一个 WPS 演示文稿。");
    }
    requireWritable(presentation);
    return presentation;
  }

  function currentSlide() {
    var slide = safe(function () { return app().ActiveWindow.View.Slide; }, null);
    if (!slide) throw new Error("请切换到普通视图并选中一张幻灯片。");
    return slide;
  }

  /* The slide the user moved since the editor opened is found by its id. */
  function slideOf(presentation, index, id) {
    var slides = presentation.Slides;
    if (index >= 1 && index <= slides.Count) {
      var slide = slides.Item(index);
      if (Number(slide.SlideID) === id) return slide;
    }
    for (var k = 1; k <= slides.Count; k++) {
      if (Number(slides.Item(k).SlideID) === id) return slides.Item(k);
    }
    throw new Error("公式所在的幻灯片已经不存在。");
  }

  function slideOfShape(shape) {
    var parent = safe(function () { return shape.Parent; }, null);
    if (parent && safe(function () { return parent.SlideID; }, 0)) return parent;
    return currentSlide();
  }

  function shapeIndex(slide, shape) {
    var id = Number(shape.Id);
    for (var index = 1; index <= slide.Shapes.Count; index++) {
      if (Number(slide.Shapes.Item(index).Id) === id) return index;
    }
    return 0;
  }

  function shapeByName(slide, name) {
    for (var index = 1; index <= slide.Shapes.Count; index++) {
      var candidate = slide.Shapes.Item(index);
      if (String(safe(function () { return candidate.Name; }, "")) === name) return candidate;
    }
    return null;
  }

  function tag(shape, name) {
    return String(safe(function () { return shape.Tags.Item(name); }, "") || "");
  }

  /* VTSetShapeTag. */
  function setTag(shape, name, value) {
    safe(function () { shape.Tags.Delete(name); });
    shape.Tags.Add(name, String(value));
    if (tag(shape, name) !== String(value)) {
      throw new Error("WPS 没有保存公式标记 " + name);
    }
  }

  function jsonNumber(value) {
    return String(Math.round(Number(value) * 1e6) / 1e6);
  }

  /* VTShapeFormulaId. */
  function shapeFormulaId(shape) {
    var value = tag(shape, TAG_FORMULA_ID);
    var name = String(safe(function () { return shape.Name; }, "") || "");
    if (!isUuid(value) && name.indexOf(SHAPE_PREFIX) === 0) value = name.slice(SHAPE_PREFIX.length);
    return isUuid(value) ? value : "";
  }

  /* VTShapeMetadata. */
  function shapeMetadata(shape) {
    var value = tag(shape, TAG_METADATA);
    if (value.indexOf(METADATA_PREFIX) !== 0) value = String(safe(function () { return shape.AlternativeText; }, "") || "");
    return value.indexOf(METADATA_PREFIX) === 0 ? value : "";
  }

  function isFormulaShape(shape) {
    return !!shape && !!shapeMetadata(shape) && tag(shape, TAG_PENDING) !== "1";
  }

  function selectedShapes() {
    var selection = safe(function () { return app().ActiveWindow.Selection; }, null);
    if (!selection || Number(safe(function () { return selection.Type; }, 0)) !== PP_SELECTION_SHAPES) return [];
    var range = selection.ShapeRange;
    var shapes = [];
    for (var index = 1; index <= range.Count; index++) shapes.push(range.Item(index));
    return shapes;
  }

  function selectedFormulaShapes() {
    return selectedShapes().filter(isFormulaShape);
  }

  /* The one formula the selection holds (VTSelectedSingleShape), or null. */
  function selectedFormula() {
    var shapes = selectedShapes();
    return shapes.length === 1 && isFormulaShape(shapes[0]) ? shapes[0] : null;
  }

  /* ---- point sizes (VTEnsurePowerPointFormulaScaleState) ---- */

  function validFontSize(size) {
    return size >= 1 && size <= 512;
  }

  function tagNumber(shape, name) {
    var value = Number(tag(shape, name));
    return isFinite(value) && Math.abs(value) <= 10000000 ? value : 0;
  }

  function writeScale(shape, scale) {
    if (!validFontSize(scale.size) || !(scale.width > 0 && scale.height > 0 && scale.width <= 10000 && scale.height <= 10000)) {
      throw new Error("公式的字号记录无效");
    }
    setTag(shape, TAG_FONT_SIZE, jsonNumber(scale.size));
    setTag(shape, TAG_REFERENCE_WIDTH, jsonNumber(scale.width));
    setTag(shape, TAG_REFERENCE_HEIGHT, jsonNumber(scale.height));
  }

  /* The formula's point size and its size at 14 pt; a formula resized by
     hand takes the point size its height gives. */
  function scaleOf(shape, write) {
    var size = tagNumber(shape, TAG_FONT_SIZE);
    var width = tagNumber(shape, TAG_REFERENCE_WIDTH);
    var height = tagNumber(shape, TAG_REFERENCE_HEIGHT);
    var changed = false;
    if (!validFontSize(size)) {
      size = DEFAULT_FONT_SIZE;
      changed = true;
    }
    if (!(width > 0 && height > 0)) {
      width = Number(shape.Width) * REFERENCE_FONT_SIZE / size;
      height = Number(shape.Height) * REFERENCE_FONT_SIZE / size;
      changed = true;
    } else {
      var observed = Number(shape.Height) / height * REFERENCE_FONT_SIZE;
      if (validFontSize(observed) && Math.abs(observed - size) > 0.05) {
        size = observed;
        changed = true;
      }
    }
    var scale = { size: size, width: width, height: height };
    if (changed && write) writeScale(shape, scale);
    return scale;
  }

  /* VTPreferredPowerPointFormulaFontSize: the size of the selected text or
     formula, else 18 pt. */
  function preferredFontSize() {
    var size = DEFAULT_FONT_SIZE;
    safe(function () {
      var selection = app().ActiveWindow.Selection;
      var type = Number(selection.Type);
      if (type === PP_SELECTION_TEXT) {
        size = Number(selection.TextRange.Font.Size);
      } else if (type === PP_SELECTION_SHAPES && selection.ShapeRange.Count === 1) {
        var shape = selection.ShapeRange.Item(1);
        if (isFormulaShape(shape)) {
          size = scaleOf(shape, false).size;
        } else if (Number(shape.HasTextFrame) === MSO_TRUE && Number(shape.TextFrame.HasText) === MSO_TRUE) {
          size = Number(shape.TextFrame.TextRange.Font.Size);
        }
      }
    });
    return validFontSize(size) ? size : DEFAULT_FONT_SIZE;
  }

  /* VTPowerPointGeometryJson. */
  function geometry(presentation, slide, shape, scale) {
    var index = shapeIndex(slide, shape);
    if (index <= 0) throw new Error("所选公式已经不在这张幻灯片上。");
    return {
      presentationIdentity: presentationIdentity(presentation),
      slideIndex: Number(slide.SlideIndex),
      slideId: Number(slide.SlideID),
      shapeIndex: index,
      shapeId: Number(shape.Id),
      shapeName: String(shape.Name),
      left: Number(shape.Left),
      top: Number(shape.Top),
      width: Number(shape.Width),
      height: Number(shape.Height),
      rotation: Number(safe(function () { return shape.Rotation; }, 0)) || 0,
      zOrder: Number(shape.ZOrderPosition),
      fontSizePt: scale && scale.size > 0 ? scale.size : null,
      referenceWidthPt: scale && scale.width > 0 ? scale.width : null,
      referenceHeightPt: scale && scale.height > 0 ? scale.height : null
    };
  }

  function objectIdentity(slide, shape) {
    return String(slide.SlideID) + ":" + String(shape.Id);
  }

  /* VTShouldForkPowerPointFormulaCopy: a pasted copy carries the original's
     tags but is a different shape, so it is edited as a new formula. */
  function shouldFork(presentation, slide, shape, formulaId) {
    var documentId = presentationIdentity(presentation);
    var objectId = objectIdentity(slide, shape);
    var ownerDocument = tag(shape, TAG_DOCUMENT);
    var ownerObject = tag(shape, TAG_OBJECT);
    if (ownerDocument && ownerDocument !== documentId) return true;
    if (ownerObject) return ownerObject !== objectId;
    /* Formulas made before the owner tags existed: the first one in slide
       order stays the original and gets the tags. */
    var canonical = null;
    var canonicalSlide = null;
    for (var s = 1; s <= presentation.Slides.Count && !canonical; s++) {
      var candidateSlide = presentation.Slides.Item(s);
      for (var k = 1; k <= candidateSlide.Shapes.Count; k++) {
        var candidate = candidateSlide.Shapes.Item(k);
        if (shapeFormulaId(candidate) === formulaId) {
          canonical = candidate;
          canonicalSlide = candidateSlide;
          break;
        }
      }
    }
    if (!canonical) return true;
    setTag(canonical, TAG_DOCUMENT, documentId);
    setTag(canonical, TAG_OBJECT, objectIdentity(canonicalSlide, canonical));
    return objectIdentity(canonicalSlide, canonical) !== objectId;
  }

  /* ---- sessions ---- */

  function startSession(spec) {
    if (active) {
      log("session=" + active.sessionId + " replaced by a new formula command");
      var previous = active;
      active = null;
      if (previous.mode === "create") safe(function () { removePlaceholder(previous); });
    }
    var sessionId = spec.sessionId;
    active = {
      sessionId: sessionId,
      mode: spec.mode,
      identity: spec.powerPoint.presentationIdentity,
      pendingMarker: spec.pendingMarker || null,
      /* A native equation: the text box and the text it had when the editor opened. */
      native: spec.native || null,
      agent: !!agentRun
    };
    var body = {
      sessionId: sessionId,
      host: "powerpoint",
      mode: spec.mode,
      displayMode: spec.displayMode || "block",
      nativeEquation: !!spec.native,
      sourceDocumentId: spec.powerPoint.presentationIdentity,
      powerPoint: spec.powerPoint,
      forkCopiedFormula: !!spec.forkCopiedFormula
    };
    ["formulaId", "sourceObjectId", "encodedMetadata", "nativeEditXml"].forEach(function (key) {
      if (spec[key]) body[key] = spec[key];
    });
    /* An agent's formula: VisualTeX gives the request back and draws the agent's LaTeX itself. */
    if (agentRun) body.agent = true;
    var posting = call("POST", "/sessions", body, 25000);
    if (agentRun) agentRun.started = posting;
    posting.then(function () {
      poll(sessionId, 0);
    }, function (error) {
      if (active && active.sessionId === sessionId) {
        var current = active;
        active = null;
        if (current.mode === "create") safe(function () { removePlaceholder(current); });
      }
      notify("VisualTeX 无法打开公式编辑器：" + error.message);
    });
  }

  function poll(sessionId, failures) {
    if (!active || active.sessionId !== sessionId) return;
    call("GET", "/sessions/" + sessionId + "?wait=" + POLL_SECONDS, undefined, (POLL_SECONDS + 10) * 1000).then(function (state) {
      if (!active || active.sessionId !== sessionId) return;
      if (state.status === "editing") {
        poll(sessionId, 0);
      } else if (state.status === "ready" && active.document) {
        applyDocumentResult(sessionId, state.payload || {});
      } else if (state.status === "ready") {
        applyResult(sessionId, state.payload || {});
      } else {
        log("session=" + sessionId + " ended status=" + state.status);
        finishActive(state.status !== "done" && state.status !== "applied");
      }
    }, function (error) {
      if (!active || active.sessionId !== sessionId) return;
      if (failures >= 20) {
        log("session=" + sessionId + " lost VisualTeX: " + error.message);
        finishActive(true);
        return;
      }
      setTimeout(function () { poll(sessionId, failures + 1); }, 1500);
    });
  }

  /* A new formula that was not written takes its placeholder with it
     (VTCancelPowerPointDispatch). */
  function finishActive(removePending) {
    var current = active;
    active = null;
    if (current && removePending && current.mode === "create") safe(function () { removePlaceholder(current); });
  }

  function removePlaceholder(session) {
    if (!session.pendingMarker) return;
    var presentations = app().Presentations;
    for (var p = 1; p <= presentations.Count; p++) {
      var slides = presentations.Item(p).Slides;
      for (var s = 1; s <= slides.Count; s++) {
        var shapes = slides.Item(s).Shapes;
        for (var k = shapes.Count; k >= 1; k--) {
          var shape = shapes.Item(k);
          if (String(safe(function () { return shape.AlternativeText; }, "")) === session.pendingMarker &&
              tag(shape, TAG_SESSION) === session.sessionId && tag(shape, TAG_PENDING) === "1") {
            shape.Delete();
            return;
          }
        }
      }
    }
  }

  function download(sessionId, name) {
    var path = wps.OAAssist.DownloadFile(url("/sessions/" + sessionId + "/files/" + encodeURIComponent(name)));
    if (!path || !safe(function () { return wps.FileSystem.Exists(path); }, true)) {
      throw new Error("WPS could not download " + name + " from VisualTeX");
    }
    return path;
  }

  function removeFile(path) {
    if (path) safe(function () { wps.FileSystem.Remove(path); });
  }

  function requireNumber(payload, key, positive) {
    var value = Number(payload[key]);
    if (payload[key] === undefined || payload[key] === "" || !isFinite(value) || Math.abs(value) > 10000000 || (positive && value <= 0)) {
      throw new Error("VisualTeX 返回的公式数据无效（" + key + "）");
    }
    return value;
  }

  /* VTRestoreZOrder. */
  function restoreZOrder(shape, position) {
    var current = Number(shape.ZOrderPosition);
    if (position <= 0 || current === position) return;
    var count = Number(safe(function () { return shape.Parent.Shapes.Count; }, current)) || current;
    if (position > count) position = count;
    var guard = 0;
    if (current > position) {
      if (position < current - position) {
        shape.ZOrder(MSO_SEND_TO_BACK);
        while (Number(shape.ZOrderPosition) < position && guard++ < 8192) shape.ZOrder(MSO_BRING_FORWARD);
      } else {
        while (Number(shape.ZOrderPosition) > position && guard++ < 8192) shape.ZOrder(MSO_SEND_BACKWARD);
      }
    } else if (count - position + 1 < position - current) {
      shape.ZOrder(MSO_BRING_TO_FRONT);
      while (Number(shape.ZOrderPosition) > position && guard++ < 8192) shape.ZOrder(MSO_SEND_BACKWARD);
    } else {
      while (Number(shape.ZOrderPosition) < position && guard++ < 8192) shape.ZOrder(MSO_BRING_FORWARD);
    }
    if (Number(shape.ZOrderPosition) !== position) throw new Error("WPS 无法恢复公式的叠放次序");
  }

  /* An Apply that already wrote this formula (the editor retried) is done. */
  function alreadyWritten(slide, payload, sessionId) {
    var existing = shapeByName(slide, payload.shapeName);
    return !!existing && tag(existing, TAG_SESSION) === sessionId && tag(existing, TAG_PENDING) === "0" &&
      shapeMetadata(existing) === payload.metadata;
  }

  /* VTFinalizePowerPointDispatch: the new picture takes the original shape's
     place, size and z-order, with the formula's identity; the original (or
     the placeholder of a new formula) goes last, so a failure leaves the
     slide as it was. */
  function writeFormula(sessionId, payload) {
    var formulaId = String(payload.formulaId || "");
    var metadata = String(payload.metadata || "");
    var shapeName = String(payload.shapeName || "");
    var sourceShapeName = String(payload.sourceShapeName || "");
    var reference = String(payload.formulaReference || "");
    if (!isUuid(formulaId) || shapeName !== SHAPE_PREFIX + formulaId || metadata.indexOf(METADATA_PREFIX) !== 0 ||
        reference.indexOf(FORMULA_REF_PREFIX + formulaId + ":") !== 0 || !sourceShapeName) {
      throw new Error("VisualTeX 返回的公式数据无效");
    }
    var left = requireNumber(payload, "targetLeft");
    var top = requireNumber(payload, "targetTop");
    var width = requireNumber(payload, "targetWidth", true);
    var height = requireNumber(payload, "targetHeight", true);
    var rotation = requireNumber(payload, "rotation");
    var zOrder = requireNumber(payload, "zOrder", true);
    var scale = {
      size: requireNumber(payload, "fontSizePt", true),
      width: requireNumber(payload, "referenceWidthPt", true),
      height: requireNumber(payload, "referenceHeightPt", true)
    };
    var presentation = findPresentation(String(payload.presentationIdentity || ""));
    requireWritable(presentation);
    var slide = slideOf(presentation, requireNumber(payload, "slideIndex", true), requireNumber(payload, "slideId", true));
    if (alreadyWritten(slide, payload, sessionId)) return shapeByName(slide, shapeName);
    var original = shapeByName(slide, sourceShapeName);
    if (!original) throw new Error("要替换的公式已经不在幻灯片上。");

    var files = payload.files || {};
    var paths = [];
    var candidate = null;
    var originalRenamed = false;
    var token = sessionId.replace(/-/g, "").slice(0, 12);
    try {
      var imagePath = download(sessionId, files.image);
      paths.push(imagePath);
      /* SVG keeps the formula sharp; the PNG is for a WPS that refuses it. */
      candidate = safe(function () { return slide.Shapes.AddPicture(imagePath, MSO_FALSE, MSO_TRUE, left, top, width, height); }, null);
      if (!candidate && files.fallback) {
        var fallbackPath = download(sessionId, files.fallback);
        paths.push(fallbackPath);
        candidate = slide.Shapes.AddPicture(fallbackPath, MSO_FALSE, MSO_TRUE, left, top, width, height);
      }
      if (!candidate) throw new Error("WPS 无法插入公式图片");
      candidate.Name = "VisualTeXPendingResult_" + token;
      candidate.LockAspectRatio = MSO_TRUE;
      if (Math.abs(rotation) > 0.01) candidate.Rotation = rotation;
      candidate.AlternativeText = metadata;
      /* WPS keeps no Title; the tag below holds the reference instead. */
      safe(function () { candidate.Title = reference; });
      setTag(candidate, TAG_FORMULA_ID, formulaId);
      setTag(candidate, TAG_DOCUMENT, presentationIdentity(presentation));
      setTag(candidate, TAG_OBJECT, objectIdentity(slide, candidate));
      setTag(candidate, TAG_SESSION, sessionId);
      setTag(candidate, TAG_PENDING, "0");
      setTag(candidate, TAG_REFERENCE, reference);
      writeScale(candidate, scale);
      safe(function () { setTag(candidate, TAG_METADATA, metadata); });
      /* Above the original, which still holds the z-order: deleting the
         original moves the new picture into its place. */
      restoreZOrder(candidate, zOrder + 1);
      original.Name = "VisualTeXOriginal_" + token;
      originalRenamed = true;
      candidate.Name = shapeName;
      if (String(candidate.Name) !== shapeName || String(candidate.AlternativeText) !== metadata ||
          tag(candidate, TAG_FORMULA_ID) !== formulaId || tag(candidate, TAG_SESSION) !== sessionId) {
        throw new Error("WPS 没有保存公式的身份信息");
      }
      original.Delete();
      return candidate;
    } catch (error) {
      if (candidate) safe(function () { candidate.Delete(); });
      if (originalRenamed) safe(function () { original.Name = sourceShapeName; });
      throw error;
    } finally {
      paths.forEach(removeFile);
    }
  }

  function applyResult(sessionId, payload) {
    if (payload.nativeEquation) {
      applyNative(sessionId, payload);
      return;
    }
    var result = { ok: true };
    var written = null;
    try {
      safe(function () { app().StartNewUndoEntry(); });
      written = writeFormula(sessionId, payload);
    } catch (error) {
      result = { ok: false, error: "WPS could not write the formula: " + (error && error.message || error) };
    }
    if (written) safe(function () { written.Select(); });
    finishResult(sessionId, result);
  }

  function finishResult(sessionId, result) {
    if (!result.ok) log("session=" + sessionId + " apply-failed " + result.error);
    active = null;
    call("POST", "/sessions/" + sessionId + "/applied", result, 10000).catch(function () {});
    invalidateControl("vtxImageFontSize");
  }

  /* ---- native equations ----
     WPS's object model has no equations: its text spells each one inline
     (structures opened by U+001A, arguments closed by U+001B), and VisualTeX
     finds them in a copy WPS saves (bridge /presentation/equations). An
     equation is written by pasting WPS's own copy of it, from a windowless
     scratch presentation that imported the one-slide file VisualTeX made. */

  function mayContainMath(text) {
    return text.indexOf(STRUCTURE_START) >= 0 || text.indexOf("\ud835") >= 0 || text.indexOf("\u210e") >= 0;
  }

  /* Whether the character at a 1-based position is part of an equation's spelling. */
  function mathCharacter(text, position) {
    if (position < 1 || position > text.length) return false;
    var c = text.charAt(position - 1);
    return c === STRUCTURE_START || c === ARGUMENT_END || c === "\u210e" || c === "\ud835" ||
      (c >= "\udc00" && c <= "\udfff" && position >= 2 && text.charAt(position - 2) === "\ud835");
  }

  function mathNear(text, start, length) {
    if (mathCharacter(text, start) || mathCharacter(text, start - 1)) return true;
    for (var position = start; position < start + length; position++) {
      if (mathCharacter(text, position)) return true;
    }
    return false;
  }

  function isTable(shape) {
    return Number(safe(function () { return shape.HasTable; }, MSO_FALSE)) === MSO_TRUE;
  }

  function liveText(shape) {
    if (Number(safe(function () { return shape.HasTextFrame; }, MSO_FALSE)) !== MSO_TRUE) return "";
    return String(safe(function () { return shape.TextFrame.TextRange.Text; }, "") || "");
  }

  function characters(shape, start, length) {
    return shape.TextFrame.TextRange.Characters(start, length);
  }

  function shapeById(slide, id) {
    for (var index = 1; index <= slide.Shapes.Count; index++) {
      var candidate = slide.Shapes.Item(index);
      if (Number(safe(function () { return candidate.Id; }, 0)) === id) return candidate;
    }
    return null;
  }

  function readSize(range) {
    var size = Number(safe(function () { return range.Font.Size; }, 0));
    return validFontSize(size) ? size : 0;
  }

  /* The size an equation in text takes (VisualTeX for Windows'
     ResolveTextContextFontSize): the selected text's, else the character the
     cursor types into, else the one before it, else the text's. */
  function contextFontSize(shape, start, length) {
    var text = liveText(shape);
    var size = safe(function () { return readSize(characters(shape, start, length)); }, 0);
    [start, start - 1].forEach(function (position) {
      if (size || position < 1 || position > text.length || /[\r\n]/.test(text.charAt(position - 1))) return;
      size = safe(function () { return readSize(characters(shape, position, 1)); }, 0);
    });
    if (!size) size = safe(function () { return readSize(shape.TextFrame.TextRange); }, 0);
    return size || DEFAULT_FONT_SIZE;
  }

  function languageOf(range) {
    var id = Number(safe(function () { return range.LanguageID; }, 0));
    return id > 0 ? id : 0;
  }

  /* The language of the text at the cursor, else the presentation's, for the
     equation's runs (ResolveTextContextLanguage; the app maps it to a tag). */
  function contextLanguage(presentation, shape, start, length) {
    var id = 0;
    if (shape) {
      var text = liveText(shape);
      [start - 1, start + length].forEach(function (position) {
        if (id || position < 1 || position > text.length || /[\r\u000b]/.test(text.charAt(position - 1))) return;
        id = safe(function () { return languageOf(characters(shape, position, 1)); }, 0);
      });
      if (!id) id = safe(function () { return languageOf(shape.TextFrame.TextRange); }, 0);
    }
    if (!id) id = Number(safe(function () { return presentation.DefaultLanguageID; }, 0)) || 0;
    return id > 0 ? id : null;
  }

  function nativeTarget(presentation, slide, shape, start, length) {
    var target = geometry(presentation, slide, shape, { size: contextFontSize(shape, start, length) });
    target.nativeTextTarget = { rangeStart: start, rangeLength: length, leadingParagraph: false, trailingParagraph: false };
    target.languageId = contextLanguage(presentation, shape, start, length);
    return target;
  }

  /* VTNewPowerPointNativeEquation: at the text cursor (or over the selected
     text); with no text cursor, a new text box on the slide gets it. */
  function createNative(displayMode) {
    var presentation = activePresentation();
    var slide = currentSlide();
    var selection = safe(function () {
      var current = app().ActiveWindow.Selection;
      if (Number(current.Type) !== PP_SELECTION_TEXT) return null;
      return { shape: current.ShapeRange.Item(1), start: Number(current.TextRange.Start), length: Number(current.TextRange.Length) };
    }, null);
    if (selection) {
      var shape = selection.shape;
      if (isTable(shape)) throw new Error(TABLE_CELLS_UNSUPPORTED);
      startSession({
        sessionId: newUuid(),
        mode: "create",
        displayMode: displayMode,
        sourceObjectId: NATIVE_TEXT_PREFIX + slide.SlideID + ":" + shape.Id + ":" + selection.start + ":" + selection.length,
        powerPoint: nativeTarget(presentation, slide, shape, selection.start, selection.length),
        native: { shapeId: Number(shape.Id), text: liveText(shape) }
      });
      return;
    }
    startSession({
      sessionId: newUuid(),
      mode: "create",
      displayMode: displayMode,
      sourceObjectId: NATIVE_SLIDE_PREFIX + slide.SlideID,
      powerPoint: {
        presentationIdentity: presentationIdentity(presentation),
        slideIndex: Number(slide.SlideIndex),
        slideId: Number(slide.SlideID),
        shapeIndex: 0, shapeId: 0, shapeName: "",
        left: 0, top: 0, width: STANDALONE_WIDTH, height: STANDALONE_HEIGHT, rotation: 0, zOrder: 0,
        fontSizePt: preferredFontSize(),
        languageId: contextLanguage(presentation, null, 0, 0)
      },
      native: { standalone: true }
    });
  }

  /* Where the equations of a text box are: WPS saves a copy of the
     presentation and VisualTeX reads it, then deletes it. */
  function savedEquations(presentation, slide, shape, text) {
    var directory = String(safe(function () { return wps.Env.GetTempPath(); }, "") || "");
    if (!directory) return Promise.reject(new Error("WPS has no temporary folder"));
    if (directory.charAt(directory.length - 1) !== "/") directory += "/";
    var path = directory + "visualtex-wpp-" + newUuid() + ".pptx";
    try {
      presentation.SaveCopyAs(path, PP_SAVE_AS_OPEN_XML_PRESENTATION);
    } catch (error) {
      removeFile(path);
      return Promise.reject(new Error("WPS could not save a copy to read the equations from: " + (error && error.message || error)));
    }
    return call("POST", "/presentation/equations",
      { path: path, slideId: Number(slide.SlideID), shapeId: Number(shape.Id), text: text }, 60000).then(function (answer) {
      removeFile(path);
      return answer.zones || [];
    }, function (error) {
      removeFile(path);
      throw error;
    });
  }

  /* The one equation that holds the cursor, or that the selection overlaps. */
  function zoneAt(zones, start, length) {
    var found = -1;
    for (var index = 0; index < zones.length; index++) {
      var zone = zones[index];
      var hit = length === 0 ? start >= zone.start && start < zone.start + zone.length
        : start < zone.start + zone.length && start + length > zone.start;
      if (!hit) continue;
      if (found >= 0) return -1;
      found = index;
    }
    return found;
  }

  /* VTEditSelectedNativeEquation: the equation at the text cursor or
     selection, or the one equation of a selected text box. `quick` (a
     double-click): only when the cursor is at an equation's characters, so a
     double-click on words saves no copy. Resolves true when an editor opens. */
  function editNative(explicit, quick, x, y) {
    var presentation = activePresentation();
    var selection = app().ActiveWindow.Selection;
    var type = Number(selection.Type);
    var shape = null;
    var start = 0;
    var length = 0;
    if (type === PP_SELECTION_TEXT) {
      shape = selection.ShapeRange.Item(1);
      start = Number(selection.TextRange.Start);
      length = Number(selection.TextRange.Length);
    } else if (type === PP_SELECTION_SHAPES && selection.ShapeRange.Count === 1) {
      shape = selection.ShapeRange.Item(1);
    }
    var text = shape && !isTable(shape) && !isFormulaShape(shape) ? liveText(shape) : "";
    var candidate = mayContainMath(text) && (!quick || (type === PP_SELECTION_TEXT && mathNear(text, start, length) && hits(shape, x, y)));
    if (!candidate) {
      if (explicit) notify(shape && isTable(shape) ? TABLE_CELLS_UNSUPPORTED : "请先选中一个 VisualTeX 公式，或把光标放在要编辑的原生公式中。");
      return Promise.resolve(false);
    }
    var slide = slideOfShape(shape);
    return savedEquations(presentation, slide, shape, text).then(function (zones) {
      var index = type === PP_SELECTION_TEXT ? zoneAt(zones, start, length) : zones.length === 1 ? 0 : -1;
      if (index < 0) {
        if (explicit) notify("请把光标放在要编辑的公式中，或选中只含一个公式的文本框。");
        return false;
      }
      if (liveText(shape) !== text) throw new Error("读取公式时文本框被修改，请重试。");
      var zone = zones[index];
      startSession({
        sessionId: newUuid(),
        mode: "edit",
        displayMode: zone.display ? "block" : "inline",
        sourceObjectId: NATIVE_TEXT_PREFIX + slide.SlideID + ":" + shape.Id + ":" + zone.start + ":" + zone.length,
        nativeEditXml: zone.zone,
        powerPoint: nativeTarget(presentation, slide, shape, zone.start, zone.length),
        native: { shapeId: Number(shape.Id), text: text }
      });
      return true;
    });
  }

  /* A display equation goes in its own paragraph: the breaks it needs before
     and after the cursor, one per assignment (WPS turns only the last CR of
     an assigned string into a paragraph break). */
  function writeBreaks(shape, start, length) {
    var text = liveText(shape);
    var before = start > 1 && text.charAt(start - 2) !== "\r";
    var after = start - 1 + length < text.length && text.charAt(start - 1 + length) !== "\r";
    var count = (before ? 1 : 0) + (after ? 1 : 0);
    if (!count) return null;
    var listItem = Number(safe(function () {
      return characters(shape, Math.min(start, Math.max(1, text.length)), 1).ParagraphFormat.Bullet.Visible;
    }, MSO_FALSE)) === MSO_TRUE;
    for (var index = 0; index < count; index++) {
      characters(shape, start + index, index === 0 ? length : 0).Text = "\r";
    }
    return {
      start: start, count: count, insertion: start + (before ? 1 : 0), replaced: text.substr(start - 1, length), text: liveText(shape),
      before: before, after: after, listItem: listItem
    };
  }

  /* In a list item, the text after a display equation goes on with the same
     item instead of starting a new one. */
  function continueListItem(shape, breaks, pasted) {
    if (!breaks || !breaks.listItem || !breaks.before || !breaks.after) return;
    safe(function () { characters(shape, pasted.start + pasted.length + 1, 1).ParagraphFormat.Bullet.Visible = MSO_FALSE; });
  }

  /* After a failed paste the text is the original again. */
  function removeBreaks(shape, breaks) {
    if (liveText(shape) === breaks.text) characters(shape, breaks.start, breaks.count).Text = breaks.replaced;
  }

  function showSlide(presentation, slide) {
    safe(function () { presentation.Windows.Item(1).Activate(); });
    safe(function () { app().ActiveWindow.View.GotoSlide(Number(slide.SlideIndex)); });
  }

  /* Nothing is pasted unless WPS selected exactly the target. */
  function requireSelection(shape, start, length, empty) {
    var selection = app().ActiveWindow.Selection;
    var exact = Number(selection.Type) === PP_SELECTION_TEXT &&
      Number(selection.ShapeRange.Item(1).Id) === Number(shape.Id) &&
      (empty || (Number(selection.TextRange.Start) === start && Number(selection.TextRange.Length) === length));
    if (!exact) throw new Error("WPS did not select the exact text the equation replaces; nothing was pasted.");
  }

  /* Pastes the copied equation (spelled `copied`) over [start, start + length)
     with the window's own Paste (a TextRange.Paste in WPS appends at the end
     of the text); the text around the pasted range must be exactly the text
     around the target. Pasting the same spelling over itself (an edit that only
     restyled the equation) leaves the text as it was, so it is pasted once. */
  function pasteCopied(presentation, slide, shape, start, length, copied) {
    var original = liveText(shape);
    if (start < 1 || length < 0 || start - 1 + length > original.length) {
      throw new Error("The WPS text insertion position is no longer valid.");
    }
    showSlide(presentation, slide);
    var all = shape.TextFrame.TextRange;
    (original.length === 0 ? all : all.Characters(start, length)).Select();
    requireSelection(shape, start, length, original.length === 0);
    var prefix = original.slice(0, start - 1);
    var suffix = original.slice(start - 1 + length);
    var same = copied === original.substr(start - 1, length);
    for (var attempt = 1; attempt <= 3; attempt++) {
      safe(function () { app().ActiveWindow.View.Paste(); });
      var after = liveText(shape);
      if (after === original && !same) continue;
      var pasted = after.length - (original.length - length);
      if (pasted <= 0 || after.slice(0, prefix.length) !== prefix || after.slice(after.length - suffix.length) !== suffix) {
        throw new Error("WPS did not paste the equation at the expected text position.");
      }
      return { start: start, length: pasted };
    }
    throw new Error("WPS did not paste the equation.");
  }

  function pasteEquation(presentation, slide, shape, start, length, path) {
    var scratch = app().Presentations.Add(MSO_FALSE);
    try {
      scratch.Slides.InsertFromFile(path, 0, 1, 1);
      if (Number(scratch.Slides.Count) !== 1) throw new Error("WPS did not import the equation slide.");
      var shapes = scratch.Slides.Item(1).Shapes;
      if (Number(shapes.Count) !== 1) throw new Error("WPS imported the equation slide without its text box.");
      var text = shapes.Item(1).TextFrame.TextRange;
      if (!(Number(text.Length) > 0)) throw new Error("WPS imported the equation slide without its equation.");
      var copied = String(text.Text);
      text.Copy();
      return pasteCopied(presentation, slide, shape, start, length, copied);
    } finally {
      safe(function () { scratch.Saved = MSO_TRUE; });
      safe(function () { scratch.Close(); });
    }
  }

  function newEquationBox(presentation, slide) {
    var width = Number(presentation.PageSetup.SlideWidth);
    var height = Number(presentation.PageSetup.SlideHeight);
    var box = slide.Shapes.AddTextbox(MSO_TEXT_ORIENTATION_HORIZONTAL,
      (width - STANDALONE_WIDTH) / 2, (height - STANDALONE_HEIGHT) / 2, STANDALONE_WIDTH, STANDALONE_HEIGHT);
    safe(function () { box.TextFrame.WordWrap = MSO_FALSE; });
    safe(function () { box.TextFrame.AutoSize = PP_AUTO_SIZE_SHAPE_TO_FIT_TEXT; });
    return box;
  }

  /* WPS drops the equation's run sizes when it pastes into an empty box; a
     display equation is centred without a bullet; the cursor goes after it. */
  function finishNative(shape, pasted, size, display) {
    var range = characters(shape, pasted.start, pasted.length);
    safe(function () { if (Math.abs(Number(range.Font.Size) - size) > 0.01) range.Font.Size = size; });
    if (display) {
      safe(function () { range.ParagraphFormat.Alignment = PP_ALIGN_CENTER; });
      safe(function () { range.ParagraphFormat.Bullet.Visible = MSO_FALSE; });
    }
    safe(function () { characters(shape, pasted.start + pasted.length, 0).Select(); });
  }

  function writeNative(sessionId, payload) {
    var current = active;
    var presentation = findPresentation(String(payload.presentationIdentity || ""));
    requireWritable(presentation);
    var slide = slideOf(presentation, Number(payload.slideIndex), Number(payload.slideId));
    var size = validFontSize(Number(payload.fontSizePt)) ? Number(payload.fontSizePt) : DEFAULT_FONT_SIZE;
    var display = payload.displayMode === "block";
    var path = download(sessionId, (payload.files || {}).slide);
    var created = null;
    var breaks = null;
    var shape;
    try {
      var start = 1;
      var length = 0;
      if (payload.standalone) {
        shape = created = newEquationBox(presentation, slide);
      } else {
        shape = shapeById(slide, Number(payload.shapeId));
        if (!shape) throw new Error("公式所在的文本框已经不存在。");
        if (current && current.native && liveText(shape) !== current.native.text) {
          throw new Error("文本框在公式编辑期间被修改，没有写入公式。请在需要的位置重新开始。");
        }
        start = Number(payload.rangeStart);
        length = Number(payload.rangeLength);
        if (display && payload.mode === "create") breaks = writeBreaks(shape, start, length);
      }
      var pasted = pasteEquation(presentation, slide, shape,
        breaks ? breaks.insertion : start, breaks ? 0 : length, path);
      finishNative(shape, pasted, size, display);
      continueListItem(shape, breaks, pasted);
      if (created) {
        safe(function () {
          created.Left = (Number(presentation.PageSetup.SlideWidth) - Number(created.Width)) / 2;
          created.Top = (Number(presentation.PageSetup.SlideHeight) - Number(created.Height)) / 2;
        });
      }
      return { presentation: presentation, slide: slide, shape: shape, pasted: pasted };
    } catch (error) {
      if (breaks) safe(function () { removeBreaks(shape, breaks); });
      if (created) safe(function () { created.Delete(); });
      throw error;
    } finally {
      removeFile(path);
    }
  }

  /* The paste is proven when an equation covers the pasted text: math
     characters in it, or for an equation of digits and operators alone, a
     native equation in WPS's saved copy. */
  function confirmNative(written) {
    var text = liveText(written.shape);
    var pasted = written.pasted;
    if (mayContainMath(text.substr(pasted.start - 1, pasted.length))) return Promise.resolve();
    return savedEquations(written.presentation, written.slide, written.shape, text).then(function (zones) {
      var covered = zones.some(function (zone) {
        return zone.start < pasted.start + pasted.length && zone.start + zone.length > pasted.start;
      });
      if (!covered) throw new Error("WPS pasted the equation without making it a native equation.");
    });
  }

  function applyNative(sessionId, payload) {
    var fail = function (error) {
      finishResult(sessionId, { ok: false, error: "WPS could not write the equation: " + (error && error.message || error) });
    };
    var written;
    try {
      safe(function () { app().StartNewUndoEntry(); });
      written = writeNative(sessionId, payload);
    } catch (error) {
      fail(error);
      return;
    }
    confirmNative(written).then(function () { finishResult(sessionId, { ok: true }); }, fail);
  }

  /* ---- LaTeX redraw and formulas back to LaTeX ----
     As Word's "LaTeX 重绘" and VisualTeX for Windows' PowerPoint: the text of
     the text boxes in scope goes to VisualTeX, which finds and renders the
     LaTeX (presentation_redraw); the add-in writes each result in place, last
     first, so the earlier positions stay valid. Formulas go back to LaTeX the
     same way: an equation becomes $...$ or $$...$$ in its text, a picture
     formula a text box holding its LaTeX. */

  var PP_SELECTION_SLIDES = 1;
  var MSO_GROUP = 6;
  var DOCUMENT_LABELS = {
    latexRedraw: "VisualTeX LaTeX 重绘",
    formulaRestore: "VisualTeX 公式转为 LaTeX"
  };

  function confirmWithUser(message) {
    /* An agent asked for it: yes. */
    if (agentRun) return true;
    if (host && typeof host.confirm === "function") return host.confirm(String(message));
    return safe(function () { return window.confirm(message); }, false);
  }

  function slideById(presentation, id) {
    var slides = presentation.Slides;
    for (var index = 1; index <= slides.Count; index++) {
      if (Number(slides.Item(index).SlideID) === id) return slides.Item(index);
    }
    throw new Error("公式所在的幻灯片已经不存在。");
  }

  /* A text box, a placeholder, or a shape drawn as nothing but its text: only
     these become a picture formula (a filled or outlined shape keeps its LaTeX). */
  var MSO_AUTO_SHAPE = 1;
  var MSO_PLACEHOLDER = 14;
  var MSO_TEXT_BOX = 17;
  function isPlainTextBox(shape) {
    var type = Number(safe(function () { return shape.Type; }, 0));
    if (type === MSO_TEXT_BOX || type === MSO_PLACEHOLDER) return true;
    return type === MSO_AUTO_SHAPE &&
      Number(safe(function () { return shape.Fill.Visible; }, MSO_TRUE)) !== MSO_TRUE &&
      Number(safe(function () { return shape.Line.Visible; }, MSO_TRUE)) !== MSO_TRUE;
  }

  /* The text boxes VisualTeX reads: shapes with text, not tables, groups or
     VisualTeX picture formulas. */
  function isTextShape(shape) {
    return Number(safe(function () { return shape.Type; }, 0)) !== MSO_GROUP && !isTable(shape) && !isFormulaShape(shape) &&
      Number(safe(function () { return shape.HasTextFrame; }, MSO_FALSE)) === MSO_TRUE;
  }

  /* What a command works on: [{slide, shape, start, length, caret}], length 0
     for a whole shape; caret is the text cursor's position (0: none). The
     selected text, shapes or thumbnails ("selection"), the slide in view, or
     every slide. */
  function scopeTargets(presentation, scope) {
    var slides = [];
    if (scope === "presentation") {
      for (var p = 1; p <= presentation.Slides.Count; p++) slides.push(presentation.Slides.Item(p));
    } else if (scope === "selection") {
      var selection = app().ActiveWindow.Selection;
      var type = Number(safe(function () { return selection.Type; }, 0));
      if (type === PP_SELECTION_TEXT) {
        var textShape = selection.ShapeRange.Item(1);
        var start = Number(selection.TextRange.Start);
        var length = Number(selection.TextRange.Length);
        return [{ slide: slideOfShape(textShape), shape: textShape, start: length > 0 ? start : 0, length: length,
          caret: length > 0 ? 0 : start }];
      }
      if (type === PP_SELECTION_SHAPES) {
        var picked = [];
        for (var k = 1; k <= selection.ShapeRange.Count; k++) {
          var shape = selection.ShapeRange.Item(k);
          picked.push({ slide: slideOfShape(shape), shape: shape, start: 0, length: 0 });
        }
        return picked;
      }
      if (type === PP_SELECTION_SLIDES) {
        var range = selection.SlideRange;
        for (var r = 1; r <= range.Count; r++) slides.push(range.Item(r));
      }
    }
    if (!slides.length) slides.push(currentSlide());
    var targets = [];
    slides.forEach(function (slide) {
      for (var index = 1; index <= slide.Shapes.Count; index++) {
        targets.push({ slide: slide, shape: slide.Shapes.Item(index), start: 0, length: 0 });
      }
    });
    return targets;
  }

  /* The text's font sizes as [start, length, size] runs (0-based in `text`). */
  function sizeRuns(shape, start, length) {
    var range = safe(function () { return characters(shape, start, length); }, null);
    var runs = [];
    var count = Number(safe(function () { return range.Runs().Count; }, 0));
    for (var k = 1; range && k <= count && k <= 4000; k++) {
      var run = safe(function () { return range.Runs(k, 1); }, null);
      if (!run) break;
      var runStart = Number(safe(function () { return run.Start; }, 0)) - start;
      var runLength = Math.min(Number(safe(function () { return run.Length; }, 0)), length - runStart);
      var size = readSize(run);
      if (size && runStart >= 0 && runLength > 0) runs.push([runStart, runLength, size]);
    }
    if (!runs.length) runs.push([0, length, (range && readSize(range)) || contextFontSize(shape, start, 0)]);
    return runs;
  }

  function startDocumentSession(spec) {
    if (active) {
      log("session=" + active.sessionId + " replaced by a document operation");
      finishActive(true);
    }
    var sessionId = newUuid();
    active = {
      sessionId: sessionId,
      mode: "document",
      document: true,
      identity: presentationIdentity(spec.presentation),
      operation: spec.operation,
      /* The text each shape had, to check before writing. */
      texts: spec.texts,
      agent: !!agentRun
    };
    var body = {
      sessionId: sessionId,
      host: "powerpoint",
      operation: spec.operation,
      sourceDocumentId: active.identity,
      bookmarkName: "VT_D_" + sessionId.replace(/-/g, ""),
      defaultFontSizePt: DEFAULT_FONT_SIZE,
      redrawScope: spec.scope === "presentation" ? "document" : "selection",
      outputKind: spec.outputKind,
      source: JSON.stringify(spec.source)
    };
    if (spec.sourceKind) body.sourceKind = spec.sourceKind;
    /* An agent's operation: the window tells the agent what it skipped instead of showing it. */
    if (agentRun) body.agent = true;
    var posting = call("POST", "/documents", body, 60000);
    if (agentRun) agentRun.started = posting.then(function (reply) {
      reply = reply || {};
      reply.sessionId = reply.sessionId || sessionId;
      return reply;
    });
    posting.then(function () {
      poll(sessionId, 0);
    }, function (error) {
      if (active && active.sessionId === sessionId) active = null;
      notify("VisualTeX 无法开始这项操作：" + error.message);
    });
  }

  function shapeKey(slideId, shapeId) {
    return slideId + ":" + shapeId;
  }

  /* "LaTeX 重绘为 OMML / VisualTeX": the LaTeX in the text of the scope. */
  function redraw(outputKind, scope, targets) {
    var presentation = activePresentation();
    if (scope === "presentation" && !confirmWithUser("将在整个演示文稿中查找 LaTeX 公式并原位重绘，开始吗？")) return;
    var segments = [];
    var texts = {};
    /* an agent's insert_rich_text names the text box (and the text) it wrote: no selection to read */
    (targets || scopeTargets(presentation, scope)).forEach(function (target) {
      var shape = target.shape;
      if (!isTextShape(shape)) return;
      var text = liveText(shape);
      if (!text) return;
      var start = target.length > 0 ? target.start : 1;
      var length = target.length > 0 ? target.length : text.length;
      var slideId = Number(target.slide.SlideID);
      var shapeId = Number(shape.Id);
      texts[shapeKey(slideId, shapeId)] = text;
      segments.push({
        slideId: slideId,
        shapeId: shapeId,
        start: start,
        text: text.substr(start - 1, length),
        whole: start === 1 && length === text.length && isPlainTextBox(shape),
        sizes: sizeRuns(shape, start, length)
      });
    });
    if (!segments.length) {
      notify(scope === "selection" ? "所选内容中没有文字。" : "没有找到含文字的文本框。");
      return;
    }
    startDocumentSession({
      presentation: presentation, operation: "latexRedraw", outputKind: outputKind, scope: scope,
      source: { segments: segments }, texts: texts
    });
  }

  /* "转为 LaTeX 代码": the equations and picture formulas of the scope. The
     equations are found in a copy WPS saves, all shapes at once. */
  function restoreToLatex(scope) {
    var presentation = activePresentation();
    if (scope === "presentation" && !confirmWithUser("将把整个演示文稿中的公式转为 LaTeX 代码，开始吗？")) return;
    var pictures = [];
    var texts = [];
    scopeTargets(presentation, scope).forEach(function (target) {
      if (isFormulaShape(target.shape)) {
        pictures.push(target);
      } else if (isTextShape(target.shape)) {
        var text = liveText(target.shape);
        if (mayContainMath(text)) texts.push({ target: target, text: text });
      }
    });
    if (!pictures.length && !texts.length) {
      notify(scope === "selection" ? "所选内容中没有找到公式。" : "没有找到公式。");
      return;
    }
    var located = texts.length ? savedEquationsOf(presentation, texts) : Promise.resolve([]);
    located.then(function (answers) {
      var targets = [];
      var expected = {};
      var problems = [];
      texts.forEach(function (entry, index) {
        var answer = answers[index] || {};
        if (answer.error) {
          problems.push(answer.error);
          return;
        }
        var target = entry.target;
        var slideId = Number(target.slide.SlideID);
        var shapeId = Number(target.shape.Id);
        expected[shapeKey(slideId, shapeId)] = entry.text;
        var zones = answer.zones || [];
        /* The text cursor in an equation: that equation; elsewhere in the box:
           all of its equations. */
        var atCaret = zones.filter(function (zone) {
          return target.caret > 0 && target.caret >= zone.start && target.caret < zone.start + zone.length;
        });
        (atCaret.length ? atCaret : zones).forEach(function (zone) {
          if (target.length > 0 && !(zone.start < target.start + target.length && zone.start + zone.length > target.start)) return;
          targets.push({
            slideId: slideId, shapeId: shapeId, kind: "omml", start: zone.start, length: zone.length,
            displayMode: zone.display ? "block" : "inline",
            fontSizePt: contextFontSize(target.shape, zone.start, zone.length),
            zone: zone.zone
          });
        });
      });
      pictures.forEach(function (target) {
        targets.push({
          slideId: Number(target.slide.SlideID), shapeId: Number(target.shape.Id), kind: "image",
          displayMode: "block", fontSizePt: scaleOf(target.shape, false).size, metadata: shapeMetadata(target.shape)
        });
      });
      if (!targets.length) {
        notify(problems.length ? "无法读取公式：" + problems[0] : "没有找到公式。");
        return;
      }
      var kinds = targets.map(function (target) { return target.kind; });
      startDocumentSession({
        presentation: presentation, operation: "formulaRestore", outputKind: "latex", scope: scope,
        sourceKind: kinds.every(function (kind) { return kind === "omml"; }) ? "omml"
          : kinds.every(function (kind) { return kind === "image"; }) ? "image" : null,
        source: { targets: targets }, texts: expected
      });
      if (problems.length) log("restore skipped shapes: " + problems.join(" | "));
    }, function (error) {
      notify("VisualTeX 无法读取公式：" + (error && error.message || error));
    });
  }

  /* The equations of several text boxes from one saved copy. */
  function savedEquationsOf(presentation, entries) {
    var directory = String(safe(function () { return wps.Env.GetTempPath(); }, "") || "");
    if (!directory) return Promise.reject(new Error("WPS has no temporary folder"));
    if (directory.charAt(directory.length - 1) !== "/") directory += "/";
    var path = directory + "visualtex-wpp-" + newUuid() + ".pptx";
    try {
      presentation.SaveCopyAs(path, PP_SAVE_AS_OPEN_XML_PRESENTATION);
    } catch (error) {
      removeFile(path);
      return Promise.reject(new Error("WPS could not save a copy to read the equations from: " + (error && error.message || error)));
    }
    var shapes = entries.map(function (entry) {
      return { slideId: Number(entry.target.slide.SlideID), shapeId: Number(entry.target.shape.Id), text: entry.text };
    });
    return call("POST", "/presentation/equations", { path: path, shapes: shapes }, 120000).then(function (answer) {
      removeFile(path);
      return answer.shapes || [];
    }, function (error) {
      removeFile(path);
      throw error;
    });
  }

  /* The equations of a redraw come in one presentation, one text box each. */
  function openEquationScratch(path) {
    var scratch = app().Presentations.Add(MSO_FALSE);
    try {
      scratch.Slides.InsertFromFile(path, 0, 1, 1);
      if (Number(scratch.Slides.Count) !== 1) throw new Error("WPS did not import the equation slide.");
    } catch (error) {
      safe(function () { scratch.Saved = MSO_TRUE; });
      safe(function () { scratch.Close(); });
      throw error;
    }
    return scratch;
  }

  function scratchBox(scratch, number) {
    var shapes = scratch.Slides.Item(1).Shapes;
    var name = "VisualTeX Equation " + number;
    for (var index = 1; index <= shapes.Count; index++) {
      if (String(safe(function () { return shapes.Item(index).Name; }, "")) === name) return shapes.Item(index);
    }
    if (number <= shapes.Count) return shapes.Item(number);
    throw new Error("WPS imported the equation slide without equation " + number + ".");
  }

  /* A picture formula in place of the text box that held its LaTeX: the
     box's centre and z-order, the formula's identity as writeFormula writes it. */
  function replaceBoxWithPicture(presentation, slide, box, item, sessionId) {
    var files = item.files || {};
    var paths = [];
    var picture = null;
    var centerX = Number(box.Left) + Number(box.Width) / 2;
    var centerY = Number(box.Top) + Number(box.Height) / 2;
    var width = Number(item.widthPt);
    var height = Number(item.heightPt);
    try {
      var imagePath = download(sessionId, files.image);
      paths.push(imagePath);
      picture = safe(function () {
        return slide.Shapes.AddPicture(imagePath, MSO_FALSE, MSO_TRUE, centerX - width / 2, centerY - height / 2, width, height);
      }, null);
      if (!picture && files.fallback) {
        var fallbackPath = download(sessionId, files.fallback);
        paths.push(fallbackPath);
        picture = slide.Shapes.AddPicture(fallbackPath, MSO_FALSE, MSO_TRUE, centerX - width / 2, centerY - height / 2, width, height);
      }
      if (!picture) throw new Error("WPS 无法插入公式图片");
      picture.LockAspectRatio = MSO_TRUE;
      picture.Name = SHAPE_PREFIX + item.formulaId;
      picture.AlternativeText = item.metadata;
      safe(function () { picture.Title = item.formulaReference; });
      setTag(picture, TAG_FORMULA_ID, item.formulaId);
      setTag(picture, TAG_DOCUMENT, presentationIdentity(presentation));
      setTag(picture, TAG_OBJECT, objectIdentity(slide, picture));
      setTag(picture, TAG_SESSION, sessionId);
      setTag(picture, TAG_PENDING, "0");
      setTag(picture, TAG_REFERENCE, item.formulaReference);
      writeScale(picture, { size: Number(item.fontSizePt), width: Number(item.referenceWidthPt), height: Number(item.referenceHeightPt) });
      safe(function () { setTag(picture, TAG_METADATA, item.metadata); });
      restoreZOrder(picture, Number(box.ZOrderPosition) + 1);
      box.Delete();
      return picture;
    } catch (error) {
      if (picture) safe(function () { picture.Delete(); });
      throw error;
    } finally {
      paths.forEach(removeFile);
    }
  }

  /* LaTeX text in place of a picture formula: a text box at its centre. */
  function replacePictureWithText(slide, picture, item) {
    var centerX = Number(picture.Left) + Number(picture.Width) / 2;
    var centerY = Number(picture.Top) + Number(picture.Height) / 2;
    var z = Number(picture.ZOrderPosition);
    var box = slide.Shapes.AddTextbox(MSO_TEXT_ORIENTATION_HORIZONTAL, centerX - 120, centerY - 20, 240, 40);
    try {
      safe(function () { box.TextFrame.WordWrap = MSO_FALSE; });
      safe(function () { box.TextFrame.AutoSize = PP_AUTO_SIZE_SHAPE_TO_FIT_TEXT; });
      var text = box.TextFrame.TextRange;
      text.Text = String(item.text);
      safe(function () { text.Font.Size = Number(item.fontSizePt); });
      safe(function () {
        box.Left = centerX - Number(box.Width) / 2;
        box.Top = centerY - Number(box.Height) / 2;
      });
      restoreZOrder(box, z + 1);
      picture.Delete();
      return box;
    } catch (error) {
      safe(function () { box.Delete(); });
      throw error;
    }
  }

  function writeDocumentItem(presentation, item, scratch, sessionId, current) {
    var slide = slideById(presentation, Number(item.slideId));
    var shape = shapeById(slide, Number(item.shapeId));
    if (!shape) throw new Error("文本框已经不存在。");
    var key = shapeKey(Number(item.slideId), Number(item.shapeId));
    var start = Number(item.start);
    var length = Number(item.length);
    if (item.kind === "picture") {
      if (!isPlainTextBox(shape)) throw new Error("只有只含这个公式的文本框才会换成图片公式。");
      if (liveText(shape).trim() !== String(item.sourceText)) throw new Error("文本框在操作期间被修改。");
      replaceBoxWithPicture(presentation, slide, shape, item, sessionId);
      return;
    }
    if (item.kind === "text" && item.replace === "picture") {
      if (!isFormulaShape(shape)) throw new Error("公式已经不存在。");
      replacePictureWithText(slide, shape, item);
      return;
    }
    /* Text: the shape's text must still be what VisualTeX read, apart from
       the items after this one, already written. */
    var text = liveText(shape);
    var original = current.texts && current.texts[key];
    if (typeof original === "string") {
      var head = original.slice(0, start - 1 + length);
      if (text.slice(0, head.length) !== head) throw new Error("文本框在操作期间被修改。");
    }
    if (item.kind === "text") {
      characters(shape, start, length).Text = String(item.text);
      return;
    }
    if (text.substr(start - 1, length) !== String(item.sourceText)) throw new Error("文本框在操作期间被修改。");
    var display = item.displayMode === "block";
    var breaks = display ? writeBreaks(shape, start, length) : null;
    try {
      var box = scratchBox(scratch, Number(item.box)).TextFrame.TextRange;
      var copied = String(box.Text);
      box.Copy();
      var pasted = pasteCopied(presentation, slide, shape, breaks ? breaks.insertion : start, breaks ? 0 : length, copied);
      finishNative(shape, pasted, validFontSize(Number(item.fontSizePt)) ? Number(item.fontSizePt) : DEFAULT_FONT_SIZE, display);
      continueListItem(shape, breaks, pasted);
    } catch (error) {
      if (breaks) safe(function () { removeBreaks(shape, breaks); });
      throw error;
    }
  }

  function applyDocumentResult(sessionId, payload) {
    var current = active;
    var presentation = safe(function () { return findPresentation(current.identity); }, null);
    if (!presentation) {
      finishResult(sessionId, { ok: false, error: "WPS 中已找不到开始这项操作的演示文稿" });
      return;
    }
    var items = (payload.items || []).slice().sort(function (left, right) { return right.index - left.index; });
    var total = items.length;
    var written = 0;
    var failures = [];
    var scratch = null;
    var equationsPath = null;
    var reported = 0;
    try {
      safe(function () { app().StartNewUndoEntry(); });
      if (payload.equations) {
        equationsPath = download(sessionId, payload.equations);
        scratch = openEquationScratch(equationsPath);
      }
      items.forEach(function (item, done) {
        try {
          writeDocumentItem(presentation, item, scratch, sessionId, current);
          written++;
        } catch (error) {
          failures.push("第 " + (Number(item.index) + 1) + " 个：" + (error && error.message || error));
        }
        if (done + 1 - reported >= Math.max(1, Math.floor(total / 10)) || done + 1 === total) {
          reported = done + 1;
          call("POST", "/sessions/" + sessionId + "/progress", { current: done + 1, total: total }, 3000).catch(function () {});
        }
      });
    } catch (error) {
      failures.push(String(error && error.message || error));
    } finally {
      if (scratch) {
        safe(function () { scratch.Saved = MSO_TRUE; });
        safe(function () { scratch.Close(); });
      }
      removeFile(equationsPath);
    }
    safe(function () { app().StatusBar = "VisualTeX: 已写入 " + written + " 项"; });
    if (!written) {
      finishResult(sessionId, { ok: false, error: "WPS 未能写入：" + (failures[0] || "没有可写入的结果") });
      return;
    }
    if (failures.length) {
      notify("VisualTeX 已写入 " + written + " 项，" + failures.length + " 项未能写入：" + failures.slice(0, 3).join("；"));
    }
    finishResult(sessionId, { ok: true });
  }

  /* ---- commands ---- */

  /* VisualTeX_NewFormula: a placeholder in the middle of the slide marks
     where the formula goes until the editor applies it. */
  function create() {
    var presentation = activePresentation();
    var slide = currentSlide();
    var size = preferredFontSize();
    var sessionId = newUuid();
    var formulaId = newUuid();
    var marker = PENDING_PREFIX + sessionId + ":" + formulaId;
    var slideWidth = Number(presentation.PageSetup.SlideWidth);
    var slideHeight = Number(presentation.PageSetup.SlideHeight);
    var placeholder = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE,
      (slideWidth - PLACEHOLDER_WIDTH) / 2, (slideHeight - PLACEHOLDER_HEIGHT) / 2, PLACEHOLDER_WIDTH, PLACEHOLDER_HEIGHT);
    try {
      placeholder.Name = SHAPE_PREFIX + formulaId;
      placeholder.Fill.Visible = MSO_FALSE;
      placeholder.Line.Visible = MSO_TRUE;
      placeholder.Line.Weight = 1;
      placeholder.Line.ForeColor.RGB = GRAY;
      var text = placeholder.TextFrame.TextRange;
      text.Text = "VisualTeX";
      safe(function () { text.Font.Color.RGB = GRAY; });
      safe(function () { text.ParagraphFormat.Alignment = PP_ALIGN_CENTER; });
      setTag(placeholder, TAG_FORMULA_ID, formulaId);
      setTag(placeholder, TAG_DOCUMENT, presentationIdentity(presentation));
      setTag(placeholder, TAG_OBJECT, objectIdentity(slide, placeholder));
      setTag(placeholder, TAG_SESSION, sessionId);
      setTag(placeholder, TAG_PENDING, "1");
      setTag(placeholder, TAG_FONT_SIZE, jsonNumber(size));
      placeholder.AlternativeText = marker;
      startSession({
        sessionId: sessionId,
        mode: "create",
        formulaId: formulaId,
        sourceObjectId: String(placeholder.Name),
        pendingMarker: marker,
        powerPoint: geometry(presentation, slide, placeholder, { size: size })
      });
    } catch (error) {
      safe(function () { placeholder.Delete(); });
      throw error;
    }
  }

  /* VTPowerPointEditShape. */
  function editShape(shape) {
    var presentation = activePresentation();
    var slide = slideOfShape(shape);
    var metadata = shapeMetadata(shape);
    if (!metadata) throw new Error("请选择一个 VisualTeX 公式。");
    var formulaId = shapeFormulaId(shape);
    var scale = scaleOf(shape, true);
    var fork = !!formulaId && shouldFork(presentation, slide, shape, formulaId);
    if (fork) {
      formulaId = newUuid();
      shape.Name = SHAPE_PREFIX + formulaId;
    }
    startSession({
      sessionId: newUuid(),
      mode: "edit",
      /* A formula without an id (made by an older add-in) takes the one its
         metadata holds. */
      formulaId: formulaId || null,
      sourceObjectId: String(shape.Name),
      encodedMetadata: metadata,
      forkCopiedFormula: fork,
      powerPoint: geometry(presentation, slide, shape, scale)
    });
  }

  function editSelection() {
    var shape = selectedFormula();
    if (shape) {
      editShape(shape);
      return;
    }
    editNative(true, false).catch(function (error) {
      notify("VisualTeX 无法读取这个公式：" + (error && error.message || error));
    });
  }

  function deleteSelection() {
    var shapes = selectedFormulaShapes();
    if (!shapes.length) {
      notify("请先选中要删除的 VisualTeX 公式。");
      return;
    }
    shapes.forEach(function (shape) { shape.Delete(); });
  }

  function openVisualTeX() {
    call("POST", "/open", {}, 10000).catch(function (error) {
      notify("VisualTeX 无法打开主窗口：" + error.message);
    });
  }

  /* ---- formula point size (VTApplyPowerPointFormulaFontSize) ---- */

  var FONT_PRESETS = [
    [5, "八号"], [5.5, "七号"], [6.5, "小六"], [7.5, "六号"], [9, "小五"], [10.5, "五号"], [12, "小四"],
    [14, "四号"], [15, "小三"], [16, "三号"], [18, "小二"], [22, "二号"], [24, "小一"], [26, "一号"],
    [36, "小初"], [42, "初号"], [48, ""], [54, ""], [72, ""], [96, ""]
  ];

  function presetLabel(index) {
    var preset = FONT_PRESETS[index];
    return preset[1] ? preset[1] + " (" + preset[0] + " pt)" : preset[0] + " pt";
  }

  function presetIndex(size) {
    for (var index = 0; index < FONT_PRESETS.length; index++) {
      if (Math.abs(FONT_PRESETS[index][0] - size) <= 0.05) return index;
    }
    return -1;
  }

  function sizeLabel(size) {
    var index = presetIndex(size);
    return index >= 0 ? presetLabel(index) : (Math.round(size * 10) / 10) + " pt";
  }

  function invalidateControl(id) {
    safe(function () { if (host && host.ribbon) host.ribbon.InvalidateControl(id); });
  }

  /* Resized about its centre; rotation and z-order stay. */
  function resizeFormula(shape, size) {
    var scale = scaleOf(shape, true);
    var width = scale.width * size / REFERENCE_FONT_SIZE;
    var height = scale.height * size / REFERENCE_FONT_SIZE;
    if (!(width > 0 && height > 0 && width <= 10000 && height <= 10000)) {
      throw new Error("这个字号会让公式过大或过小");
    }
    var left = Number(shape.Left);
    var top = Number(shape.Top);
    var oldWidth = Number(shape.Width);
    var oldHeight = Number(shape.Height);
    var centerX = left + oldWidth / 2;
    var centerY = top + oldHeight / 2;
    try {
      shape.LockAspectRatio = MSO_FALSE;
      shape.Width = width;
      shape.Height = height;
      shape.Left = centerX - width / 2;
      shape.Top = centerY - height / 2;
      shape.LockAspectRatio = MSO_TRUE;
      writeScale(shape, { size: size, width: scale.width, height: scale.height });
    } catch (error) {
      safe(function () {
        shape.LockAspectRatio = MSO_FALSE;
        shape.Left = left;
        shape.Top = top;
        shape.Width = oldWidth;
        shape.Height = oldHeight;
        shape.LockAspectRatio = MSO_TRUE;
      });
      throw error;
    }
  }

  function resizeSelection(size) {
    var shapes = selectedFormulaShapes();
    if (!shapes.length) {
      notify("请先选中一个或多个 VisualTeX 公式，再选择字号。");
      return;
    }
    safe(function () { app().StartNewUndoEntry(); });
    shapes.forEach(function (shape) { resizeFormula(shape, size); });
    invalidateControl("vtxImageFontSize");
  }

  function fontSizeItemLabel(index) {
    if (index > 0) return presetLabel(index - 1);
    var shapes = safe(selectedFormulaShapes, []);
    if (!shapes.length) return "未选公式";
    var sizes = shapes.map(function (shape) { return scaleOf(shape, false).size; });
    var mixed = sizes.some(function (size) { return Math.abs(size - sizes[0]) > 0.05; });
    return mixed ? "混合字号" : sizeLabel(sizes[0]);
  }

  function fontSizeSelectedIndex() {
    var shapes = safe(selectedFormulaShapes, []);
    if (!shapes.length) return 0;
    var sizes = shapes.map(function (shape) { return scaleOf(shape, false).size; });
    if (sizes.some(function (size) { return Math.abs(size - sizes[0]) > 0.05; })) return 0;
    var index = presetIndex(sizes[0]);
    return index >= 0 ? index + 1 : 0;
  }

  /* The size list as a menu (WPS for Mac shows no value in a drop-down box):
     the current size first, then the sizes, the current one ticked. An item is
     "vtxSize_vtxImageFontSize_<index>" (the drop-down's index), see onAction. */
  function xmlAttribute(text) {
    return String(text).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  function fontSizeMenu() {
    var selected = safe(fontSizeSelectedIndex, 0);
    var current = safe(function () { return fontSizeItemLabel(selected); }, "");
    var xml = "<menu xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
      "<button id=\"vtxSizeCurrent\" label=\"" + xmlAttribute("当前：" + current) + "\" enabled=\"false\" />" +
      "<menuSeparator id=\"vtxSizeSeparator\" />";
    for (var index = 1; index <= FONT_PRESETS.length; index++) {
      xml += "<button id=\"vtxSize_vtxImageFontSize_" + index + "\" label=\"" +
        xmlAttribute((index === selected ? "✓ " : "") + presetLabel(index - 1)) + "\" onAction=\"OnVisualTeXAction\" />";
    }
    return xml + "</menu>";
  }

  /* ---- double-click ----
     WPS Presentation raises no double-click event, so VisualTeX's Office
     double-click monitor reports clicks in WPS (bridge events). A click opens
     the editor when it lands on the one selected formula. */

  /* Where a shape is on screen, in the screen points the monitor reports. */
  function screenBox(shape) {
    var view = app().ActiveWindow;
    var left = Number(shape.Left);
    var top = Number(shape.Top);
    var x1 = Number(view.PointsToScreenPixelsX(left));
    var y1 = Number(view.PointsToScreenPixelsY(top));
    var x2 = Number(view.PointsToScreenPixelsX(left + Number(shape.Width)));
    var y2 = Number(view.PointsToScreenPixelsY(top + Number(shape.Height)));
    if (![x1, y1, x2, y2].every(isFinite) || x2 <= x1 || y2 <= y1) return null;
    return { left: x1, top: y1, right: x2, bottom: y2 };
  }

  function hits(shape, x, y) {
    var box = safe(function () { return screenBox(shape); }, null);
    /* A WPS without screen positions: the selection alone decides. */
    if (!box) return true;
    var margin = 4;
    return x >= box.left - margin && x <= box.right + margin && y >= box.top - margin && y <= box.bottom + margin;
  }

  function onDoubleClick(x, y) {
    if (active || !presentationInFront) return;
    var now = Date.now();
    if (now - lastDoubleClick < 800) return;
    /* The first click of the pair selects the shape; give WPS a moment. */
    setTimeout(function () {
      if (active) return;
      var shape = safe(selectedFormula, null);
      if (shape) {
        if (!hits(shape, x, y)) return;
        lastDoubleClick = Date.now();
        run(function () { editShape(shape); });
        return;
      }
      /* A native equation: the double-click put the text cursor in it. */
      var opening = safe(function () { return editNative(false, true, x, y); }, null);
      if (opening) {
        lastDoubleClick = Date.now();
        opening.catch(function (error) { log("double-click native edit failed " + (error && error.message || error)); });
      }
    }, 120);
  }

  function listen(after, failures) {
    if (!host || host.module !== api) return;
    var path = after === null ? "/events" : "/events?after=" + after + "&wait=" + EVENT_WAIT_SECONDS;
    call("GET", path, undefined, (EVENT_WAIT_SECONDS + 10) * 1000).then(function (answer) {
      if (!host || host.module !== api) return;
      (answer.events || []).forEach(function (event) {
        if (event && event.kind === "doubleClick") onDoubleClick(Number(event.x), Number(event.y));
      });
      listen(Number(answer.sequence) || 0, 0);
    }, function () {
      if (!host || host.module !== api) return;
      setTimeout(function () { listen(null, failures + 1); }, Math.min(30000, 2000 * (failures + 1)));
    });
  }

  function run(action) {
    try {
      action();
    } catch (error) {
      log("command failed " + (error && error.stack || error));
      notify("VisualTeX 操作失败：" + (error && error.message || error));
    }
  }

  /* ---- AI agents (docs/office-mcp/MACOS.md) -----------------------------------
     VisualTeX hands the add-in an agent's operations over the WPS connection
     (office::wps::bridge agent_commands); each answers what the PowerPoint
     add-in's VisualTeX_AgentRun answers. A native equation is "ppt:<slide
     id>:<shape id>:<n>", the n-th equation of a text box, found in a copy WPS
     saves; its a14:m goes to the app as `zone`. A command that opens the
     formula editor hands its request back: VisualTeX draws the agent's LaTeX
     in its hidden editor and the add-in writes the result as after Apply. */

  var AGENT_POLL_SECONDS = 20;
  var agentRun = null;
  var agentLoopStarted = false;

  function agentFail(code, message) {
    var error = new Error(message);
    error.agentCode = code;
    throw error;
  }

  function base64UrlUtf8(text) {
    var bytes = new TextEncoder().encode(String(text));
    var binary = "";
    for (var index = 0; index < bytes.length; index += 0x8000) {
      binary += String.fromCharCode.apply(null, bytes.subarray(index, index + 0x8000));
    }
    return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  }

  function agentPresentation(args) {
    if (!safe(function () { return app().Presentations.Count; }, 0)) {
      agentFail("document_not_found", "No presentation is open in WPS Presentation.");
    }
    if (!args.document) return app().ActivePresentation;
    var presentations = app().Presentations;
    for (var index = 1; index <= presentations.Count; index++) {
      var candidate = presentations.Item(index);
      if (presentationIdentity(candidate) === args.document) return candidate;
    }
    agentFail("document_not_found", "The presentation is no longer open; call list_documents again.");
  }

  function agentWritable(args) {
    var presentation = agentPresentation(args);
    if (Number(safe(function () { return presentation.ReadOnly; }, MSO_FALSE)) === MSO_TRUE) {
      agentFail("document_read_only", "The presentation is read-only.");
    }
    if (presentationIdentity(app().ActivePresentation) !== presentationIdentity(presentation)) {
      if (!safe(function () { return presentation.Windows.Count; }, 0)) agentFail("operation_failed", "The presentation has no window.");
      presentation.Windows.Item(1).Activate();
    }
    return presentation;
  }

  function agentShowSlide(slide) {
    var pane = app().ActiveWindow;
    if (Number(safe(function () { return pane.View.Slide.SlideID; }, 0)) !== Number(slide.SlideID)) {
      pane.View.GotoSlide(Number(slide.SlideIndex));
    }
  }

  function round2(value) {
    return Math.round(Number(value) * 100) / 100;
  }

  function agentFrame(shape) {
    return {
      left_pt: round2(shape.Left), top_pt: round2(shape.Top), width_pt: round2(shape.Width), height_pt: round2(shape.Height),
      rotation: round2(safe(function () { return shape.Rotation; }, 0) || 0),
      z_order: Number(safe(function () { return shape.ZOrderPosition; }, 0)) || 0
    };
  }

  /* item: {slide, shape} (a picture formula) or {slide, shape, number, zone, text} (the number-th
     equation of a text box, as /presentation/equations found it). */
  function agentFormula(item, contextChars) {
    var slide = item.slide;
    var shape = item.shape;
    var formula = {
      numbered: false, slide: Number(slide.SlideIndex), slide_id: Number(slide.SlideID), shape_id: Number(shape.Id),
      shape_name: String(safe(function () { return shape.Name; }, "") || ""), frame: agentFrame(shape),
      context_before: "", context_after: ""
    };
    if (!item.zone) {
      var size = scaleOf(shape, false).size;
      formula.formula_id = "vtx:" + shapeFormulaId(shape);
      formula.kind = "picture";
      formula.display = "inline";
      formula.font_size_pt = size > 0 ? round2(size) : null;
      formula.equation = null;
      formula.frame_of = "formula";
      formula.payload_b64 = base64UrlUtf8(shapeMetadata(shape));
      formula.vtx_formula_id = shapeFormulaId(shape) || null;
      return formula;
    }
    var zone = item.zone;
    formula.formula_id = "ppt:" + slide.SlideID + ":" + shape.Id + ":" + item.number;
    formula.kind = "office_equation";
    formula.display = zone.display ? "block" : "inline";
    formula.font_size_pt = round2(contextFontSize(shape, zone.start, zone.length));
    formula.equation = item.number;
    formula.frame_of = "text_box";
    formula.payload_b64 = null;
    formula.zone = zone.zone;
    if (contextChars > 0) {
      var text = item.text;
      var from = Math.max(1, zone.start - contextChars);
      formula.context_before = text.substr(from - 1, zone.start - from).replace(/[\r\u000b]/g, "\n");
      formula.context_after = text.substr(zone.start - 1 + zone.length, contextChars).replace(/[\r\u000b]/g, "\n");
    }
    return formula;
  }

  /* Every formula of the slides, in slide and shape order: pictures, and the equations of the text boxes
     (one saved copy for all of them). */
  function agentFormulas(presentation, slides) {
    var shapes = [];
    var texts = [];
    slides.forEach(function (slide) {
      for (var index = 1; index <= slide.Shapes.Count; index++) {
        var shape = slide.Shapes.Item(index);
        if (isFormulaShape(shape)) {
          shapes.push({ slide: slide, shape: shape });
        } else if (isTextShape(shape)) {
          var text = liveText(shape);
          if (!mayContainMath(text)) continue;
          var entry = { slide: slide, shape: shape, text: text, zones: [] };
          shapes.push(entry);
          texts.push({ target: { slide: slide, shape: shape }, text: text, entry: entry });
        }
      }
    });
    var located = texts.length ? savedEquationsOf(presentation, texts) : Promise.resolve([]);
    return located.then(function (answers) {
      texts.forEach(function (text, index) {
        var answer = answers[index] || {};
        text.entry.zones = answer.zones || [];
        if (answer.error) log("agent: equations of shape " + text.entry.shape.Id + " unreadable: " + answer.error);
      });
      var items = [];
      shapes.forEach(function (entry) {
        if (!entry.text) {
          items.push(entry);
          return;
        }
        entry.zones.forEach(function (zone, index) {
          items.push({ slide: entry.slide, shape: entry.shape, text: entry.text, zone: zone, number: index + 1 });
        });
      });
      return items;
    });
  }

  function allSlides(presentation) {
    var slides = [];
    for (var index = 1; index <= presentation.Slides.Count; index++) slides.push(presentation.Slides.Item(index));
    return slides;
  }

  /* The formula an agent id names: "vtx:<id>" or "ppt:<slide id>:<shape id>:<n>". */
  function agentFind(presentation, id) {
    id = String(id || "");
    if (id.indexOf("vtx:") === 0) {
      var wanted = id.slice(4);
      var slides = allSlides(presentation);
      for (var s = 0; s < slides.length; s++) {
        for (var index = 1; index <= slides[s].Shapes.Count; index++) {
          var shape = slides[s].Shapes.Item(index);
          if (isFormulaShape(shape) && shapeFormulaId(shape) === wanted) return Promise.resolve({ slide: slides[s], shape: shape });
        }
      }
      agentFail("formula_not_found", "No VisualTeX formula in this presentation has that id.");
    }
    var parts = id.split(":");
    if (parts.length !== 4 || parts[0] !== "ppt") agentFail("formula_not_found", "This is not a formula id from list_formulas.");
    var slide = safe(function () { return slideById(presentation, Number(parts[1])); }, null);
    var box = slide && shapeById(slide, Number(parts[2]));
    var number = Number(parts[3]);
    if (!box || !isTextShape(box)) agentFail("stale_formula_id", "The presentation changed; call list_formulas again.");
    var text = liveText(box);
    return savedEquations(presentation, slide, box, text).then(function (zones) {
      if (!(number >= 1 && number <= zones.length) || liveText(box) !== text) {
        agentFail("stale_formula_id", "The presentation changed; call list_formulas again.");
      }
      return { slide: slide, shape: box, text: text, zone: zones[number - 1], number: number };
    });
  }

  /* The formula shown and selected (an equation's characters). */
  function agentSelect(item) {
    agentShowSlide(item.slide);
    if (item.zone) characters(item.shape, item.zone.start, item.zone.length).Select();
    else item.shape.Select();
  }

  function agentDone(message) {
    var parts = [];
    if (message) parts.push(message);
    if (agentRun) parts = parts.concat(agentRun.notes);
    return { message: parts.length ? parts.join(" ") : null };
  }

  /* The session the command started, once VisualTeX has it: its id and (for the hidden editor) its
     request. A command that did not start one says why in its notes. */
  function agentStarted(withRequest) {
    var run = agentRun;
    var waited = 0;
    return new Promise(function (resolve, reject) {
      var check = function () {
        if (run.started) {
          run.started.then(function (reply) {
            resolve({ session_id: reply.sessionId || null, request: withRequest ? reply.request || null : null,
              slide: Number(safe(function () { return app().ActiveWindow.View.Slide.SlideIndex; }, 0)) || null,
              message: run.notes.length ? run.notes.join(" ") : null });
          }, reject);
          return;
        }
        if (run.notes.length || waited >= 120000) {
          reject(new Error(run.notes.length ? run.notes.join(" ") : "WPS did not start the operation"));
          run.notes = [];
          return;
        }
        waited += 100;
        setTimeout(check, 100);
      };
      check();
    });
  }

  var SHAPE_TYPES = { 14: "placeholder", 17: "text_box", 13: "picture", 11: "picture", 6: "group", 19: "table", 3: "chart",
    24: "smart_art", 9: "line", 16: "media", 7: "ole_object", 10: "ole_object", 12: "ole_object", 1: "shape", 5: "shape" };
  var PLACEHOLDER_TYPES = { 1: "title", 3: "title", 5: "title", 4: "subtitle", 2: "body", 6: "body", 7: "content",
    18: "picture", 15: "footer", 13: "slide_number", 16: "date" };

  var AGENT_OPERATIONS = {
    list_documents: function () {
      var presentations = app().Presentations;
      var activeKey = safe(function () { return presentationIdentity(app().ActivePresentation); }, "");
      var list = [];
      for (var index = 1; index <= presentations.Count; index++) {
        var presentation = presentations.Item(index);
        var key = presentationIdentity(presentation);
        var path = String(safe(function () { return presentation.Path; }, "") || "");
        list.push({
          key: key, name: String(safe(function () { return presentation.Name; }, "") || ""),
          path: path ? String(safe(function () { return presentation.FullName; }, "") || "") || null : null,
          active: key === activeKey,
          saved: Number(safe(function () { return presentation.Saved; }, MSO_TRUE)) !== MSO_FALSE,
          read_only: Number(safe(function () { return presentation.ReadOnly; }, MSO_FALSE)) === MSO_TRUE,
          slide_count: Number(presentation.Slides.Count),
          slide_width_pt: round2(presentation.PageSetup.SlideWidth),
          slide_height_pt: round2(presentation.PageSetup.SlideHeight)
        });
      }
      return { documents: list };
    },

    list_formulas: function (args) {
      var presentation = agentPresentation(args);
      var offset = Number(args.offset) || 0;
      var limit = args.limit === undefined ? 50 : Number(args.limit);
      var contextChars = args.contextChars === undefined ? 40 : Number(args.contextChars);
      var kinds = args.kinds || [];
      return agentFormulas(presentation, allSlides(presentation)).then(function (items) {
        var listed = [];
        var revision = "";
        var total = 0;
        items.forEach(function (item) {
          var kind = item.zone ? "office_equation" : "picture";
          if (kinds.length && kinds.indexOf(kind) < 0) return;
          revision += item.slide.SlideID + "." + item.shape.Id + (item.zone ? "." + item.number + "@" + item.zone.start : "") + ";";
          if (total >= offset && listed.length < limit) listed.push(agentFormula(item, contextChars));
          total++;
        });
        return { revision: revision, total: total, offset: offset, formulas: listed };
      });
    },

    read_formula: function (args) {
      return agentFind(agentPresentation(args), args.formulaId).then(function (item) { return agentFormula(item, 120); });
    },

    find_vtx_formula: function (args) {
      var presentation = agentPresentation(args);
      return agentFind(presentation, "vtx:" + String(args.vtxId || "")).then(function (item) {
        return { formula: agentFormula(item, 0) };
      }, function (error) {
        if (error && error.agentCode === "formula_not_found") return { formula: null };
        throw error;
      });
    },

    get_slide_layout: function (args) {
      var presentation = agentPresentation(args);
      var number = Number(args.slide) || Number(safe(function () { return app().ActiveWindow.View.Slide.SlideIndex; }, 0)) || 1;
      if (!(number >= 1 && number <= presentation.Slides.Count)) {
        agentFail("invalid_arguments", "The presentation has " + presentation.Slides.Count + " slides.");
      }
      var slide = presentation.Slides.Item(number);
      return agentFormulas(presentation, [slide]).then(function (items) {
        var ids = {};
        items.forEach(function (item) {
          var key = String(item.shape.Id);
          (ids[key] = ids[key] || []).push(item.zone ? "ppt:" + slide.SlideID + ":" + item.shape.Id + ":" + item.number
            : "vtx:" + shapeFormulaId(item.shape));
        });
        var shapes = [];
        for (var index = 1; index <= slide.Shapes.Count; index++) {
          var shape = slide.Shapes.Item(index);
          var type = Number(safe(function () { return shape.Type; }, 0));
          var text = liveText(shape).replace(/[\r\u000b]/g, "\n");
          if (text.length > 200) text = text.slice(0, 200) + "…";
          shapes.push({
            shape_id: Number(shape.Id), name: String(safe(function () { return shape.Name; }, "") || ""),
            type: isFormulaShape(shape) ? "formula" : SHAPE_TYPES[type] || "other",
            placeholder: type === 14 ? PLACEHOLDER_TYPES[Number(safe(function () { return shape.PlaceholderFormat.Type; }, 0))] || "other" : null,
            frame: agentFrame(shape), text: text || null, formula_ids: ids[String(shape.Id)] || null
          });
        }
        return {
          slide: number, slide_id: Number(slide.SlideID), slide_count: Number(presentation.Slides.Count),
          layout: safe(function () { return String(slide.CustomLayout.Name); }, null) || null,
          slide_width_pt: round2(presentation.PageSetup.SlideWidth), slide_height_pt: round2(presentation.PageSetup.SlideHeight),
          shapes: shapes
        };
      });
    },

    get_selection: function (args) {
      var presentation = agentPresentation(args);
      var pane = app().ActiveWindow;
      var slide = safe(function () { return pane.View.Slide; }, null);
      var selection = safe(function () { return pane.Selection; }, null);
      var type = Number(safe(function () { return selection.Type; }, 0));
      var answer = {
        document: presentationIdentity(presentation), kind: type === PP_SELECTION_SLIDES ? "slide" : "none",
        formula: null, text_preview: null,
        location: { slide: slide ? Number(slide.SlideIndex) : null, slide_id: slide ? Number(slide.SlideID) : null, shape_id: null },
        read_only: Number(safe(function () { return presentation.ReadOnly; }, MSO_FALSE)) === MSO_TRUE
      };
      var shape = (type === PP_SELECTION_SHAPES || type === PP_SELECTION_TEXT) && selection.ShapeRange.Count === 1
        ? selection.ShapeRange.Item(1) : null;
      if (!shape || !slide) return answer;
      answer.kind = "shape";
      answer.location.shape_id = Number(shape.Id);
      if (isFormulaShape(shape)) {
        answer.kind = "formula";
        answer.formula = agentFormula({ slide: slide, shape: shape }, 0);
        return answer;
      }
      if (type === PP_SELECTION_TEXT) {
        answer.kind = "text";
        var preview = String(safe(function () { return selection.TextRange.Text; }, "") || "").replace(/[\r\u000b]/g, "\n");
        answer.text_preview = preview.length > 200 ? preview.slice(0, 200) + "…" : preview;
      }
      var text = isTextShape(shape) ? liveText(shape) : "";
      if (!mayContainMath(text)) return answer;
      var start = type === PP_SELECTION_TEXT ? Number(selection.TextRange.Start) : 0;
      var length = type === PP_SELECTION_TEXT ? Number(selection.TextRange.Length) : 0;
      return savedEquations(presentation, slide, shape, text).then(function (zones) {
        var index = type === PP_SELECTION_TEXT ? zoneAt(zones, start, length) : zones.length === 1 ? 0 : -1;
        if (index >= 0) {
          answer.kind = "formula";
          answer.text_preview = null;
          answer.formula = agentFormula({ slide: slide, shape: shape, text: text, zone: zones[index], number: index + 1 }, 60);
        }
        return answer;
      });
    },

    select_formula: function (args) {
      var presentation = agentPresentation(args);
      if (presentationIdentity(app().ActivePresentation) !== presentationIdentity(presentation)) {
        safe(function () { presentation.Windows.Item(1).Activate(); });
      }
      return agentFind(presentation, args.formulaId).then(function (item) {
        agentSelect(item);
        return { formula: agentFormula(item, 0) };
      });
    },

    /* A native equation alone in its text box goes with the box. */
    delete_formula: function (args) {
      return agentFind(agentWritable(args), args.formulaId).then(function (item) {
        safe(function () { app().StartNewUndoEntry(); });
        if (!item.zone || item.text.replace(/[\r\u000b\s]/g, "").length <= item.zone.length) item.shape.Delete();
        else characters(item.shape, item.zone.start, item.zone.length).Delete();
        return agentDone("Deleted the formula.");
      });
    },

    set_formula_font_size: function (args) {
      var size = Number(args.font_size_pt);
      if (!(size >= 5 && size <= 200)) agentFail("invalid_arguments", "font_size_pt must be from 5 to 200.");
      return agentFind(agentWritable(args), args.formulaId).then(function (item) {
        safe(function () { app().StartNewUndoEntry(); });
        if (item.zone) characters(item.shape, item.zone.start, item.zone.length).Font.Size = size;
        else resizeFormula(item.shape, size);
        return agentDone("The formula is " + size + " pt.");
      });
    },

    /* position: the new frame in points; what is left out stays. width or height alone keeps a
       formula's proportions (a text box only takes the size given). */
    set_frame: function (args) {
      var position = args.position || {};
      return agentFind(agentWritable(args), args.formulaId).then(function (item) {
        var shape = item.shape;
        var picture = !item.zone;
        var locked = shape.LockAspectRatio;
        var ratio = Number(shape.Height) > 0 ? Number(shape.Width) / Number(shape.Height) : 0;
        var has = function (name) { return typeof position[name] === "number" && isFinite(position[name]); };
        safe(function () { app().StartNewUndoEntry(); });
        shape.LockAspectRatio = MSO_FALSE;
        if (has("width_pt")) shape.Width = position.width_pt;
        if (has("height_pt")) shape.Height = position.height_pt;
        if (picture && ratio > 0 && has("width_pt") !== has("height_pt")) {
          if (has("width_pt")) shape.Height = position.width_pt / ratio;
          else shape.Width = position.height_pt * ratio;
        }
        if (has("left_pt")) shape.Left = position.left_pt;
        if (has("top_pt")) shape.Top = position.top_pt;
        shape.LockAspectRatio = locked;
        if (picture && (has("width_pt") || has("height_pt"))) {
          /* the point size follows the new height */
          safe(function () { scaleOf(shape, true); });
        }
        return { frame: agentFrame(shape), frame_of: picture ? "formula" : "text_box" };
      });
    },

    /* location: "cursor" (the text box's cursor, or a new formula on the current slide) or {slide, ...}
       (a new formula on that slide; the app moves it to left_pt/top_pt afterwards). */
    begin_insert: function (args) {
      var presentation = agentWritable(args);
      var location = args.location;
      if (location && typeof location === "object" && location.slide !== undefined && location.slide !== null) {
        var number = Number(location.slide);
        if (!(number >= 1 && number <= presentation.Slides.Count)) {
          agentFail("invalid_arguments", "The presentation has " + presentation.Slides.Count + " slides.");
        }
        agentShowSlide(presentation.Slides.Item(number));
        safe(function () { app().ActiveWindow.Selection.Unselect(); });
      } else if (location && typeof location !== "object" && location !== "cursor") {
        agentFail("invalid_arguments", "Unknown location " + location + ".");
      }
      if (args.format === "office_equation") createNative(args.display === "block" ? "block" : "inline");
      else create();
      return agentStarted(true);
    },

    begin_update: function (args) {
      return agentFind(agentWritable(args), args.formulaId).then(function (item) {
        agentSelect(item);
        if (!item.zone) {
          editShape(item.shape);
          return agentStarted(true);
        }
        return editNative(true, false).then(function (opened) {
          if (!opened && !agentRun.notes.length) agentFail("operation_failed", "WPS did not open the equation.");
          return agentStarted(true);
        });
      });
    },

    /* The ribbon's LaTeX redraw and formulas-to-LaTeX, which open VisualTeX's progress window; the app
       waits for the session they start. */
    latex_text_to_formulas: function (args) {
      var presentation = agentWritable(args);
      var targets = null;
      if (args.target && args.target.shapeId) {
        var slide = slideById(presentation, Number(args.target.slideId));
        var shape = shapeById(slide, Number(args.target.shapeId));
        if (!shape) agentFail("operation_failed", "The text box is no longer on the slide.");
        var length = Number(args.target.length) || 0;
        targets = [{ slide: slide, shape: shape, start: length > 0 ? Number(args.target.start) : 0, length: length }];
      }
      redraw(args.format === "office_equation" ? "omml" : "image", agentScope(args.scope), targets);
      return agentStarted(false);
    },

    /* insert_rich_text on a slide: the app's paragraphs [{text, kind, bold: [{start, length}]}] into a new text box,
       or at the text caret; the LaTeX redraw that follows gets the box (and the text) from the answer. */
    write_text_box: function (args) {
      var presentation = agentWritable(args);
      var paragraphs = args.paragraphs || [];
      if (!paragraphs.length) agentFail("invalid_arguments", "text is empty.");
      var location = args.location;
      var slideNumber = location && typeof location === "object" && location.slide ? Number(location.slide) : 0;
      var selection = safe(function () { return app().ActiveWindow.Selection; }, null);
      var box, slide, start = 0, length = 0, newBox = false;
      if (!slideNumber && selection && Number(safe(function () { return selection.Type; }, 0)) === PP_SELECTION_TEXT) {
        /* at the caret, a paragraph break at a time (a "\r" inside a longer insertion stays a character) */
        box = selection.ShapeRange.Item(1);
        slide = slideOfShape(box);
        var inserted = selection.TextRange.InsertAfter(String(paragraphs[0].text));
        start = Number(inserted.Start);
        for (var k = 1; k < paragraphs.length; k++) {
          inserted = inserted.InsertAfter("\r");
          inserted = inserted.InsertAfter(String(paragraphs[k].text));
        }
        length = Number(inserted.Start) + Number(inserted.Length) - start;
        formatParagraphs(box, paragraphs, start, false, args.font_size_pt);
        characters(box, start, length).Select();
      } else {
        slide = slideNumber ? presentation.Slides.Item(slideNumber) : currentSlide();
        if (slideNumber && !(slideNumber >= 1 && slideNumber <= presentation.Slides.Count)) {
          agentFail("invalid_arguments", "The presentation has " + presentation.Slides.Count + " slides.");
        }
        agentShowSlide(slide);
        var left = location && location.left_pt !== undefined ? Number(location.left_pt) : 60;
        var top = location && location.top_pt !== undefined ? Number(location.top_pt) : 60;
        var width = location && location.width_pt !== undefined ? Number(location.width_pt)
          : Math.max(120, Number(presentation.PageSetup.SlideWidth) - left - 60);
        box = slide.Shapes.AddTextbox(1, left, top, width, 40);
        safe(function () { box.TextFrame.WordWrap = MSO_TRUE; });
        safe(function () { box.TextFrame.AutoSize = 1; });
        /* all at once with single breaks: WPS makes paragraphs of them */
        box.TextFrame.TextRange.Text = paragraphs.map(function (paragraph) { return String(paragraph.text); }).join("\r");
        formatParagraphs(box, paragraphs, 1, true, args.font_size_pt);
        newBox = true;
        safe(function () { app().ActiveWindow.Selection.Unselect(); });
        safe(function () { box.Select(); });
      }
      return { slide: Number(slide.SlideIndex), slideId: Number(slide.SlideID), shapeId: Number(box.Id), newBox: newBox,
        start: start, length: length };
    },

    formulas_to_latex_text: function (args) {
      agentWritable(args);
      restoreToLatex(agentScope(args.scope));
      return agentStarted(false);
    }
  };

  /* Lists, headings, bold spans and the size of written paragraphs (they start at character `start` of the box). */
  function formatParagraphs(box, paragraphs, start, wholeBox, fontSize) {
    var text = box.TextFrame.TextRange;
    var position = start;
    paragraphs.forEach(function (paragraph, index) {
      var body = String(paragraph.text);
      var piece = text.Characters(position, body.length);
      if (fontSize > 0) safe(function () { piece.Font.Size = Number(fontSize); });
      if (wholeBox) {
        safe(function () {
          var format = text.Paragraphs(index + 1, 1).ParagraphFormat;
          if (paragraph.kind === "bullet" || paragraph.kind === "numbered") {
            format.Bullet.Visible = MSO_TRUE;
            format.Bullet.Type = paragraph.kind === "bullet" ? 1 : 2;
          } else if (paragraph.kind === "display") {
            format.Alignment = PP_ALIGN_CENTER;
          }
        });
      }
      if (paragraph.kind === "heading") {
        safe(function () { piece.Font.Bold = MSO_TRUE; piece.Font.Size = Number(piece.Font.Size) * 1.2; });
      }
      (paragraph.bold || []).forEach(function (span) {
        safe(function () { text.Characters(position + Number(span.start), Number(span.length)).Font.Bold = MSO_TRUE; });
      });
      position += body.length + 1;
    });
  }

  function agentScope(scope) {
    return scope === "selection" || scope === "slide" ? scope : "presentation";
  }

  function runAgentCommand(command) {
    agentRun = { notes: [], started: null };
    var run = agentRun;
    var operation = AGENT_OPERATIONS[command.op];
    return Promise.resolve().then(function () {
      if (!operation) agentFail("unknown_operation", "WPS Presentation has no agent operation " + command.op + ".");
      return operation(command.arguments || {});
    }).then(function (data) {
      return { id: command.id, ok: true, data: data || {} };
    }, function (error) {
      var message = String(error && error.message || error);
      if (run.notes.length) message += " " + run.notes.join(" ");
      log("agent " + command.op + " failed: " + (error && error.stack || error));
      return { id: command.id, ok: false, error: { code: error && error.agentCode || "operation_failed", message: message } };
    }).then(function (answer) {
      agentRun = null;
      return call("POST", "/agent/results", answer, 30000).catch(function (error) {
        log("agent answer not delivered: " + error.message);
      });
    });
  }

  /* Asks VisualTeX for the next operation while this module is the add-in's (main.js loads a newer one
     when VisualTeX is updated). Off: asks again now and then. */
  function startAgentLoop() {
    if (agentLoopStarted) return;
    agentLoopStarted = true;
    var version = encodeURIComponent(String(safe(function () { return app().Version; }, "") || ""));
    var next = function (delay) { setTimeout(ask, delay); };
    var ask = function () {
      if (!host || host.module !== api) return;
      /* a VisualTeX update brings a newer module without the user touching WPS (main.js) */
      if (typeof vtxRefreshSoon === "function") safe(vtxRefreshSoon);
      call("GET", "/agent/commands?host=presentation&wait=" + AGENT_POLL_SECONDS + "&version=" + version, undefined,
        (AGENT_POLL_SECONDS + 10) * 1000).then(function (reply) {
        var commands = reply.commands || [];
        if (!commands.length) {
          next(reply.enabled ? 0 : 30000);
          return;
        }
        runAgentCommand(commands[0]).then(function () { next(0); });
      }, function () {
        next(10000);
      });
    };
    next(1000);
  }

  /* Ribbon commands by control id (addin-wpp/ribbon.xml). */
  var COMMANDS = {
    vtxRedrawOmmlSelection: function () { redraw("omml", "selection"); },
    vtxRedrawOmmlSlide: function () { redraw("omml", "slide"); },
    vtxRedrawOmmlAll: function () { redraw("omml", "presentation"); },
    vtxRedrawImageSelection: function () { redraw("image", "selection"); },
    vtxRedrawImageSlide: function () { redraw("image", "slide"); },
    vtxRedrawImageAll: function () { redraw("image", "presentation"); },
    vtxToLatexSelection: function () { restoreToLatex("selection"); },
    vtxToLatexSlide: function () { restoreToLatex("slide"); },
    vtxToLatexAll: function () { restoreToLatex("presentation"); },
    vtxOmmlInline: function () { createNative("inline"); },
    vtxOmmlDisplay: function () { createNative("block"); },
    vtxNewFormula: create,
    vtxEditSelected: editSelection,
    vtxDeleteSelected: deleteSelection,
    vtxOpen: openVisualTeX
  };

  api = {
    init: function (bootstrap) {
      host = bootstrap;
      if (host && host.debug) {
        /* The test harness reaches the helpers it checks one by one. */
        api.internals = {
          call: call, newUuid: newUuid, geometry: geometry, scaleOf: scaleOf, screenBox: screenBox,
          selectedFormula: selectedFormula, shouldFork: shouldFork, writeFormula: writeFormula,
          onDoubleClick: onDoubleClick, savedEquations: savedEquations, pasteEquation: pasteEquation,
          editNative: editNative, active: function () { return active; }
        };
      }
      /* The bootstrap sets host.module to this module before init; the
         listener stops once a newer module replaces it. */
      listen(null, 0);
      startAgentLoop();
      log("module loaded");
    },
    onAction: function (id) {
      run(function () {
        var size = /^vtxSize_vtxImageFontSize_(\d+)$/.exec(String(id));
        if (size && Number(size[1]) > 0) resizeSelection(FONT_PRESETS[Number(size[1]) - 1][0]);
        else if (COMMANDS[id]) COMMANDS[id]();
        else notify("这个 VisualTeX 功能还没有接入 WPS 演示：" + id);
      });
    },
    menuContent: function (id) {
      return id === "vtxImageFontSize" ? safe(fontSizeMenu, null) : null;
    },
    itemCount: function () {
      return FONT_PRESETS.length + 1;
    },
    itemLabel: function (id, index) {
      return safe(function () { return fontSizeItemLabel(Number(index)); }, index > 0 ? presetLabel(index - 1) : "字号");
    },
    selectedIndex: function () {
      return safe(fontSizeSelectedIndex, 0);
    },
    onSelect: function (id, index) {
      run(function () {
        if (id === "vtxImageFontSize" && index > 0) resizeSelection(FONT_PRESETS[index - 1][0]);
      });
    },
    /* A WPS that raises the event opens the editor here and skips its own
       picture action. */
    onBeforeDoubleClick: function () {
      if (active) return false;
      var shape = safe(selectedFormula, null);
      if (!shape) return false;
      var now = Date.now();
      if (now - lastDoubleClick < 800) return true;
      lastDoubleClick = now;
      setTimeout(function () { run(function () { editShape(shape); }); }, 0);
      return true;
    },
    onSelectionChange: function () {
      presentationInFront = true;
      if (sizeRefresh) clearTimeout(sizeRefresh);
      sizeRefresh = setTimeout(function () {
        sizeRefresh = null;
        invalidateControl("vtxImageFontSize");
      }, 300);
    },
    onWindowActivate: function () {
      presentationInFront = true;
    },
    onWindowDeactivate: function () {
      presentationInFront = false;
    },
    /* No formula is being edited (main.js may swap in a newer module). */
    isIdle: function () {
      return !active && !agentRun;
    }
  };
  window.VisualTeXWpsModule = api;
})();
