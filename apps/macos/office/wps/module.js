/* VisualTeX for WPS Writer: the add-in's logic, served by the running
 * VisualTeX app (the installed bootstrap main.js loads it).
 *
 * It does for WPS what the Word-for-Mac add-in (office/macos-offline/word) does for
 * the Formula group, and writes the same document state (picture Title and
 * AlternativeText, owner bookmark, document variables), so a document stays
 * editable when it moves between WPS and Word.
 */
(function () {
  "use strict";

  var METADATA_PREFIX = "visualtex:v1:deflate:";
  var FORMULA_REF_PREFIX = "visualtex:formula-ref:v1:";
  var WD_OMATH_DISPLAY = 0;
  var WD_OMATH_INLINE = 1;
  var WD_ALIGN_PARAGRAPH_LEFT = 0;
  var WD_ALIGN_PARAGRAPH_CENTER = 1;
  var WD_ALIGN_PARAGRAPH_JUSTIFY = 3;
  var WD_FIELD_REF = 3;
  var WD_FIELD_STYLEREF = 10;
  var WD_FIELD_SEQUENCE = 12;
  var WD_WITHIN_TABLE = 12;
  var WD_COLOR_AUTOMATIC = -16777216;
  var NUMBERED_IMAGE_STYLE = "VisualTeX Numbered Equation";
  var IMAGE_PLACE_REF = "VisualTeXPlaceRef";
  var POLL_SECONDS = 20;

  var host = null;
  var active = null;
  var lastDoubleClick = 0;

  function app() {
    return wps.WpsApplication();
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
    call("POST", "/log", { message: String(message) }, 3000).catch(function () {});
  }

  /* An alert, or the host's own notifier (the test harness passes one: a WPS
     alert is modal and stops every add-in until it is answered). */
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

  function documentKey(doc) {
    return "wps:" + (safe(function () { return doc.FullName; }, "") || safe(function () { return doc.Name; }, ""));
  }

  function findDocument(key) {
    var documents = app().Documents;
    for (var index = 1; index <= documents.Count; index++) {
      var doc = documents.Item(index);
      if (documentKey(doc) === key) return doc;
    }
    return null;
  }

  function parseFormulaReference(value) {
    var text = String(value || "");
    if (text.indexOf(FORMULA_REF_PREFIX) !== 0) return null;
    var parts = text.slice(FORMULA_REF_PREFIX.length).split(":");
    if (parts.length !== 3 || !/^[0-9a-f-]{36}$/.test(parts[0])) return null;
    if (parts[1] !== "inline" && parts[1] !== "block") return null;
    return { formulaId: parts[0], displayMode: parts[1], numbered: parts[2] === "1" };
  }

  /* The VisualTeX picture or the Word equation at the selection, if any. */
  function selectionTarget() {
    var doc = app().ActiveDocument;
    var selection = app().ActiveWindow.Selection;
    var shapes = safe(function () { return selection.InlineShapes; }, null);
    if (shapes && shapes.Count >= 1) {
      var shape = shapes.Item(1);
      var alt = String(safe(function () { return shape.AlternativeText; }, "") || "");
      if (alt.indexOf(METADATA_PREFIX) === 0) {
        return { kind: "image", doc: doc, shape: shape, metadata: alt,
          reference: parseFormulaReference(safe(function () { return shape.Title; }, "")) };
      }
    }
    var maths = safe(function () { return selection.OMaths; }, null);
    if (maths && maths.Count >= 1) {
      var omath = safe(function () { return equationAt(doc, selection.Start); }, null);
      if (omath) return { kind: "omml", doc: doc, omath: omath };
    }
    return { kind: "none", doc: doc, selection: selection };
  }

  function fontSizeAt(range) {
    var size = Number(safe(function () { return range.Font.Size; }, 0));
    return size > 0 && size < 1000 ? size : undefined;
  }

  function textOf(range) {
    return String(safe(function () { return range.Text; }, "") || "").replace(/[\s\x07]/g, "");
  }

  /* A display formula gets its own paragraph (VTPrepareWordCreateInsertionRange). */
  function prepareDisplayRange(doc, range) {
    var paragraph = range.Paragraphs.Item(1).Range;
    var caret = range.Start;
    var contentEnd = Math.max(paragraph.Start, paragraph.End - 1);
    var before = hasContent(doc.Range(paragraph.Start, caret));
    var after = hasContent(doc.Range(caret, contentEnd));
    if (before && after) {
      doc.Range(caret, caret).InsertAfter("\r\r");
      return doc.Range(caret + 1, caret + 1);
    }
    if (before) {
      doc.Range(caret, caret).InsertAfter("\r");
      return doc.Range(caret + 1, caret + 1);
    }
    if (after) {
      doc.Range(caret, caret).InsertAfter("\r");
      return doc.Range(caret, caret);
    }
    return doc.Range(caret, caret);
  }

  function bookmarkRange(doc, name) {
    var bookmarks = doc.Bookmarks;
    if (!safe(function () { return bookmarks.Exists(name); }, false)) return null;
    return bookmarks.Item(name).Range;
  }

  function deleteBookmark(doc, name) {
    safe(function () {
      if (doc.Bookmarks.Exists(name)) doc.Bookmarks.Item(name).Delete();
    });
  }

  /* Undo can bring back the anchor of a finished session. Drop leftovers
     while the editor opens, so the document does not collect them. */
  function removeStaleAnchors(doc, keep) {
    var bookmarks = doc.Bookmarks;
    for (var index = bookmarks.Count; index >= 1; index--) {
      var bookmark = bookmarks.Item(index);
      var name = String(safe(function () { return bookmark.Name; }, ""));
      if (name !== keep && (name.indexOf("VT_W_") === 0 || name.indexOf("VT_E_") === 0)) {
        safe(function () { bookmark.Delete(); });
      }
    }
  }

  function withUndo(label, action) {
    var recording = safe(function () { wps.UndoRecord.StartCustomRecord(label); return true; }, false);
    try {
      return action();
    } finally {
      if (recording) safe(function () { wps.UndoRecord.EndCustomRecord(); });
    }
  }

  function forgetActive(cleanup) {
    var current = active;
    active = null;
    if (!current) return;
    var doc = findDocument(current.docKey);
    if (doc && cleanup) deleteBookmark(doc, current.anchor);
  }

  function startSession(spec) {
    if (active) {
      log("session=" + active.sessionId + " replaced by a new formula command");
      forgetActive(true);
    }
    var sessionId = spec.sessionId;
    active = {
      sessionId: sessionId,
      docKey: documentKey(spec.doc),
      anchor: spec.anchor,
      mode: spec.mode,
      nativeEquation: spec.nativeEquation,
      displayMode: spec.displayMode,
      conversion: spec.operation || null,
      sourceNumbered: !!spec.sourceNumbered,
      agent: !!agentRun
    };
    var body = {
      sessionId: sessionId,
      mode: spec.mode,
      displayMode: spec.displayMode,
      nativeEquation: spec.nativeEquation,
      sourceDocumentId: documentKey(spec.doc)
    };
    ["numbered", "formulaId", "sourceObjectId", "encodedMetadata", "nativeEditXml", "fontSizePt", "operation"].forEach(function (key) {
      if (spec[key] !== undefined && spec[key] !== null) body[key] = spec[key];
    });
    /* An agent's formula: VisualTeX gives the request back and draws the agent's LaTeX itself. */
    if (agentRun) body.agent = true;
    var doc = spec.doc;
    var anchor = spec.anchor;
    var posting = call("POST", "/sessions", body, 25000);
    if (agentRun) agentRun.started = posting;
    posting.then(function () {
      poll(sessionId, 0);
    }, function (error) {
      if (active && active.sessionId === sessionId) forgetActive(true);
      notify("VisualTeX 无法打开公式编辑器：" + error.message);
    });
    setTimeout(function () {
      safe(function () { removeStaleAnchors(doc, anchor); });
    }, 0);
  }

  function poll(sessionId, failures) {
    if (!active || active.sessionId !== sessionId) return;
    call("GET", "/sessions/" + sessionId + "?wait=" + POLL_SECONDS, undefined, (POLL_SECONDS + 10) * 1000).then(function (state) {
      if (!active || active.sessionId !== sessionId) return;
      if (state.status === "editing") {
        poll(sessionId, 0);
      } else if (state.status === "preflight" && active.document) {
        answerFontSizes(sessionId, state.ranges || []);
      } else if (state.status === "ready" && active.document) {
        applyDocumentResult(sessionId, state.payload);
      } else if (state.status === "ready") {
        applyResult(sessionId, state.payload);
      } else {
        log("session=" + sessionId + " ended status=" + state.status);
        forgetActive(true);
      }
    }, function (error) {
      if (!active || active.sessionId !== sessionId) return;
      if (failures >= 20) {
        log("session=" + sessionId + " lost VisualTeX: " + error.message);
        forgetActive(true);
        return;
      }
      setTimeout(function () { poll(sessionId, failures + 1); }, 1500);
    });
  }

  function download(sessionId, name) {
    var path = wps.OAAssist.DownloadFile(url("/sessions/" + sessionId + "/files/" + encodeURIComponent(name)));
    if (!path || !safe(function () { return wps.FileSystem.Exists(path); }, true)) {
      throw new Error("WPS could not download " + name + " from VisualTeX");
    }
    return path;
  }

  function removeFile(path) {
    safe(function () { wps.FileSystem.Remove(path); });
  }

  function setVariable(doc, name, value) {
    var existing = safe(function () { return doc.Variables.Item(name); }, null);
    if (existing) existing.Value = value;
    else doc.Variables.Add(name, value);
  }

  /* WPS gives a missing variable as an item whose Value is null. */
  function deleteVariable(doc, name) {
    var existing = safe(function () { return doc.Variables.Item(name); }, null);
    if (existing && safe(function () { return existing.Value; }, null) !== null) {
      existing.Delete();
      return true;
    }
    return false;
  }

  function writeVariables(doc, variables) {
    (variables.set || []).forEach(function (pair) { setVariable(doc, pair[0], pair[1]); });
    (variables.chunked || []).forEach(function (payload) {
      for (var index = payload.count + 1; index <= payload.maxChunks; index++) {
        var name = payload.stem + "_" + ("000" + index).slice(-3);
        if (!deleteVariable(doc, name)) break;
      }
    });
  }

  /* VTExpectedImageFormulaPosition: a display picture is moved so that its
     ink centre sits where the ink centre of a line of body text would (body
     text at the Normal style's size, Word's default number-font metrics), as
     the Word add-in places it; its paragraph mark stays on the baseline. A
     numbered picture is centred on its number as drawn: the number's size and
     font (VTExpectedNumberedImageFormulaPosition).
     VTEquationNumberFontMetrics: painted top and bottom of "(", ")" and the
     digits, line ascent and descent, in em. */
  var NUMBER_FONT_METRICS = {
    "times new roman": { inkTop: 0.69434, inkBottom: -0.21387, top: 0.89111, descent: 0.21631 },
    "stix two text": { inkTop: 0.73, inkBottom: -0.186, top: 0.76199, descent: 0.23801 },
    "palatino": { inkTop: 0.77051, inkBottom: -0.16602, top: 0.82275, descent: 0.27734 },
    "helvetica neue": { inkTop: 0.731, inkBottom: -0.197, top: 0.952, descent: 0.213 }
  };
  var DEFAULT_NUMBER_FONT_METRICS = {
    inkTop: 0.70556640625, inkBottom: -0.220703125, top: 0.9501953125, descent: 0.22216796875
  };

  function numberFontMetrics(fontName) {
    return NUMBER_FONT_METRICS[String(fontName || "").trim().toLowerCase()] || DEFAULT_NUMBER_FONT_METRICS;
  }

  function normalStyleFontSize(doc) {
    var size = Number(safe(function () { return doc.Styles.Item(-1).Font.Size; }, 0));
    return size > 0 && size <= 72 ? size : 11;
  }

  /* The number of a numbered picture: its VT_N_ bookmark (Word's image
     equation numbering), else the paragraph's text after the picture. */
  function numberTextFormat(doc, shape, payload) {
    var range = bookmarkRange(doc, "VT_N_" + String(payload.formulaId).replace(/-/g, ""));
    if (!range) {
      var paragraph = shape.Range.Paragraphs.Item(1).Range;
      range = doc.Range(shape.Range.End, Math.max(shape.Range.End, paragraph.End - 1));
    }
    var size = Number(safe(function () { return range.Font.Size; }, 0));
    return {
      size: size > 0 && size <= 1638 ? size : normalStyleFontSize(doc),
      metrics: numberFontMetrics(safe(function () { return range.Font.Name; }, ""))
    };
  }

  /* WPS lays a picture it sized out in a box rounded up to whole even
     points: the picture sits at the top and the rest becomes an effectExtent
     below it (written to the DOCX, so Word shows the same). Where the box
     bottom lands, the picture's bottom is that much higher; the positions
     below are worked out from the box. */
  function layoutBox(payload) {
    var height = Number(payload.heightPoints);
    var box = Math.ceil(height / 2 - 1e-6) * 2;
    return { height: height, below: Math.max(0, box - height) };
  }

  function displayPosition(doc, shape, payload) {
    var picture = layoutBox(payload);
    var boxHeight = picture.height + picture.below;
    var ratio = Number(payload.inkCenterYRatio);
    if (!(ratio >= 0 && ratio <= 1)) ratio = 0.5;
    var anchorAboveBottom = picture.below + picture.height * (1 - ratio);
    var number = payload.numbered ? numberTextFormat(doc, shape, payload) : null;
    var textSize = number ? number.size : normalStyleFontSize(doc);
    var font = number ? number.metrics : DEFAULT_NUMBER_FONT_METRICS;
    var textHalfInk = (font.inkTop - font.inkBottom) * textSize / 2;
    var alignment = safe(function () { return shape.Range.ParagraphFormat.BaseLineAlignment; }, 4);
    var raw;
    if (alignment === 0) {
      raw = boxHeight - anchorAboveBottom - textHalfInk;
      if (font.top >= font.inkTop) raw -= (font.top - font.inkTop) * textSize;
    } else if (alignment === 1) {
      raw = boxHeight / 2 - anchorAboveBottom +
        (font.inkTop + font.inkBottom - font.top + font.descent) * textSize / 2;
    } else if (alignment === 3) {
      raw = (font.top - font.descent) * textSize / 2 - anchorAboveBottom;
    } else {
      raw = (font.inkTop + font.inkBottom) * textSize / 2 - anchorAboveBottom;
    }
    raw = Math.max(-256, Math.min(256, raw));
    /* Word's rounding, with its 0.01 pt tie tolerance (VTWordPicturePosition). */
    return raw >= 0 ? Math.floor(raw + 0.51) : -Math.floor(-raw + 0.49);
  }

  /* VTWordInlineImagePosition: lower the picture by the formula's descent. */
  function inlinePosition(shape, payload) {
    var picture = layoutBox(payload);
    var alignment = safe(function () { return shape.Range.ParagraphFormat.BaseLineAlignment; }, 4);
    if (alignment === 0 || alignment === 1) return 0;
    if (alignment === 3) {
      return -Math.floor(Math.max(0, picture.height + picture.below - payload.fontSizePt) * 0.5 + 0.5);
    }
    var descent = -payload.referenceBaselinePt * picture.height / payload.referenceHeightPt;
    if (!(descent > 0)) return 0;
    return -Math.floor(Math.min(descent + picture.below, 256) + 0.5);
  }

  function insertPicture(doc, sessionId, payload, range) {
    var path = download(sessionId, payload.files.image);
    var shape = safe(function () { return doc.InlineShapes.AddPicture(path, false, true, range); }, null);
    removeFile(path);
    if (!shape && payload.files.fallback) {
      var fallback = download(sessionId, payload.files.fallback);
      shape = doc.InlineShapes.AddPicture(fallback, false, true, range);
      removeFile(fallback);
    }
    if (!shape) throw new Error("WPS did not insert the formula picture");
    shape.LockAspectRatio = 0;
    shape.Width = payload.widthPoints;
    shape.Height = payload.heightPoints;
    shape.LockAspectRatio = -1;
    shape.AlternativeText = payload.metadata;
    shape.Title = payload.formulaReference;
    safe(function () { shape.Range.Font.Size = payload.fontSizePt; });
    if (payload.displayMode === "inline") {
      var font = shape.Range.Font;
      safe(function () { font.Superscript = 0; font.Subscript = 0; });
      font.Position = inlinePosition(shape, payload);
    } else if (payload.numbered) {
      /* The numbered paragraph (tab stops, number fields) stays as it is. */
      shape.Range.Font.Position = displayPosition(doc, shape, payload);
    } else {
      shape.Range.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER;
      shape.Range.Font.Position = displayPosition(doc, shape, payload);
      var paragraph = shape.Range.Paragraphs.Item(1).Range;
      safe(function () { doc.Range(paragraph.End - 1, paragraph.End).Font.Position = 0; });
    }
    return shape;
  }

  function applyImage(doc, sessionId, payload, anchorRange, numbering, insertOnly) {
    var start = anchorRange.Start;
    if (active.mode === "edit" && !insertOnly) {
      var old = anchorRange.InlineShapes.Count >= 1 ? anchorRange.InlineShapes.Item(1) : null;
      if (!old) throw new Error("The formula picture being edited is no longer in the document");
      old.Delete();
    }
    /* The next numbered picture follows the new one: fix its restart first,
       while the positions read for the numbering still hold. */
    if (numbering && numbering.plan.nextRestart) {
      setImageIncrement(doc, numbering.plan.nextRestart.start, numbering.nextIncrement);
    }
    var shape = insertPicture(doc, sessionId, payload, doc.Range(start, start));
    writeVariables(doc, payload.variables);
    deleteBookmark(doc, payload.ownerBookmark);
    doc.Bookmarks.Add(payload.ownerBookmark, shape.Range);
    if (numbering) shape = writeImageNumber(doc, payload, numbering);
    return shape.Range;
  }

  /* WPS clips OMath.Range to the Range an equation was found through (a
     caret inside an equation gives a one-character Range), so look the
     equation up through its whole paragraph. */
  function equationAt(doc, position) {
    var touching = null;
    /* A caret at a paragraph's first character can report the paragraph
       before it, so the paragraph of the next character is asked too. */
    var probes = [doc.Range(position, position), doc.Range(position, position + 1)];
    for (var probe = 0; probe < probes.length; probe++) {
      var maths = probes[probe].Paragraphs.Item(1).Range.OMaths;
      for (var index = 1; index <= maths.Count; index++) {
        var omath = maths.Item(index);
        var range = omath.Range;
        if (range.Start <= position && position < range.End) return omath;
        if (range.End === position && !touching) touching = omath;
      }
    }
    return touching;
  }

  /* White space, including the zero-width space and word joiner that can
     separate an equation from the text after it. */
  function isBlank(text) {
    return String(text).split("").every(function (character) {
      var code = character.charCodeAt(0);
      return /\s/.test(character) || code === 0x200b || code === 0x2060;
    });
  }

  /* Delete what WPS added after an equation if it is only white space: the
     equation DOCX's paragraph mark that InsertFile brings in (a new paragraph
     at the end of the document, a split elsewhere), or the space WPS puts
     after an equation it makes inline (VTTryInsertNativeEquationFileFast:
     merge-imported-paragraph, VTDeleteTrailingInlineNativeSeparator). */
  function deleteAddedAfter(doc, start, count, keepBreaks) {
    if (!(count > 0)) return;
    var end = equationAt(doc, start).Range.End;
    var added = doc.Range(end, end + count);
    var text = String(added.Text);
    if (isBlank(text) && !(keepBreaks && /[\r\n\x0b]/.test(text))) added.Delete();
    else log("equation left " + JSON.stringify(text));
  }

  /* Inserts an equation package at position; returns how much the document
     grew. The app marks the package's paragraph <w:doNotImportParagraphEnd/>
     (omml_write::wps_equation_docx), as WPS marks its own export of an
     equation: WPS inserts the equation alone, and the paragraph it lands in
     keeps its properties (a whole paragraph would make it a Normal paragraph
     without its list) and the Undo record stays open. An older app's package
     brings its paragraph mark, which the callers remove. */
  function insertPackage(doc, position, path) {
    var before = doc.Content.End;
    doc.Range(position, position).InsertFile(path);
    return doc.Content.End - before;
  }

  function insertEquationFile(doc, path, start) {
    var inserted = insertPackage(doc, start, path);
    if (!(inserted > 0)) throw new Error("WPS did not insert the equation");
    var maths = doc.Range(start, start + inserted).OMaths;
    if (maths.Count !== 1) throw new Error("WPS inserted " + maths.Count + " equations instead of one");
    var omath = equationAt(doc, start);
    if (!omath || omath.Range.Start !== start) throw new Error("WPS did not insert the equation at the caret");
    deleteAddedAfter(doc, start, start + inserted - omath.Range.End);
  }

  /* Text, an equation or a picture in the range. Equation characters are part
     of Range.Text, and a collapsed Range can report the equation around it. */
  function hasContent(range) {
    return range.End > range.Start && (textOf(range).length > 0 || range.InlineShapes.Count > 0);
  }

  /* An equation alone in its paragraph is a display equation (Word's rule).
     WPS reports an equation imported from a DOCX as display even inside a
     line of text, so the paragraph decides unless WPS says inline. */
  function equationDisplayMode(doc, omath) {
    if (safe(function () { return omath.Type; }, WD_OMATH_DISPLAY) === WD_OMATH_INLINE) return "inline";
    var range = omath.Range;
    var paragraph = range.Paragraphs.Item(1).Range;
    var contentEnd = Math.max(paragraph.Start, paragraph.End - 1);
    var alone = !hasContent(doc.Range(paragraph.Start, range.Start)) &&
      !hasContent(doc.Range(range.End, contentEnd));
    return alone ? "block" : "inline";
  }

  /* VTV2ApplyEquationTypography, VTV2FinishDisplayEquation and
     VTFinalizeInlineNativeEquation. The type must be set either way: an
     equation that stays "display" is saved as its own math paragraph, which
     Word shows on a line of its own. */
  function finishEquation(doc, start, payload) {
    var display = payload.displayMode === "block";
    var before = doc.Content.End;
    var omath = equationAt(doc, start);
    var endBefore = omath.Range.End;
    safe(function () { omath.Type = display ? WD_OMATH_DISPLAY : WD_OMATH_INLINE; });
    var range = equationAt(doc, start).Range;
    var size = Math.round(Number(payload.fontSizePt) * 2) / 2;
    range.Font.Position = 0;
    if (size > 0) {
      range.Font.Size = size;
      safe(function () { range.Font.SizeBi = size; });
    }
    var added = doc.Content.End - before;
    var endAfter = equationAt(doc, start).Range.End;
    if (added > 0 && endAfter - endBefore === added) {
      /* After a script or fraction at its end, WPS puts the space inside the
         equation: the paragraph mark after it is not WPS's. */
      var tail = doc.Range(endAfter - added, endAfter);
      if (isBlank(String(tail.Text)) && !/[\r\n\x0b]/.test(String(tail.Text))) tail.Delete();
      else log("equation grew by " + JSON.stringify(String(tail.Text)));
    } else {
      deleteAddedAfter(doc, start, added, true);
    }
    range = equationAt(doc, start).Range;
    var paragraph = range.Paragraphs.Item(1).Range;
    if (display) {
      var format = paragraph.ParagraphFormat;
      format.Alignment = WD_ALIGN_PARAGRAPH_CENTER;
      format.SpaceBefore = 0;
      format.SpaceAfter = 0;
      format.LeftIndent = 0;
      format.RightIndent = 0;
      format.FirstLineIndent = 0;
    } else {
      var contentEnd = Math.max(paragraph.Start, paragraph.End - 1);
      if (!hasContent(doc.Range(paragraph.Start, range.Start)) && !hasContent(doc.Range(range.End, contentEnd))) {
        paragraph.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_LEFT;
      }
    }
    return range.End;
  }

  function applyEquation(doc, sessionId, payload, anchorRange, numbering, insertOnly) {
    var path = download(sessionId, numbering ? numbering.file : payload.files.native);
    var start = anchorRange.Start;
    try {
      if (active.mode === "edit" && !insertOnly) anchorRange.Delete();
      /* A number placed before the first one of its series takes over that
         number's \r switch (VTV2TransferRestart). */
      if (numbering && numbering.plan.transferFrom !== null && numbering.plan.transferFrom !== undefined) {
        transferRestart(doc, numbering.plan.transferFrom);
      }
      insertEquationFile(doc, path, start);
      var end = finishEquation(doc, start, payload);
      if (numbering) end = updateEquationNumbersFrom(doc, start);
      return doc.Range(start, end);
    } finally {
      removeFile(path);
    }
  }

  /* ---- converting the selected formula ("转为 Word OMML", "转为图片公式") ------
     A hidden editor renders the formula in the other kind (nativeToImage,
     imageToNative); the result takes the source's place. A numbered source
     gives a numbered result, numbered like a new formula; the numbers of the
     source's series move up. */

  /* Removes a numbered picture's number: what follows the picture when it is
     only the number, tabs and spaces, the tab before it and the bookmarks. */
  function removeImageNumber(doc, shape, formulaId) {
    var compactId = String(formulaId).replace(/-/g, "");
    var paragraph = shape.Range.Paragraphs.Item(1).Range;
    var tail = doc.Range(shape.Range.End, paragraph.End - 1);
    var placeRef = placeRefOf(shape);
    var number = placeRef ? fieldRange(doc, placeRef) : null;
    var outside = number
      ? String(doc.Range(tail.Start, number.Start).Text || "") + String(doc.Range(number.End, tail.End).Text || "")
      : String(tail.Text || "");
    if (!/[^\t \u00a0\u200b\u2060\x0b]/.test(outside)) tail.Delete();
    deleteBookmark(doc, "VT_N_" + compactId);
    deleteBookmark(doc, "VT_R_" + compactId);
    var lead = doc.Range(paragraph.Start, formulaShape(doc, formulaId, paragraph.Start).Range.Start);
    if (!hasMeaningfulText(lead)) lead.Delete();
  }

  /* The cross-reference targets (hidden "_Ref" bookmarks) on a numbered
     picture's number, to put on the number of what replaces it
     (VTCaptureBodyEquationReferenceBindings). */
  function imageNumberReferences(doc, shape) {
    var placeRef = placeRefOf(shape);
    if (!placeRef) return [];
    var whole = fieldRange(doc, placeRef);
    return numberBookmarks(doc, whole.Start, whole.End);
  }

  /* The references go onto the new formula's number
     (VTRestoreBodyEquationReferenceBindings). */
  function restoreNumberReferences(doc, formulaRange, names) {
    if (!names.length) return;
    var target = null;
    var omath = safe(function () { return equationAt(doc, formulaRange.Start); }, null);
    if (omath) {
      target = equationNumberTarget(doc, omath);
    } else if (formulaRange.InlineShapes.Count >= 1) {
      var placeRef = placeRefOf(formulaRange.InlineShapes.Item(1));
      target = placeRef ? imageNumberTarget(doc, placeRef) : null;
    }
    if (target) restoreBookmarks(doc, names, target);
  }

  function applyConversion(doc, sessionId, payload, anchorRange, numbering, current) {
    var source = anchorRange;
    var references = current.sourceNumbered && current.conversion !== "imageToNative"
      ? numberBookmarks(doc, anchorRange.Start, anchorRange.End) : [];
    if (current.conversion === "imageToNative") {
      var shapes = anchorRange.InlineShapes;
      if (shapes.Count < 1) throw new Error("The formula picture being converted is no longer in the document");
      var reference = parseFormulaReference(shapes.Item(1).Title);
      if (reference && reference.numbered) {
        references = imageNumberReferences(doc, shapes.Item(1));
        removeImageNumber(doc, shapes.Item(1), reference.formulaId);
      }
      var shape = reference ? formulaShape(doc, reference.formulaId, anchorRange.Start) : shapes.Item(1);
      if (reference) deleteBookmark(doc, "VT_I_" + String(reference.formulaId).replace(/-/g, ""));
      source = shape.Range;
      if (!current.sourceNumbered && reference && reference.displayMode === "block") {
        /* A centred picture paragraph becomes an equation paragraph. */
        safe(function () { source.Paragraphs.Item(1).Range.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER; });
      }
    }
    /* The character after the source finds the place again: removing a
       numbered formula makes WPS update the other numbers. */
    var place = "VT_X_C" + sessionId.replace(/-/g, "").slice(0, 12);
    deleteBookmark(doc, place);
    doc.Bookmarks.Add(place, doc.Range(source.End, source.End + 1));
    try {
      source.Delete();
      var at = bookmarkRange(doc, place);
      var position = at ? at.Start : source.Start;
      if (current.sourceNumbered && current.conversion === "imageToNative") {
        /* The numbered picture's paragraph becomes a display equation's. */
        var paragraph = doc.Range(position, position).Paragraphs.Item(1).Range;
        safe(function () { paragraph.Style = doc.Styles.Item(-1); });
        safe(function () { paragraph.ParagraphFormat.TabStops.ClearAll(); });
      }
      var insertion = doc.Range(position, position);
      var formulaRange = current.conversion === "imageToNative"
        ? applyEquation(doc, sessionId, payload, insertion, numbering, true)
        : applyImage(doc, sessionId, payload, insertion, numbering, true);
      if (numbering) restoreNumberReferences(doc, formulaRange, references);
      if (current.sourceNumbered) updateAllNumberFields(doc);
      return formulaRange;
    } finally {
      deleteBookmark(doc, place);
    }
  }

  function convertSelected(toImage) {
    var target = selectionTarget();
    var sessionId = newUuid();
    if (toImage) {
      if (target.kind !== "omml") {
        notify("请先选中一个 OMML 公式。");
        return;
      }
      var range = target.omath.Range;
      var xml = String(range.WordOpenXML || "");
      var displayMode = equationDisplayMode(target.doc, target.omath);
      var numbered = displayMode === "block" && isNumberedEquationXml(xml);
      var anchor = "VT_E_" + sessionId.slice(0, 24).replace(/-/g, "");
      target.doc.Bookmarks.Add(anchor, range);
      startSession({
        sessionId: sessionId, doc: target.doc, anchor: anchor, mode: "edit", operation: "nativeToImage",
        nativeEquation: false, displayMode: displayMode, numbered: numbered, sourceNumbered: numbered,
        formulaId: newUuid(), sourceObjectId: anchor, nativeEditXml: xml, fontSizePt: fontSizeAt(range)
      });
      return;
    }
    if (target.kind !== "image" || !target.reference) {
      notify("请先选中一个 VisualTeX 图片公式。");
      return;
    }
    var reference = target.reference;
    var pictureAnchor = "VT_W_" + sessionId.replace(/-/g, "");
    target.doc.Bookmarks.Add(pictureAnchor, target.shape.Range);
    var pictureNumbered = reference.displayMode === "block" && reference.numbered;
    startSession({
      sessionId: sessionId, doc: target.doc, anchor: pictureAnchor, mode: "edit", operation: "imageToNative",
      nativeEquation: true, displayMode: reference.displayMode, numbered: pictureNumbered,
      sourceNumbered: pictureNumbered, formulaId: reference.formulaId, encodedMetadata: target.metadata,
      sourceObjectId: target.shape.Title,
      fontSizePt: formulaPictureSize(target.doc, target.shape) || undefined
    });
  }

  /* ---- equation numbers -----------------------------------------------------
     A new numbered display formula gets the number Word would give it: the
     add-in reads what the document holds, the app applies the Word add-in's
     rules (office::wps::numbering), and the number is written in Word's form
     (a #(SEQ) equation, or VTWriteSingleParagraphImageNumber's picture
     paragraph); then the numbers after it are updated. */

  function fieldType(field) {
    return Number(safe(function () { return field.Type; }, 0));
  }

  function fieldCode(field) {
    return String(safe(function () { return field.Code.Text; }, "") || "");
  }

  function updateFieldsOfType(fields, type) {
    for (var index = 1; index <= fields.Count; index++) {
      var field = fields.Item(index);
      if (fieldType(field) === type) safe(function () { field.Update(); });
    }
  }

  /* null for a missing variable (WPS reads its Value as null). */
  function variableValue(doc, name) {
    return safe(function () {
      var value = doc.Variables.Item(name).Value;
      return value === null || value === undefined ? null : String(value);
    }, null);
  }

  /* Headings of outline level 1 and 2 outside tables and frames
     (VTV2CollectHeadings). */
  function headingFacts(doc) {
    var headings = [];
    var paragraphs = doc.Paragraphs;
    for (var index = 1; index <= paragraphs.Count; index++) {
      var paragraph = paragraphs.Item(index);
      var level = Number(safe(function () { return paragraph.OutlineLevel; }, 10));
      if (!(level >= 1 && level <= 2)) continue;
      var range = paragraph.Range;
      if (safe(function () { return range.Information(WD_WITHIN_TABLE); }, false)) continue;
      if (safe(function () { return range.Frames.Count; }, 0) > 0) continue;
      headings.push({
        start: range.Start,
        level: level,
        listString: String(safe(function () { return range.ListFormat.ListString; }, "") || ""),
        listNumbered: Number(safe(function () { return range.ListFormat.ListType; }, 0)) !== 0,
        text: String(safe(function () { return range.Text; }, "") || "")
      });
    }
    return headings;
  }

  /* Numbered display equations: a #(number) with its SEQ field. */
  function equationFacts(doc) {
    var equations = [];
    var maths = doc.OMaths;
    for (var index = 1; index <= maths.Count; index++) {
      var range = maths.Item(index).Range;
      var fields = safe(function () { return range.Fields; }, null);
      if (!fields || fields.Count === 0) continue;
      var instruction = "";
      for (var k = 1; k <= fields.Count; k++) {
        var field = fields.Item(k);
        if (fieldType(field) === WD_FIELD_SEQUENCE) {
          /* A Word-made number reads U+2217 for "*" in WPS (repairFieldInstructions). */
          instruction = fieldCode(field).replace(/\u2217/g, "*");
          break;
        }
      }
      var text = String(safe(function () { return range.Text; }, "") || "");
      var hash = text.lastIndexOf("#");
      if (!instruction || hash < 0) continue;
      equations.push({
        start: range.Start,
        instruction: instruction,
        numberText: text.slice(hash + 1).replace(/[\s\x07]/g, "")
      });
    }
    return equations;
  }

  /* The hidden increment SEQ field of a numbered picture's number. */
  function imageIncrementField(shape) {
    var fields = safe(function () { return shape.Range.Paragraphs.Item(1).Range.Fields; }, null);
    if (!fields) return null;
    for (var index = 1; index <= fields.Count; index++) {
      var field = fields.Item(index);
      var code = fieldCode(field);
      if (fieldType(field) === WD_FIELD_SEQUENCE && /SEQ\s+VisualTeXEquation/i.test(code) && /\\h\b/i.test(code)) {
        return field;
      }
    }
    return null;
  }

  function imageFacts(doc) {
    var images = [];
    var shapes = doc.InlineShapes;
    for (var index = 1; index <= shapes.Count; index++) {
      var shape = shapes.Item(index);
      var reference = parseFormulaReference(safe(function () { return shape.Title; }, ""));
      if (!reference || reference.displayMode !== "block" || !reference.numbered) continue;
      var increment = imageIncrementField(shape);
      images.push({ start: shape.Range.Start, incrementInstruction: increment ? fieldCode(increment) : "" });
    }
    return images;
  }

  /* The numbering style chosen for a document in this WPS session ("编号格式",
     or worked out by "更新公式编号"): it wins over what the document shows
     (VTRememberEquationNumbering; never written into the document). */
  var chosenStyles = {};

  function numberingFacts(doc, position) {
    var chosen = chosenStyles[documentKey(doc)];
    return {
      position: position,
      headings: headingFacts(doc),
      equations: equationFacts(doc),
      images: imageFacts(doc),
      modeVariable: chosen ? chosen.mode : variableValue(doc, "VT_EquationNumberingMode"),
      separatorVariable: chosen ? chosen.separator : variableValue(doc, "VT_EquationNumberingSeparator")
    };
  }

  /* Removes the \r switch of the equation at equationStart: the new number
     before it now starts the series (VTV2TransferRestart). */
  function transferRestart(doc, equationStart) {
    var omath = safe(function () { return equationAt(doc, equationStart); }, null);
    if (!omath) return;
    var fields = omath.Range.Fields;
    for (var index = 1; index <= fields.Count; index++) {
      var field = fields.Item(index);
      if (fieldType(field) !== WD_FIELD_SEQUENCE) continue;
      var code = field.Code;
      var match = /\\r +\d+/i.exec(String(code.Text));
      if (match) {
        var start = code.Start + match.index;
        var switchRange = doc.Range(start, start + match[0].length);
        if (String(switchRange.Text) === match[0]) switchRange.Text = " ";
      }
      return;
    }
  }

  /* WPS turns every "*" of an equation's math text into U+2217 when it opens
     a document, the switches of a Word-made number's fields too, and such a
     field updates to "Error! Bookmark not defined.". Writing the instruction
     back with "*" repairs it (Code.Text still reads U+2217 inside math, so
     this runs before every update). */
  function repairFieldInstructions(fields) {
    for (var index = 1; index <= fields.Count; index++) {
      var field = fields.Item(index);
      var code = fieldCode(field);
      if (code.indexOf("\u2217") >= 0) safe(function () { field.Code.Text = code.replace(/\u2217/g, "*"); });
    }
  }

  /* STYLEREF before SEQ in every numbered equation from the new one on, then
     every cross-reference (VTV2UpdateNumbersFrom). Returns the new
     equation's end. */
  function updateEquationNumbersFrom(doc, start) {
    var maths = doc.OMaths;
    for (var index = 1; index <= maths.Count; index++) {
      var range = maths.Item(index).Range;
      if (range.End < start) continue;
      var fields = safe(function () { return range.Fields; }, null);
      if (!fields || fields.Count === 0) continue;
      repairFieldInstructions(fields);
      updateFieldsOfType(fields, WD_FIELD_STYLEREF);
      updateFieldsOfType(fields, WD_FIELD_SEQUENCE);
    }
    updateFieldsOfType(doc.Fields, WD_FIELD_REF);
    return equationAt(doc, start).Range.End;
  }

  /* VTEquationNumberFontNameForFormula: a picture's number takes the Word
     font of its letters. */
  function numberFontName(letterFont) {
    return {
      times: "Times New Roman",
      stix: "STIX Two Text",
      palatino: "Palatino",
      helvetica: "Helvetica Neue"
    }[String(letterFont || "")] || "Cambria Math";
  }

  function textColumnWidth(doc, paragraph) {
    var setup = safe(function () { return paragraph.Sections.Item(1).PageSetup; }, null) || doc.PageSetup;
    var width = Number(safe(function () { return setup.TextColumns.Width; }, 0));
    if (!(width > 2 && width < 5000)) {
      width = Number(setup.PageWidth) - Number(setup.LeftMargin) - Number(setup.RightMargin);
    }
    return width;
  }

  function numberedImageStyle(doc) {
    var style = safe(function () { return doc.Styles.Item(NUMBERED_IMAGE_STYLE); }, null);
    if (!style) style = safe(function () { return doc.Styles.Add(NUMBERED_IMAGE_STYLE, 1); }, null);
    if (!style) return null;
    safe(function () { style.BaseStyle = -1; });
    safe(function () { style.NextParagraphStyle = -1; });
    safe(function () { style.AutomaticallyUpdate = false; });
    safe(function () { style.Font.Superscript = false; style.Font.Subscript = false; style.Font.Position = 0; });
    safe(function () { style.QuickStyle = false; });
    return style;
  }

  /* VTConfigureNumberedEquationParagraph: justified, no indents or spacing,
     a centre tab for the formula and a right tab for the number. */
  function configureNumberedParagraph(doc, paragraph) {
    var width = textColumnWidth(doc, paragraph);
    var style = numberedImageStyle(doc);
    if (style) safe(function () { paragraph.Style = style; });
    var format = paragraph.ParagraphFormat;
    var values = {
      Alignment: WD_ALIGN_PARAGRAPH_JUSTIFY, LeftIndent: 0, RightIndent: 0, FirstLineIndent: 0,
      SpaceBefore: 0, SpaceAfter: 0, LineSpacingRule: 0, KeepWithNext: false, KeepTogether: false,
      PageBreakBefore: false, WidowControl: false
    };
    Object.keys(values).forEach(function (name) {
      safe(function () { format[name] = values[name]; });
    });
    format.TabStops.ClearAll();
    format.TabStops.Add(width / 2, 1, 0);
    format.TabStops.Add(width - 1, 2, 0);
  }

  function xmlText(text) {
    return String(text).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  /* A Flat OPC package of one paragraph, for Range.InsertXML. */
  function paragraphPackage(content) {
    return "<?xml version=\"1.0\" standalone=\"yes\"?>" +
      "<pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">" +
      "<pkg:part pkg:name=\"/_rels/.rels\" pkg:contentType=\"application/vnd.openxmlformats-package.relationships+xml\">" +
      "<pkg:xmlData><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
      "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>" +
      "</Relationships></pkg:xmlData></pkg:part>" +
      "<pkg:part pkg:name=\"/word/document.xml\" pkg:contentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\">" +
      "<pkg:xmlData><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
      "<w:body><w:p>" + content + "</w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>";
  }

  /* A tab and the picture's number as Word writes it
     (VTCreateCanonicalImagePlaceRef): MACROBUTTON VisualTeXPlaceRef whose
     instruction holds the hidden increment SEQ, "(", the prefix, the
     current-value SEQ and ")", with VT_R_ on the whole field and VT_N_ on the
     visible number. Fields.Add cannot make it: WPS gives every field it adds
     a result, and Word ends a MACROBUTTON at its instruction. */
  function imageNumberXml(compactId, numbering, fontName, fontSize) {
    var font = xmlText(fontName);
    var halfPoints = Math.round(Number(fontSize) * 2);
    /* The East Asian slot too: the U+2010 separator counts as East Asian
       text where the run is marked so, and a document's East Asian font
       (a CJK font in every document WPS makes) draws it wide. */
    var properties = "<w:rPr><w:rFonts w:ascii=\"" + font + "\" w:hAnsi=\"" + font + "\" w:eastAsia=\"" + font +
      "\" w:cs=\"" + font + "\" w:hint=\"default\"/>" +
      (halfPoints > 0 ? "<w:sz w:val=\"" + halfPoints + "\"/><w:szCs w:val=\"" + halfPoints + "\"/>" : "") + "</w:rPr>";
    function run(content) { return "<w:r>" + properties + content + "</w:r>"; }
    function fieldChar(type) { return run("<w:fldChar w:fldCharType=\"" + type + "\"/>"); }
    function instruction(text) { return run("<w:instrText xml:space=\"preserve\">" + xmlText(text) + "</w:instrText>"); }
    return run("<w:tab/>") +
      "<w:bookmarkStart w:id=\"1\" w:name=\"VT_R_" + compactId + "\"/>" +
      fieldChar("begin") + instruction(" MACROBUTTON " + IMAGE_PLACE_REF + " ") +
      fieldChar("begin") + instruction(" " + numbering.increment + " ") + fieldChar("end") +
      instruction("(") +
      "<w:bookmarkStart w:id=\"2\" w:name=\"VT_N_" + compactId + "\"/>" +
      (numbering.plan.prefixText ? instruction(numbering.plan.prefixText) : "") +
      fieldChar("begin") + instruction(" " + numbering.current + " ") + fieldChar("separate") +
      instruction("1") + fieldChar("end") +
      "<w:bookmarkEnd w:id=\"2\"/>" +
      instruction(")") + fieldChar("end") +
      "<w:bookmarkEnd w:id=\"1\"/>";
  }

  /* The picture's number in Word's layout: a numbered paragraph with a tab
     before the picture and one before the number, the VT_N_ bookmark on the
     visible number and VT_R_ on the whole field (VTWriteSingleParagraphImageNumber). */
  function writeImageNumber(doc, payload, numbering) {
    var compactId = String(payload.formulaId).replace(/-/g, "");
    var shape = bookmarkRange(doc, payload.ownerBookmark).InlineShapes.Item(1);
    var paragraphStart = shape.Range.Paragraphs.Item(1).Range.Start;
    doc.Range(paragraphStart, paragraphStart).InsertBefore("\t");
    /* The tab went into the owner bookmark, which holds the picture alone. */
    shape = bookmarkRange(doc, payload.ownerBookmark).InlineShapes.Item(1);
    deleteBookmark(doc, payload.ownerBookmark);
    doc.Bookmarks.Add(payload.ownerBookmark, shape.Range);
    deleteBookmark(doc, "VT_N_" + compactId);
    deleteBookmark(doc, "VT_R_" + compactId);
    var name = numberFontName(payload.formulaLetterFont);
    var paragraphs = doc.Paragraphs.Count;
    var afterPicture = shape.Range.End;
    doc.Range(afterPicture, afterPicture).InsertXML(
      paragraphPackage(imageNumberXml(compactId, numbering, name, payload.fontSizePt)));
    var fullNumber = bookmarkRange(doc, "VT_R_" + compactId);
    if (!fullNumber || !bookmarkRange(doc, "VT_N_" + compactId)) {
      throw new Error("WPS did not insert the picture's number");
    }
    /* InsertXML ends what it inserts with a paragraph mark of its own. */
    var mark = doc.Range(fullNumber.End, fullNumber.End + 1);
    if (doc.Paragraphs.Count === paragraphs + 1 && String(mark.Text) === "\r") mark.Delete();
    fullNumber = bookmarkRange(doc, "VT_R_" + compactId);
    configureNumberedParagraph(doc, fullNumber.Paragraphs.Item(1).Range);
    /* VTApplyStaticImageEquationNumberFormatting: plain text on the baseline,
       in the letters' font at the formula's size. */
    var font = fullNumber.Font;
    safe(function () { font.Hidden = false; font.Superscript = false; font.Subscript = false; font.Position = 0; });
    safe(function () { font.Color = WD_COLOR_AUTOMATIC; });
    safe(function () { font.Name = name; font.NameAscii = name; font.NameOther = name; font.NameFarEast = name; });
    var size = Number(payload.fontSizePt);
    if (size > 0) safe(function () { font.Size = size; });
    updateImageNumbersFrom(doc, fullNumber.Start);
    shape = bookmarkRange(doc, payload.ownerBookmark).InlineShapes.Item(1);
    shape.Range.Font.Position = displayPosition(doc, shape, payload);
    return shape;
  }

  function isImageNumberField(field) {
    if (fieldType(field) !== WD_FIELD_SEQUENCE) return false;
    var code = fieldCode(field);
    return /SEQ\s+VisualTeXEquation/i.test(code) && /\\c\b/i.test(code);
  }

  /* The visible SEQ of every picture number from position on, in order, then
     every cross-reference (VTUpdateImageNumberFieldsFrom). WPS counts a
     sequence from the field instructions, so the hidden increment fields need
     no update; updating one, or a MACROBUTTON, gives it a result that Word's
     own numbers do not have. */
  function updateImageNumbersFrom(doc, position) {
    var fields = doc.Range(position, doc.Content.End).Fields;
    for (var index = 1; index <= fields.Count; index++) {
      var field = fields.Item(index);
      if (isImageNumberField(field)) safe(function () { field.Update(); });
    }
    updateFieldsOfType(doc.Fields, WD_FIELD_REF);
  }

  /* The increment SEQ of the numbered picture at imageStart gets instruction
     (VTReconcileNumberingAfterEdit: its restart changed). */
  function setImageIncrement(doc, imageStart, instruction) {
    var shapes = doc.Range(imageStart, imageStart + 1).InlineShapes;
    if (shapes.Count < 1 || !instruction) return;
    var field = imageIncrementField(shapes.Item(1));
    if (field) field.Code.Text = " " + instruction + " ";
  }

  /* ---- every number in the document's style --------------------------------
     "更新公式编号" and "编号格式" (VTUpdateEquationNumbersCore): each number
     that is not in the document's numbering style is written again
     (VTV2RestyleNumbers for equations, VTUpgradeWordNumberFormats for
     pictures; office::wps::numbering decides), then every SEQ, STYLEREF and
     REF field is updated (VTUpdateCurrentWordNumberFields). */

  var NUMBERING_VARIABLES = ["VT_EquationNumberingMode", "VT_EquationNumberingSeparator"];

  /* Numbered display equations with what their number holds
     (VTV2NumberedEquationEntry, VTV2NumberHasStyle). */
  function numberedEquationFacts(doc) {
    var equations = [];
    var maths = doc.OMaths;
    for (var index = 1; index <= maths.Count; index++) {
      var omath = maths.Item(index);
      var range = omath.Range;
      var text = String(safe(function () { return range.Text; }, "") || "");
      var hash = text.lastIndexOf("#");
      var fields = safe(function () { return range.Fields; }, null);
      if (hash < 0 || !fields || fields.Count === 0) continue;
      if (safe(function () { return equationDisplayMode(doc, omath); }, "inline") !== "block") continue;
      var instruction = "";
      var styleRefs = [];
      var others = 0;
      var legacy = null;
      for (var k = 1; k <= fields.Count; k++) {
        var field = fields.Item(k);
        var type = fieldType(field);
        var code = fieldCode(field).replace(/\u2217/g, "*");
        var reference = type === WD_FIELD_REF ? /^\s*REF\s+VT_N_([0-9a-f]{32})\b/i.exec(code) : null;
        if (type === WD_FIELD_SEQUENCE) {
          if (!instruction) instruction = code;
        } else if (type === WD_FIELD_STYLEREF) {
          styleRefs.push(code);
        } else if (reference) {
          legacy = reference[1].toLowerCase();
        } else {
          others++;
        }
      }
      /* An older VisualTeX number refers to a hidden SEQ elsewhere
         (VTNativeEquationArrayReferenceField); it is written anew. */
      if (!instruction && !legacy) continue;
      equations.push({
        start: range.Start,
        instruction: instruction,
        numberText: text.slice(hash + 1).replace(/[\s\x07]/g, ""),
        styleRefs: styleRefs,
        otherFields: others,
        legacy: !instruction,
        legacyId: !instruction ? legacy : null
      });
    }
    return equations;
  }

  /* The MACROBUTTON place reference of a numbered picture's paragraph. */
  function placeRefOf(shape) {
    var fields = safe(function () { return shape.Range.Paragraphs.Item(1).Range.Fields; }, null);
    if (!fields) return null;
    for (var index = 1; index <= fields.Count; index++) {
      var field = fields.Item(index);
      if (isImagePlaceRef(field)) return field;
    }
    return null;
  }

  /* A numbered picture's number as VTImagePlaceRefIsCanonical reads it.
     WPS lists a field among the fields of its own code. */
  function numberedImageFact(doc, shape, reference) {
    var fact = {
      start: shape.Range.Start, formulaId: reference.formulaId, incrementInstruction: "",
      currentInstruction: "", prefixText: "", wellFormed: false
    };
    var placeRef = placeRefOf(shape);
    if (!placeRef) return fact;
    var nested = placeRef.Code.Fields;
    var hidden = [];
    var current = [];
    var others = 0;
    for (var index = 1; index <= nested.Count; index++) {
      var field = nested.Item(index);
      if (isImagePlaceRef(field)) continue;
      var code = fieldCode(field);
      if (fieldType(field) === WD_FIELD_SEQUENCE && /\\h\b/i.test(code)) hidden.push(code);
      else if (isImageNumberField(field)) current.push(field);
      else others++;
    }
    fact.incrementInstruction = hidden[0] || "";
    if (current.length) {
      fact.currentInstruction = fieldCode(current[0]);
      var target = imageNumberTarget(doc, placeRef);
      if (target) {
        fact.prefixText = String(doc.Range(target.start, current[0].Code.Start - 1).Text || "");
      }
    }
    fact.wellFormed = hidden.length === 1 && current.length === 1 && others === 0 &&
      /\)\s*$/.test(fieldCode(placeRef));
    return fact;
  }

  function numberedImageFacts(doc) {
    var images = [];
    var shapes = doc.InlineShapes;
    for (var index = 1; index <= shapes.Count; index++) {
      var shape = shapes.Item(index);
      var reference = parseFormulaReference(safe(function () { return shape.Title; }, ""));
      if (!reference || reference.displayMode !== "block" || !reference.numbered) continue;
      images.push(numberedImageFact(doc, shape, reference));
    }
    return images;
  }

  function restyleFacts(doc) {
    return {
      headings: headingFacts(doc),
      equations: numberedEquationFacts(doc),
      images: numberedImageFacts(doc),
      modeVariable: variableValue(doc, NUMBERING_VARIABLES[0]),
      separatorVariable: variableValue(doc, NUMBERING_VARIABLES[1])
    };
  }

  /* The bookmarks a number carries (cross-reference targets are hidden "_Ref"
     bookmarks, so they are found through the REF fields), to put back on the
     number written in its place (VTCaptureNativeNumberBookmarks). */
  function numberBookmarks(doc, start, end) {
    var names = [];
    var fields = doc.Fields;
    for (var index = 1; index <= fields.Count; index++) {
      var field = fields.Item(index);
      if (fieldType(field) !== WD_FIELD_REF) continue;
      var name = (/^\s*REF\s+(\S+)/i.exec(fieldCode(field)) || [])[1];
      if (!name || names.indexOf(name) >= 0) continue;
      var range = bookmarkRange(doc, name);
      if (range && range.Start >= start && range.End <= end) names.push(name);
    }
    return names;
  }

  function restoreBookmarks(doc, names, target) {
    names.forEach(function (name) {
      deleteBookmark(doc, name);
      safe(function () { doc.Bookmarks.Add(name, doc.Range(target.start, target.end)); });
    });
  }

  function markedEquation(doc, marker) {
    var marked = bookmarkRange(doc, marker);
    return marked ? equationAt(doc, marked.Start) : null;
  }

  /* Puts an equation package in front of the paragraph mark the anchor
     bookmark holds, and returns the equation. Positions are read again after
     each change: WPS updates other number fields as equations come and go. */
  function insertEquationBeforeAnchor(doc, anchor, path) {
    var at = bookmarkRange(doc, anchor);
    if (!at) return null;
    var position = at.Start;
    var paragraphs = doc.Paragraphs.Count;
    try {
      insertPackage(doc, position, path);
    } catch (error) {
      return null;
    }
    at = bookmarkRange(doc, anchor);
    if (!at) return null;
    var mark = at.End - 1;
    /* A whole package brings its own paragraph mark: [equation][mark][old mark]. */
    if (doc.Paragraphs.Count > paragraphs && String(doc.Range(mark - 1, mark).Text) === "\r") {
      doc.Range(mark - 1, mark).Delete();
    }
    at = bookmarkRange(doc, anchor);
    var maths = doc.Range(at.End - 1, at.End).Paragraphs.Item(1).Range.OMaths;
    return maths.Count >= 1 ? maths.Item(maths.Count) : null;
  }

  /* An equation with its number written again: the app built the new
     equation around the old formula and it takes the old one's place, at the
     old one's size; the paragraph stays as it is. The old equation comes back
     from its own package if WPS does not take the new one. */
  function replaceNumberedEquation(doc, marker, path, backupPath) {
    var omath = markedEquation(doc, marker);
    if (!omath) throw new Error("A numbered equation moved while it was being renumbered");
    var range = omath.Range;
    var size = fontSizeAt(range);
    var names = numberBookmarks(doc, range.Start, range.End);
    var anchor = marker + "_P";
    deleteBookmark(doc, anchor);
    /* The paragraph mark after the equation finds its place again. */
    doc.Bookmarks.Add(anchor, doc.Range(range.End, range.End + 1));
    try {
      range.Delete();
      var inserted = insertEquationBeforeAnchor(doc, anchor, path);
      if (!inserted) {
        if (backupPath) insertEquationBeforeAnchor(doc, anchor, backupPath);
        throw new Error("WPS did not take the renumbered equation");
      }
      if (size > 0) {
        var font = inserted.Range.Font;
        safe(function () { font.Size = size; });
        safe(function () { font.SizeBi = size; });
      }
      var target = equationNumberTarget(doc, inserted);
      if (target) restoreBookmarks(doc, names, target);
      return inserted;
    } finally {
      deleteBookmark(doc, anchor);
    }
  }

  /* -- older VisualTeX forms --
     The VT_* bookmarks of a formula identity go; a body cross-reference to
     one of them is kept on a "_Ref" bookmark of the formula's number
     (VTFinishPureNativeIdentity), and the formula's document variables go. */

  var IDENTITY_BOOKMARKS = ["VT_R_", "VT_N_", "VT_C_", "VT_F_", "VT_I_"];

  /* target: the formula's number ({start, end}), else the old bookmark's
     range takes the references. */
  function retireIdentity(doc, compactId, target) {
    var fields = doc.Fields;
    IDENTITY_BOOKMARKS.forEach(function (prefix) {
      var oldName = prefix + compactId;
      var old = bookmarkRange(doc, oldName);
      var source = target || (old ? { start: old.Start, end: old.End } : null);
      var newName = null;
      for (var index = 1; source && index <= fields.Count; index++) {
        var field = fields.Item(index);
        if (fieldType(field) !== WD_FIELD_REF) continue;
        var code = fieldCode(field);
        var name = (/^\s*REF\s+(\S+)/i.exec(code) || [])[1];
        if (!name || name.toLowerCase() !== oldName.toLowerCase()) continue;
        if (!newName) {
          newName = "_Ref" + newUuid().replace(/-/g, "");
          safe(function () { doc.Bookmarks.Add(newName, doc.Range(source.start, source.end)); });
        }
        safe(function () { field.Code.Text = code.replace(name, newName); });
        safe(function () { field.Update(); });
      }
      deleteBookmark(doc, oldName);
    });
    var formulaId = compactId.slice(0, 8) + "-" + compactId.slice(8, 12) + "-" + compactId.slice(12, 16) + "-" +
      compactId.slice(16, 20) + "-" + compactId.slice(20);
    deletePictureState(doc, formulaId);
    deleteVariable(doc, "VT_NativeSignature_" + formulaId.replace(/-/g, "_"));
  }

  /* The hidden paragraph that held an older number's SEQ field
     (VTDeleteEquationNumberScaffold): it goes when it holds nothing else. */
  function removeNumberScaffold(doc, compactId) {
    var number = bookmarkRange(doc, "VT_N_" + compactId) || bookmarkRange(doc, "VT_C_" + compactId);
    if (!number) return;
    var paragraph = doc.Range(number.Start, Math.max(number.End, number.Start + 1)).Paragraphs.Item(1).Range;
    if (paragraph.OMaths.Count || paragraph.InlineShapes.Count) return;
    var fields = paragraph.Fields;
    var text = String(paragraph.Text || "");
    var sequences = 0;
    for (var index = 1; index <= fields.Count; index++) {
      var field = fields.Item(index);
      if (fieldType(field) === WD_FIELD_SEQUENCE) sequences++;
      text = text.replace(String(safe(function () { return field.Result.Text; }, "") || ""), "");
    }
    if (sequences && !/[^\s\x07\u00a0\u200b\u2060]/.test(text)) paragraph.Delete();
  }

  /* An older number written anew: its scaffold and identity go, and body
     references to it follow the new number. */
  function finishLegacyNumber(doc, omath, compactId) {
    removeNumberScaffold(doc, compactId);
    var target = omath ? equationNumberTarget(doc, omath) : null;
    retireIdentity(doc, compactId, target ? { start: target.start, end: target.end } : null);
  }

  /* WPS reads "*" of Word-made number fields as U+2217 (repairFieldInstructions):
     repaired everywhere first, the field updates WPS makes by itself while
     numbers change give numbers instead of errors. */
  function repairAllNumberFields(doc) {
    var maths = doc.OMaths;
    for (var index = 1; index <= maths.Count; index++) {
      var fields = safe(function () { return maths.Item(index).Range.Fields; }, null);
      if (fields && fields.Count) repairFieldInstructions(fields);
    }
  }

  /* The whole of a field: from its begin character to its end character.
     WPS keeps a field without a result (Word's MACROBUTTON) with its end
     character at Code.End, and one with a result with its end after it. */
  function fieldRange(doc, field) {
    var codeEnd = field.Code.End;
    var atCodeEnd = String(safe(function () { return doc.Range(codeEnd, codeEnd + 1).Text; }, "") || "");
    var end = atCodeEnd === "\x15" ? codeEnd + 1 : field.Result.End + 1;
    return doc.Range(field.Code.Start - 1, end);
  }

  /* A picture's number written again (VTWriteSingleParagraphImageNumber when
     its place reference is not canonical): what follows the picture may hold
     only the old number, tabs and spaces, which give way to the new number. */
  function rewriteImageNumber(doc, shape, formulaId, item, format) {
    var compactId = String(formulaId).replace(/-/g, "");
    var paragraph = shape.Range.Paragraphs.Item(1).Range;
    var tail = doc.Range(shape.Range.End, paragraph.End - 1);
    var placeRef = placeRefOf(shape);
    var number = placeRef ? fieldRange(doc, placeRef) : null;
    var outside = number
      ? String(doc.Range(tail.Start, number.Start).Text || "") + String(doc.Range(number.End, tail.End).Text || "")
      : String(tail.Text || "");
    if (outside.replace(/[\t \u00a0\u200b\u2060\x0b]/g, "").length > 0) {
      throw new Error("编号后面还有正文，未改写这个图片公式的编号");
    }
    var names = numberBookmarks(doc, tail.Start, tail.End);
    tail.Delete();
    deleteBookmark(doc, "VT_N_" + compactId);
    deleteBookmark(doc, "VT_R_" + compactId);
    var paragraphs = doc.Paragraphs.Count;
    var afterPicture = shape.Range.End;
    var numbering = { plan: { prefixText: item.prefixText }, increment: item.increment, current: item.current };
    doc.Range(afterPicture, afterPicture).InsertXML(paragraphPackage(
      imageNumberXml(compactId, numbering, numberFontName(format.formulaLetterFont), format.fontSizePt)));
    var fullNumber = bookmarkRange(doc, "VT_R_" + compactId);
    if (!fullNumber) throw new Error("WPS did not insert the picture's number");
    var mark = doc.Range(fullNumber.End, fullNumber.End + 1);
    if (doc.Paragraphs.Count === paragraphs + 1 && String(mark.Text) === "\r") mark.Delete();
    var source = bookmarkRange(doc, "VT_N_" + compactId);
    if (source && names.length) restoreBookmarks(doc, names, { start: source.Start, end: source.End });
  }

  /* VTApplyStaticImageEquationNumberFormatting on a picture's number. */
  function formatImageNumber(doc, formulaId, letterFont, size) {
    var fullNumber = bookmarkRange(doc, "VT_R_" + String(formulaId).replace(/-/g, ""));
    if (!fullNumber) return;
    var font = fullNumber.Font;
    var name = numberFontName(letterFont);
    safe(function () { font.Hidden = false; font.Superscript = false; font.Subscript = false; font.Position = 0; });
    safe(function () { font.Color = WD_COLOR_AUTOMATIC; });
    safe(function () { font.Name = name; font.NameAscii = name; font.NameOther = name; font.NameFarEast = name; });
    if (size > 0) safe(function () { font.Size = size; });
  }

  /* The letter font and size of a picture formula, as its document variables
     record them (VT_ImageLetterFont_<id>, VT_ImageScale_<id> = "size|...",
     the id spelt with "_"). */
  function imageNumberFormat(doc, shape, formulaId) {
    var key = String(formulaId).replace(/-/g, "_");
    var scale = String(variableValue(doc, "VT_ImageScale_" + key) || "");
    var size = Number(scale.split("|")[0]);
    if (!(size > 0 && size <= 512)) size = fontSizeAt(shape.Range) || normalStyleFontSize(doc);
    return { formulaLetterFont: variableValue(doc, "VT_ImageLetterFont_" + key) || "", fontSizePt: size };
  }

  /* VTUpdateCurrentWordNumberFields: every number's fields in document order,
     then the cross-references. */
  function updateAllNumberFields(doc) {
    var maths = doc.OMaths;
    for (var index = 1; index <= maths.Count; index++) {
      var fields = safe(function () { return maths.Item(index).Range.Fields; }, null);
      if (!fields || fields.Count === 0) continue;
      repairFieldInstructions(fields);
      updateFieldsOfType(fields, WD_FIELD_STYLEREF);
      updateFieldsOfType(fields, WD_FIELD_SEQUENCE);
    }
    updateImageNumbersFrom(doc, 0);
  }

  function shapeAt(doc, start) {
    var shapes = doc.Range(start, start + 1).InlineShapes;
    return shapes.Count >= 1 ? shapes.Item(1) : null;
  }

  /* A picture's number brought into the style where it stands: the hidden
     SEQ's switches and the prefix text change in place. Deleting a number's
     fields makes WPS update every SEQ field of the sequence and give each
     hidden one a result, so a well-formed number is never rebuilt. */
  function updateImageNumberInPlace(doc, shape, formulaId, item) {
    var placeRef = placeRefOf(shape);
    if (!placeRef) return false;
    var hidden = null;
    var current = null;
    var nested = placeRef.Code.Fields;
    for (var index = 1; index <= nested.Count; index++) {
      var field = nested.Item(index);
      if (isImagePlaceRef(field)) continue;
      if (isImageNumberField(field)) current = field;
      else if (fieldType(field) === WD_FIELD_SEQUENCE && /\\h\b/i.test(fieldCode(field))) hidden = field;
    }
    if (!hidden || !current) return false;
    var whole = fieldRange(doc, placeRef);
    var names = numberBookmarks(doc, whole.Start, whole.End);
    if (fieldCode(hidden).trim() !== item.increment) hidden.Code.Text = " " + item.increment + " ";
    if (fieldCode(current).trim() !== item.current) current.Code.Text = " " + item.current + " ";
    var target = imageNumberTarget(doc, placeRefOf(shape));
    if (!target) return false;
    var prefix = doc.Range(target.start, currentFieldOf(shape).Code.Start - 1);
    if (String(prefix.Text || "") !== item.prefixText) prefix.Text = item.prefixText;
    target = imageNumberTarget(doc, placeRefOf(shape));
    var compactId = String(formulaId).replace(/-/g, "");
    deleteBookmark(doc, "VT_N_" + compactId);
    doc.Bookmarks.Add("VT_N_" + compactId, doc.Range(target.start, target.end));
    restoreBookmarks(doc, names, target);
    return true;
  }

  function currentFieldOf(shape) {
    var placeRef = placeRefOf(shape);
    var nested = placeRef.Code.Fields;
    for (var index = 1; index <= nested.Count; index++) {
      if (isImageNumberField(nested.Item(index))) return nested.Item(index);
    }
    return null;
  }

  /* The picture of a formula id: its owner bookmark VT_I_<id> moves with the
     text, unlike a position. */
  function formulaShape(doc, formulaId, fallbackStart) {
    var range = bookmarkRange(doc, "VT_I_" + String(formulaId).replace(/-/g, ""));
    var shapes = range ? range.InlineShapes : null;
    if (shapes && shapes.Count >= 1) return shapes.Item(1);
    return shapeAt(doc, fallbackStart);
  }

  function applyRestyle(doc, facts, prepared, undoLabel) {
    var result = prepared.result;
    var batchId = prepared.batchId;
    var files = prepared.files;
    var imagesByStart = {};
    facts.images.forEach(function (image) { imagesByStart[image.start] = image; });
    var work = [];
    result.equations.forEach(function (item, index) {
      if (!item.changed) return;
      var built = files ? files[index] : null;
      var fact = facts.equations.filter(function (candidate) { return candidate.start === item.start; })[0];
      work.push({ start: item.start, equation: item, marker: "VT_X_" + index, unnumbered: !!(fact && fact.unnumbered),
        legacy: fact && fact.legacy ? fact.legacyId : null,
        file: built ? built.file : null, original: built ? built.original : null, error: built ? built.error : null });
    });
    result.images.forEach(function (item) {
      var fact = imagesByStart[item.start];
      if (fact) work.push({ start: item.start, image: item, fact: fact });
    });
    work.sort(function (a, b) { return b.start - a.start; });
    var failures = [];
    var changed = 0;
    withUndo(undoLabel, function () {
      removeMarkers(doc);
      repairAllNumberFields(doc);
      /* Each equation to renumber gets a bookmark that follows it while the
         others change around it. */
      work.forEach(function (step) {
        if (!step.marker) return;
        var omath = equationAt(doc, step.start);
        if (omath) doc.Bookmarks.Add(step.marker, omath.Range);
      });
      work.forEach(function (step) {
        try {
          if (step.equation) {
            if (!bookmarkRange(doc, step.marker)) throw new Error("A numbered equation was removed while it was being renumbered");
            /* An equation is always written anew: WPS mangles a field
               instruction set inside math and turns "-" typed into math into
               a minus sign, so a number cannot be changed where it stands. */
            if (!step.file) throw new Error(step.error || "VisualTeX could not renumber an equation");
            var path = download(batchId, step.file);
            var backup = step.original ? safe(function () { return download(batchId, step.original); }, null) : null;
            var inserted = null;
            try {
              inserted = replaceNumberedEquation(doc, step.marker, path, backup);
            } finally {
              removeFile(path);
              if (backup) removeFile(backup);
            }
            if (step.legacy) finishLegacyNumber(doc, inserted, step.legacy);
            changed++;
          } else {
            var formulaId = step.fact.formulaId;
            var shape = formulaShape(doc, formulaId, step.start);
            if (!shape) throw new Error("A numbered picture was removed while it was being renumbered");
            if (step.fact.unnumbered) {
              numberImageFormula(doc, shape, formulaId, step.image, prepared.metadata[step.start]);
              changed++;
              return;
            }
            var format = imageNumberFormat(doc, shape, formulaId);
            if (step.image.changed) {
              if (!(step.fact.wellFormed && updateImageNumberInPlace(doc, shape, formulaId, step.image))) {
                rewriteImageNumber(doc, shape, formulaId, step.image, format);
              }
              changed++;
            }
            formatImageNumber(doc, formulaId, format.formulaLetterFont, format.fontSizePt);
          }
        } catch (error) {
          failures.push(error && error.message || String(error));
        } finally {
          if (step.marker) deleteBookmark(doc, step.marker);
        }
      });
      removeMarkers(doc);
      updateAllNumberFields(doc);
    });
    if (failures.length) log("restyle failures " + JSON.stringify(failures.slice(0, 5)));
    return { changed: changed, total: result.equations.length + result.images.length, failures: failures };
  }

  function removeMarkers(doc) {
    var bookmarks = doc.Bookmarks;
    for (var index = bookmarks.Count; index >= 1; index--) {
      var bookmark = bookmarks.Item(index);
      if (String(safe(function () { return bookmark.Name; }, "")).indexOf("VT_X_") === 0) {
        safe(function () { bookmark.Delete(); });
      }
    }
  }

  /* Plans every number of the active document in its style (or in the style
     chosen now), with the unnumbered display formulas "批量编号" adds, builds
     the equations to write, then writes everything in one Undo step and
     updates every number field. */
  function renumberDocument(options) {
    var doc = options.doc || app().ActiveDocument;
    var key = documentKey(doc);
    var facts = restyleFacts(doc);
    (options.equations || []).forEach(function (candidate) {
      facts.equations.push({ start: candidate.start, instruction: "", numberText: "", unnumbered: true });
    });
    (options.images || []).forEach(function (candidate) {
      facts.images.push({ start: candidate.start, formulaId: candidate.formulaId, unnumbered: true,
        metadata: candidate.metadata });
    });
    var style = options.chosen || chosenStyles[key];
    var body = { facts: facts };
    if (style) {
      body.mode = style.mode;
      body.separator = style.separator;
    }
    var unnumberedStarts = {};
    facts.equations.forEach(function (fact) { if (fact.unnumbered) unnumberedStarts[fact.start] = true; });
    var prepared = { result: null, batchId: null, files: null, metadata: {} };
    return call("POST", "/numbers/restyle", body, 30000).then(function (result) {
      /* VTPinEquationNumbering: the style is kept for this session and the
         earlier versions' document variables go. */
      chosenStyles[key] = result.style;
      prepared.result = result;
      var changedEquations = result.equations.filter(function (item) { return item.changed; });
      if (!changedEquations.length) return null;
      prepared.batchId = newUuid();
      var items = changedEquations.map(function (item) {
        var omath = equationAt(doc, item.start);
        return { xml: String(omath.Range.WordOpenXML), plan: item.plan, unnumbered: !!unnumberedStarts[item.start] };
      });
      return call("POST", "/numbers/equations", { batchId: prepared.batchId, items: items }, 60000).then(function (built) {
        var files = [];
        var position = 0;
        result.equations.forEach(function (item) {
          if (!item.changed) {
            files.push(null);
            return;
          }
          files.push({ file: built.files[position], original: (built.originals || [])[position],
            error: built.errors[position] });
          position++;
        });
        prepared.files = files;
      });
    }).then(function () {
      var images = options.images || [];
      if (!images.length) return null;
      return call("POST", "/images/numbered", { metadata: images.map(function (image) { return image.metadata; }) }, 30000)
        .then(function (marked) {
          images.forEach(function (image, index) { prepared.metadata[image.start] = marked.items[index]; });
        });
    }).then(function () {
      var current = findDocument(key);
      if (!current) throw new Error("The document was closed before its numbers were written");
      NUMBERING_VARIABLES.forEach(function (name) { deleteVariable(current, name); });
      return applyRestyle(current, facts, prepared, options.undoLabel || "VisualTeX 更新公式编号");
    });
  }

  function updateNumbers() {
    return renumberDocument({}).then(function (summary) {
      safe(function () { app().StatusBar = "VisualTeX: 已更新 " + summary.total + " 个公式编号"; });
      if (summary.failures.length) notify("VisualTeX 有 " + summary.failures.length + " 个编号未能更新：" + summary.failures[0]);
    }).catch(function (error) {
      notify("VisualTeX 更新公式编号失败：" + error.message);
    });
  }

  /* ---- "更新公式格式" (VisualTeX_UpdateEquationFormat) -----------------------
     The document's formulas in the current form, as one Undo step: every
     number in the document's style (older reference numbers written anew),
     the older VisualTeX tags gone (the VT_F_ bookmark of a tagged equation,
     the retired numbering-style variables) and inline pictures on the text
     baseline (VTV2InlineBaselineRepairs). The formulas do not change. */

  function legacyTagIds(doc) {
    var ids = [];
    var bookmarks = doc.Bookmarks;
    for (var index = 1; index <= bookmarks.Count; index++) {
      var match = /^VT_F_([0-9a-f]{32})$/i.exec(String(safe(function () { return bookmarks.Item(index).Name; }, "")));
      if (match) ids.push(match[1].toLowerCase());
    }
    return ids;
  }

  function inlineBaselineRepairs(doc) {
    var repairs = [];
    var shapes = doc.InlineShapes;
    for (var index = 1; index <= shapes.Count; index++) {
      var shape = shapes.Item(index);
      if (!isFormulaPicture(shape)) continue;
      var reference = parseFormulaReference(shape.Title);
      if (!reference || reference.displayMode !== "inline") continue;
      if (variableValue(doc, "VT_ImageScale_" + String(reference.formulaId).replace(/-/g, "_")) === null) continue;
      var scale = imageScale(doc, shape, reference.formulaId);
      var height = Number(shape.Height);
      var expected = inlinePosition(shape, {
        heightPoints: height, referenceHeightPt: height, fontSizePt: scale.size,
        referenceBaselinePt: scale.baseline * scale.size / IMAGE_REFERENCE_FONT_SIZE
      });
      var font = shape.Range.Font;
      var raised = !!safe(function () { return font.Superscript; }, 0) || !!safe(function () { return font.Subscript; }, 0);
      if (Number(font.Position) !== expected || raised) repairs.push({ shape: shape, position: expected });
    }
    return repairs;
  }

  /* The tag of a tagged equation goes (VTFinishPureNativeEquation): with the
     word joiner the add-in once wrote after it (VTRetireLegacyNativeBoundary). */
  function retireEquationTag(doc, compactId) {
    var tagged = bookmarkRange(doc, "VT_F_" + compactId);
    var omath = tagged ? safe(function () { return equationAt(doc, tagged.Start); }, null) : null;
    if (omath) {
      var end = omath.Range.End;
      var boundary = doc.Range(end, end + 1);
      if (end < doc.Content.End - 1 && String(boundary.Text) === "\u2060" && boundary.OMaths.Count === 0) {
        var next = safe(function () { return equationAt(doc, end + 1); }, null);
        if (next && next.Range.Start === end + 1) boundary.Text = " ";
        else boundary.Delete();
      }
    }
    var target = omath ? equationNumberTarget(doc, omath) : null;
    retireIdentity(doc, compactId, target);
  }

  function updateEquationFormat() {
    var doc = safe(function () { return app().ActiveDocument; }, null);
    if (!doc) {
      notify("请先打开一个 WPS 文字文档。");
      return null;
    }
    var key = documentKey(doc);
    var style = chosenStyles[key];
    var body = { facts: restyleFacts(doc) };
    if (style) {
      body.mode = style.mode;
      body.separator = style.separator;
    }
    return call("POST", "/numbers/restyle", body, 30000).then(function (result) {
      var numbers = result.equations.concat(result.images).filter(function (item) { return item.changed; }).length;
      var tags = legacyTagIds(doc);
      var tagCount = tags.length + NUMBERING_VARIABLES.filter(function (name) { return variableValue(doc, name) !== null; }).length;
      var repairs = inlineBaselineRepairs(doc);
      if (!numbers && !tagCount && !repairs.length) {
        notify("本文档中的公式已经是当前格式。");
        return null;
      }
      var message = "将本文档中的公式更新为当前 VisualTeX 格式？\n";
      if (numbers) message += "\n" + numbers + " 个公式编号将按当前编号格式重写";
      if (tagCount) message += "\n" + tagCount + " 个旧版 VisualTeX 标记将被移除";
      if (repairs.length) message += "\n" + repairs.length + " 个行内图片公式将对齐正文基线";
      message += "\n\n公式内容不变，可一步撤销。";
      if (!confirmWithUser(message)) return null;
      var record = openUndo("VisualTeX 更新公式格式");
      return renumberDocument({ doc: doc, undoLabel: "VisualTeX 更新公式格式" }).then(function (summary) {
        var current = findDocument(key);
        if (!current) throw new Error("文档已经关闭");
        legacyTagIds(current).forEach(function (compactId) { retireEquationTag(current, compactId); });
        inlineBaselineRepairs(current).forEach(function (repair) {
          var font = repair.shape.Range.Font;
          safe(function () { font.Superscript = 0; font.Subscript = 0; });
          font.Position = repair.position;
        });
        record.end();
        var text = "已更新 " + numbers + " 个公式编号，对齐 " + repairs.length + " 个行内图片公式，移除 " + tagCount + " 个旧版标记。";
        if (summary.failures.length) text += "\n有 " + summary.failures.length + " 个编号未能更新：" + summary.failures[0];
        notify(text);
      }, function (error) {
        record.end();
        throw error;
      });
    }).catch(function (error) {
      notify("VisualTeX 更新公式格式失败：" + error.message);
    });
  }

  /* "编号格式" (VTApplyEquationNumberingFormatPreset): the chosen format is
     the default for new documents and this document's style from now on. */
  function applyNumberingFormat(mode, separator) {
    var style = { mode: mode, separator: mode === "sequence" ? "." : separator };
    call("POST", "/numbering/preference", style, 10000).catch(function (error) {
      log("numbering preference not saved: " + error.message);
    });
    chosenStyles[documentKey(app().ActiveDocument)] = style;
    return renumberDocument({ chosen: style }).then(function (summary) {
      safe(function () { app().StatusBar = "VisualTeX: 已更新 " + summary.total + " 个公式编号"; });
      if (summary.failures.length) notify("VisualTeX 有 " + summary.failures.length + " 个编号未能更新：" + summary.failures[0]);
    }).catch(function (error) {
      notify("VisualTeX 更改编号格式失败：" + error.message);
    });
  }

  /* ---- batch numbering ("批量编号", VTNumberDisplayFormulas) ------------------
     Every unnumbered display formula of the selection or the document gets a
     number in the document's format, pictures and equations each in their own
     series; numbered formulas keep theirs and the ones after follow on.
     Formulas in tables or sharing their paragraph with other content are left
     alone and counted. */

  function inTable(range) {
    return !!safe(function () { return range.Information(WD_WITHIN_TABLE); }, false);
  }

  function inScope(range, scope) {
    return !scope || (range.End > scope.start && range.Start < scope.end) ||
      (range.Start === range.End && range.Start >= scope.start && range.Start <= scope.end);
  }

  function displayNumberingCandidates(doc, scope) {
    var found = { equations: [], images: [], numbered: 0, table: 0, other: 0 };
    var maths = doc.OMaths;
    for (var index = 1; index <= maths.Count; index++) {
      var omath = maths.Item(index);
      var range = omath.Range;
      if (!inScope(range, scope)) continue;
      var text = String(safe(function () { return range.Text; }, "") || "");
      if (text.indexOf("#") >= 0 && safe(function () { return range.Fields.Count; }, 0) > 0) {
        found.numbered++;
      } else if (inTable(range)) {
        found.table++;
      } else if (safe(function () { return equationDisplayMode(doc, omath); }, "inline") !== "block") {
        continue;
      } else {
        found.equations.push({ start: range.Start });
      }
    }
    var shapes = doc.InlineShapes;
    for (var k = 1; k <= shapes.Count; k++) {
      var shape = shapes.Item(k);
      var reference = parseFormulaReference(safe(function () { return shape.Title; }, ""));
      if (!reference || reference.displayMode !== "block" || !inScope(shape.Range, scope)) continue;
      if (reference.numbered) {
        found.numbered++;
        continue;
      }
      var paragraph = shape.Range.Paragraphs.Item(1).Range;
      var others = textOf(doc.Range(paragraph.Start, shape.Range.Start)) + textOf(doc.Range(shape.Range.End, paragraph.End));
      if (inTable(shape.Range)) found.table++;
      else if (others.length > 0 || paragraph.InlineShapes.Count !== 1 || paragraph.OMaths.Count > 0) found.other++;
      else found.images.push({ start: shape.Range.Start, formulaId: reference.formulaId,
        metadata: String(safe(function () { return shape.AlternativeText; }, "") || "") });
    }
    return found;
  }

  function confirmWithUser(message) {
    /* An agent asked for it: yes. */
    if (agentRun) return true;
    if (host && typeof host.confirm === "function") return host.confirm(String(message));
    try { return confirm(message); } catch (error) { return true; }
  }

  function numberDisplayFormulas(wholeDocument) {
    var doc = app().ActiveDocument;
    var selection = app().ActiveWindow.Selection;
    var scope = wholeDocument ? null : { start: selection.Start, end: selection.End };
    var scopeText = wholeDocument ? "全文" : "所选内容";
    var found = displayNumberingCandidates(doc, scope);
    var details = "";
    if (found.numbered) details += "\n已有编号：" + found.numbered + " 个";
    if (found.table) details += "\n位于表格中（未处理）：" + found.table + " 个";
    if (found.other) details += "\n与其他内容同段（未处理）：" + found.other + " 个";
    var total = found.equations.length + found.images.length;
    if (!total) {
      notify(scopeText + "中没有需要编号的行间公式。" + details);
      return null;
    }
    var message = "为" + scopeText + "中 " + total + " 个未编号的行间公式添加编号？\n";
    if (found.images.length) message += "\nVisualTeX 图片公式：" + found.images.length + " 个";
    if (found.equations.length) message += "\nWord OMML 公式：" + found.equations.length + " 个";
    message += "\n\n已有编号保持不变，其后的编号自动顺延，交叉引用同步更新。" + details + "\n\n可一步撤销。";
    if (!confirmWithUser(message)) return null;
    return renumberDocument({ equations: found.equations, images: found.images, undoLabel: "VisualTeX 批量编号" })
      .then(function (summary) {
        safe(function () { app().StatusBar = "VisualTeX: 已为 " + total + " 个行间公式编号"; });
        if (summary.failures.length) notify("VisualTeX 有 " + summary.failures.length + " 个公式未能编号：" + summary.failures[0]);
      }).catch(function (error) {
        notify("VisualTeX 批量编号失败：" + error.message);
      });
  }

  /* A display picture becomes a numbered one: its metadata and title say so,
     its letter font is recorded, and its paragraph takes the number.
     Cross-reference targets resting on the picture (a converted formula's,
     kept from its old number) move onto the new number. */
  function numberImageFormula(doc, shape, formulaId, item, marked) {
    if (!marked || marked.error || !marked.metadata) throw new Error((marked && marked.error) || "VisualTeX could not mark the picture numbered");
    var compactId = String(formulaId).replace(/-/g, "");
    var key = String(formulaId).replace(/-/g, "_");
    var references = numberBookmarks(doc, shape.Range.Start, shape.Range.End);
    shape.AlternativeText = marked.metadata;
    shape.Title = FORMULA_REF_PREFIX + formulaId + ":block:1";
    if (marked.letterFont && !variableValue(doc, "VT_ImageLetterFont_" + key)) {
      setVariable(doc, "VT_ImageLetterFont_" + key, marked.letterFont);
    }
    var owner = "VT_I_" + compactId;
    if (!bookmarkRange(doc, owner)) doc.Bookmarks.Add(owner, shape.Range);
    var format = imageNumberFormat(doc, shape, formulaId);
    var ratio = Number(variableValue(doc, "VT_ImageInkCenter_" + key));
    var payload = {
      formulaId: formulaId, ownerBookmark: owner, numbered: true, displayMode: "block",
      formulaLetterFont: format.formulaLetterFont, fontSizePt: format.fontSizePt,
      heightPoints: Number(shape.Height), inkCenterYRatio: ratio >= 0 && ratio <= 1 ? ratio : 0.5
    };
    var numbering = { plan: { prefixText: item.prefixText, nextRestart: null }, increment: item.increment, current: item.current };
    shape = writeImageNumber(doc, payload, numbering);
    if (references.length) {
      var target = imageNumberTarget(doc, placeRefOf(shape));
      if (target) restoreBookmarks(doc, references, target);
    }
  }

  /* ---- formula font sizes ("公式字号") -----------------------------------------
     "字号" resizes the selected picture formulas from their reference geometry
     (VTApplyWordImageFormulaFontSize: no new rendering); "行内转换" and
     "行间转换" resize every inline or display formula of the document and set
     the size conversions, redraws and imports give new ones
     (VTWordRibbonApplyConversionFontSize, VisualTeXConversionFontSizes.txt). */

  var IMAGE_REFERENCE_FONT_SIZE = 14;
  var FONT_PRESETS = [
    [5, "八号"], [5.5, "七号"], [6.5, "小六"], [7.5, "六号"], [9, "小五"], [10.5, "五号"], [12, "小四"],
    [14, "四号"], [15, "小三"], [16, "三号"], [18, "小二"], [22, "二号"], [24, "小一"], [26, "一号"],
    [36, "小初"], [42, "初号"], [48, ""], [54, ""], [72, ""], [96, ""]
  ];
  var conversionSizes = { inline: 0, block: 0, loaded: false };
  var sizeRefresh = null;

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

  function validFormulaSize(size) {
    return size >= 1 && size <= 512;
  }

  function loadConversionSizes() {
    if (conversionSizes.loaded) return;
    conversionSizes.loaded = true;
    call("GET", "/conversion-font-sizes", undefined, 5000).then(function (sizes) {
      conversionSizes.inline = validFormulaSize(sizes.inline) ? sizes.inline : 0;
      conversionSizes.block = validFormulaSize(sizes.display) ? sizes.display : 0;
      invalidateControl("vtxInlineConversionFontSize");
      invalidateControl("vtxDisplayConversionFontSize");
    }, function () { conversionSizes.loaded = false; });
  }

  function invalidateControl(id) {
    safe(function () { if (host && host.ribbon) host.ribbon.InvalidateControl(id); });
  }

  function isFormulaPicture(shape) {
    return String(safe(function () { return shape.AlternativeText; }, "") || "").indexOf(METADATA_PREFIX) === 0 &&
      !!parseFormulaReference(safe(function () { return shape.Title; }, ""));
  }

  /* The picture formulas the selection holds, or the one right at the caret
     (VTVisualTeXInlineShapeAtSelection). */
  function selectedFormulaPictures() {
    var doc = app().ActiveDocument;
    var selection = app().ActiveWindow.Selection;
    var pictures = [];
    var shapes = safe(function () { return selection.InlineShapes; }, null);
    if (shapes && shapes.Count > 0) {
      for (var index = 1; index <= shapes.Count; index++) {
        if (isFormulaPicture(shapes.Item(index))) pictures.push(shapes.Item(index));
      }
      return pictures;
    }
    var caret = selection.Start;
    [caret, caret - 1].forEach(function (position) {
      if (pictures.length || position < 0) return;
      var near = doc.Range(position, position + 1).InlineShapes;
      if (near.Count === 1 && isFormulaPicture(near.Item(1))) pictures.push(near.Item(1));
    });
    return pictures;
  }

  /* VT_ImageScale_<id> = "size|referenceWidth|referenceHeight|referenceBaseline|observed",
     the reference geometry at 14 pt. A picture resized by hand since gives its
     new geometry as the reference (VTEnsureWordImageScaleState). */
  function imageScale(doc, shape, formulaId) {
    var key = String(formulaId).replace(/-/g, "_");
    var fields = String(variableValue(doc, "VT_ImageScale_" + key) || "").split("|").map(Number);
    var scale = fields.length === 5 && validFormulaSize(fields[0]) && fields[1] > 0 && fields[2] > 0
      ? { size: fields[0], width: fields[1], height: fields[2], baseline: Math.min(0, Math.max(-256, fields[3] || 0)), observed: fields[4] }
      : null;
    var width = Number(shape.Width);
    var height = Number(shape.Height);
    if (!scale) {
      var size = fontSizeAt(shape.Range) || normalStyleFontSize(doc);
      return { size: size, width: width * IMAGE_REFERENCE_FONT_SIZE / size, height: height * IMAGE_REFERENCE_FONT_SIZE / size,
        baseline: 0, observed: size };
    }
    var expectedWidth = scale.width * scale.size / IMAGE_REFERENCE_FONT_SIZE;
    var expectedHeight = scale.height * scale.size / IMAGE_REFERENCE_FONT_SIZE;
    /* WPS reports a picture's size in steps of half a point, so only a clearly
       different size counts as resized by hand. */
    if (Math.abs(width - expectedWidth) > 2 || Math.abs(height - expectedHeight) > 2) {
      var ratio = scale.height > 0 ? scale.baseline / scale.height : 0;
      scale.width = width * IMAGE_REFERENCE_FONT_SIZE / scale.size;
      scale.height = height * IMAGE_REFERENCE_FONT_SIZE / scale.size;
      scale.baseline = ratio < 0 ? Math.max(-256, scale.height * ratio) : 0;
    }
    return scale;
  }

  function formulaPictureSize(doc, shape) {
    var reference = parseFormulaReference(safe(function () { return shape.Title; }, ""));
    if (!reference) return 0;
    return imageScale(doc, shape, reference.formulaId).size;
  }

  /* VTApplyWordImageFormulaFontSize. */
  function resizeFormulaPicture(doc, shape, size) {
    var reference = parseFormulaReference(safe(function () { return shape.Title; }, ""));
    if (!reference) return false;
    var scale = imageScale(doc, shape, reference.formulaId);
    var width = scale.width * size / IMAGE_REFERENCE_FONT_SIZE;
    var height = scale.height * size / IMAGE_REFERENCE_FONT_SIZE;
    if (!(width > 0 && height > 0 && width <= 10000 && height <= 10000)) {
      throw new Error("这个字号会让公式图片过大或过小");
    }
    safe(function () { shape.Range.Font.Size = size; });
    shape.LockAspectRatio = 0;
    shape.Width = width;
    shape.Height = height;
    shape.LockAspectRatio = -1;
    var key = String(reference.formulaId).replace(/-/g, "_");
    var ratio = Number(variableValue(doc, "VT_ImageInkCenter_" + key));
    var payload = {
      formulaId: reference.formulaId, displayMode: reference.displayMode, numbered: reference.numbered,
      heightPoints: height, referenceHeightPt: height, referenceBaselinePt: scale.baseline * size / IMAGE_REFERENCE_FONT_SIZE,
      fontSizePt: size, inkCenterYRatio: ratio >= 0 && ratio <= 1 ? ratio : 0.5
    };
    if (reference.displayMode === "inline") {
      shape.Range.Font.Position = inlinePosition(shape, payload);
    } else {
      if (reference.numbered) {
        formatImageNumber(doc, reference.formulaId, variableValue(doc, "VT_ImageLetterFont_" + key) || "", size);
      }
      shape.Range.Font.Position = displayPosition(doc, shape, payload);
    }
    setVariable(doc, "VT_ImageScale_" + key, [size, scale.width, scale.height, scale.baseline, size].join("|"));
    return true;
  }

  function resizeSelectedPictures(size) {
    var pictures = selectedFormulaPictures();
    if (!pictures.length) {
      notify("请先选中一个或多个 VisualTeX 图片公式，再选择字号。");
      return;
    }
    var doc = app().ActiveDocument;
    withUndo("VisualTeX 图片公式字号", function () {
      pictures.forEach(function (shape) { resizeFormulaPicture(doc, shape, size); });
    });
    invalidateControl("vtxImageFontSize");
  }

  /* VTApplyFontSizeToDocumentFormulas: every formula of one kind; a numbered
     equation is a display formula. */
  function resizeDocumentFormulas(displayMode, size) {
    var doc = app().ActiveDocument;
    var resized = 0;
    withUndo("VisualTeX 公式字号", function () {
      var shapes = doc.InlineShapes;
      var pictures = [];
      for (var index = 1; index <= shapes.Count; index++) {
        var shape = shapes.Item(index);
        var reference = isFormulaPicture(shape) ? parseFormulaReference(shape.Title) : null;
        if (reference && reference.displayMode === displayMode) pictures.push(shape);
      }
      pictures.forEach(function (shape) {
        if (safe(function () { return resizeFormulaPicture(doc, shape, size); }, false)) resized++;
      });
      var maths = doc.OMaths;
      for (var k = 1; k <= maths.Count; k++) {
        var omath = maths.Item(k);
        var mode = String(omath.Range.Text || "").indexOf("#") >= 0 ? "block" : equationDisplayMode(doc, omath);
        if (mode !== displayMode) continue;
        var font = omath.Range.Font;
        safe(function () { font.Size = size; });
        safe(function () { font.SizeBi = size; });
        resized++;
      }
    });
    return resized;
  }

  function applyConversionFontSize(id, index) {
    var displayMode = id === "vtxDisplayConversionFontSize" ? "block" : "inline";
    var size = index > 0 ? FONT_PRESETS[index - 1][0] : 0;
    if (size > 0 && app().Documents.Count > 0) {
      resizeDocumentFormulas(displayMode, size);
      invalidateControl("vtxImageFontSize");
    }
    conversionSizes[displayMode] = size;
    call("POST", "/conversion-font-sizes", { inline: conversionSizes.inline, display: conversionSizes.block }, 10000)
      .catch(function (error) { notify("VisualTeX 无法保存转换字号：" + error.message); });
    invalidateControl(id);
  }

  /* The ribbon's font-size lists (VTWordRibbonGetImageFontSizeItemLabel and
     its siblings): the first item says what the selection holds or that
     conversions follow the body text. */
  function fontSizeItemCount() {
    return FONT_PRESETS.length + 1;
  }

  function fontSizeItemLabel(id, index) {
    if (index > 0) return presetLabel(index - 1);
    if (id !== "vtxImageFontSize") return "跟随正文";
    var doc = safe(function () { return app().ActiveDocument; }, null);
    var pictures = doc ? safe(selectedFormulaPictures, []) : [];
    if (!pictures.length) return "请选择图片公式";
    var sizes = pictures.map(function (shape) { return formulaPictureSize(doc, shape); });
    var mixed = sizes.some(function (size) { return Math.abs(size - sizes[0]) > 0.05; });
    return "当前: " + (mixed ? "混合字号" : sizeLabel(sizes[0]));
  }

  function fontSizeSelectedIndex(id) {
    if (id === "vtxImageFontSize") {
      var doc = safe(function () { return app().ActiveDocument; }, null);
      var pictures = doc ? safe(selectedFormulaPictures, []) : [];
      if (pictures.length !== 1) return 0;
      var index = presetIndex(formulaPictureSize(doc, pictures[0]));
      return index >= 0 ? index + 1 : 0;
    }
    loadConversionSizes();
    var size = conversionSizes[id === "vtxDisplayConversionFontSize" ? "block" : "inline"];
    var preset = size > 0 ? presetIndex(size) : -1;
    return preset >= 0 ? preset + 1 : 0;
  }

  /* ---- formulas back to LaTeX ("将 OMML 公式转为 LaTeX", "将图片公式转为 LaTeX") ----
     Each formula becomes its LaTeX as text, "$...$" inline and "$$...$$" for
     a display formula; a number goes with its formula and the paragraph
     becomes plain text (VTRestoreCachedImageFormulasToLatex,
     VTReplaceFormulaRestoreRangeWithLatex, VTFinalizeFormulaRestoreNumbering).
     A picture's LaTeX is in its document variables (or its metadata); an
     equation is read by the app. */

  function base64UrlText(value) {
    var base64 = String(value).replace(/-/g, "+").replace(/_/g, "/");
    while (base64.length % 4) base64 += "=";
    var binary = atob(base64);
    var bytes = new Uint8Array(binary.length);
    for (var index = 0; index < binary.length; index++) bytes[index] = binary.charCodeAt(index);
    return new TextDecoder("utf-8").decode(bytes);
  }

  function chunkedVariable(doc, stem) {
    var count = Number(variableValue(doc, stem + "_Count"));
    if (!(count >= 1 && count <= 128)) return null;
    var text = "";
    for (var index = 1; index <= count; index++) {
      var chunk = variableValue(doc, stem + "_" + ("000" + index).slice(-3));
      if (chunk === null) return null;
      text += chunk;
    }
    return text;
  }

  function storedLatex(doc, formulaId) {
    var payload = chunkedVariable(doc, "VT_Latex_" + String(formulaId).replace(/-/g, "_"));
    return payload ? safe(function () { return base64UrlText(payload); }, null) : null;
  }

  /* The range a command works on: the selection, or the whole document. A
     caret on a formula means that formula. */
  function commandScope(wholeDocument) {
    var doc = app().ActiveDocument;
    if (wholeDocument) return { doc: doc, start: doc.Content.Start, end: doc.Content.End };
    var selection = app().ActiveWindow.Selection;
    var scope = { doc: doc, start: selection.Start, end: selection.End };
    if (scope.start === scope.end) {
      var target = safe(selectionTarget, { kind: "none" });
      var range = target.kind === "image" ? target.shape.Range : target.kind === "omml" ? target.omath.Range : null;
      if (range) {
        scope.start = range.Start;
        scope.end = range.End;
      }
    }
    return scope;
  }

  function withinScope(range, scope) {
    return range.Start >= scope.start && range.End <= scope.end;
  }

  /* VTNormalizePlainWordParagraph: a numbered formula's paragraph becomes a
     plain body paragraph. */
  function normalizePlainParagraph(doc, paragraph) {
    var size = normalStyleFontSize(doc);
    safe(function () { paragraph.Style = doc.Styles.Item(-1); });
    var font = paragraph.Font;
    safe(function () { font.Position = 0; font.Hidden = false; font.Color = WD_COLOR_AUTOMATIC; font.Size = size; });
    var format = paragraph.ParagraphFormat;
    var values = {
      Alignment: WD_ALIGN_PARAGRAPH_LEFT, LeftIndent: 0, RightIndent: 0, FirstLineIndent: 0, SpaceBefore: 0,
      SpaceAfter: 0, LineSpacingRule: 0, KeepTogether: false, KeepWithNext: false, PageBreakBefore: false
    };
    Object.keys(values).forEach(function (name) {
      safe(function () { format[name] = values[name]; });
    });
    safe(function () { format.TabStops.ClearAll(); });
  }

  /* Only tabs, spaces and the like (VTWordRangeHasMeaningfulText). */
  function hasMeaningfulText(range) {
    return String(safe(function () { return range.Text; }, "") || "")
      .replace(/[\x01\t\r\n\x07\u00a0\u200b\u2060 ]/g, "").length > 0;
  }

  /* Puts text where the anchor bookmark (the character after a removed
     formula) is, outside any math zone. */
  function insertTextBeforeAnchor(doc, anchor, text) {
    var at = bookmarkRange(doc, anchor);
    if (!at) throw new Error("The place of a formula was lost");
    var position = at.Start;
    doc.Range(position, position).InsertBefore(text);
    return doc.Range(position, position + text.length);
  }

  function restorePictureToLatex(doc, item) {
    var shape = formulaShape(doc, item.formulaId, item.start);
    if (!shape) throw new Error("A picture formula was removed while it was being restored");
    var compactId = String(item.formulaId).replace(/-/g, "");
    var paragraph = shape.Range.Paragraphs.Item(1).Range;
    if (item.numbered) {
      /* The number goes: everything after the picture that is only the old
         number, tabs and spaces. */
      var tail = doc.Range(shape.Range.End, paragraph.End - 1);
      var placeRef = placeRefOf(shape);
      var number = placeRef ? fieldRange(doc, placeRef) : null;
      var outside = number
        ? String(doc.Range(tail.Start, number.Start).Text || "") + String(doc.Range(number.End, tail.End).Text || "")
        : String(tail.Text || "");
      if (!/[^\t \u00a0\u200b\u2060\x0b]/.test(outside)) tail.Delete();
      deleteBookmark(doc, "VT_N_" + compactId);
      deleteBookmark(doc, "VT_R_" + compactId);
      shape = formulaShape(doc, item.formulaId, item.start);
    }
    var anchor = "VT_X_R" + compactId.slice(0, 8);
    deleteBookmark(doc, anchor);
    doc.Bookmarks.Add(anchor, doc.Range(shape.Range.End, shape.Range.End + 1));
    try {
      shape.Range.Delete();
      deleteBookmark(doc, "VT_I_" + compactId);
      var restored = insertTextBeforeAnchor(doc, anchor, item.text);
      if (item.numbered) {
        var line = restored.Paragraphs.Item(1).Range;
        var before = doc.Range(line.Start, restored.Start);
        if (!hasMeaningfulText(before)) before.Delete();
        normalizePlainParagraph(doc, restored.Paragraphs.Item(1).Range);
      }
    } finally {
      deleteBookmark(doc, anchor);
    }
  }

  function restoreEquationToLatex(doc, marker, item) {
    var omath = markedEquation(doc, marker);
    if (!omath) throw new Error("An equation was removed while it was being restored");
    var range = omath.Range;
    var anchor = marker + "_P";
    deleteBookmark(doc, anchor);
    doc.Bookmarks.Add(anchor, doc.Range(range.End, range.End + 1));
    try {
      range.Delete();
      var restored = insertTextBeforeAnchor(doc, anchor, item.text);
      safe(function () { restored.Font.Position = 0; });
      if (item.numbered) normalizePlainParagraph(doc, restored.Paragraphs.Item(1).Range);
    } finally {
      deleteBookmark(doc, anchor);
    }
  }

  function restorePicturesToLatex(wholeDocument) {
    var scope = commandScope(wholeDocument);
    var doc = scope.doc;
    var items = [];
    var shapes = doc.InlineShapes;
    for (var index = 1; index <= shapes.Count; index++) {
      var shape = shapes.Item(index);
      if (!isFormulaPicture(shape) || !withinScope(shape.Range, scope)) continue;
      var reference = parseFormulaReference(shape.Title);
      items.push({
        start: shape.Range.Start, formulaId: reference.formulaId, displayMode: reference.displayMode,
        numbered: reference.numbered, latex: storedLatex(doc, reference.formulaId),
        metadata: String(shape.AlternativeText || "")
      });
    }
    if (!items.length) {
      notify(wholeDocument ? "全文中没有可恢复的 VisualTeX 图片公式。" : "请先选中要恢复为 LaTeX 的 VisualTeX 图片公式。");
      return null;
    }
    var missing = items.filter(function (item) { return !item.latex; });
    var lookup = missing.length
      ? call("POST", "/formulas/latex", { metadata: missing.map(function (item) { return item.metadata; }) }, 30000)
        .then(function (result) { missing.forEach(function (item, index) { item.latex = result.items[index].latex || null; }); })
      : Promise.resolve();
    return lookup.then(function () {
      var failures = [];
      items.sort(function (a, b) { return b.start - a.start; });
      withUndo("VisualTeX 将图片公式转为 LaTeX", function () {
        items.forEach(function (item) {
          try {
            if (!item.latex) throw new Error("这个图片公式没有可恢复的 LaTeX");
            item.text = item.displayMode === "block" ? "$$" + item.latex + "$$" : "$" + item.latex + "$";
            restorePictureToLatex(doc, item);
          } catch (error) {
            failures.push(error && error.message || String(error));
          }
        });
        if (items.some(function (item) { return item.numbered; })) updateImageNumbersFrom(doc, 0);
      });
      if (failures.length) notify("VisualTeX 有 " + failures.length + " 个图片公式未能恢复：" + failures[0]);
    }).catch(function (error) {
      notify("VisualTeX 恢复 LaTeX 失败：" + error.message);
    });
  }

  function restoreEquationsToLatex(wholeDocument) {
    var scope = commandScope(wholeDocument);
    var doc = scope.doc;
    var items = [];
    var maths = doc.OMaths;
    for (var index = 1; index <= maths.Count; index++) {
      var omath = maths.Item(index);
      var range = omath.Range;
      if (!withinScope(range, scope) || safe(function () { return range.InlineShapes.Count; }, 0) > 0) continue;
      var text = String(safe(function () { return range.Text; }, "") || "");
      items.push({
        start: range.Start,
        displayMode: safe(function () { return equationDisplayMode(doc, omath); }, "inline"),
        numbered: text.indexOf("#") >= 0 && safe(function () { return range.Fields.Count; }, 0) > 0,
        xml: String(range.WordOpenXML)
      });
    }
    if (!items.length) {
      notify(wholeDocument ? "全文中没有 OMML 公式。" : "请先选中要恢复为 LaTeX 的 OMML 公式。");
      return null;
    }
    return call("POST", "/omml/latex", { items: items.map(function (item) {
      return { xml: item.xml, displayMode: item.numbered ? "block" : item.displayMode };
    }) }, 60000).then(function (result) {
      var failures = [];
      items.forEach(function (item, index) {
        item.text = result.items[index].text || null;
        item.error = result.items[index].error || null;
        item.marker = "VT_X_" + index;
      });
      withUndo("VisualTeX 将 OMML 公式转为 LaTeX", function () {
        removeMarkers(doc);
        if (items.some(function (item) { return item.numbered; })) repairAllNumberFields(doc);
        items.forEach(function (item) {
          var omath = equationAt(doc, item.start);
          if (omath) doc.Bookmarks.Add(item.marker, omath.Range);
        });
        items.slice().sort(function (a, b) { return b.start - a.start; }).forEach(function (item) {
          try {
            if (!item.text) throw new Error(item.error || "VisualTeX could not read the equation");
            restoreEquationToLatex(doc, item.marker, item);
          } catch (error) {
            failures.push(error && error.message || String(error));
          }
        });
        removeMarkers(doc);
        if (items.some(function (item) { return item.numbered; })) updateAllNumberFields(doc);
      });
      if (failures.length) notify("VisualTeX 有 " + failures.length + " 个公式未能恢复：" + failures[0]);
    }).catch(function (error) {
      notify("VisualTeX 恢复 LaTeX 失败：" + error.message);
    });
  }

  /* ---- document operations -------------------------------------------------
     "全文格式转换", "LaTeX 重绘" and "批量导入" work as the Word add-in's do
     (VTStartWordFormulaRestore, VTStartWordLatexRedraw,
     VisualTeX_InsertLatexMarkdownDocument): a VT_D_ bookmark marks the target
     range or the insertion point, VisualTeX renders every formula at once in
     its own window, and the add-in writes them all as one Undo step. A
     numbered formula is written without its number; then all of them are
     numbered together in the document's style, so each series counts on as
     Word's does, and cross-references follow their formulas. */

  var MAX_DOCUMENT_FORMULAS = 1000;
  var MAX_REDRAW_SOURCE_BYTES = 5 * 1024 * 1024;
  /* Stands for an imported formula until the text around it is written. */
  var IMPORT_PLACEHOLDER = "\ue000";

  function base64UrlUtf8(text) {
    var bytes = new TextEncoder().encode(String(text));
    var binary = "";
    for (var index = 0; index < bytes.length; index += 0x8000) {
      binary += String.fromCharCode.apply(null, bytes.subarray(index, index + 0x8000));
    }
    return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  }

  function utf8Length(text) {
    return new TextEncoder().encode(String(text)).length;
  }

  function documentAnchor(sessionId) {
    return "VT_D_" + sessionId.slice(0, 24).replace(/-/g, "");
  }

  /* An Undo record that may stay open while VisualTeX numbers the formulas
     (WPS keeps one record open across the calls and folds the records opened
     inside it into it). */
  function openUndo(label) {
    var recording = safe(function () { wps.UndoRecord.StartCustomRecord(label); return true; }, false);
    return {
      recording: recording,
      end: function () {
        if (recording) safe(function () { wps.UndoRecord.EndCustomRecord(); });
        recording = false;
      }
    };
  }

  function preferredFontSize(doc, range) {
    var size = fontSizeAt(range);
    return validFormulaSize(size) ? size : normalStyleFontSize(doc);
  }

  function startDocumentSession(spec) {
    if (active) {
      log("session=" + active.sessionId + " replaced by a document operation");
      forgetActive(true);
    }
    var sessionId = spec.sessionId;
    active = {
      sessionId: sessionId,
      docKey: documentKey(spec.doc),
      anchor: spec.anchor,
      document: true,
      operation: spec.operation,
      targets: spec.targets || null,
      agent: !!agentRun
    };
    var body = {
      sessionId: sessionId,
      operation: spec.operation,
      sourceDocumentId: active.docKey,
      bookmarkName: spec.anchor,
      defaultFontSizePt: spec.defaultFontSizePt
    };
    ["redrawScope", "outputKind", "sourceKind", "source"].forEach(function (key) {
      if (spec[key] !== undefined && spec[key] !== null) body[key] = spec[key];
    });
    /* An agent's operation: VisualTeX opens an import window with the agent's text itself, and the other
       windows tell the agent what they skipped instead of showing it. */
    if (agentRun) body.agent = true;
    var posting = call("POST", "/documents", body, 60000);
    if (agentRun) agentRun.started = posting.then(function (reply) {
      reply.sessionId = reply.sessionId || sessionId;
      return reply;
    });
    posting.then(function () {
      poll(sessionId, 0);
    }, function (error) {
      if (active && active.sessionId === sessionId) forgetActive(true);
      notify("VisualTeX 无法开始这项操作：" + error.message);
    });
  }

  /* -- positions of text --
     Range.Text leaves field instructions out (a range that starts inside one
     reads it, one that ends inside one does not), so after a field a text
     offset no longer is a position (VTResolveLatexRedrawWordOffsets). */

  function textLength(doc, start, end) {
    if (end <= start) return 0;
    return String(safe(function () { return doc.Range(start, end).Text; }, "") || "").length;
  }

  /* The first position p after `from` whose text from `from` holds `units`
     characters. */
  function firstPositionReaching(doc, from, units, limit) {
    if (units <= 0) return from;
    var low = from + units - 1;
    if (low > limit) return -1;
    if (textLength(doc, from, low) >= units) low = from;
    var high = Math.min(limit, from + units);
    var step = 1;
    for (var have = textLength(doc, from, high); have < units; have = textLength(doc, from, high)) {
      if (high >= limit) return -1;
      low = high;
      step = Math.max(step * 2, units - have);
      high = Math.min(limit, high + step);
    }
    while (high - low > 1) {
      var middle = Math.floor((low + high) / 2);
      if (textLength(doc, from, middle) >= units) high = middle;
      else low = middle;
    }
    return high;
  }

  /* The text `text`, `units` characters of Range.Text after `from`. */
  function textSpanAt(doc, from, units, text, limit) {
    var after = firstPositionReaching(doc, from, units + 1, limit);
    if (after < 0) return null;
    var start = after - 1;
    var end = firstPositionReaching(doc, start, text.length, limit);
    if (end < 0) return null;
    return String(safe(function () { return doc.Range(start, end).Text; }, "") || "") === text
      ? { start: start, end: end } : null;
  }

  /* Each source range of a redraw, in order, as a document range reading the
     same text; one that moved is looked for after the one before it. */
  function resolveTextSpans(doc, start, end, ranges) {
    var spans = [];
    var from = start;
    var offset = 0;
    ranges.forEach(function (range, index) {
      var text = String(range.sourceText || "");
      var span = text ? textSpanAt(doc, from, range.sourceStart - offset, text, end) : null;
      if (!span && text) {
        var visible = String(safe(function () { return doc.Range(from, end).Text; }, "") || "");
        var found = visible.indexOf(text);
        if (found >= 0) span = textSpanAt(doc, from, found, text, end);
      }
      if (!span) throw new Error("第 " + (index + 1) + " 个 LaTeX 公式的文字已改变，请重新选择后再试");
      spans.push(span);
      from = span.end;
      offset = range.sourceEnd;
    });
    return spans;
  }

  /* -- the size of a redrawn formula (VTInheritedLatexRedrawFontSize) -- */

  function visibleBodyText(text) {
    return /[^ \t\r\n\x07\x0b\x0c\u00a0]/.test(String(text || ""));
  }

  function insideSpans(position, spans) {
    return spans.some(function (span) { return position >= span.start && position < span.end; });
  }

  function nearbyFontSize(doc, start, minimum, maximum, step, spans) {
    var lower = Math.max(minimum, doc.Content.Start);
    var upper = Math.min(maximum, doc.Content.End - 1);
    for (var probe = 0, position = start; probe < 256 && position >= lower && position <= upper; probe++, position += step) {
      if (insideSpans(position, spans)) continue;
      var range = doc.Range(position, position + 1);
      if (!visibleBodyText(safe(function () { return range.Text; }, ""))) continue;
      var size = Number(safe(function () { return range.Font.Size; }, 0));
      if (validFormulaSize(size)) return size;
    }
    return 0;
  }

  function redrawFontSize(doc, spans, index, displayMode, preset) {
    if (validFormulaSize(preset)) return preset;
    var span = spans[index];
    var paragraph = doc.Range(span.start, span.end).Paragraphs.Item(1).Range;
    var bodyEnd = paragraph.End - 1;
    var beside = visibleBodyText(safe(function () { return doc.Range(paragraph.Start, span.start).Text; }, "")) ||
      visibleBodyText(safe(function () { return doc.Range(span.end, bodyEnd).Text; }, ""));
    var size = 0;
    if (displayMode === "inline" || beside) {
      size = nearbyFontSize(doc, span.start - 1, paragraph.Start, bodyEnd, -1, spans) ||
        nearbyFontSize(doc, span.end, paragraph.Start, bodyEnd, 1, spans);
    }
    size = size || nearbyFontSize(doc, paragraph.Start - 1, doc.Content.Start, doc.Content.End - 1, -1, spans) ||
      nearbyFontSize(doc, paragraph.End, doc.Content.Start, doc.Content.End - 1, 1, spans);
    if (size) return size;
    var styleSize = Number(safe(function () { return paragraph.Style.Font.Size; }, 0));
    if (validFormulaSize(styleSize)) return styleSize;
    return preferredFontSize(doc, doc.Range(span.start, span.end));
  }

  /* The redraw renderer asks for the size of each formula it found
     (VTResolveWordLatexRedrawFontsDispatch). */
  function answerFontSizes(sessionId, ranges) {
    var current = active;
    call("GET", "/conversion-font-sizes", undefined, 5000).then(null, function () {
      return { inline: 0, display: 0 };
    }).then(function (presets) {
      var body;
      try {
        var doc = findDocument(current.docKey);
        if (!doc) throw new Error("文档已经关闭");
        var target = bookmarkRange(doc, current.anchor);
        if (!target) throw new Error("要重绘的范围已不在文档中");
        var spans = resolveTextSpans(doc, target.Start, target.End, ranges);
        body = { sizes: ranges.map(function (range, index) {
          return redrawFontSize(doc, spans, index, range.displayMode, range.displayMode === "block" ? presets.display : presets.inline);
        }) };
      } catch (error) {
        body = { error: "WPS 无法读取公式字号：" + (error && error.message || error) };
      }
      return call("POST", "/sessions/" + sessionId + "/font-sizes", body, 10000);
    }).then(function () {
      poll(sessionId, 0);
    }, function (error) {
      log("session=" + sessionId + " font sizes not sent: " + error.message);
      poll(sessionId, 1);
    });
  }

  /* -- formulas of one kind in a range (VTCollectFormulaRestoreTargets) -- */

  function restoreTargets(doc, start, end, sourceKind) {
    var targets = [];
    if (sourceKind === "omml") {
      var maths = doc.OMaths;
      for (var index = 1; index <= maths.Count; index++) {
        var omath = maths.Item(index);
        var range = omath.Range;
        if (range.Start < start || range.End > end) continue;
        if (safe(function () { return range.InlineShapes.Count; }, 0) > 0) continue;
        if (!String(safe(function () { return range.Text; }, "") || "")) continue;
        targets.push({ start: range.Start, end: range.End, omath: omath });
      }
    } else {
      var shapes = doc.InlineShapes;
      for (var k = 1; k <= shapes.Count; k++) {
        var shape = shapes.Item(k);
        if (!isFormulaPicture(shape)) continue;
        var shapeRange = shape.Range;
        if (shapeRange.Start < start || shapeRange.End > end) continue;
        targets.push({ start: shapeRange.Start, end: shapeRange.End, shape: shape,
          reference: parseFormulaReference(shape.Title) });
      }
    }
    targets.sort(function (a, b) { return a.start - b.start; });
    return targets;
  }

  /* The w:document part of an equation's export: the whole package is
     styles, theme and fonts many times over (VTFormulaRestoreNativeXml). */
  function documentPart(xml) {
    var text = String(xml || "");
    var start = text.indexOf("<w:document ");
    if (start < 0) start = text.indexOf("<w:document>");
    var end = start >= 0 ? text.indexOf("</w:document>", start) : -1;
    return start >= 0 && end > start ? text.slice(start, end + "</w:document>".length) : text;
  }

  /* VTFormulaRestoreManifest: what the renderer needs of each formula. */
  function restoreManifest(doc, sessionId, anchor, start, end, sourceKind, targets, presets) {
    var lines = [
      "protocolVersion=1",
      "sessionId=" + sessionId,
      "sourceDocumentId=" + base64UrlUtf8(documentKey(doc)),
      "bookmarkName=" + anchor,
      "sourceKind=" + sourceKind,
      "targetTextBase64=" + base64UrlUtf8(String(doc.Range(start, end).Text || "")),
      "itemCount=" + targets.length
    ];
    targets.forEach(function (target, index) {
      var prefix = "item" + index;
      var range = doc.Range(target.start, target.end);
      var displayMode;
      var numbered;
      var fontSize;
      var formulaId = "";
      var metadata = "";
      var payload;
      if (sourceKind === "omml") {
        var xml = String(range.WordOpenXML || "");
        displayMode = equationDisplayMode(doc, target.omath);
        numbered = displayMode === "block" && isNumberedEquationXml(xml);
        fontSize = preferredFontSize(doc, range);
        payload = documentPart(xml);
      } else {
        displayMode = target.reference.displayMode;
        numbered = displayMode === "block" && target.reference.numbered;
        formulaId = target.reference.formulaId;
        metadata = String(target.shape.AlternativeText || "");
        fontSize = formulaPictureSize(doc, target.shape) || preferredFontSize(doc, range);
        payload = metadata;
      }
      /* A format conversion gives the formula its kind's conversion size. */
      var preset = displayMode === "block" ? presets.display : presets.inline;
      if (validFormulaSize(preset)) fontSize = preset;
      target.displayMode = displayMode;
      target.numbered = numbered;
      target.formulaId = formulaId;
      target.sourceStart = target.start - start;
      lines.push(
        prefix + "sourceStart=" + (target.start - start),
        prefix + "sourceEnd=" + (target.end - start),
        prefix + "sourceTextBase64=" + base64UrlUtf8(String(range.Text || "")),
        prefix + "displayMode=" + displayMode,
        prefix + "fontSizePt=" + fontSize,
        prefix + "formulaId=" + formulaId,
        prefix + "numbered=" + (numbered ? "1" : "0"),
        prefix + "metadata=" + metadata,
        prefix + "payloadBase64=" + base64UrlUtf8(payload));
    });
    return lines.join("\n") + "\n";
  }

  /* "全文图片公式 → Word OMML" and "全文 OMML → 图片公式"
     (VTStartWordFormulaRestore "document"). */
  function convertDocument(sourceKind, selectionOnly) {
    var doc = safe(function () { return app().ActiveDocument; }, null);
    if (!doc) {
      notify("请先打开一个 WPS 文字文档。");
      return;
    }
    /* An agent can convert the selection's formulas only. */
    var bounds = selectionOnly ? commandScope(false) : null;
    var scopeOf = function () { return bounds || { start: doc.Content.Start, end: doc.Content.End }; };
    var start = scopeOf().start;
    var end = scopeOf().end;
    var targets = restoreTargets(doc, start, end, sourceKind);
    if (!targets.length) {
      notify(sourceKind === "omml" ? "当前文档中没有可转换的 OMML 公式。" : "当前文档中没有可转换的 VisualTeX 图片公式。");
      return;
    }
    if (targets.length > MAX_DOCUMENT_FORMULAS) {
      notify("当前文档中的公式超过 " + MAX_DOCUMENT_FORMULAS + " 个，请分段转换。");
      return;
    }
    call("GET", "/conversion-font-sizes", undefined, 5000).then(null, function () {
      return { inline: 0, display: 0 };
    }).then(function (presets) {
      run(function () {
        start = scopeOf().start;
        end = scopeOf().end;
        targets = restoreTargets(doc, start, end, sourceKind);
        if (!targets.length) return;
        var sessionId = newUuid();
        var anchor = documentAnchor(sessionId);
        var manifest = restoreManifest(doc, sessionId, anchor, start, end, sourceKind, targets, presets);
        if (utf8Length(manifest) > 16 * 1024 * 1024) throw new Error("文档中的公式太多，超过 16 MB，请分段转换");
        deleteBookmark(doc, anchor);
        doc.Bookmarks.Add(anchor, doc.Range(start, end));
        startDocumentSession({
          sessionId: sessionId, doc: doc, anchor: anchor, operation: "formulaRestore",
          redrawScope: bounds ? "selection" : "document", sourceKind: sourceKind, outputKind: sourceKind === "omml" ? "image" : "omml",
          source: manifest, defaultFontSizePt: preferredFontSize(doc, doc.Range(start, end)),
          targets: targets.map(function (target) {
            return { sourceStart: target.sourceStart, formulaId: target.formulaId, numbered: target.numbered };
          })
        });
      });
    });
  }

  /* "重绘所选" and "重绘全文" (VTStartWordLatexRedraw). */
  function redrawLatex(wholeDocument, outputKind) {
    var doc = safe(function () { return app().ActiveDocument; }, null);
    if (!doc) {
      notify("请先打开一个 WPS 文字文档。");
      return;
    }
    var range = wholeDocument ? doc.Content : app().ActiveWindow.Selection.Range;
    if (!wholeDocument && range.Start === range.End) {
      notify("请先选中包含 LaTeX 代码的文字，再重绘所选。");
      return;
    }
    var text = String(range.Text || "");
    if (!text.replace(/[\s\x07]/g, "")) {
      notify("所选内容中没有可扫描的文字。");
      return;
    }
    if (utf8Length(text) > MAX_REDRAW_SOURCE_BYTES) {
      notify("要重绘的内容超过 5 MB，请选择较小的范围。");
      return;
    }
    var sessionId = newUuid();
    var anchor = documentAnchor(sessionId);
    deleteBookmark(doc, anchor);
    doc.Bookmarks.Add(anchor, doc.Range(range.Start, range.End));
    startDocumentSession({
      sessionId: sessionId, doc: doc, anchor: anchor, operation: "latexRedraw",
      redrawScope: wholeDocument ? "document" : "selection", outputKind: outputKind, source: text,
      defaultFontSizePt: preferredFontSize(doc, range)
    });
  }

  /* "批量导入" (VisualTeX_InsertLatexMarkdownDocument): the import window
     opens; the collapsed anchor is where the document goes. */
  function importDocument() {
    var doc = safe(function () { return app().ActiveDocument; }, null);
    if (!doc) {
      notify("请先打开一个 WPS 文字文档。");
      return;
    }
    var selection = app().ActiveWindow.Selection;
    var position = selection.Start;
    var sessionId = newUuid();
    var anchor = documentAnchor(sessionId);
    deleteBookmark(doc, anchor);
    doc.Bookmarks.Add(anchor, doc.Range(position, position));
    var size = Number(safe(function () { return selection.Font.Size; }, 0));
    startDocumentSession({
      sessionId: sessionId, doc: doc, anchor: anchor, operation: "documentImport",
      defaultFontSizePt: validFormulaSize(size) ? size : normalStyleFontSize(doc)
    });
  }

  /* -- writing the result -- */

  /* What insertPicture and finishEquation take for one rendered item. */
  function itemPayload(item, numbered) {
    return {
      formulaId: item.formulaId, displayMode: item.displayMode, numbered: numbered,
      nativeEquation: item.nativeEquation, fontSizePt: Number(item.fontSizePt),
      widthPoints: Number(item.widthPoints), heightPoints: Number(item.heightPoints),
      referenceWidthPt: Number(item.referenceWidthPt), referenceHeightPt: Number(item.referenceHeightPt),
      referenceBaselinePt: Number(item.referenceBaselinePt), inkCenterYRatio: item.inkCenterYRatio,
      metadata: item.metadata, formulaReference: item.formulaReference, ownerBookmark: item.ownerBookmark,
      formulaLetterFont: item.formulaLetterFont || "", files: item.files || {}, variables: item.variables
    };
  }

  /* Writes one rendered formula at position; returns its range. A numbered
     one is written without its number and listed for numbering. */
  function writeItemFormula(doc, sessionId, item, position, numbered, references) {
    var payload = itemPayload(item, false);
    var formulaRange;
    if (item.nativeEquation) {
      var path = download(sessionId, payload.files.native);
      try {
        insertEquationFile(doc, path, position);
      } finally {
        removeFile(path);
      }
      var end = finishEquation(doc, position, payload);
      formulaRange = doc.Range(position, end);
      if (item.numbered) {
        var marker = "VT_X_N" + numbered.length;
        deleteBookmark(doc, marker);
        doc.Bookmarks.Add(marker, formulaRange);
        restoreBookmarks(doc, references, { start: formulaRange.Start, end: formulaRange.End });
        numbered.push({ kind: "equation", marker: marker });
      }
    } else {
      var shape = insertPicture(doc, sessionId, payload, doc.Range(position, position));
      if (payload.variables) writeVariables(doc, payload.variables);
      deleteBookmark(doc, payload.ownerBookmark);
      doc.Bookmarks.Add(payload.ownerBookmark, shape.Range);
      formulaRange = shape.Range;
      if (item.numbered) {
        restoreBookmarks(doc, references, { start: formulaRange.Start, end: formulaRange.End });
        numbered.push({ kind: "image", formulaId: item.formulaId });
      }
    }
    return formulaRange;
  }

  /* A numbered picture's number goes with it, and the paragraph's spacing
     if nothing else is left (VTNumberedImageReplacementRange). */
  function pictureReplacementRange(doc, shape) {
    var range = { start: shape.Range.Start, end: shape.Range.End, cleared: false };
    var placeRef = placeRefOf(shape);
    if (!placeRef || inTable(shape.Range)) return range;
    var number = fieldRange(doc, placeRef);
    var paragraph = shape.Range.Paragraphs.Item(1).Range;
    var contentEnd = paragraph.End - 1;
    if (number.Start < range.end || number.End > contentEnd) return range;
    if (hasMeaningfulText(doc.Range(range.end, number.Start))) return range;
    range.end = number.End;
    if (!hasMeaningfulText(doc.Range(paragraph.Start, range.start)) &&
        doc.Range(paragraph.Start, range.start).InlineShapes.Count === 0) range.start = paragraph.Start;
    if (!hasMeaningfulText(doc.Range(number.End, contentEnd))) range.end = contentEnd;
    range.cleared = range.start === paragraph.Start && range.end === contentEnd;
    return range;
  }

  /* The document variables and bookmarks of a picture formula that is no
     longer one (VTDeleteDocumentImportedFormulaState). */
  function deletePictureState(doc, formulaId) {
    var key = String(formulaId).replace(/-/g, "_");
    var compactId = String(formulaId).replace(/-/g, "");
    ["VT_Latex_", "VT_OMML_", "VT_Metadata_"].forEach(function (prefix) {
      var stem = prefix + key;
      var count = Number(variableValue(doc, stem + "_Count")) || 0;
      for (var index = 1; index <= Math.min(count, 128); index++) deleteVariable(doc, stem + "_" + ("000" + index).slice(-3));
      deleteVariable(doc, stem + "_Count");
    });
    ["VT_Format_", "VT_ImageScale_", "VT_ImageCanvas_", "VT_ImageInkCenter_", "VT_ImageLetterFont_"].forEach(function (prefix) {
      deleteVariable(doc, prefix + key);
    });
    ["VT_I_", "VT_N_", "VT_R_", "VT_C_"].forEach(function (prefix) { deleteBookmark(doc, prefix + compactId); });
  }

  /* One source formula or LaTeX text of a range replacement: removed, and
     the rendered formula written in its place. */
  function replaceItem(doc, sessionId, payload, item, marker, numbered) {
    var range = bookmarkRange(doc, marker);
    if (!range) throw new Error("一个要替换的公式已不在文档中");
    var references = [];
    var cleared = false;
    var start = range.Start;
    var end = range.End;
    if (payload.operation === "latexRedraw") {
      if (String(range.Text || "") !== String(item.sourceText)) throw new Error("WPS 中的 LaTeX 文字已改变，请重新选择后再试");
    } else if (payload.sourceKind === "image") {
      var shape = range.InlineShapes.Count >= 1 ? range.InlineShapes.Item(1) : null;
      var reference = shape ? parseFormulaReference(shape.Title) : null;
      if (!reference || reference.formulaId !== item.formulaId) throw new Error("一个图片公式在转换时被改动了");
      if (reference.numbered) {
        references = imageNumberReferences(doc, shape);
        var replacement = pictureReplacementRange(doc, shape);
        start = replacement.start;
        end = replacement.end;
        cleared = replacement.cleared;
      }
      /* Its bookmarks go; its document variables stay: WPS's Undo does not
         bring deleted variables back, and the picture would come back
         without its LaTeX and size. */
      var compactId = String(reference.formulaId).replace(/-/g, "");
      ["VT_I_", "VT_N_", "VT_R_", "VT_C_"].forEach(function (prefix) { deleteBookmark(doc, prefix + compactId); });
    } else {
      var omath = equationAt(doc, range.Start);
      if (!omath) throw new Error("一个 OMML 公式在转换时被改动了");
      start = omath.Range.Start;
      end = omath.Range.End;
      if (item.numbered) references = numberBookmarks(doc, start, end);
    }
    /* The character after it finds the place again: removing a number makes
       WPS update the other numbers. */
    var place = marker + "_P";
    deleteBookmark(doc, place);
    doc.Bookmarks.Add(place, doc.Range(end, end + 1));
    try {
      doc.Range(start, end).Delete();
      var at = bookmarkRange(doc, place);
      var position = at ? at.Start : start;
      if (cleared) normalizePlainParagraph(doc, doc.Range(position, position).Paragraphs.Item(1).Range);
      writeItemFormula(doc, sessionId, item, position, numbered, references);
    } finally {
      deleteBookmark(doc, place);
    }
  }

  /* LaTeX redraw and format conversion (VTCommitWordLatexRedrawDispatch):
     the items replace their sources from the last one back. */
  function replaceItems(doc, sessionId, payload, current, start, end, numbered, progress) {
    var items = payload.items;
    var markers = [];
    if (payload.operation === "latexRedraw") {
      var spans = resolveTextSpans(doc, start, end, items);
      spans.forEach(function (span, index) {
        markers.push("VT_X_" + index);
        doc.Bookmarks.Add(markers[index], doc.Range(span.start, span.end));
      });
    } else {
      /* A formula conversion: the formulas found when it started, found again. */
      var live = restoreTargets(doc, start, end, payload.sourceKind);
      if (!current.targets || live.length !== current.targets.length) {
        throw new Error("转换期间文档中的公式有增减，未作修改");
      }
      repairAllNumberFields(doc);
      items.forEach(function (item, index) {
        var found = -1;
        current.targets.forEach(function (target, position) {
          if (target.sourceStart === item.sourceStart) found = position;
        });
        if (found < 0) throw new Error("转换结果与文档中的公式对不上，未作修改");
        markers.push("VT_X_" + index);
        doc.Bookmarks.Add(markers[index], doc.Range(live[found].start, live[found].end));
      });
    }
    for (var index = items.length - 1; index >= 0; index--) {
      replaceItem(doc, sessionId, payload, items[index], markers[index], numbered);
      deleteBookmark(doc, markers[index]);
      progress(items.length - index);
    }
  }

  /* -- batch import (VTCommitWordDocumentImportDispatch) --
     The text goes in first, with a placeholder for each formula, then the
     formulas take their placeholders' places from the last one back: text
     written next to an equation would go into it. */

  var IMPORT_STYLES = { normal: -1, heading1: -2, heading2: -3, heading3: -4, heading4: -5, quote: -1, code: -1 };
  var IMPORT_ALIGNMENTS = { left: 0, center: 1, right: 2, justify: 3 };

  /* VTApplyDocumentImportParagraphFormat, on the paragraph at position. */
  function formatImportParagraph(doc, position, paragraph) {
    /* A caret at a paragraph's start can report the paragraph before it. */
    var range = doc.Range(position, position + 1).Paragraphs.Item(1).Range;
    var style = paragraph.style || "normal";
    safe(function () { range.Style = doc.Styles.Item(IMPORT_STYLES[style] || -1); });
    var format = range.ParagraphFormat;
    if (style === "quote") {
      safe(function () { format.LeftIndent = 36; format.RightIndent = 36; });
    }
    safe(function () { format.Alignment = IMPORT_ALIGNMENTS[paragraph.alignment] || 0; });
    /* ApplyBulletDefault and ApplyNumberDefault take a list away from a
       paragraph that has one of that kind (one split from the item before
       it), so each is applied only to a paragraph without it. */
    var list = range.ListFormat;
    var type = Number(safe(function () { return list.ListType; }, 0));
    var numberedList = type === 1 || type === 3 || type === 4 || type === 5;
    if (paragraph.listKind === "bullet" && type !== 2) {
      if (type) safe(function () { list.RemoveNumbers(); });
      safe(function () { list.ApplyBulletDefault(); });
    } else if (paragraph.listKind === "number" && !numberedList) {
      if (type) safe(function () { list.RemoveNumbers(); });
      safe(function () { list.ApplyNumberDefault(); });
    } else if (paragraph.listKind !== "bullet" && paragraph.listKind !== "number" && type) {
      safe(function () { list.RemoveNumbers(); });
    }
    if (paragraph.listKind === "bullet" || paragraph.listKind === "number") {
      var level = Math.max(1, Math.min(9, Number(paragraph.listLevel) || 1));
      var current = doc.Range(position, position + 1).Paragraphs.Item(1).Range.ListFormat;
      if (!safe(function () { current.ListLevelNumber = level; return true; }, false)) {
        for (var step = Number(safe(function () { return current.ListLevelNumber; }, 1)); step < level; step++) {
          safe(function () { current.ListIndent(); });
        }
      }
    }
    if (style === "normal" || style === "quote" || style === "code") {
      safe(function () { format.SpaceBefore = 0; format.SpaceAfter = 0; });
    }
  }

  /* VTDocumentImportInsertText: one item is one paragraph's text, its line
     breaks manual ones; it is plain text. */
  function insertImportText(doc, position, text, code) {
    var plain = String(text).replace(/\r\n?/g, "\n").replace(/\n/g, "\x0b");
    if (!plain) return position;
    doc.Range(position, position).Text = plain;
    var inserted = doc.Range(position, position + plain.length);
    var font = inserted.Font;
    safe(function () { font.Reset(); });
    safe(function () { font.Bold = 0; font.Italic = 0; font.Underline = 0; font.Position = 0; });
    if (code) safe(function () { font.Name = "Menlo"; font.NameAscii = "Menlo"; font.NameOther = "Menlo"; });
    return position + plain.length;
  }

  /* VTFinalizeDocumentImportParagraph: the paragraph mark after the text, or
     the one already there (not the document's last). */
  function endImportParagraph(doc, position) {
    if (position < doc.Content.End - 1 && String(doc.Range(position, position + 1).Text) === "\r") return position + 1;
    doc.Range(position, position).Text = "\r";
    return position + 1;
  }

  function emptyParagraph(range) {
    return !hasMeaningfulText(range) && range.InlineShapes.Count === 0 && range.OMaths.Count === 0 &&
      safe(function () { return range.Fields.Count; }, 0) === 0;
  }

  /* A display formula's own plain paragraph, its placeholder, and the empty
     paragraph after it where the import goes on
     (VTPrepareBatchUnnumberedDisplayInsertionRange,
     VTBatchContinuationRangeAfterUnnumberedDisplayFormula). */
  function placeImportDisplay(doc, position) {
    var paragraph = doc.Range(position, position + 1).Paragraphs.Item(1).Range;
    if (position > paragraph.Start && !emptyParagraph(doc.Range(paragraph.Start, position))) {
      doc.Range(position, position).Text = "\r";
      position++;
    }
    var own = doc.Range(position, position + 1).Paragraphs.Item(1).Range;
    safe(function () { own.Style = doc.Styles.Item(-1); });
    safe(function () { own.ListFormat.RemoveNumbers(); });
    doc.Range(position, position).Text = IMPORT_PLACEHOLDER;
    var placeholder = position;
    position++;
    var next = position + 1;
    if (position < doc.Content.End - 1 && String(doc.Range(position, next).Text) === "\r" && next < doc.Content.End &&
        emptyParagraph(doc.Range(next, next + 1).Paragraphs.Item(1).Range)) {
      return { placeholder: placeholder, position: next };
    }
    doc.Range(position, position).Text = "\r";
    return { placeholder: placeholder, position: position + 1 };
  }

  function importItems(doc, sessionId, payload, start, numbered, progress) {
    var items = payload.items;
    var position = start;
    var placeholders = [];
    var code = false;
    items.forEach(function (item, index) {
      var paragraph = item.paragraph;
      if (paragraph && paragraph.start) {
        formatImportParagraph(doc, position, paragraph);
        code = paragraph.style === "code";
      }
      if (item.kind === "text") {
        position = insertImportText(doc, position, item.text, code);
      } else if (item.displayMode === "block") {
        var placed = placeImportDisplay(doc, position);
        placeholders.push({ index: index, position: placed.placeholder });
        position = placed.position;
      } else {
        doc.Range(position, position).Text = IMPORT_PLACEHOLDER;
        placeholders.push({ index: index, position: position });
        position++;
      }
      if (paragraph && paragraph.end) {
        position = endImportParagraph(doc, position);
        code = false;
      }
    });
    var before = doc.Content.End;
    var texts = items.length - placeholders.length;
    progress(texts);
    for (var k = placeholders.length - 1; k >= 0; k--) {
      var at = placeholders[k].position;
      if (String(doc.Range(at, at + 1).Text) !== IMPORT_PLACEHOLDER) throw new Error("WPS 改动了导入中的文档");
      doc.Range(at, at + 1).Delete();
      writeItemFormula(doc, sessionId, items[placeholders[k].index], at, numbered, []);
      progress(texts + placeholders.length - k);
    }
    return position + doc.Content.End - before;
  }

  /* The numbered formulas of a document operation, numbered together. */
  function numberWritten(doc, numbered, label) {
    if (!numbered.length) return Promise.resolve(null);
    var equations = [];
    var images = [];
    numbered.forEach(function (entry) {
      if (entry.kind === "equation") {
        var omath = markedEquation(doc, entry.marker);
        if (omath) equations.push({ start: omath.Range.Start });
      } else {
        var shape = formulaShape(doc, entry.formulaId, -1);
        if (shape) images.push({ start: shape.Range.Start, formulaId: entry.formulaId,
          metadata: String(shape.AlternativeText || "") });
      }
    });
    return renumberDocument({ doc: doc, equations: equations, images: images, undoLabel: label });
  }

  var DOCUMENT_LABELS = {
    documentImport: "VisualTeX 批量导入",
    latexRedraw: "VisualTeX 重绘 LaTeX 公式",
    formulaRestore: "VisualTeX 全文格式转换"
  };

  function applyDocumentResult(sessionId, payload) {
    var current = active;
    var doc = findDocument(current.docKey);
    var importing = payload.operation === "documentImport";
    if (!doc) {
      finishDocument(sessionId, { ok: false, error: "WPS 中已找不到开始这项操作的文档" }, false);
      return;
    }
    var anchor = bookmarkRange(doc, current.anchor);
    if (!anchor && payload.operation === "formulaRestore" && payload.redrawScope === "document") anchor = doc.Content;
    if (!anchor) {
      finishDocument(sessionId, { ok: false, error: "文档中要写入的位置已不存在，请重新开始" }, false);
      return;
    }
    var start = anchor.Start;
    var end = anchor.End;
    var originalEnd = doc.Content.End;
    var originalText = importing ? "" : String(anchor.Text || "");
    var total = payload.items.length;
    var reported = 0;
    var progress = function (done) {
      if (done - reported < Math.max(1, Math.floor(total / 10)) && done < total) return;
      reported = done;
      call("POST", "/sessions/" + sessionId + "/progress", { current: Math.min(done, total), total: total }, 3000).catch(function () {});
    };
    /* Removed before the Undo record, which would bring it back. */
    deleteBookmark(doc, current.anchor);
    var label = DOCUMENT_LABELS[payload.operation] || "VisualTeX";
    var record = openUndo(label);
    var numbered = [];
    var caret = null;
    try {
      removeMarkers(doc);
      if (importing) caret = importItems(doc, sessionId, payload, start, numbered, progress);
      else replaceItems(doc, sessionId, payload, current, start, end, numbered, progress);
    } catch (error) {
      /* Nothing stays half written: the record is undone and checked
         (VTCommitWordLatexRedrawDispatch's rollback). */
      var undoable = record.recording;
      record.end();
      var restored = false;
      if (undoable) {
        safe(function () { doc.Undo(1); });
        restored = doc.Content.End === originalEnd &&
          (importing || String(safe(function () { return doc.Range(start, end).Text; }, "") || "") === originalText);
      }
      removeMarkers(doc);
      if (restored && importing) safe(function () { doc.Bookmarks.Add(current.anchor, doc.Range(start, start)); });
      log("session=" + sessionId + " document apply failed: " + (error && error.stack || error));
      var reason = String(error && error.message || error) +
        (restored ? "" : "（WPS 未能自动撤回已写入的部分，请用撤销恢复）");
      finishDocument(sessionId, { ok: false, error: "WPS 未能写入：" + reason }, importing && restored);
      return;
    }
    numberWritten(doc, numbered, label).then(function (summary) {
      return summary && summary.failures.length
        ? "有 " + summary.failures.length + " 个编号未能写入：" + summary.failures[0] : "";
    }, function (error) {
      return "编号未能写入：" + error.message;
    }).then(function (problem) {
      removeMarkers(doc);
      /* Inside the record: the typing font reset is a step of its own. */
      if (caret !== null) safe(function () { doc.Range(caret, caret).Select(); resetTypingFont(); });
      record.end();
      safe(function () { app().StatusBar = "VisualTeX: 已写入 " + total + " 项"; });
      if (problem) notify("VisualTeX 已写入公式，但" + problem);
      finishDocument(sessionId, { ok: true }, false);
    });
  }

  /* An import that failed can be tried again from its window: the add-in
     waits for the next result. */
  function finishDocument(sessionId, result, retry) {
    if (!result.ok) log("session=" + sessionId + " apply-failed " + result.error);
    if (!retry) active = null;
    call("POST", "/sessions/" + sessionId + "/applied", result, 10000).catch(function () {}).then(function () {
      if (retry && active && active.sessionId === sessionId) poll(sessionId, 0);
    });
  }

  /* ---- cross-references ---------------------------------------------------
     "插入公式引用" lists the numbered equations and pictures and inserts
     "(REF)" to the chosen number (VTWordRibbonCrossReferenceMenuContent,
     VTInsertEquationNumberReferenceAtRange). */

  var CUSTOM_UI_NAMESPACE = "http://schemas.microsoft.com/office/2006/01/customui";
  var referenceMenu = null;

  function menuMessage(text) {
    return "<menu xmlns=\"" + CUSTOM_UI_NAMESPACE + "\"><button id=\"vtxRefMessage\" label=\"" +
      xmlText(text) + "\" enabled=\"false\" /></menu>";
  }

  /* The visible number of a numbered equation: its SEQ result, from its
     STYLEREF chapter prefix (VTPureNativeVisibleNumberRange) or from its
     literal prefix ("1.0-", written when the headings are not numbered by a
     list; Word references only the SEQ result then, which reads "(1)" for
     "(1.0-1)"). WPS shows each math structure boundary as a "\r". */
  function literalPrefixStart(doc, fieldStart, mathStart) {
    var start = fieldStart;
    while (start > mathStart) {
      var character = String(safe(function () { return doc.Range(start - 1, start).Text; }, "") || "");
      if (character.length !== 1 || character.charCodeAt(0) < 32 || character === "#" || character === "(") break;
      start--;
    }
    return start;
  }

  function equationNumberTarget(doc, omath) {
    var range = omath.Range;
    var text = String(safe(function () { return range.Text; }, "") || "");
    if (text.indexOf("#") < 0) return null;
    var fields = safe(function () { return range.Fields; }, null);
    if (!fields) return null;
    var sequence = null;
    var prefix = null;
    for (var index = 1; index <= fields.Count; index++) {
      var field = fields.Item(index);
      var type = fieldType(field);
      if (type === WD_FIELD_STYLEREF && !sequence) prefix = field;
      if (type === WD_FIELD_SEQUENCE) {
        sequence = field;
        break;
      }
    }
    if (!sequence) return null;
    /* Through the field's end character: WPS shortens a bookmark that ends
       inside a field result whenever the field is updated. */
    var end = fieldRange(doc, sequence).End;
    var start = prefix ? prefix.Result.Start : literalPrefixStart(doc, sequence.Code.Start - 1, range.Start);
    return { start: start, end: end };
  }

  /* The visible number of a numbered picture: from after "(" to the end of
     its current-value SEQ result (VTCurrentImageNumberRange). */
  function imageNumberTarget(doc, placeRef) {
    var fields = placeRef.Code.Fields;
    var current = null;
    for (var index = 1; index <= fields.Count; index++) {
      if (isImageNumberField(fields.Item(index))) current = fields.Item(index);
    }
    if (!current) return null;
    var fieldStart = current.Code.Start - 1;
    for (var position = fieldStart - 1; position >= placeRef.Code.Start; position--) {
      if (String(doc.Range(position, position + 1).Text) === "(") {
        return { start: position + 1, end: fieldRange(doc, current).End };
      }
    }
    return null;
  }

  function isImagePlaceRef(field) {
    return fieldType(field) === 51 && fieldCode(field).indexOf(IMAGE_PLACE_REF) >= 0;
  }

  /* Numbered equations ("omml") and pictures ("image") in document order
     (VTCurrentNumberTargets). */
  function numberTargets(doc, kind) {
    var targets = [];
    if (kind === "omml") {
      var maths = doc.OMaths;
      for (var index = 1; index <= maths.Count; index++) {
        var target = safe(function () { return equationNumberTarget(doc, maths.Item(index)); }, null);
        if (target) targets.push(target);
      }
    } else {
      var fields = doc.Fields;
      for (var k = 1; k <= fields.Count; k++) {
        var field = fields.Item(k);
        if (!isImagePlaceRef(field)) continue;
        var found = safe(function () { return imageNumberTarget(doc, field); }, null);
        if (found) targets.push(found);
      }
    }
    targets.sort(function (a, b) { return a.start - b.start; });
    return targets.filter(function (target, index) {
      return index === 0 || target.start >= targets[index - 1].end;
    });
  }

  function referenceItemsXml(doc, targets, kind, label) {
    var xml = "<menu id=\"vtxRef_" + kind + "\" label=\"" + xmlText(label) + "\">";
    targets.forEach(function (target, index) {
      var text = "(" + String(doc.Range(target.start, target.end).Text).replace(/[\r\n\t\x07\x13\x14\x15]/g, " ").trim() + ")";
      if (text.length > 180) text = text.slice(0, 177) + "...";
      xml += "<button id=\"vtxRef_" + kind + "_" + index + "\" label=\"" + xmlText(text) +
        "\" onAction=\"OnVisualTeXAction\" />";
    });
    return xml + "</menu>";
  }

  function crossReferenceMenu() {
    referenceMenu = null;
    var doc = safe(function () { return app().ActiveDocument; }, null);
    if (!doc) return menuMessage("请先打开一个文档");
    var equations = numberTargets(doc, "omml");
    var images = numberTargets(doc, "image");
    if (!equations.length && !images.length) return menuMessage("当前文档没有可引用的编号公式");
    var selection = app().ActiveWindow.Selection;
    referenceMenu = {
      docKey: documentKey(doc),
      start: selection.Start,
      end: selection.End,
      targets: { omml: equations, image: images }
    };
    var xml = "<menu xmlns=\"" + CUSTOM_UI_NAMESPACE + "\">";
    if (equations.length) xml += referenceItemsXml(doc, equations, "omml", "Word OMML 公式");
    if (images.length) xml += referenceItemsXml(doc, images, "image", "图片公式");
    return xml + "</menu>";
  }

  /* "(" REF _Ref… \h ")" at the selection, the bookmark on the number; the
     reference looks like body text at the Normal style's size
     (VTFormatBodyEquationReference). */
  function insertReference(kind, index) {
    var menu = referenceMenu;
    referenceMenu = null;
    var doc = menu && findDocument(menu.docKey);
    var target = menu && menu.targets[kind] && menu.targets[kind][index];
    if (!doc || !target) throw new Error("请重新打开公式引用列表后再选择");
    withUndo("VisualTeX 插入公式引用", function () {
      var name = "_Ref" + newUuid().replace(/-/g, "");
      doc.Bookmarks.Add(name, doc.Range(target.start, target.end));
      var start = menu.start;
      doc.Range(start, menu.end).Text = "()";
      var field = doc.Fields.Add(doc.Range(start + 1, start + 1), WD_FIELD_REF, name + " \\h", false);
      safe(function () { field.Update(); });
      var end = field.Result.End + 2;
      var reference = doc.Range(start, end);
      var size = normalStyleFontSize(doc);
      [field.Result, reference].forEach(function (range) {
        var font = range.Font;
        safe(function () { font.Hidden = false; });
        safe(function () { font.Color = WD_COLOR_AUTOMATIC; });
        safe(function () { font.Position = 0; });
        safe(function () { font.Size = size; });
      });
      doc.Range(end, end).Select();
    });
  }

  /* Text typed next goes on as body text: not lowered like an inline
     picture, not into the equation, not into a display formula's line. */
  function resetTypingFont() {
    safe(function () { app().ActiveWindow.Selection.Font.Position = 0; });
  }

  function isInlineFormulaPicture(doc, index) {
    if (index < 0) return false;
    var shapes = doc.Range(index, index + 1).InlineShapes;
    if (shapes.Count !== 1) return false;
    var shape = shapes.Item(1);
    return String(shape.AlternativeText).indexOf(METADATA_PREFIX) === 0 &&
      String(shape.Title).indexOf(":inline:") > 0;
  }

  /* WPS gives text typed next to a picture the picture's character position
     (Word does not), so text typed after an inline VisualTeX formula, and every
     paragraph started from there, sank below the line by the formula's
     descent. Whenever the caret comes to rest next to such a formula, give
     the insertion point position 0 again (WordFormulaHostLayout's
     RestoreInlineCaret does this once on Windows; WPS needs it on every move). */
  function resetCaretNextToFormula() {
    var selection = app().ActiveWindow.Selection;
    if (selection.Start !== selection.End) return;
    var doc = app().ActiveDocument;
    var caret = selection.Start;
    var next = isInlineFormulaPicture(doc, caret - 1) ||
      (selection.Paragraphs.Item(1).Range.Start === caret && isInlineFormulaPicture(doc, caret));
    if (next && Number(selection.Font.Position) !== 0) selection.Font.Position = 0;
  }

  /* VTEnsurePlainContinuationParagraph and VTV2ApplyBodyFormatToTypingParagraph:
     the paragraph below a display formula, a new one when the next paragraph
     is missing or holds a formula, formatted like the body text above. */
  function continuationAfterDisplay(doc, formulaParagraph) {
    var boundary = formulaParagraph.End;
    var next = boundary < doc.Content.End ? doc.Range(boundary, boundary + 1).Paragraphs.Item(1).Range : null;
    if (next && next.Start === boundary && next.InlineShapes.Count === 0 && next.OMaths.Count === 0 &&
        safe(function () { return next.Fields.Count; }, 0) === 0) {
      return doc.Range(boundary, boundary);
    }
    formulaParagraph.InsertParagraphAfter();
    var typing = doc.Range(boundary, boundary + 1).Paragraphs.Item(1).Range;
    var previous = formulaParagraph.Start > 0 ? doc.Range(formulaParagraph.Start - 1, formulaParagraph.Start).Paragraphs.Item(1).Range : null;
    var body = previous && previous.InlineShapes.Count === 0 && previous.OMaths.Count === 0 ? previous.ParagraphFormat : null;
    var format = typing.ParagraphFormat;
    format.Alignment = body ? body.Alignment : WD_ALIGN_PARAGRAPH_LEFT;
    ["LeftIndent", "RightIndent", "FirstLineIndent", "SpaceBefore", "SpaceAfter"].forEach(function (name) {
      safe(function () { format[name] = body ? body[name] : 0; });
    });
    return doc.Range(boundary, boundary);
  }

  /* Where the user goes on after Apply (the Word add-in's caret placement):
     an edited formula stays selected; after a new inline formula the caret
     follows it, after a new display formula it starts the next paragraph. */
  function placeCaret(doc, payload, formulaRange, mode) {
    if (mode === "edit") {
      formulaRange.Select();
      return;
    }
    if (payload.displayMode === "block") {
      continuationAfterDisplay(doc, formulaRange.Paragraphs.Item(1).Range).Select();
      resetTypingFont();
      return;
    }
    var end = formulaRange.End;
    doc.Range(end, end).Select();
    if (payload.nativeEquation) {
      /* WPS keeps a new equation open for editing; like Word's
         VTPlaceCaretAfterInlineNativeEquation, MoveRight leaves it without
         moving past the next character. */
      safe(function () { app().ActiveWindow.Selection.MoveRight(1, 1, 0); });
      if (app().ActiveWindow.Selection.Start !== end) doc.Range(end, end).Select();
    }
    resetTypingFont();
  }

  /* A new numbered formula is numbered first (office::wps::numbering reads
     the document as it is now), then written. */
  function applyResult(sessionId, payload) {
    var current = active;
    if (!(payload.numbered && (current.mode === "create" || current.conversion))) {
      writeResult(sessionId, payload, null);
      return;
    }
    var facts;
    try {
      var doc = findDocument(current.docKey);
      if (!doc) throw new Error("The WPS document of this formula is no longer open");
      var anchorRange = bookmarkRange(doc, current.anchor);
      if (!anchorRange) throw new Error("The formula's position in the document was lost; insert it again");
      facts = numberingFacts(doc, anchorRange.Start);
    } catch (error) {
      finishResult(sessionId, { ok: false, error: "WPS could not number the formula: " + (error && error.message || error) });
      return;
    }
    call("POST", "/sessions/" + sessionId + "/numbering",
      { kind: payload.nativeEquation ? "equation" : "image", facts: facts }, 20000).then(function (numbering) {
      if (!active || active.sessionId !== sessionId) return;
      writeResult(sessionId, payload, numbering);
    }, function (error) {
      if (!active || active.sessionId !== sessionId) return;
      finishResult(sessionId, { ok: false, error: "VisualTeX could not number the formula: " + error.message });
    });
  }

  function writeResult(sessionId, payload, numbering) {
    var current = active;
    var result = { ok: true };
    try {
      var doc = findDocument(current.docKey);
      if (!doc) throw new Error("The WPS document of this formula is no longer open");
      var anchorRange = bookmarkRange(doc, current.anchor);
      if (!anchorRange) throw new Error("The formula's position in the document was lost; insert it again");
      withUndo("VisualTeX 公式", function () {
        var formulaRange = current.conversion
          ? applyConversion(doc, sessionId, payload, anchorRange, numbering, current)
          : payload.nativeEquation
            ? applyEquation(doc, sessionId, payload, anchorRange, numbering)
            : applyImage(doc, sessionId, payload, anchorRange, numbering);
        deleteBookmark(doc, current.anchor);
        safe(function () { placeCaret(doc, payload, formulaRange, current.mode); });
      });
    } catch (error) {
      result = { ok: false, error: "WPS could not write the formula: " + (error && error.message || error) };
    }
    finishResult(sessionId, result);
  }

  function finishResult(sessionId, result) {
    if (!result.ok) log("session=" + sessionId + " apply-failed " + result.error);
    active = null;
    call("POST", "/sessions/" + sessionId + "/applied", result, 10000).catch(function () {});
  }

  function create(nativeEquation, displayMode, numbered) {
    var target = selectionTarget();
    var doc = target.doc;
    if (!doc) {
      notify("请先打开一个 WPS 文字文档。");
      return;
    }
    var range = app().ActiveWindow.Selection.Range;
    range = doc.Range(range.Start, range.Start);
    var fontSizePt = fontSizeAt(range);
    if (displayMode === "block") range = prepareDisplayRange(doc, range);
    var sessionId = newUuid();
    var anchor = "VT_W_" + sessionId.replace(/-/g, "");
    doc.Bookmarks.Add(anchor, range);
    startSession({
      sessionId: sessionId,
      doc: doc,
      anchor: anchor,
      mode: "create",
      nativeEquation: nativeEquation,
      displayMode: displayMode,
      numbered: numbered && displayMode === "block" ? true : undefined,
      fontSizePt: fontSizePt
    });
  }

  function editImage(target) {
    var reference = target.reference;
    if (!reference) {
      notify("这个 VisualTeX 公式缺少公式标识，无法编辑。");
      return;
    }
    var sessionId = newUuid();
    var anchor = "VT_W_" + sessionId.replace(/-/g, "");
    target.doc.Bookmarks.Add(anchor, target.shape.Range);
    startSession({
      sessionId: sessionId,
      doc: target.doc,
      anchor: anchor,
      mode: "edit",
      nativeEquation: false,
      displayMode: reference.displayMode,
      /* A numbered picture keeps its number: only the picture is replaced. */
      numbered: reference.displayMode === "block" && reference.numbered,
      formulaId: reference.formulaId,
      encodedMetadata: target.metadata,
      sourceObjectId: target.shape.Title,
      fontSizePt: Number(safe(function () { return target.shape.Range.Font.Size; }, 0)) || undefined
    });
  }

  /* A Word numbered equation: an equation array whose row ends in a "#" run
     and the (number) delimiter, the shape omml_mathml strips before editing
     (try_resolve_word_native_equation_number_host). The app puts the edited
     formula back in front of the same number (omml_write::keep_equation_number). */
  var NUMBERED_ROW_END = new RegExp(
    "#</m:t>\\s*</m:r>\\s*<m:d\\b(?:(?!<m:d\\b)[\\s\\S])*?</m:d>\\s*" +
    "(?:<m:ctrlPr>[\\s\\S]*?</m:ctrlPr>\\s*)?</m:e>\\s*" +
    "(?:<m:ctrlPr>[\\s\\S]*?</m:ctrlPr>\\s*)?</m:eqArr>");

  function isNumberedEquationXml(xml) {
    return NUMBERED_ROW_END.test(xml);
  }

  function editEquation(target) {
    var range = target.omath.Range;
    var xml = range.WordOpenXML;
    var displayMode = equationDisplayMode(target.doc, target.omath);
    var sessionId = newUuid();
    /* The anchor the Word path uses for an untagged equation edit (VT_E_). */
    var anchor = "VT_E_" + sessionId.slice(0, 24).replace(/-/g, "");
    target.doc.Bookmarks.Add(anchor, range);
    startSession({
      sessionId: sessionId,
      doc: target.doc,
      anchor: anchor,
      mode: "edit",
      nativeEquation: true,
      displayMode: displayMode,
      numbered: displayMode === "block" && isNumberedEquationXml(String(xml || "")),
      formulaId: newUuid(),
      sourceObjectId: anchor,
      nativeEditXml: xml,
      fontSizePt: fontSizeAt(range)
    });
  }

  function editSelection(explicit) {
    var target = selectionTarget();
    if (target.kind === "image") editImage(target);
    else if (target.kind === "omml") editEquation(target);
    else if (explicit) notify("请先选中一个 VisualTeX 公式或 WPS/Word 公式。");
  }

  function run(action) {
    try {
      action();
    } catch (error) {
      log("command failed " + (error && error.stack || error));
      notify("VisualTeX 操作失败：" + (error && error.message || error));
    }
  }

  function openVisualTeX() {
    call("POST", "/open", {}, 10000).catch(function (error) {
      notify("VisualTeX 无法打开主窗口：" + error.message);
    });
  }

  /* ---- AI agents (docs/office-mcp/MACOS.md) -----------------------------------
     VisualTeX hands the add-in an agent's operations over the WPS connection
     (office::wps::bridge agent_commands); each answers what the Word add-in's
     VisualTeX_AgentRun answers. A command that opens the formula editor hands
     its request back instead: VisualTeX draws the agent's LaTeX in its hidden
     editor, and the add-in writes the result as after Apply. While an
     operation runs its messages are collected for the answer (a WPS alert
     stops every add-in until it is answered) and confirmations say yes. */

  var AGENT_POLL_SECONDS = 20;
  var agentRun = null;
  var agentLoopStarted = false;

  function agentFail(code, message) {
    var error = new Error(message);
    error.agentCode = code;
    throw error;
  }

  function agentDocument(args) {
    if (!safe(function () { return app().Documents.Count; }, 0)) {
      agentFail("document_not_found", "No document is open in WPS Writer.");
    }
    if (!args.document) return app().ActiveDocument;
    var doc = findDocument(args.document);
    if (!doc) agentFail("document_not_found", "The document is no longer open; call list_documents again.");
    return doc;
  }

  function agentWritable(args) {
    var doc = agentDocument(args);
    if (safe(function () { return doc.ReadOnly; }, false)) agentFail("document_read_only", "The document is read-only.");
    if (documentKey(app().ActiveDocument) !== documentKey(doc)) doc.Activate();
    return doc;
  }

  function agentRange(item) {
    return item.shape ? item.shape.Range : item.omath.Range;
  }

  /* One formula as the app reads it (VTAgentFormulaJson): kind, place, display, numbering, size and the
     payload (a picture's metadata, the document part of an equation's WordOpenXML). */
  function agentFormula(doc, item, contextChars) {
    var range = agentRange(item);
    var formula;
    if (item.shape) {
      var reference = parseFormulaReference(item.shape.Title);
      formula = {
        formula_id: "vtx:" + reference.formulaId, kind: "picture", display: reference.displayMode,
        numbered: reference.displayMode === "block" && reference.numbered,
        font_size_pt: formulaPictureSize(doc, item.shape) || fontSizeAt(range) || null,
        vtx_formula_id: reference.formulaId,
        payload_b64: base64UrlUtf8(String(safe(function () { return item.shape.AlternativeText; }, "") || ""))
      };
    } else {
      var text = String(safe(function () { return range.Text; }, "") || "");
      var numbered = text.indexOf("#") >= 0 && safe(function () { return range.Fields.Count; }, 0) > 0;
      formula = {
        formula_id: "w:m:" + range.Start + ":" + range.End, kind: "office_equation",
        display: numbered ? "block" : safe(function () { return equationDisplayMode(doc, item.omath); }, "inline"),
        numbered: numbered, font_size_pt: fontSizeAt(range) || null, vtx_formula_id: null,
        payload_b64: base64UrlUtf8(documentPart(safe(function () { return range.WordOpenXML; }, "")))
      };
    }
    formula.start = range.Start;
    formula.end = range.End;
    var before = "";
    var after = "";
    if (contextChars > 0) {
      before = String(safe(function () { return doc.Range(Math.max(0, range.Start - contextChars), range.Start).Text; }, "") || "");
      after = String(safe(function () { return doc.Range(range.End, Math.min(doc.Content.End, range.End + contextChars)).Text; }, "") || "");
    }
    formula.context_before = before.replace(/\r/g, "\n");
    formula.context_after = after.replace(/\r/g, "\n");
    return formula;
  }

  /* The VisualTeX pictures and the equations of the document, in reading order. */
  function agentFormulas(doc) {
    var items = [];
    var shapes = doc.InlineShapes;
    for (var index = 1; index <= shapes.Count; index++) {
      var shape = shapes.Item(index);
      if (isFormulaPicture(shape)) items.push({ start: shape.Range.Start, shape: shape });
    }
    var maths = doc.OMaths;
    for (var k = 1; k <= maths.Count; k++) {
      var omath = maths.Item(k);
      var range = omath.Range;
      if (safe(function () { return range.InlineShapes.Count; }, 0) > 0 || !textOf(range)) continue;
      items.push({ start: range.Start, omath: omath });
    }
    items.sort(function (a, b) { return a.start - b.start; });
    return items;
  }

  /* The formula an agent id names: "vtx:<id>" (a picture) or "w:m:<start>:<end>" (an equation). */
  function agentFind(doc, id) {
    id = String(id || "");
    if (id.indexOf("vtx:") === 0) {
      var shapes = doc.InlineShapes;
      for (var index = 1; index <= shapes.Count; index++) {
        var shape = shapes.Item(index);
        var reference = isFormulaPicture(shape) ? parseFormulaReference(shape.Title) : null;
        if (reference && reference.formulaId === id.slice(4)) return { shape: shape };
      }
      agentFail("formula_not_found", "No VisualTeX formula in this document has that id.");
    }
    var parts = id.split(":");
    if (parts.length < 4 || parts[0] !== "w" || parts[1] !== "m") {
      agentFail("formula_not_found", "This is not a WPS formula id from list_formulas.");
    }
    var start = Number(parts[2]);
    var end = Number(parts[3]);
    var omath = start >= 0 && end > start ? safe(function () { return equationAt(doc, start); }, null) : null;
    if (!omath || omath.Range.Start !== start || omath.Range.End !== end) {
      agentFail("stale_formula_id", "The document changed; call list_formulas again.");
    }
    return { omath: omath };
  }

  function agentIsDisplay(doc, item) {
    return agentFormula(doc, item, 0).display === "block";
  }

  /* location: "cursor" (default), "document_end", {after_formula}, {before_formula} or {paragraph_end}
     (VTAgentPlaceInsertion). */
  function agentPlaceInsertion(doc, args) {
    var location = args.location;
    var target = null;
    if (location && typeof location === "object") {
      if (location.after_formula || location.before_formula) {
        var after = !!location.after_formula;
        var item = agentFind(doc, after ? location.after_formula : location.before_formula);
        var range = agentRange(item);
        if (agentIsDisplay(doc, item)) {
          var paragraph = range.Paragraphs.Item(1).Range;
          target = after ? paragraph.End - 1 : paragraph.Start;
        } else {
          target = after ? range.End : range.Start;
        }
      } else if (location.paragraph_end !== undefined && location.paragraph_end !== null) {
        var number = Number(location.paragraph_end);
        var count = doc.Paragraphs.Count;
        if (!(number >= 1 && number <= count)) {
          agentFail("invalid_arguments", "The document has " + count + " paragraphs.");
        }
        target = doc.Paragraphs.Item(number).Range.End - 1;
      }
    } else if (location === "document_end") {
      target = doc.Content.End - 1;
    } else if (location && location !== "cursor") {
      agentFail("invalid_arguments", "Unknown location " + location + ".");
    } else {
      /* a new formula goes after the user's selection, never over it */
      var selection = app().ActiveWindow.Selection;
      if (selection.Start !== selection.End) target = selection.End;
    }
    if (target !== null) doc.Range(target, target).Select();
  }

  function agentDone(message) {
    var parts = [];
    if (message) parts.push(message);
    if (agentRun) parts = parts.concat(agentRun.notes);
    return { message: parts.length ? parts.join(" ") : null };
  }

  /* The session the command started, once VisualTeX has it: its id and (for the hidden editor or an
     import) its request. A command that did not start one says why in its notes. */
  function agentStarted(withRequest) {
    var run = agentRun;
    var waited = 0;
    return new Promise(function (resolve, reject) {
      var check = function () {
        if (run.started) {
          run.started.then(function (reply) {
            resolve({ session_id: reply.sessionId || null, request: withRequest ? reply.request || null : null,
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

  var NUMBER_FORMATS = {
    "continuous": { mode: "sequence", separator: "." },
    "heading1-dot": { mode: "chapter", separator: "." },
    "heading1-dash": { mode: "chapter", separator: "-" },
    "heading2-dot": { mode: "section", separator: "." },
    "heading2-dash": { mode: "section", separator: "-" }
  };

  var AGENT_OPERATIONS = {
    list_documents: function () {
      var documents = app().Documents;
      var activeKey = safe(function () { return documentKey(app().ActiveDocument); }, "");
      var list = [];
      for (var index = 1; index <= documents.Count; index++) {
        var doc = documents.Item(index);
        var key = documentKey(doc);
        list.push({
          key: key, name: String(safe(function () { return doc.Name; }, "") || ""),
          path: safe(function () { return String(doc.FullName || "") || null; }, null),
          active: key === activeKey, saved: !!safe(function () { return doc.Saved; }, true),
          read_only: !!safe(function () { return doc.ReadOnly; }, false)
        });
      }
      return { documents: list };
    },

    list_formulas: function (args) {
      var doc = agentDocument(args);
      var offset = Number(args.offset) || 0;
      var limit = args.limit === undefined ? 50 : Number(args.limit);
      var contextChars = args.contextChars === undefined ? 40 : Number(args.contextChars);
      var kinds = args.kinds || [];
      var listed = [];
      var revision = "";
      var total = 0;
      agentFormulas(doc).forEach(function (item) {
        var kind = item.shape ? "picture" : "office_equation";
        if (kinds.length && kinds.indexOf(kind) < 0) return;
        revision += item.start + ";";
        if (total >= offset && listed.length < limit) listed.push(agentFormula(doc, item, contextChars));
        total++;
      });
      return { revision: revision, total: total, offset: offset, formulas: listed };
    },

    read_formula: function (args) {
      var doc = agentDocument(args);
      return agentFormula(doc, agentFind(doc, args.formulaId), 120);
    },

    /* A VisualTeX picture by its own id (WPS equations carry none). */
    find_vtx_formula: function (args) {
      var doc = agentDocument(args);
      var wanted = String(args.vtxId || "");
      var found = agentFormulas(doc).filter(function (item) {
        return item.shape && parseFormulaReference(item.shape.Title).formulaId === wanted;
      })[0];
      return { formula: found ? agentFormula(doc, found, 120) : null };
    },

    get_selection: function (args) {
      var doc = args.document ? agentDocument(args) : app().ActiveDocument;
      if (!doc) agentFail("document_not_found", "No document is open in WPS Writer.");
      var selection = documentKey(doc) === documentKey(app().ActiveDocument)
        ? app().ActiveWindow.Selection : doc.ActiveWindow.Selection;
      var kind = "none";
      var formula = null;
      var preview = null;
      var shapes = safe(function () { return selection.InlineShapes; }, null);
      if (shapes && shapes.Count === 1 && isFormulaPicture(shapes.Item(1))) {
        kind = "formula";
        formula = agentFormula(doc, { shape: shapes.Item(1) }, 60);
      }
      if (kind === "none" && safe(function () { return selection.OMaths.Count; }, 0) >= 1) {
        var omath = safe(function () { return equationAt(doc, selection.Start); }, null);
        if (omath && safe(function () { return omath.Range.InlineShapes.Count; }, 0) === 0) {
          kind = "formula";
          formula = agentFormula(doc, { omath: omath }, 60);
        }
      }
      if (kind === "none" && selection.End > selection.Start) {
        kind = "text";
        preview = String(safe(function () { return selection.Range.Text; }, "") || "").replace(/\r/g, "\n");
        if (preview.length > 200) preview = preview.slice(0, 200) + "…";
      }
      return {
        document: documentKey(doc), kind: kind, formula: formula, text_preview: preview,
        location: { start: selection.Start, end: selection.End },
        read_only: !!safe(function () { return doc.ReadOnly; }, false)
      };
    },

    select_formula: function (args) {
      var doc = agentDocument(args);
      if (documentKey(app().ActiveDocument) !== documentKey(doc)) doc.Activate();
      var item = agentFind(doc, args.formulaId);
      agentRange(item).Select();
      safe(function () { app().ActiveWindow.ScrollIntoView(agentRange(item), true); });
      return { formula: agentFormula(doc, item, 0) };
    },

    /* A display formula goes with its paragraph (and its number); the numbers after it follow. */
    delete_formula: function (args) {
      var doc = agentWritable(args);
      var item = agentFind(doc, args.formulaId);
      var described = agentFormula(doc, item, 0);
      var range = agentRange(item);
      var target = described.display === "block" ? range.Paragraphs.Item(1).Range : range;
      withUndo("VisualTeX 删除公式", function () {
        if (item.shape) {
          var compactId = String(described.vtx_formula_id).replace(/-/g, "");
          ["VT_I_", "VT_N_", "VT_R_"].forEach(function (prefix) { deleteBookmark(doc, prefix + compactId); });
        }
        target.Delete();
        if (described.numbered) updateAllNumberFields(doc);
      });
      return agentDone("Deleted the formula.");
    },

    set_formula_font_size: function (args) {
      var size = Number(args.font_size_pt);
      if (!(size >= 5 && size <= 200)) agentFail("invalid_arguments", "font_size_pt must be from 5 to 200.");
      var doc = agentWritable(args);
      if (args.kind) {
        /* every inline or every display formula, as the ribbon's conversion sizes do */
        var count = resizeDocumentFormulas(args.kind === "display" ? "block" : "inline", size);
        return agentDone("Resized " + count + " " + args.kind + " formula(s) to " + size + " pt.");
      }
      var item = agentFind(doc, args.formulaId);
      withUndo("VisualTeX 公式字号", function () {
        if (item.shape) {
          resizeFormulaPicture(doc, item.shape, size);
        } else {
          var font = item.omath.Range.Font;
          font.Size = size;
          safe(function () { font.SizeBi = size; });
        }
      });
      return agentDone("The formula is " + size + " pt.");
    },

    number_display_formulas: function (args) {
      var doc = agentWritable(args);
      var selection = app().ActiveWindow.Selection;
      var scope = args.scope === "selection" ? { start: selection.Start, end: selection.End } : null;
      var found = displayNumberingCandidates(doc, scope);
      var total = found.equations.length + found.images.length;
      var skipped = [];
      if (found.table) skipped.push(found.table + " in tables");
      if (found.other) skipped.push(found.other + " sharing their paragraph");
      var left = skipped.length ? " Not numbered: " + skipped.join(", ") + "." : "";
      if (!total) return agentDone("No unnumbered display formula to number." + left);
      return renumberDocument({ doc: doc, equations: found.equations, images: found.images, undoLabel: "VisualTeX 批量编号" })
        .then(function (summary) {
          var failed = summary.failures.length ? " " + summary.failures.length + " failed: " + summary.failures[0] : "";
          return agentDone("Numbered " + total + " display formula(s)." + left + failed);
        });
    },

    update_equation_numbers: function (args) {
      var doc = agentWritable(args);
      return renumberDocument({ doc: doc }).then(function (summary) {
        var failed = summary.failures.length ? " " + summary.failures.length + " failed: " + summary.failures[0] : "";
        return agentDone("Updated " + summary.total + " equation numbers." + failed);
      });
    },

    set_equation_number_format: function (args) {
      var style = NUMBER_FORMATS[args.format_id];
      if (!style) agentFail("invalid_arguments", "Unknown format_id " + args.format_id + ".");
      var doc = agentWritable(args);
      call("POST", "/numbering/preference", style, 10000).catch(function (error) {
        log("numbering preference not saved: " + error.message);
      });
      chosenStyles[documentKey(doc)] = style;
      return renumberDocument({ doc: doc, chosen: style }).then(function () { return agentDone(""); });
    },

    /* The ribbon's equation-reference menu with the equation named by its number. */
    insert_equation_reference: function (args) {
      var wanted = "(" + String(args.number || "").replace(/[()]/g, "").trim() + ")";
      var doc = agentWritable(args);
      agentPlaceInsertion(doc, args);
      var targets = { omml: numberTargets(doc, "omml"), image: numberTargets(doc, "image") };
      var sources = { omml: "office_equation", image: "picture" };
      var available = [];
      var matches = [];
      ["omml", "image"].forEach(function (kind) {
        targets[kind].forEach(function (target, index) {
          var text = "(" + String(doc.Range(target.start, target.end).Text).replace(/[\r\n\t\x07\x13\x14\x15]/g, " ").trim() + ")";
          available.push(text + " " + sources[kind]);
          if (text === wanted && (!args.source || args.source === sources[kind])) matches.push({ kind: kind, index: index });
        });
      });
      if (!matches.length) {
        agentFail("formula_not_found", "No numbered equation " + wanted + " in this document. Numbered equations: " +
          (available.join(", ") || "none") + ".");
      }
      if (matches.length > 1) {
        agentFail("invalid_arguments", "Several numbered equations are " + wanted +
          "; give source to choose one. Numbered equations: " + available.join(", ") + ".");
      }
      var selection = app().ActiveWindow.Selection;
      referenceMenu = { docKey: documentKey(doc), start: selection.Start, end: selection.End, targets: targets };
      insertReference(matches[0].kind, matches[0].index);
      return agentDone("Inserted a reference to equation " + wanted + ".");
    },

    begin_insert: function (args) {
      var doc = agentWritable(args);
      agentPlaceInsertion(doc, args);
      var display = args.display === "block" ? "block" : "inline";
      create(args.format === "office_equation", display, !!args.numbered);
      return agentStarted(true);
    },

    /* New LaTeX and/or another format: the editor's session for the formula, or the conversion of it. */
    begin_update: function (args) {
      var doc = agentWritable(args);
      var item = agentFind(doc, args.formulaId);
      agentRange(item).Select();
      if (args.format === "office_equation" && item.shape) convertSelected(false);
      else if (args.format === "picture" && item.omath) convertSelected(true);
      else editSelection(true);
      return agentStarted(true);
    },

    /* Redraw and conversion of many formulas open VisualTeX's progress window, as from the ribbon; the app
       waits for the session they start. A restore to LaTeX is written here at once. */
    latex_text_to_formulas: function (args) {
      agentWritable(args);
      redrawLatex(args.scope !== "selection", args.format === "office_equation" ? "omml" : "image");
      return agentStarted(false);
    },

    formulas_to_latex_text: function (args) {
      agentWritable(args);
      var whole = args.scope !== "selection";
      var restore = args.from === "office_equation" ? restoreEquationsToLatex(whole)
        : args.from === "picture" ? restorePicturesToLatex(whole)
          : agentFail("invalid_arguments", "from must be picture or office_equation.");
      return Promise.resolve(restore).then(function () {
        return { session_id: null, message: agentRun.notes.length ? agentRun.notes.join(" ") : null };
      });
    },

    convert_formulas: function (args) {
      agentWritable(args);
      var pair = args.from + ">" + args.to;
      var source = pair === "picture>office_equation" ? "image" : pair === "office_equation>picture" ? "omml" : null;
      if (!source) agentFail("invalid_arguments", "In WPS Writer, from and to are picture and office_equation, one of each.");
      convertDocument(source, args.scope === "selection");
      return agentStarted(false);
    },

    /* Markdown or LaTeX prose: the import at the location; the app gives the text to its import window. */
    begin_import: function (args) {
      var doc = agentWritable(args);
      agentPlaceInsertion(doc, args);
      importDocument();
      return agentStarted(true);
    }
  };

  function runAgentCommand(command) {
    agentRun = { notes: [], started: null };
    var run = agentRun;
    var operation = AGENT_OPERATIONS[command.op];
    return Promise.resolve().then(function () {
      if (!operation) agentFail("unknown_operation", "WPS Writer has no agent operation " + command.op + ".");
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

  /* Asks VisualTeX for the next operation, as long as this module is the add-in's (main.js loads a newer one
     when VisualTeX is updated). Off: asks again now and then. */
  function startAgentLoop(module) {
    if (agentLoopStarted) return;
    agentLoopStarted = true;
    var version = encodeURIComponent(String(safe(function () { return app().Version; }, "") || ""));
    var next = function (delay) { setTimeout(ask, delay); };
    var ask = function () {
      if (window.VisualTeXWpsModule !== module) return;
      /* a VisualTeX update brings a newer module without the user touching WPS (main.js) */
      if (typeof vtxRefreshSoon === "function") safe(vtxRefreshSoon);
      call("GET", "/agent/commands?host=writer&wait=" + AGENT_POLL_SECONDS + "&version=" + version, undefined,
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

  /* Ribbon commands by control id (addin/ribbon.xml). */
  var COMMANDS = {
    vtxImageInline: function () { create(false, "inline"); },
    vtxImageDisplay: function () { create(false, "block"); },
    vtxOmmlInline: function () { create(true, "inline"); },
    vtxOmmlDisplay: function () { create(true, "block"); },
    vtxEditSelected: function () { editSelection(true); },
    vtxOpen: openVisualTeX,
    vtxUpdateNumbers: updateNumbers,
    vtxNumberingSequence: function () { applyNumberingFormat("sequence", "."); },
    vtxNumberingChapterDot: function () { applyNumberingFormat("chapter", "."); },
    vtxNumberingChapterDash: function () { applyNumberingFormat("chapter", "-"); },
    vtxNumberingSectionDot: function () { applyNumberingFormat("section", "."); },
    vtxNumberingSectionDash: function () { applyNumberingFormat("section", "-"); },
    vtxNumberSelection: function () { numberDisplayFormulas(false); },
    vtxNumberDocument: function () { numberDisplayFormulas(true); },
    vtxRestoreSelectionOmml: function () { restoreEquationsToLatex(false); },
    vtxRestoreSelectionImage: function () { restorePicturesToLatex(false); },
    vtxRestoreDocumentOmml: function () { restoreEquationsToLatex(true); },
    vtxRestoreDocumentImage: function () { restorePicturesToLatex(true); },
    vtxConvertNative: function () { convertSelected(false); },
    vtxConvertImage: function () { convertSelected(true); },
    vtxConvertDocumentImagesToOmml: function () { convertDocument("image"); },
    vtxConvertDocumentOmmlToImages: function () { convertDocument("omml"); },
    vtxRedrawSelectionOmml: function () { redrawLatex(false, "omml"); },
    vtxRedrawSelectionImage: function () { redrawLatex(false, "image"); },
    vtxRedrawDocumentOmml: function () { redrawLatex(true, "omml"); },
    vtxRedrawDocumentImage: function () { redrawLatex(true, "image"); },
    vtxDocumentImport: importDocument,
    vtxUpdateEquationFormat: updateEquationFormat
  };

  window.VisualTeXWpsModule = {
    init: function (bootstrap) {
      host = bootstrap;
      if (host && host.debug) {
        /* The test harness reaches the helpers it checks one by one. */
        window.VisualTeXWpsModule.internals = {
          equationAt: equationAt, insertEquationFile: insertEquationFile, finishEquation: finishEquation,
          replaceNumberedEquation: replaceNumberedEquation, restyleFacts: restyleFacts, download: download,
          equationNumberTarget: equationNumberTarget, call: call, newUuid: newUuid,
          resolveTextSpans: resolveTextSpans, redrawFontSize: redrawFontSize, restoreTargets: restoreTargets,
          restoreManifest: restoreManifest, active: function () { return active; }
        };
      }
      log("module loaded");
      startAgentLoop(window.VisualTeXWpsModule);
    },
    onAction: function (id) {
      run(function () {
        var reference = /^vtxRef_(omml|image)_(\d+)$/.exec(String(id));
        if (reference) insertReference(reference[1], Number(reference[2]));
        else if (COMMANDS[id]) COMMANDS[id]();
        else notify("这个 VisualTeX 功能还没有接入 WPS：" + id);
      });
    },
    menuContent: function (id) {
      if (id === "vtxCrossReference") {
        return safe(crossReferenceMenu, menuMessage("无法读取公式列表"));
      }
      return null;
    },
    itemCount: function () {
      return fontSizeItemCount();
    },
    itemLabel: function (id, index) {
      return safe(function () { return fontSizeItemLabel(id, Number(index)); }, index > 0 ? presetLabel(index - 1) : "字号");
    },
    selectedIndex: function (id) {
      return safe(function () { return fontSizeSelectedIndex(id); }, 0);
    },
    onSelect: function (id, index) {
      run(function () {
        if (id === "vtxImageFontSize") {
          if (index > 0) resizeSelectedPictures(FONT_PRESETS[index - 1][0]);
        } else {
          applyConversionFontSize(id, index);
        }
      });
    },
    /* Returns true when the double-click opens VisualTeX, so WPS's own action
       (picture pane, equation editing) is cancelled. WPS can raise the event
       twice for one double-click; the second one is only cancelled. */
    onBeforeDoubleClick: function () {
      var now = Date.now();
      if (now - lastDoubleClick < 800) return true;
      var target = safe(selectionTarget, { kind: "none" });
      if (target.kind === "none") return false;
      lastDoubleClick = now;
      setTimeout(function () {
        run(function () {
          if (target.kind === "image") editImage(target);
          else editEquation(target);
        });
      }, 0);
      return true;
    },
    onSelectionChange: function () {
      safe(resetCaretNextToFormula);
      /* The picture size list shows the selection's size: refreshed once the
         caret settles, not on every move. */
      if (sizeRefresh) clearTimeout(sizeRefresh);
      sizeRefresh = setTimeout(function () {
        sizeRefresh = null;
        invalidateControl("vtxImageFontSize");
      }, 300);
    },
    /* No formula is being edited (main.js may swap in a newer module). */
    isIdle: function () {
      return !active && !agentRun;
    }
  };
})();
