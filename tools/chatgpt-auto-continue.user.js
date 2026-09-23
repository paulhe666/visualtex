// ==UserScript==
// @name         ChatGPT Auto Continue Guard
// @namespace    local.visualtex.tools
// @version      2.1.0
// @description  Continue an open ChatGPT conversation after the current answer has stopped and stayed idle.
// @match        https://chatgpt.com/*
// @match        https://chat.openai.com/*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  const STORAGE_KEY = 'chatgpt-auto-continue-guard-settings-v2';
  const PANEL_ID = 'chatgpt-auto-continue-guard';
  const STOP_MISSING_GRACE_MS = 2500;
  const COMPOSER_INPUT_GRACE_MS = 3000;
  const SEND_BUTTON_WAIT_MS = globalThis.__CHATGPT_AUTO_CONTINUE_TEST_MODE__ ? 60 : 3000;
  const DEFAULTS = Object.freeze({
    idleSeconds: 60,
    maxContinues: 12,
    sentinel: '[TASK_DONE]',
  });
  const RETRY_POLICY = Object.freeze({
    stopWaitMilliseconds: globalThis.__CHATGPT_AUTO_CONTINUE_TEST_MODE__ ? 80 : 10000,
    generationWaitMilliseconds: globalThis.__CHATGPT_AUTO_CONTINUE_TEST_MODE__ ? 100 : 12000,
    betweenAttemptsMilliseconds: globalThis.__CHATGPT_AUTO_CONTINUE_TEST_MODE__ ? 5 : 750,
  });

  const state = {
    enabled: false,
    busy: false,
    conversationKey: '',
    generationActive: false,
    hasSeenGeneration: false,
    idleSince: Date.now(),
    lastContentActivityAt: Date.now(),
    lastResponseSnapshot: null,
    seenResponseFingerprints: new Map(),
    progressObserver: null,
    observedRoot: null,
    progressRevision: 0,
    handledProgressRevision: 0,
    allowBranchSwitch: false,
    ignoredCompletionMarkerFingerprint: '',
    lastProgressSource: '尚无',
    lastUserSignature: '',
    lastComposerText: '',
    lastComposerInputAt: 0,
    stopMissingSince: 0,
    stopWasSeenInCurrentResponse: false,
    pendingSubmission: false,
    pendingSince: 0,
    pendingKind: '',
    pendingExpectedUserText: '',
    pendingUserSignatureBefore: '',
    pendingExpectedMessageCountBefore: 0,
    pendingAction: '',
    nextActionNotBefore: 0,
    continueCount: 0,
    status: '未启动',
  };

  let settings = loadSettings();
  let ui = null;

  function clampInteger(value, min, max, fallback) {
    const number = Number.parseInt(String(value), 10);
    return Number.isFinite(number) ? Math.min(max, Math.max(min, number)) : fallback;
  }

  function normalizeSentinel(value) {
    const text = String(value || '').trim();
    return text || DEFAULTS.sentinel;
  }

  function loadSettings() {
    try {
      const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}');
      return {
        idleSeconds: clampInteger(saved.idleSeconds, 10, 3600, DEFAULTS.idleSeconds),
        maxContinues: clampInteger(saved.maxContinues, 1, 100, DEFAULTS.maxContinues),
        sentinel: normalizeSentinel(saved.sentinel),
      };
    } catch (_error) {
      return { ...DEFAULTS };
    }
  }

  function saveSettings() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
  }

  function getConversationKey() {
    return `${location.pathname}${location.search}`;
  }

  function normalizeText(value) {
    return String(value || '').replace(/\s+/g, ' ').trim();
  }

  function isVisible(element) {
    if (!element || !element.isConnected) return false;
    let current = element;
    while (current?.nodeType === 1) {
      const ancestorStyle = getComputedStyle(current);
      if (ancestorStyle.display === 'none' || ancestorStyle.visibility === 'hidden'
        || Number(ancestorStyle.opacity || 1) === 0) return false;
      current = current.parentElement;
    }
    const style = getComputedStyle(element);
    const rect = element.getBoundingClientRect();
    return style.display !== 'none'
      && style.visibility !== 'hidden'
      && Number(style.opacity || 1) !== 0
      && rect.width > 0
      && rect.height > 0;
  }

  function isEnabledButton(element) {
    return isVisible(element)
      && !element.disabled
      && element.getAttribute('aria-disabled') !== 'true'
      && element.getAttribute('data-disabled') !== 'true';
  }

  function getButtonLabel(button) {
    return normalizeText([
      button?.innerText,
      button?.getAttribute?.('aria-label'),
      button?.getAttribute?.('title'),
    ].filter(Boolean).join(' '));
  }

  function findButtonByText(pattern) {
    return Array.from(document.querySelectorAll('button')).find((button) => {
      const visibleText = normalizeText(button.innerText || '');
      const ariaLabel = normalizeText(button.getAttribute('aria-label') || '');
      const title = normalizeText(button.getAttribute('title') || '');
      return isEnabledButton(button)
        && (pattern.test(visibleText) || pattern.test(ariaLabel) || pattern.test(title));
    }) || null;
  }

  function findSubmitButton() {
    const selectors = [
      '#composer-submit-button',
      'button[data-testid="send-button"]',
      'button[data-testid="stop-button"]',
      'button[data-testid="fruitjuice-send-button"]',
    ];
    for (const selector of selectors) {
      const button = document.querySelector(selector);
      if (isVisible(button)) return button;
    }
    return findButtonByText(/^(send prompt|send message|send|发送提示词|发送提示|发送消息|发送)$/i);
  }

  function getComposerControlScope(composer = findComposer()) {
    if (!composer) return null;
    const form = composer.closest?.('form');
    if (form) return form;
    let current = composer.parentElement;
    for (let depth = 0; current && depth < 6; depth += 1) {
      if (current.querySelector?.('button')) return current;
      current = current.parentElement;
    }
    return composer.parentElement || null;
  }

  function findStopButton() {
    const composer = findComposer();
    const scope = getComposerControlScope(composer);
    if (!scope) return null;
    const stopLabel = /^(停止回答|停止生成|停止响应|stop response|stop generating|stop)$/i;
    const selectors = [
      'button[data-testid="stop-button"]',
      '#composer-stop-button',
      '#composer-submit-button',
      'button[type="submit"]',
      'button',
    ];
    const candidates = [];
    for (const selector of selectors) {
      for (const button of Array.from(scope.querySelectorAll?.(selector) || [])) {
        if (!candidates.includes(button)) candidates.push(button);
      }
    }
    return candidates.find((button) => {
      if (!isEnabledButton(button)) return false;
      const testId = button.getAttribute('data-testid') || '';
      return testId === 'stop-button'
        || [button.innerText, button.getAttribute('aria-label'), button.getAttribute('title')]
          .some((label) => stopLabel.test(normalizeText(label)));
    }) || null;
  }

  function isGenerating() {
    return Boolean(findStopButton());
  }

  function findComposer() {
    const selectors = [
      '#prompt-textarea',
      '[data-testid="composer-text-input"]',
      'textarea[placeholder]',
      'div[contenteditable="true"][data-lexical-editor="true"]',
      'div[contenteditable="true"][role="textbox"]',
    ];
    for (const selector of selectors) {
      const composer = Array.from(document.querySelectorAll(selector)).find(isVisible);
      if (composer) return composer;
    }
    return null;
  }

  function getComposerText(composer) {
    if (!composer) return '';
    if (composer instanceof HTMLTextAreaElement || composer instanceof HTMLInputElement) return composer.value;
    return composer.innerText || composer.textContent || '';
  }

  function setNativeValue(element, value) {
    const prototype = element instanceof HTMLTextAreaElement
      ? HTMLTextAreaElement.prototype
      : HTMLInputElement.prototype;
    const setter = Object.getOwnPropertyDescriptor(prototype, 'value')?.set;
    if (!setter) throw new Error('无法设置 ChatGPT 输入框');
    setter.call(element, value);
    element.dispatchEvent(new Event('input', { bubbles: true }));
    element.dispatchEvent(new Event('change', { bubbles: true }));
  }

  function findProseMirrorView(element) {
    const candidates = [element, ...Array.from(element?.querySelectorAll?.('*') || []).slice(0, 40)];
    for (const candidate of candidates) {
      let descriptor = candidate?.pmViewDesc;
      const seen = new Set();
      while (descriptor && !seen.has(descriptor)) {
        seen.add(descriptor);
        const view = descriptor.view;
        if (view?.state?.tr && view.state?.doc?.content
          && typeof view.dispatch === 'function') return view;
        descriptor = descriptor.parent;
      }
    }
    return null;
  }

  function setProseMirrorText(element, value) {
    const view = findProseMirrorView(element);
    if (!view) return false;
    try {
      const { state } = view;
      const paragraphType = state.schema?.nodes?.paragraph;
      if (!paragraphType || typeof paragraphType.create !== 'function') return false;
      const blocks = String(value).split(/\r?\n/).map((line) => paragraphType.create(
        null,
        line && typeof state.schema.text === 'function' ? state.schema.text(line) : null,
      ));
      const transaction = state.tr.replaceWith(0, state.doc.content.size, blocks);
      view.dispatch(transaction);
      view.focus?.();
      const editorText = view.state?.doc?.textContent ?? getComposerText(element);
      return normalizeText(editorText) === normalizeText(value);
    } catch (_error) {
      return false;
    }
  }

  function setContentEditableText(element, value) {
    element.focus();
    if (setProseMirrorText(element, value)) return;
    const selection = window.getSelection();
    const range = document.createRange();
    const paragraphs = Array.from(element.querySelectorAll?.('p') || []);
    // ChatGPT's ProseMirror rejects text inserted at the editor root. For its
    // normal single-paragraph composer, place the selection inside that block,
    // matching where real keyboard input is inserted.
    const selectionTarget = paragraphs.length === 1 ? paragraphs[0] : element;
    range.selectNodeContents(selectionTarget);
    selection.removeAllRanges();
    selection.addRange(range);

    let inserted = false;
    try {
      inserted = value
        ? document.execCommand('insertText', false, value)
        : document.execCommand('delete', false, null);
    } catch (_error) {
      inserted = false;
    }

    if (normalizeText(getComposerText(element)) !== normalizeText(value)) {
      // Lexical/ProseMirror can reject execCommand while still accepting a
      // real selection mutation followed by beforeinput/input events.
      try {
        const beforeInput = typeof InputEvent === 'function'
          ? new InputEvent('beforeinput', {
            bubbles: true,
            cancelable: true,
            inputType: value ? 'insertText' : 'deleteContentBackward',
            data: value || null,
          })
          : new Event('beforeinput', { bubbles: true, cancelable: true });
        element.dispatchEvent(beforeInput);
        if (normalizeText(getComposerText(element)) !== normalizeText(value)) {
          range.deleteContents();
          if (value) {
            const textNode = document.createTextNode(value);
            range.insertNode(textNode);
            range.setStartAfter?.(textNode);
            range.collapse?.(true);
          }
          selection.removeAllRanges();
          selection.addRange(range);
        }
      } catch (_error) {
        // The final text check below decides whether the editor accepted it.
      }
    }

    const inputEvent = typeof InputEvent === 'function'
      ? new InputEvent('input', { bubbles: true, inputType: 'insertText', data: value })
      : new Event('input', { bubbles: true });
    element.dispatchEvent(inputEvent);
    if (normalizeText(getComposerText(element)) !== normalizeText(value)) {
      throw new Error('网页编辑器未接受续写文字；尚未发送');
    }
  }

  function dispatchComposerEnter(composer) {
    if (!composer || typeof KeyboardEvent !== 'function') return false;
    const eventOptions = {
      key: 'Enter',
      code: 'Enter',
      keyCode: 13,
      which: 13,
      bubbles: true,
      cancelable: true,
    };
    composer.focus();
    composer.dispatchEvent(new KeyboardEvent('keydown', eventOptions));
    composer.dispatchEvent(new KeyboardEvent('keypress', eventOptions));
    composer.dispatchEvent(new KeyboardEvent('keyup', eventOptions));
    return true;
  }

  function setComposerText(composer, value) {
    composer.focus();
    if (composer instanceof HTMLTextAreaElement || composer instanceof HTMLInputElement) {
      setNativeValue(composer, value);
    } else {
      setContentEditableText(composer, value);
    }
  }

  function composerHasAttachments(composer) {
    const scope = composer?.closest('form') || composer?.parentElement?.parentElement;
    if (!scope) return false;
    if (scope.querySelector('[data-testid*="attachment"], [data-testid*="file-upload"], [data-testid*="file-thumbnail"]')) return true;
    return Array.from(scope.querySelectorAll('button')).some((button) => {
      const label = getButtonLabel(button);
      return /(删除|移除|remove).*(附件|文件|attachment|file)/i.test(label);
    });
  }

  function wait(milliseconds) {
    return new Promise((resolve) => window.setTimeout(resolve, milliseconds));
  }

  async function waitUntil(predicate, timeoutMilliseconds, intervalMilliseconds = 100) {
    const startedAt = Date.now();
    while (Date.now() - startedAt < timeoutMilliseconds) {
      const result = predicate();
      if (result) return result;
      await wait(intervalMilliseconds);
    }
    return null;
  }

  async function stopStalledGeneration() {
    const stopButton = findStopButton();
    if (!stopButton) return false;
    setStatus('检测到当前输入框的“停止回答”，点击一次并等待页面切换为空闲状态');
    stopButton.click();
    const stopped = await waitUntil(
      () => !findStopButton(),
      RETRY_POLICY.stopWaitMilliseconds,
      150,
    );
    if (!stopped) {
      throw new Error('点击当前输入框的“停止回答”后等待 10 秒，按钮仍未切换为空闲状态；未发送续写');
    }
    await wait(250);
    return true;
  }

  function findSendButton() {
    const scope = getComposerControlScope();
    if (isGenerating()) return null;
    const selectors = [
      'button[type="submit"]',
      '#composer-submit-button',
      'button[data-testid="send-button"]',
      'button[data-testid="fruitjuice-send-button"]',
    ];
    for (const selector of selectors) {
      const searchRoot = scope || document;
      const button = Array.from(searchRoot.querySelectorAll(selector)).find((candidate) =>
        isEnabledButton(candidate)
        && !isVoiceControlButton(candidate)
        && candidate.getAttribute('data-testid') !== 'stop-button'
        && !/(停止回答|停止生成|stop response|stop generating)/i.test(getButtonLabel(candidate)));
      if (button) return button;
    }
    return findButtonByText(/^(send prompt|send message|send|发送提示词|发送提示|发送消息|发送)$/i);
  }

  function isVoiceControlButton(button) {
    const label = getButtonLabel(button);
    const testId = button?.getAttribute?.('data-testid') || '';
    return /(voice|speech|dictat|microphone|mic|语音|听写|麦克风)/i.test(`${testId} ${label}`);
  }

  function getComposerControlState() {
    if (findStopButton()) return 'stop';
    if (findSendButton()) return 'send';
    const scope = getComposerControlScope();
    const voiceButton = Array.from(scope?.querySelectorAll?.('button') || [])
      .find((button) => isEnabledButton(button) && isVoiceControlButton(button));
    if (voiceButton) return 'voice';
    return normalizeText(getComposerText(findComposer())) ? 'text-no-send' : 'idle-unknown';
  }

  function makeContinueText() {
    return `继续。请从中断处接着完成；若任务已经全部完成，请只回复 ${settings.sentinel}。`;
  }

  function countSubmittedMessages(expectedText) {
    const expected = normalizeText(expectedText);
    if (!expected) return 0;
    const users = getUserTurns();
    if (users.length) {
      return users.filter((element) => normalizeText(
        element.innerText || element.textContent || '',
      ).includes(expected)).length;
    }
    const root = getThreadRoot();
    const composer = findComposer();
    if (!root) return 0;
    return Array.from(root.querySelectorAll('[data-testid^="conversation-turn-"]'))
      .slice(-40)
      .filter((turn) => (!composer || !turn.contains?.(composer))
        && normalizeText(turn.innerText || turn.textContent || '').includes(expected)).length;
  }

  async function sendContinueMessage() {
    const initialComposer = findComposer();
    if (initialComposer && composerHasAttachments(initialComposer)) {
      throw new Error('输入区有附件，为避免误发送附件，已暂停');
    }
    const preservedDraft = initialComposer ? getComposerText(initialComposer) : '';
    const continueText = makeContinueText();
    const userSignatureBefore = getUserSignature();
    const expectedMessageCountBefore = countSubmittedMessages(continueText);
    const responseBefore = getResponseSnapshot();
    let action = '';
    let submittedViaForm = false;
    let submittedViaEnter = false;

    // Idle continuation must always go through the current composer. Page-wide
    // Retry/Continue buttons can belong to an old turn and must never preempt it.
    const composer = await waitUntil(findComposer, 3000);
    if (!composer) throw new Error('找不到输入框，尚未发送续写');
    if (composerHasAttachments(composer)) throw new Error('输入区有附件，为避免误发送附件，已暂停');
    try {
      setStatus('正在把“继续”写入输入框');
      setComposerText(composer, continueText);
    } catch (error) {
      if (normalizeText(getComposerText(composer)) !== normalizeText(preservedDraft)) {
        try { setComposerText(composer, preservedDraft); } catch (_restoreError) { /* Keep the original error. */ }
      }
      throw error;
    }
    await wait(300);
    if (normalizeText(getComposerText(composer)) !== normalizeText(continueText)) {
      setComposerText(composer, preservedDraft);
      throw new Error('ChatGPT 编辑器回滚了脚本写入的“继续”；尚未点击发送');
    }
    setStatus('“继续”已稳定写入输入框，等待发送箭头');
    const sendButton = await waitUntil(findSendButton, SEND_BUTTON_WAIT_MS);
    if (!sendButton) {
      const form = composer.closest?.('form');
      if (form && typeof form.requestSubmit === 'function') {
        setStatus('发送箭头未出现，尝试提交输入框表单');
        action = '表单提交“继续”';
        submittedViaForm = true;
        form.requestSubmit();
      } else if (dispatchComposerEnter(composer)) {
        setStatus('发送箭头未出现，尝试回车发送');
        action = '回车提交“继续”';
        submittedViaEnter = true;
      } else {
        setComposerText(composer, preservedDraft);
        throw new Error('发送箭头不可用且找不到可提交的表单，原草稿已恢复；尚未发送');
      }
    } else {
      setStatus('已找到发送箭头，正在点击');
      action = '发送“继续”';
      sendButton.click();
    }

    // Clicking a button, clearing the composer, or showing Stop does not prove
    // that the continuation message entered the conversation.
    const confirmAction = () => {
      if (countSubmittedMessages(continueText) > expectedMessageCountBefore) {
        return '“继续”消息已出现在对话中';
      }
      const currentComposer = findComposer();
      const composerCleared = normalizeText(getComposerText(currentComposer))
        !== normalizeText(continueText);
      if (composerCleared && isGenerating()) {
        return '输入框已清空，当前输入框显示正在回答';
      }
      const currentResponse = getResponseSnapshot();
      const responseStarted = currentResponse && (!responseBefore
        || (isSameOrNewerSnapshot(responseBefore, currentResponse)
          && isMeaningfulSnapshotChange(responseBefore, currentResponse)));
      if (composerCleared && responseStarted) return '输入框已清空，助手已开始新回复';
      return null;
    };
    let confirmation = await waitUntil(confirmAction, RETRY_POLICY.generationWaitMilliseconds, 200);

    // If the button click left the exact prompt in the composer and no user
    // turn appeared, try the form's submit path once. Never use this fallback
    // after the composer clears or the conversation changes.
    const stillUnsent = !confirmation && !normalizeText(preservedDraft)
      && getUserSignature() === userSignatureBefore
      && normalizeText(getComposerText(findComposer())) === normalizeText(continueText);
    const form = stillUnsent && !submittedViaForm ? findComposer()?.closest('form') : null;
    if (form && typeof form.requestSubmit === 'function') {
      setStatus('发送按钮未提交，尝试表单提交一次');
      try {
        form.requestSubmit();
        confirmation = await waitUntil(confirmAction, RETRY_POLICY.generationWaitMilliseconds, 200);
        action = '表单提交“继续”';
      } catch (error) {
        action = `表单提交失败：${error instanceof Error ? error.message : String(error)}`;
      }
    }

    const enterComposer = findComposer();
    const enterStillUnsent = !confirmation
      && getUserSignature() === userSignatureBefore
      && normalizeText(getComposerText(enterComposer)) === normalizeText(continueText);
    if (enterStillUnsent && !submittedViaEnter && dispatchComposerEnter(enterComposer)) {
      setStatus('表单未提交，尝试回车发送一次');
      confirmation = await waitUntil(confirmAction, RETRY_POLICY.generationWaitMilliseconds, 200);
      action = '回车提交“继续”';
    }

    const currentComposer = findComposer();
    const composerStillHasContinue = currentComposer
      && normalizeText(getComposerText(currentComposer)) === normalizeText(continueText);
    let draftRestoreError = '';
    try {
      if (preservedDraft && currentComposer
        && normalizeText(getComposerText(currentComposer)) !== normalizeText(preservedDraft)) {
        setComposerText(currentComposer, preservedDraft);
      } else if (composerStillHasContinue) {
        setComposerText(currentComposer, preservedDraft);
      }
    } catch (error) {
      draftRestoreError = error instanceof Error ? error.message : String(error);
    }
    return {
      attempted: true,
      confirmed: Boolean(confirmation),
      kind: 'message',
      expectedUserText: continueText,
      userSignatureBefore,
      expectedMessageCountBefore,
      action,
      message: confirmation
        ? `${action}，${confirmation}${preservedDraft && !draftRestoreError ? '；原草稿已恢复' : ''}${draftRestoreError ? `；草稿恢复失败：${draftRestoreError}` : ''}`
        : `${action || '续写操作'}执行后，“继续”没有出现在对话中${composerStillHasContinue ? '；发送前文字仍在输入框' : '；输入框已清空'}；不会重复发送${draftRestoreError ? `；草稿恢复失败：${draftRestoreError}` : ''}`,
    };
  }

  function getThreadRoot() {
    const composer = findComposer();
    return composer?.closest?.('main')
      || document.querySelector('[data-testid="conversation-turns"]')
      || document.querySelector('main')
      || document.querySelector('[role="main"]');
  }

  function isRenderedConversationNode(element) {
    if (!element || element.isConnected === false) return false;
    if (element.closest?.('[hidden], [aria-hidden="true"]')) return false;
    let current = element;
    while (current?.nodeType === 1) {
      const style = typeof getComputedStyle === 'function' ? getComputedStyle(current) : null;
      if (style && (style.display === 'none' || style.visibility === 'hidden'
        || Number(style.opacity || 1) === 0)) return false;
      current = current.parentElement;
    }
    return true;
  }

  function isNodeAfter(candidate, anchor) {
    if (!candidate || !anchor) return true;
    if (typeof anchor.compareDocumentPosition === 'function') {
      const relation = anchor.compareDocumentPosition(candidate);
      if (relation & 4) return true; // DOCUMENT_POSITION_FOLLOWING
      if (relation & 2) return false; // DOCUMENT_POSITION_PRECEDING
    }
    const candidateOrder = getTurnOrder(candidate);
    const anchorOrder = getTurnOrder(anchor);
    if (candidateOrder !== null && anchorOrder !== null) return candidateOrder > anchorOrder;
    return true;
  }

  function getLatestTurn() {
    const root = getThreadRoot();
    if (!root) return null;
    const composer = findComposer();
    // Branches can leave higher-numbered turns in the DOM. The last assistant
    // message in document order is the current displayed reply. Observe its
    // entire turn so earlier assistant blocks within that turn also count.
    const userMessages = Array.from(root.querySelectorAll('[data-message-author-role="user"]'))
      .filter(isRenderedConversationNode);
    const latestUserMessage = userMessages[userMessages.length - 1] || null;
    const renderedAssistantMessages = Array.from(
      root.querySelectorAll('[data-message-author-role="assistant"]'),
    ).filter((element) => (!composer || !element.contains(composer))
      && isRenderedConversationNode(element));
    const assistantMessages = renderedAssistantMessages
      .filter((element) => isNodeAfter(element, latestUserMessage));
    if (assistantMessages.length) {
      const latestMessage = assistantMessages[assistantMessages.length - 1];
      return latestMessage.closest?.('[data-testid^="conversation-turn-"]') || latestMessage;
    }
    // A newer user turn exists but its assistant turn has not mounted yet.
    // Returning an older reply here would reuse its completion marker and
    // response fingerprint for the new response session.
    if (latestUserMessage && renderedAssistantMessages.length) return null;
    const isResponseTurn = (element) => (!composer || !element.contains(composer))
      && !element.matches('[data-message-author-role="user"]')
      && !element.querySelector('[data-message-author-role="user"]');

    const numberedTurns = Array.from(root.querySelectorAll('[data-testid^="conversation-turn-"]'))
      .filter(isResponseTurn);
    if (numberedTurns.length) return numberedTurns[numberedTurns.length - 1];

    for (const selector of ['article']) {
      const candidates = Array.from(root.querySelectorAll(selector)).filter(isResponseTurn);
      if (candidates.length) return candidates[candidates.length - 1];
    }
    const contentCandidates = Array.from(root.querySelectorAll(
      '.markdown, .prose, [data-testid="assistant-message"]',
    )).filter((element) => !element.closest?.('[data-message-author-role="user"]')
      && (!composer || !element.contains(composer)));
    if (contentCandidates.length) return contentCandidates[contentCandidates.length - 1];
    return null;
  }

  function getTurnOrder(element) {
    const container = element?.matches?.('[data-testid^="conversation-turn-"]')
      ? element
      : element?.closest?.('[data-testid^="conversation-turn-"]');
    const testId = container?.getAttribute?.('data-testid') || '';
    const match = testId.match(/conversation-turn-(\d+)/i);
    if (match) return Number.parseInt(match[1], 10);

    if (element?.matches?.('[data-message-author-role="assistant"]')) {
      const assistants = Array.from(document.querySelectorAll('[data-message-author-role="assistant"]'));
      const index = assistants.indexOf(element);
      if (index >= 0) return (index * 2) + 1;
    }
    return null;
  }

  function normalizeProgressText(value) {
    return normalizeText(value)
      .replace(/(思考了|思考中|已思考|运行了|运行中)\s*(?:\d+\s*(?:小时|分钟|秒|h|m|s)\s*)+/gi, '$1 <elapsed>')
      .replace(/(thought for|thinking for|ran for|running for)\s*(?:\d+\s*(?:hours?|minutes?|seconds?|h|m|s)\s*)+/gi, '$1 <elapsed>');
  }

  function hashText(value) {
    let hash = 2166136261;
    for (let index = 0; index < value.length; index += 1) {
      hash ^= value.charCodeAt(index);
      hash = Math.imul(hash, 16777619);
    }
    return (hash >>> 0).toString(16);
  }

  function createProgressSnapshot(key, order, rawText, structure = {}) {
    const text = normalizeProgressText(structure.observedText ?? rawText);
    return {
      key,
      order: Number.isInteger(order) ? order : null,
      rawText: String(rawText || ''),
      text,
      textLength: text.length,
      signature: `${text.length}:${hashText(text)}`,
      frameCount: Number(structure.frameCount) || 0,
      toolNodeCount: Number(structure.toolNodeCount) || 0,
      descendantCount: Number(structure.descendantCount) || 0,
      toolSignature: String(structure.toolSignature || ''),
    };
  }

  function getSnapshotFingerprint(snapshot) {
    if (!snapshot) return '';
    return [
      snapshot.key,
      snapshot.signature,
      snapshot.frameCount,
      snapshot.toolNodeCount,
      snapshot.toolSignature,
    ].join('|');
  }

  function rememberResponseSnapshot(snapshot) {
    if (!snapshot) return false;
    const key = snapshot.key || 'response-fallback';
    let history = state.seenResponseFingerprints.get(key);
    if (!history) {
      history = { queue: [], values: new Set() };
      state.seenResponseFingerprints.set(key, history);
    }
    const fingerprint = getSnapshotFingerprint(snapshot);
    const isNew = !history.values.has(fingerprint);
    if (isNew) {
      history.values.add(fingerprint);
      history.queue.push(fingerprint);
      while (history.queue.length > 120) history.values.delete(history.queue.shift());
    }
    return isNew;
  }

  function hasStandaloneCompletionMarker(value, sentinel) {
    const marker = normalizeText(sentinel);
    if (!marker) return false;
    const interfaceLabels = /^(回复操作|复制回复|评价回复|分享|更多操作|朗读|重新生成|copy response|copy|feedback|share|more actions|read aloud|regenerate)$/i;
    const meaningfulLines = String(value || '')
      .split(/\r?\n/)
      .map(normalizeText)
      .filter(Boolean)
      .filter((line) => !interfaceLabels.test(line));
    return meaningfulLines.length === 1 && meaningfulLines[0] === marker;
  }

  function getResponseSnapshot() {
    const turn = getLatestTurn();
    if (!turn) return null;
    const order = getTurnOrder(turn);
    const container = turn.matches?.('[data-testid^="conversation-turn-"]')
      ? turn
      : turn.closest?.('[data-testid^="conversation-turn-"]');
    const toolScope = container || turn;
    const assistantNodes = turn.matches?.('[data-message-author-role="assistant"]')
      ? [turn]
      : Array.from(toolScope.querySelectorAll('[data-message-author-role="assistant"]'));
    const contentNodes = assistantNodes.length ? assistantNodes : [turn];
    // innerText and textContent do not always update together during streaming.
    // Track both, but use visible text for the completion-marker decision.
    const visibleText = contentNodes.map((element) => element.innerText || '').join('\n');
    const domText = contentNodes.map((element) => element.textContent || '').join('\n');
    const wholeTurnVisibleText = toolScope.innerText || '';
    const wholeTurnDomText = toolScope.textContent || '';
    const key = container?.getAttribute?.('data-testid')
      || turn.getAttribute?.('data-message-id')
      || (order !== null ? `response-${order}` : 'response-fallback');
    const frames = Array.from(toolScope.querySelectorAll('iframe'));
    const toolNodes = Array.from(toolScope.querySelectorAll([
      'iframe',
      '[data-testid*="tool"]',
      '[aria-label*="tool"]',
      '[aria-label*="工具"]',
      '[title*="tool"]',
      '[title*="工具"]',
    ].join(',')));
    const toolDetails = toolNodes.slice(-80).map((element) => {
      let embeddedText = '';
      if (element.tagName === 'IFRAME') {
        try {
          embeddedText = normalizeProgressText(
            element.contentDocument?.body?.innerText || element.contentDocument?.body?.textContent || '',
          ).slice(-500);
        } catch (_error) {
          embeddedText = '';
        }
      }
      return [
        element.tagName,
        element.getAttribute('data-testid') || '',
        element.getAttribute('src') || '',
        element.getAttribute('title') || '',
        element.getAttribute('aria-label') || '',
        element.getAttribute('aria-busy') || '',
        normalizeProgressText(element.innerText || element.textContent || '').slice(-2000),
        embeddedText,
      ].join(':');
    }).join('|');
    return createProgressSnapshot(
      key,
      order,
      visibleText || domText,
      {
        observedText: `${visibleText}\n${domText}\n${wholeTurnVisibleText}\n${wholeTurnDomText}`,
        frameCount: frames.length,
        toolNodeCount: toolNodes.length,
        descendantCount: toolScope.querySelectorAll('*').length,
        toolSignature: `${toolDetails.length}:${hashText(toolDetails)}`,
      },
    );
  }

  function isSameOrNewerSnapshot(previous, current) {
    if (!current) return false;
    if (!previous) return true;
    if (previous.order !== null && current.order !== null) return current.order >= previous.order;
    return current.key === previous.key;
  }

  function isMeaningfulSnapshotChange(previous, current) {
    if (!previous || !current) return false;
    if (previous.order !== null && current.order !== null) {
      if (current.order < previous.order) return false;
      if (current.order > previous.order) return true;
    } else if (current.key !== previous.key) {
      return false;
    }

    if (current.frameCount > previous.frameCount) return true;
    if (current.toolNodeCount > previous.toolNodeCount) return true;
    if (current.toolSignature !== previous.toolSignature) return true;
    if (current.signature === previous.signature) return false;

    // A large shrink is usually a collapsed/virtualized tool card. Do not count the
    // shrink itself, but tick() still adopts it as the next comparison baseline so
    // subsequent streaming text is visible immediately instead of waiting until the
    // reply grows past an obsolete, longer snapshot.
    const largeShrink = current.textLength + 20 < previous.textLength
      && current.textLength < previous.textLength * 0.8;
    return !largeShrink;
  }

  function captureResponseProgress(now, countAsActivity = true) {
    const responseSnapshot = getResponseSnapshot();
    const previous = state.lastResponseSnapshot;
    const switchedBranch = Boolean(state.allowBranchSwitch && previous && responseSnapshot
      && responseSnapshot.key !== previous.key);
    if (!switchedBranch && !isSameOrNewerSnapshot(previous, responseSnapshot)) return false;
    const isNew = rememberResponseSnapshot(responseSnapshot);
    const isProgress = isNew && (!previous || switchedBranch
      || isMeaningfulSnapshotChange(previous, responseSnapshot));
    state.lastResponseSnapshot = responseSnapshot;
    if (switchedBranch) state.allowBranchSwitch = false;
    if (isProgress && countAsActivity) {
      state.lastContentActivityAt = now;
      state.progressRevision += 1;
      state.lastProgressSource = '回复快照';
    }
    return isProgress && countAsActivity;
  }

  function isResponseActivityNode(element, latestUserMessage, composer) {
    if (!element) return false;
    if (composer && (element === composer || composer.contains?.(element))) return false;
    if (element.closest?.('button, [contenteditable="true"], [role="textbox"], '
      + '[data-message-author-role="user"]')) return false;
    if (latestUserMessage) {
      if (element === latestUserMessage || latestUserMessage.contains?.(element)) return false;
      if (!isNodeAfter(element, latestUserMessage)) return false;
    }
    return true;
  }

  function getResponseActivityMutation(records) {
    const root = getThreadRoot();
    if (!root) return '';
    const composer = findComposer();
    const users = Array.from(root.querySelectorAll('[data-message-author-role="user"]'))
      .filter(isRenderedConversationNode);
    const latestUserMessage = users[users.length - 1] || null;

    for (const record of records) {
      const target = record.target?.nodeType === 3
        ? record.target.parentElement : record.target;
      if (record.type === 'characterData'
        && isResponseActivityNode(target, latestUserMessage, composer)) {
        const before = normalizeProgressText(record.oldValue || '');
        const after = normalizeProgressText(record.target.data || '');
        if (after && after !== before) return '工具/文本节点';
      }
      if (record.type === 'childList') {
        const relevantAddedNodes = Array.from(record.addedNodes || []).filter((node) => {
          const element = node.nodeType === 1 ? node : node.parentElement || target;
          return isResponseActivityNode(element, latestUserMessage, composer);
        });
        const after = normalizeProgressText(relevantAddedNodes
          .map((node) => node.textContent || '').join('\n'));
        const before = normalizeProgressText(Array.from(record.removedNodes || [])
          .map((node) => node.textContent || '').join('\n'));
        if (after && after !== before) return '工具/新增节点';
      }
    }
    return '';
  }

  function resetResponseTracking(now, allowBranchSwitch = false) {
    state.lastResponseSnapshot = getResponseSnapshot();
    state.seenResponseFingerprints = new Map();
    rememberResponseSnapshot(state.lastResponseSnapshot);
    state.progressRevision = 0;
    state.handledProgressRevision = 0;
    state.allowBranchSwitch = allowBranchSwitch;
    state.lastContentActivityAt = now;
    state.lastProgressSource = '重新绑定';
    state.ignoredCompletionMarkerFingerprint = hasStandaloneCompletionMarker(
      state.lastResponseSnapshot?.rawText,
      settings.sentinel,
    ) ? getSnapshotFingerprint(state.lastResponseSnapshot) : '';
  }

  function installProgressObserver() {
    const root = getThreadRoot();
    if (root === state.observedRoot && state.progressObserver) return;
    state.progressObserver?.disconnect();
    state.progressObserver = null;
    state.observedRoot = root;
    if (!root || typeof MutationObserver !== 'function') return;
    state.progressObserver = new MutationObserver((records) => {
      if (!state.enabled || state.busy || getConversationKey() !== state.conversationKey) return;
      const stoppedAnswer = state.stopWasSeenInCurrentResponse && !isGenerating();
      const previousFingerprint = getSnapshotFingerprint(state.lastResponseSnapshot);
      const now = Date.now();
      const snapshotProgress = captureResponseProgress(now, !stoppedAnswer);
      // A streaming text node can change before the rendered turn snapshot
      // catches up. Count that mutation only when the snapshot stayed identical.
      const activitySource = !stoppedAnswer && !snapshotProgress
        && previousFingerprint === getSnapshotFingerprint(state.lastResponseSnapshot)
        ? getResponseActivityMutation(records) : '';
      if (activitySource) {
        state.lastContentActivityAt = now;
        state.progressRevision += 1;
        state.lastProgressSource = activitySource;
      }
    });
    state.progressObserver.observe(root, {
      subtree: true,
      childList: true,
      characterData: true,
      characterDataOldValue: true,
    });
  }

  function getUserTurns() {
    const root = getThreadRoot() || document;
    return Array.from(root.querySelectorAll('[data-message-author-role="user"]'));
  }

  function getLatestUserMessageText() {
    const users = getUserTurns();
    if (!users.length) return '';
    return normalizeText(users[users.length - 1].innerText || users[users.length - 1].textContent || '');
  }

  function getUserSignature() {
    const users = getUserTurns();
    if (!users.length) return '0:';
    const text = getLatestUserMessageText();
    return `${users.length}:${text.length}:${hashText(text)}`;
  }

  function readSettingsFromUi() {
    if (!ui) return settings;
    return {
      idleSeconds: clampInteger(ui.idleInput.value, 10, 3600, settings.idleSeconds),
      maxContinues: clampInteger(ui.maxInput.value, 1, 100, settings.maxContinues),
      sentinel: normalizeSentinel(ui.sentinelInput.value),
    };
  }

  function setStatus(message) {
    state.status = message;
    render();
  }

  function start() {
    settings = readSettingsFromUi();
    saveSettings();
    const now = Date.now();
    state.enabled = true;
    state.busy = false;
    state.conversationKey = getConversationKey();
    state.generationActive = isGenerating();
    state.hasSeenGeneration = state.generationActive;
    state.idleSince = now;
    resetResponseTracking(now, false);
    installProgressObserver();
    state.lastUserSignature = getUserSignature();
    state.lastComposerText = getComposerText(findComposer());
    state.lastComposerInputAt = 0;
    state.stopMissingSince = 0;
    state.stopWasSeenInCurrentResponse = state.generationActive;
    state.pendingSubmission = false;
    state.pendingSince = 0;
    state.pendingKind = '';
    state.pendingExpectedUserText = '';
    state.pendingUserSignatureBefore = '';
    state.pendingExpectedMessageCountBefore = 0;
    state.pendingAction = '';
    state.nextActionNotBefore = 0;
    state.continueCount = 0;
    setStatus(state.generationActive ? '已检测到：GPT 正在回答' : '已检测到：GPT 当前空闲');
  }

  function pause(reason = '已暂停') {
    state.enabled = false;
    state.busy = false;
    state.generationActive = false;
    state.stopMissingSince = 0;
    state.pendingSubmission = false;
    state.pendingKind = '';
    state.pendingExpectedUserText = '';
    state.pendingUserSignatureBefore = '';
    state.pendingExpectedMessageCountBefore = 0;
    state.pendingAction = '';
    state.progressObserver?.disconnect();
    state.progressObserver = null;
    state.observedRoot = null;
    setStatus(reason);
  }

  async function requestContinue(triggerLabel, { stopIfGenerating = false } = {}) {
    if (state.busy) return;
    if (state.pendingSubmission) {
      setStatus('上次续写尚未确认，暂不重复发送');
      return;
    }
    const generatingAtStart = isGenerating();
    if (generatingAtStart && !stopIfGenerating) {
      setStatus('GPT 仍在回答，本次不发送');
      return;
    }
    if (state.continueCount >= settings.maxContinues) {
      pause(`已达到上限 ${settings.maxContinues} 次`);
      return;
    }

    state.busy = true;
    state.nextActionNotBefore = Date.now() + 10000;
    setStatus(`${triggerLabel}：正在发送续写请求`);
    try {
      if (generatingAtStart) {
        setStatus(`${triggerLabel}：当前输入框仍显示停止按钮`);
        await stopStalledGeneration();
        setStatus('回答已停止，正在发送“继续”');
      } else {
        const composerState = getComposerControlState();
        const stateLabel = composerState === 'voice' ? '空输入框和语音按钮'
          : composerState === 'send' ? '可发送状态'
            : composerState === 'text-no-send' ? '有文字但没有发送按钮'
              : '空闲状态';
        setStatus(`${triggerLabel}：当前输入框为${stateLabel}，开始写入并发送“继续”`);
      }
      const result = await sendContinueMessage();
      const completedAt = Date.now();
      if (!result.confirmed && result.kind === 'message') {
        pause(result.message);
        return;
      }
      if (result.confirmed) state.continueCount += 1;
      state.idleSince = completedAt;
      resetResponseTracking(completedAt, true);
      state.lastUserSignature = getUserSignature();
      state.lastComposerText = getComposerText(findComposer());
      state.lastComposerInputAt = 0;
      state.stopMissingSince = 0;
      state.generationActive = isGenerating();
      state.stopWasSeenInCurrentResponse = state.generationActive;
      state.pendingSubmission = !result.confirmed;
      state.pendingSince = state.pendingSubmission ? completedAt : 0;
      state.pendingKind = state.pendingSubmission ? result.kind : '';
      state.pendingExpectedUserText = state.pendingSubmission ? result.expectedUserText : '';
      state.pendingUserSignatureBefore = state.pendingSubmission ? result.userSignatureBefore : '';
      state.pendingExpectedMessageCountBefore = state.pendingSubmission
        ? result.expectedMessageCountBefore : 0;
      state.pendingAction = state.pendingSubmission ? result.action : '';
      state.nextActionNotBefore = completedAt + (state.pendingSubmission ? 120000 : 10000);
      setStatus(result.message);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      if (/(网页编辑器未接受|编辑器回滚|输入区有附件|无法设置 ChatGPT 输入框|停止回答.*仍未切换)/u.test(message)) {
        pause(message);
      } else {
        state.nextActionNotBefore = Date.now() + 10000;
        setStatus(message);
      }
    } finally {
      state.busy = false;
      render();
    }
  }

  async function tick() {
    if (!state.enabled || state.busy) {
      render();
      return;
    }
    if (getConversationKey() !== state.conversationKey) {
      pause('页面或会话已切换，已暂停');
      return;
    }

    const now = Date.now();
    installProgressObserver();
    const composerText = getComposerText(findComposer());
    if (composerText !== state.lastComposerText) {
      state.lastComposerText = composerText;
      state.lastComposerInputAt = now;
    }
    const userSignature = getUserSignature();
    const userSignatureChanged = userSignature !== state.lastUserSignature;
    if (userSignatureChanged) {
      state.lastUserSignature = userSignature;
      state.idleSince = now;
      state.lastContentActivityAt = now;
      state.stopWasSeenInCurrentResponse = false;
      state.stopMissingSince = 0;
      state.seenResponseFingerprints = new Map();
      state.allowBranchSwitch = true;
      state.ignoredCompletionMarkerFingerprint = '';
    }

    const requiredIdleMilliseconds = settings.idleSeconds * 1000;
    const generatingNow = isGenerating();
    const stoppedAnswer = state.stopWasSeenInCurrentResponse && !generatingNow;
    captureResponseProgress(now, !stoppedAnswer);
    const responseProgress = state.progressRevision !== state.handledProgressRevision;
    state.handledProgressRevision = state.progressRevision;
    const responseSnapshot = state.lastResponseSnapshot;
    const completionFingerprint = getSnapshotFingerprint(state.lastResponseSnapshot);
    const hasCompletionMarker = hasStandaloneCompletionMarker(
      state.lastResponseSnapshot?.rawText,
      settings.sentinel,
    );
    if (hasCompletionMarker
      && completionFingerprint !== state.ignoredCompletionMarkerFingerprint) {
        pause(`检测到完成标记 ${settings.sentinel}`);
        return;
    }
    if (!hasCompletionMarker && state.lastResponseSnapshot) {
      state.ignoredCompletionMarkerFingerprint = '';
    }

    if (state.pendingSubmission) {
      const composerStillContainsExpected = normalizeText(getComposerText(findComposer()))
        === normalizeText(state.pendingExpectedUserText);
      const exactMessageAppeared = countSubmittedMessages(state.pendingExpectedUserText)
        > state.pendingExpectedMessageCountBefore;
      const pendingConfirmed = state.pendingKind === 'message'
        ? exactMessageAppeared
          || (!composerStillContainsExpected && (responseProgress || generatingNow))
        : responseProgress || generatingNow;
      if (pendingConfirmed) {
        state.continueCount += 1;
        state.pendingSubmission = false;
        state.pendingSince = 0;
        state.pendingKind = '';
        state.pendingExpectedUserText = '';
        state.pendingUserSignatureBefore = '';
        state.pendingExpectedMessageCountBefore = 0;
        state.pendingAction = '';
        state.nextActionNotBefore = now + 10000;
      } else if (now - state.pendingSince < 120000) {
        setStatus('已点击续写，等待页面确认；不会重复发送');
        return;
      } else {
        const action = state.pendingAction || '续写操作';
        pause(`${action}后仍无法确认是否送达；为避免重复发送已暂停`);
        return;
      }
    }

    if (!responseSnapshot && !generatingNow && !responseProgress) {
      setStatus('尚未定位到助手回复区域，等待页面加载；不会盲目发送');
      return;
    }

    if (generatingNow) {
      state.stopWasSeenInCurrentResponse = true;
      state.stopMissingSince = 0;
      if (!state.generationActive || responseProgress) state.lastContentActivityAt = now;
      state.generationActive = true;
      state.hasSeenGeneration = true;
      state.idleSince = 0;
      const noUpdateMilliseconds = now - state.lastContentActivityAt;
      if (now >= state.nextActionNotBefore && noUpdateMilliseconds >= requiredIdleMilliseconds
        && now - state.lastComposerInputAt >= COMPOSER_INPUT_GRACE_MS) {
        await requestContinue('回答超过阈值没有新内容', { stopIfGenerating: true });
        return;
      }
      setStatus('已检测到：GPT 正在回答');
      return;
    }

    if (state.stopWasSeenInCurrentResponse) {
      // Once an observed Stop control disappears, later old-content redraws
      // cannot override the stopped-answer clock.
      if (state.generationActive) {
        if (!state.stopMissingSince) state.stopMissingSince = now;
        const typingRecently = now - state.lastComposerInputAt < COMPOSER_INPUT_GRACE_MS;
        if (now - state.stopMissingSince < STOP_MISSING_GRACE_MS || typingRecently) {
          setStatus('正在确认回答是否停止');
          return;
        }
        state.generationActive = false;
        state.idleSince = state.stopMissingSince;
        state.stopMissingSince = 0;
      }
    } else if (responseProgress) {
      // Some ChatGPT layouts never expose a Stop control. New assistant content
      // still proves that a response is in progress and refreshes its timer.
      state.generationActive = true;
      state.hasSeenGeneration = true;
      state.lastContentActivityAt = now;
      state.idleSince = 0;
    }

    if (state.generationActive) {
      const noUpdateMilliseconds = now - state.lastContentActivityAt;
      if (noUpdateMilliseconds < requiredIdleMilliseconds) {
        setStatus('助手内容正在更新，静默计时中');
        return;
      }
      state.generationActive = false;
      state.idleSince = state.lastContentActivityAt;
    }

    if (!state.idleSince) state.idleSince = now;
    const quietSince = state.idleSince;
    const idleMilliseconds = now - quietSince;

    if (now - state.lastComposerInputAt < COMPOSER_INPUT_GRACE_MS) {
      setStatus('正在输入，暂缓续写');
      return;
    }
    if (now < state.nextActionNotBefore || idleMilliseconds < requiredIdleMilliseconds) {
      setStatus(state.hasSeenGeneration ? '回答已停止，静默计时中' : 'GPT 当前空闲，静默计时中');
      return;
    }
    await requestContinue('已静默到设定时间');
  }

  function render() {
    if (!ui) return;
    ui.toggleButton.textContent = state.enabled ? '暂停' : '开始监督';
    ui.toggleButton.classList.toggle('running', state.enabled);
    ui.testButton.disabled = state.busy || state.generationActive || state.pendingSubmission;
    ui.idleInput.disabled = state.enabled;
    ui.maxInput.disabled = state.enabled;
    ui.sentinelInput.disabled = state.enabled;

    let detail = state.status;
    if (state.enabled && !state.busy) {
      if (state.generationActive) {
        const quietSeconds = Math.max(0, Math.floor((Date.now() - state.lastContentActivityAt) / 1000));
        detail = `${state.status} · ${state.lastProgressSource} ${quietSeconds}s 前更新 · 已续 ${state.continueCount}/${settings.maxContinues}`;
      } else {
        const quietSince = state.idleSince || Date.now();
        const remainingSeconds = Math.max(0, Math.ceil(
          (settings.idleSeconds * 1000 - (Date.now() - quietSince)) / 1000,
        ));
        detail = `${state.status} · ${remainingSeconds}s 后续写 · 已续 ${state.continueCount}/${settings.maxContinues}`;
      }
    }
    ui.status.textContent = detail;
  }

  function createPanel() {
    if (document.getElementById(PANEL_ID)) return;
    const host = document.createElement('div');
    host.id = PANEL_ID;
    const shadow = host.attachShadow({ mode: 'open' });
    shadow.innerHTML = `
      <style>
        :host { all: initial; }
        .panel { position: fixed; right: 16px; bottom: 16px; z-index: 2147483647; width: 286px; box-sizing: border-box; padding: 12px; border: 1px solid rgba(255,255,255,.18); border-radius: 12px; background: rgba(25,25,25,.95); color: #f5f5f5; box-shadow: 0 8px 30px rgba(0,0,0,.28); font: 13px/1.4 -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif; }
        .title { margin-bottom: 9px; font-size: 14px; font-weight: 650; }
        .grid { display: grid; grid-template-columns: 1fr 108px; gap: 7px 8px; align-items: center; }
        label { color: #d4d4d4; }
        input { min-width: 0; box-sizing: border-box; width: 100%; padding: 5px 7px; border: 1px solid #555; border-radius: 6px; background: #111; color: #fff; font: inherit; }
        input:disabled { opacity: .6; }
        .actions { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin-top: 10px; }
        button { box-sizing: border-box; width: 100%; padding: 7px 8px; border: 0; border-radius: 7px; background: #10a37f; color: white; cursor: pointer; font: inherit; font-weight: 650; }
        button.secondary { background: #3f3f46; }
        button.running { background: #b45309; }
        button:disabled { cursor: not-allowed; opacity: .5; }
        .status { min-height: 36px; margin-top: 8px; color: #b8e6da; overflow-wrap: anywhere; }
        .note { margin-top: 5px; color: #a8a8a8; font-size: 11px; }
      </style>
      <section class="panel" aria-label="ChatGPT 自动继续监督器">
        <div class="title">ChatGPT 自动继续 v2.1</div>
        <div class="grid">
          <label for="idle">无更新阈值</label><input id="idle" type="number" min="10" max="3600" step="1">
          <label for="max">最多续写</label><input id="max" type="number" min="1" max="100" step="1">
          <label for="sentinel">完成标记</label><input id="sentinel" type="text">
        </div>
        <div class="actions">
          <button id="toggle" type="button">开始监督</button>
          <button id="test" class="secondary" type="button">测试续写一次</button>
        </div>
        <div id="status" class="status" role="status"></div>
        <div class="note">只有“继续”消息进入对话才算发送成功；按钮未提交时最多补一次表单提交。</div>
      </section>
    `;
    document.documentElement.appendChild(host);

    ui = {
      idleInput: shadow.getElementById('idle'),
      maxInput: shadow.getElementById('max'),
      sentinelInput: shadow.getElementById('sentinel'),
      toggleButton: shadow.getElementById('toggle'),
      testButton: shadow.getElementById('test'),
      status: shadow.getElementById('status'),
    };
    ui.idleInput.value = String(settings.idleSeconds);
    ui.maxInput.value = String(settings.maxContinues);
    ui.sentinelInput.value = settings.sentinel;
    ui.toggleButton.addEventListener('click', () => {
      if (state.enabled) pause();
      else start();
    });
    ui.testButton.addEventListener('click', () => {
      settings = readSettingsFromUi();
      saveSettings();
      void requestContinue('手动测试');
    });
    render();
  }

  if (globalThis.__CHATGPT_AUTO_CONTINUE_TEST_MODE__) {
    globalThis.__chatgptAutoContinueTestApi = {
      getComposerText,
      setContentEditableText,
      getSnapshotFingerprint,
      createProgressSnapshot,
      hasStandaloneCompletionMarker,
      isGenerating,
      getComposerControlState,
      findSendButton,
      isMeaningfulSnapshotChange,
      normalizeProgressText,
      getLatestTurn,
      getResponseSnapshot,
      resetResponseTracking,
      start,
      tick,
      state,
      requestContinue,
      sendContinueMessage,
      stopStalledGeneration,
    };
    return;
  }

  createPanel();
  window.setInterval(() => { void tick(); }, 1000);
})();
