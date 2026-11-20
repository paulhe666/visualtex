'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const userMessages = [];

class FakeInput {
  constructor(value = '') {
    this._value = value;
    this.disabled = false;
    this.isConnected = true;
  }

  get value() { return this._value; }
  set value(value) { this._value = String(value); }
  focus() {}
  dispatchEvent() {}
  getBoundingClientRect() { return { width: 300, height: 40 }; }
  closest(selector) { return selector === 'form' ? fakeForm : null; }
}

class FakeButton {
  constructor(composer) {
    this.composer = composer;
    this.disabled = false;
    this.isConnected = true;
    this.innerText = '';
    this.attributes = {
      'aria-label': '发送提示词',
      'aria-disabled': 'false',
      'data-testid': 'send-button',
      title: '',
    };
    this.sentText = '';
    this.sendClicks = 0;
    this.showStopAfterSend = true;
    this.showUserTurnAfterSend = true;
    this.stopClicks = 0;
    this.stopFailuresBeforeSuccess = 0;
    this.sendFailuresBeforeSuccess = 0;
  }

  getAttribute(name) { return this.attributes[name] ?? null; }
  getBoundingClientRect() { return { width: 40, height: 40 }; }
  click() {
    if (/停止回答/u.test(this.attributes['aria-label'])) {
      this.stopClicks += 1;
      if (this.stopFailuresBeforeSuccess > 0) {
        this.stopFailuresBeforeSuccess -= 1;
        return;
      }
      this.attributes['aria-label'] = '发送提示词';
      this.attributes['data-testid'] = 'send-button';
      return;
    }
    this.sendClicks += 1;
    this.sentText = this.composer.value;
    if (this.sendFailuresBeforeSuccess > 0) {
      this.sendFailuresBeforeSuccess -= 1;
      return;
    }
    this.composer.value = '';
    if (this.showUserTurnAfterSend) {
      userMessages.push({ innerText: this.sentText, textContent: this.sentText });
    }
    if (this.showStopAfterSend) {
      this.attributes['aria-label'] = '停止回答';
      this.attributes['data-testid'] = 'stop-button';
    }
  }
}

const fakeForm = {
  requestSubmitCalls: 0,
  requestSubmit() {
    this.requestSubmitCalls += 1;
    submitButton.click();
  },
  querySelector() { return null; },
  querySelectorAll(selector) {
    const controls = dedicatedStopButton ? [submitButton, dedicatedStopButton] : [submitButton];
    if (selector === 'button') return controls;
    if (selector === 'button[type="submit"]' || selector === '#composer-submit-button') return [submitButton];
    if (selector === 'button[data-testid="stop-button"]') {
      return controls.filter((button) => button.getAttribute('data-testid') === 'stop-button');
    }
    if (selector === '#composer-stop-button') return [];
    if (selector === 'button[data-testid="send-button"]') {
      return controls.filter((button) => button.getAttribute('data-testid') === 'send-button');
    }
    if (selector === 'button[data-testid="fruitjuice-send-button"]') return [];
    return [];
  },
};
const composer = new FakeInput('用户未发送的草稿');
const submitButton = new FakeButton(composer);
let dedicatedStopButton = null;
let staleStopButton = null;
let staleRetryButton = null;
let threadRoot = null;
let disabledComposerSubmit = null;

const document = {
  querySelector(selector) {
    if (selector === '#composer-submit-button') return submitButton;
    if (selector === 'main') return threadRoot;
    return null;
  },
  querySelectorAll(selector) {
    if (selector === 'button') {
      return [submitButton, dedicatedStopButton, staleStopButton, staleRetryButton].filter(Boolean);
    }
    if (selector === 'button[data-testid="stop-button"]') {
      return [dedicatedStopButton, staleStopButton].filter(Boolean);
    }
    if (selector === '#composer-submit-button') return disabledComposerSubmit ? [disabledComposerSubmit] : [submitButton];
    if (selector === 'button[data-testid="send-button"]') return submitButton.attributes['data-testid'] === 'send-button' ? [submitButton] : [];
    if (selector === '#prompt-textarea') return [composer];
    if (selector === '[data-message-author-role="user"]') return userMessages;
    return [];
  },
};

const sessionValues = new Map();
const sessionStorage = {
  getItem(key) { return sessionValues.has(key) ? sessionValues.get(key) : null; },
  setItem(key, value) { sessionValues.set(key, String(value)); },
  removeItem(key) { sessionValues.delete(key); },
};
let reloadCalls = 0;
const location = {
  pathname: '/c/test',
  search: '',
  reload() { reloadCalls += 1; },
};

const context = vm.createContext({
  __CHATGPT_AUTO_CONTINUE_TEST_MODE__: true,
  console,
  document,
  Event: class Event {},
  HTMLInputElement: class HTMLInputElement {},
  HTMLTextAreaElement: FakeInput,
  InputEvent: class InputEvent {},
  localStorage: { getItem() { return null; }, setItem() {} },
  sessionStorage,
  location,
  getComputedStyle() {
    return { display: 'block', visibility: 'visible', opacity: '1' };
  },
  setTimeout,
  clearTimeout,
  window: { setTimeout, setInterval() {} },
});

const scriptPath = path.join(__dirname, 'chatgpt-auto-continue.user.js');
const source = fs.readFileSync(scriptPath, 'utf8');
vm.runInContext(source, context, { filename: scriptPath });

const api = context.__chatgptAutoContinueTestApi;
assert.ok(api, 'test API should be exposed in test mode');
assert.equal(api.isGenerating(), false, '发送提示词 must mean idle');

const timer60 = api.createProgressSnapshot('conversation-turn-9', 9, '思考中 60s Waiting for a tool result.');
const timer61 = api.createProgressSnapshot('conversation-turn-9', 9, '思考中 61s Waiting for a tool result.');
assert.equal(timer60.signature, timer61.signature, 'elapsed-time labels must not count as progress');

const olderTurn = api.createProgressSnapshot('conversation-turn-7', 7, '滚动后重新渲染出来的更早回复，而且文字更长');
assert.equal(api.isMeaningfulSnapshotChange(timer60, olderTurn), false, 'scrolling to an older turn must not reset activity');

const grownTurn = api.createProgressSnapshot('conversation-turn-9', 9, '思考中 61s Waiting for a tool result. 新的有效输出');
assert.equal(api.isMeaningfulSnapshotChange(timer60, grownTurn), true, 'new text in the same turn must count as progress');

const collapsedTurn = api.createProgressSnapshot('conversation-turn-9', 9, 'Waiting');
assert.equal(api.isMeaningfulSnapshotChange(timer60, collapsedTurn), false, 'collapsed or virtualized content must not count as progress');

const resumedAfterCollapse = api.createProgressSnapshot('conversation-turn-9', 9, 'Waiting 新吐出的信息');
assert.equal(api.isMeaningfulSnapshotChange(collapsedTurn, resumedAfterCollapse), true, 'streaming after a collapsed render must be detected against the current baseline');

const waitingTool = api.createProgressSnapshot('conversation-turn-9', 9, '正在思考', {
  frameCount: 3,
  toolNodeCount: 3,
  descendantCount: 100,
  toolSignature: 'waiting',
});
const newToolResult = api.createProgressSnapshot('conversation-turn-9', 9, '正在思考', {
  frameCount: 4,
  toolNodeCount: 4,
  descendantCount: 120,
  toolSignature: 'result',
});
assert.equal(api.isMeaningfulSnapshotChange(waitingTool, newToolResult), true, 'a newly rendered tool result must count as progress even without parent text');
assert.notEqual(api.getSnapshotFingerprint(waitingTool), api.getSnapshotFingerprint(newToolResult), 'tool-card changes must produce a different progress fingerprint');

const unfinishedText = '还未完成的是后三个真实 Ribbon 方向、10.5/11.5pt 特殊源字号回归。因此目前不能回复 [TASK_DONE]，也还没有生成最终安装包。';
assert.equal(api.hasStandaloneCompletionMarker(unfinishedText, '[TASK_DONE]'), false, 'mentioning the marker in an unfinished sentence must not stop continuation');
assert.equal(api.hasStandaloneCompletionMarker('[TASK_DONE]', '[TASK_DONE]'), true, 'a standalone marker must stop continuation');
assert.equal(api.hasStandaloneCompletionMarker('目前不能回复：\n[TASK_DONE]', '[TASK_DONE]'), false, 'a marker following unfinished prose must not stop continuation');
assert.equal(api.hasStandaloneCompletionMarker('[TASK_DONE]\n复制回复\n分享\n更多操作', '[TASK_DONE]'), true, 'ChatGPT action labels after an otherwise exact marker should be ignored');

submitButton.attributes['aria-label'] = '停止回答';
assert.equal(api.isGenerating(), true, '停止回答 must mean generating');

(async () => {
  await api.stopStalledGeneration();
  assert.equal(api.isGenerating(), false, 'a stalled generation should be stopped');
  assert.equal(submitButton.stopClicks, 1, 'the stop button should be clicked once');

  let staleRetryClicks = 0;
  staleRetryButton = new FakeButton(composer);
  staleRetryButton.innerText = '重试';
  staleRetryButton.attributes['aria-label'] = '重试';
  staleRetryButton.attributes['data-testid'] = 'retry-button';
  staleRetryButton.click = () => { staleRetryClicks += 1; };
  const result = await api.sendContinueMessage();
  assert.equal(staleRetryClicks, 0,
    'a stale page-wide Retry button must be ignored during idle continuation');
  assert.match(submitButton.sentText, /^继续。/u, 'the continuation prompt should be sent');
  assert.equal(composer.value, '用户未发送的草稿', 'the existing draft should be restored');
  assert.equal(result.confirmed, true, 'the new user turn confirms the send');
  assert.match(result.message, /草稿已恢复/u);
  staleRetryButton = null;

  await api.stopStalledGeneration();
  submitButton.stopClicks = 0;
  submitButton.stopFailuresBeforeSuccess = 1;
  submitButton.attributes['aria-label'] = '停止回答';
  submitButton.attributes['data-testid'] = 'stop-button';
  await assert.rejects(api.stopStalledGeneration(), /仍未切换为空闲状态/u);
  assert.equal(submitButton.stopClicks, 1,
    'a real Stop control should be clicked once and observed instead of blindly retried');
  submitButton.stopFailuresBeforeSuccess = 0;
  submitButton.attributes['aria-label'] = '发送提示词';
  submitButton.attributes['data-testid'] = 'send-button';

  composer.value = '';
  submitButton.sendFailuresBeforeSuccess = 1;
  const sendClicksBeforeFallback = submitButton.sendClicks;
  const fallbackResult = await api.sendContinueMessage();
  assert.equal(fallbackResult.confirmed, true,
    'form submission should recover when the Send click leaves the prompt untouched');
  assert.equal(submitButton.sendClicks, sendClicksBeforeFallback + 2,
    'the fallback should submit only once after the failed click');

  await api.stopStalledGeneration();
  composer.value = '用户未发送的草稿';
  submitButton.sendFailuresBeforeSuccess = 2;
  const sendClicksBeforeFailure = submitButton.sendClicks;
  const unconfirmedResult = await api.sendContinueMessage();
  assert.equal(unconfirmedResult.confirmed, false, 'an unconfirmed click is reported honestly');
  assert.equal(submitButton.sendClicks, sendClicksBeforeFailure + 1,
    'an unconfirmed send with a saved draft must not retry or submit that draft');
  assert.equal(composer.value, '用户未发送的草稿', 'an unconfirmed send restores the draft');

  submitButton.sendFailuresBeforeSuccess = 0;
  submitButton.showStopAfterSend = false;
  const noStopResult = await api.sendContinueMessage();
  assert.equal(noStopResult.confirmed, true, 'a new user turn confirms sending without a Stop control');
  assert.match(noStopResult.message, /消息已出现在对话中/u);
  submitButton.showUserTurnAfterSend = false;
  const clearedWithoutUserTurn = await api.sendContinueMessage();
  assert.equal(clearedWithoutUserTurn.confirmed, false,
    'a cleared composer alone must not claim that the message was sent');
  submitButton.showStopAfterSend = true;
  submitButton.attributes['aria-label'] = '发送提示词';
  submitButton.attributes['data-testid'] = 'send-button';
  const responseStateConfirmation = await api.sendContinueMessage();
  assert.equal(responseStateConfirmation.confirmed, true,
    'a cleared composer plus the current composer Stop control confirms that an idle send started a response');
  assert.match(responseStateConfirmation.message, /当前输入框显示正在回答/u);
  await api.stopStalledGeneration();
  submitButton.showUserTurnAfterSend = true;
  submitButton.showStopAfterSend = true;

  composer.value = '';
  submitButton.attributes['aria-label'] = '开始语音模式';
  submitButton.attributes['data-testid'] = 'composer-speech-button';
  assert.equal(api.getComposerControlState(), 'voice',
    'an empty composer with the blue voice control is an explicit idle state');
  const formSubmitsBeforeVoiceState = fakeForm.requestSubmitCalls;
  const voiceStateResult = await api.sendContinueMessage();
  assert.equal(voiceStateResult.confirmed, true,
    'an empty composer showing the voice button must submit through the form');
  assert.equal(fakeForm.requestSubmitCalls, formSubmitsBeforeVoiceState + 1,
    'the voice button must not be clicked as though it were Send');
  assert.match(voiceStateResult.message, /表单提交/u);
  await api.stopStalledGeneration();

  composer.value = '';
  submitButton.attributes['aria-label'] = '开始语音模式';
  submitButton.attributes['data-testid'] = 'composer-speech-button';
  const originalComposerClosest = composer.closest;
  const originalComposerDispatchEvent = composer.dispatchEvent;
  composer.closest = () => null;
  context.KeyboardEvent = class KeyboardEvent {
    constructor(type) { this.type = type; }
  };
  composer.dispatchEvent = (event) => {
    if (event.type !== 'keydown') return;
    const sentText = composer.value;
    composer.value = '';
    userMessages.push({ innerText: sentText, textContent: sentText });
  };
  const enterResult = await api.sendContinueMessage();
  assert.equal(enterResult.confirmed, true,
    'Enter must be attempted once when neither a Send button nor a form is available');
  assert.match(enterResult.message, /回车提交/u);
  composer.closest = originalComposerClosest;
  composer.dispatchEvent = originalComposerDispatchEvent;
  submitButton.attributes['aria-label'] = '发送提示词';
  submitButton.attributes['data-testid'] = 'send-button';
  composer.value = '用户未发送的草稿';

  submitButton.sendFailuresBeforeSuccess = 2;
  const sendClicksBeforeRequest = submitButton.sendClicks;
  await api.requestContinue('测试未确认发送');
  assert.equal(submitButton.sendClicks, sendClicksBeforeRequest + 1,
    'the full continuation flow must not retry when a draft is present');
  assert.equal(api.state.pendingSubmission, false,
    'an unconfirmed message attempt must stop with a precise result instead of entering a vague pending state');
  assert.equal(api.state.continueCount, 0, 'unconfirmed clicks must not count as successful continuations');
  assert.equal(api.state.enabled, false, 'requestContinue must not silently enable supervision');
  assert.match(api.state.status, /即将刷新当前页面后重试一次/u,
    'an unconfirmed send must enter the one-time page refresh recovery');
  await new Promise((resolve) => setTimeout(resolve, 10));
  assert.equal(reloadCalls, 1,
    'the full continuation flow should refresh once after an unconfirmed send');
  assert.ok(api.readSendRecovery(), 'the failed continuation must survive the refresh');
  const messagesBeforeReloadRecovery = userMessages.length;
  submitButton.sendFailuresBeforeSuccess = 0;
  await api.recoverSendAfterReload();
  assert.equal(userMessages.length, messagesBeforeReloadRecovery + 1,
    'after refresh the saved continuation must be sent through the recovered composer');
  assert.equal(api.readSendRecovery(), null,
    'successful post-refresh sending must consume the recovery marker');
  assert.equal(reloadCalls, 1,
    'a failed post-refresh attempt must not start a refresh loop');
  await api.stopStalledGeneration();
  reloadCalls = 0;

  disabledComposerSubmit = new FakeButton(composer);
  disabledComposerSubmit.disabled = true;
  assert.equal(api.findSendButton(), submitButton,
    'a disabled primary control must not hide another enabled Send button');
  disabledComposerSubmit = null;

  let replacedEditorDom = false;
  const editable = {
    innerText: '草稿',
    textContent: '草稿',
    focus() {},
    dispatchEvent() {},
    replaceChildren() { replacedEditorDom = true; },
  };
  document.createRange = () => ({ selectNodeContents() {} });
  document.execCommand = () => false;
  context.window.getSelection = () => ({ removeAllRanges() {}, addRange() {} });
  assert.throws(() => api.setContentEditableText(editable, '继续'), /网页编辑器未接受/u,
    'an editor that rejects insertion must fail before any Send click');
  assert.equal(replacedEditorDom, false,
    'the script must not fake a successful editor update by replacing DOM children');
  const rangeEditable = {
    _text: '旧草稿',
    get innerText() { return this._text; },
    get textContent() { return this._text; },
    focus() {},
    dispatchEvent() {},
  };
  document.createRange = () => ({
    selectNodeContents() {},
    deleteContents() { rangeEditable._text = ''; },
    insertNode(node) { rangeEditable._text = node.textContent; },
    setStartAfter() {},
    collapse() {},
  });
  document.createTextNode = (text) => ({ textContent: text });
  api.setContentEditableText(rangeEditable, '继续');
  assert.equal(rangeEditable.innerText, '继续',
    'a selection mutation must recover when execCommand is rejected');
  const proseMirrorEditable = {
    _text: '',
    get innerText() { return this._text; },
    get textContent() { return this._text; },
    focus() {},
    querySelectorAll() { return []; },
  };
  const proseMirrorState = {
    doc: { content: { size: 2 }, textContent: '' },
    schema: {
      text(value) { return { textContent: value }; },
      nodes: {
        paragraph: {
          create(_attributes, content) { return { textContent: content?.textContent || '' }; },
        },
      },
    },
    tr: {
      replaceWith(_from, _to, blocks) { return { blocks }; },
    },
  };
  const proseMirrorView = {
    state: proseMirrorState,
    focus() {},
    dispatch(transaction) {
      const text = transaction.blocks.map((block) => block.textContent).join('\n');
      proseMirrorEditable._text = text;
      proseMirrorState.doc.textContent = text;
    },
  };
  proseMirrorEditable.pmViewDesc = { view: proseMirrorView };
  api.setContentEditableText(proseMirrorEditable, '继续处理工程');
  assert.equal(proseMirrorEditable.innerText, '继续处理工程',
    'the ProseMirror transaction must update the editor state used by ChatGPT');

  // A draft changes the composer submit control to Send while a separate
  // dedicated Stop control still shows that the assistant is generating.
  submitButton.attributes['aria-label'] = '发送提示词';
  submitButton.attributes['data-testid'] = 'send-button';
  dedicatedStopButton = new FakeButton(composer);
  dedicatedStopButton.attributes['aria-label'] = '停止回答';
  dedicatedStopButton.attributes['data-testid'] = 'stop-button';
  assert.equal(api.isGenerating(), true, 'typing must not hide a separate Stop control');
  const activeDedicatedStopButton = dedicatedStopButton;
  dedicatedStopButton = null;
  staleStopButton = new FakeButton(composer);
  staleStopButton.attributes['aria-label'] = '停止回答';
  staleStopButton.attributes['data-testid'] = 'stop-button';
  submitButton.attributes['aria-label'] = '开始语音模式';
  submitButton.attributes['data-testid'] = 'composer-speech-button';
  assert.equal(api.isGenerating(), false,
    'a stale page-level Stop button must be ignored when the current composer shows the idle voice control');
  staleStopButton = null;
  dedicatedStopButton = activeDedicatedStopButton;
  submitButton.attributes['aria-label'] = '发送提示词';
  submitButton.attributes['data-testid'] = 'send-button';

  let assistantNodes = [];
  let threadUsers = [];
  let toolCardText = '';
  let laterStaticNode = null;
  const toolCard = {
    tagName: 'DIV',
    innerText: '',
    textContent: '',
    getAttribute(name) { return name === 'data-testid' ? 'tool-result' : null; },
  };
  const container = {
    innerText: '',
    textContent: '',
    getAttribute(name) { return name === 'data-testid' ? 'conversation-turn-9' : null; },
    querySelector(selector) {
      return selector === '[data-message-author-role="assistant"]' ? assistantNodes[0] : null;
    },
    querySelectorAll(selector) {
      if (selector === '[data-message-author-role="assistant"]') return assistantNodes;
      if (selector.includes('[data-testid*="tool"]')) return toolCardText ? [toolCard] : [];
      return [];
    },
    matches(selector) { return selector === '[data-testid^="conversation-turn-"]'; },
    contains(element) { return element === assistant || element === laterStaticNode; },
  };
  const assistant = {
    innerText: '最初的回答',
    textContent: '最初的回答',
    contains() { return false; },
    closest(selector) { return selector === '[data-testid^="conversation-turn-"]' ? container : null; },
    matches() { return false; },
    getAttribute() { return null; },
  };
  assistantNodes = [assistant];
  let numberedTurnNodes = [container];
  threadRoot = {
    querySelectorAll(selector) {
      if (selector === '[data-message-author-role="assistant"]') return assistantNodes;
      if (selector === '[data-message-author-role="user"]') return threadUsers;
      if (selector === '[data-testid^="conversation-turn-"]') return numberedTurnNodes;
      return [];
    },
  };
  assert.equal(api.getLatestTurn(), container, 'the entire latest assistant turn must be the progress target');
  const oldBranch = {
    ...container,
    getAttribute(name) { return name === 'data-testid' ? 'conversation-turn-99' : null; },
  };
  const oldBranchAssistant = {
    ...assistant,
    closest(selector) { return selector === '[data-testid^="conversation-turn-"]' ? oldBranch : null; },
  };
  oldBranch.querySelectorAll = (selector) => {
    if (selector === '[data-message-author-role="assistant"]') return [oldBranchAssistant];
    return [];
  };
  assistantNodes = [oldBranchAssistant, assistant];
  assert.equal(api.getLatestTurn(), container,
    'a higher-numbered old branch must not displace the last displayed assistant turn');
  const hiddenOldBranchAssistant = {
    ...oldBranchAssistant,
    closest(selector) {
      if (selector === '[hidden], [aria-hidden="true"]') return { hidden: true };
      return selector === '[data-testid^="conversation-turn-"]' ? oldBranch : null;
    },
  };
  assistantNodes = [assistant, hiddenOldBranchAssistant];
  assert.equal(api.getLatestTurn(), container,
    'a hidden stale branch at the end of the DOM must not become the response target');
  let toolOnlyTurn = null;
  const newerUser = {
    innerText: '完成后的新问题',
    textContent: '完成后的新问题',
    contains() { return false; },
    closest() { return null; },
    compareDocumentPosition(candidate) {
      if (candidate === toolOnlyTurn) return 4;
      return candidate === assistant || candidate === container ? 2 : 0;
    },
  };
  threadUsers = [newerUser];
  assistantNodes = [assistant];
  assert.equal(api.getLatestTurn(), null,
    'an old assistant reply before the latest user message must not become the current response');
  toolOnlyTurn = {
    ...container,
    innerText: 'Devspace 工具正在运行',
    textContent: 'Devspace 工具正在运行',
    getAttribute(name) { return name === 'data-testid' ? 'conversation-turn-10' : null; },
    querySelector() { return null; },
    querySelectorAll() { return []; },
    contains() { return false; },
  };
  numberedTurnNodes = [container, toolOnlyTurn];
  assert.equal(api.getLatestTurn(), toolOnlyTurn,
    'a rendered tool-only turn after the latest user must be a valid response region');
  numberedTurnNodes = [container];
  threadUsers = [];
  assistantNodes = [assistant];

  laterStaticNode = {
    ...assistant,
    innerText: '静态工具说明',
    textContent: '静态工具说明',
  };
  assistantNodes = [assistant, laterStaticNode];
  const beforeEarlierNodeStreams = api.getResponseSnapshot();
  assistant.innerText = '最初的回答，有新信息';
  const afterEarlierNodeStreams = api.getResponseSnapshot();
  assert.equal(api.isMeaningfulSnapshotChange(beforeEarlierNodeStreams, afterEarlierNodeStreams), true,
    'streaming in an earlier assistant node must count even when a later node is static');
  const beforeDomOnlyUpdate = api.getResponseSnapshot();
  assistant.textContent = '最初的回答，有新信息，以及仅在 DOM 文本中出现的新增内容';
  const afterDomOnlyUpdate = api.getResponseSnapshot();
  assert.equal(api.isMeaningfulSnapshotChange(beforeDomOnlyUpdate, afterDomOnlyUpdate), true,
    'textContent-only streaming must count when innerText lags');
  toolCardText = '工具开始';
  toolCard.innerText = toolCardText;
  const beforeToolUpdate = api.getResponseSnapshot();
  toolCardText = '工具的新结果';
  toolCard.innerText = toolCardText;
  const afterToolUpdate = api.getResponseSnapshot();
  assert.equal(api.isMeaningfulSnapshotChange(beforeToolUpdate, afterToolUpdate), true,
    'tool-card output outside assistant text must refresh progress');
  const beforeTurnTextUpdate = api.getResponseSnapshot();
  container.textContent = '工具卡片之外的新结果';
  const afterTurnTextUpdate = api.getResponseSnapshot();
  assert.equal(api.isMeaningfulSnapshotChange(beforeTurnTextUpdate, afterTurnTextUpdate), true,
    'new text elsewhere in the assistant turn must refresh progress');
  assistantNodes = [assistant];
  toolCardText = '';
  container.textContent = '';

  const observers = [];
  context.MutationObserver = class MutationObserver {
    constructor(callback) { this.callback = callback; observers.push(this); }
    observe() {}
    disconnect() {}
  };
  let fakeNow = 1_000_000;
  context.Date = class extends Date { static now() { return fakeNow; } };
  api.start();
  assert.equal(api.state.generationActive, true);
  fakeNow += 20_000;
  composer.value = '用户正在输入新问题';
  assistant.innerText = '最初的回答，加上正常生成的新信息';
  await api.tick();
  assert.equal(api.state.generationActive, true, 'composer typing must not stop generation');
  assert.equal(api.state.lastContentActivityAt, fakeNow, 'new assistant text must refresh the generation timer');

  dedicatedStopButton = null;
  fakeNow += 1_000;
  await api.tick();
  assert.equal(api.state.generationActive, true, 'a brief missing Stop control while typing must not count as stopped');
  dedicatedStopButton = new FakeButton(composer);
  dedicatedStopButton.attributes['aria-label'] = '停止回答';
  dedicatedStopButton.attributes['data-testid'] = 'stop-button';
  fakeNow += 1_000;
  await api.tick();
  assert.equal(api.state.stopMissingSince, 0, 'a returning Stop control clears the pending stop');

  dedicatedStopButton = null;
  fakeNow += 4_000;
  const stoppedAt = fakeNow;
  await api.tick();
  fakeNow += 3_000;
  assistant.innerText = '最初的回答，加上正常生成的新信息，随后重绘的旧信息';
  await api.tick();
  assert.equal(api.state.generationActive, false, 'a confirmed missing Stop control means the answer stopped');
  assert.equal(api.state.idleSince, stoppedAt, 'idle timing starts when Stop first disappeared');
  const contentTimeAtStop = api.state.lastContentActivityAt;
  fakeNow += 1_000;
  assistant.innerText += '，再次重绘';
  observers[observers.length - 1].callback([]);
  await api.tick();
  assert.equal(api.state.idleSince, stoppedAt, 'late old output must not reset the stopped-answer timer');
  assert.equal(api.state.lastContentActivityAt, contentTimeAtStop, 'late old output must not become generation progress');

  // Some layouts never expose a Stop control. Streaming text itself must then
  // establish and refresh the active-response timer.
  assistant.innerText = '无停止按钮时的初始回答';
  composer.value = '';
  fakeNow += 1_000;
  api.start();
  assert.equal(api.state.stopWasSeenInCurrentResponse, false);
  fakeNow += 20_000;
  assistant.innerText += '，第一段新内容';
  await api.tick();
  assert.equal(api.state.generationActive, true, 'new assistant content infers generation without Stop');
  assert.equal(api.state.lastContentActivityAt, fakeNow);
  fakeNow += 30_000;
  assistant.innerText += '，第二段新内容';
  await api.tick();
  assert.equal(api.state.lastContentActivityAt, fakeNow, 'subsequent content refreshes the timer without Stop');
  const lastContentAt = fakeNow;
  api.state.nextActionNotBefore = fakeNow + 1_000_000;
  fakeNow += 61_000;
  await api.tick();
  assert.equal(api.state.generationActive, false, 'no new content eventually ends inferred generation');
  assert.equal(api.state.idleSince, lastContentAt, 'idle timer follows the last actual content update');

  fakeNow += 1_000;
  assistant.innerText = '新的流式回答';
  api.start();
  fakeNow += 400;
  assistant.innerText += ' 第一批';
  observers[observers.length - 1].callback([]);
  assert.equal(api.state.lastContentActivityAt, fakeNow,
    'the observer records a streamed chunk before the next one-second poll');
  fakeNow += 400;
  assistant.innerText += ' 第二批';
  observers[observers.length - 1].callback([]);
  assert.equal(api.state.lastContentActivityAt, fakeNow,
    'each distinct streamed chunk refreshes the activity time');
  fakeNow += 400;
  observers[observers.length - 1].callback([{
    type: 'childList',
    target: assistant,
    addedNodes: [{ nodeType: 3, parentElement: assistant, textContent: ' React 替换节点吐出的新内容' }],
    removedNodes: [],
  }]);
  assert.equal(api.state.lastContentActivityAt, fakeNow,
    'React child-node streaming must refresh activity between polls');
  fakeNow += 200;
  await api.tick();
  assert.equal(api.state.generationActive, true,
    'the next poll consumes throughput observed between polls');

  // The text node can change before innerText/textContent of the turn is
  // reflected in a snapshot. Its old/new payload is still real throughput.
  assistant.innerText = '稳定快照';
  assistant.textContent = '稳定快照';
  fakeNow += 1_000;
  api.start();
  const streamNode = { nodeType: 3, parentElement: assistant, data: '稳定快照，新吐出的字' };
  fakeNow += 200;
  observers[observers.length - 1].callback([{
    type: 'characterData', target: streamNode, oldValue: '稳定快照',
  }]);
  assert.equal(api.state.lastContentActivityAt, fakeNow,
    'a text-node stream must refresh activity even while the turn snapshot lags');
  const afterStream = fakeNow;
  fakeNow += 200;
  streamNode.data = '思考中 2s';
  observers[observers.length - 1].callback([{
    type: 'characterData', target: streamNode, oldValue: '思考中 1s',
  }]);
  assert.equal(api.state.lastContentActivityAt, afterStream,
    'an elapsed-time label must not look like assistant throughput');
  fakeNow += 200;
  observers[observers.length - 1].callback([{
    type: 'characterData', target: { nodeType: 3, parentElement: composer, data: '输入中' },
    oldValue: '',
  }]);
  assert.equal(api.state.lastContentActivityAt, afterStream,
    'typing in the composer must not refresh assistant activity');
  fakeNow += 200;
  api.state.stopWasSeenInCurrentResponse = true;
  streamNode.data = '停止后重绘的旧字';
  observers[observers.length - 1].callback([{
    type: 'characterData', target: streamNode, oldValue: '稳定快照',
  }]);
  assert.equal(api.state.lastContentActivityAt, afterStream,
    'missing Stop control has priority over later DOM text mutations');

  // A regenerated branch may use a lower conversation-turn number. A newly
  // submitted user turn authorizes this one branch switch.
  api.state.stopWasSeenInCurrentResponse = false;
  assistantNodes = [oldBranchAssistant];
  api.resetResponseTracking(fakeNow, true);
  assert.equal(api.state.lastResponseSnapshot.key, 'conversation-turn-99');
  assistantNodes = [assistant];
  fakeNow += 200;
  observers[observers.length - 1].callback([]);
  assert.equal(api.state.lastResponseSnapshot.key, 'conversation-turn-9',
    'the active displayed branch must replace a higher-numbered old branch');
  assert.equal(api.state.allowBranchSwitch, false);
  assert.equal(api.state.lastContentActivityAt, fakeNow,
    'new content on a lower-numbered branch must refresh activity');

  // Devspace/terminal tool cards can be siblings of assistant messages rather
  // than descendants of the latest assistant turn.
  const devspaceBlock = {
    nodeType: 1,
    textContent: '已调用工具 Devspace Ran command 81 lines',
    parentElement: threadRoot,
    closest() { return null; },
    contains() { return false; },
  };
  const toolRequestUser = {
    innerText: '请继续处理工程',
    textContent: '请继续处理工程',
    contains() { return false; },
    closest() { return null; },
    compareDocumentPosition(candidate) {
      if (candidate === assistant || candidate === container) return 2;
      if (candidate === devspaceBlock) return 4;
      return 0;
    },
  };
  threadUsers = [toolRequestUser];
  assistantNodes = [assistant];
  fakeNow += 1_000;
  api.start();
  assert.equal(api.state.lastResponseSnapshot, null,
    'the prior assistant reply must not stand in for a tool-only current response');
  fakeNow += 250;
  observers[observers.length - 1].callback([{
    type: 'childList', target: threadRoot, addedNodes: [devspaceBlock], removedNodes: [],
  }]);
  assert.equal(api.state.lastContentActivityAt, fakeNow,
    'a Devspace block outside the assistant turn must refresh activity');
  assert.equal(api.state.lastProgressSource, '工具/新增节点');
  const devspaceText = {
    nodeType: 3,
    parentElement: devspaceBlock,
    data: 'Ran command 82 lines',
  };
  fakeNow += 250;
  observers[observers.length - 1].callback([{
    type: 'characterData', target: devspaceText, oldValue: 'Ran command 81 lines',
  }]);
  assert.equal(api.state.lastContentActivityAt, fakeNow,
    'a running Devspace card changing its output count must refresh activity');
  assert.equal(api.state.lastProgressSource, '工具/文本节点');
  await api.tick();
  assert.equal(api.state.generationActive, true,
    'tool-only DOM activity must establish an active response without an assistant snapshot');
  api.state.nextActionNotBefore = fakeNow + 1_000_000;
  fakeNow += 61_000;
  await api.tick();
  assert.equal(api.state.generationActive, false,
    'tool-only activity must be allowed to finish its quiet timer without an assistant snapshot');
  assert.doesNotMatch(api.state.status, /尚未定位到助手回复区域/u,
    'previously observed tool activity must not be rejected by the old assistant-region guard');
  threadUsers = [];

  // Reopening supervision after an already handled completion marker must not
  // instantly pause on that old marker again.
  assistant.innerText = '[TASK_DONE]';
  assistant.textContent = '[TASK_DONE]';
  fakeNow += 1_000;
  api.start();
  fakeNow += 1_000;
  await api.tick();
  assert.equal(api.state.enabled, true,
    'an existing completion marker is only a baseline when supervision is reopened');
  assistant.innerText = '新任务仍在进行';
  assistant.textContent = '新任务仍在进行';
  fakeNow += 1_000;
  await api.tick();
  assistant.innerText = '[TASK_DONE]';
  assistant.textContent = '[TASK_DONE]';
  fakeNow += 1_000;
  await api.tick();
  assert.equal(api.state.enabled, false,
    'a newly produced completion marker still pauses supervision');

  api.state.enabled = true;
  api.state.continueCount = 3;
  assert.equal(api.scheduleSendRecoveryReload({
    expectedUserText: '继续。测试刷新恢复',
    userSignatureBefore: '2:10:abc',
    expectedMessageCountBefore: 1,
  }), true, 'a stuck send should schedule one reload recovery');
  const savedRecovery = api.readSendRecovery();
  assert.equal(savedRecovery.supervisionEnabled, true,
    'reload recovery must remember that supervision was running');
  assert.equal(savedRecovery.continueCount, 3,
    'reload recovery must preserve the continuation count');
  assert.equal(savedRecovery.expectedUserText, '继续。测试刷新恢复');
  await new Promise((resolve) => setTimeout(resolve, 10));
  assert.equal(reloadCalls, 1, 'a stuck send should refresh the page exactly once');
  api.clearSendRecovery();
  assert.equal(api.readSendRecovery(), null, 'the reload marker must be cleared after recovery');

  console.log('chatgpt-auto-continue: all checks passed');
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
