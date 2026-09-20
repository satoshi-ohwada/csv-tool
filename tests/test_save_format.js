const fs = require('fs');
const path = require('path');
const vm = require('vm');

function createDummyEl(tagName) {
  const listeners = {};
  const children = [];
  const classList = new Set();
  const style = {};
  const el = {
    tagName: tagName.toUpperCase(),
    id: '',
    value: '',
    textContent: '',
    innerHTML: '',
    style,
    classList: {
      add: (c) => classList.add(c),
      remove: (c) => classList.delete(c),
      contains: (c) => classList.has(c)
    },
    addEventListener(type, fn) {
      if (!listeners[type]) listeners[type] = [];
      listeners[type].push(fn);
    },
    removeEventListener(type, fn) {
      if (!listeners[type]) return;
      listeners[type] = listeners[type].filter(f => f !== fn);
    },
    dispatchEvent(event) {
      const fns = listeners[event.type] || [];
      fns.forEach(fn => fn(event));
    },
    appendChild(child) { children.push(child); return child; },
    removeChild(child) {
      const idx = children.indexOf(child);
      if (idx !== -1) children.splice(idx, 1);
      return child;
    },
    setAttribute(k, v) { el[k] = v; },
    getAttribute(k) { return el[k] || null; },
    scrollIntoView() {},
    focus() {},
    blur() {},
    click() {
      if (listeners['click']) {
        listeners['click'].forEach(fn => fn({ type: 'click', target: el }));
      }
    }
  };
  return el;
}

const mockDoc = {
  addEventListener() {},
  removeEventListener() {},
  getElementById(id) {
    if (!this._elements) this._elements = {};
    if (!this._elements[id]) {
      this._elements[id] = createDummyEl('div');
      this._elements[id].id = id;
    }
    return this._elements[id];
  },
  querySelector(sel) { return createDummyEl('div'); },
  querySelectorAll() { return []; },
  createElement(tag) { return createDummyEl(tag); },
  body: createDummyEl('body'),
  documentElement: createDummyEl('html')
};

const sandbox = {
  console,
  addEventListener() {},
  removeEventListener() {},
  document: mockDoc,
  navigator: { userAgent: 'NodeTest' },
  localStorage: { getItem: () => null, setItem: () => {} },
  setTimeout,
  clearTimeout,
  setInterval,
  clearInterval,
  Tabulator: function() {
    this.setData = () => Promise.resolve();
    this.getData = () => [{ 'col_1': '100', 'col_2': '東京' }];
    this.getColumns = () => [];
    this.getRows = () => [];
    this.setFilter = () => {};
    this.clearFilter = () => {};
    this.setSort = () => {};
    this.clearSort = () => {};
    this.getSorters = () => [];
  },
  Papa: {
    unparse: (rows, opts) => JSON.stringify(rows),
    parse: () => ({ data: [] })
  },
  Encoding: {
    convert: (arr) => arr,
    stringToCode: (s) => Array.from(s).map(c => c.charCodeAt(0)),
    codeToString: (arr) => String.fromCharCode(...arr),
    detect: () => 'UTF8'
  },
  Blob: class {
    constructor(parts, opts) { this.parts = parts; this.type = opts ? opts.type : ''; }
  }
};
sandbox.window = sandbox;
sandbox.global = sandbox;

const appJsPath = path.resolve(__dirname, "../app.js");
const appJsCode = fs.readFileSync(appJsPath, 'utf8');

vm.createContext(sandbox);
vm.runInContext(appJsCode, sandbox);

let passed = 0;
let failed = 0;
function assert(desc, cond) {
  if (cond) {
    passed++;
    console.log(`  ✓ ${desc}`);
  } else {
    failed++;
    console.error(`  ✗ ${desc}`);
  }
}

console.log("\n========================================");
console.log("TEST SUITE: 保存ダイアログ・保存フォーマット選択");
console.log("========================================");

const {
  tabs,
  showSaveModal,
  hideSaveModal,
  handleSave,
  applySaveSettingsToCurrentTab,
  generateExportData
} = sandbox;

const saveModal = mockDoc.getElementById('save-modal');
const modalFilename = mockDoc.getElementById('modal-filename');
const selectEncoding = mockDoc.getElementById('select-encoding');
const selectDelimiter = mockDoc.getElementById('select-delimiter');
const selectNewline = mockDoc.getElementById('select-newline');
const selectCommaMode = mockDoc.getElementById('select-comma-mode');
const existingActions = mockDoc.getElementById('save-modal-existing-actions');
const newActions = mockDoc.getElementById('save-modal-new-actions');

// テスト用タブの設定を VM 内で実行
vm.runInContext(`
  tabs = [{
    id: 'tab_test_1',
    name: 'test.csv',
    isNew: false,
    isModified: true,
    encoding: 'Shift_JIS',
    delimiter: '\\t',
    newline: '\\n',
    commaMode: 'remove',
    columns: [{ field: 'col_1', title: '数値' }, { field: 'col_2', title: '地域' }],
    data: [{ 'col_1': '1,234', 'col_2': '大阪' }]
  }];
  activeTabId = 'tab_test_1';
  currentTable = new Tabulator();
`, sandbox);

const activeTab = vm.runInContext('tabs[0]', sandbox);

// 1. showSaveModal でタブの設定がモーダルのセレクトに正しく反映されるか
showSaveModal(activeTab);
assert("showSaveModal: モーダルが表示される (display === 'flex')", saveModal.style.display === 'flex');
assert("showSaveModal: ファイル名が表示される", modalFilename.textContent === 'test.csv');
assert("showSaveModal: 文字コードが反映される", selectEncoding.value === 'Shift_JIS');
assert("showSaveModal: 区切り文字が反映される", selectDelimiter.value === '\t');
assert("showSaveModal: 改行コードが反映される", selectNewline.value === '\n');
assert("showSaveModal: 既存ファイルの場合は上書き/別名ボタンが表示される", existingActions.style.display === 'flex' && newActions.style.display === 'none');

// 2. 新規ファイルの場合の表示切り替え
const newTab = { ...activeTab, isNew: true, name: '無題-1.csv' };
showSaveModal(newTab);
assert("showSaveModal: 新規ファイルの場合は名前を付けて保存ボタンが表示される", newActions.style.display === 'flex' && existingActions.style.display === 'none');

// 3. モーダルでフォーマットを変更して適用
selectEncoding.value = 'UTF-8';
selectDelimiter.value = ',';
selectNewline.value = '\r\n';
selectCommaMode.value = 'keep';
applySaveSettingsToCurrentTab();

assert("applySaveSettingsToCurrentTab: tab.encoding が UTF-8 に更新される", activeTab.encoding === 'UTF-8');
assert("applySaveSettingsToCurrentTab: tab.delimiter が カンマ に更新される", activeTab.delimiter === ',');
assert("applySaveSettingsToCurrentTab: tab.newline が CRLF に更新される", activeTab.newline === '\r\n');
assert("applySaveSettingsToCurrentTab: tab.commaMode が keep に更新される", activeTab.commaMode === 'keep');

// 4. hideSaveModal でモーダルが閉じる
hideSaveModal();
assert("hideSaveModal: モーダルが非表示になる (display === 'none')", saveModal.style.display === 'none');

// 5. handleSave() の呼び出しでモーダルが開くか
handleSave();
assert("handleSave: 保存ボタン押下でモーダルが開く", saveModal.style.display === 'flex');

console.log("\n========================================");
console.log(`TEST SUMMARY: Passed: ${passed}, Failed: ${failed}`);
console.log("========================================");

if (failed > 0) process.exit(1);
process.exit(0);
