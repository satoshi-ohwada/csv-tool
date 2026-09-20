const fs = require('fs');
const vm = require('vm');
const path = require('path');
const assert = require('assert');

// 1. 環境モックの構築
function createDummyEl(tagName = 'div') {
  const children = [];
  const listeners = {};
  const el = {
    tagName: tagName.toUpperCase(),
    value: '',
    textContent: '',
    innerHTML: '',
    checked: false,
    style: {},
    classList: {
      _classes: new Set(),
      add(...cls) { cls.forEach(c => this._classes.add(c)); },
      remove(...cls) { cls.forEach(c => this._classes.delete(c)); },
      contains(c) { return this._classes.has(c); },
      toggle(c) { if (this._classes.has(c)) this._classes.delete(c); else this._classes.add(c); }
    },
    addEventListener(type, fn) {
      if (!listeners[type]) listeners[type] = [];
      listeners[type].push(fn);
    },
    removeEventListener(type, fn) {
      if (listeners[type]) listeners[type] = listeners[type].filter(f => f !== fn);
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
    querySelector() { return createDummyEl('div'); },
    querySelectorAll() { return []; },
    scrollIntoView() {},
    focus() {},
    blur() {}
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
  querySelector() { return createDummyEl('div'); },
  querySelectorAll() { return []; },
  createElement(tag) { return createDummyEl(tag); },
  body: createDummyEl('body'),
  documentElement: createDummyEl('html')
};

function MockTabulator(container, options = {}) {
  this.options = options;
  this.data = options.data || [];
  this.columns = options.columns || [];
  this.destroy = () => {};
  this.getData = () => this.data;
  this.getRows = () => this.data.map((d, i) => ({
    getData: () => d,
    update: (nd) => { Object.assign(d, nd); },
    getCell: () => ({
      getElement: () => createDummyEl('div'),
      getValue: () => '',
      setValue: () => {}
    })
  }));
  this.getColumns = () => this.columns.map(c => ({
    getField: () => c.field,
    getDefinition: () => c,
    getElement: () => createDummyEl('div')
  }));
  this.getRanges = () => [];
  this.on = (evt, fn) => {};
  this.getSorters = () => [];
  this.clearSort = () => {};
  this.setSort = () => {};
}

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
  Tabulator: MockTabulator
};
sandbox.window = sandbox;
sandbox.global = sandbox;

// 2. app.js のロード
const appJsPath = path.resolve(__dirname, "../app.js");
const appJsCode = fs.readFileSync(appJsPath, 'utf8');

vm.createContext(sandbox);
vm.runInContext(appJsCode, sandbox);

let passed = 0;
let failed = 0;

function it(desc, fn) {
  try {
    fn();
    console.log(`  ✓ ${desc}`);
    passed++;
  } catch (err) {
    console.error(`  ✗ ${desc}`);
    console.error(`    ${err.message}`);
    failed++;
  }
}

console.log('========================================');
console.log('TEST SUITE: データ整形 - セル内改行削除 (remove_newlines)');
console.log('========================================\n');

// テスト1: removeNewlines 単体テスト
it('removeNewlines: CRLF, LF, CR の各種改行コードをすべて削除する', () => {
  const { removeNewlines } = sandbox;
  assert.strictEqual(removeNewlines('東京都\r\n港区\n芝公園\r1-1'), '東京都港区芝公園1-1');
  assert.strictEqual(removeNewlines('改行なし文字列'), '改行なし文字列');
  assert.strictEqual(removeNewlines('\n\n先頭と末尾\r\n'), '先頭と末尾');
  assert.strictEqual(removeNewlines(''), '');
});

// テスト2: applyCleanTransformation('remove_newlines')
it('applyCleanTransformation: remove_newlines アクションの変換', () => {
  const { applyCleanTransformation } = sandbox;
  assert.strictEqual(applyCleanTransformation('Line1\nLine2\r\nLine3', 'remove_newlines'), 'Line1Line2Line3');
});

// テスト3: テーブル全体のセル内改行削除の実行
it('executeDataClean("remove_newlines"): 表全体のセルから改行が削除され、isModified が true になる', () => {
  vm.runInContext(`
    (() => {
      const hdrs = ['住所', '備考'];
      const cols = buildTabulatorColumns(hdrs);
      const rows = [
        { _id: 1, col_0: '東京都\\n千代田区', col_1: '特記事項なし' },
        { _id: 2, col_0: '大阪府\\r\\n大阪市', col_1: '1行目\\r2行目\\n3行目' }
      ];
      const t = {
        id: 'tab-clean-1',
        name: 'address.csv',
        headers: hdrs,
        columns: cols,
        data: rows,
        isModified: false
      };
      tabs = [t];
      activeTabId = 'tab-clean-1';
      currentTable = new Tabulator(null, { data: t.data, columns: t.columns });
      executeDataClean('remove_newlines');
    })()
  `, sandbox);

  const tab = vm.runInContext(`tabs.find(t => t.id === activeTabId)`, sandbox);
  assert.strictEqual(tab.isModified, true);
  assert.strictEqual(tab.data[0].col_0, '東京都千代田区');
  assert.strictEqual(tab.data[0].col_1, '特記事項なし');
  assert.strictEqual(tab.data[1].col_0, '大阪府大阪市');
  assert.strictEqual(tab.data[1].col_1, '1行目2行目3行目');
});

// テスト4: 改行がない場合は変更なし（通知・フラグ管理）
it('executeDataClean("remove_newlines"): 改行がないセルのみの場合は変更されない', () => {
  vm.runInContext(`
    (() => {
      const hdrs = ['名前'];
      const cols = buildTabulatorColumns(hdrs);
      const rows = [
        { _id: 1, col_0: '山田太郎' }
      ];
      const t = {
        id: 'tab-clean-2',
        name: 'names.csv',
        headers: hdrs,
        columns: cols,
        data: rows,
        isModified: false
      };
      tabs = [t];
      activeTabId = 'tab-clean-2';
      currentTable = new Tabulator(null, { data: t.data, columns: t.columns });
      executeDataClean('remove_newlines');
    })()
  `, sandbox);

  const tab = vm.runInContext(`tabs.find(t => t.id === activeTabId)`, sandbox);
  assert.strictEqual(tab.isModified, false);
  assert.strictEqual(tab.data[0].col_0, '山田太郎');
});

console.log('\n========================================');
console.log(`TEST SUMMARY: Passed: ${passed}, Failed: ${failed}`);
console.log('========================================');

if (failed > 0) {
  process.exit(1);
} else {
  console.log('ALL CLEAN TESTS PASSED! 🎉\n');
  process.exit(0);
}
