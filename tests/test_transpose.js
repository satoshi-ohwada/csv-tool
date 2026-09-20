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
  this.on = (evt, fn) => {
    if (evt === 'tableBuilt') {
      setTimeout(fn, 0);
    }
  };
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

// テストランナー
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
console.log('TEST SUITE: 行と列の入れ替え (転置: Transpose)');
console.log('========================================\n');

// ヘルパー: VM内でテーブルをセットアップ
function setupTable(tabId, headers, rowData) {
  return vm.runInContext(`
    (() => {
      const hdrs = ${JSON.stringify(headers)};
      const cols = buildTabulatorColumns(hdrs);
      const rows = ${JSON.stringify(rowData)};
      const t = {
        id: '${tabId}',
        name: 'test.csv',
        headers: hdrs,
        columns: cols,
        data: rows,
        isModified: false
      };
      tabs = [t];
      activeTabId = '${tabId}';
      currentTable = new Tabulator(null, { data: t.data, columns: t.columns });
      return t;
    })()
  `, sandbox);
}

function getActiveTab() {
  return vm.runInContext(`tabs.find(t => t.id === activeTabId)`, sandbox);
}

// テスト1: 基本的な長方形データの転置と往復（インボリューション性）
it('3列×2行のテーブルが転置されて3行×3列になり、再転置で元に戻る', () => {
  setupTable('tab-1', ['名前', '国語', '数学'], [
    { _id: 1, col_0: '田中', col_1: '80', col_2: '90' },
    { _id: 2, col_0: '鈴木', col_1: '70', col_2: '85' }
  ]);

  // 1回目の転置を実行
  vm.runInContext(`transposeCurrentTable()`, sandbox);
  let tab = getActiveTab();

  // 検証: 新ヘッダーは元の縦第1列 = ['名前', '田中', '鈴木']
  assert.deepStrictEqual([...tab.headers], ['名前', '田中', '鈴木']);
  assert.strictEqual(tab.data.length, 2);
  // 新行1: ['国語', '80', '70']
  assert.strictEqual(tab.data[0].col_0, '国語');
  assert.strictEqual(tab.data[0].col_1, '80');
  assert.strictEqual(tab.data[0].col_2, '70');
  // 新行2: ['数学', '90', '85']
  assert.strictEqual(tab.data[1].col_0, '数学');
  assert.strictEqual(tab.data[1].col_1, '90');
  assert.strictEqual(tab.data[1].col_2, '85');
  assert.strictEqual(tab.isModified, true);

  // 2回目の転置を実行（元に戻ることを検証）
  vm.runInContext(`transposeCurrentTable()`, sandbox);
  tab = getActiveTab();

  assert.deepStrictEqual([...tab.headers], ['名前', '国語', '数学']);
  assert.strictEqual(tab.data.length, 2);
  assert.strictEqual(tab.data[0].col_0, '田中');
  assert.strictEqual(tab.data[0].col_1, '80');
  assert.strictEqual(tab.data[0].col_2, '90');
  assert.strictEqual(tab.data[1].col_0, '鈴木');
  assert.strictEqual(tab.data[1].col_1, '70');
  assert.strictEqual(tab.data[1].col_2, '85');
});

// テスト2: 1列のみのテーブルの転置
it('1列×3行のテーブルを転置すると4列×1行になる', () => {
  setupTable('tab-2', ['品名'], [
    { _id: 1, col_0: 'りんご' },
    { _id: 2, col_0: 'みかん' },
    { _id: 3, col_0: 'バナナ' }
  ]);

  vm.runInContext(`transposeCurrentTable()`, sandbox);
  const tab = getActiveTab();

  // 元は (ヘッダー1行 + データ3行) = 4行 × 1列
  // 転置後は 1行 × 4列 (ヘッダー1行、データ行なしだが空行1行補完)
  assert.deepStrictEqual([...tab.headers], ['品名', 'りんご', 'みかん', 'バナナ']);
  assert.strictEqual(tab.data.length, 1);
  assert.strictEqual(tab.data[0].col_0, '');
  assert.strictEqual(tab.data[0].col_1, '');
  assert.strictEqual(tab.data[0].col_2, '');
  assert.strictEqual(tab.data[0].col_3, '');
});

// テスト3: 空セルや空白を含むテーブルの転置
it('空文字セルがある場合でも安全にデフォルト列名が補完され転置される', () => {
  setupTable('tab-3', ['A', ''], [
    { _id: 1, col_0: '', col_1: 'val1' },
    { _id: 2, col_0: 'val2', col_1: '' }
  ]);

  vm.runInContext(`transposeCurrentTable()`, sandbox);
  const tab = getActiveTab();

  // 新ヘッダー: 元の第1列 (A, '', 'val2') -> 2番目の空文字は '列 2' に補完
  assert.strictEqual(tab.headers[0], 'A');
  assert.strictEqual(tab.headers[1], '列 2');
  assert.strictEqual(tab.headers[2], 'val2');

  // 新行1: 元の第2列 (列 2, 'val1', '')
  assert.strictEqual(tab.data[0].col_0, '列 2');
  assert.strictEqual(tab.data[0].col_1, 'val1');
  assert.strictEqual(tab.data[0].col_2, '');
});

// テスト4: executeDataClean('transpose') 経由での呼び出し
it('executeDataClean("transpose") が正しく transposeCurrentTable をディスパッチする', () => {
  setupTable('tab-4', ['X', 'Y'], [
    { _id: 1, col_0: '10', col_1: '20' }
  ]);

  vm.runInContext(`executeDataClean('transpose')`, sandbox);
  const tab = getActiveTab();

  assert.deepStrictEqual([...tab.headers], ['X', '10']);
  assert.strictEqual(tab.data.length, 1);
  assert.strictEqual(tab.data[0].col_0, 'Y');
  assert.strictEqual(tab.data[0].col_1, '20');
});

// テスト5: 列定義 (columns) とカラム文字 (A, B, C...) の再構築
it('転置後のcolumnsプロパティに適切なcolLetterとtitleFormatterParamsが設定される', () => {
  setupTable('tab-5', ['C1', 'C2', 'C3', 'C4'], [
    { _id: 1, col_0: 'a', col_1: 'b', col_2: 'c', col_3: 'd' }
  ]);

  vm.runInContext(`transposeCurrentTable()`, sandbox);
  const tab = getActiveTab();

  // 元は 2行×4列 -> 転置後は 4行×2列 (ヘッダー2列、データ3行)
  assert.strictEqual(tab.headers.length, 2);
  // columns[0]はrownum列、columns[1]はcol_0, columns[2]はcol_1
  assert.strictEqual(tab.columns.length, 3);
  assert.strictEqual(tab.columns[1].field, 'col_0');
  assert.strictEqual(tab.columns[1].colLetter, 'A');
  assert.strictEqual(tab.columns[2].field, 'col_1');
  assert.strictEqual(tab.columns[2].colLetter, 'B');
});

console.log('\n========================================');
console.log(`TEST SUMMARY: Passed: ${passed}, Failed: ${failed}`);
console.log('========================================');

if (failed > 0) {
  process.exit(1);
} else {
  console.log('ALL TRANSPOSE TESTS PASSED! 🎉\n');
  process.exit(0);
}
