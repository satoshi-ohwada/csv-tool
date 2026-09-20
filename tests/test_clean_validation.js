const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

// DOM モック作成
function createDummyEl(tagName = 'div') {
  const el = {
    tagName: tagName.toUpperCase(),
    style: {},
    classList: {
      _classes: new Set(),
      add: function(c) { this._classes.add(c); },
      remove: function(c) { this._classes.delete(c); },
      contains: function(c) { return this._classes.has(c); },
      toggle: function(c) { if (this.contains(c)) this.remove(c); else this.add(c); }
    },
    dataset: {},
    children: [],
    appendChild: function(child) { this.children.push(child); return child; },
    addEventListener: () => {},
    removeEventListener: () => {},
    setAttribute: (k, v) => { el[k] = v; },
    getAttribute: (k) => el[k],
    closest: (sel) => null,
    querySelector: (sel) => null,
    querySelectorAll: (sel) => [],
    innerHTML: '',
    textContent: ''
  };
  return el;
}

const elMap = {};
const mockDoc = {
  readyState: 'complete',
  createElement: (tag) => createDummyEl(tag),
  getElementById: (id) => {
    if (!elMap[id]) {
      const el = createDummyEl('div');
      el.id = id;
      elMap[id] = el;
    }
    return elMap[id];
  },
  querySelector: () => createDummyEl('div'),
  querySelectorAll: () => [],
  addEventListener: () => {},
  removeEventListener: () => {}
};

function MockTabulator(container, options) {
  this.options = options || {};
  this.data = (options && options.data) ? JSON.parse(JSON.stringify(options.data)) : [];
  this.columns = (options && options.columns) ? options.columns : [];

  this.getData = () => this.data;
  this.setData = (newData) => { this.data = newData; return Promise.resolve(); };
  this.getRows = () => this.data.map((row, idx) => ({
    getData: () => row,
    update: (fields) => Object.assign(row, fields),
    getCell: (field) => ({
      getElement: () => createDummyEl('div'),
      getValue: () => row[field],
      setValue: (val) => { row[field] = val; }
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
  Tabulator: MockTabulator,
  Papa: {
    unparse: (rows) => JSON.stringify(rows),
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

console.log('\n========================================');
console.log('TEST SUITE: データクレンジング - 重複チェック＆列診断サマリー');
console.log('========================================\n');

// テスト1: 重複値の検出ロジック (highlightDuplicatesInColumn)
it('highlightDuplicatesInColumn: 日時列の重複値（同一時刻が複数行）を検出する', () => {
  vm.runInContext(`
    (() => {
      const hdrs = ['日時', '値'];
      const cols = buildTabulatorColumns(hdrs);
      const rows = [
        { _id: 1, col_0: '2026-09-01 10:00:00', col_1: '10' },
        { _id: 2, col_0: '2026-09-01 11:00:00', col_1: '20' },
        { _id: 3, col_0: '2026-09-01 10:00:00', col_1: '30' }, // 重複
        { _id: 4, col_0: '2026-09-01 12:00:00', col_1: '40' }
      ];
      tabs = [{
        id: 'tab_dup_test',
        name: 'timeseries.csv',
        headers: hdrs,
        columns: cols,
        data: rows,
        isModified: false
      }];
      activeTabId = 'tab_dup_test';
      currentTable = new Tabulator(null, { data: rows, columns: cols });
      highlightDuplicatesInColumn('col_0');
    })();
  `, sandbox);

  // highlightedColumnField が col_0 にセットされていること
  const hCol = vm.runInContext('highlightedColumnField', sandbox);
  assert.strictEqual(hCol, 'col_0');
});

// テスト2: 列データ診断サマリーモーダル生成 (showColumnSummaryModal)
it('showColumnSummaryModal: 欠測率、秘匿x、ゼロ、数値範囲、重複状況が正しく集計される', () => {
  vm.runInContext(`
    (() => {
      const hdrs = ['タイムスタンプ', '売上', '備考'];
      const cols = buildTabulatorColumns(hdrs);
      const rows = [
        { _id: 1, col_0: '2026-09-01', col_1: '100', col_2: '正常' },
        { _id: 2, col_0: '2026-09-02', col_1: '0', col_2: 'ゼロ' },
        { _id: 3, col_0: '2026-09-03', col_1: 'NA', col_2: '欠測' },
        { _id: 4, col_0: '2026-09-04', col_1: 'x', col_2: '秘匿' },
        { _id: 5, col_0: '2026-09-01', col_1: '250', col_2: '' } // col_0重複, col_2空セル
      ];
      tabs = [{
        id: 'tab_summary_test',
        name: 'summary.csv',
        headers: hdrs,
        columns: cols,
        data: rows,
        isModified: false
      }];
      activeTabId = 'tab_summary_test';
      currentTable = new Tabulator(null, { data: rows, columns: cols });
      showColumnSummaryModal();
    })();
  `, sandbox);

  const tbody = sandbox.document.getElementById('summary-table-tbody');
  assert.ok(tbody.children.length === 3, '3列分のサマリー行が生成されている');

  // 列1: タイムスタンプ
  assert.ok(tbody.children[0].innerHTML.includes('タイムスタンプ'));
  assert.ok(tbody.children[0].innerHTML.includes('日付/日時'));
  assert.ok(tbody.children[0].innerHTML.includes('1 種類の重複')); // 2026-09-01 が重複

  // 列2: 売上
  assert.ok(tbody.children[1].innerHTML.includes('売上'));
  assert.ok(tbody.children[1].innerHTML.includes('数値'));
  assert.ok(tbody.children[1].innerHTML.includes('1 件 (20.0%)')); // NAが1件 (欠測20%)
  assert.ok(tbody.children[1].innerHTML.includes('1 件')); // 秘匿xが1件
  assert.ok(tbody.children[1].innerHTML.includes('0 〜 250')); // 最小〜最大
});

console.log(`\n========================================`);
console.log(`TEST SUMMARY: Passed: ${passed}, Failed: ${failed}`);
console.log('========================================');

if (failed > 0) process.exit(1);
process.exit(0);
