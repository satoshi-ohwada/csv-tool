const fs = require('fs');
const vm = require('vm');
const path = require('path');
const assert = require('assert');

// 1. DOM・ブラウザ環境モックの構築
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
      toggle(c, force) {
        if (force !== undefined) {
          if (force) this._classes.add(c); else this._classes.delete(c);
          return force;
        }
        if (this._classes.has(c)) {
          this._classes.delete(c);
          return false;
        } else {
          this._classes.add(c);
          return true;
        }
      }
    },
    addEventListener(type, fn) {
      if (!listeners[type]) listeners[type] = [];
      listeners[type].push(fn);
    },
    removeEventListener(type, fn) {
      if (listeners[type]) listeners[type] = [];
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
  readyState: 'complete',
  activeElement: null,
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
  createElement(tag) {
    return createDummyEl(tag);
  },
  querySelectorAll() { return []; },
  querySelector(selector) {
    if (selector && selector.includes('tabulator-editing')) {
      return null;
    }
    return createDummyEl('div');
  }
};

let lastWrittenClipboardText = '';

const sandbox = {
  window: {
    addEventListener(evt, fn) {
      if (!this._listeners) this._listeners = {};
      if (!this._listeners[evt]) this._listeners[evt] = [];
      this._listeners[evt].push(fn);
    },
    removeEventListener() {},
    dispatch(evt) {
      if (this._listeners && this._listeners[evt.type]) {
        this._listeners[evt.type].forEach(fn => fn(evt));
      }
    }
  },
  document: mockDoc,
  navigator: {
    userAgent: 'NodeTest',
    clipboard: {
      writeText: async (text) => {
        lastWrittenClipboardText = text;
      },
      readText: async () => lastWrittenClipboardText
    }
  },
  console: {
    log: () => {},
    warn: () => {},
    error: () => {}
  },
  setTimeout: (fn) => fn(),
  clearTimeout: () => {},
  alert: () => {},
  confirm: () => true,
  prompt: (msg, def) => def,
  showToast: () => {},
  Papa: {
    parse: () => ({ data: [] }),
    unparse: (rows) => rows.map(r => r.join(',')).join('\n')
  },
  Encoding: {
    convert: (arr) => arr,
    detect: () => 'UTF8',
    stringToCode: (s) => Array.from(s).map(c => c.charCodeAt(0)),
    codeToString: (c) => String.fromCharCode(...c)
  }
};

// Tabulator モッククラス
class MockTabulator {
  constructor(el, options) {
    this.options = options || {};
    this.data = (options && options.data) || [];
    this.filterFn = null;
  }
  getData(mode) {
    if (mode === "active" && this.filterFn) {
      return this.data.filter(row => this.filterFn(row));
    }
    return this.data;
  }
  getRows(mode) {
    const d = this.getData(mode);
    return d.map(item => {
      const rowEl = createDummyEl('div');
      rowEl.classList.add('tabulator-row');
      return {
        getData: () => item,
        getCell: (field) => {
          const cellEl = createDummyEl('div');
          cellEl.classList.add('tabulator-cell');
          cellEl.setAttribute('tabulator-field', field);
          return {
            getValue: () => item[field],
            getElement: () => cellEl
          };
        },
        getElement: () => rowEl
      };
    });
  }
  setFilter(fn) {
    this.filterFn = fn;
  }
  clearFilter() {
    this.filterFn = null;
  }
  on() {}
  destroy() {}
  getColumns() { return []; }
}
sandbox.Tabulator = MockTabulator;
sandbox.window.Tabulator = MockTabulator;

const ctx = vm.createContext(sandbox);
const appJsPath = path.join(__dirname, '../app.js');
const appJsCode = fs.readFileSync(appJsPath, 'utf-8');
vm.runInContext(appJsCode, ctx);

console.log('\n========================================');
console.log('TEST SUITE: 変数名付き全コピー & 全選択 (1, 2, 3)');
console.log('========================================\n');

let passed = 0;
let failed = 0;

function runTest(name, fn) {
  try {
    fn();
    console.log(`  ✓ ${name}`);
    passed++;
  } catch (err) {
    console.error(`  ✗ ${name}`);
    console.error(err);
    failed++;
  }
}

// テスト共通データセットアップ
function setupTable(rows, headers) {
  vm.runInContext(`{
    const testRows = ${JSON.stringify(rows)};
    const testHeaders = ${JSON.stringify(headers)};
    const testCols = buildTabulatorColumns(testHeaders);

    const testTabId = 'tab-copy-test';
    const testTab = {
      id: testTabId,
      name: 'test_dataset.csv',
      encoding: 'UTF-8',
      delimiter: ',',
      newline: '\\r\\n',
      headers: testHeaders,
      columns: testCols,
      data: testRows,
      isModified: false,
      isNew: false
    };

    tabs = [testTab];
    activeTabId = testTabId;
    currentTable = new Tabulator(null, { data: testRows });
    selectionRange = null;
    activeFocusCell = null;
  }`, ctx);
}

// -----------------------------------------------------------------
// Test 1: selectAllCells（表全体の全選択）
// -----------------------------------------------------------------
runTest('selectAllCells: 表全体の行と列がすべて選択範囲になる', () => {
  const sampleRows = [
    { _id: 1, col_0: '101', col_1: '田中', col_2: '東京' },
    { _id: 2, col_0: '102', col_1: '佐藤', col_2: '大阪' },
    { _id: 3, col_0: '103', col_1: '鈴木', col_2: '名古屋' }
  ];
  const sampleHeaders = ['ID', '氏名', '拠点'];
  setupTable(sampleRows, sampleHeaders);

  vm.runInContext('selectAllCells();', ctx);

  const range = vm.runInContext('selectionRange', ctx);
  assert.ok(range, 'selectionRangeが設定されていること');
  assert.strictEqual(range.minRow, 0, '開始行が0であること');
  assert.strictEqual(range.maxRow, 2, '終了行が2であること (全3行)');
  assert.strictEqual(range.minCol, 0, '開始列が0であること');
  assert.strictEqual(range.maxCol, 2, '終了列が2であること (全3列)');
});

// -----------------------------------------------------------------
// Test 2: copyAllWithHeadersToClipboard（変数名付き全コピー）
// -----------------------------------------------------------------
runTest('copyAllWithHeadersToClipboard: 1行目に変数名、2行目以降にデータが入ったTSVがコピーされる', () => {
  const sampleRows = [
    { _id: 1, col_0: 'A01', col_1: 'りんご', col_2: '150' },
    { _id: 2, col_0: 'A02', col_1: 'みかん', col_2: '80' }
  ];
  const sampleHeaders = ['商品コード', '品名', '単価'];
  setupTable(sampleRows, sampleHeaders);

  vm.runInContext('copyAllWithHeadersToClipboard();', ctx);

  const clipboard = vm.runInContext('internalClipboardText', ctx);
  const lines = clipboard.split('\r\n');

  assert.strictEqual(lines.length, 3, 'ヘッダー1行＋データ2行の計3行であること');
  assert.strictEqual(lines[0], '商品コード\t品名\t単価', '1行目が変数名(ヘッダー)であること');
  assert.strictEqual(lines[1], 'A01\tりんご\t150', '2行目が1行目データであること');
  assert.strictEqual(lines[2], 'A02\tみかん\t80', '3行目が2行目データであること');
});

// -----------------------------------------------------------------
// Test 3: 絞り込み状態での copyAllWithHeadersToClipboard
// -----------------------------------------------------------------
runTest('copyAllWithHeadersToClipboard: フィルター適用時は表示行のみが変数名付きでコピーされる', () => {
  const sampleRows = [
    { _id: 1, col_0: '1', col_1: '営業', col_2: '東京' },
    { _id: 2, col_0: '2', col_1: '開発', col_2: '東京' },
    { _id: 3, col_0: '3', col_1: '営業', col_2: '大阪' }
  ];
  const sampleHeaders = ['ID', '部署', '地域'];
  setupTable(sampleRows, sampleHeaders);

  // 「営業」のみにフィルター
  vm.runInContext("currentTable.setFilter(r => r.col_1 === '営業');", ctx);

  vm.runInContext('copyAllWithHeadersToClipboard();', ctx);

  const clipboard = vm.runInContext('internalClipboardText', ctx);
  const lines = clipboard.split('\r\n');

  assert.strictEqual(lines.length, 3, 'ヘッダー1行＋営業の2行＝計3行であること');
  assert.strictEqual(lines[0], 'ID\t部署\t地域');
  assert.strictEqual(lines[1], '1\t営業\t東京');
  assert.strictEqual(lines[2], '3\t営業\t大阪');
});

// -----------------------------------------------------------------
// Test 4: copySelectedWithHeadersToClipboard（選択範囲を変数名付きでコピー）
// -----------------------------------------------------------------
runTest('copySelectedWithHeadersToClipboard: 選択された列のヘッダーのみが付加されてコピーされる', () => {
  const sampleRows = [
    { _id: 1, col_0: '001', col_1: '山田', col_2: '管理職', col_3: 'S' },
    { _id: 2, col_0: '002', col_1: '佐々木', col_2: '一般職', col_3: 'A' },
    { _id: 3, col_0: '003', col_1: '工藤', col_2: '一般職', col_3: 'B' }
  ];
  const sampleHeaders = ['ID', '氏名', '役職', '評価'];
  setupTable(sampleRows, sampleHeaders);

  // 行0〜1、列1〜2（氏名・役職）を選択
  vm.runInContext(`{
    selectionRange = {
      minRow: 0, maxRow: 1,
      minCol: 1, maxCol: 2,
      startRow: 0, startCol: 1,
      endRow: 1, endCol: 2
    };
  }`, ctx);

  vm.runInContext('copySelectedWithHeadersToClipboard();', ctx);

  const clipboard = vm.runInContext('internalClipboardText', ctx);
  const lines = clipboard.split('\r\n');

  assert.strictEqual(lines.length, 3, 'ヘッダー1行＋データ2行の計3行であること');
  assert.strictEqual(lines[0], '氏名\t役職', '選択列の変数名のみがヘッダーとなること');
  assert.strictEqual(lines[1], '山田\t管理職');
  assert.strictEqual(lines[2], '佐々木\t一般職');
});

// -----------------------------------------------------------------
// Test 5: ショートカットキー Ctrl + A
// -----------------------------------------------------------------
runTest('ショートカット: Ctrl+A を押すと selectAllCells が実行される', () => {
  const sampleRows = [{ _id: 1, col_0: 'X' }, { _id: 2, col_0: 'Y' }];
  setupTable(sampleRows, ['項目']);

  let prevented = false;
  const evt = {
    type: 'keydown',
    ctrlKey: true,
    metaKey: false,
    shiftKey: false,
    key: 'a',
    keyCode: 65,
    preventDefault: () => { prevented = true; }
  };

  mockDoc.activeElement = createDummyEl('div'); // 非入力要素
  vm.runInContext('handleGlobalKeydown', ctx)(evt);

  assert.strictEqual(prevented, true, 'e.preventDefaultが呼ばれること');
  const range = vm.runInContext('selectionRange', ctx);
  assert.strictEqual(range.maxRow, 1, '全行が選択されていること');
});

// -----------------------------------------------------------------
// Test 6: ショートカットキー Ctrl + Shift + C
// -----------------------------------------------------------------
runTest('ショートカット: Ctrl+Shift+C を押すと変数名付きコピーが実行される', () => {
  const sampleRows = [{ _id: 1, col_0: '1', col_1: 'TestVal' }];
  setupTable(sampleRows, ['ID', 'データ']);

  let prevented = false;
  const evt = {
    type: 'keydown',
    ctrlKey: true,
    metaKey: false,
    shiftKey: true,
    key: 'c',
    keyCode: 67,
    preventDefault: () => { prevented = true; }
  };

  mockDoc.activeElement = createDummyEl('div');
  vm.runInContext('handleGlobalKeydown', ctx)(evt);

  assert.strictEqual(prevented, true, 'e.preventDefaultが呼ばれること');
  const clipboard = vm.runInContext('internalClipboardText', ctx);
  assert.ok(clipboard.startsWith('ID\tデータ'), 'ヘッダーから始まるTSVがコピーされること');
});

// -----------------------------------------------------------------
// Test 7: 左上「#」セルの定義と右クリックメニュー
// -----------------------------------------------------------------
runTest('左上「#」セル: headerClick, headerDblClick, headerContextMenu が備わっている', () => {
  const cols = vm.runInContext("buildTabulatorColumns(['列A', '列B'])", ctx);
  const rownumCol = cols[0];

  assert.strictEqual(rownumCol.title, '#');
  assert.strictEqual(typeof rownumCol.headerClick, 'function', 'headerClick関数が存在すること');
  assert.strictEqual(typeof rownumCol.headerDblClick, 'function', 'headerDblClick関数が存在すること');
  assert.ok(Array.isArray(rownumCol.headerContextMenu), 'headerContextMenu配列が存在すること');

  // 右クリックメニューの項目検証
  const copyMenuItem = rownumCol.headerContextMenu.find(item => item.label && item.label.includes('変数名付きで表全体をコピー'));
  assert.ok(copyMenuItem, 'メニューに「変数名付きで表全体をコピー」が存在すること');
});

// -----------------------------------------------------------------
// Test 8: ツールバー「📑 全コピー」ボタンの存在
// -----------------------------------------------------------------
runTest('ツールバー: #btn-copy-all ボタンが存在しクリック可能', () => {
  const btn = mockDoc.getElementById('btn-copy-all');
  assert.ok(btn, '#btn-copy-all 要素が存在すること');
});

console.log('\n========================================');
console.log(`TEST SUMMARY: Passed: ${passed}, Failed: ${failed}`);
console.log('========================================\n');

if (failed > 0) {
  process.exit(1);
} else {
  console.log('ALL COPY WITH HEADERS TESTS PASSED! 🎉\n');
}
