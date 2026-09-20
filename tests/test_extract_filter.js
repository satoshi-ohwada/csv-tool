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
  querySelector() { return createDummyEl('div'); }
};

let lastToastMessage = null;
let lastToastType = null;

const sandbox = {
  window: {
    addEventListener() {},
    removeEventListener() {},
    showSaveFilePicker: async () => {},
  },
  document: mockDoc,
  navigator: { userAgent: 'NodeTest', clipboard: { writeText: async () => {}, readText: async () => '' } },
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
  showToast: (msg, type) => {
    lastToastMessage = msg;
    lastToastType = type;
  },
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
    return d.map(item => ({
      getData: () => item
    }));
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
console.log('TEST SUITE: 絞り込み結果の新規タブ抽出機能（案B）');
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

// -----------------------------------------------------------------
// Test 1: generateExtractedTabName の命名規則と重複防止
// -----------------------------------------------------------------
runTest('generateExtractedTabName: 基本の拡張子付きファイル名', () => {
  const result = vm.runInContext("generateExtractedTabName('社員名簿.csv')", ctx);
  assert.strictEqual(result, '社員名簿_抽出.csv');
});

runTest('generateExtractedTabName: TSVファイル名の拡張子維持', () => {
  const result = vm.runInContext("generateExtractedTabName('log_data.tsv')", ctx);
  assert.strictEqual(result, 'log_data_抽出.tsv');
});

runTest('generateExtractedTabName: 拡張子なしファイル名', () => {
  const result = vm.runInContext("generateExtractedTabName('exported_dataset')", ctx);
  assert.strictEqual(result, 'exported_dataset_抽出.csv');
});

runTest('generateExtractedTabName: 既存タブに同名が存在する場合は連番付与', () => {
  vm.runInContext(`
    tabs.push({ id: 'dummy-1', name: '売上_抽出.csv' });
  `, ctx);
  const result1 = vm.runInContext("generateExtractedTabName('売上.csv')", ctx);
  assert.strictEqual(result1, '売上_抽出_2.csv');

  vm.runInContext(`
    tabs.push({ id: 'dummy-2', name: '売上_抽出_2.csv' });
  `, ctx);
  const result2 = vm.runInContext("generateExtractedTabName('売上.csv')", ctx);
  assert.strictEqual(result2, '売上_抽出_3.csv');

  // クリーンアップ
  vm.runInContext(`
    tabs = tabs.filter(t => !t.id.startsWith('dummy-'));
  `, ctx);
});

// -----------------------------------------------------------------
// Test 2: 絞り込み結果の新規タブ抽出（正常系）
// -----------------------------------------------------------------
runTest('extractFilteredRowsToNewTab: 絞り込み条件に合致する行のみが新しいタブに抽出される', () => {
  vm.runInContext(`{
    const originalRows = [
      { _id: 1, col_0: '101', col_1: '田中', col_2: '営業部', col_3: '500000' },
      { _id: 2, col_0: '102', col_1: '佐藤', col_2: '開発部', col_3: '650000' },
      { _id: 3, col_0: '103', col_1: '鈴木', col_2: '営業部', col_3: '480000' },
      { _id: 4, col_0: '104', col_1: '高橋', col_2: '総務部', col_3: '420000' },
      { _id: 5, col_0: '105', col_1: '伊藤', col_2: '営業部', col_3: '530000' }
    ];
    const originalHeaders = ['社員ID', '氏名', '部署', '給与'];
    const originalCols = buildTabulatorColumns(originalHeaders);

    const origTabId = 'tab-orig-1';
    const origTab = {
      id: origTabId,
      name: '従業員リスト.csv',
      fileHandle: { createWritable: async () => {} },
      rawBytes: null,
      encoding: 'Shift_JIS',
      delimiter: ',',
      newline: '\\r\\n',
      commaMode: 'keep',
      headers: originalHeaders,
      columns: originalCols,
      data: originalRows,
      isModified: false,
      isNew: false
    };

    tabs = [origTab];
    activeTabId = origTabId;
    currentTable = new Tabulator(null, { data: originalRows });

    // 「営業部」で絞り込み中の状態を再現
    currentTable.setFilter(row => row.col_2 === '営業部');
    searchInput.value = '営業部';

    // 抽出を実行
    globalThis.__testNewTab = extractFilteredRowsToNewTab();
  }`, ctx);

  const tabs = vm.runInContext('tabs', ctx);
  const activeTabId = vm.runInContext('activeTabId', ctx);
  const origTab = tabs[0];
  const newTab = tabs[1];

  // 検証: 新規タブが作成され、アクティブになっているか
  assert.ok(newTab, '新規タブが返されること');
  assert.strictEqual(activeTabId, newTab.id, '新規タブがアクティブ化されていること');
  assert.strictEqual(tabs.length, 2, 'タブ総数が2つになっていること');

  // 検証: 新規タブのデータと行数
  assert.strictEqual(newTab.name, '従業員リスト_抽出.csv', 'タブ名が「_抽出.csv」となっていること');
  assert.strictEqual(newTab.data.length, 3, '営業部の3名のみが抽出されていること');
  assert.strictEqual(newTab.data[0].col_1, '田中');
  assert.strictEqual(newTab.data[1].col_1, '鈴木');
  assert.strictEqual(newTab.data[2].col_1, '伊藤');

  // 検証: 行ID (_id) が 1 から連番で再採番されていること
  assert.strictEqual(newTab.data[0]._id, 1);
  assert.strictEqual(newTab.data[1]._id, 2);
  assert.strictEqual(newTab.data[2]._id, 3);

  // 検証: 元タブの設定が継承され、かつ新規保存用に初期化されていること
  assert.strictEqual(newTab.encoding, 'Shift_JIS', 'Shift_JIS文字コードが継承されること');
  assert.strictEqual(newTab.delimiter, ',', '区切り文字が継承されること');
  assert.strictEqual(newTab.newline, '\r\n', '改行コードが継承されること');
  assert.strictEqual(newTab.fileHandle, null, '元ファイルへの上書き防止のためfileHandleはnullであること');
  assert.strictEqual(newTab.isNew, true, 'isNewがtrueであること');
  assert.strictEqual(newTab.isModified, true, 'isModifiedがtrueであること');

  // 検証: 元のタブのデータ（5名分）が無傷のまま残っていること
  assert.strictEqual(origTab.data.length, 5, '元タブのデータ行数が維持されていること');
  assert.strictEqual(origTab.name, '従業員リスト.csv', '元タブの名前が変わっていないこと');
});

// -----------------------------------------------------------------
// Test 3: 0件ヒット時の安全ガード
// -----------------------------------------------------------------
runTest('extractFilteredRowsToNewTab: 絞り込み条件に一致する行が0件の場合は抽出を中断し警告を表示', () => {
  vm.runInContext(`{
    const sampleRows = [
      { _id: 1, col_0: 'A', col_1: '10' }
    ];
    const origTabId = 'tab-guard-1';
    const origTab = {
      id: origTabId,
      name: 'guard_test.csv',
      columns: buildTabulatorColumns(['列1', '列2']),
      data: sampleRows
    };
    tabs = [origTab];
    activeTabId = origTabId;
    currentTable = new Tabulator(null, { data: sampleRows });

    // 存在しないキーワードで0件ヒット状態
    searchInput.value = '存在しない文字列xyz';
    currentTable.setFilter(() => false);

    globalThis.__testResult = extractFilteredRowsToNewTab();
  }`, ctx);

  const tabs = vm.runInContext('tabs', ctx);
  const result = vm.runInContext('globalThis.__testResult', ctx);

  assert.strictEqual(result, null, '0件時はnullが返されること');
  assert.strictEqual(tabs.length, 1, 'タブは追加されないこと');
});

// -----------------------------------------------------------------
// Test 4: フィルターなし時の全行複製・抽出
// -----------------------------------------------------------------
runTest('extractFilteredRowsToNewTab: 絞り込みなし時は全行が新しいタブに複製・抽出される', () => {
  vm.runInContext(`{
    const sampleRows = [
      { _id: 1, col_0: '1', col_1: '商品A' },
      { _id: 2, col_0: '2', col_1: '商品B' }
    ];
    const origTabId = 'tab-full-1';
    const origTab = {
      id: origTabId,
      name: 'master.csv',
      columns: buildTabulatorColumns(['ID', '商品名']),
      data: sampleRows,
      encoding: 'UTF-8',
      delimiter: ',',
      newline: '\\n'
    };
    tabs = [origTab];
    activeTabId = origTabId;
    currentTable = new Tabulator(null, { data: sampleRows });
    searchInput.value = '';
    currentTable.clearFilter();

    globalThis.__testNewTabFull = extractFilteredRowsToNewTab();
  }`, ctx);

  const newTabFull = vm.runInContext('globalThis.__testNewTabFull', ctx);
  assert.ok(newTabFull);
  assert.strictEqual(newTabFull.name, 'master_抽出.csv');
  assert.strictEqual(newTabFull.data.length, 2);
  assert.strictEqual(newTabFull.data[0].col_1, '商品A');
  assert.strictEqual(newTabFull.data[1].col_1, '商品B');
});


// -----------------------------------------------------------------
// Test 5: UI連動（検索窓に文字があるときのボタン強調クラス）
// -----------------------------------------------------------------
runTest('updateExtractFilterButtonUI: 絞り込み状態に応じて抽出ボタンのhas-filterクラスがトグルされる', () => {
  const btn = mockDoc.getElementById('btn-extract-filter');
  assert.ok(btn, '抽出ボタン要素が存在すること');

  // 検索クエリあり -> has-filter が付与
  vm.runInContext(`
    searchInput.value = 'テスト';
    updateExtractFilterButtonUI();
  `, ctx);
  assert.strictEqual(btn.classList.contains('has-filter'), true, '検索クエリ入力時にhas-filterが付与されること');

  // 検索クエリなし -> has-filter が解除
  vm.runInContext(`
    searchInput.value = '';
    updateExtractFilterButtonUI();
  `, ctx);
  assert.strictEqual(btn.classList.contains('has-filter'), false, '検索クエリクリア時にhas-filterが解除されること');
});

console.log('\n========================================');
console.log(`TEST SUMMARY: Passed: ${passed}, Failed: ${failed}`);
console.log('========================================\n');

if (failed > 0) {
  process.exit(1);
} else {
  console.log('ALL EXTRACT FILTER TESTS PASSED! 🎉\n');
}
