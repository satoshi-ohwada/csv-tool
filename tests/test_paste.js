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
    querySelector() { return null; },
    querySelectorAll() { return []; },
    closest() { return null; },
    scrollIntoView() {},
    focus() {},
    blur() {}
  };
  return el;
}

let activeEditingElement = null;
let currentActiveElement = null;

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
  querySelector(selector) {
    if (selector === '.tabulator-cell.tabulator-editing') {
      return activeEditingElement;
    }
    return null;
  },
  querySelectorAll() { return []; },
  createElement(tag) { return createDummyEl(tag); },
  body: createDummyEl('body'),
  documentElement: createDummyEl('html'),
  get activeElement() {
    return currentActiveElement || mockDoc.body;
  }
};

function MockTabulator(container, options = {}) {
  this.options = options;
  this.data = options.data ? [...options.data] : [];
  this.columns = options.columns ? [...options.columns] : [];
  this._rows = [];

  this._buildRows = () => {
    this._rows = this.data.map((d, i) => {
      const rowEl = createDummyEl('div');
      rowEl.classList.add('tabulator-row');
      const cells = {};
      return {
        _data: d,
        getData: () => d,
        getElement: () => rowEl,
        getIndex: () => d._id !== undefined ? d._id : i + 1,
        update: (nd) => { Object.assign(d, nd); },
        delete: () => {
          const idx = this.data.indexOf(d);
          if (idx !== -1) {
            this.data.splice(idx, 1);
            this._buildRows();
          }
        },
        getPosition: () => i + 1,
        getCell: (field) => {
          if (!cells[field]) {
            const cellEl = createDummyEl('div');
            cellEl.classList.add('tabulator-cell');
            cellEl.setAttribute('tabulator-field', field);
            cellEl.closest = (s) => (s === '.tabulator-row' ? rowEl : null);
            cells[field] = {
              getElement: () => cellEl,
              getValue: () => d[field],
              setValue: (v) => { d[field] = v; },
              edit: (select) => {
                this._lastEditedCell = { rowIndex: i, field };
              }
            };
          }
          return cells[field];
        },
        getCells: () => {
          return this.columns.filter(c => c.field).map(c => this.getCell(c.field));
        }
      };
    });
  };
  this._buildRows();

  this.destroy = () => {};
  this.getData = () => this.data;
  this.getRows = () => this._rows;
  this.replaceData = (newData) => {
    this.data = newData || [];
    this._buildRows();
    return Promise.resolve();
  };
  this.setData = (newData) => {
    this.data = newData || [];
    this._buildRows();
    return Promise.resolve();
  };
  this.setColumns = (newCols) => {
    this.columns = newCols || [];
    this._buildRows();
    return Promise.resolve();
  };
  this.addRow = (rowObj, addToTop, targetRow) => {
    if (targetRow) {
      const targetData = targetRow.getData ? targetRow.getData() : targetRow;
      const idx = this.data.indexOf(targetData);
      if (idx !== -1) {
        const insertIdx = addToTop ? idx : idx + 1;
        this.data.splice(insertIdx, 0, rowObj);
        this._buildRows();
        return Promise.resolve(this._rows[insertIdx]);
      }
    }
    if (addToTop) {
      this.data.unshift(rowObj);
      this._buildRows();
      return Promise.resolve(this._rows[0]);
    } else {
      this.data.push(rowObj);
      this._buildRows();
      return Promise.resolve(this._rows[this.data.length - 1]);
    }
  };
  this.getSelectedRows = () => [];
  this.getColumn = (field) => {
    const colDef = this.columns.find(c => c.field === field);
    if (!colDef) return null;
    return {
      getField: () => field,
      getDefinition: () => colDef,
      delete: () => {
        const idx = this.columns.findIndex(c => c.field === field);
        if (idx !== -1) {
          this.columns.splice(idx, 1);
          this._buildRows();
        }
      }
    };
  };
  this.addColumn = (colDef, before, targetCol) => {
    if (targetCol) {
      const idx = this.columns.findIndex(c => c.field === targetCol.getField());
      if (idx !== -1) {
        const insertIdx = before ? idx : idx + 1;
        this.columns.splice(insertIdx, 0, colDef);
        this._buildRows();
        return Promise.resolve();
      }
    }
    this.columns.push(colDef);
    this._buildRows();
    return Promise.resolve();
  };
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

const PapaReal = require('../lib/papaparse.min.js');

let lastWrittenClipboard = '';

const mockWindow = {
  document: mockDoc,
  Tabulator: MockTabulator,
  Papa: PapaReal,
  localStorage: { getItem: () => null, setItem: () => {}, removeItem: () => {} },
  addEventListener() {},
  removeEventListener() {},
  alert() {},
  confirm: () => true,
  prompt: () => '',
  setTimeout: (fn, ms) => {
    if (typeof fn === 'function') fn();
    return 1;
  },
  clearTimeout: (id) => {},
  setInterval() {},
  clearInterval() {},
  navigator: {
    clipboard: {
      writeText: (t) => { lastWrittenClipboard = t; return Promise.resolve(); },
      readText: () => Promise.resolve(lastWrittenClipboard)
    }
  },
  Date: Date,
  parseInt: parseInt,
  parseFloat: parseFloat,
  String: String,
  Array: Array,
  Object: Object,
  console: console
};

const sandbox = {
  window: mockWindow,
  document: mockDoc,
  Tabulator: MockTabulator,
  Papa: PapaReal,
  localStorage: mockWindow.localStorage,
  setTimeout: mockWindow.setTimeout,
  clearTimeout: mockWindow.clearTimeout,
  setInterval: mockWindow.setInterval,
  clearInterval: mockWindow.clearInterval,
  alert: mockWindow.alert,
  confirm: mockWindow.confirm,
  prompt: mockWindow.prompt,
  console: console,
  navigator: mockWindow.navigator
};

const ctx = vm.createContext(sandbox);

const appJsPath = path.join(__dirname, '../app.js');
const appJsCode = fs.readFileSync(appJsPath, 'utf-8');
vm.runInContext(appJsCode, ctx);

console.log('\n========================================');
console.log('TEST SUITE: Excel範囲コピー & ペースト機能');
console.log('========================================\n');

let passed = 0;
let failed = 0;

async function runTest(name, fn) {
  try {
    await fn();
    console.log(`  ✓ ${name}`);
    passed++;
  } catch (err) {
    console.error(`  ✗ ${name}`);
    console.error(err);
    failed++;
  }
}

function setupTestTab(data, headers) {
  return vm.runInContext(`
    (() => {
      const testHeaders = ${JSON.stringify(headers)};
      const testCols = buildTabulatorColumns(testHeaders);
      const testData = ${JSON.stringify(data)};
      const tabId = 'tab-paste-test-' + Date.now();
      const testTab = {
        id: tabId,
        name: 'test.csv',
        encoding: 'UTF-8',
        delimiter: ',',
        newline: '\\r\\n',
        headers: testHeaders,
        columns: testCols,
        data: testData,
        isModified: false,
        isNew: false
      };
      tabs = [testTab];
      activeTabId = tabId;
      currentTable = new Tabulator(null, { data: testData, columns: testCols });
      selectionRange = null;
      activeFocusCell = null;
      return testTab;
    })()
  `, ctx);
}

async function start() {
  // Test 1: Excelからの基本TSV貼り付け（2行 × 2列）
  await runTest('基本TSV貼り付け: Excelからコピーした2x2の値が指定セルを起点に展開される', async () => {
    const initialRows = [
      { _id: 1, col_0: 'A', col_1: 'B', col_2: 'C' },
      { _id: 2, col_0: 'D', col_1: 'E', col_2: 'F' },
      { _id: 3, col_0: 'G', col_1: 'H', col_2: 'I' }
    ];
    setupTestTab(initialRows, ['列1', '列2', '列3']);

    const res = await vm.runInContext(`
      (async () => {
        selectSingleCell(1, 1); // B2セル(col_1, row 1)を起点
        const pasteEvt = {
          preventDefault: () => {},
          clipboardData: {
            getData: () => "X1\\tY1\\nX2\\tY2"
          }
        };
        await handleGlobalPaste(pasteEvt);
        const t = tabs[0];
        return {
          row0: t.data[0],
          row1: t.data[1],
          row2: t.data[2],
          range: selectionRange
        };
      })()
    `, ctx);

    assert.strictEqual(res.row1.col_1, 'X1');
    assert.strictEqual(res.row1.col_2, 'Y1');
    assert.strictEqual(res.row2.col_1, 'X2');
    assert.strictEqual(res.row2.col_2, 'Y2');
    assert.strictEqual(res.row0.col_1, 'B'); // 未変更
    assert.strictEqual(res.range.minRow, 1);
    assert.strictEqual(res.range.maxRow, 2);
    assert.strictEqual(res.range.minCol, 1);
    assert.strictEqual(res.range.maxCol, 2);
  });

  // Test 2: 改行・クォートを含むExcelセルデータのパース
  await runTest('Excel改行セル貼り付け: セル内に改行を含むデータも壊れずに1セルとして貼り付けられる', async () => {
    const initialRows = [
      { _id: 1, col_0: '', col_1: '' },
      { _id: 2, col_0: '', col_1: '' }
    ];
    setupTestTab(initialRows, ['列1', '列2']);

    const res = await vm.runInContext(`
      (async () => {
        selectSingleCell(0, 0);
        const excelTsv = '"東京都\\n港区"\\t"営業部"\\r\\n"大阪府\\n北区"\\t"開発部"';
        const pasteEvt = {
          preventDefault: () => {},
          clipboardData: { getData: () => excelTsv }
        };
        await handleGlobalPaste(pasteEvt);
        return tabs[0].data;
      })()
    `, ctx);

    assert.strictEqual(res[0].col_0, '東京都\n港区');
    assert.strictEqual(res[0].col_1, '営業部');
    assert.strictEqual(res[1].col_0, '大阪府\n北区');
    assert.strictEqual(res[1].col_1, '開発部');
  });

  // Test 3: CSV形式テキスト（カンマ区切り）の貼り付け
  await runTest('CSV貼り付け: タブではなくカンマ区切りのデータも正しく複数列に分解されて貼り付けられる', async () => {
    const initialRows = [
      { _id: 1, col_0: '', col_1: '' }
    ];
    setupTestTab(initialRows, ['列1', '列2']);

    const res = await vm.runInContext(`
      (async () => {
        selectSingleCell(0, 0);
        const csvText = '商品A,1500';
        const pasteEvt = {
          preventDefault: () => {},
          clipboardData: { getData: () => csvText }
        };
        await handleGlobalPaste(pasteEvt);
        return tabs[0].data[0];
      })()
    `, ctx);

    assert.strictEqual(res.col_0, '商品A');
    assert.strictEqual(res.col_1, '1500');
  });

  // Test 4: 1セルコピー × 複数セル範囲選択への一括貼り付け（Excel同様の充填）
  await runTest('Excel互換充填: 1セルをコピーして複数セル範囲を選択して貼り付けると選択セル全てに同一値が入る', async () => {
    const initialRows = [
      { _id: 1, col_0: '1', col_1: '2' },
      { _id: 2, col_0: '3', col_1: '4' },
      { _id: 3, col_0: '5', col_1: '6' }
    ];
    setupTestTab(initialRows, ['列1', '列2']);

    const res = await vm.runInContext(`
      (async () => {
        // 0行〜1行、0列〜1列を選択
        selectionRange = { minRow: 0, maxRow: 1, minCol: 0, maxCol: 1 };
        activeFocusCell = { rowIndex: 0, colIndex: 0 };
        const pasteEvt = {
          preventDefault: () => {},
          clipboardData: { getData: () => "OK" }
        };
        await handleGlobalPaste(pasteEvt);
        return tabs[0].data;
      })()
    `, ctx);

    assert.strictEqual(res[0].col_0, 'OK');
    assert.strictEqual(res[0].col_1, 'OK');
    assert.strictEqual(res[1].col_0, 'OK');
    assert.strictEqual(res[1].col_1, 'OK');
    assert.strictEqual(res[2].col_0, '5'); // 選択外はそのまま
  });

  // Test 5: 行数が不足している場合の自動行追加
  await runTest('行自動追加: 表の末尾を超える行数のデータを貼り付けた場合、自動で行が追加される', async () => {
    const initialRows = [
      { _id: 1, col_0: '1', col_1: '2' }
    ];
    setupTestTab(initialRows, ['列1', '列2']);

    const res = await vm.runInContext(`
      (async () => {
        selectSingleCell(0, 0);
        const multiRowTsv = "R1C1\\tR1C2\\nR2C1\\tR2C2\\nR3C1\\tR3C2";
        const pasteEvt = {
          preventDefault: () => {},
          clipboardData: { getData: () => multiRowTsv }
        };
        await handleGlobalPaste(pasteEvt);
        return tabs[0].data;
      })()
    `, ctx);

    assert.strictEqual(res.length, 3);
    assert.strictEqual(res[2].col_0, 'R3C1');
    assert.strictEqual(res[2].col_1, 'R3C2');
  });

  // Test 6: 列数が不足している場合の自動列追加
  await runTest('列自動追加: 表の右端を超える列数のデータを貼り付けた場合、自動で列が追加される', async () => {
    const initialRows = [
      { _id: 1, col_0: '1', col_1: '2' }
    ];
    setupTestTab(initialRows, ['列1', '列2']);

    const res = await vm.runInContext(`
      (async () => {
        selectSingleCell(0, 1); // 2列目(col_1)を起点に3列幅のデータを貼り付け (計4列必要)
        const multiColTsv = "V1\\tV2\\tV3";
        const pasteEvt = {
          preventDefault: () => {},
          clipboardData: { getData: () => multiColTsv }
        };
        await handleGlobalPaste(pasteEvt);
        const t = tabs[0];
        return {
          activeColsCount: t.columns.filter(c => c.field).length,
          row0: t.data[0]
        };
      })()
    `, ctx);

    assert.strictEqual(res.activeColsCount, 4);
    assert.strictEqual(res.row0.col_1, 'V1');
  });

  // Test 7: pasteFromClipboard (コンテキストメニューやボタンからの呼び出し)
  await runTest('pasteFromClipboard: 引数なしでnavigator.clipboardから直接読み取って貼り付けられる', async () => {
    const initialRows = [
      { _id: 1, col_0: '', col_1: '' }
    ];
    setupTestTab(initialRows, ['列1', '列2']);

    lastWrittenClipboard = "FromClipboard1\tFromClipboard2";

    await vm.runInContext(`
      (async () => {
        selectSingleCell(0, 0);
        await pasteFromClipboard();
      })()
    `, ctx);

    const d = vm.runInContext('tabs[0].data[0]', ctx);
    assert.strictEqual(d.col_0, 'FromClipboard1');
    assert.strictEqual(d.col_1, 'FromClipboard2');
  });

  // Test 8: 右クリックコンテキストメニューに貼り付け・コピー・切り取りが整理されている
  await runTest('コンテキストメニュー: 貼り付け(Ctrl+V)、コピー(Ctrl+C)、切り取り(Ctrl+X)が整理されている', () => {
    const menu = vm.runInContext('getRowContextMenu()', ctx);
    const labels = menu.map(m => m.label || '---');
    
    assert.ok(labels.some(l => l.includes('貼り付け (Ctrl+V)')));
    assert.ok(labels.some(l => l.includes('コピー (Ctrl+C)')));
    assert.ok(labels.some(l => l.includes('切り取り (Ctrl+X)')));
    assert.ok(labels.some(l => l.includes('変数名付きでコピー')));
    assert.ok(labels.some(l => l.includes('セルをクリア')));
    // 紛らわしい重複した貼り付け項目（内部バッファ貼り付け等）が排除されていること
    assert.ok(!labels.some(l => l.includes('コピーした行を下に貼り付け')));
  });

  // Test 9: selectEntireColumn (列ヘッダークリックで列全体選択)
  await runTest('selectEntireColumn: 列ヘッダーをクリックするとその列全体が選択される', () => {
    const initialRows = [
      { _id: 1, col_0: '1', col_1: '2' },
      { _id: 2, col_0: '3', col_1: '4' },
      { _id: 3, col_0: '5', col_1: '6' }
    ];
    setupTestTab(initialRows, ['列1', '列2']);

    vm.runInContext('selectEntireColumn(1)', ctx); // 2列目 (col_1)
    const range = vm.runInContext('selectionRange', ctx);

    assert.strictEqual(range.minRow, 0);
    assert.strictEqual(range.maxRow, 2);
    assert.strictEqual(range.minCol, 1);
    assert.strictEqual(range.maxCol, 1);
  });

  // Test 10: selectEntireRow (行番号クリックで行全体選択)
  await runTest('selectEntireRow: 行番号セルをクリックするとその行全体が選択される', () => {
    const initialRows = [
      { _id: 1, col_0: '1', col_1: '2' },
      { _id: 2, col_0: '3', col_1: '4' }
    ];
    setupTestTab(initialRows, ['列1', '列2']);

    vm.runInContext('selectEntireRow(1)', ctx); // 2行目 (インデックス1)
    const range = vm.runInContext('selectionRange', ctx);

    assert.strictEqual(range.minRow, 1);
    assert.strictEqual(range.maxRow, 1);
    assert.strictEqual(range.minCol, 0);
    assert.strictEqual(range.maxCol, 1);
  });

  // Test 11: cutSelectedCellsToClipboard (切り取り)
  await runTest('cutSelectedCellsToClipboard: 選択セルがコピーされた後に値がクリアされる', () => {
    const initialRows = [
      { _id: 1, col_0: 'CutMe', col_1: 'KeepMe' }
    ];
    setupTestTab(initialRows, ['列1', '列2']);

    vm.runInContext(`
      selectSingleCell(0, 0);
      cutSelectedCellsToClipboard();
    `, ctx);

    const d = vm.runInContext('tabs[0].data[0]', ctx);
    assert.strictEqual(d.col_0, '');
    assert.strictEqual(d.col_1, 'KeepMe');
    assert.strictEqual(lastWrittenClipboard, 'CutMe');
  });

  // Test 12: Excel側で列全体を選択してコピーした場合（大量の空行を含むケース）
  await runTest('Excel列全体コピー貼り付け: 末尾の数千行の空行が自動トリムされ、実データのみが正しく貼り付けられる', async () => {
    const initialRows = [
      { _id: 1, col_0: 'old1', col_1: 'A' },
      { _id: 2, col_0: 'old2', col_1: 'B' },
      { _id: 3, col_0: 'old3', col_1: 'C' }
    ];
    setupTestTab(initialRows, ['列1', '列2']);

    // Excelで1列全体を選択してコピーした時、実データ3行 + 2000行の空行が含まれるシミュレーション
    let excelColumnCopyText = 'New1\r\nNew2\r\nNew3\r\n';
    for (let i = 0; i < 2000; i++) {
      excelColumnCopyText += '\r\n';
    }

    const res = await vm.runInContext(`
      (async () => {
        selectSingleCell(0, 0); // 1列目の先頭を起点
        const pasteEvt = {
          preventDefault: () => {},
          clipboardData: { getData: () => ${JSON.stringify(excelColumnCopyText)} }
        };
        await handleGlobalPaste(pasteEvt);
        const t = tabs[0];
        return {
          rowsCount: t.data.length,
          data: t.data
        };
      })()
    `, ctx);

    assert.strictEqual(res.rowsCount, 3); // 2000行も追加されず3行のまま
    assert.strictEqual(res.data[0].col_0, 'New1');
    assert.strictEqual(res.data[1].col_0, 'New2');
    assert.strictEqual(res.data[2].col_0, 'New3');
    assert.strictEqual(res.data[0].col_1, 'A'); // 他列は保持
  });

  // Test 13: Excel側で行全体を選択してコピーした場合（大量の空タブを含むケース）
  await runTest('Excel行全体コピー貼り付け: 末尾の数百列の空タブが自動トリムされ、実データのみが正しく貼り付けられる', async () => {
    const initialRows = [
      { _id: 1, col_0: 'old1', col_1: 'old2' }
    ];
    setupTestTab(initialRows, ['列1', '列2']);

    // Excelで1行全体を選択してコピーした時、実データ2列 + 500列の空タブが含まれるシミュレーション
    let excelRowCopyText = 'ValA\tValB';
    for (let i = 0; i < 500; i++) {
      excelRowCopyText += '\t';
    }
    excelRowCopyText += '\r\n';

    const res = await vm.runInContext(`
      (async () => {
        selectSingleCell(0, 0); // 1行目の先頭を起点
        const pasteEvt = {
          preventDefault: () => {},
          clipboardData: { getData: () => ${JSON.stringify(excelRowCopyText)} }
        };
        await handleGlobalPaste(pasteEvt);
        const t = tabs[0];
        return {
          colsCount: t.columns.filter(c => c.field).length,
          row0: t.data[0]
        };
      })()
    `, ctx);

    assert.strictEqual(res.colsCount, 2); // 500列も追加されず2列のまま
    assert.strictEqual(res.row0.col_0, 'ValA');
    assert.strictEqual(res.row0.col_1, 'ValB');
  });

  // Test 14: 列ヘッダーのコンテキストメニューに貼り付け・コピー項目が存在する
  await runTest('列ヘッダーコンテキストメニュー: 貼り付け、コピー、切り取りが統一ラベルで用意されている', () => {
    const menu = vm.runInContext('getHeaderContextMenu()', ctx);
    const labels = menu.map(m => m.label || '---');

    assert.ok(labels.some(l => l.includes('貼り付け (Ctrl+V)')));
    assert.ok(labels.some(l => l.includes('コピー (Ctrl+C)')));
    assert.ok(labels.some(l => l.includes('切り取り (Ctrl+X)')));
  });

  // Test 15: 列全体を選択した状態での列データ貼り付け
  await runTest('列選択時の貼り付け: 列ヘッダークリックで列選択後、Excelからコピーした複数行の列データが綺麗に貼り付けられる', async () => {
    const initialRows = [
      { _id: 1, col_0: '1', col_1: 'X' },
      { _id: 2, col_0: '2', col_1: 'Y' },
      { _id: 3, col_0: '3', col_1: 'Z' }
    ];
    setupTestTab(initialRows, ['列1', '列2']);

    const res = await vm.runInContext(`
      (async () => {
        selectEntireColumn(1); // 2列目(col_1)全体を選択
        const colText = "Alpha\\r\\nBeta\\r\\nGamma\\r\\n";
        const pasteEvt = {
          preventDefault: () => {},
          clipboardData: { getData: () => colText }
        };
        await handleGlobalPaste(pasteEvt);
        return tabs[0].data;
      })()
    `, ctx);

    assert.strictEqual(res[0].col_1, 'Alpha');
    assert.strictEqual(res[1].col_1, 'Beta');
    assert.strictEqual(res[2].col_1, 'Gamma');
    assert.strictEqual(res[0].col_0, '1'); // 1列目は影響を受けない
  });

  // Test 16: selectColumnRange による複数列（A〜C列など）の範囲選択
  await runTest('selectColumnRange: 複数列の範囲選択が正しく機能する', () => {
    const initialRows = [
      { _id: 1, col_0: 'A', col_1: 'B', col_2: 'C', col_3: 'D' }
    ];
    setupTestTab(initialRows, ['列A', '列B', '列C', '列D']);

    vm.runInContext('selectColumnRange(0, 2)', ctx); // 列0から列2 (A〜C)
    const range = vm.runInContext('selectionRange', ctx);

    assert.strictEqual(range.minCol, 0);
    assert.strictEqual(range.maxCol, 2);
    assert.strictEqual(range.minRow, 0);
    assert.strictEqual(range.maxRow, 0);
  });

  // Test 17: 700行 × 15列 の大規模データの超高速貼り付け
  await runTest('700行×15列貼り付け: 大規模データもバッチ更新により超高速に一括反映される', async () => {
    // 初期状態: 10行 × 5列
    const initialRows = [];
    for (let i = 0; i < 10; i++) {
      initialRows.push({ _id: i + 1, col_0: `Init_${i}_0`, col_1: `Init_${i}_1` });
    }
    setupTestTab(initialRows, ['列1', '列2', '列3', '列4', '列5']);

    // 700行 × 15列 の TSV 文字列を生成
    const lines = [];
    for (let r = 0; r < 700; r++) {
      const rowVals = [];
      for (let c = 0; c < 15; c++) {
        rowVals.push(`Data_${r}_${c}`);
      }
      lines.push(rowVals.join('\t'));
    }
    const tsvData = lines.join('\r\n');

    const startTime = Date.now();
    await vm.runInContext(`
      (async () => {
        selectEntireColumn(0); // 列0から貼り付け開始
        const pasteEvt = {
          preventDefault: () => {},
          clipboardData: { getData: () => ${JSON.stringify(tsvData)} }
        };
        await handleGlobalPaste(pasteEvt);
      })()
    `, ctx);
    const elapsed = Date.now() - startTime;

    const data = vm.runInContext('tabs[0].data', ctx);
    const cols = vm.runInContext('tabs[0].columns', ctx);

    assert.strictEqual(data.length, 700, '行数が自動拡張されて700行になっていること');
    assert.ok(cols.length >= 15, '列数が自動拡張されて15列以上になっていること');
    assert.strictEqual(data[0].col_0, 'Data_0_0');
    assert.strictEqual(data[699].col_0, 'Data_699_0');
    assert.ok(elapsed < 2000, `実行時間が2秒未満（実際は${elapsed}ms）で超高速に完了すること`);
  });

  // Test 18: ブラウザ権限制限時の貼り付けアシストモーダル表示と反映
  await runTest('権限制限時の貼り付けアシスト: readText失敗時にモーダルが表示され、入力テキストが正しく反映される', async () => {
    const initialRows = [
      { _id: 1, col_0: 'Old_0', col_1: 'Old_1' }
    ];
    setupTestTab(initialRows, ['列1', '列2']);

    await vm.runInContext(`
      (async () => {
        selectSingleCell(0, 0);
        const origClipboard = navigator.clipboard;
        navigator.clipboard = {
          readText: async () => { throw new Error("Clipboard read requires browser privileges"); }
        };
        internalClipboardText = '';

        await pasteFromClipboard();

        const modal = document.getElementById('paste-assist-modal');
        if (modal.style.display !== 'flex') {
          throw new Error('paste-assist-modal should be displayed as flex');
        }

        // テキストエリアへのペーストイベントを発火
        const textarea = document.getElementById('paste-assist-input');
        const pasteEvt = {
          preventDefault: () => {},
          clipboardData: { getData: () => "New_0\\tNew_1" }
        };
        textarea.dispatchEvent(Object.assign(pasteEvt, { type: 'paste' }));

        navigator.clipboard = origClipboard;
      })()
    `, ctx);

    const data = vm.runInContext('tabs[0].data', ctx);
    const modalDisplay = vm.runInContext("document.getElementById('paste-assist-modal').style.display", ctx);

    assert.strictEqual(data[0].col_0, 'New_0');
    assert.strictEqual(data[0].col_1, 'New_1');
    assert.strictEqual(modalDisplay, 'none', '反映後にモーダルが非表示になること');
  });

  // Test 19: promoteRowToHeader による1行目の列名（変数名）一括設定
  await runTest('promoteRowToHeader: 1行目のデータが列名（変数名）に設定され、データ行から除外される', async () => {
    const initialRows = [
      { _id: 1, col_0: '氏名', col_1: '年齢', col_2: '居住地' },
      { _id: 2, col_0: '田中', col_1: '28', col_2: '東京都' },
      { _id: 3, col_0: '佐藤', col_1: '34', col_2: '大阪府' }
    ];
    setupTestTab(initialRows, ['列 1', '列 2', '列 3']);

    vm.runInContext('promoteRowToHeader(0)', ctx);

    const tab = vm.runInContext('tabs[0]', ctx);
    assert.strictEqual(tab.headers[0], '氏名');
    assert.strictEqual(tab.headers[1], '年齢');
    assert.strictEqual(tab.headers[2], '居住地');
    assert.strictEqual(tab.columns[1].title, '氏名');
    assert.strictEqual(tab.columns[2].title, '年齢');
    assert.strictEqual(tab.columns[3].title, '居住地');
    assert.strictEqual(tab.data.length, 2, '1行目が削除されてデータは2行になること');
    assert.strictEqual(tab.data[0].col_0, '田中');
    assert.strictEqual(tab.data[0]._id, 1);
    assert.strictEqual(tab.data[1].col_0, '佐藤');
    assert.strictEqual(tab.data[1]._id, 2);
    assert.strictEqual(tab.isModified, true);
  });

  // Test 20: demoteHeaderToRow による列名の1行目データへの変換（逆変換）
  await runTest('demoteHeaderToRow: 現在の列名が1行目のデータとして挿入され、列名がデフォルトに戻る', async () => {
    const initialRows = [
      { _id: 1, col_0: '田中', col_1: '28' }
    ];
    setupTestTab(initialRows, ['氏名', '年齢']);

    vm.runInContext('demoteHeaderToRow()', ctx);

    const tab = vm.runInContext('tabs[0]', ctx);
    assert.strictEqual(tab.headers[0], '列 1');
    assert.strictEqual(tab.headers[1], '列 2');
    assert.strictEqual(tab.columns[1].title, '列 1');
    assert.strictEqual(tab.columns[2].title, '列 2');
    assert.strictEqual(tab.data.length, 2, '列名行が挿入されて2行になること');
    assert.strictEqual(tab.data[0].col_0, '氏名');
    assert.strictEqual(tab.data[0].col_1, '年齢');
    assert.strictEqual(tab.data[1].col_0, '田中');
    assert.strictEqual(tab.data[1].col_1, '28');
  });

  // Test 21: 空白セルと重複列名のハンドリング
  await runTest('promoteRowToHeader 空白・重複処理: 空白セルにはフォールバック名、重複には連番サフィックスが付与される', async () => {
    const initialRows = [
      { _id: 1, col_0: 'スコア', col_1: '', col_2: 'スコア' },
      { _id: 2, col_0: '100', col_1: 'A', col_2: '90' }
    ];
    setupTestTab(initialRows, ['初期1', '初期2', '初期3']);

    vm.runInContext('promoteRowToHeader(0)', ctx);

    const tab = vm.runInContext('tabs[0]', ctx);
    assert.strictEqual(tab.headers[0], 'スコア');
    assert.strictEqual(tab.headers[1], '初期2', '空セルは元の列名にフォールバック');
    assert.strictEqual(tab.headers[2], 'スコア_2', '重複列名には連番サフィックス');
    assert.strictEqual(tab.data.length, 1);
  });

  // Test 22: executeDataClean による promote_header / demote_header ディスパッチ
  await runTest('executeDataClean: promote_header および demote_header アクションが正しくディスパッチされる', async () => {
    const initialRows = [
      { _id: 1, col_0: '項目A', col_1: '項目B' },
      { _id: 2, col_0: '値1', col_1: '値2' }
    ];
    setupTestTab(initialRows, ['列 1', '列 2']);

    vm.runInContext('executeDataClean("promote_header")', ctx);
    let tab = vm.runInContext('tabs[0]', ctx);
    assert.strictEqual(tab.headers[0], '項目A');
    assert.strictEqual(tab.data.length, 1);

    vm.runInContext('executeDataClean("demote_header")', ctx);
    tab = vm.runInContext('tabs[0]', ctx);
    assert.strictEqual(tab.headers[0], '列 1');
    assert.strictEqual(tab.data.length, 2);
    assert.strictEqual(tab.data[0].col_0, '項目A');
  });

  // Test 23: コンテキストメニューに「列名（変数名）に設定」項目が存在する
  await runTest('コンテキストメニュー: 行メニューとヘッダーメニューに「列名（変数名）に設定」が存在する', () => {
    const rowMenu = vm.runInContext('getRowContextMenu()', ctx);
    const rowLabels = rowMenu.map(m => m.label || '---');
    assert.ok(rowLabels.some(l => l.includes('列名（変数名）に設定')), '行メニューに項目が存在すること');

    const headerMenu = vm.runInContext('getHeaderContextMenu()', ctx);
    const headerLabels = headerMenu.map(m => m.label || '---');
    assert.ok(headerLabels.some(l => l.includes('1行目を列名（変数名）に設定')), 'ヘッダーメニューに項目が存在すること');
  });

  console.log('\n========================================');
  console.log(`TEST SUMMARY: Passed: ${passed}, Failed: ${failed}`);
  console.log('========================================\n');

  if (failed > 0) {
    process.exit(1);
  } else {
    console.log('ALL PASTE & SELECTION TESTS PASSED! 🎉\n');
  }
}

start().catch(err => {
  console.error(err);
  process.exit(1);
});
