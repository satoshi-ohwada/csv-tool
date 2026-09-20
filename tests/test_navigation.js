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

const mockWindow = {
  document: mockDoc,
  Tabulator: MockTabulator,
  Papa: { parse: () => ({ data: [] }), unparse: (d) => '' },
  localStorage: { getItem: () => null, setItem: () => {}, removeItem: () => {} },
  addEventListener() {},
  removeEventListener() {},
  alert() {},
  confirm: () => true,
  prompt: () => '',
  setTimeout: (fn, ms) => {
    if (typeof fn === 'function') {
      // 適切な非同期タイミングで実行
      return setTimeout(fn, 0);
    }
    return 1;
  },
  clearTimeout: (id) => clearTimeout(id),
  setInterval() {},
  clearInterval() {},
  navigator: { clipboard: { writeText: () => Promise.resolve() } },
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
  Papa: mockWindow.Papa,
  localStorage: mockWindow.localStorage,
  setTimeout: mockWindow.setTimeout,
  clearTimeout: mockWindow.clearTimeout,
  setInterval: mockWindow.setInterval,
  clearInterval: mockWindow.clearInterval,
  alert: mockWindow.alert,
  confirm: mockWindow.confirm,
  prompt: mockWindow.prompt,
  console: console,
  navigator: mockWindow.navigator,
  Date: Date,
  parseInt: parseInt,
  parseFloat: parseFloat,
  String: String,
  Array: Array,
  Object: Object
};

vm.createContext(sandbox);

const appJsCode = fs.readFileSync(path.join(__dirname, '../app.js'), 'utf8');
vm.runInContext(appJsCode, sandbox);

// -------------------------------------------------------------
// テストスイートの実行
// -------------------------------------------------------------
console.log('========================================');
console.log('TEST SUITE: キーボード入力ナビゲーション（案1: Enter / Tab）');
console.log('========================================\n');

let passed = 0;
let failed = 0;
const tests = [];

function it(desc, fn) {
  tests.push({ desc, fn });
}

// テスト用テーブルのセットアップヘルパー (4列 × 4行: A〜D列、1〜4行)
function setup4x4Table() {
  vm.runInContext(`
    (() => {
      const columns = [
        { title: '列 A', field: 'col_a', colLetter: 'A' },
        { title: '列 B', field: 'col_b', colLetter: 'B' },
        { title: '列 C', field: 'col_c', colLetter: 'C' },
        { title: '列 D', field: 'col_d', colLetter: 'D' }
      ];
      const data = [
        { _id: 1, col_a: 'A1', col_b: 'B1', col_c: 'C1', col_d: 'D1' },
        { _id: 2, col_a: 'A2', col_b: 'B2', col_c: 'C2', col_d: 'D2' },
        { _id: 3, col_a: 'A3', col_b: 'B3', col_c: 'C3', col_d: 'D3' },
        { _id: 4, col_a: 'A4', col_b: 'B4', col_c: 'C4', col_d: 'D4' }
      ];

      tabs = [{
        id: 'tab_test_nav',
        name: 'test.csv',
        columns: columns,
        headers: columns.map(c => c.title),
        data: data,
        isModified: false
      }];
      activeTabId = 'tab_test_nav';
      currentTable = new Tabulator(null, { data, columns });
      activeFocusCell = null;
      selectionRange = null;
    })()
  `, sandbox);
}

// セル編集モード状態の擬似DOM作成ヘルパー
function simulateEditingCell(rowIndex, colIndex, field, value) {
  const currentTable = vm.runInContext(`currentTable`, sandbox);
  const row = currentTable.getRows()[rowIndex];
  const cell = row.getCell(field);
  const cellEl = cell.getElement();
  cellEl.classList.add('tabulator-editing');

  const inputEl = createDummyEl('input');
  inputEl.value = value;
  cellEl.querySelector = (s) => (s === 'input' ? inputEl : null);

  activeEditingElement = cellEl;
  currentActiveElement = inputEl;

  return { cellEl, inputEl };
}

function clearEditing() {
  if (activeEditingElement) {
    activeEditingElement.classList.remove('tabulator-editing');
  }
  activeEditingElement = null;
  currentActiveElement = null;
}

const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

// -------------------------------------------------------------
// 1. Enter キーのテスト
// -------------------------------------------------------------

it('Enterキー: 最下行かつ最右列 (D4) でEnterを押すと、新しい行 (5行目) が追加され、次行の先頭 (A5) が選択される', async () => {
  setup4x4Table();
  const initialRowCount = vm.runInContext(`getAllRows().length`, sandbox);
  assert.strictEqual(initialRowCount, 4, '初期行数は4');

  simulateEditingCell(3, 3, 'col_d', 'New-D4-Value');

  // Enterイベント発火
  const enterEvent = {
    key: 'Enter',
    isComposing: false,
    keyCode: 13,
    preventDefault() {},
    stopPropagation() {}
  };

  vm.runInContext(`handleRightmostEnterKey`, sandbox)(enterEvent);
  clearEditing();
  await sleep(60);

  const newRowCount = vm.runInContext(`getAllRows().length`, sandbox);
  const tab = vm.runInContext(`tabs[0]`, sandbox);
  const focus = vm.runInContext(`activeFocusCell`, sandbox);

  // 行が追加されて5行になったか
  assert.strictEqual(newRowCount, 5, '5行目に新しい行が追加されていること');
  assert.strictEqual(tab.isModified, true, 'isModifiedがtrueになること');

  // 選択されたセルが (rowIndex: 4, colIndex: 0) = A5 であること
  assert.strictEqual(focus.rowIndex, 4, '行インデックスが4（5行目）であること');
  assert.strictEqual(focus.colIndex, 0, '列インデックスが0（A列）であること');
});

it('Enterキー: 途中の行の最右列 (D1) でEnterを押すと、行は追加されず、既存の次行の先頭 (A2) が選択される', async () => {
  setup4x4Table();
  const initialRowCount = vm.runInContext(`getAllRows().length`, sandbox);
  assert.strictEqual(initialRowCount, 4, '初期行数は4');

  simulateEditingCell(0, 3, 'col_d', 'Edited-D1');

  const enterEvent = {
    key: 'Enter',
    isComposing: false,
    keyCode: 13,
    preventDefault() {},
    stopPropagation() {}
  };

  vm.runInContext(`handleRightmostEnterKey`, sandbox)(enterEvent);
  clearEditing();
  await sleep(60);

  const rowCount = vm.runInContext(`getAllRows().length`, sandbox);
  const focus = vm.runInContext(`activeFocusCell`, sandbox);

  // 行数は4のままであること
  assert.strictEqual(rowCount, 4, '行数は4のまま（追加されない）');

  // 選択されたセルが A2 (rowIndex: 1, colIndex: 0) であること
  assert.strictEqual(focus.rowIndex, 1, '行インデックスが1（2行目）であること');
  assert.strictEqual(focus.colIndex, 0, '列インデックスが0（A列）であること');
});

it('Enterキー: 途中の列 (B1) でEnterを押すと、行は追加されず、一つ右隣のセル (C1) が選択される', async () => {
  setup4x4Table();
  const initialRowCount = vm.runInContext(`getAllRows().length`, sandbox);
  assert.strictEqual(initialRowCount, 4, '初期行数は4');

  simulateEditingCell(0, 1, 'col_b', 'Edited-B1');

  const enterEvent = {
    key: 'Enter',
    isComposing: false,
    keyCode: 13,
    preventDefault() {},
    stopPropagation() {}
  };

  vm.runInContext(`handleRightmostEnterKey`, sandbox)(enterEvent);
  clearEditing();
  await sleep(60);

  const rowCount = vm.runInContext(`getAllRows().length`, sandbox);
  const focus = vm.runInContext(`activeFocusCell`, sandbox);

  // 行数は4のままであること
  assert.strictEqual(rowCount, 4, '行数は4のまま');

  // 選択されたセルが C1 (rowIndex: 0, colIndex: 2) であること
  assert.strictEqual(focus.rowIndex, 0, '行インデックスが0であること');
  assert.strictEqual(focus.colIndex, 2, '列インデックスが2（C列）であること');
});

it('Enterキー: 最下行・最右列でも行全体が空の場合は、不要な空行を追加せずその場にとどまる', async () => {
  setup4x4Table();
  // 4行目を空にする
  vm.runInContext(`
    tabs[0].data[3] = { _id: 4, col_a: '', col_b: '', col_c: '', col_d: '' };
    currentTable.data[3] = { _id: 4, col_a: '', col_b: '', col_c: '', col_d: '' };
    currentTable._buildRows();
  `, sandbox);

  simulateEditingCell(3, 3, 'col_d', '');

  const enterEvent = {
    key: 'Enter',
    isComposing: false,
    keyCode: 13,
    preventDefault() {},
    stopPropagation() {}
  };

  vm.runInContext(`handleRightmostEnterKey`, sandbox)(enterEvent);
  clearEditing();
  await sleep(60);

  const rowCount = vm.runInContext(`getAllRows().length`, sandbox);
  const focus = vm.runInContext(`activeFocusCell`, sandbox);

  // 行数は4のまま
  assert.strictEqual(rowCount, 4, '空行の場合は新しい行を追加しない');
  assert.strictEqual(focus.rowIndex, 3, '現在行にとどまること');
  assert.strictEqual(focus.colIndex, 3, '現在列にとどまること');
});

// -------------------------------------------------------------
// 2. Tab キーのテスト
// -------------------------------------------------------------

it('Tabキー: 途中の列 (A1) でTabを押すと、右隣のセル (B1) へ移動して編集モードになる', async () => {
  setup4x4Table();
  simulateEditingCell(0, 0, 'col_a', 'A1');

  vm.runInContext(`handleTabNavigation`, sandbox)(false);
  clearEditing();
  await sleep(60);

  const focus = vm.runInContext(`activeFocusCell`, sandbox);
  const currentTable = vm.runInContext(`currentTable`, sandbox);

  // B1 (rowIndex: 0, col_b) のセルが編集モードになっていること
  assert.strictEqual(focus.rowIndex, 0, '行インデックスは0');
  assert.strictEqual(focus.colIndex, 1, '列インデックスは1（B列）');
  assert.strictEqual(currentTable._lastEditedCell.rowIndex, 0);
  assert.strictEqual(currentTable._lastEditedCell.field, 'col_b');
});

it('Tabキー: 最右列 (D4) でTabを押すと、新しい列 (E列) が自動追加され、E4へ移動して編集モードになる', async () => {
  setup4x4Table();
  const initialColCount = vm.runInContext(`getActiveColumns().length`, sandbox);
  assert.strictEqual(initialColCount, 4, '初期列数は4');

  simulateEditingCell(3, 3, 'col_d', 'D4');

  vm.runInContext(`handleTabNavigation`, sandbox)(false);
  clearEditing();
  await sleep(80);

  const activeCols = vm.runInContext(`getActiveColumns()`, sandbox);
  const tab = vm.runInContext(`tabs[0]`, sandbox);
  const focus = vm.runInContext(`activeFocusCell`, sandbox);
  const currentTable = vm.runInContext(`currentTable`, sandbox);

  // 新しい列（E列）が追加されて列数が5になったこと
  assert.strictEqual(activeCols.length, 5, '列数が5（E列が追加）になること');
  assert.strictEqual(tab.isModified, true, 'isModifiedがtrueになること');

  const newCol = activeCols[4];
  assert.strictEqual(newCol.colLetter, 'E', '新列の記号がEであること');

  // E4 (rowIndex: 3, newField) のセルが編集モードになっていること
  assert.ok(focus !== null, 'focusがnullではないこと');
  assert.strictEqual(focus.rowIndex, 3, '行インデックスは3');
  assert.strictEqual(focus.colIndex, 4, '列インデックスは4（E列）');
  assert.strictEqual(currentTable._lastEditedCell.rowIndex, 3);
  assert.strictEqual(currentTable._lastEditedCell.field, newCol.field);
});

it('Tabキー: 途中の行の最右列 (D1) でTabを押しても、新しい列 (E列) が自動追加され、E1へ移動して編集モードになる', async () => {
  setup4x4Table();
  const initialColCount = vm.runInContext(`getActiveColumns().length`, sandbox);
  assert.strictEqual(initialColCount, 4, '初期列数は4');

  simulateEditingCell(0, 3, 'col_d', 'D1');

  vm.runInContext(`handleTabNavigation`, sandbox)(false);
  clearEditing();
  await sleep(80);

  const activeCols = vm.runInContext(`getActiveColumns()`, sandbox);
  assert.strictEqual(activeCols.length, 5, '列数が5になること');
  const newCol = activeCols[4];
  const focus = vm.runInContext(`activeFocusCell`, sandbox);
  const currentTable = vm.runInContext(`currentTable`, sandbox);

  // E1 (rowIndex: 0, newField) のセルが編集モードになっていること
  assert.ok(focus !== null, 'focusがnullではないこと');
  assert.strictEqual(focus.rowIndex, 0, '行インデックスは0');
  assert.strictEqual(focus.colIndex, 4, '列インデックスは4');
  assert.strictEqual(currentTable._lastEditedCell.rowIndex, 0);
  assert.strictEqual(currentTable._lastEditedCell.field, newCol.field);
});

it('Shift+Tabキー: 途中の列 (B1) でShift+Tabを押すと、左隣のセル (A1) へ戻って編集モードになる', async () => {
  setup4x4Table();
  simulateEditingCell(0, 1, 'col_b', 'B1');

  vm.runInContext(`handleTabNavigation`, sandbox)(true);
  clearEditing();
  await sleep(60);

  const focus = vm.runInContext(`activeFocusCell`, sandbox);
  const currentTable = vm.runInContext(`currentTable`, sandbox);

  assert.strictEqual(focus.rowIndex, 0, '行インデックスは0');
  assert.strictEqual(focus.colIndex, 0, '列インデックスは0（A列）');
  assert.strictEqual(currentTable._lastEditedCell.rowIndex, 0);
  assert.strictEqual(currentTable._lastEditedCell.field, 'col_a');
});

it('Shift+Tabキー: 行頭の列 (A2) でShift+Tabを押すと、前行の右端セル (D1) へ移動する', async () => {
  setup4x4Table();
  simulateEditingCell(1, 0, 'col_a', 'A2');

  vm.runInContext(`handleTabNavigation`, sandbox)(true);
  clearEditing();
  await sleep(60);

  const focus = vm.runInContext(`activeFocusCell`, sandbox);
  const currentTable = vm.runInContext(`currentTable`, sandbox);

  assert.strictEqual(focus.rowIndex, 0, '前行（0行目）へ移動');
  assert.strictEqual(focus.colIndex, 3, '最右列（D列）へ移動');
  assert.strictEqual(currentTable._lastEditedCell.rowIndex, 0);
  assert.strictEqual(currentTable._lastEditedCell.field, 'col_d');
});

// -------------------------------------------------------------
// 3. 検索窓フォーカス時の矢印キー・Enterキーテスト
// -------------------------------------------------------------

it('検索窓フォーカス時: 矢印キーを押してもセルの移動や選択が発生しない', async () => {
  setup4x4Table();
  // 初期フォーカスセルを (1, 1) に設定
  vm.runInContext(`activeFocusCell = { rowIndex: 1, colIndex: 1 }; selectionRange = { minRow: 1, maxRow: 1, minCol: 1, maxCol: 1 };`, sandbox);

  // 検索窓にフォーカスがある状態を模擬
  const searchEl = createDummyEl('input');
  searchEl.id = 'search-input';
  currentActiveElement = searchEl;

  // 矢印キーイベントを発火
  ['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].forEach(key => {
    let prevented = false;
    const arrowEvent = {
      key: key,
      isComposing: false,
      keyCode: 37,
      preventDefault() { prevented = true; },
      stopPropagation() {}
    };
    vm.runInContext(`handleArrowKeyNavigation`, sandbox)(arrowEvent);
    // 検索窓内なので preventDefault されずセルも移動しないこと
    assert.strictEqual(prevented, false, `${key} で preventDefault されないこと`);
  });

  const focus = vm.runInContext(`activeFocusCell`, sandbox);
  assert.strictEqual(focus.rowIndex, 1, '行インデックスが1のまま（セル移動していないこと）');
  assert.strictEqual(focus.colIndex, 1, '列インデックスが1のまま（セル移動していないこと）');

  currentActiveElement = null;
});

it('検索窓フォーカス時: Enterキーを押してもセル移動や行追加が発生しない', async () => {
  setup4x4Table();
  vm.runInContext(`activeFocusCell = { rowIndex: 3, colIndex: 3 };`, sandbox);

  const searchEl = createDummyEl('input');
  searchEl.id = 'search-input';
  searchEl.value = 'query';
  currentActiveElement = searchEl;

  let prevented = false;
  const enterEvent = {
    key: 'Enter',
    isComposing: false,
    keyCode: 13,
    preventDefault() { prevented = true; },
    stopPropagation() {}
  };

  vm.runInContext(`handleRightmostEnterKey`, sandbox)(enterEvent);

  const rowCount = vm.runInContext(`getAllRows().length`, sandbox);
  const focus = vm.runInContext(`activeFocusCell`, sandbox);

  assert.strictEqual(prevented, false, 'Enterキーで preventDefault されないこと');
  assert.strictEqual(rowCount, 4, '行数は4のまま（行追加されないこと）');
  assert.strictEqual(focus.rowIndex, 3, 'フォーカスセルが変わらないこと');
  assert.strictEqual(focus.colIndex, 3, 'フォーカスセルが変わらないこと');

  currentActiveElement = null;
});

async function runTests() {
  for (const t of tests) {
    try {
      await t.fn();
      console.log(`  ✓ ${t.desc}`);
      passed++;
    } catch (err) {
      console.error(`  ✗ ${t.desc}`);
      console.error(err);
      failed++;
    }
  }

  console.log('\n========================================');
  console.log(`TEST SUMMARY: Passed: ${passed}, Failed: ${failed}`);
  console.log('========================================');

  if (failed > 0) {
    process.exit(1);
  } else {
    console.log('ALL NAVIGATION (OPTION 1) TESTS PASSED! 🎉\n');
  }
}

runTests();
