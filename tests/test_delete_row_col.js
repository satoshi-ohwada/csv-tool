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
  this.columns = options.columns ? [...options.columns] : [];
  this.destroy = () => {};
  this.getData = () => this.data;
  this.getRows = () => this.data.map((d, i) => ({
    getData: () => d,
    update: (nd) => { Object.assign(d, nd); },
    delete: () => {
      const idx = this.data.indexOf(d);
      if (idx !== -1) this.data.splice(idx, 1);
    },
    getPosition: () => i + 1,
    getCell: () => ({
      getElement: () => createDummyEl('div'),
      getValue: () => '',
      setValue: () => {}
    })
  }));
  this.addRow = (rowObj, addToTop, targetRow) => {
    if (targetRow) {
      const idx = this.data.indexOf(targetRow.getData ? targetRow.getData() : targetRow);
      if (idx !== -1) {
        const insertIdx = addToTop ? idx : idx + 1;
        this.data.splice(insertIdx, 0, rowObj);
        return Promise.resolve(rowObj);
      }
    }
    if (addToTop) {
      this.data.unshift(rowObj);
    } else {
      this.data.push(rowObj);
    }
    return Promise.resolve(rowObj);
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
        if (idx !== -1) this.columns.splice(idx, 1);
      }
    };
  };
  this.addColumn = (colDef, before, targetCol) => {
    if (targetCol) {
      const idx = this.columns.findIndex(c => c.field === targetCol.getField());
      if (idx !== -1) {
        const insertIdx = before ? idx : idx + 1;
        this.columns.splice(insertIdx, 0, colDef);
        return;
      }
    }
    this.columns.push(colDef);
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

let promptLog = [];
let promptResponse = null;
let confirmLog = [];
let confirmResponse = true;

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
  prompt: (msg, def) => {
    promptLog.push({ msg, def });
    return promptResponse;
  },
  confirm: (msg) => {
    confirmLog.push(msg);
    return confirmResponse;
  },
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
    promptLog = [];
    promptResponse = null;
    confirmLog = [];
    confirmResponse = true;
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
console.log('TEST SUITE: 行・列の指定・確認対話型フロー (Prompt & Confirm)');
console.log('========================================\n');

function setupTable(rowsCount = 3, colsCount = 3) {
  vm.runInContext(`
    (() => {
      const hdrs = [];
      for (let c = 0; c < ${colsCount}; c++) hdrs.push('列 ' + (c + 1));
      const cols = buildTabulatorColumns(hdrs);
      const rows = [];
      for (let r = 0; r < ${rowsCount}; r++) {
        const obj = { _id: r + 1 };
        for (let c = 0; c < ${colsCount}; c++) obj['col_' + c] = 'val_' + r + '_' + c;
        rows.push(obj);
      }
      const t = {
        id: 'tab-test',
        name: 'test.csv',
        headers: hdrs,
        columns: cols,
        data: rows,
        isModified: false
      };
      tabs = [t];
      activeTabId = 'tab-test';
      currentTable = new Tabulator(null, { data: t.data, columns: t.columns });
      activeFocusCell = null;
      selectionRange = null;
    })()
  `, sandbox);
}

// 1. 行追加 (addRowPrompt)
it('addRowPrompt: 位置を指定し確認でOKするとその位置に行が追加される', () => {
  setupTable(3, 2);
  promptResponse = '2'; // 2行目の位置に挿入
  confirmResponse = true;
  vm.runInContext(`addRowPrompt()`, sandbox);

  assert.strictEqual(promptLog.length, 1);
  assert.ok(promptLog[0].msg.includes('追加する行の位置'));
  assert.strictEqual(confirmLog.length, 1);
  assert.ok(confirmLog[0].includes('行 2 の位置に行を追加（挿入）してもよろしいですか？'));

  const tab = vm.runInContext(`tabs.find(t => t.id === activeTabId)`, sandbox);
  assert.strictEqual(tab.data.length, 4);
  assert.strictEqual(tab.isModified, true);
});

it('addRowPrompt: プロンプトでキャンセル(null)した場合は追加されない', () => {
  setupTable(3, 2);
  promptResponse = null;
  vm.runInContext(`addRowPrompt()`, sandbox);

  assert.strictEqual(confirmLog.length, 0);
  const tab = vm.runInContext(`tabs.find(t => t.id === activeTabId)`, sandbox);
  assert.strictEqual(tab.data.length, 3);
});

// 2. 行削除 (deleteRowPrompt)
it('deleteRowPrompt: 行番号を指定し確認でOKするとその行が削除される', () => {
  setupTable(3, 2);
  promptResponse = '2'; // 2行目を削除
  confirmResponse = true;
  vm.runInContext(`deleteRowPrompt()`, sandbox);

  assert.strictEqual(promptLog.length, 1);
  assert.strictEqual(confirmLog.length, 1);
  assert.ok(confirmLog[0].includes('行 2を削除してもよろしいですか？'));

  const tab = vm.runInContext(`tabs.find(t => t.id === activeTabId)`, sandbox);
  assert.strictEqual(tab.data.length, 2);
  assert.strictEqual(tab.isModified, true);
});

it('deleteRowPrompt: 確認ダイアログでキャンセル(false)した場合は削除されない', () => {
  setupTable(3, 2);
  promptResponse = '2';
  confirmResponse = false;
  vm.runInContext(`deleteRowPrompt()`, sandbox);

  assert.strictEqual(confirmLog.length, 1);
  const tab = vm.runInContext(`tabs.find(t => t.id === activeTabId)`, sandbox);
  assert.strictEqual(tab.data.length, 3);
});

it('deleteRowPrompt: 残り1行の時は最低保護でプロンプトすら出ずに中断される', () => {
  setupTable(1, 2);
  vm.runInContext(`deleteRowPrompt()`, sandbox);

  assert.strictEqual(promptLog.length, 0);
  assert.strictEqual(confirmLog.length, 0);
  const tab = vm.runInContext(`tabs.find(t => t.id === activeTabId)`, sandbox);
  assert.strictEqual(tab.data.length, 1);
});

// 3. 列追加 (addColumnPrompt)
it('addColumnPrompt: 列記号を指定し確認でOKすると新しい列が追加される', () => {
  setupTable(2, 3);
  promptResponse = 'B'; // B列の位置に挿入
  confirmResponse = true;
  vm.runInContext(`addColumnPrompt()`, sandbox);

  assert.strictEqual(promptLog.length, 1);
  assert.ok(promptLog[0].msg.includes('追加する列の位置'));
  assert.strictEqual(confirmLog.length, 1);
  assert.ok(confirmLog[0].includes('列 B の位置に新しい列を追加（挿入）してもよろしいですか？'));

  const tab = vm.runInContext(`tabs.find(t => t.id === activeTabId)`, sandbox);
  const activeCols = tab.columns.filter(c => c.field);
  assert.strictEqual(activeCols.length, 4);
  assert.strictEqual(tab.isModified, true);
});

it('addColumnPrompt: 末尾指定で末尾に列が追加される', () => {
  setupTable(2, 3);
  promptResponse = 'D'; // 末尾
  confirmResponse = true;
  vm.runInContext(`addColumnPrompt()`, sandbox);

  assert.strictEqual(confirmLog.length, 1);
  assert.ok(confirmLog[0].includes('末尾（列 D）に新しい列を追加してもよろしいですか？'));

  const tab = vm.runInContext(`tabs.find(t => t.id === activeTabId)`, sandbox);
  const activeCols = tab.columns.filter(c => c.field);
  assert.strictEqual(activeCols.length, 4);
});

// 4. 列削除 (deleteColumnPrompt)
it('deleteColumnPrompt: 列記号(例: B)を指定し確認でOKするとその列が削除される', () => {
  setupTable(2, 3);
  promptResponse = 'B';
  confirmResponse = true;
  vm.runInContext(`deleteColumnPrompt()`, sandbox);

  assert.strictEqual(promptLog.length, 1);
  assert.strictEqual(confirmLog.length, 1);
  assert.ok(confirmLog[0].includes('列 B「列 2」を削除してもよろしいですか？'));

  const tab = vm.runInContext(`tabs.find(t => t.id === activeTabId)`, sandbox);
  const activeCols = tab.columns.filter(c => c.field);
  assert.strictEqual(activeCols.length, 2);
  assert.strictEqual(tab.isModified, true);
});

it('deleteColumnPrompt: 列名(例: 列 2)を指定して削除できる', () => {
  setupTable(2, 3);
  promptResponse = '列 2';
  confirmResponse = true;
  vm.runInContext(`deleteColumnPrompt()`, sandbox);

  assert.strictEqual(confirmLog.length, 1);
  assert.ok(confirmLog[0].includes('列 B「列 2」を削除してもよろしいですか？'));

  const tab = vm.runInContext(`tabs.find(t => t.id === activeTabId)`, sandbox);
  const activeCols = tab.columns.filter(c => c.field);
  assert.strictEqual(activeCols.length, 2);
});

it('deleteColumnPrompt: 残り1列の時は最低保護でプロンプトが出ずに中断される', () => {
  setupTable(2, 1);
  vm.runInContext(`deleteColumnPrompt()`, sandbox);

  assert.strictEqual(promptLog.length, 0);
  assert.strictEqual(confirmLog.length, 0);
  const tab = vm.runInContext(`tabs.find(t => t.id === activeTabId)`, sandbox);
  const activeCols = tab.columns.filter(c => c.field);
  assert.strictEqual(activeCols.length, 1);
});

// 5. 右クリックメニュー deleteRow: 複数行選択時の範囲削除
it('deleteRow: 複数行が範囲選択されている場合に一括削除される', () => {
  setupTable(5, 2);
  vm.runInContext(`selectionRange = { minRow: 1, maxRow: 3, minCol: 0, maxCol: 1 };`, sandbox); // 行2〜行4の3行
  confirmResponse = true;
  vm.runInContext(`
    const rows = currentTable.getRows();
    deleteRow(rows[1]);
  `, sandbox);

  assert.strictEqual(confirmLog.length, 1);
  assert.ok(confirmLog[0].includes('選択中の 行 2 〜 行 4 (3行) を削除してもよろしいですか？'));

  const tab = vm.runInContext(`tabs.find(t => t.id === activeTabId)`, sandbox);
  assert.strictEqual(tab.data.length, 2);
  assert.strictEqual(tab.isModified, true);
});

console.log('\n========================================');
console.log(`TEST SUMMARY: Passed: ${passed}, Failed: ${failed}`);
console.log('========================================');

if (failed > 0) {
  process.exit(1);
} else {
  console.log('ALL PROMPT & CONFIRM TESTS PASSED! 🎉\n');
  process.exit(0);
}
