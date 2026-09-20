const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

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
  querySelector(sel) { return createDummyEl('div'); },
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
  this.setData = (d) => { this.data = d; return Promise.resolve(); };
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
  Tabulator: MockTabulator,
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
console.log('TEST SUITE: データ整形 - 日付・時刻・和暦の統一と変換');
console.log('========================================\n');

const {
  applyCleanTransformation,
  executeDataClean
} = sandbox;

// テスト1: 西暦の各種日付表記を YYYY/MM/DD に統一 (format_date_slash)
it('format_date_slash: スラッシュ・ハイフン・ドット・漢字・8桁・2桁年の正規化', () => {
  assert.strictEqual(applyCleanTransformation('2026/9/1', 'format_date_slash'), '2026/09/01');
  assert.strictEqual(applyCleanTransformation('2026-9-1', 'format_date_slash'), '2026/09/01');
  assert.strictEqual(applyCleanTransformation('2026.9.1', 'format_date_slash'), '2026/09/01');
  assert.strictEqual(applyCleanTransformation('2026年9月1日', 'format_date_slash'), '2026/09/01');
  assert.strictEqual(applyCleanTransformation('20260901', 'format_date_slash'), '2026/09/01');
  assert.strictEqual(applyCleanTransformation('26/9/1', 'format_date_slash'), '2026/09/01');
});

// テスト2: 和暦（元号）を含む日付を西暦 YYYY/MM/DD に統一
it('format_date_slash: 令和・平成・昭和などの和暦を西暦 YYYY/MM/DD に統一', () => {
  assert.strictEqual(applyCleanTransformation('令和8年9月1日', 'format_date_slash'), '2026/09/01');
  assert.strictEqual(applyCleanTransformation('令和元年5月1日', 'format_date_slash'), '2019/05/01');
  assert.strictEqual(applyCleanTransformation('R8.9.1', 'format_date_slash'), '2026/09/01');
  assert.strictEqual(applyCleanTransformation('H30/4/1', 'format_date_slash'), '2018/04/01');
  assert.strictEqual(applyCleanTransformation('S60-1-15', 'format_date_slash'), '1985/01/15');
});

// テスト3: 時刻を含む場合は時刻を維持して桁揃え
it('format_date_slash: 時刻付きデータは時刻を維持して統一 (YYYY/MM/DD HH:mm:ss)', () => {
  assert.strictEqual(applyCleanTransformation('2026/9/1 9:5', 'format_date_slash'), '2026/09/01 09:05:00');
  assert.strictEqual(applyCleanTransformation('2026-9-1 14:30:15', 'format_date_slash'), '2026/09/01 14:30:15');
  assert.strictEqual(applyCleanTransformation('令和8年9月1日 14時30分', 'format_date_slash'), '2026/09/01 14:30:00');
  assert.strictEqual(applyCleanTransformation('R8/9/1 18:00:00', 'format_date_slash'), '2026/09/01 18:00:00');
});

// テスト4: ハイフン区切りへの統一 (format_date_hyphen)
it('format_date_hyphen: YYYY-MM-DD 形式に統一（時刻付きは時刻維持）', () => {
  assert.strictEqual(applyCleanTransformation('2026/9/1', 'format_date_hyphen'), '2026-09-01');
  assert.strictEqual(applyCleanTransformation('令和8年9月1日', 'format_date_hyphen'), '2026-09-01');
  assert.strictEqual(applyCleanTransformation('2026/9/1 14:30', 'format_date_hyphen'), '2026-09-01 14:30:00');
});

// テスト5: 時刻切り捨て (strip_time_to_date)
it('strip_time_to_date: 時刻付きデータから時刻を切り捨てて年月日のみに統一', () => {
  assert.strictEqual(applyCleanTransformation('2026/09/01 14:30:45', 'strip_time_to_date'), '2026/09/01');
  assert.strictEqual(applyCleanTransformation('2026-09-01T09:00:00Z', 'strip_time_to_date'), '2026/09/01');
  assert.strictEqual(applyCleanTransformation('令和8年9月1日 14時30分', 'strip_time_to_date'), '2026/09/01');
  assert.strictEqual(applyCleanTransformation('2026/9/1', 'strip_time_to_date'), '2026/09/01');
});

// テスト6: 和暦テキストを西暦テキストに変換 (wareki_to_seireki)
it('wareki_to_seireki: 和暦漢字・略記を西暦に変換', () => {
  assert.strictEqual(applyCleanTransformation('令和8年9月20日', 'wareki_to_seireki'), '2026年9月20日');
  assert.strictEqual(applyCleanTransformation('令和元年5月1日', 'wareki_to_seireki'), '2019年5月1日');
  assert.strictEqual(applyCleanTransformation('R8.9.20', 'wareki_to_seireki'), '2026.9.20');
  assert.strictEqual(applyCleanTransformation('H30/04/01', 'wareki_to_seireki'), '2018/04/01');
  assert.strictEqual(applyCleanTransformation('昭和60年', 'wareki_to_seireki'), '1985年');
});

// テスト7: 非日付セルの保護（通常テキストや数値が誤って変換されないこと）
it('保護確認: 日付でない通常の文字列や数値は変換されず保持される', () => {
  assert.strictEqual(applyCleanTransformation('りんご', 'format_date_slash'), 'りんご');
  assert.strictEqual(applyCleanTransformation('商品コード12345', 'format_date_slash'), '商品コード12345');
  assert.strictEqual(applyCleanTransformation('12345', 'format_date_slash'), '12345');
  assert.strictEqual(applyCleanTransformation('99999999', 'format_date_slash'), '99999999'); // 月日が不正な8桁数値
  assert.strictEqual(applyCleanTransformation('2026/13/40', 'format_date_slash'), '2026/13/40'); // 暦として不正
});

// テスト8: executeDataClean によるテーブル更新動作
it('executeDataClean("format_date_slash"): テーブル内の混在日付が一括統一される', () => {
  vm.runInContext(`
    (() => {
      const hdrs = ['ID', '登録日時'];
      const cols = buildTabulatorColumns(hdrs);
      const rows = [
        { _id: 1, col_0: '001', col_1: '2026/9/1' },
        { _id: 2, col_0: '002', col_1: '令和8年9月1日 14:30' },
        { _id: 3, col_0: '003', col_1: 'R8.9.1' },
        { _id: 4, col_0: '004', col_1: '通常テキスト' }
      ];
      const t = {
        id: 'tab_date_test',
        name: 'test.csv',
        headers: hdrs,
        columns: cols,
        data: rows,
        isModified: false
      };
      tabs = [t];
      activeTabId = 'tab_date_test';
      currentTable = new Tabulator(null, { data: t.data, columns: t.columns });
      executeDataClean('format_date_slash');
    })();
  `, sandbox);

  const tab = vm.runInContext(`tabs.find(t => t.id === activeTabId)`, sandbox);
  assert.strictEqual(tab.isModified, true);
  assert.strictEqual(tab.data[0].col_1, '2026/09/01');
  assert.strictEqual(tab.data[1].col_1, '2026/09/01 14:30:00');
  assert.strictEqual(tab.data[2].col_1, '2026/09/01');
  assert.strictEqual(tab.data[3].col_1, '通常テキスト'); // 保護される
});

console.log(`\n========================================`);
console.log(`TEST SUMMARY: Passed: ${passed}, Failed: ${failed}`);
console.log('========================================');

if (failed > 0) process.exit(1);
process.exit(0);
