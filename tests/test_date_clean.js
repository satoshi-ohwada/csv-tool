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

// テスト1: 基本ハイフン形式 (format_date_hyphen) - 粒度保持 (日時・日付・年月・時刻のみ・2桁年)
it('format_date_hyphen: 日時(フル)は時刻を維持して YYYY-MM-DD hh:mm:ss に統一', () => {
  assert.strictEqual(applyCleanTransformation('2026/9/1 14:30:15', 'format_date_hyphen'), '2026-09-01 14:30:15');
  assert.strictEqual(applyCleanTransformation('2026-9-1 14:30', 'format_date_hyphen'), '2026-09-01 14:30:00');
  assert.strictEqual(applyCleanTransformation('令和8年9月1日 9:05', 'format_date_hyphen'), '2026-09-01 09:05:00');
  assert.strictEqual(applyCleanTransformation('26/9/1 14:30', 'format_date_hyphen'), '2026-09-01 14:30:00'); // 2桁年
});

it('format_date_hyphen: 年月日(時刻なし)は時刻を付加せず YYYY-MM-DD で終了', () => {
  assert.strictEqual(applyCleanTransformation('2026/9/1', 'format_date_hyphen'), '2026-09-01');
  assert.strictEqual(applyCleanTransformation('2026-9-1', 'format_date_hyphen'), '2026-09-01');
  assert.strictEqual(applyCleanTransformation('令和8年9月1日', 'format_date_hyphen'), '2026-09-01');
  assert.strictEqual(applyCleanTransformation('26/9/1', 'format_date_hyphen'), '2026-09-01'); // 2桁年
  assert.strictEqual(applyCleanTransformation('99-12-31', 'format_date_hyphen'), '1999-12-31'); // 2桁年(1900年代)
});

it('format_date_hyphen: 年月(日なし)は01日を勝手に補完せず YYYY-MM で終了', () => {
  assert.strictEqual(applyCleanTransformation('2026/9', 'format_date_hyphen'), '2026-09');
  assert.strictEqual(applyCleanTransformation('2026-9', 'format_date_hyphen'), '2026-09');
  assert.strictEqual(applyCleanTransformation('2026年9月', 'format_date_hyphen'), '2026-09');
  assert.strictEqual(applyCleanTransformation('令和8年9月', 'format_date_hyphen'), '2026-09');
  assert.strictEqual(applyCleanTransformation('R8.9', 'format_date_hyphen'), '2026-09');
  assert.strictEqual(applyCleanTransformation('26-9', 'format_date_hyphen'), '2026-09'); // 2桁年
});

it('format_date_hyphen: 時刻のみ(日付なし)は年月日は付加せず hh:mm:ss で整形', () => {
  assert.strictEqual(applyCleanTransformation('14:30:15', 'format_date_hyphen'), '14:30:15');
  assert.strictEqual(applyCleanTransformation('14:30', 'format_date_hyphen'), '14:30:00');
  assert.strictEqual(applyCleanTransformation('9:05', 'format_date_hyphen'), '09:05:00');
  assert.strictEqual(applyCleanTransformation('午後2:30', 'format_date_hyphen'), '14:30:00');
  assert.strictEqual(applyCleanTransformation('2:30 PM', 'format_date_hyphen'), '14:30:00');
  assert.strictEqual(applyCleanTransformation('14時30分', 'format_date_hyphen'), '14:30:00');
});

// テスト2: スラッシュ形式 (format_date_slash) - 粒度保持
it('format_date_slash: スラッシュ統一でも粒度（日時・日付・年月・時刻）が維持される', () => {
  assert.strictEqual(applyCleanTransformation('2026-9-1 14:30', 'format_date_slash'), '2026/09/01 14:30:00');
  assert.strictEqual(applyCleanTransformation('2026-9-1', 'format_date_slash'), '2026/09/01');
  assert.strictEqual(applyCleanTransformation('2026-9', 'format_date_slash'), '2026/09');
  assert.strictEqual(applyCleanTransformation('14:30', 'format_date_slash'), '14:30:00');
});

// テスト3: 和暦西暦化 (wareki_to_seireki)
it('wareki_to_seireki: 和暦テキストを粒度維持のまま西暦に置換', () => {
  assert.strictEqual(applyCleanTransformation('令和8年9月20日', 'wareki_to_seireki'), '2026年9月20日');
  assert.strictEqual(applyCleanTransformation('令和8年9月', 'wareki_to_seireki'), '2026年9月');
  assert.strictEqual(applyCleanTransformation('令和元年5月1日', 'wareki_to_seireki'), '2019年5月1日');
  assert.strictEqual(applyCleanTransformation('R8.9.20', 'wareki_to_seireki'), '2026.9.20');
  assert.strictEqual(applyCleanTransformation('H30/04/01', 'wareki_to_seireki'), '2018/04/01');
});

// テスト4: 欠測値表記の空セル統一 (na_to_empty)
it('na_to_empty: NA, null等の欠測表記を空セルに統一し、0や秘匿xは保護する', () => {
  assert.strictEqual(applyCleanTransformation('NA', 'na_to_empty'), '');
  assert.strictEqual(applyCleanTransformation('N/A', 'na_to_empty'), '');
  assert.strictEqual(applyCleanTransformation('NaN', 'na_to_empty'), '');
  assert.strictEqual(applyCleanTransformation('null', 'na_to_empty'), '');
  assert.strictEqual(applyCleanTransformation('NULL', 'na_to_empty'), '');
  assert.strictEqual(applyCleanTransformation('none', 'na_to_empty'), '');
  assert.strictEqual(applyCleanTransformation('#N/A', 'na_to_empty'), '');
  assert.strictEqual(applyCleanTransformation('ND', 'na_to_empty'), '');

  // 保護ガード: 0, 0.0, 秘匿 x は絶対に空セルにしない！
  assert.strictEqual(applyCleanTransformation('0', 'na_to_empty'), '0');
  assert.strictEqual(applyCleanTransformation('0.0', 'na_to_empty'), '0.0');
  assert.strictEqual(applyCleanTransformation('x', 'na_to_empty'), 'x');
  assert.strictEqual(applyCleanTransformation('X', 'na_to_empty'), 'X');
  assert.strictEqual(applyCleanTransformation('リンゴ', 'na_to_empty'), 'リンゴ');
});

// テスト5: 数値末尾の単位記号除去 (strip_units)
it('strip_units: %, 円, kg などの単位記号を除去し、秘匿xや通常文字は保持する', () => {
  assert.strictEqual(applyCleanTransformation('12.5%', 'strip_units'), '12.5');
  assert.strictEqual(applyCleanTransformation('1,200円', 'strip_units'), '1200'); // カンマも除去
  assert.strictEqual(applyCleanTransformation('150kg', 'strip_units'), '150');
  assert.strictEqual(applyCleanTransformation('25.4℃', 'strip_units'), '25.4');
  assert.strictEqual(applyCleanTransformation('50人', 'strip_units'), '50');
  assert.strictEqual(applyCleanTransformation('3000千円', 'strip_units'), '3000');

  // 保護ガード: 秘匿記号 x や単位のみの文字は保持
  assert.strictEqual(applyCleanTransformation('x', 'strip_units'), 'x');
  assert.strictEqual(applyCleanTransformation('円', 'strip_units'), '円');
  assert.strictEqual(applyCleanTransformation('商品A', 'strip_units'), '商品A');
});

// テスト6: 数値の3桁カンマ除去 (remove_commas)
it('remove_commas: 数値中の3桁区切りカンマを除去する', () => {
  assert.strictEqual(applyCleanTransformation('1,234', 'remove_commas'), '1234');
  assert.strictEqual(applyCleanTransformation('1,234,567.89', 'remove_commas'), '1234567.89');
  assert.strictEqual(applyCleanTransformation('東京,大阪', 'remove_commas'), '東京,大阪'); // 文字列内のカンマは保護
});

// テスト7: 非日付セルの保護（通常テキストや不正数値が誤って変換されないこと）
it('保護確認: 日付でない通常の文字列や数値は変換されず保持される', () => {
  assert.strictEqual(applyCleanTransformation('りんご', 'format_date_hyphen'), 'りんご');
  assert.strictEqual(applyCleanTransformation('商品コード12345', 'format_date_hyphen'), '商品コード12345');
  assert.strictEqual(applyCleanTransformation('12345', 'format_date_hyphen'), '12345');
  assert.strictEqual(applyCleanTransformation('99999999', 'format_date_hyphen'), '99999999'); // 月日が不正な8桁数値
  assert.strictEqual(applyCleanTransformation('2026/13/40', 'format_date_hyphen'), '2026/13/40'); // 暦として不正
});

// テスト8: executeDataClean によるテーブル更新動作 (format_date_hyphen)
it('executeDataClean("format_date_hyphen"): テーブル内の混在日付が粒度を保って一括統一される', () => {
  vm.runInContext(`
    (() => {
      const hdrs = ['ID', '登録日時', '測定値'];
      const cols = buildTabulatorColumns(hdrs);
      const rows = [
        { _id: 1, col_0: '001', col_1: '2026/9/1 14:30', col_2: '12.5%' },
        { _id: 2, col_0: '002', col_1: '2026/9/1', col_2: 'NA' },
        { _id: 3, col_0: '003', col_1: '2026/9', col_2: '0' },
        { _id: 4, col_0: '004', col_1: '14:30', col_2: 'x' },
        { _id: 5, col_0: '005', col_1: '通常テキスト', col_2: '1,500円' }
      ];
      const t = {
        id: 'tab_clean_test',
        name: 'test.csv',
        headers: hdrs,
        columns: cols,
        data: rows,
        isModified: false
      };
      tabs = [t];
      activeTabId = 'tab_clean_test';
      currentTable = new Tabulator(null, { data: t.data, columns: t.columns });
      executeDataClean('format_date_hyphen');
    })();
  `, sandbox);

  const tab = vm.runInContext(`tabs.find(t => t.id === activeTabId)`, sandbox);
  assert.strictEqual(tab.isModified, true);
  assert.strictEqual(tab.data[0].col_1, '2026-09-01 14:30:00'); // 日時は時刻維持
  assert.strictEqual(tab.data[1].col_1, '2026-09-01');          // 年月日はDDまで
  assert.strictEqual(tab.data[2].col_1, '2026-09');             // 年月はYYYY-MMまで(日を捏造しない)
  assert.strictEqual(tab.data[3].col_1, '14:30:00');             // 時刻のみは年月日を付加しない
  assert.strictEqual(tab.data[4].col_1, '通常テキスト');         // 通常テキスト保護
});

// テスト9: executeDataClean による欠測値統一 (na_to_empty)
it('executeDataClean("na_to_empty"): NAやnullが空セルに統一され、0や秘匿xは残る', () => {
  vm.runInContext(`
    (() => {
      const hdrs = ['ID', '値'];
      const cols = buildTabulatorColumns(hdrs);
      const rows = [
        { _id: 1, col_0: '001', col_1: 'NA' },
        { _id: 2, col_0: '002', col_1: 'null' },
        { _id: 3, col_0: '003', col_1: '0' },
        { _id: 4, col_0: '004', col_1: 'x' },
        { _id: 5, col_0: '005', col_1: '123' }
      ];
      const t = {
        id: 'tab_na_test',
        name: 'na.csv',
        headers: hdrs,
        columns: cols,
        data: rows,
        isModified: false
      };
      tabs = [t];
      activeTabId = 'tab_na_test';
      currentTable = new Tabulator(null, { data: t.data, columns: t.columns });
      executeDataClean('na_to_empty');
    })();
  `, sandbox);

  const tab = vm.runInContext(`tabs.find(t => t.id === activeTabId)`, sandbox);
  assert.strictEqual(tab.isModified, true);
  assert.strictEqual(tab.data[0].col_1, '');    // NA → 空白
  assert.strictEqual(tab.data[1].col_1, '');    // null → 空白
  assert.strictEqual(tab.data[2].col_1, '0');   // 0は保護
  assert.strictEqual(tab.data[3].col_1, 'x');   // 秘匿xは保護
  assert.strictEqual(tab.data[4].col_1, '123'); // 数値は保護
});

// テスト10: executeDataClean による単位記号除去 (strip_units)
it('executeDataClean("strip_units"): 数値末尾の単位記号とカンマが除去される', () => {
  vm.runInContext(`
    (() => {
      const hdrs = ['割合', '金額'];
      const cols = buildTabulatorColumns(hdrs);
      const rows = [
        { _id: 1, col_0: '12.5%', col_1: '1,200円' },
        { _id: 2, col_0: '100%', col_1: 'x' }
      ];
      const t = {
        id: 'tab_unit_test',
        name: 'units.csv',
        headers: hdrs,
        columns: cols,
        data: rows,
        isModified: false
      };
      tabs = [t];
      activeTabId = 'tab_unit_test';
      currentTable = new Tabulator(null, { data: t.data, columns: t.columns });
      executeDataClean('strip_units');
    })();
  `, sandbox);

  const tab = vm.runInContext(`tabs.find(t => t.id === activeTabId)`, sandbox);
  assert.strictEqual(tab.isModified, true);
  assert.strictEqual(tab.data[0].col_0, '12.5');
  assert.strictEqual(tab.data[0].col_1, '1200');
  assert.strictEqual(tab.data[1].col_0, '100');
  assert.strictEqual(tab.data[1].col_1, 'x'); // 秘匿xは保護
});

console.log(`\n========================================`);
console.log(`TEST SUMMARY: Passed: ${passed}, Failed: ${failed}`);
console.log('========================================');

if (failed > 0) process.exit(1);
process.exit(0);

