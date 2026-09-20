const fs = require('fs');
const vm = require('vm');
const path = require('path');

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
    querySelector(sel) {
      if (sel === 'input[type="checkbox"]') {
        return children.find(c => c.type === 'checkbox') || null;
      }
      return null;
    },
    querySelectorAll(sel) {
      if (sel === '.filter-item') {
        return children.filter(c => c.classList.contains('filter-item'));
      }
      return [];
    },
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
  Tabulator: function() {
    this.setData = () => {};
    this.setColumns = () => {};
    this.on = () => {};
    this.setFilter = (fn) => { this._filterFn = fn; };
    this.clearFilter = () => { this._filterFn = null; };
    this.setSort = () => {};
    this.clearSort = () => {};
    this.getSorters = () => [];
    this.getColumns = () => [];
  }
};
sandbox.window = sandbox;
sandbox.global = sandbox;

// 2. app.js のロード
const appJsPath = path.resolve(__dirname, "../app.js");
const appJsCode = fs.readFileSync(appJsPath, 'utf8');

vm.createContext(sandbox);
vm.runInContext(appJsCode, sandbox);

// 抽出する関数・オブジェクト
const {
  createSmartMatcher,
  excelSmartSorter,
  handleSearch,
  handleFilterSearchInput,
  toggleSearchHelpPopup,
  showSearchHelpPopup,
  hideSearchHelpPopup
} = sandbox;

// 3. テストランナーの実装
let passedCount = 0;
let failedCount = 0;
const failures = [];

function assert(desc, condition, details = '') {
  if (condition) {
    passedCount++;
    console.log(`  ✓ ${desc}`);
  } else {
    failedCount++;
    const msg = `  ✗ ${desc} ${details ? '(' + details + ')' : ''}`;
    console.error(msg);
    failures.push(msg);
  }
}

function runSection(title, fn) {
  console.log(`\n========================================\n[TEST SECTION] ${title}\n========================================`);
  try {
    fn();
  } catch (e) {
    failedCount++;
    console.error(`  ✗ Section Exception: ${e.message}`);
    failures.push(`Section ${title}: ${e.message}`);
  }
}

// 4. テストスイートの実行

runSection('Suite 1: createSmartMatcher - 通常の部分一致・単語境界安全性', () => {
  const m1 = createSmartMatcher('東京');
  assert('通常一致: "東京都港区" に一致', m1('東京都港区') === true);
  assert('通常不一致: "大阪府北区" に不一致', m1('大阪府北区') === false);
  assert('null/undefinedはfalse', m1(null) === false && m1(undefined) === false);

  // 単語境界安全性: "and", "or" を部分文字列に含む英単語が誤って分割されないこと
  const mBrand = createSmartMatcher('brand');
  assert('単語 brand: "brand new" に一致', mBrand('brand new') === true);
  assert('単語 brand: "no b r a n d" に不一致', mBrand('no b r a n d') === false);

  const mOrange = createSmartMatcher('orange');
  assert('単語 orange: "orange juice" に一致', mOrange('orange juice') === true);
  assert('単語 orange: "apple juice" に不一致', mOrange('apple juice') === false);

  const mOrder = createSmartMatcher('order');
  assert('単語 order: "order status" に一致', mOrder('order status') === true);
  assert('単語 order: "delivery status" に不一致', mOrder('delivery status') === false);

  const mStand = createSmartMatcher('stand');
  assert('単語 stand: "standard edition" に一致', mStand('standard edition') === true);
});

runSection('Suite 2: createSmartMatcher - AND 検索 (and, AND, 全角スペース)', () => {
  const mAnd1 = createSmartMatcher('東京 and 営業');
  assert('A and B: "東京都 営業部" に一致 (両方含む)', mAnd1('東京都 営業部') === true);
  assert('A and B: "東京都 総務部" に不一致 (片方のみ)', mAnd1('東京都 総務部') === false);
  assert('A and B: "大阪府 営業部" に不一致 (片方のみ)', mAnd1('大阪府 営業部') === false);
  assert('A and B: "北海道 開発部" に不一致 (両方なし)', mAnd1('北海道 開発部') === false);

  const mAndUpper = createSmartMatcher('ノート AND パソコン');
  assert('A AND B: "ノートパソコン (Core i7)" に一致', mAndUpper('ノートパソコン (Core i7)') === true);
  assert('A AND B: "デスクトップパソコン" に不一致', mAndUpper('デスクトップパソコン') === false);

  const mAndWide = createSmartMatcher('東京　and　開発');
  assert('全角スペース A　and　B: "東京都 開発部" に一致', mAndWide('東京都 開発部') === true);

  const mAnd3 = createSmartMatcher('東京 and 営業 and 主任');
  assert('3要素AND: "東京都 営業部 主任" に一致', mAnd3('東京都 営業部 主任') === true);
  assert('3要素AND: "東京都 営業部 課長" に不一致', mAnd3('東京都 営業部 課長') === false);
});

runSection('Suite 3: createSmartMatcher - OR 検索 (or, OR, カンマ, 読点, パイプ)', () => {
  const mOr1 = createSmartMatcher('ノート or デスクトップ');
  assert('A or B: "ノートPC" に一致', mOr1('ノートPC') === true);
  assert('A or B: "デスクトップPC" に一致', mOr1('デスクトップPC') === true);
  assert('A or B: "タブレット" に不一致', mOr1('タブレット') === false);

  const mOrUpper = createSmartMatcher('東京 OR 大阪');
  assert('A OR B: "東京都" に一致', mOrUpper('東京都') === true);
  assert('A OR B: "大阪府" に一致', mOrUpper('大阪府') === true);
  assert('A OR B: "名古屋市" に不一致', mOrUpper('名古屋市') === false);

  const mComma = createSmartMatcher('ノート, デスクトップ');
  assert('カンマ区切り: "ノートPC" に一致', mComma('ノートPC') === true);
  assert('カンマ区切り: "デスクトップPC" に一致', mComma('デスクトップPC') === true);
  assert('カンマ区切り: "スマホ" に不一致', mComma('スマホ') === false);

  const mTouten = createSmartMatcher('東京、大阪');
  assert('読点区切り: "東京都" に一致', mTouten('東京都') === true);
  assert('読点区切り: "大阪府" に一致', mTouten('大阪府') === true);

  const mPipe = createSmartMatcher('ノート|デスクトップ');
  assert('パイプ記号: "ノートPC" に一致', mPipe('ノートPC') === true);
  assert('パイプ記号: "デスクトップPC" に一致', mPipe('デスクトップPC') === true);
  assert('パイプ記号: "スマホ" に不一致', mPipe('スマホ') === false);

  const mOr3 = createSmartMatcher('東京 or 大阪 or 名古屋');
  assert('3要素OR: "名古屋支店" に一致', mOr3('名古屋支店') === true);
  assert('3要素OR: "福岡支店" に不一致', mOr3('福岡支店') === false);
});

runSection('Suite 4: createSmartMatcher - 複合検索 (AND + OR)', () => {
  const mMixed1 = createSmartMatcher('東京 and 営業 or 大阪 and 総務');
  assert('複合: "東京都 営業部" に一致 (前者のAND充足)', mMixed1('東京都 営業部') === true);
  assert('複合: "大阪府 総務部" に一致 (後者のAND充足)', mMixed1('大阪府 総務部') === true);
  assert('複合: "東京都 総務部" に不一致 (どちらのANDも満たさない)', mMixed1('東京都 総務部') === false);
  assert('複合: "大阪府 営業部" に不一致', mMixed1('大阪府 営業部') === false);

  const mMixed2 = createSmartMatcher('ノート or デスクトップ and Dell');
  assert('複合: "ノートPC" に一致', mMixed2('ノートPC') === true);
  assert('複合: "デスクトップ Dell" に一致', mMixed2('デスクトップ Dell') === true);
  assert('複合: "デスクトップ HP" に不一致', mMixed2('デスクトップ HP') === false);
});

runSection('Suite 5: createSmartMatcher - 数値条件・範囲条件', () => {
  const mGte = createSmartMatcher('>=300');
  assert('>=300: 300 に一致', mGte('300') === true);
  assert('>=300: 301 に一致', mGte('301') === true);
  assert('>=300: 299 に不一致', mGte('299') === false);
  assert('>=300: "1,200" に一致 (カンマ付き数値)', mGte('1,200') === true);

  const mLte = createSmartMatcher('<=100');
  assert('<=100: 100 に一致', mLte('100') === true);
  assert('<=100: 50 に一致', mLte('50') === true);
  assert('<=100: 101 に不一致', mLte('101') === false);

  const mRange = createSmartMatcher('100..200');
  assert('100..200: 150 に一致', mRange('150') === true);
  assert('100..200: 100 に一致', mRange('100') === true);
  assert('100..200: 200 に一致', mRange('200') === true);
  assert('100..200: 99 に不一致', mRange('99') === false);
  assert('100..200: 201 に不一致', mRange('201') === false);

  const mJp = createSmartMatcher('100以上');
  assert('100以上: 100 に一致', mJp('100') === true);
  assert('100以上: 99 に不一致', mJp('99') === false);
});

runSection('Suite 6: createSmartMatcher - ワイルドカード・正規表現', () => {
  const mWild = createSmartMatcher('*phone');
  assert('*phone: "smartphone" に一致', mWild('smartphone') === true);
  assert('*phone: "tablet" に不一致', mWild('tablet') === false);

  const mRegexSlash = createSmartMatcher('/^item-\\d+$/i');
  assert('/^item-\\d+$/i: "ITEM-123" に一致', mRegexSlash('ITEM-123') === true);
  assert('/^item-\\d+$/i: "ITEM-ABC" に不一致', mRegexSlash('ITEM-ABC') === false);
});

runSection('Suite 7: handleSearch - クイック検索バー（テーブル全体フィルター）', () => {
  // テスト用データ行
  const sampleRows = [
    { col_0: '1', col_1: 'ノートパソコン Pro', col_2: '128,000', col_3: '営業部', col_4: '東京都港区' },
    { col_0: '2', col_1: 'ノートパソコン Air', col_2: '98,000', col_3: '開発部', col_4: '東京都渋谷区' },
    { col_0: '3', col_1: 'デスクトップ ワークステーション', col_2: '250,000', col_3: '開発部', col_4: '大阪府北区' },
    { col_0: '4', col_1: '液晶モニター 27インチ', col_2: '35,000', col_3: '総務部', col_4: '東京都港区' },
    { col_0: '5', col_1: 'タブレット 10インチ', col_2: '45,000', col_3: '営業部', col_4: '大阪府中央区' },
    { col_0: '6', col_1: 'スマートフォン', col_2: '80,000', col_3: '総務部', col_4: '名古屋市中区' }
  ];

  // コンテキスト内での初期化
  vm.runInContext(`
    currentTable = new Tabulator();
    tabs = [{
      id: 'tab-1',
      columns: [
        { field: 'col_0', title: 'ID' },
        { field: 'col_1', title: '商品名' },
        { field: 'col_2', title: '単価' },
        { field: 'col_3', title: '担当部署' },
        { field: 'col_4', title: '所在地' }
      ]
    }];
    activeTabId = 'tab-1';
    activeColumnFilters.clear();
  `, sandbox);

  function testQuery(query) {
    vm.runInContext(`
      searchInput.value = ${JSON.stringify(query)};
      handleSearch();
    `, sandbox);
    const filterFn = vm.runInContext('currentTable._filterFn', sandbox);
    if (!filterFn) return sampleRows;
    return sampleRows.filter(r => filterFn(r));
  }

  // テスト 7.1: スペース区切り AND
  const r1 = testQuery('東京 営業');
  assert('スペースAND "東京 営業": 1件ヒット (行1のみ)', r1.length === 1 && r1[0].col_0 === '1');

  // テスト 7.2: "and" 区切り AND
  const r2 = testQuery('東京 and 営業');
  assert('"and"区切りAND "東京 and 営業": 1件ヒット (行1のみ)', r2.length === 1 && r2[0].col_0 === '1');

  // テスト 7.3: "or" 区切り OR
  const r3 = testQuery('大阪 or 名古屋');
  assert('"or"区切りOR "大阪 or 名古屋": 3件ヒット (行3, 行5, 行6)', r3.length === 3 && r3.map(r => r.col_0).sort().join(',') === '3,5,6');

  // テスト 7.4: カンマ区切り OR
  const r4 = testQuery('モニター, タブレット');
  assert('カンマ区切りOR: 2件ヒット (行4, 行5)', r4.length === 2 && r4.map(r => r.col_0).sort().join(',') === '4,5');

  // テスト 7.5: 列指定検索
  const r5 = testQuery('商品名:ノート');
  assert('列指定 "商品名:ノート": 2件ヒット (行1, 行2)', r5.length === 2 && r5.every(r => r.col_1.includes('ノート')));

  // テスト 7.6: 列指定 + 数値条件
  const r6 = testQuery('単価:>=100000');
  assert('列指定数値 "単価:>=100000": 2件ヒット (128,000と250,000)', r6.length === 2 && r6.map(r => r.col_0).sort().join(',') === '1,3');

  // テスト 7.7: 列指定 + AND
  const r7 = testQuery('所在地:東京 and 担当部署:開発');
  assert('列指定AND "所在地:東京 and 担当部署:開発": 1件ヒット (行2)', r7.length === 1 && r7[0].col_0 === '2');

  // テスト 7.8: 列指定 + OR
  const r8 = testQuery('商品名:モニター or 商品名:スマートフォン');
  assert('列指定OR "商品名:モニター or 商品名:スマートフォン": 2件ヒット (行4, 行6)', r8.length === 2 && r8.map(r => r.col_0).sort().join(',') === '4,6');

  // テスト 7.9: 複合検索 (節1 AND または 節2 AND)
  const r9 = testQuery('所在地:東京 担当部署:営業 or 所在地:大阪 担当部署:開発');
  assert('複合検索 "所在地:東京 営業 or 所在地:大阪 開発": 2件ヒット (行1, 行3)', r9.length === 2 && r9.map(r => r.col_0).sort().join(',') === '1,3');

  // テスト 7.10: 括弧 () による優先順位グループ化
  const r10 = testQuery('(東京 or 大阪) 営業');
  assert('括弧グループ化 "(東京 or 大阪) 営業": 2件ヒット (行1, 行5のみ)', r10.length === 2 && r10.map(r => r.col_0).sort().join(',') === '1,5');

  // テスト 7.11: 括弧と列指定・数値条件の組み合わせ
  const r11 = testQuery('(商品名:ノート or 商品名:スマートフォン) 単価:>=80000');
  assert('括弧+列指定+数値 "(商品名:ノート or 商品名:スマートフォン) 単価:>=80000": 3件ヒット (行1, 行2, 行6)', r11.length === 3 && r11.map(r => r.col_0).sort().join(',') === '1,2,6');

  // テスト 7.12: 全角括弧 （） にも対応
  const r12 = testQuery('（東京 or 大阪） 営業');
  assert('全角括弧 "（東京 or 大阪） 営業": 2件ヒット (行1, 行5のみ)', r12.length === 2 && r12.map(r => r.col_0).sort().join(',') === '1,5');
});

runSection('Suite 8: handleFilterSearchInput - 列フィルターポップアップ内の候補選択', () => {
  // DOM のセットアップ
  const filterSearchBox = sandbox.document.getElementById('filter-search-box');
  const filterItemsList = sandbox.document.getElementById('filter-items-list');
  sandbox.filterSearchBox = filterSearchBox;
  sandbox.filterItemsList = filterItemsList;
  sandbox.currentFilterSearchMode = 'auto';

  // 候補アイテム（DOM要素）を作成
  const candidateValues = ['ノートパソコン Pro', 'ノートパソコン Air', 'デスクトップ ワークステーション', '液晶モニター', 'タブレット', 'スマートフォン'];
  const items = candidateValues.map(v => {
    const itemEl = createDummyEl('div');
    itemEl.classList.add('filter-item');
    itemEl.dataset = { value: v };

    const textEl = createDummyEl('span');
    textEl.classList.add('filter-item-text');
    textEl.textContent = v;
    itemEl.appendChild(textEl);

    const cb = createDummyEl('input');
    cb.type = 'checkbox';
    cb.checked = true;
    itemEl.appendChild(cb);

    // itemEl.querySelector のモック対応
    itemEl.querySelector = (sel) => {
      if (sel === 'input[type="checkbox"]') return cb;
      if (sel === '.filter-item-text') return textEl;
      return null;
    };

    filterItemsList.appendChild(itemEl);
    return { itemEl, cb, val: v };
  });

  // テスト 8.1: 'ノート or タブレット'
  filterSearchBox.value = 'ノート or タブレット';
  handleFilterSearchInput();

  const checked1 = items.filter(it => it.cb.checked).map(it => it.val);
  assert('列フィルターOR検索 "ノート or タブレット": 3件チェックON',
    checked1.length === 3 &&
    checked1.includes('ノートパソコン Pro') &&
    checked1.includes('ノートパソコン Air') &&
    checked1.includes('タブレット')
  );

  // テスト 8.2: 検索窓クリアで全件復帰
  filterSearchBox.value = '';
  handleFilterSearchInput();
  const checked2 = items.filter(it => it.cb.checked).map(it => it.val);
  assert('検索窓クリア時: 全件(6件)チェックONに復帰', checked2.length === candidateValues.length);

  // テスト 8.3: 数値候補に対する数値条件
  // アイテムリストをクリア
  while (filterItemsList.querySelectorAll('.filter-item').length > 0) {
    const list = filterItemsList.querySelectorAll('.filter-item');
    list.forEach(el => filterItemsList.removeChild(el));
  }
  const numCandidateValues = ['50', '100', '300', '1,200', '128,000', '250,000'];
  const numItems = numCandidateValues.map(v => {
    const itemEl = createDummyEl('div');
    itemEl.classList.add('filter-item');
    itemEl.dataset = { value: v };

    const textEl = createDummyEl('span');
    textEl.classList.add('filter-item-text');
    textEl.textContent = v;
    itemEl.appendChild(textEl);

    const cb = createDummyEl('input');
    cb.type = 'checkbox';
    cb.checked = true;
    itemEl.appendChild(cb);

    itemEl.querySelector = (sel) => {
      if (sel === 'input[type="checkbox"]') return cb;
      if (sel === '.filter-item-text') return textEl;
      return null;
    };

    filterItemsList.appendChild(itemEl);
    return { itemEl, cb, val: v };
  });

  filterSearchBox.value = '>=300';
  handleFilterSearchInput();
  const checkedNum = numItems.filter(it => it.cb.checked).map(it => it.val);
  assert('列フィルター数値条件 ">=300": 300以上の候補(300, 1,200, 128,000, 250,000)すべてチェックON',
    checkedNum.length === 4 &&
    checkedNum.includes('300') &&
    checkedNum.includes('1,200') &&
    checkedNum.includes('128,000') &&
    checkedNum.includes('250,000') &&
    !checkedNum.includes('50') &&
    !checkedNum.includes('100')
  );
});

runSection('Suite 9: excelSmartSorter - 検索後の並べ替え連携', () => {
  // Tabulator のソートシミュレーション
  function tabSort(arr, dir) {
    return [...arr].sort((e, t) => {
      const a = dir === 'asc' ? e : t;
      const l = dir === 'asc' ? t : e;
      return excelSmartSorter(a, l, null, null, null, dir);
    });
  }

  const dataset = ['1,200', '300', '128,000', '¥500', '220,000', '', '10', 'abc', '(空白)'];

  const ascResult = tabSort(dataset, 'asc');
  assert('昇順ソート: 数値が小から大に正しくソートされる',
    ascResult.slice(0, 6).join(',') === '10,300,¥500,1,200,128,000,220,000'
  );
  assert('昇順ソート: 文字列が数値の後、空白が末尾',
    ascResult[6] === 'abc' && (ascResult[7] === '' || ascResult[7] === '(空白)')
  );

  const descResult = tabSort(dataset, 'desc');
  assert('降順ソート: 文字列が先頭、数値が大から小にソートされる',
    descResult[0] === 'abc' &&
    descResult.slice(1, 7).join(',') === '220,000,128,000,1,200,¥500,300,10'
  );
  assert('降順ソート: 空白が末尾に配置される',
    (descResult[7] === '' || descResult[7] === '(空白)') && (descResult[8] === '' || descResult[8] === '(空白)')
  );
});

runSection('機能: 検索・絞り込みガイドポップアップの表示・非表示', () => {
  const popup = sandbox.document.getElementById('search-help-popup');

  hideSearchHelpPopup();
  assert('hideSearchHelpPopup: style.display が none になる', popup.style.display === 'none');

  showSearchHelpPopup();
  assert('showSearchHelpPopup: style.display が block になる', popup.style.display === 'block');

  toggleSearchHelpPopup();
  assert('toggleSearchHelpPopup: 開いている状態でトグルすると none になる', popup.style.display === 'none');

  toggleSearchHelpPopup();
  assert('toggleSearchHelpPopup: 閉じている状態でトグルすると block になる', popup.style.display === 'block');

  hideSearchHelpPopup();
  assert('hideSearchHelpPopup: 再び none になる', popup.style.display === 'none');
});

// 5. 最終サマリー表示
console.log(`\n========================================`);
console.log(`TEST SUMMARY:`);
console.log(`  Passed: ${passedCount}`);
console.log(`  Failed: ${failedCount}`);
console.log(`========================================`);

if (failedCount > 0) {
  console.error('\nFAILURES:');
  failures.forEach(f => console.error(f));
  process.exit(1);
} else {
  console.log('\nALL TESTS PASSED SUCCESSFULLY! 🎉');
  process.exit(0);
}
