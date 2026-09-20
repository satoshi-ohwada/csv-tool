// CSV編集ツール

let tabs = [];
let activeTabId = null;
let currentTable = null;
let tabCounter = 1;

// DOM要素の取得（初期化時に確実に取得）
let tabsContainer, btnAddTab, fileInput, btnOpen, btnNew, btnSave;
let btnUndo, btnRedo, btnAddRow, btnAddCol, btnDelRow, btnDelCol;
let btnRemoveComma, btnAddComma, selectCommaMode;
let searchInput, btnClearSearch, selectEncoding, selectDelimiter, selectNewline;
let tableArea, emptyState, dropOverlay;
let statusFilename, statusModified, statusRows, statusCols, statusEncoding, statusDelimiter, statusNewline;
let saveModal, btnModalClose, btnModalOverwrite, btnModalSaveAs, btnModalSaveNew, btnModalCancel, saveModalExistingActions, saveModalNewActions, toastContainer;

// 新機能用 DOM要素
let btnCleanMenu, cleanDropdownMenu, dropdownCleanContainer;
let btnFindReplace, findReplaceModal, btnFindClose, findKeyword, replaceKeyword, findOptCase, findOptExact, findOptRegex;
let btnFindNext, btnReplaceOne, btnReplaceAll, btnFindCancel;
let filterActiveBadge, filterPopup, filterSearchBox, filterItemsList, btnFilterSelectAll, btnFilterClearAll, btnFilterApply, btnFilterClose, btnFilterPopupClose, btnFilterClearColumn;
let filterAddSelectionRow, filterAddToSelection, btnFilterSearchClear, filterSearchPreviousQuery = '';
let btnToggleNumRange, filterNumRangePanel, filterNumMin, filterNumMax, btnFilterNumApply, btnFilterNumClear;
let statusCalcArea, calcAvg, calcCount, calcSum;
let btnSearchHelp, searchHelpPopup, btnSearchHelpClose;

// 行コピー用クリップボードバッファ
let rowClipboardBuffer = null;

// 範囲選択・コピペ・セルフォーカス管理 (Excel風)
let activeFocusCell = null;      // { rowIndex: number, colIndex: number } 現在のアクティブセル
let selectionAnchor = null;      // { rowIndex: number, colIndex: number } 範囲選択の起点セル
let selectionRange = null;       // { minRow, maxRow, minCol, maxCol, startRow, startCol, endRow, endCol } 選択範囲
let copiedRange = null;          // { minRow, maxRow, minCol, maxCol } コピー中の範囲 (緑点線枠用)
let internalClipboardText = '';  // ブラウザ権限制限対策の内部クリップボードTSV文字列
let isMouseDraggingRange = false;// マウスドラッグで範囲選択中か

// 入力モード管理
let currentEditingCell = null;   // 現在編集中（入力モード）の CellComponent
let editingOriginalValue = '';   // 編集前の元の値（ESCキャンセル用）

// 列幅自動計算用 Canvas 2D コンテキスト
let textMeasureCtx = null;

// 初期化
function init() {
  console.log("CSV編集ツール initializing...");
  getDOMElements();
  setupEventListeners();
  setupGlobalDragAndDrop();

  // 文字幅計測用キャンバスの初期化
  try {
    const canvas = document.createElement('canvas');
    textMeasureCtx = canvas.getContext('2d');
    textMeasureCtx.font = '13px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", sans-serif';
  } catch (e) {
    console.warn("Canvas measureText not available:", e);
  }

  // 初期状態で空のタブを作成
  createNewTab("無題-1.csv");
  console.log("CSV編集ツール initialized successfully.");
}

function getDOMElements() {
  tabsContainer = document.getElementById('tabs-container');
  btnAddTab = document.getElementById('btn-add-tab');
  fileInput = document.getElementById('file-input');
  btnOpen = document.getElementById('btn-open');
  btnNew = document.getElementById('btn-new');
  btnSave = document.getElementById('btn-save');
  btnUndo = document.getElementById('btn-undo');
  btnRedo = document.getElementById('btn-redo');
  btnAddRow = document.getElementById('btn-add-row');
  btnAddCol = document.getElementById('btn-add-col');
  btnDelRow = document.getElementById('btn-del-row');
  btnDelCol = document.getElementById('btn-del-col');
  btnRemoveComma = document.getElementById('btn-remove-comma');
  btnAddComma = document.getElementById('btn-add-comma');
  selectCommaMode = document.getElementById('select-comma-mode');
  searchInput = document.getElementById('search-input');
  btnClearSearch = document.getElementById('btn-clear-search');
  selectEncoding = document.getElementById('select-encoding');
  selectDelimiter = document.getElementById('select-delimiter');
  selectNewline = document.getElementById('select-newline');
  tableArea = document.getElementById('table-area');
  emptyState = document.getElementById('empty-state');
  dropOverlay = document.getElementById('drop-overlay');

  statusFilename = document.getElementById('status-filename');
  statusModified = document.getElementById('status-modified');
  statusRows = document.getElementById('status-rows');
  statusCols = document.getElementById('status-cols');
  statusEncoding = document.getElementById('status-encoding');
  statusDelimiter = document.getElementById('status-delimiter');
  statusNewline = document.getElementById('status-newline');

  saveModal = document.getElementById('save-modal');
  btnModalClose = document.getElementById('btn-modal-close');
  btnModalOverwrite = document.getElementById('btn-modal-overwrite');
  btnModalSaveAs = document.getElementById('btn-modal-save-as');
  btnModalSaveNew = document.getElementById('btn-modal-save-new');
  btnModalCancel = document.getElementById('btn-modal-cancel');
  saveModalExistingActions = document.getElementById('save-modal-existing-actions');
  saveModalNewActions = document.getElementById('save-modal-new-actions');
  toastContainer = document.getElementById('toast-container');

  // データ整形
  btnCleanMenu = document.getElementById('btn-clean-menu');
  cleanDropdownMenu = document.getElementById('clean-dropdown-menu');
  dropdownCleanContainer = document.getElementById('dropdown-clean-container');

  // 検索・置換
  btnFindReplace = document.getElementById('btn-find-replace');
  findReplaceModal = document.getElementById('find-replace-modal');
  btnFindClose = document.getElementById('btn-find-close');
  findKeyword = document.getElementById('find-keyword');
  replaceKeyword = document.getElementById('replace-keyword');
  findOptCase = document.getElementById('find-opt-case');
  findOptExact = document.getElementById('find-opt-exact');
  findOptRegex = document.getElementById('find-opt-regex');
  btnFindNext = document.getElementById('btn-find-next');
  btnReplaceOne = document.getElementById('btn-replace-one');
  btnReplaceAll = document.getElementById('btn-replace-all');
  btnFindCancel = document.getElementById('btn-find-cancel');

  // 検索・絞り込みガイド
  btnSearchHelp = document.getElementById('btn-search-help');
  searchHelpPopup = document.getElementById('search-help-popup');
  btnSearchHelpClose = document.getElementById('btn-search-help-close');

  // オートフィルター
  filterActiveBadge = document.getElementById('filter-active-badge');
  filterPopup = document.getElementById('filter-popup');
  filterSearchBox = document.getElementById('filter-search-box');
  filterItemsList = document.getElementById('filter-items-list');
  btnFilterSelectAll = document.getElementById('btn-filter-select-all');
  btnFilterClearAll = document.getElementById('btn-filter-clear-all');
  btnFilterApply = document.getElementById('btn-filter-apply');
  btnFilterClose = document.getElementById('btn-filter-close');
  btnFilterPopupClose = document.getElementById('btn-filter-popup-close');
  btnFilterClearColumn = document.getElementById('btn-filter-clear-column');
  btnToggleNumRange = document.getElementById('btn-toggle-num-range');
  filterNumRangePanel = document.getElementById('filter-num-range-panel');
  filterNumMin = document.getElementById('filter-num-min');
  filterNumMax = document.getElementById('filter-num-max');
  btnFilterNumApply = document.getElementById('btn-filter-num-apply');
  btnFilterNumClear = document.getElementById('btn-filter-num-clear');
  filterAddSelectionRow = document.getElementById('filter-add-selection-row');
  filterAddToSelection = document.getElementById('filter-add-to-selection');
  btnFilterSearchClear = document.getElementById('btn-filter-search-clear');

  // ステータスバー簡易集計
  statusCalcArea = document.getElementById('status-calc-area');
  calcAvg = document.getElementById('calc-avg');
  calcCount = document.getElementById('calc-count');
  calcSum = document.getElementById('calc-sum');
}

function setupEventListeners() {
  fileInput.addEventListener('change', handleFileSelect);

  // ファイルを開く（File System Access API 対応）
  if (btnOpen) {
    btnOpen.addEventListener('click', (e) => {
      e.preventDefault();
      openFileWithPicker();
    });
  }

  // 空状態のファイル選択ボタン
  const btnEmptyOpen = document.getElementById('btn-empty-open');
  if (btnEmptyOpen) {
    btnEmptyOpen.addEventListener('click', (e) => {
      e.preventDefault();
      openFileWithPicker();
    });
  }

  btnNew.addEventListener('click', () => createNewTab());
  btnAddTab.addEventListener('click', () => createNewTab());
  
  // 保存ボタン（新規なら新規保存、既存なら上書き/別名選択）
  if (btnSave) btnSave.addEventListener('click', () => handleSave());

  // 保存モーダルのイベント
  if (btnModalClose) btnModalClose.addEventListener('click', hideSaveModal);
  if (btnModalCancel) btnModalCancel.addEventListener('click', hideSaveModal);
  if (btnModalOverwrite) {
    btnModalOverwrite.addEventListener('click', () => {
      applySaveSettingsToCurrentTab();
      hideSaveModal();
      saveOverwriteCurrentTab();
    });
  }
  if (btnModalSaveAs) {
    btnModalSaveAs.addEventListener('click', () => {
      applySaveSettingsToCurrentTab();
      hideSaveModal();
      saveAsCurrentTab();
    });
  }
  if (btnModalSaveNew) {
    btnModalSaveNew.addEventListener('click', () => {
      applySaveSettingsToCurrentTab();
      hideSaveModal();
      saveAsCurrentTab();
    });
  }
  if (saveModal) {
    saveModal.addEventListener('click', (e) => {
      if (e.target === saveModal) hideSaveModal();
    });
  }

  btnUndo.addEventListener('click', () => {
    if (currentTable) currentTable.undo();
  });

  btnRedo.addEventListener('click', () => {
    if (currentTable) currentTable.redo();
  });

  btnAddRow.addEventListener('click', addRowPrompt);
  btnAddCol.addEventListener('click', addColumnPrompt);
  btnDelRow.addEventListener('click', deleteRowPrompt);
  if (btnDelCol) btnDelCol.addEventListener('click', deleteColumnPrompt);

  if (btnRemoveComma) btnRemoveComma.addEventListener('click', removeCommasFromTable);
  if (btnAddComma) btnAddComma.addEventListener('click', addCommasToTable);

  // データ整形ドロップダウン
  if (btnCleanMenu) {
    btnCleanMenu.addEventListener('click', (e) => {
      e.stopPropagation();
      toggleCleanDropdown();
    });
  }
  if (cleanDropdownMenu) {
    cleanDropdownMenu.addEventListener('click', (e) => {
      const item = e.target.closest('.dropdown-item');
      if (item) {
        if (item.classList.contains('submenu-trigger')) {
          e.stopPropagation();
          item.blur();
          return;
        }
        const action = item.getAttribute('data-action');
        if (action) {
          executeDataClean(action);
          hideCleanDropdown();
        }
      }
    });

    // マウスホバーでサブメニューを即座に切り替え（前のサブメニューが残るのを防止）
    const submenuItems = document.querySelectorAll('.dropdown-submenu-item');
    if (submenuItems && submenuItems.forEach) {
      submenuItems.forEach(item => {
        if (item && item.addEventListener) {
          item.addEventListener('mouseenter', () => {
            submenuItems.forEach(other => {
              if (other && other.classList) other.classList.remove('submenu-active');
            });
            if (item && item.classList) item.classList.add('submenu-active');
          });
        }
      });
    }
  }
  window.addEventListener('click', (e) => {
    if (cleanDropdownMenu && !cleanDropdownMenu.contains(e.target) && e.target !== btnCleanMenu) {
      hideCleanDropdown();
    }
  });

  // 列データ診断サマリーモーダル
  const btnSummaryClose = document.getElementById('btn-summary-close');
  const btnSummaryCloseX = document.getElementById('btn-summary-close-x');
  const modalColumnSummary = document.getElementById('modal-column-summary');
  if (btnSummaryClose) btnSummaryClose.addEventListener('click', hideColumnSummaryModal);
  if (btnSummaryCloseX) btnSummaryCloseX.addEventListener('click', hideColumnSummaryModal);
  if (modalColumnSummary) {
    modalColumnSummary.addEventListener('click', (e) => {
      if (e.target === modalColumnSummary) hideColumnSummaryModal();
    });
  }

  // 検索・置換モーダル
  if (btnFindReplace) {
    btnFindReplace.addEventListener('click', () => openFindReplaceModal('find'));
  }
  if (btnFindClose) btnFindClose.addEventListener('click', closeFindReplaceModal);
  if (btnFindCancel) btnFindCancel.addEventListener('click', closeFindReplaceModal);
  if (btnFindNext) btnFindNext.addEventListener('click', () => executeFindNext());
  if (btnReplaceOne) btnReplaceOne.addEventListener('click', () => executeReplaceOne());
  if (btnReplaceAll) btnReplaceAll.addEventListener('click', () => executeReplaceAll());
  if (findKeyword) {
    findKeyword.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        executeFindNext();
      }
    });
  }

  // 検索・絞り込みガイドポップアップ
  if (btnSearchHelp) {
    btnSearchHelp.addEventListener('click', (e) => {
      e.stopPropagation();
      toggleSearchHelpPopup();
    });
  }
  if (btnSearchHelpClose) {
    btnSearchHelpClose.addEventListener('click', hideSearchHelpPopup);
  }

  // オートフィルターポップアップ
  if (filterActiveBadge) {
    filterActiveBadge.addEventListener('click', clearColumnFilter);
  }
  if (btnFilterClose) btnFilterClose.addEventListener('click', hideFilterPopup);
  if (btnFilterPopupClose) btnFilterPopupClose.addEventListener('click', hideFilterPopup);
  if (btnFilterClearColumn) btnFilterClearColumn.addEventListener('click', clearCurrentColumnFilter);
  if (btnFilterSelectAll) btnFilterSelectAll.addEventListener('click', filterPopupSelectAll);
  if (btnFilterClearAll) btnFilterClearAll.addEventListener('click', filterPopupClearAll);
  if (btnFilterApply) btnFilterApply.addEventListener('click', applyColumnFilter);
  if (filterSearchBox) {
    filterSearchBox.addEventListener('input', handleFilterSearchInput);
    filterSearchBox.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        applyColumnFilter();
      } else if (e.key === 'Escape') {
        e.preventDefault();
        hideFilterPopup();
      }
    });
  }
  if (filterAddToSelection) {
    filterAddToSelection.addEventListener('change', () => {
      handleFilterSearchInput();
    });
  }
  if (btnFilterSearchClear) {
    btnFilterSearchClear.addEventListener('click', (e) => {
      e.stopPropagation();
      clearFilterSearch();
    });
  }

  // フィルター検索モードボタン切替 (*?, .*, 自動)
  document.querySelectorAll('.filter-mode-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const mode = btn.dataset.mode;
      if (mode) setFilterSearchMode(mode);
    });
  });

  // 数値範囲指定パネルのトグルと操作
  if (btnToggleNumRange) {
    btnToggleNumRange.addEventListener('click', (e) => {
      e.stopPropagation();
      toggleNumericRangePanel();
    });
  }
  if (btnFilterNumApply) {
    btnFilterNumApply.addEventListener('click', (e) => {
      e.stopPropagation();
      applyNumericRangeSelection();
    });
  }
  if (btnFilterNumClear) {
    btnFilterNumClear.addEventListener('click', (e) => {
      e.stopPropagation();
      clearNumericRangeInputs();
    });
  }
  if (filterNumMin) {
    filterNumMin.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        applyNumericRangeSelection();
      }
    });
  }
  if (filterNumMax) {
    filterNumMax.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        applyNumericRangeSelection();
      }
    });
  }

  // ポップアップの外側をクリックしたら閉じる
  document.addEventListener('click', (e) => {
    if (filterPopup && filterPopup.style.display === 'block') {
      if (!filterPopup.contains(e.target) && !e.target.closest('.btn-col-filter')) {
        hideFilterPopup();
      }
    }
    if (searchHelpPopup && searchHelpPopup.style.display !== 'none') {
      if (!searchHelpPopup.contains(e.target) && e.target !== btnSearchHelp && !btnSearchHelp.contains(e.target)) {
        hideSearchHelpPopup();
      }
    }
  });

  // 簡易絞り込み検索
  searchInput.addEventListener('input', handleSearch);
  btnClearSearch.addEventListener('click', () => {
    searchInput.value = '';
    handleSearch();
  });
  searchInput.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      e.preventDefault();
      searchInput.value = '';
      handleSearch();
      searchInput.blur();
    }
  });

  // 保存モーダル内の設定変更（タブの保存設定に反映）
  if (selectEncoding) selectEncoding.addEventListener('change', applySaveSettingsToCurrentTab);
  if (selectDelimiter) selectDelimiter.addEventListener('change', applySaveSettingsToCurrentTab);
  if (selectNewline) selectNewline.addEventListener('change', applySaveSettingsToCurrentTab);
  if (selectCommaMode) selectCommaMode.addEventListener('change', applySaveSettingsToCurrentTab);

  // ショートカットキー・キーボード操作（Excel準拠）
  window.addEventListener('keydown', handleGlobalKeydown);

  // Enter キーによるセル移動・行挿入
  window.addEventListener('keydown', handleRightmostEnterKey, true);

  // 矢印キー（カーソルキー）によるセル間移動
  window.addEventListener('keydown', handleArrowKeyNavigation, true);

  // クリップボード貼り付け (Ctrl+V)
  window.addEventListener('paste', handleGlobalPaste);

  // マウスドラッグ範囲選択イベント
  window.addEventListener('mouseup', handleGlobalMouseUp);
  if (tableArea) {
    tableArea.addEventListener('mousedown', handleCellMouseDown);
    tableArea.addEventListener('mouseover', handleCellMouseMove);
    tableArea.addEventListener('dblclick', (e) => {
      const headerEl = e.target.closest('.tabulator-col');
      if (headerEl && currentTable) {
        const field = headerEl.getAttribute('tabulator-field');
        if (field) {
          const col = currentTable.getColumn(field);
          if (col) {
            renameColumn(col);
          }
        }
      }
    });
  }
}

// グローバルキーボードショートカット処理 (Excel風)
function handleGlobalKeydown(e) {
  const isEditing = !!document.querySelector('.tabulator-cell.tabulator-editing');

  // 表示モード（非編集中）での日本語入力（IME）開始検知:
  // 全角入力キーが押されたら即座にアクティブセルを入力モード（input）にしてIMEを受け取れるようにする
  if (!isEditing && (e.isComposing || e.keyCode === 229 || e.key === 'Process')) {
    const tag = document.activeElement ? document.activeElement.tagName.toLowerCase() : '';
    if (tag !== 'input' && tag !== 'textarea' && tag !== 'select') {
      if (!activeFocusCell) {
        activeFocusCell = { rowIndex: 0, colIndex: 0 };
      }
      const cell = getCellAt(activeFocusCell.rowIndex, activeFocusCell.colIndex);
      if (cell && typeof cell.edit === 'function') {
        cell.edit(true);
        setTimeout(() => {
          const input = cell.getElement().querySelector('input');
          if (input) input.focus();
        }, 20);
      }
    }
    return;
  }

  // 入力モード（編集中）の日本語入力変換中キーはショートカット処理を行わない
  if (e.isComposing || e.keyCode === 229) return;

  // Ctrl + S (保存)
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
    e.preventDefault();
    if (e.shiftKey) {
      saveAsCurrentTab();
    } else {
      handleSave();
    }
    return;
  }

  // Ctrl + O (開く)
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'o') {
    e.preventDefault();
    openFileWithPicker();
    return;
  }

  // Ctrl + F (検索)
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'f') {
    e.preventDefault();
    openFindReplaceModal('find');
    return;
  }

  // Ctrl + H (置換)
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'h') {
    e.preventDefault();
    openFindReplaceModal('replace');
    return;
  }

  // Ctrl + Z / Ctrl + Y (Undo / Redo)
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
    if (!isEditing) {
      e.preventDefault();
      if (e.shiftKey) {
        if (currentTable) currentTable.redo();
      } else {
        if (currentTable) currentTable.undo();
      }
    }
    return;
  }
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') {
    if (!isEditing) {
      e.preventDefault();
      if (currentTable) currentTable.redo();
    }
    return;
  }

  // Ctrl + C (コピー) - 非編集中
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c') {
    if (!isEditing) {
      e.preventDefault();
      copySelectedCellsToClipboard();
    }
    return;
  }

  // Delete / Backspace (セル一括クリア) - 非編集中
  if (e.key === 'Delete' || e.key === 'Backspace') {
    if (!isEditing) {
      // 検索バーやモーダルにフォーカスがない場合のみ
      const tag = document.activeElement ? document.activeElement.tagName.toLowerCase() : '';
      if (tag !== 'input' && tag !== 'textarea') {
        e.preventDefault();
        clearSelectedCells();
      }
    }
    return;
  }

  // F2 (セル編集開始、テキスト末尾にキャレット)
  if (e.key === 'F2') {
    e.preventDefault();
    startCellEditAtEnd();
    return;
  }

  // Tab / Shift + Tab (左右セル移動)
  if (e.key === 'Tab') {
    e.preventDefault();
    handleTabNavigation(e.shiftKey);
    return;
  }

  // Escape (入力モードキャンセル→表示モード復帰、各種モーダル・ポップアップ・コピー点線解除)
  if (e.key === 'Escape') {
    const modalColSummary = document.getElementById('modal-column-summary');
    if (modalColSummary && modalColSummary.style.display !== 'none') {
      hideColumnSummaryModal();
      return;
    }
    if (findReplaceModal && findReplaceModal.style.display !== 'none') {
      closeFindReplaceModal();
      return;
    }
    if (saveModal && saveModal.style.display !== 'none') {
      hideSaveModal();
      return;
    }
    if (filterPopup && filterPopup.style.display !== 'none') {
      hideFilterPopup();
      return;
    }
    if (searchHelpPopup && searchHelpPopup.style.display !== 'none') {
      hideSearchHelpPopup();
      return;
    }
    if (cleanDropdownMenu && cleanDropdownMenu.style.display !== 'none') {
      hideCleanDropdown();
      return;
    }
    if (copiedRange) {
      copiedRange = null;
      updateRangeHighlight();
    }
    if (isEditing) {
      e.preventDefault();
      cancelCellEdit(); // 入力モードをキャンセルして元の値に戻し、表示モードへ切り替える！
      return;
    }
    return;
  }

  // 表示モード（非編集中）に文字・数字・記号キーが押されたら即座に入力モードに切り替えて入力開始 (Excel風)
  if (!isEditing && !e.ctrlKey && !e.metaKey && !e.altKey && e.key.length === 1) {
    const tag = document.activeElement ? document.activeElement.tagName.toLowerCase() : '';
    if (tag !== 'input' && tag !== 'textarea' && tag !== 'select') {
      if (!activeFocusCell) {
        activeFocusCell = { rowIndex: 0, colIndex: 0 };
      }
      const cell = getCellAt(activeFocusCell.rowIndex, activeFocusCell.colIndex);
      if (cell && typeof cell.edit === 'function') {
        cell.edit(true);
        const initialChar = e.key;
        setTimeout(() => {
          const input = cell.getElement().querySelector('input');
          if (input) {
            input.value = initialChar;
            input.setSelectionRange(initialChar.length, initialChar.length);
          }
        }, 30);
      }
    }
  }
}

// ----------------------------------------------------
// 範囲選択・セルフォーカス・キーボード操作エンジン (Excel風)
// ----------------------------------------------------

// 現在アクティブな列定義（fieldを持つ列）を取得
function getActiveColumns() {
  if (!activeTabId) return [];
  const tab = tabs.find(t => t.id === activeTabId);
  return tab ? tab.columns.filter(c => c.field) : [];
}

// 現在の可視行（フィルター通過後の表示行）を取得 (Excel互換)
function getAllRows() {
  if (!currentTable) return [];
  try {
    return currentTable.getRows("active");
  } catch (e) {
    return [];
  }
}

// 指定した行・列インデックスのセルオブジェクトを取得
function getCellAt(rowIndex, colIndex) {
  const rows = getAllRows();
  const cols = getActiveColumns();
  if (rowIndex < 0 || rowIndex >= rows.length || colIndex < 0 || colIndex >= cols.length) return null;
  const row = rows[rowIndex];
  if (!row) return null;
  return row.getCell(cols[colIndex].field);
}

// セルまたは要素から (rowIndex, colIndex) を逆引き
function getCellIndices(target) {
  if (!currentTable || !activeTabId || !target) return null;
  const cols = getActiveColumns();
  const rows = getAllRows();
  if (cols.length === 0 || rows.length === 0) return null;

  let cellEl = null;
  if (target.getElement && typeof target.getElement === 'function') {
    cellEl = target.getElement();
  } else if (target.classList && target.classList.contains('tabulator-cell')) {
    cellEl = target;
  } else if (target.closest) {
    cellEl = target.closest('.tabulator-cell');
  }
  if (!cellEl) return null;

  const field = cellEl.getAttribute('tabulator-field');
  const colIndex = cols.findIndex(c => c.field === field);
  if (colIndex === -1) return null;

  const rowEl = cellEl.closest('.tabulator-row');
  if (!rowEl) return null;
  const rowIndex = rows.findIndex(r => r.getElement() === rowEl);
  if (rowIndex === -1) return null;

  return { rowIndex, colIndex, field, rowEl, cellEl };
}

// 単一セルを「表示モード」（選択状態）にする
function selectSingleCell(rowIndex, colIndex) {
  const rows = getAllRows();
  const cols = getActiveColumns();
  if (rowIndex < 0 || rowIndex >= rows.length || colIndex < 0 || colIndex >= cols.length) return;

  activeFocusCell = { rowIndex, colIndex };
  selectionAnchor = { rowIndex, colIndex };
  selectionRange = {
    minRow: rowIndex,
    maxRow: rowIndex,
    minCol: colIndex,
    maxCol: colIndex,
    startRow: rowIndex,
    startCol: colIndex,
    endRow: rowIndex,
    endCol: colIndex
  };

  updateRangeHighlight();

  const targetCell = getCellAt(rowIndex, colIndex);
  if (targetCell) {
    const el = targetCell.getElement();
    if (el) {
      el.focus();
      el.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    }
  }
}

// 入力モードの編集をキャンセルし、元の値に戻して「表示モード」へ切り替える (ESCキー用)
function cancelCellEdit() {
  if (!currentTable) return;
  const editingCellEl = document.querySelector('.tabulator-cell.tabulator-editing');
  if (!editingCellEl) return;

  const indices = getCellIndices(editingCellEl);

  if (currentEditingCell && typeof currentEditingCell.cancelEdit === 'function') {
    try {
      currentEditingCell.cancelEdit();
    } catch (e) {
      console.warn("cancelEdit warning:", e);
    }
  } else {
    const inputEl = editingCellEl.querySelector('input');
    if (inputEl) {
      inputEl.value = editingOriginalValue;
      inputEl.blur();
    }
  }

  currentEditingCell = null;

  if (indices) {
    setTimeout(() => {
      selectSingleCell(indices.rowIndex, indices.colIndex);
    }, 40);
  }
}

// 範囲選択ハイライトの描画更新 (Excel風)
function updateRangeHighlight() {
  if (!currentTable || !tableArea) return;

  // 1. 既存の全ハイライトクラスを削除
  const oldElements = tableArea.querySelectorAll(
    '.tabulator-range-selected, .tabulator-range-top, .tabulator-range-bottom, ' +
    '.tabulator-range-left, .tabulator-range-right, .tabulator-cell-focused, ' +
    '.tabulator-copied-top, .tabulator-copied-bottom, .tabulator-copied-left, ' +
    '.tabulator-copied-right, .tabulator-copied-bg'
  );
  oldElements.forEach(el => {
    el.classList.remove(
      'tabulator-range-selected', 'tabulator-range-top', 'tabulator-range-bottom',
      'tabulator-range-left', 'tabulator-range-right', 'tabulator-cell-focused',
      'tabulator-copied-top', 'tabulator-copied-bottom', 'tabulator-copied-left',
      'tabulator-copied-right', 'tabulator-copied-bg'
    );
  });

  const rows = getAllRows();
  const cols = getActiveColumns();
  if (rows.length === 0 || cols.length === 0) return;

  // 2. コピー中の点線枠（copiedRange）をハイライト
  if (copiedRange) {
    const minR = Math.max(0, Math.min(copiedRange.minRow, rows.length - 1));
    const maxR = Math.max(0, Math.min(copiedRange.maxRow, rows.length - 1));
    const minC = Math.max(0, Math.min(copiedRange.minCol, cols.length - 1));
    const maxC = Math.max(0, Math.min(copiedRange.maxCol, cols.length - 1));

    for (let r = minR; r <= maxR; r++) {
      const row = rows[r];
      if (!row) continue;
      for (let c = minC; c <= maxC; c++) {
        const cell = row.getCell(cols[c].field);
        if (cell) {
          const el = cell.getElement();
          if (el) {
            el.classList.add('tabulator-copied-bg');
            if (r === minR) el.classList.add('tabulator-copied-top');
            if (r === maxR) el.classList.add('tabulator-copied-bottom');
            if (c === minC) el.classList.add('tabulator-copied-left');
            if (c === maxC) el.classList.add('tabulator-copied-right');
          }
        }
      }
    }
  }

  // 3. 選択範囲（selectionRange）のハイライト
  if (selectionRange) {
    const minR = Math.max(0, Math.min(selectionRange.minRow, rows.length - 1));
    const maxR = Math.max(0, Math.min(selectionRange.maxRow, rows.length - 1));
    const minC = Math.max(0, Math.min(selectionRange.minCol, cols.length - 1));
    const maxC = Math.max(0, Math.min(selectionRange.maxCol, cols.length - 1));
    const isSingle = (minR === maxR && minC === maxC);

    for (let r = minR; r <= maxR; r++) {
      const row = rows[r];
      if (!row) continue;
      for (let c = minC; c <= maxC; c++) {
        const cell = row.getCell(cols[c].field);
        if (cell) {
          const el = cell.getElement();
          if (el) {
            if (isSingle) {
              el.classList.add('tabulator-cell-focused');
            } else {
              el.classList.add('tabulator-range-selected');
              if (r === minR) el.classList.add('tabulator-range-top');
              if (r === maxR) el.classList.add('tabulator-range-bottom');
              if (c === minC) el.classList.add('tabulator-range-left');
              if (c === maxC) el.classList.add('tabulator-range-right');
            }
          }
        }
      }
    }

    // 複数セル選択時は簡易集計バーを更新
    updateCalculationStatusBarWithRange(minR, maxR, minC, maxC);
  } else if (activeFocusCell) {
    // 単一セルのみ選択
    const cell = getCellAt(activeFocusCell.rowIndex, activeFocusCell.colIndex);
    if (cell) {
      const el = cell.getElement();
      if (el) el.classList.add('tabulator-cell-focused');
    }
    hideCalculationStatusBar();
  } else {
    hideCalculationStatusBar();
  }
}

// 選択セル範囲の簡易集計（ステータスバー: 平均・個数・合計）
function updateCalculationStatusBarWithRange(minRow, maxRow, minCol, maxCol) {
  if (!statusCalcArea || !calcCount || !calcAvg || !calcSum) return;

  const isSingle = (minRow === maxRow && minCol === maxCol);
  if (isSingle) {
    hideCalculationStatusBar();
    return;
  }

  const rows = getAllRows();
  const cols = getActiveColumns();

  let totalCount = 0;
  let numericCount = 0;
  let sum = 0;

  for (let r = minRow; r <= maxRow; r++) {
    if (r >= rows.length) break;
    const row = rows[r];
    if (!row) continue;
    for (let c = minCol; c <= maxCol; c++) {
      if (c >= cols.length) break;
      totalCount++;
      const cell = row.getCell(cols[c].field);
      if (cell) {
        const val = cell.getValue();
        if (val !== null && val !== undefined && val !== '') {
          const cleanVal = String(val).replace(/,/g, '').trim();
          const num = Number(cleanVal);
          if (!isNaN(num) && cleanVal !== '') {
            sum += num;
            numericCount++;
          }
        }
      }
    }
  }

  if (totalCount <= 1) {
    hideCalculationStatusBar();
    return;
  }

  calcCount.textContent = totalCount.toLocaleString();
  if (numericCount > 0) {
    const avg = sum / numericCount;
    calcAvg.textContent = (avg % 1 === 0) ? avg.toLocaleString() : avg.toFixed(2);
    calcSum.textContent = (sum % 1 === 0) ? sum.toLocaleString() : sum.toFixed(2);
  } else {
    calcAvg.textContent = '-';
    calcSum.textContent = '-';
  }

  statusCalcArea.style.display = 'inline-flex';
}

function hideCalculationStatusBar() {
  if (statusCalcArea) {
    statusCalcArea.style.display = 'none';
  }
}

// カーソルキー（矢印キー）によるセル移動および Ctrl+矢印 / Shift+矢印 による範囲指定
function handleArrowKeyNavigation(e) {
  const key = e.key;
  if (key !== 'ArrowLeft' && key !== 'ArrowRight' && key !== 'ArrowUp' && key !== 'ArrowDown') {
    return;
  }

  // 日本語入力（IME）変換中は無視
  if (e.isComposing || e.keyCode === 229) return;

  if (!currentTable || !activeTabId) return;
  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab) return;

  const activeCols = getActiveColumns();
  const rows = getAllRows();
  if (activeCols.length === 0 || rows.length === 0) return;

  // 範囲選択キー判定（Ctrl + 矢印 または Shift + 矢印）
  const isRangeSelectKey = (e.ctrlKey || e.shiftKey) && !e.altKey && !e.metaKey;

  // 1. セル編集中（input要素が存在する場合）
  const editingCellEl = document.querySelector('.tabulator-cell.tabulator-editing');
  if (editingCellEl) {
    const inputEl = editingCellEl.querySelector('input');
    if (!inputEl) return;

    const indices = getCellIndices(editingCellEl);
    if (!indices) return;

    if (isRangeSelectKey) {
      // 編集中であっても Ctrl+矢印 / Shift+矢印 が押されたら編集を確定して範囲選択を開始！
      e.preventDefault();
      e.stopPropagation();
      inputEl.blur();

      selectionAnchor = { rowIndex: indices.rowIndex, colIndex: indices.colIndex };
      activeFocusCell = { rowIndex: indices.rowIndex, colIndex: indices.colIndex };

      let nextRow = indices.rowIndex;
      let nextCol = indices.colIndex;
      if (key === 'ArrowLeft') nextCol = Math.max(0, nextCol - 1);
      else if (key === 'ArrowRight') nextCol = Math.min(activeCols.length - 1, nextCol + 1);
      else if (key === 'ArrowUp') nextRow = Math.max(0, nextRow - 1);
      else if (key === 'ArrowDown') nextRow = Math.min(rows.length - 1, nextRow + 1);

      activeFocusCell = { rowIndex: nextRow, colIndex: nextCol };
      selectionRange = {
        minRow: Math.min(selectionAnchor.rowIndex, activeFocusCell.rowIndex),
        maxRow: Math.max(selectionAnchor.rowIndex, activeFocusCell.rowIndex),
        minCol: Math.min(selectionAnchor.colIndex, activeFocusCell.colIndex),
        maxCol: Math.max(selectionAnchor.colIndex, activeFocusCell.colIndex),
        startRow: selectionAnchor.rowIndex,
        startCol: selectionAnchor.colIndex,
        endRow: activeFocusCell.rowIndex,
        endCol: activeFocusCell.colIndex
      };

      updateRangeHighlight();
      const targetCell = getCellAt(nextRow, nextCol);
      if (targetCell && targetCell.getElement()) {
        targetCell.getElement().scrollIntoView({ block: 'nearest', inline: 'nearest' });
      }
      return;
    }

    // 通常の矢印キー（単独）
    const isAtStart = (inputEl.selectionStart === 0 && inputEl.selectionEnd === 0);
    const isAtEnd = (inputEl.selectionStart === inputEl.value.length && inputEl.selectionEnd === inputEl.value.length);
    const isEmpty = (inputEl.value === '');

    let shouldMove = false;
    let nextRowIndex = indices.rowIndex;
    let nextColIndex = indices.colIndex;

    if (key === 'ArrowLeft' && (isAtStart || isEmpty) && indices.colIndex > 0) {
      shouldMove = true;
      nextColIndex = indices.colIndex - 1;
    } else if (key === 'ArrowRight' && (isAtEnd || isEmpty) && indices.colIndex < activeCols.length - 1) {
      shouldMove = true;
      nextColIndex = indices.colIndex + 1;
    } else if (key === 'ArrowUp' && indices.rowIndex > 0) {
      shouldMove = true;
      nextRowIndex = indices.rowIndex - 1;
    } else if (key === 'ArrowDown' && indices.rowIndex < rows.length - 1) {
      shouldMove = true;
      nextRowIndex = indices.rowIndex + 1;
    }

    if (shouldMove) {
      e.preventDefault();
      e.stopPropagation();
      inputEl.blur();

      const nextRow = rows[nextRowIndex];
      const nextField = activeCols[nextColIndex].field;

      setTimeout(() => {
        moveToAndEditCell(nextRow, nextField);
      }, 30);
      return;
    }
    return;
  }

  // 2. セル非編集中
  const activeEl = document.activeElement;
  const tag = activeEl ? activeEl.tagName.toLowerCase() : '';
  if (tag === 'input' || tag === 'textarea' || tag === 'select') {
    // 検索窓やモーダル入力欄などにフォーカスがある場合は矢印キーによるセル移動をしない
    return;
  }

  if (!activeFocusCell) {
    const indices = getCellIndices(activeEl);
    if (indices) {
      activeFocusCell = { rowIndex: indices.rowIndex, colIndex: indices.colIndex };
      selectionAnchor = { ...activeFocusCell };
    } else {
      activeFocusCell = { rowIndex: 0, colIndex: 0 };
      selectionAnchor = { rowIndex: 0, colIndex: 0 };
    }
  }

  if (isRangeSelectKey) {
    // 【重要】Ctrl + 矢印 または Shift + 矢印 による範囲指定！
    e.preventDefault();
    e.stopPropagation();

    if (!selectionAnchor) {
      selectionAnchor = { ...activeFocusCell };
    }

    let nextRow = activeFocusCell.rowIndex;
    let nextCol = activeFocusCell.colIndex;

    if (key === 'ArrowLeft') {
      nextCol = Math.max(0, nextCol - 1);
    } else if (key === 'ArrowRight') {
      nextCol = Math.min(activeCols.length - 1, nextCol + 1);
    } else if (key === 'ArrowUp') {
      nextRow = Math.max(0, nextRow - 1);
    } else if (key === 'ArrowDown') {
      nextRow = Math.min(rows.length - 1, nextRow + 1);
    }

    activeFocusCell = { rowIndex: nextRow, colIndex: nextCol };
    selectionRange = {
      minRow: Math.min(selectionAnchor.rowIndex, activeFocusCell.rowIndex),
      maxRow: Math.max(selectionAnchor.rowIndex, activeFocusCell.rowIndex),
      minCol: Math.min(selectionAnchor.colIndex, activeFocusCell.colIndex),
      maxCol: Math.max(selectionAnchor.colIndex, activeFocusCell.colIndex),
      startRow: selectionAnchor.rowIndex,
      startCol: selectionAnchor.colIndex,
      endRow: activeFocusCell.rowIndex,
      endCol: activeFocusCell.colIndex
    };

    updateRangeHighlight();

    const targetCell = getCellAt(nextRow, nextCol);
    if (targetCell && targetCell.getElement()) {
      targetCell.getElement().scrollIntoView({ block: 'nearest', inline: 'nearest' });
    }
    return;
  }

  // 通常の矢印キー（単独）: 単一セルの移動
  if (!e.ctrlKey && !e.shiftKey && !e.altKey && !e.metaKey) {
    e.preventDefault();
    e.stopPropagation();

    let nextRow = activeFocusCell.rowIndex;
    let nextCol = activeFocusCell.colIndex;

    if (key === 'ArrowLeft' && nextCol > 0) {
      nextCol--;
    } else if (key === 'ArrowRight' && nextCol < activeCols.length - 1) {
      nextCol++;
    } else if (key === 'ArrowUp' && nextRow > 0) {
      nextRow--;
    } else if (key === 'ArrowDown' && nextRow < rows.length - 1) {
      nextRow++;
    }

    activeFocusCell = { rowIndex: nextRow, colIndex: nextCol };
    selectionAnchor = { rowIndex: nextRow, colIndex: nextCol };
    selectionRange = {
      minRow: nextRow,
      maxRow: nextRow,
      minCol: nextCol,
      maxCol: nextCol,
      startRow: nextRow,
      startCol: nextCol,
      endRow: nextRow,
      endCol: nextCol
    };

    updateRangeHighlight();

    const targetCell = getCellAt(nextRow, nextCol);
    if (targetCell) {
      const el = targetCell.getElement();
      if (el) {
        el.focus();
        el.scrollIntoView({ block: 'nearest', inline: 'nearest' });
      }
    }
  }
}

// セルクリック時の選択ハンドラ
function handleCellClick(e, cell) {
  const indices = getCellIndices(cell);
  if (!indices) return;

  if (e.shiftKey && selectionAnchor) {
    activeFocusCell = { rowIndex: indices.rowIndex, colIndex: indices.colIndex };
    selectionRange = {
      minRow: Math.min(selectionAnchor.rowIndex, indices.rowIndex),
      maxRow: Math.max(selectionAnchor.rowIndex, indices.rowIndex),
      minCol: Math.min(selectionAnchor.colIndex, indices.colIndex),
      maxCol: Math.max(selectionAnchor.colIndex, indices.colIndex),
      startRow: selectionAnchor.rowIndex,
      startCol: selectionAnchor.colIndex,
      endRow: indices.rowIndex,
      endCol: indices.colIndex
    };
  } else {
    activeFocusCell = { rowIndex: indices.rowIndex, colIndex: indices.colIndex };
    selectionAnchor = { rowIndex: indices.rowIndex, colIndex: indices.colIndex };
    selectionRange = {
      minRow: indices.rowIndex,
      maxRow: indices.rowIndex,
      minCol: indices.colIndex,
      maxCol: indices.colIndex,
      startRow: indices.rowIndex,
      startCol: indices.colIndex,
      endRow: indices.rowIndex,
      endCol: indices.colIndex
    };
  }
  updateRangeHighlight();
}

// セル上での mousedown (ドラッグ選択の開始)
function handleCellMouseDown(e) {
  if (e.button !== 0) return;
  const indices = getCellIndices(e.target);
  if (!indices) return;

  if (e.target.tagName.toLowerCase() === 'input') return;

  if (e.shiftKey && selectionAnchor) {
    activeFocusCell = { rowIndex: indices.rowIndex, colIndex: indices.colIndex };
    selectionRange = {
      minRow: Math.min(selectionAnchor.rowIndex, indices.rowIndex),
      maxRow: Math.max(selectionAnchor.rowIndex, indices.rowIndex),
      minCol: Math.min(selectionAnchor.colIndex, indices.colIndex),
      maxCol: Math.max(selectionAnchor.colIndex, indices.colIndex),
      startRow: selectionAnchor.rowIndex,
      startCol: selectionAnchor.colIndex,
      endRow: indices.rowIndex,
      endCol: indices.colIndex
    };
  } else {
    activeFocusCell = { rowIndex: indices.rowIndex, colIndex: indices.colIndex };
    selectionAnchor = { rowIndex: indices.rowIndex, colIndex: indices.colIndex };
    selectionRange = {
      minRow: indices.rowIndex,
      maxRow: indices.rowIndex,
      minCol: indices.colIndex,
      maxCol: indices.colIndex,
      startRow: indices.rowIndex,
      startCol: indices.colIndex,
      endRow: indices.rowIndex,
      endCol: indices.colIndex
    };
    isMouseDraggingRange = true;
  }
  updateRangeHighlight();
}

// マウスドラッグ中のセル追随
function handleCellMouseMove(e) {
  if (!isMouseDraggingRange || !selectionAnchor) return;
  const indices = getCellIndices(e.target);
  if (!indices) return;

  if (activeFocusCell && activeFocusCell.rowIndex === indices.rowIndex && activeFocusCell.colIndex === indices.colIndex) {
    return;
  }

  activeFocusCell = { rowIndex: indices.rowIndex, colIndex: indices.colIndex };
  selectionRange = {
    minRow: Math.min(selectionAnchor.rowIndex, indices.rowIndex),
    maxRow: Math.max(selectionAnchor.rowIndex, indices.rowIndex),
    minCol: Math.min(selectionAnchor.colIndex, indices.colIndex),
    maxCol: Math.max(selectionAnchor.colIndex, indices.colIndex),
    startRow: selectionAnchor.rowIndex,
    startCol: selectionAnchor.colIndex,
    endRow: indices.rowIndex,
    endCol: indices.colIndex
  };
  updateRangeHighlight();
}

// マウスドラッグ終了
function handleGlobalMouseUp() {
  isMouseDraggingRange = false;
}

// 指定した行・列のセルへ移動して編集モードを開始
function moveToAndEditCell(targetRow, field) {
  try {
    const cell = targetRow.getCell(field);
    if (cell) {
      const indices = getCellIndices(cell);
      if (indices) {
        activeFocusCell = { rowIndex: indices.rowIndex, colIndex: indices.colIndex };
        selectionAnchor = { rowIndex: indices.rowIndex, colIndex: indices.colIndex };
        selectionRange = {
          minRow: indices.rowIndex, maxRow: indices.rowIndex,
          minCol: indices.colIndex, maxCol: indices.colIndex,
          startRow: indices.rowIndex, startCol: indices.colIndex,
          endRow: indices.rowIndex, endCol: indices.colIndex
        };
        updateRangeHighlight();
      }
      if (typeof cell.edit === 'function') {
        cell.edit(true);
      } else {
        const el = cell.getElement();
        if (el) el.focus();
      }
    }
  } catch (err) {
    console.warn("moveToAndEditCell error:", err);
  }
}

// セルでの Enter キー処理（入力モード確定＆移動、または表示モードでの入力モード開始）
function handleRightmostEnterKey(e) {
  if (e.key !== 'Enter') return;
  // 日本語入力（IME）変換中の Enter は無視
  if (e.isComposing || e.keyCode === 229) return;
  if (e.shiftKey || e.ctrlKey || e.metaKey || e.altKey) return;

  if (!currentTable || !activeTabId) return;
  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab) return;

  const activeCols = getActiveColumns();
  const rows = getAllRows();
  if (activeCols.length === 0 || rows.length === 0) return;

  // 1. セル編集中（入力モード中）
  const editingCellEl = document.querySelector('.tabulator-cell.tabulator-editing');
  if (editingCellEl) {
    const inputEl = editingCellEl.querySelector('input');
    if (!inputEl) return;

    const indices = getCellIndices(editingCellEl);
    if (!indices) return;

    const targetRow = rows[indices.rowIndex];
    if (!targetRow) return;

    const currentInputVal = inputEl.value || '';
    const isRightmost = (indices.colIndex === activeCols.length - 1);

    e.preventDefault();
    e.stopPropagation();

    // 入力内容を確定
    inputEl.blur();

    if (isRightmost) {
      // 一番右端の列の場合
      if (isRowEmpty(targetRow, tab, indices.field, currentInputVal)) {
        // 行全体が空なら追加せず確定のみでそのセルを表示モードに
        setTimeout(() => {
          selectSingleCell(indices.rowIndex, indices.colIndex);
        }, 30);
        return;
      }

      const isBottomRow = (indices.rowIndex >= rows.length - 1);
      if (isBottomRow) {
        // 最下行（D4など）の場合：下に追加された新しい空行の一番左のセル（第1列：A5）に移動！
        setTimeout(() => {
          insertEmptyRowBelow(targetRow, tab);
        }, 30);
        return;
      } else {
        // 途中の行（D1〜D3など）の場合：行は追加せず、既存の次行の先頭（第1列：A2〜A4）へ移動！
        setTimeout(() => {
          selectSingleCell(indices.rowIndex + 1, 0);
        }, 30);
        return;
      }
    } else {
      // 一番右端ではない列の場合 → 一つ右隣の列のセルに移動（表示モード）！
      const nextField = activeCols[indices.colIndex + 1].field;
      setTimeout(() => {
        moveToNextCell(targetRow, nextField);
      }, 30);
      return;
    }
  }

  // 2. 表示モード（非編集状態）での Enter キー:
  // データを編集または入力したいセルでリターンキーを押すと入力モードに切り替わる！
  const tag = document.activeElement ? document.activeElement.tagName.toLowerCase() : '';
  if (tag !== 'input' && tag !== 'textarea' && tag !== 'select') {
    e.preventDefault();
    e.stopPropagation();
    startCellEditAtEnd();
  }
}

// 同じ行の次のセルに移動して「表示モード」（選択状態）にする
function moveToNextCell(targetRow, nextField) {
  try {
    const cell = targetRow.getCell(nextField);
    if (cell) {
      const indices = getCellIndices(cell);
      if (indices) {
        selectSingleCell(indices.rowIndex, indices.colIndex);
      }
    }
  } catch (err) {
    console.warn("moveToNextCell error:", err);
  }
}

// 行が完全に空欄（すべてのセルが空）かどうか判定
function isRowEmpty(targetRow, tab, currentEditingField, currentEditingValue) {
  if (!targetRow || !tab) return true;
  const rowData = targetRow.getData();
  const activeFields = tab.columns.filter(c => c.field).map(c => c.field);

  return activeFields.every(field => {
    let val = rowData[field];
    if (field === currentEditingField && currentEditingValue !== undefined) {
      val = currentEditingValue;
    }
    return val === null || val === undefined || String(val).trim() === '';
  });
}

// 指定行の直下に空の行を挿入し、追加行の一番左のセル（第1列）を「表示モード」にする
function insertEmptyRowBelow(targetRow, tab) {
  const newRow = { _id: Date.now() };
  tab.columns.filter(c => c.field).forEach(c => newRow[c.field] = '');

  currentTable.addRow(newRow, false, targetRow).then((addedRow) => {
    tab.isModified = true;
    syncCurrentTabData();
    renderTabs();
    updateStatusBar(tab);

    // 新しく挿入した行の一番左端のセル（第1列）を表示モードにする
    if (addedRow) {
      setTimeout(() => {
        try {
          const rows = getAllRows();
          const newRowIdx = rows.findIndex(r => r === addedRow || r.getData()._id === addedRow.getData()._id);
          if (newRowIdx !== -1) {
            selectSingleCell(newRowIdx, 0);
          } else {
            const cells = addedRow.getCells();
            if (cells.length > 0) {
              const indices = getCellIndices(cells[0]);
              if (indices) selectSingleCell(indices.rowIndex, indices.colIndex);
            }
          }
        } catch (err) {
          console.warn("Could not select new row first cell:", err);
        }
      }, 40);
    }
  }).catch(err => {
    console.error("Error inserting row below:", err);
  });
}

// ブラウザのデフォルトのドロップ保存を完全に防ぎ、確実にファイルを取り込む
function setupGlobalDragAndDrop() {
  let dragCounter = 0;

  function onDragOver(e) {
    e.preventDefault();
    if (e.dataTransfer) {
      e.dataTransfer.dropEffect = 'copy';
    }
  }

  function onDragEnter(e) {
    e.preventDefault();
    if (e.dataTransfer) {
      e.dataTransfer.dropEffect = 'copy';
    }
    dragCounter++;
    if (dropOverlay) {
      dropOverlay.classList.add('active');
    }
  }

  function onDragLeave(e) {
    e.preventDefault();
    dragCounter--;
    if (dragCounter <= 0) {
      dragCounter = 0;
      if (dropOverlay) {
        dropOverlay.classList.remove('active');
      }
    }
  }

  async function onDrop(e) {
    e.preventDefault();
    e.stopPropagation();
    dragCounter = 0;
    if (dropOverlay) {
      dropOverlay.classList.remove('active');
    }

    const dt = e.dataTransfer;
    if (!dt) return;

    // 1. 同期的にすべてのファイルとハンドルPromiseを取得
    // ※await でイベントループが進むとブラウザのセキュリティ仕様により DataTransfer が消滅するため、同期コンテキストでの抽出が必須
    const fileEntries = [];

    if (dt.files && dt.files.length > 0) {
      const files = Array.from(dt.files);
      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        let handlePromise = null;
        if (dt.items && dt.items[i] && typeof dt.items[i].getAsFileSystemHandle === 'function') {
          try {
            handlePromise = dt.items[i].getAsFileSystemHandle();
          } catch (err) {
            handlePromise = null;
          }
        }
        fileEntries.push({ file, handlePromise });
      }
    } else if (dt.items && dt.items.length > 0) {
      for (let i = 0; i < dt.items.length; i++) {
        const item = dt.items[i];
        if (item.kind === 'file') {
          const file = item.getAsFile();
          let handlePromise = null;
          if (typeof item.getAsFileSystemHandle === 'function') {
            try {
              handlePromise = item.getAsFileSystemHandle();
            } catch (err) {
              handlePromise = null;
            }
          }
          if (file) {
            fileEntries.push({ file, handlePromise });
          }
        }
      }
    }

    if (fileEntries.length === 0) {
      console.warn("No files found in drop event.");
      return;
    }

    console.log(`Processing dropped ${fileEntries.length} file(s)...`);

    // 2. 取得したファイル（およびハンドル）を順次読み込み
    for (const entry of fileEntries) {
      let handle = null;
      if (entry.handlePromise) {
        try {
          const h = await entry.handlePromise;
          if (h && h.kind === 'file') {
            handle = h;
          }
        } catch (err) {
          console.warn("Could not retrieve file system handle from drop:", err);
        }
      }
      readFileWithHandle(entry.file, handle);
    }
  }

  window.addEventListener('dragenter', onDragEnter, false);
  window.addEventListener('dragover', onDragOver, false);
  window.addEventListener('dragleave', onDragLeave, false);
  window.addEventListener('drop', onDrop, false);
}

// ファイルピッカーでファイルを開く (File System Access API 対応)
async function openFileWithPicker() {
  if ('showOpenFilePicker' in window) {
    try {
      const handles = await window.showOpenFilePicker({
        multiple: true,
        types: [
          {
            description: 'CSV / TSV / テキストファイル',
            accept: {
              'text/csv': ['.csv'],
              'text/tab-separated-values': ['.tsv'],
              'text/plain': ['.txt', '.csv', '.tsv']
            }
          }
        ]
      });
      for (const handle of handles) {
        const file = await handle.getFile();
        readFileWithHandle(file, handle);
      }
      return;
    } catch (err) {
      if (err.name === 'AbortError') return; // ユーザーがダイアログを閉じた場合は無視
      console.warn('showOpenFilePicker failed, falling back to input:', err);
    }
  }
  // フォールバック: input[type=file]
  if (fileInput) fileInput.click();
}

// ファイル選択ハンドラ (input[type=file] からの読み込み)
function handleFileSelect(e) {
  if (e.target.files && e.target.files.length > 0) {
    console.log(`Selected ${e.target.files.length} file(s)`);
    loadFiles(e.target.files);
    e.target.value = ''; // リセットして同じファイルも再選択可能にする
  }
}

// 複数ファイル読み込み (Handleなし)
function loadFiles(fileList) {
  Array.from(fileList).forEach(file => {
    readFileWithHandle(file, null);
  });
}

// Handle 付きファイル読み込み処理
function readFileWithHandle(file, fileHandle = null) {
  console.log(`Reading file: ${file.name}, size: ${file.size}, hasHandle: ${!!fileHandle}`);
  const reader = new FileReader();
  reader.onload = (event) => {
    try {
      const buffer = event.target.result;
      const uint8 = new Uint8Array(buffer);
      parseAndCreateTab(file.name, uint8, fileHandle);
    } catch (err) {
      console.error("Error parsing file:", err);
      alert(`ファイル「${file.name}」の読み込み中にエラーが発生しました:\n${err.message}`);
    }
  };
  reader.onerror = (err) => {
    console.error("FileReader error:", err);
    alert(`ファイル「${file.name}」を開けませんでした。`);
  };
  reader.readAsArrayBuffer(file);
}

// 文字コード判定
function detectEncoding(uint8Array) {
  if (typeof Encoding !== 'undefined') {
    const detected = Encoding.detect(uint8Array);
    console.log("Detected encoding by encoding.js:", detected);
    if (detected === 'SJIS') return 'Shift_JIS';
    if (detected === 'UTF8') return 'UTF-8';
    if (detected === 'EUCJP') return 'EUC-JP';
    if (detected === 'ASCII') return 'UTF-8';
  }
  return 'UTF-8';
}

// バイト配列を文字列にデコード
function decodeBytes(uint8Array, encoding) {
  // BOMのスキップ判定
  let dataToDecode = uint8Array;
  if (uint8Array.length >= 3 && uint8Array[0] === 0xEF && uint8Array[1] === 0xBB && uint8Array[2] === 0xBF) {
    dataToDecode = uint8Array.subarray(3);
  }

  try {
    if (encoding === 'Shift_JIS' || encoding === 'SJIS') {
      return new TextDecoder('shift_jis').decode(dataToDecode);
    } else if (encoding === 'EUC-JP') {
      return new TextDecoder('euc-jp').decode(dataToDecode);
    } else {
      return new TextDecoder('utf-8').decode(dataToDecode);
    }
  } catch (err) {
    console.warn("TextDecoder fallback to encoding.js", err);
    if (typeof Encoding !== 'undefined') {
      const unicodeArray = Encoding.convert(dataToDecode, 'UNICODE', encoding);
      return Encoding.codeToString(unicodeArray);
    }
    return new TextDecoder().decode(dataToDecode);
  }
}

// 文字列を指定エンコーディングの Uint8Array に変換
function encodeString(text, encoding) {
  if (typeof Encoding !== 'undefined') {
    if (encoding === 'Shift_JIS' || encoding === 'SJIS') {
      const sjisArray = Encoding.convert(Encoding.stringToCode(text), 'SJIS', 'UNICODE');
      return new Uint8Array(sjisArray);
    } else if (encoding === 'EUC-JP') {
      const eucArray = Encoding.convert(Encoding.stringToCode(text), 'EUCJP', 'UNICODE');
      return new Uint8Array(eucArray);
    }
  }

  if (encoding === 'UTF-8-BOM') {
    const utf8Bytes = new TextEncoder().encode(text);
    const result = new Uint8Array(3 + utf8Bytes.length);
    result.set([0xEF, 0xBB, 0xBF], 0);
    result.set(utf8Bytes, 3);
    return result;
  } else {
    return new TextEncoder().encode(text);
  }
}

// CSV/TSV パースとタブ作成
function parseAndCreateTab(fileName, uint8Array, fileHandle = null) {
  const isBOM = uint8Array.length >= 3 && uint8Array[0] === 0xEF && uint8Array[1] === 0xBB && uint8Array[2] === 0xBF;
  const rawEncoding = detectEncoding(uint8Array);
  const currentEncoding = isBOM ? 'UTF-8-BOM' : rawEncoding;

  const text = decodeBytes(uint8Array, rawEncoding);

  // PapaParse による解析
  const parsed = Papa.parse(text, {
    skipEmptyLines: false
  });

  const delimiter = parsed.meta.delimiter || (fileName.endsWith('.tsv') ? '\t' : ',');
  const newline = parsed.meta.linebreak || '\r\n';

  const rawRows = parsed.data || [];
  
  // テーブルデータへの変換
  let headers = [];
  let rowData = [];

  if (rawRows.length > 0) {
    headers = rawRows[0];
    const numCols = headers.length;

    for (let i = 1; i < rawRows.length; i++) {
      const row = rawRows[i];
      // 末尾の完全空行をスキップ
      if (i === rawRows.length - 1 && row.length === 1 && row[0] === '') continue;
      
      const rowObj = { _id: i };
      for (let c = 0; c < numCols; c++) {
        rowObj[`col_${c}`] = (row[c] !== undefined) ? row[c] : '';
      }
      rowData.push(rowObj);
    }
  } else {
    headers = ['列 1', '列 2', '列 3'];
    rowData = [{ _id: 1, col_0: '', col_1: '', col_2: '' }];
  }

  // カラム定義の生成
  const columns = buildTabulatorColumns(headers);

  const tabId = 'tab-' + Date.now() + '-' + Math.random().toString(36).substr(2, 5);
  const tab = {
    id: tabId,
    name: fileName,
    fileHandle: fileHandle,
    rawBytes: uint8Array,
    encoding: currentEncoding,
    delimiter: delimiter,
    newline: newline,
    headers: headers,
    columns: columns,
    data: rowData,
    isModified: false,
    isNew: false
  };

  // 同名ファイルがすでに開かれているかチェック（意図しない重複オープンを防止）
  const existingTab = tabs.find(t => t.name === fileName);
  if (existingTab) {
    if (!existingTab.isModified) {
      // 既存タブが未編集なら、そのタブをアクティブにして終了
      if (fileHandle) existingTab.fileHandle = fileHandle;
      activateTab(existingTab.id);
      return;
    } else {
      // 編集済みの場合は、別名（例: sample (2).csv）にして区別
      let count = 2;
      let newName = fileName.replace(/(\.[^.]+)$/, ` (${count})$1`);
      if (newName === fileName) newName = `${fileName} (${count})`;
      while (tabs.some(t => t.name === newName)) {
        count++;
        newName = fileName.replace(/(\.[^.]+)$/, ` (${count})$1`);
        if (newName === fileName) newName = `${fileName} (${count})`;
      }
      tab.name = newName;
    }
  }

  // もし最初のタブが無編集の「無題-1.csv」なら置き換える
  if (tabs.length === 1 && tabs[0].name.startsWith('無題') && !tabs[0].isModified && tabs[0].data.length <= 3) {
    tabs = [tab];
  } else {
    tabs.push(tab);
  }

  renderTabs();
  activateTab(tabId);
}

// Excel風の列座標記号 (0 -> A, 1 -> B, 25 -> Z, 26 -> AA...)
function getColumnLetter(colIndex) {
  let letter = '';
  let temp = colIndex;
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

// Excel風の列座標記号からインデックスを逆算 (A -> 0, B -> 1, Z -> 25, AA -> 26...)
function getColumnIndexFromLetter(letterStr) {
  if (!letterStr || typeof letterStr !== 'string') return -1;
  const clean = letterStr.trim().toUpperCase();
  if (!/^[A-Z]+$/.test(clean)) return -1;
  let index = 0;
  for (let i = 0; i < clean.length; i++) {
    index = index * 26 + (clean.charCodeAt(i) - 64);
  }
  return index - 1;
}

// 2段ヘッダーフォーマッター (上段: A, B, C... 座標欄と昇順・降順・解除サイクル / 下段: 変数名)
function compositeHeaderFormatter(cell, formatterParams, onRendered) {
  const col = (cell && typeof cell.getColumn === 'function') ? cell.getColumn() : cell;
  const field = (col && typeof col.getField === 'function') ? col.getField() : '';
  const def = (col && typeof col.getDefinition === 'function') ? col.getDefinition() : {};
  const rawTitle = def.title || '';

  // 列インデックスの決定 (右方向に A → B → C...)
  let colIndex = 0;
  if (formatterParams && typeof formatterParams.colIndex === 'number') {
    colIndex = formatterParams.colIndex;
  } else if (typeof def.colIndex === 'number') {
    colIndex = def.colIndex;
  } else if (field && field.startsWith('col_')) {
    const num = parseInt(field.replace('col_', ''), 10);
    if (!isNaN(num)) colIndex = num;
  }

  const colLetter = (def && def.colLetter) ? def.colLetter : getColumnLetter(colIndex);

  // 現在のソート状態を取得
  let currentSortDir = null;
  if (currentTable) {
    const sorts = currentTable.getSorters();
    const currentSort = sorts.find(s => s.field === field);
    if (currentSort) currentSortDir = currentSort.dir; // 'asc' または 'desc'
  }

  const container = document.createElement('div');
  container.className = 'col-composite-header';

  // 上段: 座標・ソート・フィルターバー
  const coordBar = document.createElement('div');
  coordBar.className = `col-coord-bar ${currentSortDir ? 'is-sorted' : ''}`;
  coordBar.setAttribute('data-field', field);

  // 左側: 座標 & ソートトリガー（クリックで 昇順 → 降順 → 解除）
  const sortTrigger = document.createElement('div');
  sortTrigger.className = 'col-coord-main';
  if (!currentSortDir) {
    sortTrigger.title = `列 ${colLetter}: クリックで昇順に並べ替え (A→Z, 1→9)`;
  } else if (currentSortDir === 'asc') {
    sortTrigger.title = `列 ${colLetter}: 現在「昇順」 - クリックで降順に並べ替え (Z→A, 9→1)`;
  } else {
    sortTrigger.title = `列 ${colLetter}: 現在「降順」 - クリックで並べ替えを解除 (何もしない)`;
  }

  // 座標アルファベット (A, B, C...)
  const letterSpan = document.createElement('span');
  letterSpan.className = 'col-coord-letter';
  letterSpan.textContent = colLetter;
  sortTrigger.appendChild(letterSpan);

  // ソートインジケーター（当初の昇順・降順・解除サイクル表示）
  const sortIndicator = document.createElement('span');
  sortIndicator.className = `col-sort-indicator ${currentSortDir ? ('active ' + currentSortDir) : 'inactive'}`;
  if (currentSortDir === 'asc') {
    sortIndicator.innerHTML = '▲';
  } else if (currentSortDir === 'desc') {
    sortIndicator.innerHTML = '▼';
  } else {
    sortIndicator.innerHTML = '⇅';
  }
  sortTrigger.appendChild(sortIndicator);

  // クリックでソート状態をサイクル切り替え
  sortTrigger.addEventListener('click', (e) => {
    e.stopPropagation();
    e.preventDefault();
    cycleColumnSort(col);
  });

  coordBar.appendChild(sortTrigger);

  // 右側: フィルターボタン (エクセル風オートフィルター)
  const isFiltered = activeColumnFilters && activeColumnFilters.has(field);
  const btnFilter = document.createElement('button');
  btnFilter.type = 'button';
  btnFilter.className = `btn-col-filter ${isFiltered ? 'active' : ''}`;
  btnFilter.innerHTML = isFiltered ? '⧩' : '🔍';
  btnFilter.title = isFiltered
    ? `列「${rawTitle || colLetter}」でフィルター適用中 (クリックして変更または解除)`
    : `列「${rawTitle || colLetter}」を値で絞り込み (フィルター)`;
  btnFilter.addEventListener('click', (e) => {
    e.stopPropagation();
    e.preventDefault();
    openColumnFilter(e, col);
  });

  coordBar.appendChild(btnFilter);

  // 下段: 変数名バー（ダブルクリックで編集、クリックでソートは発動しない）
  const titleBar = document.createElement('div');
  titleBar.className = 'col-title-bar';
  titleBar.title = 'ダブルクリックで列名（変数名）を変更';

  const titleText = document.createElement('span');
  titleText.className = 'col-title-text';
  titleText.textContent = rawTitle;
  titleBar.appendChild(titleText);

  container.appendChild(coordBar);
  container.appendChild(titleBar);

  return container;
}

// 当初と同様の3段階サイクル切り替え: [ソートなし] → [昇順 ▲] → [降順 ▼] → [ソートなし (解除)]
function cycleColumnSort(column) {
  if (!currentTable) return;
  const field = column.getField();
  const def = column.getDefinition();
  const colTitle = def.title || field;

  const sorts = currentTable.getSorters();
  const currentSort = sorts.find(s => s.field === field);

  if (!currentSort) {
    // 1回目: 昇順に切り替え
    currentTable.setSort([{ column: field, dir: 'asc', sorter: excelSmartSorter }]);
    updateSortButtonsUI();
    showToast(`列「${colTitle}」を昇順で並べ替えました`, 'info');
  } else if (currentSort.dir === 'asc') {
    // 2回目: 降順に切り替え
    currentTable.setSort([{ column: field, dir: 'desc', sorter: excelSmartSorter }]);
    updateSortButtonsUI();
    showToast(`列「${colTitle}」を降順で並べ替えました`, 'info');
  } else {
    // 3回目: 並べ替えを解除（何もしない状態に戻す）
    currentTable.clearSort();
    updateSortButtonsUI();
    showToast(`列「${colTitle}」の並べ替えを解除しました`, 'info');
  }
}

// ソート表示の見た目を最新状態に更新
function updateSortButtonsUI() {
  if (!currentTable) return;
  const sorts = currentTable.getSorters();
  const cols = currentTable.getColumns();

  cols.forEach(c => {
    const field = c.getField();
    if (!field) return;
    const el = c.getElement();
    if (!el) return;

    const coordBar = el.querySelector('.col-coord-bar');
    const sortTrigger = el.querySelector('.col-coord-main');
    const indicator = el.querySelector('.col-sort-indicator');
    const letterSpan = el.querySelector('.col-coord-letter');
    const activeSort = sorts.find(s => s.field === field);
    const colLetter = letterSpan ? letterSpan.textContent : '';

    if (coordBar && indicator && sortTrigger) {
      if (activeSort && activeSort.dir === 'asc') {
        coordBar.classList.add('is-sorted');
        sortTrigger.title = `列 ${colLetter}: 現在「昇順」 - クリックで降順に並べ替え (Z→A)`;
        indicator.className = 'col-sort-indicator active asc';
        indicator.innerHTML = '▲';
      } else if (activeSort && activeSort.dir === 'desc') {
        coordBar.classList.add('is-sorted');
        sortTrigger.title = `列 ${colLetter}: 現在「降順」 - クリックで並べ替えを解除 (何もしない)`;
        indicator.className = 'col-sort-indicator active desc';
        indicator.innerHTML = '▼';
      } else {
        coordBar.classList.remove('is-sorted');
        sortTrigger.title = `列 ${colLetter}: クリックで昇順に並べ替え (A→Z)`;
        indicator.className = 'col-sort-indicator inactive';
        indicator.innerHTML = '⇅';
      }
    }
  });

  updateFilterButtonsUI();
}

// 列の座標記号 (A, B, C...) を最新の列順序に合わせて再描画（右方向に A → B → C... と順に振る）
function refreshColumnLetters() {
  if (!currentTable) return;
  const dataCols = currentTable.getColumns().filter(c => c.getField());
  dataCols.forEach((col, idx) => {
    const el = col.getElement();
    if (el) {
      const letterSpan = el.querySelector('.col-coord-letter');
      if (letterSpan) {
        letterSpan.textContent = getColumnLetter(idx);
      }
    }
  });
}

// 数値判定正規表現 (整数、小数、負数、指数表記、3桁カンマ区切り数値対応)
const NUMERIC_PATTERN = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/;

function isNumericValue(val) {
  if (val === null || val === undefined) return false;
  const s = String(val).trim();
  if (!s) return false;
  const clean = s.replace(/,/g, '');
  if (!NUMERIC_PATTERN.test(clean)) return false;
  return !isNaN(Number(clean));
}

// セル表示フォーマッター (文字列は左寄せ、数値は右寄せ - Excel完全互換)
function excelCellFormatter(cell, formatterParams, onRendered) {
  const val = cell.getValue();
  const el = cell.getElement();
  const isNum = isNumericValue(val);

  if (el) {
    if (isNum) {
      el.classList.add('cell-align-right');
      el.classList.remove('cell-align-left');
    } else {
      el.classList.add('cell-align-left');
      el.classList.remove('cell-align-right');
    }
  }

  return (val !== null && val !== undefined) ? String(val) : '';
}

// カンマ区切り数値、通貨記号、負数、パーセントに対応したスマートソーター (Excel完全互換)
function excelSmartSorter(a, b, aRow, bRow, column, dir, sorterParams) {
  const isAEmpty = a === null || a === undefined || a === '' || a === '(空白)';
  const isBEmpty = b === null || b === undefined || b === '' || b === '(空白)';

  if (isAEmpty && isBEmpty) return 0;
  // 空値は昇順・降順問わず常に末尾に配置 (Tabulatorの内部挙動に合わせてdirで符号調整)
  if (isAEmpty) return dir === 'desc' ? -1 : 1;
  if (isBEmpty) return dir === 'desc' ? 1 : -1;

  // 全角数字を半角化し、カンマ・通貨記号・パーセント・前後の余白を除去
  const toCleanStr = (val) => {
    if (typeof val === 'number') return val;
    let s = String(val).trim();
    s = s.replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0));
    return s.replace(/[,¥$€%\s]/g, '');
  };

  const cleanA = toCleanStr(a);
  const cleanB = toCleanStr(b);

  const numA = typeof cleanA === 'number' ? cleanA : Number(cleanA);
  const numB = typeof cleanB === 'number' ? cleanB : Number(cleanB);

  const isANum = !isNaN(numA) && cleanA !== '';
  const isBNum = !isNaN(numB) && cleanB !== '';

  if (isANum && isBNum) {
    return numA - numB;
  }
  // 数値と文字列が混在する場合：Excelと同様に数値が優先
  if (isANum) return -1;
  if (isBNum) return 1;

  // 文字列同士は自然順ソート（数値混じりの文字列も自然な順序で比較）
  return String(a).localeCompare(String(b), 'ja', { numeric: true, sensitivity: 'base' });
}

// カラム定義作成
function buildTabulatorColumns(headers) {
  // 行番号列（固定、並べ替え移動不可）
  const cols = [
    {
      formatter: "rownum",
      hozAlign: "center",
      width: 45,
      headerSort: false,
      resizable: false,
      frozen: true,
      movable: false,
      cssClass: "tabulator-rownum-cell",
      title: "#"
    }
  ];

  headers.forEach((h, idx) => {
    const letter = getColumnLetter(idx);
    cols.push({
      title: h || `列 ${idx + 1}`,
      field: `col_${idx}`,
      colIndex: idx,
      colLetter: letter,
      editor: "input",
      formatter: excelCellFormatter, // 文字列は左寄せ、数値は右寄せ
      sorter: excelSmartSorter, // カンマ区切り数値・自然順対応スマートソーター
      headerSort: false, // 勝手な昇順・降順ソートを無効化（ダブルクリックでの変数名変更を確実にする）
      titleFormatter: compositeHeaderFormatter, // 2段ヘッダー（A, B, C... 座標と変数名）
      titleFormatterParams: { colIndex: idx, colLetter: letter },
      headerTooltip: "ダブルクリックまたは右クリックで列名（変数名）を変更",
      resizable: true,
      minWidth: 75,
      headerContextMenu: getHeaderContextMenu()
    });
  });

  return cols;
}

// 列ヘッダーの右クリックメニュー定義
function getHeaderContextMenu() {
  return [
    {
      label: "✏️ 列名（変数名）を変更",
      action: (e, column) => renameColumn(column)
    },
    { separator: true },
    {
      label: "🔼 昇順で並べ替え (A→Z, 1→9)",
      action: (e, column) => sortColumn(column, 'asc')
    },
    {
      label: "🔽 降順で並べ替え (Z→A, 9→1)",
      action: (e, column) => sortColumn(column, 'desc')
    },
    {
      label: "🔄 並べ替えを解除 (元の順序に戻す)",
      action: (e, column) => clearColumnSort()
    },
    { separator: true },
    {
      label: "⬅️ 左に列を追加",
      action: (e, column) => addColumnRelative(column, 'left')
    },
    {
      label: "➡️ 右に列を追加",
      action: (e, column) => addColumnRelative(column, 'right')
    },
    { separator: true },
    {
      label: "📐 この列幅を自動調整",
      action: (e, column) => autoFitColumnWidth(column)
    },
    {
      label: "📐 すべての列幅を自動調整",
      action: (e, column) => autoFitAllColumns()
    },
    { separator: true },
    {
      label: "🔍 この列で絞り込み (フィルター)",
      action: (e, column) => openColumnFilter(e, column)
    },
    { separator: true },
    {
      label: "🔢 この列のカンマを外す (1,234 → 1234)",
      action: (e, column) => removeCommasFromColumn(column)
    },
    {
      label: "🔢 この列にカンマを付ける (1234 → 1,234)",
      action: (e, column) => addCommasToColumn(column)
    },
    { separator: true },
    {
      label: "📅 この列の日時・日付を統一 (YYYY-MM-DD)",
      action: (e, column) => cleanSpecificColumn(column, 'format_date_hyphen')
    },
    {
      label: "🔍 この列の重複値をハイライト",
      action: (e, column) => highlightDuplicatesInColumn(column)
    },
    {
      label: "⚪ この列の欠測表記を空セルに統一",
      action: (e, column) => cleanSpecificColumn(column, 'na_to_empty')
    },
    {
      label: "🏷️ この列の数値単位記号を除去",
      action: (e, column) => cleanSpecificColumn(column, 'strip_units')
    },
    { separator: true },
    {
      label: "🗑️ この列を削除",
      action: (e, column) => deleteColumn(column)
    }
  ];
}

// 列の並べ替え関数
function sortColumn(column, dir) {
  if (!currentTable) return;
  currentTable.setSort([{ column: column.getField(), dir: dir, sorter: excelSmartSorter }]);
  updateSortButtonsUI();
  showToast(`列「${column.getDefinition().title}」を${dir === 'asc' ? '昇順' : '降順'}で並べ替えました`, 'info');
}

// 並べ替え解除
function clearColumnSort() {
  if (!currentTable) return;
  currentTable.clearSort();
  updateSortButtonsUI();
  showToast('並べ替えを解除しました', 'info');
}

// 新規タブの作成
function createNewTab(customName) {
  const name = customName || `無題-${tabCounter++}.csv`;
  const headers = ['列 1', '列 2', '列 3', '列 4'];
  const columns = buildTabulatorColumns(headers);
  const rowData = [
    { _id: 1, col_0: '', col_1: '', col_2: '', col_3: '' },
    { _id: 2, col_0: '', col_1: '', col_2: '', col_3: '' },
    { _id: 3, col_0: '', col_1: '', col_2: '', col_3: '' }
  ];

  const tabId = 'tab-' + Date.now() + '-' + Math.random().toString(36).substr(2, 5);
  const tab = {
    id: tabId,
    name: name,
    fileHandle: null,
    rawBytes: null,
    encoding: 'UTF-8',
    delimiter: ',',
    newline: '\r\n',
    headers: headers,
    columns: columns,
    data: rowData,
    isModified: false,
    isNew: true
  };

  tabs.push(tab);
  renderTabs();
  activateTab(tabId);
}

// タブのレンダリング
function renderTabs() {
  if (!tabsContainer) return;
  tabsContainer.innerHTML = '';

  tabs.forEach(tab => {
    const tabEl = document.createElement('div');
    tabEl.className = 'tab-item' + (tab.id === activeTabId ? ' active' : '');
    tabEl.onclick = (e) => {
      if (!e.target.classList.contains('tab-close-btn')) {
        activateTab(tab.id);
      }
    };

    const titleEl = document.createElement('span');
    titleEl.className = 'tab-title';
    titleEl.textContent = tab.name;
    titleEl.title = tab.name; // フルパスや長いファイル名のツールチップ

    const modifiedEl = document.createElement('span');
    modifiedEl.className = 'tab-modified';
    modifiedEl.textContent = tab.isModified ? ' ●' : '';

    const closeBtn = document.createElement('button');
    closeBtn.className = 'tab-close-btn';
    closeBtn.textContent = '×';
    closeBtn.title = 'タブを閉じる';
    closeBtn.onclick = (e) => {
      e.stopPropagation();
      closeTab(tab.id);
    };

    tabEl.appendChild(titleEl);
    tabEl.appendChild(modifiedEl);
    tabEl.appendChild(closeBtn);
    tabsContainer.appendChild(tabEl);

    // アクティブなタブを可視領域にスクロール
    if (tab.id === activeTabId) {
      setTimeout(() => {
        tabEl.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'nearest' });
      }, 50);
    }
  });

  if (tabs.length === 0) {
    if (emptyState) emptyState.style.display = 'flex';
    if (tableArea) tableArea.style.display = 'none';
  } else {
    if (emptyState) emptyState.style.display = 'none';
    if (tableArea) tableArea.style.display = 'block';
  }
}

// 現在のテーブルの編集データをタブオブジェクトに同期
function syncCurrentTabData() {
  if (!currentTable || !activeTabId) return;
  const currentTab = tabs.find(t => t.id === activeTabId);
  if (!currentTab) return;

  try {
    if (typeof currentTable.getData === 'function') {
      currentTab.data = currentTable.getData();
    }
  } catch (err) {
    console.warn("syncCurrentTabData warning:", err);
  }
}

// タブのアクティブ化
function activateTab(tabId) {
  // 切り替え前に現在のアクティブタブの編集内容とフィルター状態を確実に保存
  syncCurrentTabData();
  if (activeTabId) {
    const prevTab = tabs.find(t => t.id === activeTabId);
    if (prevTab) {
      prevTab.activeFilters = new Map(activeColumnFilters);
    }
  }

  const tab = tabs.find(t => t.id === tabId);
  if (!tab) return;

  activeTabId = tabId;
  activeColumnFilters.clear();
  if (tab.activeFilters) {
    for (const [k, v] of tab.activeFilters.entries()) {
      activeColumnFilters.set(k, v);
    }
  }

  // ツールバーのコントロール更新
  if (selectEncoding) selectEncoding.value = tab.encoding;
  if (selectDelimiter) selectDelimiter.value = tab.delimiter;
  if (selectNewline) selectNewline.value = tab.newline;

  renderTabs();
  renderTableForTab(tab);
  updateStatusBar(tab);
}

// タブのテーブル描画
function renderTableForTab(tab) {
  if (!tableArea) return;

  if (currentTable) {
    try {
      currentTable.destroy();
    } catch (e) {
      console.warn("Table destroy warning:", e);
    }
    currentTable = null;
  }

  try {
    currentTable = new Tabulator(tableArea, {
      data: tab.data,
      columns: tab.columns,
      layout: "fitDataFill",
      history: true,
      clipboard: false, // 自前の高度なTSV矩形コピペエンジンを使用
      selectableRange: false, // 自前のExcel風完全互換範囲選択エンジンを使用
      movableColumns: true,
      editTriggerEvent: "dblclick",
      rowContextMenu: getRowContextMenu(),
      cellClick: (e, cell) => {
        handleCellClick(e, cell);
      },
      cellDblClick: (e, cell) => {
        if (typeof cell.edit === 'function') {
          cell.edit(true);
        }
      },
      headerDblClick: (e, col) => {
        if (!col || !col.getField()) return;
        renameColumn(col);
      }
    });

    // セル編集（入力モード）の監視とESCキャンセル連携
    currentTable.on("cellEditing", (cell) => {
      currentEditingCell = cell;
      const v = cell.getValue();
      editingOriginalValue = (v !== null && v !== undefined) ? String(v) : '';
      const isNum = isNumericValue(v);
      const el = cell.getElement();
      if (el) {
        el.classList.toggle('cell-align-right', isNum);
        el.classList.toggle('cell-align-left', !isNum);
        const input = el.querySelector('input');
        if (input) {
          input.style.textAlign = isNum ? 'right' : 'left';
        }
      }
    });

    currentTable.on("cellEditCancelled", (cell) => {
      currentEditingCell = null;
      const indices = getCellIndices(cell);
      if (indices) {
        setTimeout(() => {
          selectSingleCell(indices.rowIndex, indices.colIndex);
        }, 30);
      }
    });

    // テーブル構築完了時：初期状態を「表示モード」にする（(0,0)セルを選択）
    currentTable.on("tableBuilt", () => {
      refreshColumnLetters();
      if (activeColumnFilters.size > 0) {
        reapplyTableFilters();
      } else {
        updateSortButtonsUI();
      }
      if (!activeFocusCell) {
        const rows = getAllRows();
        const cols = getActiveColumns();
        if (rows.length > 0 && cols.length > 0) {
          selectSingleCell(0, 0);
        }
      } else {
        updateRangeHighlight();
      }
    });

    // 仮想DOMスクロール時や再描画時に範囲ハイライトを再描画
    currentTable.on("renderComplete", () => {
      refreshColumnLetters();
      updateRangeHighlight();
      updateSortButtonsUI();
    });
    currentTable.on("scrollVertical", updateRangeHighlight);
    currentTable.on("scrollHorizontal", updateRangeHighlight);

    // ソート実行時のボタン表示更新
    currentTable.on("dataSorted", () => {
      updateSortButtonsUI();
    });

    // 列の並べ替え移動
    currentTable.on("columnMoved", (column, columns) => {
      handleColumnMoved(column, columns);
    });

    // 変更イベントのリスニング（変更をタブデータに常時同期）
    currentTable.on("cellEdited", (cell) => {
      if (cell) {
        const el = cell.getElement();
        if (el) {
          const isNum = isNumericValue(cell.getValue());
          el.classList.toggle('cell-align-right', isNum);
          el.classList.toggle('cell-align-left', !isNum);
        }
      }
      tab.isModified = true;
      syncCurrentTabData();
      renderTabs();
      updateStatusBar(tab);
    });

    currentTable.on("historyUndo", () => {
      tab.isModified = true;
      syncCurrentTabData();
      renderTabs();
      updateStatusBar(tab);
    });

    currentTable.on("historyRedo", () => {
      tab.isModified = true;
      syncCurrentTabData();
      renderTabs();
      updateStatusBar(tab);
    });

    currentTable.on("dataLoaded", () => {
      updateStatusBar(tab);
    });
  } catch (err) {
    console.error("Tabulator initialization error:", err);
    tableArea.innerHTML = `<div style="padding: 20px; color: red;">テーブルの初期化に失敗しました: ${err.message}</div>`;
  }
}

// タブを閉じる
function closeTab(tabId) {
  const index = tabs.findIndex(t => t.id === tabId);
  if (index === -1) return;

  const tab = tabs[index];
  if (tab.isModified) {
    if (!confirm(`「${tab.name}」には保存されていない変更があります。閉じてもよろしいですか？`)) {
      return;
    }
  }

  tabs.splice(index, 1);

  if (tabs.length > 0) {
    const nextTab = tabs[Math.max(0, index - 1)];
    activateTab(nextTab.id);
  } else {
    activeTabId = null;
    if (currentTable) {
      try { currentTable.destroy(); } catch(e){}
      currentTable = null;
    }
    renderTabs();
    updateStatusBar(null);
  }
}

// ステータスバー更新
function updateStatusBar(tab) {
  if (!statusFilename) return;

  if (!tab || !currentTable) {
    statusFilename.textContent = 'ファイルなし';
    statusModified.style.display = 'none';
    statusRows.textContent = '0 行';
    statusCols.textContent = '0 列';
    return;
  }

  let totalRows = tab.data.length;
  let activeRows = totalRows;
  try {
    if (currentTable.getData) {
      totalRows = currentTable.getData().length;
    }
    if (currentTable.getRows) {
      activeRows = currentTable.getRows("active").length;
    }
  } catch (e) {}

  const colCount = tab.columns.filter(c => c.field).length;

  statusFilename.textContent = tab.name;
  statusModified.style.display = tab.isModified ? 'inline' : 'none';
  if (activeRows !== totalRows) {
    statusRows.textContent = `${activeRows} / ${totalRows} 行 (フィルター中)`;
  } else {
    statusRows.textContent = `${totalRows} 行`;
  }
  statusCols.textContent = `${colCount} 列`;
  statusEncoding.textContent = tab.encoding;
  statusDelimiter.textContent = tab.delimiter === '\t' ? 'TSV (タブ)' : (tab.delimiter === ';' ? 'セミコロン (;)' : 'CSV (カンマ)');
  statusNewline.textContent = tab.newline === '\r\n' ? 'CRLF (Win)' : 'LF (Unix)';
}

// エクスポート用データおよびBlob生成（3桁カンマ処理・文字コード変換）
function generateExportData(tab) {
  let data = [];
  try {
    data = currentTable.getData();
  } catch (e) {
    data = tab.data;
  }

  const activeCols = tab.columns.filter(c => c.field);
  const headers = activeCols.map(c => c.title);

  // 保存時の3桁カンマ処理オプション
  const commaMode = (selectCommaMode && selectCommaMode.value) || tab.commaMode || 'keep';

  // 2次元配列の作成
  const rows = [headers];
  data.forEach(item => {
    const row = activeCols.map(c => {
      let val = item[c.field] !== undefined ? item[c.field] : '';
      if (commaMode === 'remove') {
        val = removeCommasFromNumber(val);
      } else if (commaMode === 'add') {
        val = addCommasToNumber(val);
      }
      return val;
    });
    rows.push(row);
  });

  // PapaParse unparse（3桁カンマのない数値は "" で囲まない）
  const csvString = Papa.unparse(rows, {
    delimiter: tab.delimiter,
    newline: tab.newline,
    quotes: function(value) {
      if (value === null || value === undefined) return false;
      const str = String(value);
      // 3桁カンマのない純粋な数値（整数・小数）はクォートしない（""で囲まない）
      if (/^-?\d+(\.\d+)?$/.test(str.trim())) {
        return false;
      }
      // カンマ、ダブルクォート、改行を含む場合は必ずクォート
      if (str.includes(',') || str.includes('"') || str.includes('\n') || str.includes('\r')) {
        return true;
      }
      return false;
    }
  });

  // 指定文字コードへエンコード
  const encodedBytes = encodeString(csvString, tab.encoding);
  const isTsv = tab.delimiter === '\t';
  const mimeType = isTsv ? 'text/tab-separated-values;charset=binary;' : 'text/csv;charset=binary;';
  const blob = new Blob([encodedBytes], { type: mimeType });

  return { blob, encodedBytes, csvString };
}

// 保存の統括処理（保存フォーマット設定と保存方法のモーダルを表示）
function handleSave() {
  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab || !currentTable) return;

  // 保存設定と保存方法を選択するモーダルを表示
  showSaveModal(tab);
}

// 上書き保存の実行
async function saveOverwriteCurrentTab() {
  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab || !currentTable) return;

  // 既存のファイルハンドルがある場合は直接上書き
  if (tab.fileHandle) {
    try {
      const { blob } = generateExportData(tab);
      const writable = await tab.fileHandle.createWritable();
      await writable.write(blob);
      await writable.close();

      tab.isModified = false;
      tab.isNew = false;
      renderTabs();
      updateStatusBar(tab);
      showToast(`「${tab.name}」に上書き保存しました`, 'success');
      return;
    } catch (err) {
      if (err.name === 'AbortError') return;
      console.warn("Direct overwrite failed, falling back to picker:", err);
    }
  }

  // ファイルハンドルがない場合（D&D等で開いたファイル）は、現在のファイル名を初期値として保存ダイアログを開く
  // ユーザーが既存の同名ファイルを選択すればOSが上書き確認して上書き完了
  await saveAsCurrentTab(tab.name);
}

// 別名で保存 / 新規保存 (suggestedName: 初期ファイル名候補)
async function saveAsCurrentTab(suggestedName = null) {
  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab || !currentTable) return;

  const targetName = suggestedName || tab.name;
  const { blob } = generateExportData(tab);

  // File System Access API (showSaveFilePicker) による保存
  if ('showSaveFilePicker' in window) {
    try {
      const isTsv = tab.delimiter === '\t' || targetName.endsWith('.tsv');
      const ext = isTsv ? '.tsv' : '.csv';
      const mime = isTsv ? 'text/tab-separated-values' : 'text/csv';
      const desc = isTsv ? 'TSVファイル (*.tsv)' : 'CSVファイル (*.csv)';

      const handle = await window.showSaveFilePicker({
        suggestedName: targetName,
        types: [
          {
            description: desc,
            accept: { [mime]: [ext] }
          },
          {
            description: 'テキストファイル (*.txt)',
            accept: { 'text/plain': ['.txt'] }
          }
        ]
      });

      const writable = await handle.createWritable();
      await writable.write(blob);
      await writable.close();

      // 新しいファイル名とハンドルを反映
      tab.fileHandle = handle;
      tab.name = handle.name;
      tab.isNew = false;
      tab.isModified = false;
      renderTabs();
      updateStatusBar(tab);
      showToast(`「${tab.name}」として保存しました`, 'success');
      return;
    } catch (err) {
      if (err.name === 'AbortError') return; // ユーザーがキャンセルした場合は何もしない
      console.warn("showSaveFilePicker failed, falling back to browser download:", err);
    }
  }

  // フォールバック: ブラウザ標準ダウンロード
  let fileName = prompt('保存するファイル名を入力してください:', targetName);
  if (!fileName) return; // キャンセル

  tab.name = fileName;
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = tab.name;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  tab.isNew = false;
  tab.isModified = false;
  renderTabs();
  updateStatusBar(tab);
  showToast(`「${tab.name}」をダウンロードしました`, 'success');
}

// 保存フォーマット設定を現在のアクティブタブに反映
function applySaveSettingsToCurrentTab() {
  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab) return;
  if (selectEncoding && selectEncoding.value) {
    tab.encoding = selectEncoding.value;
  }
  if (selectDelimiter && selectDelimiter.value) {
    tab.delimiter = selectDelimiter.value;
  }
  if (selectNewline && selectNewline.value) {
    tab.newline = selectNewline.value;
  }
  if (selectCommaMode && selectCommaMode.value) {
    tab.commaMode = selectCommaMode.value;
  }
  updateStatusBar(tab);
}

// 保存ダイアログ（保存設定・保存方法選択）の表示・非表示
function showSaveModal(tab) {
  if (!saveModal) return;
  const filenameEl = document.getElementById('modal-filename');
  if (filenameEl && tab) {
    filenameEl.textContent = tab.name;
  }

  // 現在のタブの設定をモーダル内のセレクトボックスに初期反映
  if (tab) {
    if (selectEncoding && tab.encoding) selectEncoding.value = tab.encoding;
    if (selectDelimiter && tab.delimiter) selectDelimiter.value = tab.delimiter;
    if (selectNewline && tab.newline) selectNewline.value = tab.newline;
    if (selectCommaMode && tab.commaMode) selectCommaMode.value = tab.commaMode;
  }

  // 新規作成ファイルか既存ファイルかで保存ボタングループを切り替え
  if (saveModalExistingActions && saveModalNewActions) {
    if (tab && tab.isNew) {
      saveModalExistingActions.style.display = 'none';
      saveModalNewActions.style.display = 'flex';
    } else {
      saveModalExistingActions.style.display = 'flex';
      saveModalNewActions.style.display = 'none';
    }
  }

  saveModal.style.display = 'flex';
}

function hideSaveModal() {
  if (saveModal) {
    saveModal.style.display = 'none';
  }
}

// トースト通知の表示
function showToast(message, type = 'info') {
  if (!toastContainer) {
    toastContainer = document.getElementById('toast-container');
  }
  if (!toastContainer) return;

  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  const icon = type === 'success' ? '✅' : (type === 'error' ? '⚠️' : 'ℹ️');
  toast.innerHTML = `<span>${icon}</span><span>${message}</span>`;
  toastContainer.appendChild(toast);

  setTimeout(() => {
    if (toast && toast.parentNode) {
      toast.parentNode.removeChild(toast);
    }
  }, 4000);
}

// 3桁カンマ関連の処理
function addCommasToNumber(value) {
  if (value === null || value === undefined) return value;
  const str = String(value).trim();
  if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(str)) return str;
  if (/^-?\d{4,}(\.\d+)?$/.test(str)) {
    const parts = str.split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return parts.join('.');
  }
  return value;
}

function removeCommasFromNumber(value) {
  if (value === null || value === undefined) return value;
  const str = String(value).trim();
  if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(str)) {
    return str.replace(/,/g, '');
  }
  return value;
}

// 表全体の3桁カンマを外す
function removeCommasFromTable() {
  if (!currentTable || !activeTabId) return;
  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab) return;

  const data = currentTable.getData();
  const activeFields = tab.columns.filter(c => c.field).map(c => c.field);
  let changedCount = 0;

  data.forEach(row => {
    activeFields.forEach(field => {
      const original = row[field];
      const updated = removeCommasFromNumber(original);
      if (original !== updated) {
        row[field] = updated;
        changedCount++;
      }
    });
  });

  if (changedCount > 0) {
    currentTable.replaceData(data);
    tab.isModified = true;
    syncCurrentTabData();
    renderTabs();
    updateStatusBar(tab);
    alert(`${changedCount} 箇所の数値から3桁カンマを外しました。\n（保存時に "" で囲まれなくなります）`);
  } else {
    alert("3桁カンマを含む数値は見つかりませんでした。");
  }
}

// 表全体に3桁カンマを付ける
function addCommasToTable() {
  if (!currentTable || !activeTabId) return;
  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab) return;

  const data = currentTable.getData();
  const activeFields = tab.columns.filter(c => c.field).map(c => c.field);
  let changedCount = 0;

  data.forEach(row => {
    activeFields.forEach(field => {
      const original = row[field];
      const updated = addCommasToNumber(original);
      if (original !== updated) {
        row[field] = updated;
        changedCount++;
      }
    });
  });

  if (changedCount > 0) {
    currentTable.replaceData(data);
    tab.isModified = true;
    syncCurrentTabData();
    renderTabs();
    updateStatusBar(tab);
    alert(`${changedCount} 箇所の数値に3桁カンマを付けました。\n（保存時に "" で囲まれます）`);
  } else {
    alert("3桁カンマを付与できる4桁以上の数値は見つかりませんでした。");
  }
}

// 特定の列のカンマを外す
function removeCommasFromColumn(column) {
  if (!currentTable || !activeTabId) return;
  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab) return;

  const field = column.getField();
  const data = currentTable.getData();
  let changedCount = 0;

  data.forEach(row => {
    const original = row[field];
    const updated = removeCommasFromNumber(original);
    if (original !== updated) {
      row[field] = updated;
      changedCount++;
    }
  });

  if (changedCount > 0) {
    currentTable.replaceData(data);
    tab.isModified = true;
    syncCurrentTabData();
    renderTabs();
    updateStatusBar(tab);
  }
}

// 特定の列にカンマを付ける
function addCommasToColumn(column) {
  if (!currentTable || !activeTabId) return;
  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab) return;

  const field = column.getField();
  const data = currentTable.getData();
  let changedCount = 0;

  data.forEach(row => {
    const original = row[field];
    const updated = addCommasToNumber(original);
    if (original !== updated) {
      row[field] = updated;
      changedCount++;
    }
  });

  if (changedCount > 0) {
    currentTable.replaceData(data);
    tab.isModified = true;
    syncCurrentTabData();
    renderTabs();
    updateStatusBar(tab);
  }
}

// 行・列の操作

// 行の追加（位置指定プロンプト付き）
function addRowPrompt() {
  if (!currentTable || !activeTabId) return;
  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab) return;

  const allRows = currentTable.getRows();
  const maxPos = allRows.length + 1;

  let defaultPos = maxPos;
  if (activeFocusCell && activeFocusCell.rowIndex !== undefined && activeFocusCell.rowIndex >= 0 && activeFocusCell.rowIndex < allRows.length) {
    defaultPos = activeFocusCell.rowIndex + 1;
  }

  const input = prompt(
    `追加する行の位置（行番号 1〜${maxPos}）を指定してください:\n（※ 1: 先頭に行を挿入、${maxPos}: 末尾に行を追加）`,
    String(defaultPos)
  );
  if (input === null) return;

  const pos = parseInt(input.trim(), 10);
  if (isNaN(pos) || pos < 1 || pos > maxPos) {
    showToast(`有効な行番号（1〜${maxPos}）を入力してください`, 'error');
    return;
  }

  const confirmMsg = (pos <= allRows.length)
    ? `行 ${pos} の位置に行を追加（挿入）してもよろしいですか？`
    : `末尾（行 ${pos}）に行を追加してもよろしいですか？`;

  if (!confirm(confirmMsg)) return;

  const newRow = { _id: Date.now() + Math.random() };
  tab.columns.filter(c => c.field).forEach(c => newRow[c.field] = '');

  if (pos > allRows.length) {
    currentTable.addRow(newRow, false);
  } else {
    const targetRow = allRows[pos - 1];
    currentTable.addRow(newRow, true, targetRow);
  }

  tab.isModified = true;
  syncCurrentTabData();
  renderTabs();
  updateStatusBar(tab);

  setTimeout(() => {
    selectSingleCell(pos - 1, activeFocusCell ? activeFocusCell.colIndex : 0);
  }, 30);

  showToast(`行 ${pos} を追加しました`, 'success');
}

// 行の削除（番号指定プロンプト付き）
function deleteRowPrompt() {
  if (!currentTable || !activeTabId) return;
  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab) return;

  const allRows = currentTable.getRows();
  if (allRows.length <= 1) {
    showToast('これ以上行を削除できません (最低1行必要です)', 'error');
    return;
  }

  let defaultRow = allRows.length;
  if (activeFocusCell && activeFocusCell.rowIndex !== undefined && activeFocusCell.rowIndex >= 0 && activeFocusCell.rowIndex < allRows.length) {
    defaultRow = activeFocusCell.rowIndex + 1;
  }

  const input = prompt(
    `削除する行番号（1〜${allRows.length}）を指定してください:`,
    String(defaultRow)
  );
  if (input === null) return;

  const rowNum = parseInt(input.trim(), 10);
  if (isNaN(rowNum) || rowNum < 1 || rowNum > allRows.length) {
    showToast(`有効な行番号（1〜${allRows.length}）を入力してください`, 'error');
    return;
  }

  if (!confirm(`行 ${rowNum}を削除してもよろしいですか？`)) return;

  const targetRow = allRows[rowNum - 1];
  if (!targetRow) return;

  targetRow.delete();

  tab.isModified = true;
  syncCurrentTabData();
  renderTabs();
  updateStatusBar(tab);

  const remainingRows = currentTable.getRows();
  const nextRowIndex = Math.min(rowNum - 1, remainingRows.length - 1);
  if (activeFocusCell) {
    activeFocusCell.rowIndex = nextRowIndex;
    if (selectionRange) {
      selectionRange.minRow = nextRowIndex;
      selectionRange.maxRow = nextRowIndex;
    }
    if (typeof updateRangeHighlight === 'function') updateRangeHighlight();
  }

  showToast(`行 ${rowNum} を削除しました`, 'info');
}

// 列の追加（位置指定プロンプト付き）
function addColumnPrompt() {
  if (!currentTable || !activeTabId) return;
  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab) return;

  const activeCols = tab.columns.filter(c => c.field);
  const maxCols = activeCols.length;
  const nextLetter = getColumnLetter(maxCols);

  let defaultLetter = nextLetter;
  if (activeFocusCell && activeFocusCell.colIndex !== undefined && activeFocusCell.colIndex >= 0 && activeFocusCell.colIndex < maxCols) {
    defaultLetter = getColumnLetter(activeFocusCell.colIndex + 1);
  }

  const input = prompt(
    `追加する列の位置（列記号 A〜${nextLetter} または 番号 1〜${maxCols + 1}）を指定してください:\n（※ 末尾に追加する場合は「${nextLetter}」）`,
    defaultLetter
  );
  if (input === null) return;

  const trimmed = input.trim();
  let targetIndex = -1;

  if (/^\d+$/.test(trimmed)) {
    const num = parseInt(trimmed, 10);
    if (num >= 1 && num <= maxCols + 1) {
      targetIndex = num - 1;
    }
  } else {
    const letterIdx = getColumnIndexFromLetter(trimmed);
    if (letterIdx >= 0 && letterIdx <= maxCols) {
      targetIndex = letterIdx;
    }
  }

  if (targetIndex < 0 || targetIndex > maxCols) {
    showToast(`有効な列位置（A〜${nextLetter}）を指定してください`, 'error');
    return;
  }

  const targetLetter = getColumnLetter(targetIndex);
  const confirmMsg = (targetIndex < maxCols)
    ? `列 ${targetLetter} の位置に新しい列を追加（挿入）してもよろしいですか？`
    : `末尾（列 ${targetLetter}）に新しい列を追加してもよろしいですか？`;

  if (!confirm(confirmMsg)) return;

  const colTimestamp = Date.now();
  const newField = `col_${colTimestamp}`;
  const newTitle = `列 ${maxCols + 1}`;

  const newColDef = {
    title: newTitle,
    field: newField,
    colIndex: targetIndex,
    colLetter: targetLetter,
    editor: "input",
    formatter: excelCellFormatter,
    sorter: excelSmartSorter,
    headerSort: false,
    titleFormatter: compositeHeaderFormatter,
    titleFormatterParams: { colIndex: targetIndex, colLetter: targetLetter },
    headerTooltip: "ダブルクリックまたは右クリックで列名（変数名）を変更",
    resizable: true,
    minWidth: 80,
    headerContextMenu: getHeaderContextMenu()
  };

  if (targetIndex >= maxCols) {
    tab.columns.push(newColDef);
    currentTable.addColumn(newColDef);
  } else {
    const targetColumn = currentTable.getColumn(activeCols[targetIndex].field);
    const spliceIdx = tab.columns.findIndex(c => c.field === activeCols[targetIndex].field);
    if (spliceIdx !== -1) {
      tab.columns.splice(spliceIdx, 0, newColDef);
    } else {
      tab.columns.push(newColDef);
    }
    currentTable.addColumn(newColDef, true, targetColumn);
  }

  tab.headers = tab.columns.filter(c => c.field).map(c => c.title);
  tab.isModified = true;
  renderTabs();
  updateStatusBar(tab);

  setTimeout(() => {
    refreshColumnLetters();
    updateSortButtonsUI();
    selectSingleCell(activeFocusCell ? activeFocusCell.rowIndex : 0, targetIndex);
  }, 50);

  showToast(`列 ${targetLetter} を追加しました`, 'success');
}

// 列の削除（記号/列名指定プロンプト付き）
function deleteColumnPrompt() {
  if (!currentTable || !activeTabId) return;
  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab) return;

  const activeCols = tab.columns.filter(c => c.field);
  if (activeCols.length <= 1) {
    showToast('これ以上列を削除できません (最低1列必要です)', 'error');
    return;
  }

  const lastLetter = getColumnLetter(activeCols.length - 1);

  let defaultLetter = lastLetter;
  if (activeFocusCell && activeFocusCell.colIndex !== undefined && activeFocusCell.colIndex >= 0 && activeFocusCell.colIndex < activeCols.length) {
    defaultLetter = getColumnLetter(activeFocusCell.colIndex);
  }

  const input = prompt(
    `削除する列（列記号 A〜${lastLetter} または 列名）を指定してください:`,
    defaultLetter
  );
  if (input === null) return;

  const trimmed = input.trim();
  let targetColDef = null;
  let targetIndex = -1;

  const letterIdx = getColumnIndexFromLetter(trimmed);
  if (letterIdx >= 0 && letterIdx < activeCols.length) {
    targetIndex = letterIdx;
    targetColDef = activeCols[targetIndex];
  }

  if (!targetColDef && /^\d+$/.test(trimmed)) {
    const num = parseInt(trimmed, 10);
    if (num >= 1 && num <= activeCols.length) {
      targetIndex = num - 1;
      targetColDef = activeCols[targetIndex];
    }
  }

  if (!targetColDef) {
    const foundIdx = activeCols.findIndex(c => c.title.toLowerCase() === trimmed.toLowerCase());
    if (foundIdx !== -1) {
      targetIndex = foundIdx;
      targetColDef = activeCols[foundIdx];
    }
  }

  if (!targetColDef) {
    showToast(`指定された列「${trimmed}」が見つかりませんでした`, 'error');
    return;
  }

  const targetColumn = currentTable.getColumn(targetColDef.field);
  if (!targetColumn) return;

  const colTitle = targetColDef.title || `列 ${targetIndex + 1}`;
  const letter = getColumnLetter(targetIndex);
  const colLabel = `列 ${letter}「${colTitle}」`;

  if (!confirm(`${colLabel}を削除してもよろしいですか？`)) return;

  const field = targetColDef.field;
  targetColumn.delete();

  tab.columns = tab.columns.filter(c => c.field !== field);
  tab.headers = tab.columns.filter(c => c.field).map(c => c.title);
  tab.isModified = true;

  const remainingCols = tab.columns.filter(c => c.field);
  const nextColIndex = Math.min(targetIndex, remainingCols.length - 1);
  if (activeFocusCell) {
    activeFocusCell.colIndex = nextColIndex;
    if (selectionRange) {
      selectionRange.minCol = Math.min(selectionRange.minCol, remainingCols.length - 1);
      selectionRange.maxCol = Math.min(selectionRange.maxCol, remainingCols.length - 1);
    }
    if (typeof updateRangeHighlight === 'function') updateRangeHighlight();
  }

  renderTabs();
  updateStatusBar(tab);
  setTimeout(() => {
    refreshColumnLetters();
    updateSortButtonsUI();
  }, 50);

  showToast(`${colLabel} を削除しました`, 'info');
}

function addRowBottom() {
  addRowPrompt();
}

function addRowRelative(targetRow, position) {
  if (!currentTable) return;
  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab) return;

  const newRow = { _id: Date.now() };
  tab.columns.filter(c => c.field).forEach(c => newRow[c.field] = '');
  currentTable.addRow(newRow, position === 'above', targetRow);
  tab.isModified = true;
  syncCurrentTabData();
  renderTabs();
  updateStatusBar(tab);
}

function duplicateRow(targetRow) {
  if (!currentTable) return;
  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab) return;

  const data = Object.assign({}, targetRow.getData(), { _id: Date.now() });
  currentTable.addRow(data, false, targetRow);
  tab.isModified = true;
  syncCurrentTabData();
  renderTabs();
  updateStatusBar(tab);
}

function deleteRow(targetRow, skipConfirm = false) {
  if (!currentTable) return;
  const allRows = currentTable.getRows();
  if (allRows.length <= 1) {
    showToast('これ以上行を削除できません (最低1行必要です)', 'error');
    return;
  }

  // 複数行が範囲選択されている場合はまとめて削除
  if (selectionRange && selectionRange.minRow !== undefined && selectionRange.maxRow !== undefined) {
    const minRow = selectionRange.minRow;
    const maxRow = selectionRange.maxRow;
    const count = maxRow - minRow + 1;
    if (count > 1) {
      if (allRows.length - count < 1) {
        showToast('すべての行を削除することはできません (最低1行必要です)', 'error');
        return;
      }
      const msg = `選択中の 行 ${minRow + 1} 〜 行 ${maxRow + 1} (${count}行) を削除してもよろしいですか？`;
      if (!skipConfirm && !confirm(msg)) return;

      const rowsToDelete = [];
      for (let r = minRow; r <= maxRow; r++) {
        if (allRows[r]) rowsToDelete.push(allRows[r]);
      }
      rowsToDelete.forEach(r => r.delete());

      const tab = tabs.find(t => t.id === activeTabId);
      if (tab) {
        tab.isModified = true;
        syncCurrentTabData();
        renderTabs();
        updateStatusBar(tab);
      }
      showToast(`${count} 行を削除しました`, 'info');
      return;
    }
  }

  const rowPos = (typeof targetRow.getPosition === 'function') ? targetRow.getPosition(true) : null;
  const rowLabel = rowPos ? `行 ${rowPos}` : 'この行';
  if (!skipConfirm && !confirm(`${rowLabel}を削除してもよろしいですか？`)) return;

  targetRow.delete();
  const tab = tabs.find(t => t.id === activeTabId);
  if (tab) {
    tab.isModified = true;
    syncCurrentTabData();
    renderTabs();
    updateStatusBar(tab);
    if (!skipConfirm) showToast(`${rowLabel} を削除しました`, 'info');
  }
}

function deleteSelectedOrLastRow() {
  deleteRowPrompt();
}

function deleteSelectedOrLastColumn() {
  deleteColumnPrompt();
}

function addColumnRight() {
  addColumnPrompt();
}

function addColumnRelative(targetColumn, position) {
  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab || !currentTable) return;

  const colIndex = Date.now();
  const newField = `col_${colIndex}`;
  const newTitle = prompt("新しい列名を入力してください:", `列 ${tab.columns.length}`);
  if (newTitle === null) return;

  const targetField = targetColumn.getField();
  const idx = tab.columns.findIndex(c => c.field === targetField);
  const letterIdx = idx !== -1 ? (position === 'left' ? idx : idx + 1) : tab.columns.length;
  const letter = getColumnLetter(letterIdx);

  const newColDef = {
    title: newTitle || `列 ${tab.columns.length}`,
    field: newField,
    colIndex: letterIdx,
    colLetter: letter,
    editor: "input",
    formatter: excelCellFormatter,
    sorter: excelSmartSorter,
    headerSort: false,
    titleFormatter: compositeHeaderFormatter,
    titleFormatterParams: { colIndex: letterIdx, colLetter: letter },
    headerTooltip: "ダブルクリックまたは右クリックで列名（変数名）を変更",
    resizable: true,
    minWidth: 80,
    headerContextMenu: getHeaderContextMenu()
  };

  if (idx !== -1) {
    const insertIdx = position === 'left' ? idx : idx + 1;
    tab.columns.splice(insertIdx, 0, newColDef);
  } else {
    tab.columns.push(newColDef);
  }

  currentTable.addColumn(newColDef, position === 'left', targetColumn);
  tab.isModified = true;
  renderTabs();
  updateStatusBar(tab);
  setTimeout(() => {
    refreshColumnLetters();
    updateSortButtonsUI();
  }, 50);
}

function renameColumn(column) {
  const currentTitle = column.getDefinition().title;
  const newTitle = prompt("列名（変数名）を変更:", currentTitle);
  if (newTitle !== null && newTitle.trim() !== "") {
    const trimmedTitle = newTitle.trim();
    column.updateDefinition({ title: trimmedTitle });
    const tab = tabs.find(t => t.id === activeTabId);
    if (tab) {
      const colDef = tab.columns.find(c => c.field === column.getField());
      if (colDef) colDef.title = trimmedTitle;
      const activeCols = tab.columns.filter(c => c.field);
      tab.headers = activeCols.map(c => c.title);
      tab.isModified = true;
      renderTabs();
      updateStatusBar(tab);
    }
    const el = column.getElement();
    if (el) {
      const titleText = el.querySelector('.col-title-text');
      if (titleText) titleText.textContent = trimmedTitle;
    }
    showToast(`列名を「${trimmedTitle}」に変更しました`, 'info');
  }
}

function deleteColumn(column) {
  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab || !currentTable) return;
  const activeCols = tab.columns.filter(c => c.field);
  if (activeCols.length <= 1) {
    showToast('これ以上列を削除できません (最低1列必要です)', 'error');
    return;
  }

  const colTitle = column.getDefinition().title || 'この列';
  const letter = column.getDefinition().colLetter || '';
  const colLabel = letter ? `列 ${letter}「${colTitle}」` : `列「${colTitle}」`;

  if (!confirm(`${colLabel}を削除してもよろしいですか？`)) return;
  const field = column.getField();
  column.delete();

  tab.columns = tab.columns.filter(c => c.field !== field);
  tab.headers = tab.columns.filter(c => c.field).map(c => c.title);
  tab.isModified = true;

  const remainingCols = tab.columns.filter(c => c.field);
  if (activeFocusCell) {
    activeFocusCell.colIndex = Math.min(activeFocusCell.colIndex, remainingCols.length - 1);
    if (selectionRange) {
      selectionRange.minCol = Math.min(selectionRange.minCol, remainingCols.length - 1);
      selectionRange.maxCol = Math.min(selectionRange.maxCol, remainingCols.length - 1);
    }
    if (typeof updateRangeHighlight === 'function') updateRangeHighlight();
  }

  renderTabs();
  updateStatusBar(tab);
  setTimeout(() => {
    refreshColumnLetters();
    updateSortButtonsUI();
  }, 50);

  showToast(`${colLabel} を削除しました`, 'info');
}

// ----------------------------------------------------
// 共通ヘルパー: ワイルドカード / 正規表現 スマートマッチャー & フィルター値正規化 & HTMLエスケープ
// ----------------------------------------------------

/**
 * HTMLエスケープヘルパー
 */
function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .split('&').join('&amp;')
    .split('<').join('&lt;')
    .split('>').join('&gt;')
    .split('"').join('&quot;')
    .split("'").join('&#39;');
}

/**
 * フィルター用値の正規化 (前後の空白・改行コードをトリムし、空または空白のみは '(空白)' に統一)
 */
function normalizeFilterValue(v) {
  if (v === null || v === undefined) return '(空白)';
  const s = String(v).trim();
  return s === '' ? '(空白)' : s;
}

/**
 * ワイルドカードパターン (*, ?, ~*, ~?, ~~) を正規表現文字列に変換する
 */
function wildcardToRegexStr(pat) {
  let res = '';
  for (let i = 0; i < pat.length; i++) {
    const ch = pat[i];
    if (ch === '~' && i + 1 < pat.length) {
      const next = pat[i + 1];
      if (next === '*' || next === '?' || next === '~') {
        res += '\\' + next;
        i++;
        continue;
      }
    }
    if (ch === '*') {
      res += '.*';
    } else if (ch === '?') {
      res += '.';
    } else if (/[.\\+^$()[\]{}|]/.test(ch)) {
      res += '\\' + ch;
    } else {
      res += ch;
    }
  }
  return res;
}

/**
 * クエリ文字列から数値条件（範囲・不等号）を抽出する
 * 返り値: { type: 'range', min, max } または { type: 'cmp', op, value } または null
 */
function parseNumericCondition(query) {
  if (!query || typeof query !== 'string') return null;
  const s = query.trim().replace(/,/g, '');

  // 1. 範囲指定: min..max, min~max, min〜max
  const rangeMatch = s.match(/^(-?\d+(?:\.\d+)?)\s*(?:\.\.|\~|\〜|から)\s*(-?\d+(?:\.\d+)?)(?:以下|まで)?$/);
  if (rangeMatch) {
    const min = parseFloat(rangeMatch[1]);
    const max = parseFloat(rangeMatch[2]);
    return { type: 'range', min: Math.min(min, max), max: Math.max(min, max) };
  }

  // 2. 日本語の範囲: 100以上200以下
  const jpRangeMatch = s.match(/^(-?\d+(?:\.\d+)?)\s*以上\s*(-?\d+(?:\.\d+)?)\s*以下$/);
  if (jpRangeMatch) {
    const min = parseFloat(jpRangeMatch[1]);
    const max = parseFloat(jpRangeMatch[2]);
    return { type: 'range', min: Math.min(min, max), max: Math.max(min, max) };
  }

  // 3. ハイフン範囲: min-max (正数同士)
  const hyphenRange = s.match(/^(\d+(?:\.\d+)?)\s*-\s*(\d+(?:\.\d+)?)$/);
  if (hyphenRange) {
    const min = parseFloat(hyphenRange[1]);
    const max = parseFloat(hyphenRange[2]);
    return { type: 'range', min: Math.min(min, max), max: Math.max(min, max) };
  }

  // 4. 不等号: <=, >=, <, >, =, !=, <>
  const cmpMatch = s.match(/^([<>!=]=?|<>)\s*(-?\d+(?:\.\d+)?)$/);
  if (cmpMatch) {
    const op = cmpMatch[1];
    const val = parseFloat(cmpMatch[2]);
    return { type: 'cmp', op, value: val };
  }

  // 5. 日本語不等号: 100以下, 100以上, 100未満, 100超
  const jpCmpMatch = s.match(/^(-?\d+(?:\.\d+)?)\s*(以下|以上|未満|超|より大きい|より小さい)$/);
  if (jpCmpMatch) {
    const val = parseFloat(jpCmpMatch[1]);
    const term = jpCmpMatch[2];
    let op = '<=';
    if (term === '以下' || term === 'より小さい') op = '<=';
    else if (term === '以上') op = '>=';
    else if (term === '未満') op = '<';
    else if (term === '超' || term === 'より大きい') op = '>';
    return { type: 'cmp', op, value: val };
  }

  return null;
}

/**
 * 文字列値が数値条件に合致するか判定する (カンマ区切り対応)
 */
function testNumericCondition(cond, textVal) {
  if (textVal === null || textVal === undefined || textVal === '(空白)') return false;
  const clean = String(textVal).replace(/,/g, '').trim();
  const num = parseFloat(clean);
  if (isNaN(num)) return false;

  if (cond.type === 'range') {
    return num >= cond.min && num <= cond.max;
  }
  if (cond.type === 'cmp') {
    switch (cond.op) {
      case '<=': case '=<': return num <= cond.value;
      case '>=': case '=>': return num >= cond.value;
      case '<': return num < cond.value;
      case '>': return num > cond.value;
      case '=': case '==': return num === cond.value;
      case '!=': case '<>': return num !== cond.value;
      default: return false;
    }
  }
  return false;
}

/**
 * 検索文字列とモード（'auto' | 'wildcard' | 'regex' | 'normal'）から
 * 値の合否を判定するマッチャー関数 (val: any) => boolean を生成する
 */
function createSmartMatcher(query, mode = 'auto', options = {}) {
  const { caseSensitive = false, exactMatch = false } = options;
  if (!query || typeof query !== 'string') {
    return () => true;
  }

  const trimmed = query.trim();
  if (!trimmed) {
    return () => true;
  }

  const flags = caseSensitive ? '' : 'i';
  let effectiveMode = mode;
  let regexPattern = null;
  let customFlags = flags;

  if (effectiveMode === 'auto') {
    // 0. 数値条件（<=100, 100..200, 100~200, 100以下 等）の自動検出
    const numCond = parseNumericCondition(trimmed);
    if (numCond) {
      return (val) => testNumericCondition(numCond, val);
    }

    // 1. /pattern/flags スラッシュ記法の検出
    const slashMatch = trimmed.match(/^\/(.+)\/([gimsuy]*)$/);
    if (slashMatch) {
      effectiveMode = 'regex';
      regexPattern = slashMatch[1];
      customFlags = slashMatch[2] || (caseSensitive ? '' : 'i');
    } else if (trimmed.includes('*') || trimmed.includes('?')) {
      effectiveMode = 'wildcard';
    } else {
      // 2. OR条件の検出 (例: "ノート or デスクトップ", "東京 OR 大阪", "A, B", "A、B")
      const orParts = trimmed.split(/[\s　]+(?:or|OR)[\s　]+|[,、]/).map(p => p.trim()).filter(p => p.length > 0);
      if (orParts.length > 1) {
        const subMatchers = orParts.map(p => createSmartMatcher(p, 'auto', options));
        return (val) => subMatchers.some(m => m(val));
      }

      // 3. AND条件の検出 (例: "東京 and 営業", "ノート AND パソコン")
      const andParts = trimmed.split(/[\s　]+(?:and|AND)[\s　]+/).map(p => p.trim()).filter(p => p.length > 0);
      if (andParts.length > 1) {
        const subMatchers = andParts.map(p => createSmartMatcher(p, 'auto', options));
        return (val) => subMatchers.every(m => m(val));
      }

      // 4. 正規表現記号のチェック
      if (/[\\^$|()[\]{}+]/.test(trimmed)) {
        try {
          new RegExp(trimmed, flags);
          effectiveMode = 'regex';
          regexPattern = trimmed;
        } catch (e) {
          effectiveMode = 'normal';
        }
      } else {
        effectiveMode = 'normal';
      }
    }
  }

  if (effectiveMode === 'regex') {
    const pat = regexPattern !== null ? regexPattern : query;
    try {
      const reg = new RegExp(pat, customFlags);
      return (val) => {
        if (val === null || val === undefined) return false;
        return reg.test(String(val));
      };
    } catch (e) {
      const qTarget = caseSensitive ? query : query.toLowerCase();
      return (val) => {
        if (val === null || val === undefined) return false;
        const s = caseSensitive ? String(val) : String(val).toLowerCase();
        return exactMatch ? s === qTarget : s.includes(qTarget);
      };
    }
  }

  if (effectiveMode === 'wildcard') {
    const hasWildcard = query.includes('*') || query.includes('?');
    const pat = wildcardToRegexStr(query);
    const fullPat = hasWildcard ? `^${pat}$` : pat;
    try {
      const reg = new RegExp(fullPat, flags);
      return (val) => {
        if (val === null || val === undefined) return false;
        return reg.test(String(val));
      };
    } catch (e) {
      const qTarget = caseSensitive ? query : query.toLowerCase();
      return (val) => {
        if (val === null || val === undefined) return false;
        const s = caseSensitive ? String(val) : String(val).toLowerCase();
        return exactMatch ? s === qTarget : s.includes(qTarget);
      };
    }
  }

  // normal
  const qTarget = caseSensitive ? query : query.toLowerCase();
  return (val) => {
    if (val === null || val === undefined) return false;
    const s = caseSensitive ? String(val) : String(val).toLowerCase();
    return exactMatch ? s === qTarget : s.includes(qTarget);
  };
}

// 検索クエリのトークナイザー（括弧 ()、OR、AND、クォート、列指定トークンに対応）
function tokenizeSearchQuery(query) {
  const tokens = [];
  let i = 0;
  const len = query.length;

  while (i < len) {
    const ch = query[i];

    if (/[\s　]/.test(ch)) {
      tokens.push({ type: 'SPACE' });
      i++;
      continue;
    }

    if (ch === '(' || ch === '（') {
      tokens.push({ type: 'LPAREN' });
      i++;
      continue;
    }
    if (ch === ')' || ch === '）') {
      tokens.push({ type: 'RPAREN' });
      i++;
      continue;
    }

    if (ch === ',' || ch === '、') {
      tokens.push({ type: 'OR' });
      i++;
      continue;
    }

    if (ch === '"' || ch === "'") {
      const quote = ch;
      i++;
      let str = '';
      while (i < len && query[i] !== quote) {
        if (query[i] === '\\' && i + 1 < len) {
          i++;
          str += query[i];
        } else {
          str += query[i];
        }
        i++;
      }
      if (i < len && query[i] === quote) i++;
      tokens.push({ type: 'LITERAL', value: str });
      continue;
    }

    let word = '';
    while (i < len && !/[\s　(),、（）"']/.test(query[i])) {
      word += query[i];
      i++;
    }

    const lower = word.toLowerCase();
    if (lower === 'or') {
      tokens.push({ type: 'OR' });
    } else if (lower === 'and') {
      tokens.push({ type: 'AND' });
    } else if (word) {
      tokens.push({ type: 'LITERAL', value: word });
    }
  }

  // 空白トークンの正規化:
  // [LITERALまたはRPAREN] SPACE [LITERALまたはLPAREN] の場合、暗黙の AND として扱う
  const normalized = [];
  for (let j = 0; j < tokens.length; j++) {
    const tok = tokens[j];
    if (tok.type === 'SPACE') {
      const prev = normalized[normalized.length - 1];
      let next = null;
      for (let k = j + 1; k < tokens.length; k++) {
        if (tokens[k].type !== 'SPACE') {
          next = tokens[k];
          break;
        }
      }
      if (prev && next) {
        const prevCanBeOperand = (prev.type === 'LITERAL' || prev.type === 'RPAREN');
        const nextCanBeOperand = (next.type === 'LITERAL' || next.type === 'LPAREN');
        if (prevCanBeOperand && nextCanBeOperand) {
          normalized.push({ type: 'AND' });
        }
      }
    } else {
      normalized.push(tok);
    }
  }

  return normalized;
}

// 検索トークンを評価する単一マッチャー関数を作成
function createSearchTokenMatcher(tokenStr, activeCols) {
  const colMatch = tokenStr.match(/^([^:=><]+)[:=](.+)$/);
  let targetField = null;
  let queryPart = tokenStr;

  if (colMatch && activeCols && activeCols.length > 0) {
    const colName = colMatch[1].trim().toLowerCase();
    const valStr = colMatch[2].trim();
    const matchedCol = activeCols.find(c => {
      const title = (c.title || '').toLowerCase();
      const field = (c.field || '').toLowerCase();
      return title === colName || title.includes(colName) || field === colName;
    });

    if (matchedCol) {
      targetField = matchedCol.field;
      queryPart = valStr;
    }
  }

  const matcher = createSmartMatcher(queryPart, 'auto');

  return (data) => {
    if (targetField) {
      const val = data[targetField];
      return val !== null && val !== undefined && matcher(String(val));
    } else {
      return Object.values(data).some(val => {
        if (val === null || val === undefined) return false;
        return matcher(String(val));
      });
    }
  };
}

// 検索クエリを再帰下降パーサーで述語関数 (row) => boolean に変換
function parseSearchQueryToPredicate(query, activeCols) {
  const tokens = tokenizeSearchQuery(query);
  if (tokens.length === 0) return () => true;

  let pos = 0;

  function peek() {
    return tokens[pos] || null;
  }

  function consume(expectedType) {
    const tok = tokens[pos];
    if (!tok || (expectedType && tok.type !== expectedType)) {
      return null;
    }
    pos++;
    return tok;
  }

  function parseExpression() {
    return parseOr();
  }

  function parseOr() {
    let left = parseAnd();
    while (peek() && peek().type === 'OR') {
      consume('OR');
      const right = parseAnd();
      const prevLeft = left;
      left = (data) => prevLeft(data) || right(data);
    }
    return left;
  }

  function parseAnd() {
    let left = parsePrimary();
    while (peek() && peek().type === 'AND') {
      consume('AND');
      const right = parsePrimary();
      const prevLeft = left;
      left = (data) => prevLeft(data) && right(data);
    }
    return left;
  }

  function parsePrimary() {
    const tok = peek();
    if (!tok) return () => true;

    if (tok.type === 'LPAREN') {
      consume('LPAREN');
      const expr = parseExpression();
      if (peek() && peek().type === 'RPAREN') {
        consume('RPAREN');
      }
      return expr;
    }

    if (tok.type === 'LITERAL') {
      consume('LITERAL');
      return createSearchTokenMatcher(tok.value, activeCols);
    }

    pos++;
    return () => true;
  }

  return parseExpression();
}

// 検索ハンドラ (クイック検索: 複数列スペース区切りAND、OR、()グルーピング、列指定、ワイルドカード、数値条件対応)
function handleSearch() {
  if (!currentTable) return;
  const rawQuery = searchInput.value.trim();
  if (!rawQuery) {
    if (activeColumnFilters && activeColumnFilters.size > 0) {
      reapplyTableFilters();
    } else {
      currentTable.clearFilter();
    }
    return;
  }

  const tab = tabs.find(t => t.id === activeTabId);
  const activeCols = tab ? tab.columns.filter(c => c.field) : [];
  const queryPredicate = parseSearchQueryToPredicate(rawQuery, activeCols);

  currentTable.setFilter((data) => {
    // 1. 列フィルター（オートフィルター）の条件をまずチェック
    if (activeColumnFilters && activeColumnFilters.size > 0) {
      for (const [field, allowedSet] of activeColumnFilters.entries()) {
        const v = data[field];
        const norm = normalizeFilterValue(v);
        if (!allowedSet.has(norm)) return false;
      }
    }

    // 2. 検索バーの条件チェック (括弧・AND・OR を総合評価)
    return queryPredicate(data);
  });
}

// 検索・絞り込みガイド ポップアップ制御
function toggleSearchHelpPopup() {
  if (!searchHelpPopup) return;
  const isShown = searchHelpPopup.style.display === 'block';
  if (isShown) {
    hideSearchHelpPopup();
  } else {
    showSearchHelpPopup();
  }
}

function showSearchHelpPopup() {
  if (!searchHelpPopup) return;
  searchHelpPopup.style.display = 'block';
}

function hideSearchHelpPopup() {
  if (!searchHelpPopup) return;
  searchHelpPopup.style.display = 'none';
}

// エンコーディング変更ハンドラ（文字コード手動再解釈含む）
function handleEncodingChange(e) {
  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab) return;

  const newEncoding = e.target.value;
  if (tab.rawBytes && tab.rawBytes.length > 0 && tab.encoding !== newEncoding) {
    const shouldReload = confirm(`文字コードを「${newEncoding}」に変更してファイルを再読み込みしますか？\n（現在の未保存の変更は破棄されます）`);
    if (shouldReload) {
      const bytes = tab.rawBytes;
      const fileName = tab.name;
      closeTab(tab.id);
      
      const rawEncoding = newEncoding.replace('-BOM', '');
      const text = decodeBytes(bytes, rawEncoding);
      const parsed = Papa.parse(text, { skipEmptyLines: false });
      const rawRows = parsed.data || [];
      
      let headers = rawRows.length > 0 ? rawRows[0] : ['列 1'];
      let rowData = [];
      for (let i = 1; i < rawRows.length; i++) {
        const row = rawRows[i];
        if (i === rawRows.length - 1 && row.length === 1 && row[0] === '') continue;
        const rowObj = { _id: i };
        for (let c = 0; c < headers.length; c++) {
          rowObj[`col_${c}`] = (row[c] !== undefined) ? row[c] : '';
        }
        rowData.push(rowObj);
      }

      const newTabId = 'tab-' + Date.now();
      const newTab = {
        id: newTabId,
        name: fileName,
        rawBytes: bytes,
        encoding: newEncoding,
        delimiter: tab.delimiter,
        newline: tab.newline,
        headers: headers,
        columns: buildTabulatorColumns(headers),
        data: rowData,
        isModified: false
      };
      tabs.push(newTab);
      renderTabs();
      activateTab(newTabId);
      return;
    }
  }
  tab.encoding = newEncoding;
  updateStatusBar(tab);
}

function handleDelimiterChange(e) {
  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab) return;
  tab.delimiter = e.target.value;
  tab.isModified = true;
  renderTabs();
  updateStatusBar(tab);
}

function handleNewlineChange(e) {
  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab) return;
  tab.newline = e.target.value;
  tab.isModified = true;
  renderTabs();
  updateStatusBar(tab);
}

// ----------------------------------------------------
// 機能1 & 2: 行コンテキストメニュー & クリップボード・範囲操作
// ----------------------------------------------------

function getRowContextMenu() {
  return [
    {
      label: "⬆️ 上に行を挿入",
      action: (e, row) => addRowRelative(row, 'above')
    },
    {
      label: "⬇️ 下に行を挿入",
      action: (e, row) => addRowRelative(row, 'below')
    },
    {
      label: "📄 この行を複製",
      action: (e, row) => duplicateRow(row)
    },
    { separator: true },
    {
      label: "✂️ 行を切り取り",
      action: (e, row) => {
        rowClipboardBuffer = Object.assign({}, row.getData());
        deleteRow(row, true);
        showToast('行を切り取りました', 'info');
      }
    },
    {
      label: "📋 行をコピー",
      action: (e, row) => {
        rowClipboardBuffer = Object.assign({}, row.getData());
        showToast('行をコピーしました', 'info');
      }
    },
    {
      label: "📋 下に貼り付け",
      action: (e, row) => {
        if (!rowClipboardBuffer) {
          showToast('コピーされた行データがありません', 'error');
          return;
        }
        const newRow = Object.assign({}, rowClipboardBuffer, { _id: Date.now() });
        currentTable.addRow(newRow, false, row).then(() => {
          const tab = tabs.find(t => t.id === activeTabId);
          if (tab) {
            tab.isModified = true;
            syncCurrentTabData();
            renderTabs();
            updateStatusBar(tab);
          }
          showToast('行を貼り付けました', 'success');
        });
      }
    },
    { separator: true },
    {
      label: "🗑️ この行（または選択中の行）を削除",
      action: (e, row) => deleteRow(row)
    }
  ];
}

// // 選択セルまたは範囲のTSVコピー (Excel互換)
function copySelectedCellsToClipboard() {
  if (!currentTable || !activeTabId) return;

  const rows = getAllRows();
  const cols = getActiveColumns();
  if (rows.length === 0 || cols.length === 0) return;

  let rangeToCopy = null;

  if (selectionRange) {
    rangeToCopy = {
      minRow: Math.max(0, Math.min(selectionRange.minRow, rows.length - 1)),
      maxRow: Math.max(0, Math.min(selectionRange.maxRow, rows.length - 1)),
      minCol: Math.max(0, Math.min(selectionRange.minCol, cols.length - 1)),
      maxCol: Math.max(0, Math.min(selectionRange.maxCol, cols.length - 1))
    };
  } else if (activeFocusCell) {
    rangeToCopy = {
      minRow: activeFocusCell.rowIndex,
      maxRow: activeFocusCell.rowIndex,
      minCol: activeFocusCell.colIndex,
      maxCol: activeFocusCell.colIndex
    };
  } else {
    const activeEl = document.activeElement;
    const indices = getCellIndices(activeEl);
    if (indices) {
      rangeToCopy = {
        minRow: indices.rowIndex,
        maxRow: indices.rowIndex,
        minCol: indices.colIndex,
        maxCol: indices.colIndex
      };
    }
  }

  if (!rangeToCopy) return;

  const lines = [];
  for (let r = rangeToCopy.minRow; r <= rangeToCopy.maxRow; r++) {
    const row = rows[r];
    if (!row) continue;
    const rowVals = [];
    for (let c = rangeToCopy.minCol; c <= rangeToCopy.maxCol; c++) {
      const cell = row.getCell(cols[c].field);
      const val = cell ? cell.getValue() : '';
      rowVals.push(val !== null && val !== undefined ? String(val) : '');
    }
    lines.push(rowVals.join('\t'));
  }

  const tsvText = lines.join('\r\n');
  internalClipboardText = tsvText;

  // コピー範囲を記録して緑点線枠（Excel風マーキー）を表示
  copiedRange = { ...rangeToCopy };
  updateRangeHighlight();

  const rowCount = rangeToCopy.maxRow - rangeToCopy.minRow + 1;
  const colCount = rangeToCopy.maxCol - rangeToCopy.minCol + 1;

  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(tsvText).then(() => {
      showToast(`コピーしました (${rowCount}行 × ${colCount}列)`, 'info');
    }).catch(err => {
      console.warn("Clipboard write failed, using internal buffer:", err);
      showToast(`コピーしました (${rowCount}行 × ${colCount}列)`, 'info');
    });
  } else {
    showToast(`コピーしました (${rowCount}行 × ${colCount}列)`, 'info');
  }
}

// クリップボードからの貼り付け (TSV / CSV展開、Excel互換)
async function handleGlobalPaste(e) {
  if (!currentTable || !activeTabId) return;

  const tag = document.activeElement ? document.activeElement.tagName.toLowerCase() : '';
  if (tag === 'input' || tag === 'textarea') {
    if (document.activeElement.closest('.modal') || document.activeElement.closest('.filter-popup')) {
      return;
    }
  }

  let text = '';
  if (e.clipboardData) {
    text = e.clipboardData.getData('text/plain');
  }
  if (!text && internalClipboardText) {
    text = internalClipboardText;
  }
  if (!text) {
    try {
      text = await navigator.clipboard.readText();
    } catch (err) {
      console.warn("Clipboard read failed:", err);
    }
  }
  if (!text) return;

  e.preventDefault();

  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab) return;
  const activeCols = getActiveColumns();
  if (activeCols.length === 0) return;

  // 貼り付け先の起点セル（左上）を特定
  let startRowIndex = 0;
  let startColIndex = 0;

  if (activeFocusCell) {
    startRowIndex = activeFocusCell.rowIndex;
    startColIndex = activeFocusCell.colIndex;
  } else if (selectionRange) {
    startRowIndex = selectionRange.minRow;
    startColIndex = selectionRange.minCol;
  } else {
    const indices = getCellIndices(document.activeElement);
    if (indices) {
      startRowIndex = indices.rowIndex;
      startColIndex = indices.colIndex;
    }
  }

  // TSVテキストを行・列に分解
  const rawLines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  if (rawLines.length > 0 && rawLines[rawLines.length - 1] === '') {
    rawLines.pop();
  }
  if (rawLines.length === 0) return;

  const pasteMatrix = rawLines.map(line => line.split('\t'));
  const pasteRowCount = pasteMatrix.length;
  const pasteColCount = Math.max(...pasteMatrix.map(row => row.length));

  const allRows = getAllRows();

  // 行数が不足している場合、自動で末尾に行を追加（Excelと同等の動作）
  const neededRows = startRowIndex + pasteRowCount;
  if (neededRows > allRows.length) {
    const rowsToAddCount = neededRows - allRows.length;
    for (let i = 0; i < rowsToAddCount; i++) {
      const newRowObj = { _id: Date.now() + Math.random() };
      activeCols.forEach(c => newRowObj[c.field] = '');
      const addedRow = await currentTable.addRow(newRowObj, false);
      allRows.push(addedRow);
    }
  }

  // 矩形領域に値を展開
  for (let r = 0; r < pasteRowCount; r++) {
    const targetRowIdx = startRowIndex + r;
    const targetRow = allRows[targetRowIdx];
    if (!targetRow) continue;

    const rowData = targetRow.getData();
    const cols = pasteMatrix[r];
    for (let c = 0; c < cols.length; c++) {
      const targetColIdx = startColIndex + c;
      if (targetColIdx < activeCols.length) {
        const field = activeCols[targetColIdx].field;
        rowData[field] = cols[c];
      }
    }
    targetRow.update(rowData);
  }

  // 貼り付け後の選択範囲をペースト領域に設定してハイライト！
  const endRowIndex = Math.min(allRows.length - 1, startRowIndex + pasteRowCount - 1);
  const endColIndex = Math.min(activeCols.length - 1, startColIndex + pasteColCount - 1);

  selectionAnchor = { rowIndex: startRowIndex, colIndex: startColIndex };
  activeFocusCell = { rowIndex: startRowIndex, colIndex: startColIndex };
  selectionRange = {
    minRow: startRowIndex,
    maxRow: endRowIndex,
    minCol: startColIndex,
    maxCol: endColIndex,
    startRow: startRowIndex,
    startCol: startColIndex,
    endRow: endRowIndex,
    endCol: endColIndex
  };

  tab.isModified = true;
  syncCurrentTabData();
  renderTabs();
  updateStatusBar(tab);

  updateRangeHighlight();
  showToast(`貼り付けました (${pasteRowCount}行 × ${pasteColCount}列)`, 'success');
}

// 選択セルまたは範囲の一括クリア (Delete / Backspace)
function clearSelectedCells() {
  if (!currentTable || !activeTabId) return;

  const rows = getAllRows();
  const cols = getActiveColumns();
  if (rows.length === 0 || cols.length === 0) return;

  let clearedCount = 0;

  if (selectionRange) {
    const minR = Math.max(0, Math.min(selectionRange.minRow, rows.length - 1));
    const maxR = Math.max(0, Math.min(selectionRange.maxRow, rows.length - 1));
    const minC = Math.max(0, Math.min(selectionRange.minCol, cols.length - 1));
    const maxC = Math.max(0, Math.min(selectionRange.maxCol, cols.length - 1));

    for (let r = minR; r <= maxR; r++) {
      const row = rows[r];
      if (!row) continue;
      const data = row.getData();
      for (let c = minC; c <= maxC; c++) {
        const field = cols[c].field;
        data[field] = '';
        clearedCount++;
      }
      row.update(data);
    }
  } else if (activeFocusCell) {
    const row = rows[activeFocusCell.rowIndex];
    if (row && cols[activeFocusCell.colIndex]) {
      const field = cols[activeFocusCell.colIndex].field;
      const data = row.getData();
      data[field] = '';
      row.update(data);
      clearedCount = 1;
    }
  }

  if (clearedCount > 0) {
    const tab = tabs.find(t => t.id === activeTabId);
    if (tab) {
      tab.isModified = true;
      syncCurrentTabData();
      renderTabs();
      updateStatusBar(tab);
    }
    updateRangeHighlight();
    showToast(`${clearedCount}個のセルをクリアしました`, 'info');
  }
}

// ----------------------------------------------------
// 機能3: キーボードナビゲーション (Tab / F2)
// ----------------------------------------------------

function appendNewColumnAndFocus(tab, targetRowIndex) {
  if (!currentTable || !tab) return;
  const activeCols = tab.columns.filter(c => c.field);
  const maxCols = activeCols.length;
  const nextLetter = getColumnLetter(maxCols);
  const colTimestamp = Date.now();
  const newField = `col_${colTimestamp}`;
  const newTitle = `列 ${maxCols + 1}`;

  const newColDef = {
    title: newTitle,
    field: newField,
    colIndex: maxCols,
    colLetter: nextLetter,
    editor: "input",
    formatter: excelCellFormatter,
    sorter: excelSmartSorter,
    headerSort: false,
    titleFormatter: compositeHeaderFormatter,
    titleFormatterParams: { colIndex: maxCols, colLetter: nextLetter },
    headerTooltip: "ダブルクリックまたは右クリックで列名（変数名）を変更",
    resizable: true,
    minWidth: 80,
    headerContextMenu: getHeaderContextMenu()
  };

  tab.columns.push(newColDef);
  if (Array.isArray(tab.data)) {
    tab.data.forEach(r => {
      if (r && typeof r === 'object' && !(newField in r)) {
        r[newField] = '';
      }
    });
  }

  const addPromise = currentTable.addColumn(newColDef);
  tab.headers = tab.columns.filter(c => c.field).map(c => c.title);
  tab.isModified = true;
  renderTabs();
  updateStatusBar(tab);

  const finishNav = () => {
    refreshColumnLetters();
    updateSortButtonsUI();
    const rows = getAllRows();
    const targetRow = (targetRowIndex >= 0 && targetRowIndex < rows.length) ? rows[targetRowIndex] : (rows[0] || null);
    if (targetRow) {
      moveToAndEditCell(targetRow, newField);
    }
  };

  if (addPromise && typeof addPromise.then === 'function') {
    addPromise.then(() => {
      setTimeout(finishNav, 40);
    }).catch(() => {
      setTimeout(finishNav, 40);
    });
  } else {
    setTimeout(finishNav, 50);
  }
}

function handleTabNavigation(shiftKey) {
  if (!currentTable || !activeTabId) return;
  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab) return;
  const activeCols = getActiveColumns();
  if (activeCols.length === 0) return;

  const rows = getAllRows();
  if (rows.length === 0) return;

  let currentRowIndex = -1;
  let currentColIndex = -1;

  const editingCellEl = document.querySelector('.tabulator-cell.tabulator-editing');
  if (editingCellEl) {
    const indices = getCellIndices(editingCellEl);
    if (indices) {
      currentRowIndex = indices.rowIndex;
      currentColIndex = indices.colIndex;
    } else {
      const field = editingCellEl.getAttribute('tabulator-field');
      currentColIndex = activeCols.findIndex(c => c.field === field);
      const rowEl = editingCellEl.closest('.tabulator-row');
      if (rowEl) currentRowIndex = rows.findIndex(r => r.getElement() === rowEl);
    }
    const inputEl = editingCellEl.querySelector('input');
    if (inputEl) inputEl.blur();
  } else {
    const activeEl = document.activeElement;
    if (activeEl) {
      const indices = getCellIndices(activeEl);
      if (indices) {
        currentRowIndex = indices.rowIndex;
        currentColIndex = indices.colIndex;
      } else if (activeEl.classList && activeEl.classList.contains('tabulator-cell')) {
        const field = activeEl.getAttribute('tabulator-field');
        currentColIndex = activeCols.findIndex(c => c.field === field);
        const rowEl = activeEl.closest('.tabulator-row');
        if (rowEl) currentRowIndex = rows.findIndex(r => r.getElement() === rowEl);
      }
    }
  }

  if (currentRowIndex === -1 || currentColIndex === -1) {
    if (activeFocusCell && activeFocusCell.rowIndex !== undefined && activeFocusCell.colIndex !== undefined) {
      currentRowIndex = activeFocusCell.rowIndex;
      currentColIndex = activeFocusCell.colIndex;
    } else {
      currentRowIndex = 0;
      currentColIndex = 0;
    }
  }

  if (!shiftKey) {
    if (currentColIndex < activeCols.length - 1) {
      // 途中の列（A〜C列など）➔ 右隣のセルへ移動して編集開始
      const nextRow = rows[currentRowIndex];
      const nextField = activeCols[currentColIndex + 1].field;
      setTimeout(() => {
        moveToAndEditCell(nextRow, nextField);
      }, 40);
    } else {
      // 右端列（D列など）➔ 新しい列（E列）を自動追加し、右隣の新しいセル（E4など）へ移動して編集開始！
      appendNewColumnAndFocus(tab, currentRowIndex);
    }
  } else {
    // Shift + Tab: 左へ戻る
    if (currentColIndex > 0) {
      currentColIndex--;
    } else {
      if (currentRowIndex > 0) {
        currentRowIndex--;
        currentColIndex = activeCols.length - 1;
      }
    }
    const nextRow = rows[currentRowIndex];
    const nextField = activeCols[currentColIndex].field;
    setTimeout(() => {
      moveToAndEditCell(nextRow, nextField);
    }, 40);
  }
}

function startCellEditAtEnd() {
  if (!currentTable || !activeTabId) return;

  let targetCell = null;
  if (activeFocusCell) {
    targetCell = getCellAt(activeFocusCell.rowIndex, activeFocusCell.colIndex);
  }
  if (!targetCell) {
    const activeEl = document.activeElement;
    const indices = getCellIndices(activeEl);
    if (indices) {
      targetCell = getCellAt(indices.rowIndex, indices.colIndex);
    }
  }
  if (!targetCell) {
    const rows = getAllRows();
    const cols = getActiveColumns();
    if (rows.length > 0 && cols.length > 0) {
      targetCell = getCellAt(0, 0);
    }
  }
  if (!targetCell) return;

  if (typeof targetCell.edit === 'function') {
    targetCell.edit(true);
    setTimeout(() => {
      const cellEl = targetCell.getElement();
      if (cellEl) {
        const input = cellEl.querySelector('input');
        if (input) {
          input.focus();
          const len = input.value.length;
          input.setSelectionRange(len, len);
        }
      }
    }, 30);
  }
}

// ----------------------------------------------------
// 機能4: 列幅の自動調整 (オートフィット)
// ----------------------------------------------------

function autoFitColumnWidth(column) {
  if (!column || !textMeasureCtx) return;
  const def = column.getDefinition();
  const field = def.field;
  if (!field) return;

  let maxWidth = textMeasureCtx.measureText(def.title || '').width + 42;

  const rows = currentTable.getRows();
  rows.forEach(r => {
    const val = r.getData()[field];
    if (val !== null && val !== undefined) {
      const w = textMeasureCtx.measureText(String(val)).width + 24;
      if (w > maxWidth) maxWidth = w;
    }
  });

  const finalWidth = Math.max(65, Math.min(650, Math.ceil(maxWidth)));
  column.setWidth(finalWidth);
}

function autoFitAllColumns() {
  if (!currentTable) return;
  const cols = currentTable.getColumns();
  cols.forEach(col => {
    if (col.getField()) autoFitColumnWidth(col);
  });
  showToast('すべての列幅を自動調整しました', 'info');
}

// ----------------------------------------------------
// 機能7: 列のドラッグ並べ替え同期
// ----------------------------------------------------

function handleColumnMoved(column, columns) {
  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab) return;

  const dataCols = columns.filter(c => c.getField());
  tab.headers = dataCols.map(c => c.getDefinition().title);
  tab.columns = columns.map(c => c.getDefinition());
  tab.isModified = true;
  renderTabs();
  updateStatusBar(tab);
  refreshColumnLetters();
  updateSortButtonsUI();
  showToast(`列「${column.getDefinition().title}」を移動しました`, 'info');
}

// ----------------------------------------------------
// 機能6: ステータスバー簡易集計 (Excel風)
// ----------------------------------------------------

function updateCalculationStatusBar() {
  if (selectionRange) {
    updateCalculationStatusBarWithRange(
      selectionRange.minRow,
      selectionRange.maxRow,
      selectionRange.minCol,
      selectionRange.maxCol
    );
  } else {
    hideCalculationStatusBar();
  }
}

// ----------------------------------------------------
// 機能5: 検索と置換 (Ctrl+F / Ctrl+H)
// ----------------------------------------------------

let findSearchIndex = -1;
let findMatchedCells = [];

function openFindReplaceModal(mode = 'find') {
  if (!findReplaceModal) return;
  findReplaceModal.style.display = 'flex';
  if (findKeyword) {
    if (!findKeyword.value && searchInput && searchInput.value.trim()) {
      findKeyword.value = searchInput.value.trim();
    }
    if (mode === 'replace' && findKeyword.value && replaceKeyword) {
      replaceKeyword.focus();
      replaceKeyword.select();
    } else {
      findKeyword.focus();
      findKeyword.select();
    }
  }
  findSearchIndex = -1;
  findMatchedCells = [];
}

function closeFindReplaceModal() {
  if (findReplaceModal) {
    findReplaceModal.style.display = 'none';
  }
}

function collectMatchingCells(keyword, caseSensitive, exactMatch, useRegex) {
  const matches = [];
  if (!currentTable || !activeTabId || !keyword) return matches;

  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab) return matches;
  const activeCols = tab.columns.filter(c => c.field);
  const rows = currentTable.getRows();

  const matcher = createSmartMatcher(
    keyword,
    useRegex ? 'auto' : 'normal',
    { caseSensitive, exactMatch }
  );

  rows.forEach(row => {
    const data = row.getData();
    activeCols.forEach(col => {
      const field = col.field;
      const rawVal = data[field] !== undefined && data[field] !== null ? String(data[field]) : '';
      if (matcher(rawVal)) {
        matches.push({ row, field, cell: row.getCell(field) });
      }
    });
  });

  return matches;
}

function executeFindNext() {
  const keyword = findKeyword ? findKeyword.value : '';
  if (!keyword) {
    showToast('検索キーワードを入力してください', 'error');
    return;
  }

  const caseSensitive = findOptCase ? findOptCase.checked : false;
  const exactMatch = findOptExact ? findOptExact.checked : false;
  const useRegex = findOptRegex ? findOptRegex.checked : false;

  findMatchedCells = collectMatchingCells(keyword, caseSensitive, exactMatch, useRegex);

  if (findMatchedCells.length === 0) {
    showToast('一致するセルは見つかりませんでした', 'info');
    return;
  }

  findSearchIndex = (findSearchIndex + 1) % findMatchedCells.length;
  const match = findMatchedCells[findSearchIndex];

  match.row.scrollTo().then(() => {
    moveToAndEditCell(match.row, match.field);
    showToast(`一致: ${findSearchIndex + 1} / ${findMatchedCells.length} 件`, 'info');
  });
}

function executeReplaceOne() {
  if (findMatchedCells.length === 0 || findSearchIndex === -1) {
    executeFindNext();
    return;
  }

  const match = findMatchedCells[findSearchIndex];
  const replaceText = replaceKeyword ? replaceKeyword.value : '';
  const keyword = findKeyword ? findKeyword.value : '';
  const caseSensitive = findOptCase ? findOptCase.checked : false;
  const exactMatch = findOptExact ? findOptExact.checked : false;
  const useRegex = findOptRegex ? findOptRegex.checked : false;

  const rowData = match.row.getData();
  const currentVal = String(rowData[match.field] || '');

  let newVal = '';
  if (exactMatch) {
    newVal = replaceText;
  } else if (useRegex) {
    try {
      const isWildcard = (keyword.includes('*') || keyword.includes('?')) && !/^\/.+\/[gimsuy]*$/.test(keyword);
      const pat = isWildcard ? wildcardToRegexStr(keyword) : keyword;
      const flags = caseSensitive ? 'g' : 'gi';
      const reg = new RegExp(pat, flags);
      newVal = currentVal.replace(reg, replaceText);
    } catch (e) {
      newVal = currentVal.split(keyword).join(replaceText);
    }
  } else {
    if (caseSensitive) {
      newVal = currentVal.split(keyword).join(replaceText);
    } else {
      const reg = new RegExp(keyword.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
      newVal = currentVal.replace(reg, replaceText);
    }
  }

  rowData[match.field] = newVal;
  match.row.update(rowData);

  const tab = tabs.find(t => t.id === activeTabId);
  if (tab) {
    tab.isModified = true;
    syncCurrentTabData();
    renderTabs();
    updateStatusBar(tab);
  }

  showToast('1件置換しました', 'success');
  executeFindNext();
}

function executeReplaceAll() {
  const keyword = findKeyword ? findKeyword.value : '';
  if (!keyword) {
    showToast('検索キーワードを入力してください', 'error');
    return;
  }

  const replaceText = replaceKeyword ? replaceKeyword.value : '';
  const caseSensitive = findOptCase ? findOptCase.checked : false;
  const exactMatch = findOptExact ? findOptExact.checked : false;
  const useRegex = findOptRegex ? findOptRegex.checked : false;

  const matches = collectMatchingCells(keyword, caseSensitive, exactMatch, useRegex);
  if (matches.length === 0) {
    showToast('置換対象のセルが見つかりませんでした', 'info');
    return;
  }

  matches.forEach(match => {
    const rowData = match.row.getData();
    const currentVal = String(rowData[match.field] || '');
    let newVal = '';
    if (exactMatch) {
      newVal = replaceText;
    } else if (useRegex) {
      try {
        const isWildcard = (keyword.includes('*') || keyword.includes('?')) && !/^\/.+\/[gimsuy]*$/.test(keyword);
        const pat = isWildcard ? wildcardToRegexStr(keyword) : keyword;
        const flags = caseSensitive ? 'g' : 'gi';
        const reg = new RegExp(pat, flags);
        newVal = currentVal.replace(reg, replaceText);
      } catch (e) {
        newVal = currentVal.split(keyword).join(replaceText);
      }
    } else {
      if (caseSensitive) {
        newVal = currentVal.split(keyword).join(replaceText);
      } else {
        const reg = new RegExp(keyword.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
        newVal = currentVal.replace(reg, replaceText);
      }
    }
    rowData[match.field] = newVal;
    match.row.update(rowData);
  });

  const tab = tabs.find(t => t.id === activeTabId);
  if (tab) {
    tab.isModified = true;
    syncCurrentTabData();
    renderTabs();
    updateStatusBar(tab);
  }

  showToast(`${matches.length} 件すべて置換しました`, 'success');
  closeFindReplaceModal();
}

// ----------------------------------------------------
// 機能8: Excel風オートフィルター (値チェックボックス絞り込み & コピー連動)
// ----------------------------------------------------

// 列ごとのフィルター条件管理 (field => Set of allowed string values)
const activeColumnFilters = new Map();
let activeFilterColumnField = null;
let activeFilterColumnDef = null;
let currentFilterSearchMode = 'auto'; // 'auto' | 'wildcard' | 'regex'

function setFilterSearchMode(mode) {
  currentFilterSearchMode = mode;
  document.querySelectorAll('.filter-mode-btn').forEach(btn => {
    if (btn.dataset.mode === mode) {
      btn.classList.add('active');
    } else {
      btn.classList.remove('active');
    }
  });
  handleFilterSearchInput();
  if (filterSearchBox) filterSearchBox.focus();
}

function openColumnFilter(e, column) {
  if (!column || !filterPopup || !currentTable) return;
  activeFilterColumnField = column.getField();
  activeFilterColumnDef = column.getDefinition();
  const colTitle = activeFilterColumnDef.title || activeFilterColumnField;

  // 他の列に適用されているフィルター条件を抽出（カスケード連動用）
  const otherActiveFilters = new Map();
  for (const [field, allowedSet] of activeColumnFilters.entries()) {
    if (field !== activeFilterColumnField) {
      otherActiveFilters.set(field, allowedSet);
    }
  }

  // 他の列のフィルター条件を満たす行のみを候補の集計対象にする
  const tab = tabs.find(t => t.id === activeTabId);
  const allData = (tab && tab.data && tab.data.length > 0) ? tab.data : currentTable.getData();
  const targetRows = otherActiveFilters.size === 0
    ? allData
    : allData.filter(row => {
        for (const [field, allowedSet] of otherActiveFilters.entries()) {
          const v = row[field];
          const norm = normalizeFilterValue(v);
          if (!allowedSet.has(norm)) return false;
        }
        return true;
      });

  // ポップアップタイトルの更新 (適用順序と対象行数を明示)
  const filterKeys = Array.from(activeColumnFilters.keys());
  const existingOrder = filterKeys.indexOf(activeFilterColumnField);
  let orderInfo = '';
  if (existingOrder !== -1) {
    orderInfo = ` [適用順: ${existingOrder + 1}番目]`;
  } else if (filterKeys.length > 0) {
    orderInfo = ` [適用すると ${filterKeys.length + 1}番目]`;
  }

  const titleEl = document.getElementById('filter-popup-title');
  if (titleEl) {
    if (otherActiveFilters.size > 0) {
      titleEl.textContent = `🔍「${colTitle}」(${targetRows.length}行対象)${orderInfo}`;
    } else {
      titleEl.textContent = `🔍「${colTitle}」のフィルター${orderInfo}`;
    }
  }

  // 「この列を解除」ボタンの表示制御
  const btnClearCol = document.getElementById('btn-filter-clear-column');
  if (btnClearCol) {
    btnClearCol.style.display = activeColumnFilters.has(activeFilterColumnField) ? 'inline-block' : 'none';
  }

  const valueCounts = new Map();
  targetRows.forEach(row => {
    const v = row[activeFilterColumnField];
    const norm = normalizeFilterValue(v);
    valueCounts.set(norm, (valueCounts.get(norm) || 0) + 1);
  });

  // 現在この列で保存されている許可値セット
  const currentAllowed = activeColumnFilters.get(activeFilterColumnField);

  filterItemsList.innerHTML = '';
  // ユニークな値（1度だけ候補）を自然順（数値優先・文字コード順）でソート
  const sortedValues = Array.from(valueCounts.keys()).sort((a, b) => {
    if (a === '(空白)') return 1;
    if (b === '(空白)') return -1;
    const numA = Number(a);
    const numB = Number(b);
    if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
    return a.localeCompare(b, 'ja');
  });

  sortedValues.forEach(val => {
    const isChecked = currentAllowed ? currentAllowed.has(val) : true;
    const count = valueCounts.get(val) || 0;
    const itemEl = document.createElement('label');
    itemEl.className = 'filter-item';
    itemEl.title = `${val} (${count}行)`;
    itemEl.innerHTML = `
      <input type="checkbox" value="${encodeURIComponent(val)}" ${isChecked ? 'checked' : ''}>
      <span class="filter-item-text" title="${val}">${val}</span>
    `;
    filterItemsList.appendChild(itemEl);
  });

  if (filterSearchBox) {
    filterSearchBox.value = '';
  }
  filterSearchPreviousQuery = '';
  if (btnFilterSearchClear) {
    btnFilterSearchClear.style.display = 'none';
  }
  if (filterAddSelectionRow) {
    filterAddSelectionRow.style.display = 'none';
  }
  if (filterAddToSelection) {
    filterAddToSelection.checked = false;
  }
  if (filterNumMin) filterNumMin.value = '';
  if (filterNumMax) filterNumMax.value = '';

  const countBadge = document.getElementById('filter-search-count');
  if (countBadge) {
    countBadge.textContent = '';
  }
  setFilterSearchMode('auto');

  // 列に数値データが含まれているか判定して範囲パネルの表示を制御
  const hasNumeric = sortedValues.some(v => v !== '(空白)' && isNumericValue(v));
  if (hasNumeric) {
    toggleNumericRangePanel(true);
  } else {
    toggleNumericRangePanel(false);
  }

  // 表示位置の計算
  const triggerEl = e.currentTarget || e.target;
  const rect = triggerEl ? triggerEl.getBoundingClientRect() : { left: 200, bottom: 80 };
  const popupWidth = 270;
  const popupHeight = hasNumeric ? 385 : 340;

  let left = rect.left;
  if (left + popupWidth > window.innerWidth - 10) {
    left = window.innerWidth - popupWidth - 10;
  }
  if (left < 10) left = 10;

  let top = rect.bottom + 4;
  if (top + popupHeight > window.innerHeight - 10) {
    const topAbove = rect.top - popupHeight - 4;
    if (topAbove >= 10) {
      top = topAbove;
    } else {
      top = Math.max(10, window.innerHeight - popupHeight - 10);
    }
  }

  filterPopup.style.top = `${top}px`;
  filterPopup.style.left = `${left}px`;
  filterPopup.style.display = 'block';

  setTimeout(() => {
    if (filterSearchBox) filterSearchBox.focus();
  }, 50);
}

function hideFilterPopup() {
  if (filterPopup) filterPopup.style.display = 'none';
}

function toggleNumericRangePanel(forceState = null) {
  if (!filterNumRangePanel) return;
  const shouldOpen = forceState !== null ? forceState : (filterNumRangePanel.style.display === 'none');
  if (shouldOpen) {
    filterNumRangePanel.style.display = 'block';
    if (btnToggleNumRange) btnToggleNumRange.classList.add('active');
  } else {
    filterNumRangePanel.style.display = 'none';
    if (btnToggleNumRange) btnToggleNumRange.classList.remove('active');
  }
}

function clearNumericRangeInputs() {
  if (filterNumMin) filterNumMin.value = '';
  if (filterNumMax) filterNumMax.value = '';
  if (filterSearchBox) filterSearchBox.value = '';
  filterPopupSelectAll();
  handleFilterSearchInput();
}

/**
 * 最小値・最大値入力欄の値に基づいてリスト項目を選択
 */
function applyNumericRangeSelection() {
  if (!filterItemsList) return;
  const minStr = filterNumMin ? filterNumMin.value.trim() : '';
  const maxStr = filterNumMax ? filterNumMax.value.trim() : '';

  if (!minStr && !maxStr) {
    showToast('最小値または最大値を入力してください', 'info');
    return;
  }

  const minVal = minStr !== '' ? parseFloat(minStr) : -Infinity;
  const maxVal = maxStr !== '' ? parseFloat(maxStr) : Infinity;

  // 検索ボックスにも反映 (双方向同期)
  if (filterSearchBox) {
    if (minStr && maxStr) {
      filterSearchBox.value = `${minStr}..${maxStr}`;
    } else if (minStr) {
      filterSearchBox.value = `>=${minStr}`;
    } else {
      filterSearchBox.value = `<=${maxStr}`;
    }
  }

  const items = filterItemsList.querySelectorAll('.filter-item');
  let matchedCount = 0;

  items.forEach(it => {
    const textEl = it.querySelector('.filter-item-text');
    const text = textEl ? textEl.textContent : '';
    const cb = it.querySelector('input[type="checkbox"]');
    
    // カンマ区切り対応の数値判定
    const clean = text.replace(/,/g, '').trim();
    const num = parseFloat(clean);
    const isMatch = !isNaN(num) && num >= minVal && num <= maxVal;

    it.style.display = isMatch ? 'flex' : 'none';
    if (cb) cb.checked = isMatch;
    if (isMatch) matchedCount++;
  });

  const countBadge = document.getElementById('filter-search-count');
  if (countBadge) {
    countBadge.textContent = `${matchedCount}/${items.length}件`;
  }

  showToast(`範囲内の ${matchedCount} 件を選択しました（「適用」で確定）`, 'info');
}

function clearFilterSearch() {
  if (filterSearchBox) {
    filterSearchBox.value = '';
    filterSearchBox.focus();
  }
  handleFilterSearchInput();
}

function handleFilterSearchInput() {
  if (!filterSearchBox || !filterItemsList) return;
  const rawQuery = filterSearchBox.value;
  const trimmed = rawQuery.trim();
  const matcher = createSmartMatcher(rawQuery, currentFilterSearchMode);
  const items = filterItemsList.querySelectorAll('.filter-item');
  let matchedCount = 0;

  const isAddingSelection = filterAddToSelection && filterAddToSelection.checked;

  // 検索クリアボタンの表示制御
  if (btnFilterSearchClear) {
    btnFilterSearchClear.style.display = trimmed ? 'flex' : 'none';
  }

  // 「現在の選択に追加」チェックボックスの表示制御
  if (filterAddSelectionRow) {
    filterAddSelectionRow.style.display = trimmed ? 'flex' : 'none';
  }

  // 検索条件が解除された（空になった）場合
  if (!trimmed) {
    items.forEach(it => {
      it.style.display = 'flex';
      const cb = it.querySelector('input[type="checkbox"]');
      // 検索解除時: その段階で対象となるすべてのデータが選択されている状態に戻す
      if (cb) cb.checked = true;
      matchedCount++;
    });

    filterSearchPreviousQuery = '';

    const countBadge = document.getElementById('filter-search-count');
    if (countBadge) {
      countBadge.textContent = '';
    }
    if (filterNumMin) filterNumMin.value = '';
    if (filterNumMax) filterNumMax.value = '';
    return;
  }

  items.forEach(it => {
    const textEl = it.querySelector('.filter-item-text');
    const text = textEl ? textEl.textContent : '';
    const cb = it.querySelector('input[type="checkbox"]');
    const isMatch = matcher(text);

    it.style.display = isMatch ? 'flex' : 'none';
    if (isMatch) matchedCount++;

    if (isAddingSelection) {
      // 現在の選択に追加がONの場合: マッチした項目を新たに追加チェック（非表示の項目のチェックはそのまま保持）
      if (isMatch && cb) cb.checked = true;
    } else {
      // 通常検索: 検索条件に合致した表示項目にチェックを入れ、非合致項目はチェックを外す
      if (cb) cb.checked = isMatch;
    }
  });

  filterSearchPreviousQuery = trimmed;

  const countBadge = document.getElementById('filter-search-count');
  if (countBadge) {
    countBadge.textContent = `${matchedCount}/${items.length}件`;
  }

  // 検索窓が数値条件の場合はGUIミニフォーム側にも同期
  const numCond = parseNumericCondition(rawQuery);
  if (numCond) {
    if (numCond.type === 'range') {
      if (filterNumMin) filterNumMin.value = numCond.min;
      if (filterNumMax) filterNumMax.value = numCond.max;
    } else if (numCond.type === 'cmp') {
      if (numCond.op === '<=' || numCond.op === '<') {
        if (filterNumMin) filterNumMin.value = '';
        if (filterNumMax) filterNumMax.value = numCond.value;
      } else if (numCond.op === '>=' || numCond.op === '>') {
        if (filterNumMin) filterNumMin.value = numCond.value;
        if (filterNumMax) filterNumMax.value = '';
      }
    }
  }
}

function filterPopupSelectAll() {
  if (!filterItemsList) return;
  // 表示中の項目（検索絞り込み中ならその結果のみ）を選択
  filterItemsList.querySelectorAll('.filter-item').forEach(it => {
    if (it.style.display !== 'none') {
      const cb = it.querySelector('input[type="checkbox"]');
      if (cb) cb.checked = true;
    }
  });
}

function filterPopupClearAll() {
  if (!filterItemsList) return;
  // 表示中の項目を解除
  filterItemsList.querySelectorAll('.filter-item').forEach(it => {
    if (it.style.display !== 'none') {
      const cb = it.querySelector('input[type="checkbox"]');
      if (cb) cb.checked = false;
    }
  });
}

function applyColumnFilter() {
  if (!currentTable || !activeFilterColumnField) return;

  try {
    const allCheckboxes = filterItemsList.querySelectorAll('input[type="checkbox"]');
    const checkedBoxes = filterItemsList.querySelectorAll('input[type="checkbox"]:checked');

    // 全てチェックされている場合は、フィルター条件を設定しない（解除と同じ）
    if (checkedBoxes.length === allCheckboxes.length) {
      activeColumnFilters.delete(activeFilterColumnField);
    } else {
      const allowedValues = new Set(Array.from(checkedBoxes).map(cb => decodeURIComponent(cb.value)));
      activeColumnFilters.set(activeFilterColumnField, allowedValues);
    }

    reapplyTableFilters();
  } catch (err) {
    console.error('Error applying column filter:', err);
  } finally {
    hideFilterPopup();
  }

  const colTitle = activeFilterColumnDef ? activeFilterColumnDef.title : activeFilterColumnField;
  const visibleCount = currentTable.getRows("active").length;
  showToast(`列「${colTitle}」にフィルターを適用しました (${visibleCount}行表示中)`, 'info');
}

// 現在の列のフィルターを解除
function clearCurrentColumnFilter() {
  if (!activeFilterColumnField) return;
  activeColumnFilters.delete(activeFilterColumnField);
  reapplyTableFilters();
  hideFilterPopup();
  showToast('この列のフィルターを解除しました', 'info');
}

// 全列のフィルターを一括解除
function clearColumnFilter() {
  activeColumnFilters.clear();
  reapplyTableFilters();
  hideFilterPopup();
  showToast('すべてのフィルターを解除しました', 'info');
}

// フィルター条件をテーブルに再適用
function reapplyTableFilters() {
  if (!currentTable) return;

  try {
    if (activeColumnFilters.size === 0) {
      currentTable.clearFilter();
    } else {
      currentTable.setFilter((data) => {
        for (const [field, allowedSet] of activeColumnFilters.entries()) {
          const v = data[field];
          const norm = normalizeFilterValue(v);
          if (!allowedSet.has(norm)) return false;
        }
        return true;
      });
    }

    updateFilterButtonsUI();
    updateSortButtonsUI();
    refreshColumnLetters();
    updateRangeHighlight();
    updateCalculationStatusBar();

    // タブの行数表示更新
    const tab = tabs.find(t => t.id === activeTabId);
    if (tab) updateStatusBar(tab);
  } catch (err) {
    console.error('Error in reapplyTableFilters:', err);
  }
}

// フィルターボタンの見た目（アクティブ状態および適用順序）を更新
function updateFilterButtonsUI() {
  if (!currentTable) return;
  const cols = currentTable.getColumns();
  const filterKeys = Array.from(activeColumnFilters.keys());

  cols.forEach(c => {
    const field = c.getField();
    if (!field) return;
    const el = c.getElement();
    if (!el) return;

    const btnFilter = el.querySelector('.btn-col-filter');
    const coordBar = el.querySelector('.col-coord-bar');
    const colDef = c.getDefinition();
    const colTitle = colDef ? (colDef.title || field) : field;

    const orderIdx = filterKeys.indexOf(field); // 0-indexed
    const hasFilter = (orderIdx !== -1);

    if (btnFilter) {
      btnFilter.classList.toggle('active', hasFilter);
      if (hasFilter) {
        const orderNum = orderIdx + 1;
        btnFilter.innerHTML = `⧩<span class="col-filter-order-badge" title="適用順: ${orderNum}番目">${orderNum}</span>`;
        btnFilter.title = `列「${colTitle}」でフィルター適用中 (適用順: ${orderNum}番目) - クリックして変更または解除`;
      } else {
        btnFilter.innerHTML = '🔍';
        btnFilter.title = `列「${colTitle}」を値で絞り込み (フィルター)`;
      }
    }

    if (coordBar) {
      coordBar.classList.toggle('is-filtered', hasFilter);
    }
  });

  if (filterActiveBadge) {
    if (activeColumnFilters.size > 0) {
      filterActiveBadge.style.display = 'inline-flex';
      const orderLabels = filterKeys.map((f, idx) => {
        const col = cols.find(c => c.getField() === f);
        const title = col ? (col.getDefinition().title || f) : f;
        return `<span class="filter-order-tag" title="適用順: ${idx + 1}番目"><span class="filter-order-tag-num">${idx + 1}</span>${escapeHtml(title)}</span>`;
      }).join('<span class="filter-order-arrow">➔</span>');

      filterActiveBadge.innerHTML = `
        <span class="filter-order-container">
          <span class="filter-order-heading">🔍 フィルター順:</span>
          ${orderLabels}
        </span>
        <span class="filter-order-clear-btn" title="すべてのフィルターを一括解除">✕ 全解除</span>
      `;
    } else {
      filterActiveBadge.style.display = 'none';
    }
  }
}

// ----------------------------------------------------
// 機能9: データ整形 (クレンジング)
// ----------------------------------------------------

function toggleCleanDropdown() {
  if (!cleanDropdownMenu) return;
  const isShown = cleanDropdownMenu.style.display === 'flex';
  cleanDropdownMenu.style.display = isShown ? 'none' : 'flex';
}

function hideCleanDropdown() {
  if (cleanDropdownMenu) {
    cleanDropdownMenu.style.display = 'none';
    const submenuItems = document.querySelectorAll('.dropdown-submenu-item');
    if (submenuItems && submenuItems.forEach) {
      submenuItems.forEach(item => {
        if (item && item.classList) item.classList.remove('submenu-active');
      });
    }
  }
}

function zenkakuToHankaku(str) {
  return str.replace(/[！-～]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
            .replace(/　/g, ' ');
}

function hankakuToZenkaku(str) {
  return str.replace(/[!-~]/g, s => String.fromCharCode(s.charCodeAt(0) + 0xFEE0))
            .replace(/ /g, '　');
}

const KANA_CONVERT_MAP = {
  'ｶﾞ': 'ガ', 'ｷﾞ': 'ギ', 'ｸﾞ': 'グ', 'ｹﾞ': 'ゲ', 'ｺﾞ': 'ゴ',
  'ｻﾞ': 'ザ', 'ｼﾞ': 'ジ', 'ｽﾞ': 'ズ', 'ｾﾞ': 'ゼ', 'ｿﾞ': 'ゾ',
  'ﾀﾞ': 'ダ', 'ﾁﾞ': 'ヂ', 'ﾂﾞ': 'ヅ', 'ﾃﾞ': 'デ', 'ﾄﾞ': 'ド',
  'ﾊﾞ': 'バ', 'ﾋﾞ': 'ビ', 'ﾌﾞ': 'ブ', 'ﾍﾞ': 'ベ', 'ﾎﾞ': 'ボ',
  'ﾊﾟ': 'パ', 'ﾋﾟ': 'ピ', 'ﾌﾟ': 'プ', 'ﾍﾟ': 'ペ', 'ﾎﾟ': 'ポ',
  'ｳﾞ': 'ヴ',
  'ｱ': 'ア', 'ｲ': 'イ', 'ｳ': 'ウ', 'ｴ': 'エ', 'ｵ': 'オ',
  'ｶ': 'カ', 'ｷ': 'キ', 'ｸ': 'ク', 'ｹ': 'ケ', 'ｺ': 'コ',
  'ｻ': 'サ', 'ｼ': 'シ', 'ｽ': 'ス', 'ｾ': 'セ', 'ｿﾞ': 'ソ',
  'ﾀ': 'タ', 'ﾁ': 'チ', 'ﾂ': 'ツ', 'ﾃ': 'テ', 'ﾄ': 'ト',
  'ﾅ': 'ナ', 'ﾆ': 'ニ', 'ﾇ': 'ヌ', 'ﾈ': 'ネ', 'ﾉ': 'ノ',
  'ﾊ': 'ハ', 'ﾋ': 'ヒ', 'ﾌ': 'フ', 'ﾍ': 'ヘ', 'ﾎ': 'ホ',
  'ﾏ': 'マ', 'ﾐ': 'ミ', 'ﾑ': 'ム', 'ﾒ': 'メ', 'ﾓ': 'モ',
  'ﾔ': 'ヤ', 'ﾕ': 'ユ', 'ﾖ': 'ヨ',
  'ﾗ': 'ラ', 'ﾘ': 'リ', 'ﾙ': 'ル', 'ﾚ': 'レ', 'ﾛ': 'ロ',
  'ﾜ': 'ワ', 'ｦ': 'ヲ', 'ﾝ': 'ン',
  'ｧ': 'ァ', 'ｨ': 'ィ', 'ｩ': 'ゥ', 'ｪ': 'ェ', 'ｫ': 'ォ',
  'ｯ': 'ッ', 'ｬ': 'ャ', 'ｭ': 'ュ', 'ｮ': 'ョ',
  'ｰ': 'ー', '｡': '。', '｢': '「', '｣': '」', '､': '、', '･': '・'
};

function hankakuKanaToZenkaku(str) {
  let res = str;
  for (const [k, v] of Object.entries(KANA_CONVERT_MAP)) {
    if (k.length > 1) res = res.split(k).join(v);
  }
  for (const [k, v] of Object.entries(KANA_CONVERT_MAP)) {
    if (k.length === 1) res = res.split(k).join(v);
  }
  return res;
}

function trimSpaces(str) {
  return str.replace(/^[\s　]+|[\s　]+$/g, '');
}

function removeNewlines(str) {
  return str.replace(/\r\n|\r|\n/g, '');
}

// ----------------------------------------------------
// 日付・日時・和暦の正規化・フォーマット関数群
// ----------------------------------------------------

const ERA_OFFSETS = {
  '令和': 2018, '令': 2018, 'R': 2018, 'r': 2018,
  '平成': 1988, '平': 1988, 'H': 1988, 'h': 1988,
  '昭和': 1925, '昭': 1925, 'S': 1925, 's': 1925,
  '大正': 1911, '大': 1911, 'T': 1911, 't': 1911,
  '明治': 1867, '明': 1867, 'M': 1867, 'm': 1867
};

function pad2(n) {
  return String(n).padStart(2, '0');
}

function isLeapYear(y) {
  return (y % 4 === 0 && y % 100 !== 0) || (y % 400 === 0);
}

function isValidDate(y, m, d) {
  if (y < 1868 || y > 2150) return false;
  if (m < 1 || m > 12) return false;
  if (d < 1 || d > 31) return false;
  const daysInMonth = [31, (isLeapYear(y) ? 29 : 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  return d <= daysInMonth[m - 1];
}

// 2桁年 (YY) を4桁西暦 (YYYY) に補正
function normalizeTwoDigitYear(y) {
  if (y >= 100) return y;
  return y >= 70 ? 1900 + y : 2000 + y;
}

// 和暦表記を西暦テキストに置換（例: 令和8年9月20日 → 2026年9月20日, 令和8年9月 → 2026年9月, R8.9.20 → 2026.9.20）
function convertWarekiToSeirekiText(str) {
  if (typeof str !== 'string' || !str) return str;
  let res = str;

  // 1. 漢字元号: 令和8年 / 令和元年 / 令和08年 / 令8年
  res = res.replace(/(令和|平成|昭和|大正|明治|令|平|昭|大|明)(?:元年|(\d{1,2})年)/g, (match, era, numStr) => {
    const base = ERA_OFFSETS[era];
    if (!base) return match;
    const eraYear = numStr ? parseInt(numStr, 10) : 1;
    const seirekiYear = base + eraYear;
    return `${seirekiYear}年`;
  });

  // 2. アルファベット・記号略記: R8.9.20 / H31/4/30 / S60-1-15 / R08.09.20 / R8.9
  res = res.replace(/\b([RrHhSsTtMm])\s*(\d{1,2})([./\-])/g, (match, eraKey, numStr, sep) => {
    const base = ERA_OFFSETS[eraKey];
    if (!base) return match;
    const eraYear = parseInt(numStr, 10);
    const seirekiYear = base + eraYear;
    return `${seirekiYear}${sep}`;
  });

  return res;
}

// 日付・時刻文字列のパース（粒度自動判定: time_only, datetime, date, year_month）
function parseDateTimeString(rawStr) {
  if (typeof rawStr !== 'string' && typeof rawStr !== 'number') {
    return { isValid: false };
  }
  const str = String(rawStr).trim();
  if (!str) return { isValid: false };

  // 1. まず純粋な時刻のみ (time_only) か判定
  // 例: "14:30:15", "14:30", "9:05", "14時30分15秒", "14時30分", "午後2:30", "2:30 PM"
  const pureTimeColon = str.match(/^(?:([+-]?\d+)\s*)?^(?:(午前|午後)\s*)?(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?(?:\.\d+)?(?:\s*(AM|PM|午前|午後))?$/i);
  const pureTimeJp = str.match(/^(?:(午前|午後)\s*)?(\d{1,2})時\s*(\d{1,2})分(?:\s*(\d{1,2})秒)?$/);

  if (pureTimeColon) {
    let h = parseInt(pureTimeColon[3], 10);
    const m = parseInt(pureTimeColon[4], 10);
    const s = pureTimeColon[5] ? parseInt(pureTimeColon[5], 10) : 0;
    const ap = (pureTimeColon[2] || pureTimeColon[6] || '').toUpperCase();
    if (ap === 'PM' || ap === '午後') {
      if (h < 12) h += 12;
    } else if (ap === 'AM' || ap === '午前') {
      if (h === 12) h = 0;
    }
    if (h >= 0 && h <= 23 && m >= 0 && m <= 59 && s >= 0 && s <= 59) {
      return {
        isValid: true,
        granularity: 'time_only',
        hours: h,
        minutes: m,
        seconds: s
      };
    }
  } else if (pureTimeJp) {
    let h = parseInt(pureTimeJp[2], 10);
    const m = parseInt(pureTimeJp[3], 10);
    const s = pureTimeJp[4] ? parseInt(pureTimeJp[4], 10) : 0;
    const ap = pureTimeJp[1] || '';
    if (ap === '午後') {
      if (h < 12) h += 12;
    } else if (ap === '午前') {
      if (h === 12) h = 0;
    }
    if (h >= 0 && h <= 23 && m >= 0 && m <= 59 && s >= 0 && s <= 59) {
      return {
        isValid: true,
        granularity: 'time_only',
        hours: h,
        minutes: m,
        seconds: s
      };
    }
  }

  // 2. 和暦を西暦に変換
  const normalizedStr = convertWarekiToSeirekiText(str);

  let remaining = normalizedStr;
  let hasTime = false;
  let hours = 0;
  let minutes = 0;
  let seconds = 0;

  // 末尾の時刻表記の抽出
  const timeJpMatch = remaining.match(/\s+(\d{1,2})時\s*(\d{1,2})分(?:\s*(\d{1,2})秒)?$/);
  const timeColonMatch = remaining.match(/(?:[T\s]+)(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?(?:\.\d+)?(?:\s*(AM|PM|午前|午後))?(?:Z|[+-]\d{2}:?\d{2})?$/i);

  if (timeJpMatch) {
    hasTime = true;
    hours = parseInt(timeJpMatch[1], 10);
    minutes = parseInt(timeJpMatch[2], 10);
    seconds = timeJpMatch[3] ? parseInt(timeJpMatch[3], 10) : 0;
    remaining = remaining.substring(0, remaining.length - timeJpMatch[0].length).trim();
  } else if (timeColonMatch) {
    hasTime = true;
    hours = parseInt(timeColonMatch[1], 10);
    minutes = parseInt(timeColonMatch[2], 10);
    seconds = timeColonMatch[3] ? parseInt(timeColonMatch[3], 10) : 0;
    const ampm = (timeColonMatch[4] || '').toUpperCase();
    if (ampm === 'PM' || ampm === '午後') {
      if (hours < 12) hours += 12;
    } else if (ampm === 'AM' || ampm === '午前') {
      if (hours === 12) hours = 0;
    }
    remaining = remaining.substring(0, remaining.length - timeColonMatch[0].length).trim();
  }

  if (hasTime) {
    if (hours < 0 || hours > 23 || minutes < 0 || minutes > 59 || seconds < 0 || seconds > 59) {
      return { isValid: false };
    }
  }

  // 3. 日付（年月日）の抽出
  // パターンA: YYYY/MM/DD, YYYY-MM-DD, YYYY.MM.DD または YY/MM/DD
  const sepMatch = remaining.match(/^(\d{2,4})[\/\.\-](\d{1,2})[\/\.\-](\d{1,2})$/);
  // パターンB: YYYY年MM月DD日 または YY年MM月DD日
  const kanjiMatch = remaining.match(/^(\d{2,4})年\s*(\d{1,2})月\s*(\d{1,2})日?$/);
  // パターンC: 8桁連続数値 (例: 20260920) ※時刻がない場合のみ
  const eightDigitMatch = remaining.match(/^(\d{4})(\d{2})(\d{2})$/);

  if (sepMatch) {
    const rawY = parseInt(sepMatch[1], 10);
    const y = normalizeTwoDigitYear(rawY);
    const m = parseInt(sepMatch[2], 10);
    const d = parseInt(sepMatch[3], 10);
    if (isValidDate(y, m, d)) {
      return {
        isValid: true,
        granularity: hasTime ? 'datetime' : 'date',
        year: y,
        month: m,
        day: d,
        hours,
        minutes,
        seconds
      };
    }
  }

  if (kanjiMatch) {
    const rawY = parseInt(kanjiMatch[1], 10);
    const y = normalizeTwoDigitYear(rawY);
    const m = parseInt(kanjiMatch[2], 10);
    const d = parseInt(kanjiMatch[3], 10);
    if (isValidDate(y, m, d)) {
      return {
        isValid: true,
        granularity: hasTime ? 'datetime' : 'date',
        year: y,
        month: m,
        day: d,
        hours,
        minutes,
        seconds
      };
    }
  }

  if (eightDigitMatch && !hasTime) {
    const y = parseInt(eightDigitMatch[1], 10);
    const m = parseInt(eightDigitMatch[2], 10);
    const d = parseInt(eightDigitMatch[3], 10);
    if (isValidDate(y, m, d)) {
      return {
        isValid: true,
        granularity: 'date',
        year: y,
        month: m,
        day: d,
        hours: 0,
        minutes: 0,
        seconds: 0
      };
    }
  }

  // 4. 年月（日なし）の抽出 (月次統計データ)
  // 例: "2026/9", "2026-9", "2026.9", "2026年9月", "26/9", "26-9", "26年9月"
  if (!hasTime) {
    const ymSepMatch = remaining.match(/^(\d{2,4})[\/\.\-](\d{1,2})$/);
    const ymKanjiMatch = remaining.match(/^(\d{2,4})年\s*(\d{1,2})月$/);

    if (ymSepMatch) {
      const rawY = parseInt(ymSepMatch[1], 10);
      const y = normalizeTwoDigitYear(rawY);
      const m = parseInt(ymSepMatch[2], 10);
      if (y >= 1868 && y <= 2150 && m >= 1 && m <= 12) {
        return {
          isValid: true,
          granularity: 'year_month',
          year: y,
          month: m,
          day: 0,
          hours: 0,
          minutes: 0,
          seconds: 0
        };
      }
    }

    if (ymKanjiMatch) {
      const rawY = parseInt(ymKanjiMatch[1], 10);
      const y = normalizeTwoDigitYear(rawY);
      const m = parseInt(ymKanjiMatch[2], 10);
      if (y >= 1868 && y <= 2150 && m >= 1 && m <= 12) {
        return {
          isValid: true,
          granularity: 'year_month',
          year: y,
          month: m,
          day: 0,
          hours: 0,
          minutes: 0,
          seconds: 0
        };
      }
    }
  }

  return { isValid: false };
}

// 日付・日時・年月・時刻のハイフン形式統一 (基本形式: YYYY-MM-DD [hh:mm:ss])
function formatToHyphenDate(str) {
  const p = parseDateTimeString(str);
  if (!p.isValid) return str;

  if (p.granularity === 'time_only') {
    return `${pad2(p.hours)}:${pad2(p.minutes)}:${pad2(p.seconds)}`;
  }
  if (p.granularity === 'year_month') {
    return `${p.year}-${pad2(p.month)}`;
  }
  if (p.granularity === 'date') {
    return `${p.year}-${pad2(p.month)}-${pad2(p.day)}`;
  }
  if (p.granularity === 'datetime') {
    return `${p.year}-${pad2(p.month)}-${pad2(p.day)} ${pad2(p.hours)}:${pad2(p.minutes)}:${pad2(p.seconds)}`;
  }
  return str;
}

// 日付・日時・年月・時刻のスラッシュ形式統一 (YYYY/MM/DD [hh:mm:ss])
function formatToSlashDate(str) {
  const p = parseDateTimeString(str);
  if (!p.isValid) return str;

  if (p.granularity === 'time_only') {
    return `${pad2(p.hours)}:${pad2(p.minutes)}:${pad2(p.seconds)}`;
  }
  if (p.granularity === 'year_month') {
    return `${p.year}/${pad2(p.month)}`;
  }
  if (p.granularity === 'date') {
    return `${p.year}/${pad2(p.month)}/${pad2(p.day)}`;
  }
  if (p.granularity === 'datetime') {
    return `${p.year}/${pad2(p.month)}/${pad2(p.day)} ${pad2(p.hours)}:${pad2(p.minutes)}:${pad2(p.seconds)}`;
  }
  return str;
}

// 欠測値表記を空セルに統一 (NA, null, NaN等 → 空欄)
// ※ 0 や 0.0、秘匿記号 x は絶対に保持
function normalizeMissingToEmpty(str) {
  if (typeof str !== 'string' && typeof str !== 'number') return str;
  const trimmed = String(str).trim();
  if (trimmed === '') return '';

  // 0, 0.0, 0.00 などの数値0は絶対に保護
  if (/^[-+]?0(?:\.0+)?$/.test(trimmed)) return str;

  // 秘匿記号 x, X は絶対に保護
  if (trimmed === 'x' || trimmed === 'X') return str;

  // 欠測値パターン: NA, N/A, n/a, NaN, null, NULL, none, None, #N/A, #VALUE!, ND, nd
  if (/^(NA|N\/A|NaN|null|none|#N\/A|#VALUE!|ND)$/i.test(trimmed)) {
    return '';
  }
  return str;
}

// 数値末尾の単位記号を除去 (例: 12.5% → 12.5, 1,200円 → 1200, 150kg → 150)
function stripUnitsFromNumber(str) {
  if (typeof str !== 'string' && typeof str !== 'number') return str;
  const trimmed = String(str).trim();
  if (trimmed === '') return str;

  // 秘匿記号 x, X は保護
  if (trimmed === 'x' || trimmed === 'X') return str;

  // 数値 + 単位記号
  const unitMatch = trimmed.match(/^([+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?)\s*(%|％|円|千円|百万円|万円|人|kg|g|t|m|cm|mm|km|℃|度|個|件|台|点|本|枚|回|歳|才)$/);
  if (unitMatch) {
    return unitMatch[1].replace(/,/g, '');
  }
  return str;
}

// 数値の3桁カンマを除去 (1,000 → 1000)
function removeCommasFromNumber(str) {
  if (typeof str !== 'string' && typeof str !== 'number') return str;
  const trimmed = String(str).trim();
  if (/^[+-]?(?:\d{1,3}(?:,\d{3})+)(?:\.\d+)?$/.test(trimmed)) {
    return trimmed.replace(/,/g, '');
  }
  return str;
}


// 表全体の行と列を入れ替え（転置: Transpose）
function transposeCurrentTable() {
  if (!currentTable || !activeTabId) return;
  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab) return;

  // 1. 最新のテーブル編集内容を確実に同期
  syncCurrentTabData();

  const activeCols = tab.columns.filter(c => c.field);
  if (activeCols.length === 0) {
    showToast('転置する列がありません', 'info');
    return;
  }

  // 2. 元のヘッダー一覧 (長さ C)
  const originalHeaders = activeCols.map((c, idx) => c.title || `列 ${idx + 1}`);

  // 3. 元のデータ行一覧 (行数 R)
  const originalData = tab.data || [];

  // 4. 二次元グリッド matrix を構築: (R + 1) 行 × C 列
  //    行0: 元ヘッダー
  //    行1〜R: 各データ行の値
  const matrix = [];
  matrix.push(originalHeaders);

  originalData.forEach(item => {
    const rowVals = activeCols.map(c => {
      const val = item[c.field];
      return (val !== null && val !== undefined) ? String(val) : '';
    });
    matrix.push(rowVals);
  });

  const oldRowCount = matrix.length; // R + 1
  const oldColCount = matrix[0].length; // C

  // 5. 転置行列 transposedMatrix: C 行 × (R + 1) 列
  const transposedMatrix = [];
  for (let c = 0; c < oldColCount; c++) {
    const newRow = [];
    for (let r = 0; r < oldRowCount; r++) {
      newRow.push(matrix[r][c]);
    }
    transposedMatrix.push(newRow);
  }

  // 6. 新ヘッダーと新データ行の構築
  //    新ヘッダー: transposedMatrix[0] (長さ R + 1)
  //    新データ行: transposedMatrix[1..C-1] (各行 長さ R + 1)
  const newHeaders = transposedMatrix[0].map((h, idx) => {
    const trimmed = (h !== null && h !== undefined) ? String(h).trim() : '';
    return trimmed !== '' ? trimmed : `列 ${idx + 1}`;
  });

  const newRowData = [];
  for (let r = 1; r < transposedMatrix.length; r++) {
    const rowObj = { _id: r };
    for (let c = 0; c < newHeaders.length; c++) {
      rowObj[`col_${c}`] = transposedMatrix[r][c];
    }
    newRowData.push(rowObj);
  }

  // 元データがヘッダーのみ（データ行0行）だった場合は空行を1行用意
  if (newRowData.length === 0) {
    const emptyRowObj = { _id: 1 };
    for (let c = 0; c < newHeaders.length; c++) {
      emptyRowObj[`col_${c}`] = '';
    }
    newRowData.push(emptyRowObj);
  }

  // 7. フィルターや選択状態・コピー枠のクリア（構造変化に伴う安全措置）
  if (activeColumnFilters) activeColumnFilters.clear();
  if (typeof updateFilterBadgeUI === 'function') updateFilterBadgeUI();
  activeFocusCell = null;
  selectionAnchor = null;
  selectionRange = null;
  copiedRange = null;

  // 8. タブオブジェクトの更新
  tab.headers = newHeaders;
  tab.columns = buildTabulatorColumns(newHeaders);
  tab.data = newRowData;
  tab.isModified = true;

  // 9. テーブル再構築・再描画
  renderTableForTab(tab);
  renderTabs();
  updateStatusBar(tab);

  showToast(`行と列を入れ替えました (${oldRowCount}行×${oldColCount}列 → ${transposedMatrix.length}行×${newHeaders.length}列)`, 'success');
}

// 特定の列にデータクレンジングを適用
function cleanSpecificColumn(column, action) {
  if (!currentTable || !activeTabId) return;
  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab) return;
  const field = column.getField();
  const title = column.getDefinition().title || field;

  let changedCount = 0;
  const rows = currentTable.getRows();
  rows.forEach(r => {
    const data = r.getData();
    const original = data[field];
    if (original !== null && original !== undefined) {
      const str = String(original);
      const updated = applyCleanTransformation(str, action);
      if (original !== updated) {
        data[field] = updated;
        changedCount++;
      }
    }
    r.update(data);
  });

  if (changedCount > 0) {
    tab.isModified = true;
    syncCurrentTabData();
    renderTabs();
    updateStatusBar(tab);
    showToast(`列「${title}」の ${changedCount} セルをクレンジングしました`, 'success');
  } else {
    showToast(`列「${title}」に変更対象のデータはありませんでした`, 'info');
  }
}

// 重複ハイライト管理
let highlightedColumnField = null;

function clearDuplicateHighlights() {
  if (typeof document === 'undefined') return;
  const cells = document.querySelectorAll('.highlight-duplicate-cell');
  cells.forEach(el => el.classList.remove('highlight-duplicate-cell'));
  highlightedColumnField = null;
}

function highlightDuplicatesInColumn(column) {
  if (!currentTable || !activeTabId) return;
  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab) return;

  const field = typeof column === 'string' ? column : column.getField();
  const colDef = tab.columns.find(c => c.field === field);
  const colTitle = colDef ? colDef.title : field;

  // すでに同じ列がハイライトされている場合はトグル解除
  if (highlightedColumnField === field) {
    clearDuplicateHighlights();
    showToast(`列「${colTitle}」の重複ハイライトを解除しました`, 'info');
    return;
  }

  clearDuplicateHighlights();

  const rows = currentTable.getRows();
  const valueCounts = {};

  // 1. 各値の出現回数をカウント（空文字は除外）
  rows.forEach(r => {
    const val = r.getData()[field];
    if (val !== null && val !== undefined) {
      const str = String(val).trim();
      if (str !== '') {
        valueCounts[str] = (valueCounts[str] || 0) + 1;
      }
    }
  });

  // 重複値（2回以上出現）のセット
  const duplicateValues = new Set(
    Object.keys(valueCounts).filter(val => valueCounts[val] > 1)
  );

  let duplicateCellCount = 0;

  if (duplicateValues.size === 0) {
    showToast(`列「${colTitle}」に重複値はありませんでした`, 'info');
    return;
  }

  // 2. 重複セルにクラスを付与
  rows.forEach(r => {
    const val = r.getData()[field];
    if (val !== null && val !== undefined) {
      const str = String(val).trim();
      if (duplicateValues.has(str)) {
        const cell = r.getCell(field);
        if (cell) {
          const el = cell.getElement();
          if (el) {
            el.classList.add('highlight-duplicate-cell');
            duplicateCellCount++;
          }
        }
      }
    }
  });

  highlightedColumnField = field;
  showToast(`列「${colTitle}」で ${duplicateValues.size} 種類の重複値 (${duplicateCellCount} セル) を強調表示しました（再実行で解除）`, 'warning');
}

function highlightDuplicatesInSelectedColumn() {
  if (!currentTable || !activeTabId) return;
  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab) return;

  let targetField = null;
  if (activeFocusCell && activeFocusCell.colField) {
    targetField = activeFocusCell.colField;
  } else {
    const ranges = typeof currentTable.getRanges === 'function' ? currentTable.getRanges() : [];
    if (ranges && ranges.length > 0) {
      const cells = ranges[0].getCells();
      if (cells && cells.length > 0) {
        targetField = cells[0].getField();
      }
    }
  }

  if (!targetField) {
    const activeCols = tab.columns.filter(c => c.field);
    if (activeCols.length > 0) {
      targetField = activeCols[0].field;
    }
  }

  if (!targetField) {
    showToast('重複チェックを行う列を選択してください', 'info');
    return;
  }

  highlightDuplicatesInColumn(targetField);
}

// 列データ診断サマリーモーダル
function showColumnSummaryModal() {
  if (!currentTable || !activeTabId) return;
  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab) return;

  syncCurrentTabData();

  const tbody = document.getElementById('summary-table-tbody');
  if (!tbody) return;
  tbody.innerHTML = '';

  const activeCols = tab.columns.filter(c => c.field);
  const rows = tab.data || [];
  const totalRows = rows.length;

  activeCols.forEach(col => {
    const field = col.field;
    const title = col.title || field;

    let emptyCount = 0;
    let naCount = 0;
    let confidentialCount = 0;
    let zeroCount = 0;
    let validCount = 0;

    let numCount = 0;
    let numSum = 0;
    let numMin = Infinity;
    let numMax = -Infinity;

    let dateCount = 0;
    const valueMap = {};

    rows.forEach(r => {
      const val = r[field];
      if (val === null || val === undefined || String(val).trim() === '') {
        emptyCount++;
        return;
      }
      const str = String(val).trim();
      valueMap[str] = (valueMap[str] || 0) + 1;

      if (/^(NA|N\/A|NaN|null|none|#N\/A|#VALUE!|ND)$/i.test(str)) {
        naCount++;
      } else if (str === 'x' || str === 'X') {
        confidentialCount++;
        validCount++;
      } else if (/^[-+]?0(?:\.0+)?$/.test(str)) {
        zeroCount++;
        validCount++;
        numCount++;
        if (0 < numMin) numMin = 0;
        if (0 > numMax) numMax = 0;
      } else {
        validCount++;

        const cleanNumStr = str.replace(/,/g, '');
        if (/^[+-]?(?:\d+)(?:\.\d+)?$/.test(cleanNumStr)) {
          const numVal = parseFloat(cleanNumStr);
          numCount++;
          numSum += numVal;
          if (numVal < numMin) numMin = numVal;
          if (numVal > numMax) numMax = numVal;
        }

        const dateCheck = parseDateTimeString(str);
        if (dateCheck.isValid) {
          dateCount++;
        }
      }
    });

    const totalMissing = emptyCount + naCount;
    const missingRate = totalRows > 0 ? ((totalMissing / totalRows) * 100).toFixed(1) : '0.0';

    let inferredType = '文字列';
    let typeClass = 'type-string';
    if (validCount > 0) {
      if (dateCount / validCount >= 0.6) {
        inferredType = '日付/日時';
        typeClass = 'type-datetime';
      } else if (numCount / validCount >= 0.6) {
        inferredType = '数値';
        typeClass = 'type-number';
      }
    }

    let duplicateTypes = 0;
    Object.keys(valueMap).forEach(k => {
      if (valueMap[k] > 1) duplicateTypes++;
    });

    let rangeText = '-';
    if (numCount > 0 && numMin !== Infinity && numMax !== -Infinity) {
      const avg = (numSum / numCount).toFixed(2);
      rangeText = `${numMin} 〜 ${numMax} (平均: ${avg})`;
    }

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${escapeHtml(title)}</strong></td>
      <td><span class="summary-badge ${typeClass}">${inferredType}</span></td>
      <td>${totalRows} 行</td>
      <td>
        ${totalMissing} 件 (${missingRate}%)
        ${totalMissing > 0 ? (missingRate > 20 ? '<span class="summary-badge badge-warn">欠測多</span>' : '') : '<span class="summary-badge badge-ok">完全</span>'}
      </td>
      <td>${confidentialCount > 0 ? `<span class="summary-badge badge-warn">${confidentialCount} 件</span>` : '0 件'}</td>
      <td>${zeroCount} 件</td>
      <td><small style="font-variant-numeric: tabular-nums;">${rangeText}</small></td>
      <td>
        ${duplicateTypes > 0 ? `<span class="summary-badge badge-warn">${duplicateTypes} 種類の重複</span>` : '<span class="summary-badge badge-ok">一意</span>'}
      </td>
    `;
    tbody.appendChild(tr);
  });

  const modal = document.getElementById('modal-column-summary');
  if (modal) {
    modal.style.display = 'flex';
  }
}

function hideColumnSummaryModal() {
  const modal = document.getElementById('modal-column-summary');
  if (modal) {
    modal.style.display = 'none';
  }
}

function applyCleanTransformation(str, action) {
  if (action === 'trim') return trimSpaces(str);
  if (action === 'remove_newlines') return removeNewlines(str);
  if (action === 'format_date_hyphen') return formatToHyphenDate(str);
  if (action === 'format_date_slash') return formatToSlashDate(str);
  if (action === 'wareki_to_seireki') return convertWarekiToSeirekiText(str);
  if (action === 'na_to_empty') return normalizeMissingToEmpty(str);
  if (action === 'strip_units') return stripUnitsFromNumber(str);
  if (action === 'remove_commas') return removeCommasFromNumber(str);
  if (action === 'zen2han') return zenkakuToHankaku(str);
  if (action === 'han2zen') return hankakuToZenkaku(str);
  if (action === 'kana2zen') return hankakuKanaToZenkaku(str);
  if (action === 'upper') return str.toUpperCase();
  if (action === 'lower') return str.toLowerCase();
  return str;
}

function executeDataClean(action) {
  if (!currentTable || !activeTabId) return;

  if (action === 'transpose') {
    transposeCurrentTable();
    return;
  }

  if (action === 'highlight_duplicates') {
    highlightDuplicatesInSelectedColumn();
    return;
  }

  if (action === 'column_summary') {
    showColumnSummaryModal();
    return;
  }

  const tab = tabs.find(t => t.id === activeTabId);
  if (!tab) return;
  const activeCols = tab.columns.filter(c => c.field);
  const ranges = typeof currentTable.getRanges === 'function' ? currentTable.getRanges() : [];

  let cellsToClean = [];

  if (ranges && ranges.length > 0) {
    cellsToClean = ranges[0].getCells();
  }

  let changedCount = 0;

  if (cellsToClean.length > 0) {
    cellsToClean.forEach(cell => {
      const original = cell.getValue();
      if (original !== null && original !== undefined) {
        const str = String(original);
        const updated = applyCleanTransformation(str, action);

        if (original !== updated) {
          cell.setValue(updated);
          changedCount++;
        }
      }
    });
  } else {
    const rows = currentTable.getRows();
    rows.forEach(r => {
      const data = r.getData();
      activeCols.forEach(col => {
        const field = col.field;
        const original = data[field];
        if (original !== null && original !== undefined) {
          const str = String(original);
          const updated = applyCleanTransformation(str, action);

          if (original !== updated) {
            data[field] = updated;
            changedCount++;
          }
        }
      });
      r.update(data);
    });
  }

  if (changedCount > 0) {
    tab.isModified = true;
    syncCurrentTabData();
    renderTabs();
    updateStatusBar(tab);
    let detail = `${changedCount} セルを変換しました`;
    if (action === 'remove_newlines') {
      detail = `${changedCount} セルから改行を削除しました`;
    } else if (action === 'format_date_hyphen' || action === 'format_date_slash') {
      detail = `${changedCount} セルの日付・日時・年月を統一しました`;
    } else if (action === 'wareki_to_seireki') {
      detail = `${changedCount} セルの和暦を西暦に変換しました`;
    } else if (action === 'na_to_empty') {
      detail = `${changedCount} セルの欠測値表記を空セルに統一しました`;
    } else if (action === 'strip_units') {
      detail = `${changedCount} セルから数値の単位記号を除去しました`;
    } else if (action === 'remove_commas') {
      detail = `${changedCount} セルの数値から3桁カンマを外しました`;
    }
    showToast(`データクレンジング完了: ${detail}`, 'success');
  } else {
    showToast('変換対象のデータはありませんでした', 'info');
  }
}

// 起動ハンドラ（DOMの準備完了状態を安全にチェック）
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init);
} else {
  init();
}
