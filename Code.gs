const NEW_GAME_TEST_STATION_NAME = '新遊戲測試站';
const NEW_GAME_TEST_STATION_CODE = '999';
// G008、G009、G010維持正式開放；G011左右拍拍樂保留在站點999試玩。
const NEW_GAME_TEST_IDS = ['G011'];

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const tables = {
    '個案資料': [
      '個案ID', '個案代碼', '姓名', '出生日期', '年齡',
      '教育程度', '主要帶領者', '服務備註', '建檔日期',
      '狀態', '建議遊戲難度', '最後SPMSQ日期', '站別'
    ],

    'SPMSQ評估': [
      '評估ID', '個案ID', '評估日期', '評估者',
      'Q1', 'Q2', 'Q3', 'Q4', 'Q5',
      'Q6', 'Q7', 'Q8', 'Q9', 'Q10',
      '原始錯誤數', '教育校正值（未採用）', '計分錯誤數',
      '服務觀察區間', '系統建議難度',
      '工作人員決定難度', '評估備註', '建立時間'
    ],

    '遊戲設定': [
      '遊戲ID', '遊戲名稱', '訓練面向', '啟用狀態',
      '預設難度', '題目數', '預估分鐘', '帶領提示',
      '建立日期', '更新日期', '模組代碼', '版本',
      '遊戲模式', '可遊玩狀態'
    ],

    '活動紀錄': [
      '活動ID', '個案ID', '活動日期', '遊戲ID',
      '遊戲名稱', '使用難度', '完成度',
      '協助程度', '情緒反應', '服務觀察',
      '帶領人員', '建立時間', '站別', '裝置ID',
      '裝置名稱', '遊戲版本', '遊戲得分', '遊玩秒數',
      '答對數', '答錯數', '提示次數', '是否完成',
      '遊戲結果明細'
    ],

    '個案狀態紀錄': [
      '紀錄ID', '個案ID', '變更時間', '原狀態', '新狀態',
      '變更原因', '操作人員', '站別', '裝置ID', '裝置名稱'
    ]
  };

  // 將空白的預設工作表改為「個案資料」
  let defaultSheet =
    ss.getSheetByName('工作表1') ||
    ss.getSheetByName('Sheet1');

  if (
    !ss.getSheetByName('個案資料') &&
    defaultSheet &&
    defaultSheet.getLastRow() === 0
  ) {
    defaultSheet.setName('個案資料');
  }

  Object.entries(tables).forEach(([sheetName, headers]) => {
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#2e6856')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true);

    sheet.setFrozenRows(1);
    sheet.setRowHeight(1, 42);
    sheet.autoResizeColumns(1, headers.length);
  });

  // 建立下拉選單
  setDropdown('個案資料', 6, [
    '國小以下',
    '國中',
    '高中以上'
  ]);

  setDropdown('個案資料', 10, [
    '服務中',
    '暫停服務',
    '已結案'
  ]);

  setDropdown('個案資料', 11, [
    '入門',
    '標準',
    '進階'
  ]);

  for (let column = 5; column <= 14; column++) {
    setDropdown('SPMSQ評估', column, [
      '正確',
      '錯誤／拒答'
    ]);
  }

  setDropdown('SPMSQ評估', 20, [
    '入門',
    '標準',
    '進階'
  ]);

  setDropdown('遊戲設定', 4, [
    '啟用',
    '停用'
  ]);

  setDropdown('遊戲設定', 5, [
    '入門',
    '標準',
    '進階'
  ]);

  setDropdown('遊戲設定', 14, [
    '可玩',
    '待開發'
  ]);

  setDropdown('活動紀錄', 6, [
    '入門',
    '標準',
    '進階'
  ]);

  setDropdown('活動紀錄', 8, [
    '自行完成',
    '少量提示',
    '較多協助'
  ]);

  setDropdown('活動紀錄', 9, [
    '愉快',
    '平穩',
    '不安',
    '抗拒'
  ]);

  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    '資料表建立完成',
    '已建立個案資料、SPMSQ評估、遊戲設定、活動紀錄及個案狀態紀錄。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function setDropdown(sheetName, column, options) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(sheetName);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(options, true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange(
    2,
    column,
    Math.max(sheet.getMaxRows() - 1, 1),
    1
  ).setDataValidation(rule);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const adminMenu = ui.createMenu('管理者資料管理')
    .addItem('刪除選取的錯植個案', 'adminDeleteSelectedClient')
    .addItem('刪除選取的SPMSQ評估', 'adminDeleteSelectedAssessment')
    .addItem('刪除選取的活動紀錄', 'adminDeleteSelectedActivity')
    .addItem('刪除選取的設備', 'adminDeleteSelectedDevice');

  ui
    .createMenu('失智遊戲平台')
    .addItem('正式上線前檢查', 'runProductionReadinessCheck')
    .addSubMenu(adminMenu)
    .addSeparator()
    .addItem('安裝 SPMSQ 自動計分', 'setupSpmsqScoring')
    .addItem('計算選取的 SPMSQ', 'calculateSelectedSpmsq')
    .addSeparator()
    .addItem('安裝／更新遊戲模組', 'setupPlayableGames')
    .addItem('建立新遊戲測試站（999）', 'setupNewGameTestStation')
    .addItem('升級站別與設備設定', 'setupDeviceTracking')
    .addItem('重新套用下拉選單', 'refreshValidations')
    .addToUi();
}

/**
 * 正式上線前唯讀檢查。
 * 僅盤點資料表、測試資料、站別／設備及遊戲狀態，不會修改或刪除任何資料。
 */
function runProductionReadinessCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = [
    '個案資料', 'SPMSQ評估', '遊戲設定', '活動紀錄',
    '裝置設定', '站別清單', '個案狀態紀錄'
  ];
  const issues = [];
  const passed = [];
  const missingSheets = requiredSheets.filter(name => !ss.getSheetByName(name));

  if (missingSheets.length) {
    issues.push('缺少必要工作表：' + missingSheets.join('、'));
  } else {
    passed.push('必要工作表均已建立');
  }

  const clientSheet = ss.getSheetByName('個案資料');
  const testClientIds = {};
  let clientCount = 0;
  let missingClientStationCount = 0;
  let testClientCount = 0;
  let testStationClientCount = 0;

  if (clientSheet && clientSheet.getLastRow() >= 2) {
    const rows = clientSheet
      .getRange(2, 1, clientSheet.getLastRow() - 1, 13)
      .getDisplayValues();

    rows.forEach(row => {
      if (!row[0]) return;
      if (isNewGameTestStation_(row[12])) {
        testStationClientCount += 1;
        return;
      }
      clientCount += 1;

      const searchText = [row[1], row[2], row[7]].join(' ');
      if (/TEST|測試|虛構/i.test(searchText)) {
        testClientIds[String(row[0])] = true;
        testClientCount += 1;
      }
      if (!String(row[12] || '').trim()) {
        missingClientStationCount += 1;
      }
    });
  }

  if (!clientCount) {
    passed.push('個案資料目前為空白，可於正式使用後建檔');
  } else {
    passed.push('已建立 ' + clientCount + ' 位個案');
  }
  if (testClientCount) {
    issues.push('發現 ' + testClientCount + ' 位疑似測試／虛構個案，正式使用前請確認是否保留');
  } else {
    passed.push('未發現明顯測試個案');
  }
  if (testStationClientCount) {
    passed.push(
      '站點999共有 ' + testStationClientCount + ' 位新遊戲測試個案，已與正式檢查分開'
    );
  }
  if (missingClientStationCount) {
    issues.push('有 ' + missingClientStationCount + ' 位個案尚未設定站別');
  } else if (clientCount) {
    passed.push('個案均已設定站別');
  }

  const assessmentSheet = ss.getSheetByName('SPMSQ評估');
  const activitySheet = ss.getSheetByName('活動紀錄');
  const testAssessmentCount = countRowsForClientIds_(assessmentSheet, 2, testClientIds);
  const testActivityCount = countRowsForClientIds_(activitySheet, 2, testClientIds);

  if (testAssessmentCount) {
    issues.push('測試個案仍有 ' + testAssessmentCount + ' 筆 SPMSQ 評估紀錄');
  }
  if (testActivityCount) {
    issues.push('測試個案仍有 ' + testActivityCount + ' 筆活動紀錄');
  }

  const deviceSheet = ss.getSheetByName('裝置設定');
  let enabledDeviceCount = 0;
  let testDeviceCount = 0;
  let incompleteDeviceCount = 0;
  let testStationDeviceCount = 0;

  if (deviceSheet && deviceSheet.getLastRow() >= 2) {
    const rows = deviceSheet
      .getRange(2, 1, deviceSheet.getLastRow() - 1, 8)
      .getDisplayValues();

    rows.forEach(row => {
      if (!row[0]) return;
      if (isNewGameTestStation_(row[2])) {
        testStationDeviceCount += 1;
        return;
      }
      if (String(row[3] || '').trim() === '啟用') {
        enabledDeviceCount += 1;
        if (!String(row[1] || '').trim() ||
            !String(row[2] || '').trim() ||
            !String(row[6] || '').trim()) {
          incompleteDeviceCount += 1;
        }
      }

      const searchText = [row[1], row[6], row[7]].join(' ');
      if (/TEST|測試/i.test(searchText)) {
        testDeviceCount += 1;
      }
    });
  }

  if (!enabledDeviceCount) {
    issues.push('目前沒有啟用中的正式設備');
  } else {
    passed.push('已有 ' + enabledDeviceCount + ' 台啟用設備');
  }
  if (testDeviceCount) {
    issues.push('發現 ' + testDeviceCount + ' 台疑似測試設備，請確認是否停用');
  }
  if (incompleteDeviceCount) {
    issues.push('有 ' + incompleteDeviceCount + ' 台啟用設備缺少名稱、站別或預設帶領人員');
  }
  if (testStationDeviceCount) {
    passed.push(
      '站點999共有 ' + testStationDeviceCount + ' 台測試設備，已與正式檢查分開'
    );
  }

  const requiredGameIds = [
    'G001', 'G003', 'G004', 'G005', 'G006', 'G007', 'G008'
  ];
  const readyGameIds = {};
  const gameSheet = ss.getSheetByName('遊戲設定');

  if (gameSheet && gameSheet.getLastRow() >= 2) {
    gameSheet
      .getRange(2, 1, gameSheet.getLastRow() - 1, 14)
      .getDisplayValues()
      .forEach(row => {
        if (row[0] && row[3] === '啟用' && row[13] === '可玩') {
          readyGameIds[String(row[0])] = true;
        }
      });
  }

  const unavailableGames = requiredGameIds.filter(id => !readyGameIds[id]);
  if (unavailableGames.length) {
    issues.push('下列遊戲尚未同時設為「啟用／可玩」：' + unavailableGames.join('、'));
  } else {
    passed.push('七款正式遊戲均為啟用且可玩');
  }

  const lines = [];
  lines.push(issues.length
    ? '檢查完成：尚有 ' + issues.length + ' 項需要確認。'
    : '檢查完成：目前未發現阻擋正式使用的項目。');

  if (issues.length) {
    lines.push('', '【待確認】');
    issues.forEach((item, index) => lines.push((index + 1) + '. ' + item));
  }

  if (passed.length) {
    lines.push('', '【已通過】');
    passed.forEach(item => lines.push('✓ ' + item));
  }

  lines.push('', '本檢查不會修改或刪除任何資料。');

  SpreadsheetApp.getUi().alert(
    issues.length ? '正式上線前檢查' : '正式上線前檢查通過',
    lines.join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  return {
    ready: issues.length === 0,
    issues: issues,
    passed: passed,
    checkedAt: new Date()
  };
}

function countRowsForClientIds_(sheet, clientIdColumn, clientIds) {
  if (!sheet || sheet.getLastRow() < 2 || !Object.keys(clientIds).length) {
    return 0;
  }

  return sheet
    .getRange(2, clientIdColumn, sheet.getLastRow() - 1, 1)
    .getDisplayValues()
    .filter(row => clientIds[String(row[0] || '')])
    .length;
}

/**
 * 清除已由管理者確認的3位測試／虛構個案及其連動紀錄。
 * 執行前會再次顯示姓名與筆數，並在同一試算表建立隱藏備份工作表。
 */
function cleanupConfirmedTestClientData() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const preview = getConfirmedTestClientCleanupPreview_(ss);

  // 本次使用者明確確認的是檢查報告中的3位；數量若變動就停止，避免擴大刪除範圍。
  if (preview.clients.length !== 3) {
    ui.alert(
      '未執行清除',
      '目前偵測到 ' + preview.clients.length +
        ' 位疑似測試個案，與已確認的3位不一致。\n' +
        '請重新執行「正式上線前檢查」後再確認。',
      ui.ButtonSet.OK
    );
    return { success: false, reason: 'target_count_changed' };
  }

  const clientLines = preview.clients.map((client, index) =>
    (index + 1) + '. ' + client.name +
      (client.code ? '（' + client.code + '）' : '')
  );
  const response = ui.alert(
    '確認清除測試個案',
    [
      '即將清除以下3位測試／虛構個案：',
      clientLines.join('\n'),
      '',
      '連動資料：',
      '・SPMSQ評估：' + preview.assessmentCount + ' 筆',
      '・活動紀錄：' + preview.activityCount + ' 筆',
      '・服務狀態紀錄：' + preview.statusLogCount + ' 筆',
      '',
      '系統會先建立隱藏備份工作表，再執行刪除。',
      '9台疑似測試設備及其他正式資料不會變更。',
      '',
      '確定要繼續嗎？'
    ].join('\n'),
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return { success: false, reason: 'cancelled' };
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const current = getConfirmedTestClientCleanupPreview_(ss);
    const previewIds = preview.clients.map(client => client.id).sort().join('|');
    const currentIds = current.clients.map(client => client.id).sort().join('|');

    if (current.clients.length !== 3 || currentIds !== previewIds) {
      throw new Error('確認後資料已有變動，為避免誤刪，本次已停止。請重新檢查。');
    }

    const backupNames = createCleanupBackupSheets_(ss, [
      '個案資料', 'SPMSQ評估', '活動紀錄', '個案狀態紀錄'
    ]);
    const clientIds = {};
    current.clients.forEach(client => {
      clientIds[client.id] = true;
    });

    const deletedAssessmentCount = deleteRowsForClientIds_(
      ss.getSheetByName('SPMSQ評估'), 2, clientIds
    );
    const deletedActivityCount = deleteRowsForClientIds_(
      ss.getSheetByName('活動紀錄'), 2, clientIds
    );
    const deletedStatusLogCount = deleteRowsForClientIds_(
      ss.getSheetByName('個案狀態紀錄'), 2, clientIds
    );
    const deletedClientCount = deleteRowsForClientIds_(
      ss.getSheetByName('個案資料'), 1, clientIds
    );

    SpreadsheetApp.flush();

    ui.alert(
      '測試個案清除完成',
      [
        '已清除個案：' + deletedClientCount + ' 位',
        '已清除SPMSQ評估：' + deletedAssessmentCount + ' 筆',
        '已清除活動紀錄：' + deletedActivityCount + ' 筆',
        '已清除服務狀態紀錄：' + deletedStatusLogCount + ' 筆',
        '',
        '已建立 ' + backupNames.length + ' 個隱藏備份工作表。',
        '疑似測試設備未處理。',
        '',
        '請再次執行「正式上線前檢查」。'
      ].join('\n'),
      ui.ButtonSet.OK
    );

    return {
      success: true,
      deletedClientCount: deletedClientCount,
      deletedAssessmentCount: deletedAssessmentCount,
      deletedActivityCount: deletedActivityCount,
      deletedStatusLogCount: deletedStatusLogCount,
      backupSheets: backupNames
    };
  } finally {
    lock.releaseLock();
  }
}

function getConfirmedTestClientCleanupPreview_(ss) {
  const clientSheet = ss.getSheetByName('個案資料');
  const clients = [];

  if (clientSheet && clientSheet.getLastRow() >= 2) {
    clientSheet
      .getRange(2, 1, clientSheet.getLastRow() - 1, 13)
      .getDisplayValues()
      .forEach(row => {
        if (!row[0]) return;
        const searchText = [row[1], row[2], row[7]].join(' ');

        if (/TEST|測試|虛構/i.test(searchText)) {
          clients.push({
            id: String(row[0]),
            code: String(row[1] || ''),
            name: String(row[2] || '')
          });
        }
      });
  }

  const clientIds = {};
  clients.forEach(client => {
    clientIds[client.id] = true;
  });

  return {
    clients: clients,
    assessmentCount: countRowsForClientIds_(
      ss.getSheetByName('SPMSQ評估'), 2, clientIds
    ),
    activityCount: countRowsForClientIds_(
      ss.getSheetByName('活動紀錄'), 2, clientIds
    ),
    statusLogCount: countRowsForClientIds_(
      ss.getSheetByName('個案狀態紀錄'), 2, clientIds
    )
  };
}

function createCleanupBackupSheets_(ss, sheetNames) {
  const timestamp = Utilities.formatDate(
    new Date(),
    ss.getSpreadsheetTimeZone() || 'Asia/Taipei',
    'yyyyMMdd_HHmmss'
  );

  return sheetNames
    .map(sheetName => ss.getSheetByName(sheetName))
    .filter(Boolean)
    .map(sourceSheet => {
      const baseName = ('清理備份_' + timestamp + '_' + sourceSheet.getName())
        .substring(0, 95);
      let backupName = baseName;
      let suffix = 2;

      while (ss.getSheetByName(backupName)) {
        backupName = (baseName.substring(0, 90) + '_' + suffix).substring(0, 99);
        suffix += 1;
      }

      const backupSheet = sourceSheet.copyTo(ss).setName(backupName);
      backupSheet.hideSheet();
      return backupName;
    });
}

function deleteRowsForClientIds_(sheet, clientIdColumn, clientIds) {
  if (!sheet || sheet.getLastRow() < 2 || !Object.keys(clientIds).length) {
    return 0;
  }

  const values = sheet
    .getRange(2, clientIdColumn, sheet.getLastRow() - 1, 1)
    .getDisplayValues();
  const rowsToDelete = [];

  values.forEach((row, index) => {
    if (clientIds[String(row[0] || '')]) {
      rowsToDelete.push(index + 2);
    }
  });

  rowsToDelete
    .sort((a, b) => b - a)
    .forEach(rowNumber => sheet.deleteRow(rowNumber));

  return rowsToDelete.length;
}

/**
 * 管理者刪除目前在「個案資料」選取的錯植個案。
 * 會連動刪除該個案的SPMSQ、活動及服務狀態紀錄。
 */
function adminDeleteSelectedClient() {
  const context = getAdminSelectedRow_('個案資料');
  if (!context) return { success: false };

  const ui = SpreadsheetApp.getUi();
  const values = context.sheet
    .getRange(context.row, 1, 1, 13)
    .getDisplayValues()[0];
  const clientId = String(values[0] || '').trim();

  if (!clientId) {
    ui.alert('此列沒有個案ID，未執行刪除。');
    return { success: false };
  }

  const clientIds = {};
  clientIds[clientId] = true;
  const assessmentCount = countRowsForClientIds_(
    context.ss.getSheetByName('SPMSQ評估'), 2, clientIds
  );
  const activityCount = countRowsForClientIds_(
    context.ss.getSheetByName('活動紀錄'), 2, clientIds
  );
  const statusCount = countRowsForClientIds_(
    context.ss.getSheetByName('個案狀態紀錄'), 2, clientIds
  );
  const reason = getAdminDeletionReason_(ui, '刪除錯植個案');

  if (!reason) return { success: false, reason: 'cancelled' };

  const confirmed = ui.alert(
    '確認刪除錯植個案',
    [
      '個案：' + (values[2] || '未填姓名'),
      '個案代碼：' + (values[1] || '未填代碼'),
      '個案ID：' + clientId,
      '',
      '將連動刪除：',
      '・SPMSQ評估：' + assessmentCount + ' 筆',
      '・活動紀錄：' + activityCount + ' 筆',
      '・服務狀態紀錄：' + statusCount + ' 筆',
      '',
      '刪除前會建立隱藏備份。確定繼續嗎？'
    ].join('\n'),
    ui.ButtonSet.YES_NO
  );

  if (confirmed !== ui.Button.YES) {
    return { success: false, reason: 'cancelled' };
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const currentId = context.sheet
      .getRange(context.row, 1)
      .getDisplayValue()
      .trim();
    if (currentId !== clientId) {
      throw new Error('選取列已變動，為避免誤刪，本次已停止。');
    }

    const backupNames = createCleanupBackupSheets_(context.ss, [
      '個案資料', 'SPMSQ評估', '活動紀錄', '個案狀態紀錄'
    ]);
    const deletedAssessmentCount = deleteRowsForClientIds_(
      context.ss.getSheetByName('SPMSQ評估'), 2, clientIds
    );
    const deletedActivityCount = deleteRowsForClientIds_(
      context.ss.getSheetByName('活動紀錄'), 2, clientIds
    );
    const deletedStatusCount = deleteRowsForClientIds_(
      context.ss.getSheetByName('個案狀態紀錄'), 2, clientIds
    );
    const deletedClientCount = deleteRowsForClientIds_(
      context.sheet, 1, clientIds
    );

    logAdminDeletion_(context.ss, {
      action: '刪除錯植個案及連動紀錄',
      targetSheet: '個案資料',
      targetId: clientId,
      targetName: String(values[2] || ''),
      reason: reason,
      backupNames: backupNames,
      details: [
        '個案' + deletedClientCount,
        'SPMSQ' + deletedAssessmentCount,
        '活動' + deletedActivityCount,
        '狀態' + deletedStatusCount
      ].join('；')
    });
    SpreadsheetApp.flush();
    ui.alert('錯植個案及連動紀錄已刪除。');

    return { success: true, clientId: clientId };
  } finally {
    lock.releaseLock();
  }
}

function adminDeleteSelectedAssessment() {
  const context = getAdminSelectedRow_('SPMSQ評估');
  if (!context) return { success: false };

  const ui = SpreadsheetApp.getUi();
  const values = context.sheet
    .getRange(context.row, 1, 1, 23)
    .getDisplayValues()[0];
  const assessmentId = String(values[0] || '').trim();
  const clientId = String(values[1] || '').trim();

  if (!assessmentId || !clientId) {
    ui.alert('此列缺少評估ID或個案ID，未執行刪除。');
    return { success: false };
  }

  const reason = getAdminDeletionReason_(ui, '刪除錯植SPMSQ評估');
  if (!reason) return { success: false, reason: 'cancelled' };

  const confirmed = ui.alert(
    '確認刪除SPMSQ評估',
    [
      '評估ID：' + assessmentId,
      '個案ID：' + clientId,
      '評估日期：' + (values[2] || '未填日期'),
      '',
      '刪除後會依剩餘的最新評估，自動重算個案目前難度與最後評估日期。',
      '刪除前會建立隱藏備份。確定繼續嗎？'
    ].join('\n'),
    ui.ButtonSet.YES_NO
  );

  if (confirmed !== ui.Button.YES) {
    return { success: false, reason: 'cancelled' };
  }

  return deleteSelectedSingleRow_(context, {
    id: assessmentId,
    idColumn: 1,
    backupSheets: ['SPMSQ評估', '個案資料'],
    action: '刪除錯植SPMSQ評估',
    targetName: clientId,
    reason: reason,
    afterDelete: function() {
      recalculateClientAssessmentState_(clientId);
    },
    successMessage: '錯植SPMSQ評估已刪除，個案難度與最後評估日期已重新計算。'
  });
}

function adminDeleteSelectedActivity() {
  const context = getAdminSelectedRow_('活動紀錄');
  if (!context) return { success: false };

  const ui = SpreadsheetApp.getUi();
  const values = context.sheet
    .getRange(context.row, 1, 1, 23)
    .getDisplayValues()[0];
  const activityId = String(values[0] || '').trim();

  if (!activityId) {
    ui.alert('此列沒有活動ID，未執行刪除。');
    return { success: false };
  }

  const reason = getAdminDeletionReason_(ui, '刪除錯植活動紀錄');
  if (!reason) return { success: false, reason: 'cancelled' };

  const confirmed = ui.alert(
    '確認刪除活動紀錄',
    [
      '活動ID：' + activityId,
      '個案ID：' + (values[1] || '未填'),
      '遊戲：' + (values[4] || '未填'),
      '活動日期：' + (values[2] || '未填'),
      '',
      '刪除後首頁、個案紀錄與統計會自動更新。',
      '刪除前會建立隱藏備份。確定繼續嗎？'
    ].join('\n'),
    ui.ButtonSet.YES_NO
  );

  if (confirmed !== ui.Button.YES) {
    return { success: false, reason: 'cancelled' };
  }

  return deleteSelectedSingleRow_(context, {
    id: activityId,
    idColumn: 1,
    backupSheets: ['活動紀錄'],
    action: '刪除錯植活動紀錄',
    targetName: String(values[4] || ''),
    reason: reason,
    successMessage: '錯植活動紀錄已刪除。'
  });
}

function adminDeleteSelectedDevice() {
  const context = getAdminSelectedRow_('裝置設定');
  if (!context) return { success: false };

  const ui = SpreadsheetApp.getUi();
  const values = context.sheet
    .getRange(context.row, 1, 1, 8)
    .getDisplayValues()[0];
  const deviceId = String(values[0] || '').trim();

  if (!deviceId) {
    ui.alert('此列沒有裝置ID，未執行刪除。');
    return { success: false };
  }

  const historyCount = countRowsForValue_(
    context.ss.getSheetByName('活動紀錄'), 14, deviceId
  );
  const reason = getAdminDeletionReason_(ui, '刪除錯植設備');
  if (!reason) return { success: false, reason: 'cancelled' };

  const confirmed = ui.alert(
    '確認刪除設備',
    [
      '設備名稱：' + (values[1] || '未填'),
      '站別：' + (values[2] || '未填'),
      '裝置ID：' + deviceId,
      '',
      '既有活動中引用此設備的紀錄：' + historyCount + ' 筆。',
      '活動歷史會保留；該設備下次開啟網頁時須重新登錄。',
      '刪除前會建立隱藏備份。確定繼續嗎？'
    ].join('\n'),
    ui.ButtonSet.YES_NO
  );

  if (confirmed !== ui.Button.YES) {
    return { success: false, reason: 'cancelled' };
  }

  return deleteSelectedSingleRow_(context, {
    id: deviceId,
    idColumn: 1,
    backupSheets: ['裝置設定'],
    action: '刪除錯植設備',
    targetName: String(values[1] || ''),
    reason: reason,
    details: '保留既有活動紀錄' + historyCount + '筆',
    successMessage: '錯植設備已刪除；該設備下次開啟時須重新登錄。'
  });
}

function getAdminSelectedRow_(expectedSheetName) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (!sheet || sheet.getName() !== expectedSheetName) {
    ui.alert('請先切換到「' + expectedSheetName + '」工作表，並選取要處理的資料列。');
    return null;
  }

  const range = sheet.getActiveRange();
  const row = range ? range.getRow() : 0;
  if (row < 2) {
    ui.alert('請選取第2列以後的資料，不能選取標題列。');
    return null;
  }

  return { ss: ss, sheet: sheet, row: row };
}

function getAdminDeletionReason_(ui, actionName) {
  const response = ui.prompt(
    actionName,
    '請輸入刪除原因（至少2個字），例如：重複建檔、測試資料或資料輸入錯誤。',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return '';
  const reason = response.getResponseText().trim().substring(0, 300);

  if (reason.length < 2) {
    ui.alert('刪除原因至少需要2個字，本次未執行。');
    return '';
  }

  return reason;
}

function deleteSelectedSingleRow_(context, options) {
  const ui = SpreadsheetApp.getUi();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const currentId = context.sheet
      .getRange(context.row, options.idColumn)
      .getDisplayValue()
      .trim();
    if (currentId !== options.id) {
      throw new Error('選取列已變動，為避免誤刪，本次已停止。');
    }

    const backupNames = createCleanupBackupSheets_(
      context.ss, options.backupSheets
    );
    context.sheet.deleteRow(context.row);

    if (options.afterDelete) options.afterDelete();

    logAdminDeletion_(context.ss, {
      action: options.action,
      targetSheet: context.sheet.getName(),
      targetId: options.id,
      targetName: options.targetName || '',
      reason: options.reason,
      backupNames: backupNames,
      details: options.details || '刪除1筆'
    });
    SpreadsheetApp.flush();
    ui.alert(options.successMessage);
    return { success: true, deletedId: options.id };
  } finally {
    lock.releaseLock();
  }
}

function recalculateClientAssessmentState_(clientId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clientSheet = ss.getSheetByName('個案資料');
  const assessmentSheet = ss.getSheetByName('SPMSQ評估');
  const clientRow = findClientRow(clientId);

  if (!clientRow || !clientSheet) return;

  const rows = !assessmentSheet || assessmentSheet.getLastRow() < 2
    ? []
    : assessmentSheet
        .getRange(2, 1, assessmentSheet.getLastRow() - 1, 23)
        .getValues()
        .filter(row => String(row[1] || '') === String(clientId));

  if (!rows.length) {
    clientSheet.getRange(clientRow, 11).setValue('待評估');
    clientSheet.getRange(clientRow, 12).clearContent();
    return;
  }

  rows.sort((a, b) => getAssessmentRowTime_(b) - getAssessmentRowTime_(a));
  const latest = rows[0];
  const latestDifficulty = ['入門', '標準', '進階'].includes(latest[19])
    ? latest[19]
    : (['入門', '標準', '進階'].includes(latest[18]) ? latest[18] : '待評估');

  clientSheet.getRange(clientRow, 11).setValue(latestDifficulty);
  clientSheet.getRange(clientRow, 12).setValue(latest[2] || '');
}

function getAssessmentRowTime_(row) {
  const value = row[21] || row[2];
  const time = new Date(value).getTime();
  return isNaN(time) ? 0 : time;
}

function countRowsForValue_(sheet, column, value) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  return sheet
    .getRange(2, column, sheet.getLastRow() - 1, 1)
    .getDisplayValues()
    .filter(row => String(row[0] || '') === String(value))
    .length;
}

function logAdminDeletion_(ss, data) {
  const sheet = ensureAdminDeletionLogSheet_(ss);
  let operator = '未取得帳號';

  try {
    operator = Session.getActiveUser().getEmail() || operator;
  } catch (error) {
    // 某些部署方式不會提供電子郵件，仍保留其他刪除軌跡。
  }

  sheet.appendRow([
    new Date(),
    data.action,
    data.targetSheet,
    data.targetId,
    data.targetName,
    data.reason,
    operator,
    (data.backupNames || []).join('、'),
    data.details || ''
  ]);
  sheet.getRange(sheet.getLastRow(), 1).setNumberFormat('yyyy/mm/dd hh:mm:ss');
}

function ensureAdminDeletionLogSheet_(ss) {
  const headers = [
    '操作時間', '操作類型', '來源工作表', '資料ID', '資料名稱',
    '刪除原因', '操作帳號', '備份工作表', '處理明細'
  ];
  let sheet = ss.getSheetByName('管理刪除紀錄');

  if (!sheet) {
    sheet = ss.insertSheet('管理刪除紀錄');
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#8f3f36')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * 正式啟用前清空全部測試營運資料。
 * 保留程式、表頭、格式、下拉選單、正式站別清單及遊戲設定。
 */
function resetAllTestOperationalData() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheets = [
    '個案資料', 'SPMSQ評估', '活動紀錄',
    '個案狀態紀錄', '裝置設定'
  ];
  const counts = getOperationalDataCounts_(ss, targetSheets);
  const firstConfirmation = ui.alert(
    '確認清空全部測試營運資料',
    [
      '以下資料將全部清空：',
      '・個案資料：' + counts['個案資料'] + ' 筆',
      '・SPMSQ評估：' + counts['SPMSQ評估'] + ' 筆',
      '・活動紀錄：' + counts['活動紀錄'] + ' 筆',
      '・個案狀態紀錄：' + counts['個案狀態紀錄'] + ' 筆',
      '・裝置設定：' + counts['裝置設定'] + ' 筆',
      '',
      '系統會先建立5個隱藏備份工作表。',
      '工作表結構、正式站別清單及遊戲設定會保留。',
      '',
      '是否進入最後確認？'
    ].join('\n'),
    ui.ButtonSet.YES_NO
  );

  if (firstConfirmation !== ui.Button.YES) {
    return { success: false, reason: 'cancelled' };
  }

  const typedConfirmation = ui.prompt(
    '最後確認',
    '請輸入「清空全部測試資料」後按確定。',
    ui.ButtonSet.OK_CANCEL
  );

  if (typedConfirmation.getSelectedButton() !== ui.Button.OK ||
      typedConfirmation.getResponseText().trim() !== '清空全部測試資料') {
    ui.alert(
      '未執行清空',
      '確認文字不符或已取消，所有資料均未變更。',
      ui.ButtonSet.OK
    );
    return { success: false, reason: 'confirmation_failed' };
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const currentCounts = getOperationalDataCounts_(ss, targetSheets);

    // 若確認過程中資料有新增或刪除，停止執行，避免把剛輸入的正式資料清掉。
    const dataChanged = targetSheets.some(name => currentCounts[name] !== counts[name]);
    if (dataChanged) {
      throw new Error('確認期間資料筆數已有變動，為避免誤刪，本次已停止。請重新執行。');
    }

    const backupNames = createCleanupBackupSheets_(ss, targetSheets);
    const clearedCounts = {};

    targetSheets.forEach(sheetName => {
      clearedCounts[sheetName] = clearOperationalSheetData_(
        ss.getSheetByName(sheetName)
      );
    });

    SpreadsheetApp.flush();

    ui.alert(
      '測試營運資料已清空',
      [
        '已清空個案資料：' + clearedCounts['個案資料'] + ' 筆',
        '已清空SPMSQ評估：' + clearedCounts['SPMSQ評估'] + ' 筆',
        '已清空活動紀錄：' + clearedCounts['活動紀錄'] + ' 筆',
        '已清空個案狀態紀錄：' + clearedCounts['個案狀態紀錄'] + ' 筆',
        '已清空裝置設定：' + clearedCounts['裝置設定'] + ' 筆',
        '',
        '已建立 ' + backupNames.length + ' 個隱藏備份工作表。',
        '正式站別清單與遊戲設定均已保留。',
        '',
        '下一步請先登錄一台正式設備，再執行正式上線前檢查。'
      ].join('\n'),
      ui.ButtonSet.OK
    );

    return {
      success: true,
      clearedCounts: clearedCounts,
      backupSheets: backupNames
    };
  } finally {
    lock.releaseLock();
  }
}

function getOperationalDataCounts_(ss, sheetNames) {
  const result = {};

  sheetNames.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet || sheet.getLastRow() < 2) {
      result[sheetName] = 0;
      return;
    }

    result[sheetName] = sheet
      .getRange(2, 1, sheet.getLastRow() - 1, 1)
      .getDisplayValues()
      .filter(row => String(row[0] || '').trim())
      .length;
  });

  return result;
}

function clearOperationalSheetData_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) {
    return 0;
  }

  const dataRowCount = sheet.getLastRow() - 1;
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  sheet.getRange(2, 1, dataRowCount, lastColumn).clearContent();
  return dataRowCount;
}

function initializeTestData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const clientSheet = ss.getSheetByName('個案資料');
  const gameSheet = ss.getSheetByName('遊戲設定');

  if (!clientSheet || !gameSheet) {
    throw new Error(
      '找不到「個案資料」或「遊戲設定」工作表。'
    );
  }

  refreshValidations();

  // 建立虛構測試個案
  if (!columnContains(clientSheet, 2, 'TEST-001')) {
    clientSheet.appendRow([
      createId('C'),
      'TEST-001',
      '王小花（測試）',
      new Date(1944, 2, 15),
      82,
      '國小以下',
      '林社工',
      '虛構測試資料，非真實個案',
      new Date(),
      '服務中',
      '待評估',
      ''
    ]);
  }

  // 建立預設遊戲；完整遊戲清單由 setupPlayableGames 統一安裝／更新。
  const games = [
    [
      'G001',
      '生活圖卡配對',
      '短期記憶',
      '啟用',
      '標準',
      6,
      8,
      '先讓長者熟悉圖卡，再鼓勵自行配對。',
      new Date(),
      new Date(),
      'memory',
      '1.1.0',
      '單人',
      '可玩'
    ],
    [
      'G002',
      '日常順序記憶',
      '邏輯排序',
      '停用',
      '入門',
      5,
      10,
      '以熟悉的生活步驟為主，必要時提供口語提示。',
      new Date(),
      new Date(),
      'sequence',
      '1.0.0',
      '單人',
      '待開發'
    ],
    [
      'G003',
      '懷舊拼圖',
      '視覺辨識',
      '啟用',
      '標準',
      9,
      10,
      '依長者反應調整拼圖片數，不強調完成速度。',
      new Date(),
      new Date(),
      'puzzle',
      '1.1.0',
      '單人',
      '可玩'
    ]
  ];

  games.forEach(game => {
    if (!columnContains(gameSheet, 1, game[0])) {
      gameSheet.appendRow(game);
    }
  });

  // 日期格式
  clientSheet.getRange('D2:D')
    .setNumberFormat('yyyy/mm/dd');

  clientSheet.getRange('I2:I')
    .setNumberFormat('yyyy/mm/dd');

  clientSheet.getRange('L2:L')
    .setNumberFormat('yyyy/mm/dd');

  gameSheet.getRange('I2:J')
    .setNumberFormat('yyyy/mm/dd');

  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    '測試資料建立完成',
    '已確認一位虛構個案及三款預設遊戲。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function refreshValidations() {
  setDropdown('個案資料', 6, [
    '國小以下',
    '國中',
    '高中以上'
  ]);

  setDropdown('個案資料', 10, [
    '服務中',
    '暫停服務',
    '已結案'
  ]);

  setDropdown('個案資料', 11, [
    '待評估',
    '入門',
    '標準',
    '進階'
  ]);

  for (let column = 5; column <= 14; column++) {
    setDropdown('SPMSQ評估', column, [
      '正確',
      '錯誤／拒答'
    ]);
  }

  setDropdown('SPMSQ評估', 20, [
    '入門',
    '標準',
    '進階'
  ]);

  setDropdown('遊戲設定', 4, [
    '啟用',
    '停用'
  ]);

  setDropdown('遊戲設定', 5, [
    '入門',
    '標準',
    '進階'
  ]);

  setDropdown('遊戲設定', 14, [
    '可玩',
    '待開發'
  ]);

  setDropdown('活動紀錄', 6, [
    '入門',
    '標準',
    '進階'
  ]);

  setDropdown('活動紀錄', 8, [
    '自行完成',
    '少量提示',
    '較多協助'
  ]);

  setDropdown('活動紀錄', 9, [
    '愉快',
    '平穩',
    '不安',
    '抗拒'
  ]);
}

/**
 * 安裝或更新目前可遊玩的遊戲模組。
 * 只在原欄位右側增欄，不移動既有活動紀錄的 A～O 欄。
 */
function setupPlayableGames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gameSheet = ss.getSheetByName('遊戲設定');
  const activitySheet = ss.getSheetByName('活動紀錄');

  if (!gameSheet || !activitySheet) {
    throw new Error('請先執行「建立資料表」。');
  }

  const gameHeaders = [
    '遊戲ID', '遊戲名稱', '訓練面向', '啟用狀態',
    '預設難度', '題目數', '預估分鐘', '帶領提示',
    '建立日期', '更新日期', '模組代碼', '版本',
    '遊戲模式', '可遊玩狀態'
  ];

  const activityExtraHeaders = [
    '遊戲版本', '遊戲得分', '遊玩秒數', '答對數',
    '答錯數', '提示次數', '是否完成', '遊戲結果明細'
  ];

  gameSheet
    .getRange(1, 1, 1, gameHeaders.length)
    .setValues([gameHeaders]);

  activitySheet
    .getRange(1, 16, 1, activityExtraHeaders.length)
    .setValues([activityExtraHeaders]);

  [gameSheet, activitySheet].forEach(sheet => {
    const width = sheet === gameSheet ? 14 : 23;

    sheet.getRange(1, 1, 1, width)
      .setBackground('#2e6856')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true);

    sheet.setFrozenRows(1);
    sheet.setRowHeight(1, 42);
  });

  const now = new Date();
  const games = [
    [
      'G001', '生活圖卡配對', '短期記憶', '啟用',
      '標準', 6, 8,
      '先讓長者熟悉圖卡，再鼓勵自行配對。',
      now, now, 'memory', '1.1.0', '單人', '可玩'
    ],
    [
      'G003', '懷舊拼圖', '視覺辨識', '啟用',
      '標準', 9, 10,
      '可先查看原圖；不強調完成速度。',
      now, now, 'puzzle', '1.1.0', '單人', '可玩'
    ],
    [
      'G004', '剪影猜猜看', '視覺辨識', '啟用',
      '入門', 5, 8,
      '可使用提示顯示輪廓特徵，再鼓勵作答。',
      now, now, 'silhouette', '1.1.0', '單人', '可玩'
    ],
    [
      'G005', '超市購物算帳', '計算與生活功能', '啟用',
      '標準', 5, 10,
      '可逐項讀出商品與數量，必要時提供紙筆。',
      now, now, 'shopping', '1.1.0', '單人', '可玩'
    ],
    [
      'G006', '骰子認知挑戰', '數量與邏輯', '啟用',
      '標準', 5, 10,
      '先確認看清骰子點數，再依題意比較、加總或排序。',
      now, now, 'dice', '1.1.0', '單人', '可玩'
    ],
    [
      'G007', '雙人對決搶答', '注意力與反應', '啟用',
      '進階', 5, 12,
      '由兩位參與者輪流搶答；帶領者可協助朗讀題目。',
      now, now, 'battle', '1.1.0', '雙人', '可玩'
    ],
    [
      'G008', '數字搜尋挑戰', '視覺搜尋與注意力', '啟用',
      '標準', 12, 8,
      '不強調速度；可先使用朗讀題目，必要時再提供視覺提示。',
      now, now, 'numberSearch', '1.1.0', '單人', '可玩'
    ],
    [
      'G009', '數字對位', '數字辨識與位置對應', '啟用',
      '入門', 48, 10,
      '抽一張數字卡，再點相同數字的格子放入；每格都有數字，不強調速度。入門3×3、5×5、5×5，抽8、20、20張；標準5×5、7×7、7×7，抽20、35、35張；進階7×7、9×9、9×9，每輪固定抽35張。第三輪底盤數字隨機換位，每格仍有數字。',
      now, now, 'numberPlacement', '1.3.0', '單人', '可玩'
    ],
    [
      'G010', '穩穩送茶', '眼手協調與平板穩定控制', '啟用',
      '入門', 6, 5,
      '坐姿小幅度操作，每輪2杯、共3輪6杯；先校正傾斜方向再練習，也可用手指拖曳。路徑包含斜線、弧線與S型。每杯送完手動繼續，可休息或放棄。分數僅為活動完成度，不代表身體平衡或醫療評估。傾斜與觸控模式分開觀察。',
      now, now, 'teaDelivery', '0.1.4', '單人', '可玩'
    ],
    [
      'G011', '左右拍拍樂（試玩）', '左右操作、眼手協調與注意力', '啟用',
      '入門', 24, 5,
      '平板平放桌上，左手拍左鼓、右手拍右鼓；先自由練習。入門每輪8拍、標準10拍、進階12拍，共3輪，不比速度。雙手版須兩鼓同時按住，可先按住一邊再拍另一邊；單手或桌機可用依序版。每組與每輪手動繼續，可休息或放棄，紀錄區分操作模式。',
      now, now, 'drumPads', '0.1.0', '單人', '可玩'
    ]
  ];

  const existingRows = gameSheet.getLastRow() < 2
    ? []
    : gameSheet
        .getRange(2, 1, gameSheet.getLastRow() - 1, 14)
        .getValues();

  const rowById = {};
  existingRows.forEach((row, index) => {
    if (row[0]) {
      rowById[String(row[0])] = index + 2;
    }
  });

  games.forEach(game => {
    const rowNumber = rowById[game[0]];

    if (rowNumber) {
      // 保留管理者已調整的名稱、啟用狀態、難度、題數、時間與提示。
      // 只更新模組代碼、版本、模式與可玩狀態，避免更新時洗掉現場設定。
      const existingGame = gameSheet
        .getRange(rowNumber, 1, 1, 14)
        .getValues()[0];

      // 送茶轉正式：只移除原預設名稱的試玩標示，保留現場自訂名稱。
      if (game[0] === 'G010' && String(existingGame[1] || '') === '穩穩送茶（試玩）') {
        existingGame[1] = '';
      }

      // 同步已確認的送茶六杯版預設值，保留現場自行調整的提示與題數。
      if (game[0] === 'G010' && String(existingGame[11] || '') === '0.1.0') {
        if (Number(existingGame[5]) === 9) existingGame[5] = '';
        if (String(existingGame[7] || '') === '坐姿小幅度操作，每輪3杯、共3輪；先測試傾斜並校正，也可用手指拖曳。每杯送完手動繼續，可休息或放棄。分數僅為活動完成度，不代表身體平衡或醫療評估。傾斜與觸控模式分開觀察。') existingGame[7] = '';
      }

      // G009舊測試版的預設題數與提示需同步更新；保留管理者自行改過的內容。
      if (game[0] === 'G009' && ['1.0.0', '1.1.0', '1.2.0'].includes(String(existingGame[11] || ''))) {
        const oldDefaultCounts = { '1.0.0': 18, '1.1.0': 34, '1.2.0': 40 };
        if (Number(existingGame[5]) === oldDefaultCounts[String(existingGame[11])]) existingGame[5] = '';
        const oldHint = '抽一張數字卡，再點相同數字的格子放入；每格都有數字，不強調速度。入門3×3、4×4、5×5；標準5×5、6×6、7×7；進階7×7、8×8、10×10，各三輪。';
        const previousHint = '抽一張數字卡，再點相同數字的格子放入；每格都有數字，不強調速度。入門3×3、4×4、5×5，抽6、11、17張；標準5×5、6×6、7×7，抽17、24、33張；進階7×7、8×8、10×10，抽33、43、50張。每輪取格數三分之二並進位，上限50張；100格依確認採50張。';
        const shuffledBoardHint = '抽一張數字卡，再點相同數字的格子放入；每格都有數字，不強調速度。入門3×3、5×5、5×5，抽6、17、17張；標準5×5、7×7、7×7，抽17、33、33張；進階7×7、9×9、9×9，抽33、50、50張。第三輪底盤數字隨機換位，每格仍有數字。每輪取格數三分之二並進位，上限50張；9×9依上限採50張。';
        if ([oldHint, previousHint, shuffledBoardHint].includes(String(existingGame[7] || ''))) existingGame[7] = '';
      }

      for (let index = 1; index <= 7; index += 1) {
        if (
          existingGame[index] !== '' &&
          existingGame[index] !== null
        ) {
          game[index] = existingGame[index];
        }
      }

      game[8] = existingGame[8] || now;
      game[9] = now;
      gameSheet.getRange(rowNumber, 1, 1, 14).setValues([game]);
    } else {
      gameSheet.appendRow(game);
    }
  });

  // 舊版測試用順序遊戲沒有遊玩模組，避免出現在前台。
  const sequenceRow = rowById.G002;

  if (sequenceRow) {
    gameSheet.getRange(sequenceRow, 4).setValue('停用');
    gameSheet.getRange(sequenceRow, 11, 1, 4)
      .setValues([['sequence', '1.0.0', '單人', '待開發']]);
  }

  gameSheet.getRange('I2:J').setNumberFormat('yyyy/mm/dd');
  activitySheet.getRange('C2:C').setNumberFormat('yyyy/mm/dd');
  activitySheet.getRange('L2:L').setNumberFormat('yyyy/mm/dd hh:mm:ss');
  activitySheet.getRange('R2:R').setNumberFormat('0');

  ensureNewGameTestStation_();

  refreshValidations();
  gameSheet.autoResizeColumns(1, 14);
  activitySheet.autoResizeColumns(16, 8);
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    '遊戲模組更新完成',
    [
      '已安裝或更新九款正式遊戲與一款試玩模組，活動紀錄沿用 P～W 欄。',
      'G010 穩穩送茶已正式開放；G011 左右拍拍樂僅在站點999顯示。',
      '新遊戲測試站：' + NEW_GAME_TEST_STATION_CODE + ' ' + NEW_GAME_TEST_STATION_NAME +
        '（本次測試：左右拍拍樂）'
    ].join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function createId(prefix) {
  const code = Utilities.getUuid()
    .replace(/-/g, '')
    .substring(0, 10)
    .toUpperCase();

  return prefix + '-' + code;
}

function columnContains(sheet, column, value) {
  if (sheet.getLastRow() < 2) {
    return false;
  }

  const values = sheet
    .getRange(2, column, sheet.getLastRow() - 1, 1)
    .getDisplayValues()
    .flat();

  return values.includes(value);
}

// 使用者直接在試算表輸入資料時，自動產生編號
function onEdit(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();

  if (row === 1) {
    return;
  }

  const sheetName = sheet.getName();

  if (sheetName === '個案資料') {
    const name = sheet.getRange(row, 3).getValue();

    if (name && !sheet.getRange(row, 1).getValue()) {
      sheet.getRange(row, 1).setValue(createId('C'));
      sheet.getRange(row, 9).setValue(new Date());
      sheet.getRange(row, 10).setValue('服務中');
      sheet.getRange(row, 11).setValue('待評估');
    }
  }

  if (sheetName === 'SPMSQ評估') {
    const clientId = sheet.getRange(row, 2).getValue();

    if (clientId && !sheet.getRange(row, 1).getValue()) {
      sheet.getRange(row, 1).setValue(createId('S'));
      sheet.getRange(row, 3).setValue(new Date());
      sheet.getRange(row, 22).setValue(new Date());
    }
  }

  if (sheetName === '遊戲設定') {
    const gameName = sheet.getRange(row, 2).getValue();

    if (gameName && !sheet.getRange(row, 1).getValue()) {
      sheet.getRange(row, 1).setValue(createId('G'));
      sheet.getRange(row, 9).setValue(new Date());
      sheet.getRange(row, 10).setValue(new Date());
    }
  }

  if (sheetName === '活動紀錄') {
    const clientId = sheet.getRange(row, 2).getValue();

    if (clientId && !sheet.getRange(row, 1).getValue()) {
      sheet.getRange(row, 1).setValue(createId('A'));
      sheet.getRange(row, 3).setValue(new Date());
      sheet.getRange(row, 12).setValue(new Date());
    }
  }
}
function setupSpmsqScoring() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('SPMSQ評估');

  // 新增遊戲建議說明欄
  sheet.getRange('V1').copyTo(
    sheet.getRange('W1'),
    SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
    false
  );

  sheet.getRange('W1').setValue('遊戲建議說明');

  // 保留原欄位位置以相容既有資料，但停止教育程度加減分。
  sheet.getRange('P1').setValue('教育校正值（未採用）');
  sheet.getRange('Q1').setValue('計分錯誤數');
  sheet.getRange('P1').setNote(
    '目前未採用教育程度加減分；新評估一律記錄為0。'
  );
  sheet.getRange('Q1').setNote(
    '本欄與原始錯誤數相同，作為服務活動建議的計分依據。'
  );

  // 在 Q1～Q10 欄位加入題目提示
  const questionNotes = [[
    '請個案說出今天的完整日期',
    '請個案說出今天是星期幾',
    '請個案說出目前所在場所',
    '請個案說出電話號碼；無電話時依正式題本改詢問地址',
    '請個案說出自己的年齡',
    '請個案說出自己的出生日期',
    '請個案說出現任總統姓名',
    '請個案說出前任總統姓名',
    '請個案說出母親姓名',
    '請個案從20開始，每次減3'
  ]];

  sheet.getRange(1, 5, 1, 10).setNotes(questionNotes);

  // 避免重複安裝觸發器
  ScriptApp.getProjectTriggers()
    .filter(trigger =>
      trigger.getHandlerFunction() === 'spmsqEditTrigger'
    )
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger('spmsqEditTrigger')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  refreshValidations();
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    '安裝完成',
    'SPMSQ資料填寫完整後，系統將自動計分及推薦遊戲難度。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function spmsqEditTrigger(e) {
  const sheet = e.range.getSheet();

  if (
    sheet.getName() !== 'SPMSQ評估' ||
    e.range.getRow() === 1
  ) {
    return;
  }

  const row = e.range.getRow();
  const column = e.range.getColumn();

  // 編輯Q1～Q10時，嘗試自動計分
  if (column >= 5 && column <= 14) {
    calculateSpmsqRow(row, false);
  }

  // 工作人員調整難度時，同步回個案資料
  if (column === 20) {
    applyDifficultyFromAssessment(row);
  }
}

function calculateSelectedSpmsq() {
  const sheet = SpreadsheetApp.getActiveSheet();

  if (sheet.getName() !== 'SPMSQ評估') {
    SpreadsheetApp.getUi().alert(
      '請先切換到「SPMSQ評估」工作表。'
    );
    return;
  }

  const row = sheet.getActiveRange().getRow();

  if (row === 1) {
    SpreadsheetApp.getUi().alert(
      '請選擇一筆SPMSQ評估資料。'
    );
    return;
  }

  calculateSpmsqRow(row, true);
}

function calculateSpmsqRow(row, showAlert) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('SPMSQ評估');

  const values = sheet
    .getRange(row, 1, 1, 23)
    .getValues()[0];

  const clientId = values[1];
  const answers = values.slice(4, 14);

  if (!clientId) {
    if (showAlert) {
      SpreadsheetApp.getUi().alert('此筆資料沒有個案ID。');
    }
    return;
  }

  const validAnswers = [
    '正確',
    '錯誤／拒答'
  ];

  const isComplete = answers.every(answer =>
    validAnswers.includes(answer)
  );

  if (!isComplete) {
    if (showAlert) {
      SpreadsheetApp.getUi().alert(
        '尚未完成',
        '請確認Q1至Q10都已選擇「正確」或「錯誤／拒答」。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return;
  }

  const clientRow = findClientRow(clientId);

  if (!clientRow) {
    if (showAlert) {
      SpreadsheetApp.getUi().alert(
        '找不到對應的個案資料。'
      );
    }
    return;
  }

  const rawErrors = answers.filter(
    answer => answer === '錯誤／拒答'
  ).length;

  // 尚無可核對的台灣教育程度加減分依據，因此不進行校正。
  // P、Q欄仍保留，以免破壞既有工作表欄位及歷史資料。
  const adjustment = 0;
  const scoringErrors = rawErrors;

  const band = getSpmsqBand(scoringErrors);
  const recommendation =
    getGameRecommendation(scoringErrors);

  const assessmentId =
    values[0] || createId('S');

  const assessmentDate =
    values[2] || new Date();

  const workerDecision =
    values[19] || recommendation.level;

  const existingNote =
    values[20] || '';

  const assessor =
    values[3] || '工作人員';

  // 補齊基本資料
  sheet.getRange(row, 1).setValue(assessmentId);
  sheet.getRange(row, 3).setValue(assessmentDate);
  sheet.getRange(row, 4).setValue(assessor);

  // 寫入O至W欄
  sheet.getRange(row, 15, 1, 9).setValues([[
    rawErrors,
    adjustment,
    scoringErrors,
    band,
    recommendation.level,
    workerDecision,
    existingNote,
    new Date(),
    recommendation.description
  ]]);

  applyDifficultyFromAssessment(row);

  if (showAlert) {
    SpreadsheetApp.getUi().alert(
      'SPMSQ計算完成',
      [
        '錯誤數：' + rawErrors,
        '教育程度加減分：未採用',
        '服務觀察區間：' + band,
        '建議遊戲難度：' + recommendation.level,
        '',
        '遊戲難度為本平台的服務活動建議，不代表醫療診斷或SPMSQ正式分級。'
      ].join('\n'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function getSpmsqBand(errors) {
  if (errors <= 2) {
    return '錯誤較少';
  }

  if (errors <= 4) {
    return '需持續觀察';
  }

  if (errors <= 7) {
    return '需較多支持';
  }

  return '建議就醫';
}

function getGameRecommendation(errors) {
  if (errors <= 2) {
    return {
      level: '標準',
      description:
        '建議6至8個項目、8至12分鐘；先讓長者自行嘗試，再提供簡短提示。'
    };
  }

  if (errors <= 4) {
    return {
      level: '入門',
      description:
        '建議4至6個項目、不強調速度；一次提供一個步驟，必要時給口語提示。'
    };
  }

  if (errors <= 7) {
    return {
      level: '入門',
      description:
        '建議3至4個項目、取消時間限制；先示範一次並提供單一步驟提示。'
    };
  }

  return {
    level: '入門',
    description:
      '建議2至3個項目、短時段及一對一帶領；以互動與情緒安定為優先，並建議就醫。'
  };
}

function findClientRow(clientId) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('個案資料');

  if (sheet.getLastRow() < 2) {
    return null;
  }

  const ids = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 1)
    .getDisplayValues()
    .flat();

  const index = ids.indexOf(String(clientId));

  return index === -1 ? null : index + 2;
}

function applyDifficultyFromAssessment(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const assessmentSheet = ss.getSheetByName('SPMSQ評估');

  const clientId =
    assessmentSheet.getRange(row, 2).getDisplayValue();

  const assessmentDate =
    assessmentSheet.getRange(row, 3).getValue();

  const workerDecision =
    assessmentSheet.getRange(row, 20).getDisplayValue();

  if (
    !clientId ||
    !['入門', '標準', '進階'].includes(workerDecision)
  ) {
    return;
  }

  const clientRow = findClientRow(clientId);

  if (!clientRow) {
    return;
  }

  const clientSheet = ss.getSheetByName('個案資料');

  clientSheet
    .getRange(clientRow, 11)
    .setValue(workerDecision);

  clientSheet
    .getRange(clientRow, 12)
    .setValue(assessmentDate || new Date());
}

function createTestSpmsq() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clientSheet = ss.getSheetByName('個案資料');
  const assessmentSheet = ss.getSheetByName('SPMSQ評估');

  const clientData = clientSheet
    .getRange(
      2,
      1,
      Math.max(clientSheet.getLastRow() - 1, 1),
      12
    )
    .getValues();

  const testClient = clientData.find(
    row => row[1] === 'TEST-001'
  );

  if (!testClient) {
    SpreadsheetApp.getUi().alert(
      '找不到TEST-001測試個案。'
    );
    return;
  }

  const existingNotes =
    assessmentSheet.getLastRow() < 2
      ? []
      : assessmentSheet
          .getRange(
            2,
            21,
            assessmentSheet.getLastRow() - 1,
            1
          )
          .getDisplayValues()
          .flat();

  if (existingNotes.includes('虛構SPMSQ測試資料')) {
    SpreadsheetApp.getUi().alert(
      'SPMSQ測試資料已經存在。'
    );
    return;
  }

  assessmentSheet.appendRow([
    createId('S'),
    testClient[0],
    new Date(),
    '林社工',
    '正確',
    '正確',
    '正確',
    '錯誤／拒答',
    '正確',
    '正確',
    '錯誤／拒答',
    '正確',
    '錯誤／拒答',
    '正確',
    '',
    '',
    '',
    '',
    '',
    '',
    '虛構SPMSQ測試資料',
    new Date(),
    ''
  ]);

  const newRow = assessmentSheet.getLastRow();

  calculateSpmsqRow(newRow, true);
  assessmentSheet.setActiveRange(
    assessmentSheet.getRange(newRow, 1, 1, 23)
  );
}
function include(filename) {
  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();
}

function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('暖憶｜認知活動工作台')
    .addMetaTag(
      'viewport',
      'width=device-width, initial-scale=1'
    );
}

function getClientOptions(deviceId, stationName) {
  const device = getEnabledDevice_(deviceId);
  const activeStationName = getActiveStationName_(stationName, device.stationName);
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('個案資料');

  if (sheet.getLastRow() < 2) {
    return [];
  }

  const values = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 13)
    .getValues();

  return values
    .filter(row =>
      row[0] &&
      row[9] !== '已結案' &&
      String(row[12] || '').trim() === activeStationName
    )
    .map(row => ({
      id: String(row[0]),
      code: String(row[1] || ''),
      name: String(row[2] || ''),
      age: Number(row[4] || 0),
      education: String(row[5] || ''),
      worker: String(row[6] || ''),
      note: String(row[7] || ''),
      status: String(row[9] || ''),
      difficulty: String(row[10] || '待評估'),
      lastSpmsqDate: formatDateForWeb(row[11]),
      stationName: String(row[12] || '')
    }));
}

/**
 * 取得個案紀錄查詢清單。與服務流程清單不同，這裡保留已結案個案，
 * 讓歷史資料仍可查閱及在需要時恢復服務。
 */
function getClientRecordOptions(deviceId, stationName) {
  const device = getEnabledDevice_(deviceId);
  const activeStationName = getActiveStationName_(stationName, device.stationName);
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('個案資料');

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const statusRank = {
    '服務中': 1,
    '暫停服務': 2,
    '已結案': 3
  };

  return sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 13)
    .getValues()
    .filter(row =>
      row[0] &&
      String(row[12] || '').trim() === activeStationName
    )
    .map(row => ({
      id: String(row[0] || ''),
      name: String(row[2] || ''),
      age: Number(row[4] || 0),
      worker: String(row[6] || ''),
      status: String(row[9] || '服務中'),
      difficulty: String(row[10] || '待評估'),
      stationName: String(row[12] || '')
    }))
    .sort((a, b) =>
      (statusRank[a.status] || 9) - (statusRank[b.status] || 9) ||
      a.name.localeCompare(b.name, 'zh-Hant')
    );
}

/**
 * 讀取單一個案的認知活動紀錄。僅允許目前設備所選站別的個案。
 */
function getClientGameRecords(query) {
  query = query || {};

  const device = getEnabledDevice_(query.deviceId);
  const activeStationName = getActiveStationName_(
    query.stationName,
    device.stationName
  );
  const clientId = cleanWebText(query.clientId);
  const period = ['1m', '3m', '6m', 'all'].includes(query.period)
    ? query.period
    : '3m';
  const resultStatus = ['all', 'completed', 'abandoned'].includes(query.resultStatus)
    ? query.resultStatus
    : 'all';
  const gameId = cleanWebText(query.gameId);

  if (!clientId) {
    throw new Error('請選擇要查詢的個案。');
  }

  const clientRow = findClientRow(clientId);

  if (!clientRow) {
    throw new Error('找不到所選個案。');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clientSheet = ss.getSheetByName('個案資料');
  const clientValues = clientSheet
    .getRange(clientRow, 1, 1, 13)
    .getValues()[0];
  const clientStationName = String(clientValues[12] || '').trim();

  if (clientStationName !== activeStationName) {
    throw new Error('此個案不屬於目前選擇的站別。');
  }

  const client = {
    id: clientId,
    name: String(clientValues[2] || ''),
    age: Number(clientValues[4] || 0),
    worker: String(clientValues[6] || ''),
    status: String(clientValues[9] || '服務中'),
    difficulty: String(clientValues[10] || '待評估'),
    stationName: clientStationName
  };
  const activitySheet = ss.getSheetByName('活動紀錄');

  if (!activitySheet || activitySheet.getLastRow() < 2) {
    return buildClientGameRecordResult_(client, [], [], 0);
  }

  const allRows = activitySheet
    .getRange(2, 1, activitySheet.getLastRow() - 1, 23)
    .getValues()
    .filter(row => String(row[1] || '') === clientId);

  const gameMap = {};
  allRows.forEach(row => {
    const id = String(row[3] || '');
    if (id) gameMap[id] = String(row[4] || id);
  });

  let cutoff = null;
  if (period !== 'all') {
    const months = period === '1m' ? 1 : period === '6m' ? 6 : 3;
    cutoff = new Date();
    cutoff.setMonth(cutoff.getMonth() - months);
  }

  const filteredRows = allRows
    .filter(row => {
      const recordDate = row[11] || row[2];
      const date = new Date(recordDate);
      const isCompleted = String(row[21] || '').trim() !== '否';

      if (cutoff && (!date || isNaN(date.getTime()) || date < cutoff)) {
        return false;
      }
      if (gameId && String(row[3] || '') !== gameId) {
        return false;
      }
      if (resultStatus === 'completed' && !isCompleted) {
        return false;
      }
      if (resultStatus === 'abandoned' && isCompleted) {
        return false;
      }
      return true;
    })
    .sort((a, b) =>
      getActivityRowTime_(b) - getActivityRowTime_(a)
    );

  const games = Object.keys(gameMap)
    .map(id => ({ id: id, name: gameMap[id] }))
    .sort((a, b) => a.name.localeCompare(b.name, 'zh-Hant'));

  return buildClientGameRecordResult_(
    client,
    filteredRows.slice(0, 50),
    games,
    allRows.length,
    filteredRows.length,
    filteredRows
  );
}

function buildClientGameRecordResult_(client, rows, games, allGameCount, filteredCount, summaryRows) {
  const rowsForSummary = summaryRows || rows;
  const completedCount = rowsForSummary.filter(row =>
    String(row[21] || '').trim() !== '否'
  ).length;
  const completionTotal = rowsForSummary.reduce((sum, row) =>
    sum + Number(row[6] || 0), 0
  );
  const remainder = allGameCount % 3;

  return {
    client: client,
    games: games || [],
    records: rows.map(row => ({
      activityId: String(row[0] || ''),
      date: formatDateTimeForWeb_(row[11] || row[2]),
      gameId: String(row[3] || ''),
      gameName: String(row[4] || ''),
      level: String(row[5] || ''),
      completion: Number(row[6] || 0),
      help: String(row[7] || '未填寫'),
      mood: String(row[8] || '未填寫'),
      note: String(row[9] || ''),
      worker: String(row[10] || ''),
      score: Number(row[16] || 0),
      durationSeconds: Number(row[17] || 0),
      correctCount: Number(row[18] || 0),
      wrongCount: Number(row[19] || 0),
      hintCount: Number(row[20] || 0),
      completed: String(row[21] || '').trim() !== '否'
    })),
    summary: {
      totalCount: Number(filteredCount == null ? rows.length : filteredCount),
      shownCount: rows.length,
      completedCount: completedCount,
      abandonedCount: rowsForSummary.length - completedCount,
      averageCompletion: rowsForSummary.length
        ? Math.round(completionTotal / rowsForSummary.length)
        : 0,
      allGameCount: allGameCount,
      nextReviewIn: remainder === 0 ? 3 : 3 - remainder
    }
  };
}

function getActivityRowTime_(row) {
  const value = row[11] || row[2];
  const time = new Date(value).getTime();
  return isNaN(time) ? 0 : time;
}

function formatDateTimeForWeb_(value) {
  if (!value) return '';
  const date = new Date(value);
  if (isNaN(date.getTime())) return String(value);
  return Utilities.formatDate(
    date,
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'Asia/Taipei',
    'yyyy/MM/dd HH:mm'
  );
}

/**
 * 結束或恢復平台服務狀態。此操作不刪除資料，且每次變更均保留軌跡。
 */
function updateClientServiceStatus(data) {
  data = data || {};

  const device = getEnabledDevice_(data.deviceId);
  const activeStationName = getActiveStationName_(
    data.stationName,
    device.stationName
  );
  const clientId = cleanWebText(data.clientId);
  const targetStatus = cleanWebText(data.targetStatus);
  const reason = cleanWebText(data.reason).substring(0, 300);
  const operator = cleanWebText(data.operator || device.defaultWorker).substring(0, 30);

  if (!clientId) throw new Error('請選擇個案。');
  if (!['已結案', '服務中'].includes(targetStatus)) {
    throw new Error('服務狀態設定不正確。');
  }
  if (reason.length < 2) {
    throw new Error('請填寫至少2個字的變更原因。');
  }
  if (!operator) {
    throw new Error('設備尚未設定預設帶領人員，無法記錄操作人員。');
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const clientSheet = ss.getSheetByName('個案資料');
    const clientRow = findClientRow(clientId);

    if (!clientRow) throw new Error('找不到所選個案。');

    const row = clientSheet.getRange(clientRow, 1, 1, 13).getValues()[0];
    const clientStationName = String(row[12] || '').trim();
    const previousStatus = String(row[9] || '服務中');

    if (clientStationName !== activeStationName) {
      throw new Error('此個案不屬於目前選擇的站別。');
    }
    if (previousStatus === targetStatus) {
      throw new Error('此個案目前已是「' + targetStatus + '」。');
    }

    const now = new Date();
    clientSheet.getRange(clientRow, 10).setValue(targetStatus);

    const logSheet = ensureClientStatusLogSheet_();
    logSheet.appendRow([
      createId('CS'),
      clientId,
      now,
      previousStatus,
      targetStatus,
      reason,
      operator,
      activeStationName,
      String(device.deviceId || ''),
      String(device.deviceName || '')
    ]);
    logSheet.getRange(logSheet.getLastRow(), 3)
      .setNumberFormat('yyyy/mm/dd hh:mm:ss');

    SpreadsheetApp.flush();

    return {
      success: true,
      clientId: clientId,
      clientName: String(row[2] || ''),
      previousStatus: previousStatus,
      status: targetStatus
    };
  } finally {
    lock.releaseLock();
  }
}

function ensureClientStatusLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headers = [
    '紀錄ID', '個案ID', '變更時間', '原狀態', '新狀態',
    '變更原因', '操作人員', '站別', '裝置ID', '裝置名稱'
  ];
  let sheet = ss.getSheetByName('個案狀態紀錄');

  if (!sheet) {
    sheet = ss.insertSheet('個案狀態紀錄');
  }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#2e6856')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setFrozenRows(1);
  return sheet;
}

function createClientFromWeb(form) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('個案資料');

    const name = cleanWebText(form.name);
    const education = cleanWebText(form.education);
    const worker = cleanWebText(form.worker);
    const note = cleanWebText(form.note);
    const birthDateText = cleanWebText(form.birthDate);
    const device = getEnabledDevice_(form.deviceId);
    const activeStationName = getActiveStationName_(
      form.stationName,
      device.stationName
    );

    if (!name) {
      throw new Error('請輸入個案姓名。');
    }

    if (
      ![
        '國小以下',
        '國中',
        '高中以上'
      ].includes(education)
    ) {
      throw new Error('請選擇教育程度。');
    }

    let birthDate = '';

    if (birthDateText) {
      birthDate = new Date(
        birthDateText + 'T12:00:00'
      );

      if (isNaN(birthDate.getTime())) {
        throw new Error('出生日期格式不正確。');
      }
    }

    const age = birthDate
      ? calculateAge(birthDate)
      : Number(form.age || 0);

    if (!Number.isInteger(age) || age < 55 || age > 120) {
      throw new Error('年齡限55至120歲，請確認年齡或出生日期。');
    }

    const clientId = createId('C');
    const clientCode = createCaseCode();

    sheet.appendRow([
      clientId,
      clientCode,
      name,
      birthDate,
      age,
      education,
      worker || '未指定',
      note,
      new Date(),
      '服務中',
      '待評估',
      '',
      activeStationName
    ]);

    const newRow = sheet.getLastRow();

    sheet
      .getRange(newRow, 4)
      .setNumberFormat('yyyy/mm/dd');

    sheet
      .getRange(newRow, 9)
      .setNumberFormat('yyyy/mm/dd');

    SpreadsheetApp.flush();

    return {
      success: true,
      client: {
        id: clientId,
        code: clientCode,
        name: name,
        age: age,
        education: education,
        worker: worker || '未指定',
        difficulty: '待評估',
        stationName: activeStationName
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function saveSpmsqFromWeb(data) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const clientId = cleanWebText(data.clientId);
    const assessor =
      cleanWebText(data.assessor) || '工作人員';

    if (!clientId) {
      throw new Error('缺少個案ID。');
    }

    const clientRow = findClientRow(clientId);

    if (!clientRow) {
      throw new Error('找不到對應的個案資料。');
    }

    const assessmentClientStatus = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('個案資料')
      .getRange(clientRow, 10)
      .getDisplayValue();

    if (assessmentClientStatus === '已結案') {
      throw new Error('此個案已結束服務，無法新增SPMSQ評估。');
    }

    if (
      !Array.isArray(data.answers) ||
      data.answers.length !== 10
    ) {
      throw new Error('SPMSQ必須完成全部10題。');
    }

    const answers = data.answers.map(answer =>
      answer === true || answer === '正確'
        ? '正確'
        : '錯誤／拒答'
    );

    const questionNotes = Array.isArray(data.questionNotes)
      ? data.questionNotes
      : [];

    const noteParts = questionNotes
      .map((note, index) => {
        const text = cleanWebText(note);

        return text
          ? 'Q' + (index + 1) + '：' + text
          : '';
      })
      .filter(Boolean);

    const generalNote = cleanWebText(data.note);

    if (generalNote) {
      noteParts.push('整體觀察：' + generalNote);
    }

    const assessmentSheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('SPMSQ評估');

    assessmentSheet.appendRow([
      createId('S'),
      clientId,
      new Date(),
      assessor,
      ...answers,
      '',
      '',
      '',
      '',
      '',
      '',
      noteParts.join('\n'),
      new Date(),
      ''
    ]);

    const newRow = assessmentSheet.getLastRow();

    calculateSpmsqRow(newRow, false);

    const result = assessmentSheet
      .getRange(newRow, 15, 1, 9)
      .getDisplayValues()[0];

    SpreadsheetApp.flush();

    return {
      success: true,
      assessmentId:
        assessmentSheet.getRange(newRow, 1)
          .getDisplayValue(),

      rawErrors: Number(result[0]),
      scoringErrors: Number(result[2]),
      band: result[3],
      recommendedLevel: result[4],
      workerDecision: result[5],
      recommendation: result[8]
    };
  } finally {
    lock.releaseLock();
  }
}

function updateClientDifficulty(clientId, level) {
  const validLevels = [
    '入門',
    '標準',
    '進階'
  ];

  if (!validLevels.includes(level)) {
    throw new Error('遊戲難度設定不正確。');
  }

  const clientRow = findClientRow(clientId);

  if (!clientRow) {
    throw new Error('找不到對應的個案資料。');
  }

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('個案資料');

  sheet.getRange(clientRow, 11).setValue(level);

  SpreadsheetApp.flush();

  return {
    success: true,
    clientId: clientId,
    level: level
  };
}

function calculateAge(birthDate) {
  const today = new Date();

  let age =
    today.getFullYear() -
    birthDate.getFullYear();

  const monthDifference =
    today.getMonth() -
    birthDate.getMonth();

  if (
    monthDifference < 0 ||
    (
      monthDifference === 0 &&
      today.getDate() < birthDate.getDate()
    )
  ) {
    age--;
  }

  return age;
}

function createCaseCode() {
  const timeZone = Session.getScriptTimeZone();

  return 'CASE-' + Utilities.formatDate(
    new Date(),
    timeZone,
    'yyyyMMdd-HHmmss'
  );
}

function formatDateForWeb(value) {
  if (!value) {
    return '';
  }

  const date = new Date(value);

  if (isNaN(date.getTime())) {
    return String(value);
  }

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'yyyy/MM/dd'
  );
}

function cleanWebText(value) {
  return String(value || '')
    .trim()
    .substring(0, 1000);
}

function testWebBackend() {
  const clients = getClientOptions(getFirstActiveDeviceId_());

  SpreadsheetApp.getUi().alert(
    '後端測試成功',
    '目前可供選擇的服務個案共有 ' +
      clients.length +
      ' 位。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function getClientHistory(clientId) {
  const cleanClientId = cleanWebText(clientId);

  if (!cleanClientId) {
    throw new Error('缺少個案ID。');
  }

  const clientRow = findClientRow(cleanClientId);

  if (!clientRow) {
    throw new Error('找不到對應的個案資料。');
  }

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('SPMSQ評估');

  if (sheet.getLastRow() < 2) {
    return {
      history: [],
      summary: '尚無SPMSQ評估紀錄。',
      difference: 0,
      difficultyChange: ''
    };
  }

  const rows = sheet
    .getRange(
      2,
      1,
      sheet.getLastRow() - 1,
      23
    )
    .getValues()
    .filter(row =>
      String(row[1]) === cleanClientId
    )
    .sort((a, b) =>
      new Date(b[2]).getTime() -
      new Date(a[2]).getTime()
    );

  const history = rows.map(row => ({
    assessmentId: String(row[0] || ''),
    date: formatDateForWeb(row[2]),
    assessor: String(row[3] || ''),
    rawErrors: Number(row[14] || 0),
    scoringErrors: Number(row[14] || 0),
    band: String(row[17] || ''),
    recommendedLevel: String(row[18] || ''),
    workerDecision: String(row[19] || ''),
    note: String(row[20] || ''),
    recommendation: String(row[22] || '')
  }));

  if (history.length === 0) {
    return {
      history: [],
      summary: '尚無SPMSQ評估紀錄。',
      difference: 0,
      difficultyChange: ''
    };
  }

  if (history.length === 1) {
    return {
      history: history,
      summary:
        '目前只有一次評估，尚不足以形成變化趨勢。',
      difference: 0,
      difficultyChange: ''
    };
  }

  const latest = history[0];
  const previous = history[1];

  const difference =
    latest.rawErrors -
    previous.rawErrors;

  let summary = '';

  if (difference > 0) {
    summary =
      '錯誤數較前次增加 ' +
      difference +
      ' 題。建議先確認當日精神、情緒、聽力及環境因素，並持續觀察；如變化持續，請依機構流程處理。';
  } else if (difference < 0) {
    summary =
      '錯誤數較前次減少 ' +
      Math.abs(difference) +
      ' 題。可維持目前活動方式，仍應搭配現場觀察綜合判讀。';
  } else {
    summary =
      '錯誤數與前次相同，近期結果大致穩定。';
  }

  let difficultyChange = '';

  if (
    latest.workerDecision &&
    previous.workerDecision &&
    latest.workerDecision !== previous.workerDecision
  ) {
    difficultyChange =
      '遊戲難度由「' +
      previous.workerDecision +
      '」調整為「' +
      latest.workerDecision +
      '」。';
  } else if (latest.workerDecision) {
    difficultyChange =
      '目前維持「' +
      latest.workerDecision +
      '」難度。';
  }

  return {
    history: history,
    summary: summary,
    difference: difference,
    difficultyChange: difficultyChange
  };
}

function testClientHistoryBackend() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('SPMSQ評估');

  if (sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert(
      '目前沒有SPMSQ評估資料。'
    );
    return;
  }

  const clientId = sheet
    .getRange(sheet.getLastRow(), 2)
    .getDisplayValue();

  const result = getClientHistory(clientId);

  SpreadsheetApp.getUi().alert(
    '歷次評估後端測試成功',
    [
      '找到評估紀錄：' +
        result.history.length +
        ' 筆',
      '',
      result.summary,
      '',
      result.difficultyChange
    ].join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/** 活動安排排序；非認知或醫療評分。 */
function sortGamesForClientActivity_(games, activityRows, clientLevel) {
  const timeOf = row => {
    const value = row[11] || row[2];
    const time = value instanceof Date ? value.getTime() : new Date(value).getTime();
    return Number.isFinite(time) ? time : 0;
  };
  const ordered = activityRows.map((row, index) => ({ row: row, index: index }))
    .sort((a, b) => timeOf(b.row) - timeOf(a.row) || b.index - a.index)
    .map(item => item.row);
  const recentSix = ordered.slice(0, 6);
  const recentThree = ordered.slice(0, 3);
  games.forEach((game, index) => {
    const sameGame = ordered.filter(row => String(row[3]) === game.id);
    const recentGame = recentSix.filter(row => String(row[3]) === game.id);
    const sameLevel = recentGame.filter(row => String(row[5]) === clientLevel);
    const latestMood = recentGame.length ? String(recentGame[0][8] || '') : '';
    const lowCompletion = sameLevel.length >= 2 && sameLevel.slice(0, 2).every(row =>
      row[6] !== '' && row[6] !== null && Number.isFinite(Number(row[6])) && Number(row[6]) < 60);
    game.activitySort = {
      caution: latestMood === '抗拒' ? 2 : lowCompletion ? 1 : 0,
      recent: recentThree.some(row => String(row[3]) === game.id) ? 1 : 0,
      count: recentGame.length,
      lastTime: sameGame.length ? timeOf(sameGame[0]) : 0,
      index: index
    };
    if (latestMood === '抗拒') {
      game.recommendationReason = '近期出現抗拒，請先確認參與意願。';
    } else if (lowCompletion) {
      game.recommendationReason = '同難度最近兩次完成度低於60%，請留意疲累或活動困難。';
    } else if (!sameGame.length) {
      game.recommendationReason = '可嘗試新活動：尚無遊玩紀錄。';
    } else if (!game.activitySort.recent) {
      game.recommendationReason = '建議換換遊戲：最近3次未玩，最近6次玩過' + recentGame.length + '次。';
    } else {
      game.recommendationReason = '最近6次玩過' + recentGame.length + '次，可依長輩意願選擇。';
    }
    if (latestMood === '不安') game.recommendationReason += ' 上次出現不安，請先確認當日狀況。';
  });
  games.sort((a, b) => {
    const x = a.activitySort, y = b.activitySort;
    return x.caution - y.caution || x.recent - y.recent || x.count - y.count ||
      a.sortDistance - b.sortDistance || x.lastTime - y.lastTime || x.index - y.index;
  });
  games.forEach(game => { delete game.activitySort; });
  return games;
}

function getRecommendedGames(clientId) {
  const cleanClientId = cleanWebText(clientId);
  const clientRow = findClientRow(cleanClientId);

  if (!clientRow) {
    throw new Error('找不到對應的個案資料。');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clientSheet = ss.getSheetByName('個案資料');
  const gameSheet = ss.getSheetByName('遊戲設定');

  const clientValues = clientSheet
    .getRange(clientRow, 1, 1, 13)
    .getValues()[0];

  const clientStationName = String(clientValues[12] || '').trim();
  const isTestStation = isNewGameTestStation_(clientStationName);

  if (String(clientValues[9] || '') === '已結案') {
    throw new Error('此個案已結束服務，無法開始新的認知活動。');
  }

  let clientLevel = String(
    clientValues[10] || '待評估'
  );

  if (
    !['入門', '標準', '進階'].includes(clientLevel)
  ) {
    clientLevel = '入門';
  }

  if (gameSheet.getLastRow() < 2) {
    return {
      client: {
        id: cleanClientId,
        name: String(clientValues[2] || ''),
        difficulty: clientLevel
      },
      games: []
    };
  }

  const levelRank = {
    '入門': 1,
    '標準': 2,
    '進階': 3
  };

  const rows = gameSheet
    .getRange(
      2,
      1,
      gameSheet.getLastRow() - 1,
      14
    )
    .getValues()
    .filter(row => {
      if (!row[0] || String(row[3]) !== '啟用') return false;
      const gameId = String(row[0]);
      const isTestGame = NEW_GAME_TEST_IDS.includes(gameId);
      return isTestStation ? isTestGame : !isTestGame;
    });

  // 同個案的最近活動；中途放棄也列入輪替次數。
  const activitySheet = ss.getSheetByName('活動紀錄');
  const activityRows = activitySheet && activitySheet.getLastRow() >= 2
    ? activitySheet.getRange(2, 1, activitySheet.getLastRow() - 1,
        Math.min(23, activitySheet.getLastColumn())).getValues()
        .filter(row => row[0] && String(row[1]) === cleanClientId)
    : [];

  const games = rows
    .map(row => {
      const defaultLevel =
        String(row[4] || '入門');

      const originalItems =
        Number(row[5] || 4);

      let suggestedItems = originalItems;
      let guidance = '';

      if (clientLevel === '入門') {
        suggestedItems = Math.max(
          2,
          Math.min(originalItems, 4)
        );

        guidance =
          '使用熟悉題材，不強調速度；一次提供一個步驟，必要時給予口語提示。';
      }

      if (clientLevel === '標準') {
        suggestedItems = originalItems;

        guidance =
          '先讓長者自行嘗試，再依反應提供簡短提示。';
      }

      if (clientLevel === '進階') {
        suggestedItems = Math.min(
          20,
          originalItems + 2
        );

        guidance =
          '可逐步增加項目或變化，但仍需依當日精神與情緒調整。';
      }

      return {
        id: String(row[0]),
        name: String(row[1]),
        skill: String(row[2]),
        defaultLevel: defaultLevel,
        suggestedLevel: clientLevel,
        suggestedItems: suggestedItems,
        duration: Number(row[6] || 0),
        originalHint: String(row[7] || ''),
        moduleKey: String(row[10] || ''),
        version: String(row[11] || ''),
        mode: String(row[12] || '單人'),
        playable: String(row[13] || '') === '可玩',
        isTestGame: NEW_GAME_TEST_IDS.includes(String(row[0])),
        guidance: guidance,
        sortDistance: Math.abs(
          levelRank[defaultLevel] -
          levelRank[clientLevel]
        )
      };
    })
    ;
  sortGamesForClientActivity_(games, activityRows, clientLevel);

  return {
    client: {
      id: cleanClientId,
      code: String(clientValues[1] || ''),
      name: String(clientValues[2] || ''),
      difficulty: clientLevel,
      stationName: clientStationName,
      isTestStation: isTestStation
    },
    games: games
  };
}

function saveActivityFromWeb(activityData) {
  if (!activityData) {
    throw new Error('沒有收到活動紀錄資料。');
  }

  const clientId = String(activityData.clientId || '').trim();
  const gameId = String(activityData.gameId || '').trim();
  const level = String(activityData.level || '').trim();
  const help = String(activityData.help || '').trim();
  const mood = String(activityData.mood || '').trim();
  const note = String(activityData.note || '').trim();
  const submittedWorker = String(activityData.worker || '').trim();
  const deviceId = String(activityData.deviceId || '').trim();
  const requestedStationName = String(activityData.stationName || '').trim();
  const gameResult = activityData.gameResult || {};

  const completionNumber = Number(activityData.completion);

  if (!clientId) {
    throw new Error('請選擇個案。');
  }

  if (!gameId) {
    throw new Error('請選擇遊戲。');
  }

  if (!deviceId) {
    throw new Error('這台設備尚未完成登錄。');
  }

  if (!['入門', '標準', '進階'].includes(level)) {
    throw new Error('遊戲難度資料不正確。');
  }

  if (
    !Number.isFinite(completionNumber) ||
    completionNumber < 0 ||
    completionNumber > 100
  ) {
    throw new Error('完成度必須介於0至100。');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 確認個案存在
  const clientRow = findClientRow(clientId);
  if (!clientRow) {
    throw new Error('找不到所選個案。');
  }

  // 從遊戲設定取得正式遊戲名稱
  const gameSheet = ss.getSheetByName('遊戲設定');

  if (!gameSheet || gameSheet.getLastRow() < 2) {
    throw new Error('找不到遊戲設定資料。');
  }

  const games = gameSheet
    .getRange(2, 1, gameSheet.getLastRow() - 1, 14)
    .getValues();

  const game = games.find(row =>
    String(row[0]) === gameId
  );

  if (!game) {
    throw new Error('找不到所選遊戲。');
  }

  if (String(game[3]) !== '啟用') {
    throw new Error('此遊戲目前未啟用。');
  }

  const gameName = String(game[1]);
  const gameVersion = String(game[11] || '1.0.0');

  const resultScore = clampActivityNumber(gameResult.score, 0, 1000);
  const durationSeconds = clampActivityNumber(
    gameResult.durationSeconds,
    0,
    86400
  );
  const correctCount = clampActivityNumber(
    gameResult.correctCount,
    0,
    10000
  );
  const wrongCount = clampActivityNumber(
    gameResult.wrongCount,
    0,
    10000
  );
  const hintCount = clampActivityNumber(
    gameResult.hintCount,
    0,
    10000
  );
  const gameCompleted = gameResult.completed === true ? '是' : '否';
  const resultDetails = JSON.stringify({
    moduleKey: String(game[10] || ''),
    mode: String(game[12] || '單人'),
    details: gameResult.details || {}
  }).slice(0, 50000);

  // 確認設備存在且仍為啟用狀態
  const deviceSheet = ss.getSheetByName('裝置設定');

  if (!deviceSheet || deviceSheet.getLastRow() < 2) {
    throw new Error('找不到設備登錄資料。');
  }

  const deviceRows = deviceSheet
    .getRange(2, 1, deviceSheet.getLastRow() - 1, 8)
    .getValues();

  const deviceIndex = deviceRows.findIndex(row =>
    String(row[0]) === deviceId
  );

  if (deviceIndex === -1) {
    throw new Error('此設備尚未登錄。');
  }

  const deviceRow = deviceRows[deviceIndex];

  if (String(deviceRow[3]) !== '啟用') {
    throw new Error('此設備目前已停用，無法儲存活動紀錄。');
  }

  // 設備只代表實體設備；站別以本次使用者選取內容為準。
  const deviceName = String(deviceRow[1]);
  const stationName = getActiveStationName_(
    requestedStationName,
    String(deviceRow[2] || '')
  );
  const isTestStation = isNewGameTestStation_(stationName);
  const isTestGame = NEW_GAME_TEST_IDS.includes(gameId);

  if (isTestGame && !isTestStation) {
    throw new Error('此遊戲仍在測試中，只能於站點999使用。');
  }

  if (isTestStation && !isTestGame) {
    throw new Error('站點999僅供測試新遊戲。');
  }

  const defaultWorker = String(deviceRow[6] || '').trim();
  const worker = submittedWorker || defaultWorker;

  if (!worker) {
    throw new Error('請輸入帶領人員。');
  }

  const clientSheet = ss.getSheetByName('個案資料');
  const clientStatus = String(
    clientSheet.getRange(clientRow, 10).getDisplayValue() || ''
  ).trim();

  if (clientStatus === '已結案') {
    throw new Error('此個案已結束服務，無法儲存新的活動紀錄。');
  }

  const clientStation = String(
    clientSheet.getRange(clientRow, 13).getDisplayValue() || ''
  ).trim();

  if (!clientStation) {
    throw new Error('此個案尚未設定站別，請先在「個案資料」M欄補上站別。');
  }

  if (clientStation !== stationName) {
    throw new Error('此個案不屬於目前設備的站別，無法儲存活動紀錄。');
  }

  const activitySheet = ss.getSheetByName('活動紀錄');

  if (!activitySheet) {
    throw new Error('找不到「活動紀錄」工作表。');
  }

  const now = new Date();
  const activityId = createId('A');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    activitySheet.appendRow([
      activityId,              // A 活動ID
      clientId,                // B 個案ID
      now,                     // C 活動日期
      gameId,                  // D 遊戲ID
      gameName,                // E 遊戲名稱
      level,                   // F 使用難度
      completionNumber,        // G 完成度
      help,                    // H 協助程度
      mood,                    // I 情緒反應
      note,                    // J 服務觀察
      worker,                  // K 帶領人員
      now,                     // L 建立時間
      stationName,             // M 站別
      deviceId,                // N 裝置ID
      deviceName,              // O 裝置名稱
      gameVersion,             // P 遊戲版本
      resultScore,             // Q 遊戲得分
      durationSeconds,         // R 遊玩秒數
      correctCount,            // S 答對數
      wrongCount,              // T 答錯數
      hintCount,               // U 提示次數
      gameCompleted,           // V 是否完成
      resultDetails            // W 遊戲結果明細
    ]);

    const newRow = activitySheet.getLastRow();

    activitySheet
      .getRange(newRow, 3)
      .setNumberFormat('yyyy/mm/dd');

    activitySheet
      .getRange(newRow, 12)
      .setNumberFormat('yyyy/mm/dd hh:mm:ss');

    // 更新設備最後使用時間
    deviceSheet
      .getRange(deviceIndex + 2, 6)
      .setValue(now)
      .setNumberFormat('yyyy/mm/dd hh:mm:ss');

    // C欄僅記錄最近使用站別，不再限制此設備只能在單一站別使用。
    deviceSheet
      .getRange(deviceIndex + 2, 3)
      .setValue(stationName);

    return {
      success: true,
      activityId: activityId,
      clientId: clientId,
      gameName: gameName,
      stationName: stationName,
      deviceName: deviceName,
      gameCompleted: gameCompleted === '是'
    };
  } finally {
    lock.releaseLock();
  }
}

function clampActivityNumber(value, minimum, maximum) {
  const number = Number(value);

  if (!Number.isFinite(number)) {
    return 0;
  }

  return Math.round(Math.min(maximum, Math.max(minimum, number)));
}

function getActivityHistory(clientId) {
  const cleanClientId = cleanWebText(clientId);
  const clientRow = findClientRow(cleanClientId);

  if (!clientRow) {
    throw new Error('找不到對應的個案資料。');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clientSheet = ss.getSheetByName('個案資料');
  const clientDifficulty = String(
    clientSheet.getRange(clientRow, 11).getDisplayValue() || '待評估'
  );
  const sheet = ss.getSheetByName('活動紀錄');

  if (sheet.getLastRow() < 2) {
    return {
      history: [],
      averageCompletion: 0,
      summary: '尚無認知活動紀錄。',
      gameCount: 0,
      shouldReviewDifficulty: false,
      clientDifficulty: clientDifficulty
    };
  }

  const rows = sheet
    .getRange(
      2,
      1,
      sheet.getLastRow() - 1,
      22
    )
    .getValues()
    .filter(row =>
      String(row[1]) === cleanClientId
    )
    .sort((a, b) =>
      new Date(b[2]).getTime() -
      new Date(a[2]).getTime()
    );

  // 每一筆活動紀錄都列入三次檢視，包括中途放棄。
  const gameCount = rows.length;

  const shouldReviewDifficulty =
    gameCount > 0 &&
    gameCount % 3 === 0;

  const history = rows
    .slice(0, 10)
    .map(row => ({
      activityId: String(row[0] || ''),
      date: formatDateForWeb(row[2]),
      gameId: String(row[3] || ''),
      gameName: String(row[4] || ''),
      level: String(row[5] || ''),
      completion: Number(row[6] || 0),
      help: String(row[7] || ''),
      mood: String(row[8] || ''),
      note: String(row[9] || ''),
      worker: String(row[10] || '')
    }));

  if (history.length === 0) {
    return {
      history: [],
      averageCompletion: 0,
      summary: '尚無認知活動紀錄。',
      gameCount: gameCount,
      shouldReviewDifficulty: shouldReviewDifficulty,
      clientDifficulty: clientDifficulty
    };
  }

  const averageCompletion = Math.round(
    history.reduce(
      (sum, item) => sum + item.completion,
      0
    ) / history.length
  );

  let summary =
    '近期平均完成度為 ' +
    averageCompletion +
    '%。';

  if (history.length >= 2) {
    const difference =
      history[0].completion -
      history[1].completion;

    if (difference > 0) {
      summary +=
        ' 最近一次較前次增加 ' +
        difference +
        ' 個百分點。';
    } else if (difference < 0) {
      summary +=
        ' 最近一次較前次減少 ' +
        Math.abs(difference) +
        ' 個百分點，建議搭配當日精神、情緒與協助程度觀察。';
    } else {
      summary +=
        ' 最近兩次完成度相同。';
    }
  }

  return {
    history: history,
    averageCompletion: averageCompletion,
    summary: summary,
    gameCount: gameCount,
    shouldReviewDifficulty: shouldReviewDifficulty,
    clientDifficulty: clientDifficulty
  };
}

function findGameById(gameId) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('遊戲設定');

  if (sheet.getLastRow() < 2) {
    return null;
  }

  const rows = sheet
    .getRange(
      2,
      1,
      sheet.getLastRow() - 1,
      14
    )
    .getValues();

  const row = rows.find(item =>
    String(item[0]) === String(gameId)
  );

  if (!row) {
    return null;
  }

  return {
    id: String(row[0]),
    name: String(row[1]),
    skill: String(row[2]),
    status: String(row[3]),
    level: String(row[4]),
    items: Number(row[5] || 0),
    duration: Number(row[6] || 0),
    hint: String(row[7] || ''),
    moduleKey: String(row[10] || ''),
    version: String(row[11] || ''),
    mode: String(row[12] || '單人'),
    playable: String(row[13] || '') === '可玩'
  };
}

function testGameBackend() {
  const assessmentSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('SPMSQ評估');

  if (assessmentSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert(
      '目前沒有可測試的SPMSQ個案。'
    );
    return;
  }

  const clientId = assessmentSheet
    .getRange(
      assessmentSheet.getLastRow(),
      2
    )
    .getDisplayValue();

  const result = getRecommendedGames(clientId);

  SpreadsheetApp.getUi().alert(
    '遊戲推薦後端測試成功',
    [
      '個案：' + result.client.name,
      '目前難度：' + result.client.difficulty,
      '可推薦遊戲：' + result.games.length + '款',
      '',
      result.games.length
        ? '第一順位：' + result.games[0].name
        : '目前沒有啟用中的遊戲'
    ].join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
/**
 * 建立或升級站別／設備辨識需要的資料欄位。
 * 可重複執行，不會覆蓋既有活動紀錄。
 */
function setupDeviceTracking() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const clientSheet = ss.getSheetByName('個案資料');
  const activitySheet = ss.getSheetByName('活動紀錄');

  if (!clientSheet || !activitySheet) {
    throw new Error('找不到「個案資料」或「活動紀錄」工作表。');
  }

  // 一、在個案資料加入站別欄，既有資料盡量從活動紀錄回填。
  clientSheet
    .getRange(1, 13)
    .setValue('站別')
    .setBackground('#2e6856')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  const stationByClient = {};

  if (activitySheet.getLastRow() >= 2) {
    const activityRows = activitySheet
      .getRange(2, 1, activitySheet.getLastRow() - 1, 13)
      .getDisplayValues();

    activityRows.forEach(row => {
      const clientId = String(row[1] || '').trim();
      const stationName = String(row[12] || '').trim();

      if (!clientId || !stationName) return;
      if (!stationByClient[clientId]) {
        stationByClient[clientId] = {};
      }
      stationByClient[clientId][stationName] = true;
    });
  }

  let inferredClientCount = 0;
  let unresolvedClientCount = 0;

  if (clientSheet.getLastRow() >= 2) {
    const clientRows = clientSheet
      .getRange(2, 1, clientSheet.getLastRow() - 1, 13)
      .getValues();

    clientRows.forEach((row, index) => {
      if (!row[0] || String(row[12] || '').trim()) return;

      const stations = Object.keys(
        stationByClient[String(row[0])] || {}
      );

      if (stations.length === 1) {
        clientSheet.getRange(index + 2, 13).setValue(stations[0]);
        inferredClientCount += 1;
      } else {
        unresolvedClientCount += 1;
      }
    });
  }

  clientSheet.autoResizeColumn(13);

  // 建立正式站別清單，供網頁搜尋與確認站名。
  const stationDirectorySheet = ensureStationDirectorySheet_();

  // 二、建立／升級裝置設定工作表。
  let deviceSheet = ss.getSheetByName('裝置設定');

  const deviceHeaders = [
    '裝置ID',
    '裝置名稱',
    '最近使用站別',
    '啟用狀態',
    '首次登錄時間',
    '最後使用時間',
    '預設帶領人員',
    '設備連接碼'
  ];

  if (!deviceSheet) {
    deviceSheet = ss.insertSheet('裝置設定');
  }

  deviceSheet
    .getRange(1, 1, 1, deviceHeaders.length)
    .setValues([deviceHeaders])
    .setBackground('#2e6856')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  deviceSheet.setFrozenRows(1);
  deviceSheet.setRowHeight(1, 42);
  deviceSheet.autoResizeColumns(1, deviceHeaders.length);

  const statusRule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(['啟用', '停用'], true)
    .setAllowInvalid(false)
    .build();

  deviceSheet
    .getRange(
      2,
      4,
      Math.max(deviceSheet.getMaxRows() - 1, 1),
      1
    )
    .setDataValidation(statusRule);

  if (deviceSheet.getLastRow() >= 2) {
    const deviceRows = deviceSheet
      .getRange(2, 1, deviceSheet.getLastRow() - 1, 8)
      .getValues();

    const usedCodes = {};
    deviceRows.forEach(row => {
      const code = normalizeDevicePairingCode_(row[7]);
      if (code) usedCodes[code] = true;
    });

    deviceRows.forEach((row, index) => {
      if (!row[0] || row[7]) return;
      const code = createDevicePairingCode_(usedCodes);
      usedCodes[code] = true;
      deviceSheet.getRange(index + 2, 8).setValue(code);
    });
  }

  // 三、確認活動紀錄保留原有站別與裝置欄位。
  activitySheet
    .getRange(1, 13, 1, 3)
    .setValues([[
      '站別',
      '裝置ID',
      '裝置名稱'
    ]])
    .setBackground('#2e6856')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  activitySheet.autoResizeColumns(13, 3);

  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    '站別與設備設定完成',
    [
      '個案站別已由既有活動紀錄回填：' + inferredClientCount + '筆',
      '仍需人工在「個案資料」M欄補站別：' + unresolvedClientCount + '筆',
      '站別清單：' + Math.max(stationDirectorySheet.getLastRow() - 1, 0) + '筆',
      '設備可切換站別，C欄只保留最近使用站別。'
    ].join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function normalizeDevicePairingCode_(value) {
  return String(value || '')
    .trim()
    .toUpperCase()
    .replace(/\s+/g, '-');
}

function createDevicePairingCode_(usedCodes) {
  usedCodes = usedCodes || {};
  let code = '';

  do {
    code = String(Math.floor(100000 + Math.random() * 900000));
  } while (usedCodes[code]);

  return code;
}

function isNewGameTestStation_(stationName) {
  return String(stationName || '').trim() === NEW_GAME_TEST_STATION_NAME;
}

function setupNewGameTestStation() {
  const result = ensureNewGameTestStation_();

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    '新遊戲測試站已建立',
    [
      '站點代號：' + NEW_GAME_TEST_STATION_CODE,
      '站點名稱：' + NEW_GAME_TEST_STATION_NAME,
      '測試個案：' + result.clientCount + '位（入門、標準、進階）',
      '此站目前只顯示測試中的新遊戲。'
    ].join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function ensureNewGameTestStation_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stationSheet = ensureStationDirectorySheet_();
  const stationRows = stationSheet.getLastRow() >= 2
    ? stationSheet
        .getRange(2, 1, stationSheet.getLastRow() - 1, 2)
        .getDisplayValues()
    : [];
  const stationIndex = stationRows.findIndex(row =>
    String(row[0] || '').trim() === NEW_GAME_TEST_STATION_NAME ||
    String(row[1] || '').trim().padStart(3, '0') === NEW_GAME_TEST_STATION_CODE
  );

  if (stationIndex === -1) {
    stationSheet.appendRow([
      NEW_GAME_TEST_STATION_NAME,
      NEW_GAME_TEST_STATION_CODE
    ]);
  } else {
    stationSheet
      .getRange(stationIndex + 2, 1, 1, 2)
      .setValues([[
        NEW_GAME_TEST_STATION_NAME,
        NEW_GAME_TEST_STATION_CODE
      ]]);
  }

  stationSheet
    .getRange(2, 2, Math.max(stationSheet.getLastRow() - 1, 1), 1)
    .setNumberFormat('@');

  const clientSheet = ss.getSheetByName('個案資料');

  if (!clientSheet) {
    throw new Error('找不到「個案資料」工作表。');
  }

  const now = new Date();
  const definitions = [
    ['C-TEST-G008-EASY', 'TEST-G008-EASY', '新遊戲測試－入門', '入門'],
    ['C-TEST-G008-MEDIUM', 'TEST-G008-MEDIUM', '新遊戲測試－標準', '標準'],
    ['C-TEST-G008-HARD', 'TEST-G008-HARD', '新遊戲測試－進階', '進階']
  ];
  const existingRows = clientSheet.getLastRow() >= 2
    ? clientSheet
        .getRange(2, 1, clientSheet.getLastRow() - 1, 13)
        .getValues()
    : [];

  definitions.forEach(definition => {
    const rowIndex = existingRows.findIndex(row =>
      String(row[0] || '') === definition[0] ||
      String(row[1] || '') === definition[1]
    );
    const row = [
      definition[0],
      definition[1],
      definition[2],
      new Date(1945, 0, 1),
      80,
      '國小以下',
      '測試人員',
      '站點999專用虛構測試資料，不得作正式服務紀錄。',
      now,
      '服務中',
      definition[3],
      '',
      NEW_GAME_TEST_STATION_NAME
    ];

    if (rowIndex === -1) {
      clientSheet.appendRow(row);
      existingRows.push(row);
    } else {
      const sheetRow = rowIndex + 2;
      const createdAt = existingRows[rowIndex][8] || now;
      row[8] = createdAt;
      clientSheet.getRange(sheetRow, 1, 1, 13).setValues([row]);
      existingRows[rowIndex] = row;
    }
  });

  stationSheet.autoResizeColumns(1, 2);
  clientSheet.getRange('D2:D').setNumberFormat('yyyy/mm/dd');
  clientSheet.getRange('I2:I').setNumberFormat('yyyy/mm/dd');

  return {
    stationName: NEW_GAME_TEST_STATION_NAME,
    stationCode: NEW_GAME_TEST_STATION_CODE,
    clientCount: definitions.length
  };
}

function ensureStationDirectorySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('站別清單');
  let needsSetup = false;

  if (!sheet) {
    sheet = ss.insertSheet('站別清單');
    needsSetup = true;
  }

  if (sheet.getLastRow() < 2) {
    needsSetup = true;
  }

  // 站別清單建立完成後，後續查詢只能讀取，不再每次重寫標題、凍結列及
  // 自動調整欄寬。首頁同時發出多個請求時，可避免不必要的試算表寫入互卡。
  if (needsSetup) {
    sheet
      .getRange(1, 1, 1, 2)
      .setValues([['站別', '站別代碼']])
      .setBackground('#2e6856')
      .setFontColor('#ffffff')
      .setFontWeight('bold');

    const rows = getBuiltInStationDirectory_();

    if (rows.length) {
      sheet.getRange(2, 2, rows.length, 1).setNumberFormat('@');
      sheet.getRange(2, 1, rows.length, 2).setValues(rows);
    }

    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 2);
  }

  return sheet;
}

function getStationOptions() {
  const sheet = ensureStationDirectorySheet_();

  if (sheet.getLastRow() < 2) return [];

  return sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 2)
    .getDisplayValues()
    .filter(row => row[0])
    .map(row => ({
      name: String(row[0]).trim(),
      code: String(row[1] || '').trim().padStart(3, '0')
    }));
}

function getActiveStationName_(requestedStationName, fallbackStationName) {
  const requested = String(requestedStationName || '').trim();
  const fallback = String(fallbackStationName || '').trim();
  const value = requested || fallback;

  if (!value) {
    throw new Error('請選擇本次使用站別。');
  }

  const options = getStationOptions();
  const normalizedCode = value.replace(/\D/g, '').padStart(3, '0');
  const matched = options.find(item =>
    item.name === value ||
    (value.length <= 3 && item.code === normalizedCode)
  );

  if (!matched) {
    throw new Error('請從站別清單選擇正式站名，不能只輸入簡稱。');
  }

  return matched.name;
}

function getBuiltInStationDirectory_() {
  const text = `
台中A區中區站	400
台中A區北屯2站	030
台中A區北屯站	406
台中A區北區站	404
台中A區石岡站	422
台中A區西屯2站	005
台中A區西屯站	407
台中A區西區2站	043
台中A區西區站	403
台中A區和平站	424
台中A區東區站	401
台中A區東勢站	423
台中A區南屯2站	045
台中A區南屯站	408
台中A區南區站	402
台中A區烏日站	414
台中A區新社站	426
台中A區霧峰站	413
台中B區大甲站	437
台中B區大安站	439
台中B區大肚站	432
台中B區大里2站	032
台中B區大里站	412
台中B區大雅站	428
台中B區太平2站	029
台中B區太平站	411
台中B區外埔站	438
台中B區后里站	421
台中B區沙鹿站	433
台中B區神岡站	429
台中B區梧棲站	435
台中B區清水站	436
台中B區潭子2站	042
台中B區潭子站	427
台中B區龍井站	434
台中B區豐原2站	034
台中B區豐原站	420
台北市士林2站	028
台北市士林站	111
台北市大同站	103
台北市大安2站	002
台北市大安站	106
台北市中山2站	008
台北市中山站	104
台北市中正2站	000
台北市中正站	100
台北市內湖2站	018
台北市內湖站	114
台北市文山站	116
台北市北投站	112
台北市松山2站	001
台北市松山站	105
台北市信義站	110
台北市南港站	115
台北市萬華2站	022
台北市萬華站	108
台南A下營	735
台南A大內	742
台南A山上	743
台南A中西	700
台南A北門	727
台南A北區	704
台南A北區2站	039
台南A白河	732
台南A安平	708
台南A安南	709
台南A安南2站	027
台南A官田	720
台南A東山	733
台南A東區	701
台南A東區2站	015
台南A南區	702
台南A後壁	731
台南A將軍	725
台南A麻豆	721
台南A新化	712
台南A學甲	726
台南B七股	724
台南B仁德	717
台南B左鎮	713
台南B永康1站	710
台南B永康2站	012
台南B玉井	714
台南B安定	745
台南B西港	723
台南B佳里	722
台南B南化	716
台南B柳營	736
台南B善化	741
台南B新市	744
台南B新營	730
台南B楠西	715
台南B龍崎	719
台南B歸仁1站	711
台南B歸仁2站	051
台南B關廟	718
台南B鹽水	737
宜蘭縣三星站	266
宜蘭縣大同站	267
宜蘭縣五結站	268
宜蘭縣冬山站	269
宜蘭縣壯圍站	263
宜蘭縣宜蘭2站	049
宜蘭縣宜蘭站	260
宜蘭縣南澳站	272
宜蘭縣員山站	264
宜蘭縣頭城站	261
宜蘭縣礁溪站	262
宜蘭縣羅東2站	055
宜蘭縣蘇澳站	270
東縣大武	965
東縣太麻里	963
東縣台東	950
東縣台東2站	041
東縣成功	961
東縣池上	958
東縣卑南	954
東縣延平	953
東縣東河	959
東縣金峰	964
東縣長濱	962
東縣海端	957
東縣鹿野	955
東縣達仁	966
東縣綠島	951
東縣關山	956
花蓮縣玉里站	981
花蓮縣光復站	976
花蓮縣吉安2站	050
花蓮縣吉安站	973
花蓮縣秀林站	972
花蓮縣卓溪站	982
花蓮縣花蓮市2站	040
花蓮縣花蓮站	970
花蓮縣富里站	983
花蓮縣新城站	971
花蓮縣瑞穗站	978
花蓮縣萬榮站	979
花蓮縣壽豐站	974
花蓮縣鳳林站	975
花蓮縣豐濱站	977
金門金沙	890
金門金城	893
金門金湖	891
金門金寧	892
金門烈嶼	894
南投縣中寮站	541
南投縣仁愛站	546
南投縣水里站	553
南投縣名間站	551
南投縣竹山站	557
南投縣信義站	556
南投縣南投2站	056
南投縣南投站	540
南投縣埔里站	545
南投縣草屯站	542
南投縣國姓站	544
南投縣魚池站	555
南投縣鹿谷站	558
南投縣集集站	552
屏縣九如	904
屏縣三地門	901
屏縣內埔	912
屏縣竹田	911
屏縣牡丹	945
屏縣里港	905
屏縣佳冬	931
屏縣東港	928
屏縣枋寮	940
屏縣林邊	927
屏縣長治	908
屏縣南州	926
屏縣屏東	900
屏縣屏東2站	037
屏縣恆春	946
屏縣崁頂	924
屏縣泰武	921
屏縣高樹	906
屏縣新埤	925
屏縣新園	932
屏縣萬丹	913
屏縣萬巒	923
屏縣瑪家	903
屏縣潮州	920
屏縣霧台	902
屏縣麟洛	909
屏縣鹽埔	907
苗栗縣三義站	367
苗栗縣三灣站	352
苗栗縣大湖站	364
苗栗縣公館站	363
苗栗縣竹南1站	350
苗栗縣竹南2站	053
苗栗縣西湖站	368
苗栗縣卓蘭站	369
苗栗縣南庄站	353
苗栗縣後龍站	356
苗栗縣苑裡站	358
苗栗縣苗栗1站	360
苗栗縣苗栗2站	048
苗栗縣通霄站	357
苗栗縣造橋站	361
苗栗縣獅潭站	354
苗栗縣銅鑼站	366
苗栗縣頭份站	351
苗栗縣頭屋站	362
桃園市八德站	334
桃園市大園站	337
桃園市大溪站	335
桃園市中壢1站	010
桃園市中壢2站	320
桃園市平鎮1站	324
桃園市平鎮2站	026
桃園市桃園1站	330
桃園市桃園2站	009
桃園市復興站	336
桃園市新屋站	327
桃園市楊梅站	326
桃園市龍潭站	325
桃園市龜山1站	333
桃園市龜山2站	047
桃園市蘆竹站	338
桃園市觀音站	328
馬祖站	209
高雄A區三民1站	807
高雄A區三民2站	003
高雄A區小港站	812
高雄A區左營1站	813
高雄A區左營2站	007
高雄A區前金站	801
高雄A區前鎮1站	806
高雄A區前鎮2站	031
高雄A區苓雅1站	802
高雄A區苓雅2站	006
高雄A區新興站	800
高雄A區楠梓1站	017
高雄A區楠梓2站	811
高雄A區鼓山1站	804
高雄A區鼓山2站	035
高雄A區旗津站	805
高雄A區鹽埕站	803
高雄B區大社站	815
高雄B區大寮2站	033
高雄B區大寮站	831
高雄B區大樹站	840
高雄B區仁武站	814
高雄B區內門站	845
高雄B區六龜站	844
高雄B區永安區	828
高雄B區田寮站	823
高雄B區甲仙站	847
高雄B區杉林站	846
高雄B區那瑪夏站	849
高雄B區岡山站	820
高雄B區林園站	832
高雄B區阿蓮站	822
高雄B區美濃站	843
高雄B區茂林站	851
高雄B區茄萣站	852
高雄B區桃源站	848
高雄B區梓官站	826
高雄B區鳥松站	833
高雄B區湖內站	829
高雄B區路竹站	821
高雄B區旗山站	842
高雄B區鳳山2站	016
高雄B區鳳山3站	054
高雄B區鳳山站	830
高雄B區橋頭站	825
高雄B區燕巢站	824
高雄B區彌陀站	827
基隆市七堵站	206
基隆市中山站	203
基隆市中正站	202
基隆市仁愛站	200
基隆市安樂站	204
基隆市信義站	201
基隆市暖暖站	205
雲林縣二崙站	649
雲林縣口湖站	653
雲林縣土庫站	633
雲林縣大埤站	631
雲林縣元長站	655
雲林縣斗六1站	640
雲林縣斗六2站	046
雲林縣斗南站	630
雲林縣水林站	652
雲林縣北港站	651
雲林縣古坑站	646
雲林縣台西站	636
雲林縣四湖站	654
雲林縣西螺站	648
雲林縣東勢站	635
雲林縣林內站	643
雲林縣虎尾站	632
雲林縣崙背站	637
雲林縣麥寮站	638
雲林縣莿桐站	647
雲林縣褒忠站	634
新北市八里站	249
新北市三芝站	252
新北市三重1站	241
新北市三重2站	013
新北市三峽站	237
新北市土城站	236
新北市中和1站	235
新北市中和2站	025
新北市五股站	248
新北市永和1站	234
新北市永和2站	024
新北市石門站	253
新北市汐止1站	221
新北市汐止2站	023
新北市板橋1站	220
新北市板橋2站	004
新北市金山站	208
新北市泰山站	243
新北市烏來站	233
新北市貢寮站	228
新北市淡水站	251
新北市深坑站	222
新北市新莊1站	242
新北市新莊2站	014
新北市瑞芳站	224
新北市萬里站	207
新北市樹林1站	238
新北市樹林2站	020
新北市蘆洲1站	247
新北市蘆洲2站	036
新北市鶯歌站	239
新竹縣北埔站	314
新竹縣尖石站	313
新竹縣竹北2站	021
新竹縣竹北站	302
新竹縣竹東站	310
新竹縣芎林站	307
新竹縣峨眉站	315
新竹縣湖口站	303
新竹縣新竹市2站	011
新竹縣新竹市3站	044
新竹縣新竹市站	300
新竹縣新埔站	305
新竹縣新豐站	304
新竹縣橫山站	312
新竹縣關西站	306
新竹縣寶山站	308
嘉義縣大林站	622
嘉義縣大埔站	607
嘉義縣中埔站	606
嘉義縣六腳站	615
嘉義縣太保站	612
嘉義縣水上站	608
嘉義縣布袋站	625
嘉義縣民雄站	621
嘉義縣朴子站	613
嘉義縣竹崎站	604
嘉義縣東石站	614
嘉義縣阿里山站	605
嘉義縣梅山站	603
嘉義縣鹿草站	611
嘉義縣番路站	602
嘉義縣新港站	616
嘉義縣溪口站	623
嘉義縣義竹站	624
嘉義縣嘉義市1站	600
嘉義縣嘉義市2站	052
彰化縣二林站	526
彰化縣大村站	515
彰化縣大城站	527
彰化縣北斗站	521
彰化縣永靖站	512
彰化縣田中站	520
彰化縣田尾站	522
彰化縣竹塘站	525
彰化縣伸港站	509
彰化縣秀水站	504
彰化縣和美站	508
彰化縣社頭站	511
彰化縣芬園站	502
彰化縣花壇站	503
彰化縣芳苑站	528
彰化縣員林站	510
彰化縣埔心站	513
彰化縣埔鹽站	516
彰化縣埤頭站	523
彰化縣鹿港站	505
彰化縣溪州站	524
彰化縣溪湖站	514
彰化縣彰化站	500
彰化縣福興站	506
彰化縣線西站	507
澎湖馬公	880
澎湖湖西	885
澎縣白沙	884
澎縣西嶼	881`;

  return text
    .split('\n')
    .map(line => line.split('\t'))
    .filter(row => row[0])
    .map(row => [row[0], String(row[1] || '').padStart(3, '0')]);
}

function deviceRowToObject_(row) {
  return {
    success: true,
    found: true,
    enabled: String(row[3]) === '啟用',
    deviceId: String(row[0] || ''),
    deviceName: String(row[1] || ''),
    stationName: String(row[2] || ''),
    status: String(row[3] || ''),
    defaultWorker: String(row[6] || ''),
    pairingCode: String(row[7] || '')
  };
}

function getEnabledDevice_(deviceId) {
  const device = getDeviceSetting(deviceId);

  if (!device.found) {
    throw new Error('此設備尚未完成登錄。');
  }

  if (!device.enabled) {
    throw new Error(device.message || '此設備目前已停用。');
  }

  return device;
}

function getFirstActiveDeviceId_() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('裝置設定');

  if (!sheet || sheet.getLastRow() < 2) {
    throw new Error('目前沒有已登錄的設備。');
  }

  const rows = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 4)
    .getDisplayValues();

  const row = rows.find(item =>
    item[0] && item[3] === '啟用'
  );

  if (!row) {
    throw new Error('目前沒有啟用中的設備。');
  }

  return String(row[0]);
}
/**
 * 由網頁登錄或更新設備
 */
function registerDeviceFromWeb(deviceData) {
  if (!deviceData) {
    throw new Error('沒有收到設備資料。');
  }

  const deviceId = String(deviceData.deviceId || '').trim();
  const deviceName = String(deviceData.deviceName || '').trim();
  const requestedStationName = String(deviceData.stationName || '').trim();
  const defaultWorker = String(deviceData.defaultWorker || '').trim();
  const requestedPairingCode = normalizeDevicePairingCode_(
    deviceData.pairingCode
  );

  if (
    requestedPairingCode &&
    !/^[A-Z0-9_-]{4,30}$/.test(requestedPairingCode)
  ) {
    throw new Error('設備連接碼限4至30碼英文字母、數字、連字號或底線。');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('裝置設定');

    if (!sheet) {
      throw new Error(
        '找不到「裝置設定」工作表，請先執行 setupDeviceTracking。'
      );
    }

    const now = new Date();
    const rows = sheet.getLastRow() >= 2
      ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues()
      : [];

    let deviceIndex = rows.findIndex(row =>
      deviceId && String(row[0]) === deviceId
    );

    const matchedByDeviceId = deviceIndex !== -1;

    if (deviceIndex === -1 && requestedPairingCode) {
      deviceIndex = rows.findIndex(row =>
        normalizeDevicePairingCode_(row[7]) === requestedPairingCode
      );
    }

    if (deviceIndex !== -1) {
      const row = rows[deviceIndex];
      const sheetRow = deviceIndex + 2;

      if (String(row[3]) === '停用') {
        throw new Error('此設備已由管理者停用，無法使用。');
      }

      if (matchedByDeviceId) {
        const nextName = deviceName || String(row[1] || '');
        const nextStation = getActiveStationName_(
          requestedStationName,
          String(row[2] || '')
        );
        const nextWorker = defaultWorker || String(row[6] || '');
        const nextCode = requestedPairingCode ||
          normalizeDevicePairingCode_(row[7]) ||
          createDevicePairingCode_();

        const duplicateCode = rows.some((item, index) =>
          index !== deviceIndex &&
          normalizeDevicePairingCode_(item[7]) === nextCode
        );

        if (duplicateCode) {
          throw new Error('此設備連接碼已被其他設備使用。');
        }

        if (!nextName || !nextStation || !nextWorker) {
          throw new Error('請完整填寫本次站別、設備名稱及預設帶領人員。');
        }

        sheet.getRange(sheetRow, 2, 1, 2).setValues([[
          nextName,
          nextStation
        ]]);
        sheet.getRange(sheetRow, 6, 1, 3).setValues([[
          now,
          nextWorker,
          nextCode
        ]]);
      } else {
        // 由另一個瀏覽器輸入連接碼時，直接沿用原設備資料。
        sheet.getRange(sheetRow, 6).setValue(now);
      }

      sheet
        .getRange(sheetRow, 5, 1, 2)
        .setNumberFormat('yyyy/mm/dd hh:mm:ss');

      return deviceRowToObject_(
        sheet.getRange(sheetRow, 1, 1, 8).getValues()[0]
      );
    }

    if (!deviceName || !requestedStationName || !defaultWorker) {
      throw new Error(
        requestedPairingCode
          ? '找不到此連接碼；新增設備時請完整填寫本次站別、設備名稱及預設帶領人員。'
          : '新增設備時請完整填寫本次站別、設備名稱及預設帶領人員。'
      );
    }

    const stationName = getActiveStationName_(requestedStationName, '');

    const newDeviceId = deviceId || 'DEV-' + Utilities.getUuid();
    const usedCodes = {};
    rows.forEach(row => {
      const code = normalizeDevicePairingCode_(row[7]);
      if (code) usedCodes[code] = true;
    });
    const pairingCode = requestedPairingCode ||
      createDevicePairingCode_(usedCodes);

    if (usedCodes[pairingCode]) {
      throw new Error('此設備連接碼已被其他設備使用。');
    }

    sheet.appendRow([
      newDeviceId,
      deviceName,
      stationName,
      '啟用',
      now,
      now,
      defaultWorker,
      pairingCode
    ]);

    const newRow = sheet.getLastRow();
    sheet
      .getRange(newRow, 5, 1, 2)
      .setNumberFormat('yyyy/mm/dd hh:mm:ss');

    return deviceRowToObject_(
      sheet.getRange(newRow, 1, 1, 8).getValues()[0]
    );
  } finally {
    lock.releaseLock();
  }
}


/**
 * 取得已登錄的設備資料
 */
function getDeviceSetting(deviceId) {
  deviceId = String(deviceId || '').trim();

  if (!deviceId) {
    return {
      found: false
    };
  }

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('裝置設定');

  if (!sheet || sheet.getLastRow() < 2) {
    return {
      found: false
    };
  }

  const rows = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 8)
    .getValues();

  const row = rows.find(item =>
    String(item[0]) === deviceId
  );

  if (!row) {
    return {
      found: false
    };
  }

  if (row[3] !== '啟用') {
    return {
      found: true,
      enabled: false,
      message: '此設備目前已停用。'
    };
  }

  const device = deviceRowToObject_(row);
  const stationIsOfficial = getStationOptions().some(item =>
    item.name === device.stationName
  );

  if (!stationIsOfficial) {
    device.stationName = '';
    device.message = '請重新選擇本次使用站別。';
  }

  return device;
}


/**
 * 後端測試
 */
function testDeviceBackend() {
  const result = registerDeviceFromWeb({
    deviceName: '測試用iPad',
    stationName: '台北市士林站',
    defaultWorker: '測試人員',
    pairingCode: 'TEST-DEVICE-001'
  });

  SpreadsheetApp.getUi().alert(
    '設備登錄測試成功',
    [
      '裝置ID：' + result.deviceId,
      '裝置名稱：' + result.deviceName,
      '站別：' + result.stationName,
      '預設帶領人員：' + result.defaultWorker,
      '設備連接碼：' + result.pairingCode
    ].join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 取得最近的活動紀錄
 */
function getRecentActivities(limit, deviceId, stationName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const device = getEnabledDevice_(deviceId);
  const activeStationName = getActiveStationName_(stationName, device.stationName);
  const activitySheet = ss.getSheetByName('活動紀錄');
  const clientSheet = ss.getSheetByName('個案資料');

  const recordLimit = Math.max(
    1,
    Math.min(Number(limit) || 5, 20)
  );

  if (
    !activitySheet ||
    activitySheet.getLastRow() < 2
  ) {
    return [];
  }

  // 建立個案ID與姓名的對照
  const clientMap = {};

  if (clientSheet && clientSheet.getLastRow() >= 2) {
    const clients = clientSheet
      .getRange(
        2,
        1,
        clientSheet.getLastRow() - 1,
        3
      )
      .getDisplayValues();

    clients.forEach(row => {
      clientMap[String(row[0])] = {
        code: String(row[1]),
        name: String(row[2])
      };
    });
  }

  const rows = activitySheet
    .getRange(
      2,
      1,
      activitySheet.getLastRow() - 1,
      15
    )
    .getValues();

  const timeZone =
    ss.getSpreadsheetTimeZone() || 'Asia/Taipei';

  return rows
    .filter(row =>
      row[0] &&
      row[1] &&
      String(row[12] || '').trim() === activeStationName
    )
    .sort((a, b) => {
      const timeA =
        a[11] instanceof Date
          ? a[11].getTime()
          : new Date(a[11] || a[2]).getTime();

      const timeB =
        b[11] instanceof Date
          ? b[11].getTime()
          : new Date(b[11] || b[2]).getTime();

      return timeB - timeA;
    })
    .slice(0, recordLimit)
    .map(row => {
      const clientId = String(row[1]);
      const client = clientMap[clientId] || {};

      const activityDate =
        row[2] instanceof Date
          ? Utilities.formatDate(
              row[2],
              timeZone,
              'yyyy/MM/dd'
            )
          : String(row[2] || '');

      return {
        activityId: String(row[0]),
        clientId: clientId,
        clientCode: client.code || '',
        clientName: client.name || '未命名個案',
        activityDate: activityDate,
        gameId: String(row[3] || ''),
        gameName: String(row[4] || ''),
        level: String(row[5] || ''),
        completion: Number(row[6]) || 0,
        help: String(row[7] || ''),
        mood: String(row[8] || ''),
        note: String(row[9] || ''),
        worker: String(row[10] || ''),
        stationName: String(row[12] || ''),
        deviceId: String(row[13] || ''),
        deviceName: String(row[14] || '')
      };
    });
}


/**
 * 測試近期活動查詢
 */
function testRecentActivities() {
  const deviceId = getFirstActiveDeviceId_();
  const activities = getRecentActivities(5, deviceId);

  if (activities.length === 0) {
    SpreadsheetApp.getUi().alert(
      '查詢完成',
      '目前沒有活動紀錄。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    return;
  }

  const latest = activities[0];

  SpreadsheetApp.getUi().alert(
    '近期活動查詢成功',
    [
      '共讀取：' + activities.length + '筆',
      '個案：' + latest.clientName,
      '遊戲：' + latest.gameName,
      '完成度：' + latest.completion + '%',
      '站別：' + latest.stationName,
      '設備：' + latest.deviceName
    ].join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 取得首頁統計資料
 */
function getDashboardSummary(deviceId, stationName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const device = getEnabledDevice_(deviceId);
  const activeStationName = getActiveStationName_(stationName, device.stationName);

  const clientSheet = ss.getSheetByName('個案資料');
  const gameSheet = ss.getSheetByName('遊戲設定');
  const activitySheet = ss.getSheetByName('活動紀錄');

  let activeClientCount = 0;
  let enabledGameCount = 0;
  let weeklyActivityCount = 0;
  let totalActivityCount = 0;

  // 服務中個案數
  if (clientSheet && clientSheet.getLastRow() >= 2) {
    const clientRows = clientSheet
      .getRange(
        2,
        1,
        clientSheet.getLastRow() - 1,
        13
      )
      .getDisplayValues();

    activeClientCount = clientRows.filter(row =>
      row[0] &&
      row[9] === '服務中' &&
      String(row[12] || '').trim() === activeStationName
    ).length;
  }

  // 啟用遊戲數
  if (gameSheet && gameSheet.getLastRow() >= 2) {
    const gameRows = gameSheet
      .getRange(
        2,
        1,
        gameSheet.getLastRow() - 1,
        4
      )
      .getDisplayValues();

    const isTestStation = isNewGameTestStation_(activeStationName);

    enabledGameCount = gameRows.filter(row => {
      if (!row[0] || row[3] !== '啟用') return false;
      const isTestGame = NEW_GAME_TEST_IDS.includes(String(row[0]));
      return isTestStation ? isTestGame : !isTestGame;
    }).length;
  }

  // 活動紀錄統計
  if (
    activitySheet &&
    activitySheet.getLastRow() >= 2
  ) {
    const activityRows = activitySheet
      .getRange(
        2,
        1,
        activitySheet.getLastRow() - 1,
        13
      )
      .getValues();

    const validRows = activityRows.filter(row =>
      row[0] &&
      String(row[12] || '').trim() === activeStationName
    );

    totalActivityCount = validRows.length;

    const timeZone =
      ss.getSpreadsheetTimeZone() || 'Asia/Taipei';

    const now = new Date();

    // 星期一為每週第一天
    const dayNumber = Number(
      Utilities.formatDate(now, timeZone, 'u')
    );

    const weekStart = new Date(
      now.getTime() -
      (dayNumber - 1) * 24 * 60 * 60 * 1000
    );

    const weekStartKey = Utilities.formatDate(
      weekStart,
      timeZone,
      'yyyyMMdd'
    );

    const todayKey = Utilities.formatDate(
      now,
      timeZone,
      'yyyyMMdd'
    );

    weeklyActivityCount = validRows.filter(row => {
      if (!(row[2] instanceof Date)) {
        return false;
      }

      const activityDateKey = Utilities.formatDate(
        row[2],
        timeZone,
        'yyyyMMdd'
      );

      return (
        activityDateKey >= weekStartKey &&
        activityDateKey <= todayKey
      );
    }).length;
  }

  return {
    activeClientCount: activeClientCount,
    enabledGameCount: enabledGameCount,
    weeklyActivityCount: weeklyActivityCount,
    totalActivityCount: totalActivityCount
  };
}


/**
 * 測試首頁統計
 */
function testDashboardSummary() {
  const summary = getDashboardSummary(getFirstActiveDeviceId_());

  SpreadsheetApp.getUi().alert(
    '首頁統計查詢成功',
    [
      '服務中個案：' +
        summary.activeClientCount +
        '位',

      '啟用遊戲：' +
        summary.enabledGameCount +
        '款',

      '本週活動：' +
        summary.weeklyActivityCount +
        '場',

      '累計活動：' +
        summary.totalActivityCount +
        '場'
    ].join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 取得單一個案的活動歷程與趨勢
 */
function getClientActivityTrend(clientId) {
  clientId = String(clientId || '').trim();

  if (!clientId) {
    throw new Error('缺少個案ID。');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activitySheet = ss.getSheetByName('活動紀錄');
  const clientSheet = ss.getSheetByName('個案資料');

  if (
    !activitySheet ||
    activitySheet.getLastRow() < 2
  ) {
    return {
      clientId: clientId,
      clientName: '',
      totalCount: 0,
      averageCompletion: 0,
      latestCompletion: 0,
      completionChange: 0,
      moreHelpCount: 0,
      unsettledCount: 0,
      records: []
    };
  }

  // 取得個案姓名
  let clientName = '';

  if (clientSheet && clientSheet.getLastRow() >= 2) {
    const clients = clientSheet
      .getRange(
        2,
        1,
        clientSheet.getLastRow() - 1,
        3
      )
      .getDisplayValues();

    const client = clients.find(row =>
      String(row[0]) === clientId
    );

    if (client) {
      clientName = String(client[2] || '');
    }
  }

  const rows = activitySheet
    .getRange(
      2,
      1,
      activitySheet.getLastRow() - 1,
      Math.min(23, activitySheet.getLastColumn())
    )
    .getValues()
    .filter(row =>
      row[0] &&
      String(row[1]) === clientId
    )
    .sort((a, b) => {
      const timeA =
        a[11] instanceof Date
          ? a[11].getTime()
          : new Date(a[11] || a[2]).getTime();

      const timeB =
        b[11] instanceof Date
          ? b[11].getTime()
          : new Date(b[11] || b[2]).getTime();

      return timeA - timeB;
    });

  if (rows.length === 0) {
    return {
      clientId: clientId,
      clientName: clientName,
      totalCount: 0,
      averageCompletion: 0,
      latestCompletion: 0,
      completionChange: 0,
      moreHelpCount: 0,
      unsettledCount: 0,
      records: []
    };
  }

  const timeZone =
    ss.getSpreadsheetTimeZone() || 'Asia/Taipei';

  // 畫面最多呈現最近12筆
  const recentRows = rows.slice(-12);

  const records = recentRows.map(row => ({
    activityId: String(row[0]),
    activityDate:
      row[2] instanceof Date
        ? Utilities.formatDate(
            row[2],
            timeZone,
            'yyyy/MM/dd'
          )
        : String(row[2] || ''),
    gameId: String(row[3] || ''),
    gameName: String(row[4] || ''),
    completionStatus: row[21] === '是' ? '完成' : row[21] === '否' ? '中途放棄' : '未記錄完成狀況',
    level: String(row[5] || ''),
    completion: Number(row[6]) || 0,
    help: String(row[7] || ''),
    mood: String(row[8] || ''),
    note: String(row[9] || ''),
    worker: String(row[10] || ''),
    stationName: String(row[12] || ''),
    deviceName: String(row[14] || '')
  }));

  const lastPlayedByGame = {};
  rows.forEach(row => {
    const date = row[2] instanceof Date
      ? Utilities.formatDate(row[2], timeZone, 'yyyy/MM/dd')
      : String(row[2] || '');
    lastPlayedByGame[String(row[3] || '')] = date;
  });

  const completions = rows.map(row =>
    Number(row[6]) || 0
  );

  const averageCompletion = Math.round(
    completions.reduce(
      (sum, value) => sum + value,
      0
    ) / completions.length
  );

  const latestCompletion =
    completions[completions.length - 1];

  // 最近3次與前3次相比；不足6筆時比較首筆與末筆
  let completionChange = 0;

  if (completions.length >= 6) {
    const recentThree = completions.slice(-3);
    const previousThree = completions.slice(-6, -3);

    const recentAverage =
      recentThree.reduce((sum, value) => sum + value, 0) /
      recentThree.length;

    const previousAverage =
      previousThree.reduce((sum, value) => sum + value, 0) /
      previousThree.length;

    completionChange = Math.round(
      recentAverage - previousAverage
    );
  } else if (completions.length >= 2) {
    completionChange =
      latestCompletion - completions[0];
  }

  const moreHelpCount = rows.filter(row =>
    String(row[7]) === '較多協助'
  ).length;

  const unsettledCount = rows.filter(row =>
    ['不安', '抗拒'].includes(String(row[8]))
  ).length;

  return {
    clientId: clientId,
    clientName: clientName,
    totalCount: rows.length,
    averageCompletion: averageCompletion,
    latestCompletion: latestCompletion,
    completionChange: completionChange,
    moreHelpCount: moreHelpCount,
    unsettledCount: unsettledCount,
    lastPlayedByGame: lastPlayedByGame,
    records: records
  };
}


/**
 * 測試個案活動趨勢
 * 自動使用最近一筆活動的個案
 */
function testClientActivityTrend() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activitySheet = ss.getSheetByName('活動紀錄');

  if (
    !activitySheet ||
    activitySheet.getLastRow() < 2
  ) {
    SpreadsheetApp.getUi().alert(
      '目前沒有可測試的活動紀錄。'
    );

    return;
  }

  const clientId = String(
    activitySheet
      .getRange(activitySheet.getLastRow(), 2)
      .getDisplayValue()
  );

  const result = getClientActivityTrend(clientId);

  SpreadsheetApp.getUi().alert(
    '個案活動趨勢查詢成功',
    [
      '個案：' + (result.clientName || clientId),
      '活動次數：' + result.totalCount + '次',
      '平均完成度：' +
        result.averageCompletion +
        '%',
      '最近完成度：' +
        result.latestCompletion +
        '%',
      '完成度變化：' +
        (
          result.completionChange > 0
            ? '+'
            : ''
        ) +
        result.completionChange +
        '個百分點',
      '較多協助：' +
        result.moreHelpCount +
        '次',
      '不安／抗拒：' +
        result.unsettledCount +
        '次'
    ].join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}