/**
 * NEXUS55 Gate1 - 社内問い合わせ管理システム
 * Copy this file to Code.gs in a spreadsheet-bound Apps Script project.
 * Run setupGate1() once, then reload the spreadsheet.
 */

var GATE1 = Object.freeze({
  SHEETS: Object.freeze({
    INQUIRIES: 'inquiries',
    USERS: 'users',
    HISTORIES: 'inquiry_histories'
  }),
  INQUIRY_HEADERS: Object.freeze([
    'id', 'title', 'body', 'requester_name', 'status', 'assignee_id',
    'priority', 'created_at', 'updated_at', 'closed_at'
  ]),
  USER_HEADERS: Object.freeze(['id', 'name', 'role', 'is_active']),
  HISTORY_HEADERS: Object.freeze([
    'id', 'inquiry_id', 'changed_by', 'field_name', 'old_value',
    'new_value', 'changed_at'
  ]),
  STATUS_LABELS: Object.freeze({
    NEW: '未対応',
    IN_PROGRESS: '対応中',
    PENDING: '保留',
    DONE: '完了'
  }),
  TRANSITIONS: Object.freeze({
    NEW: Object.freeze(['IN_PROGRESS']),
    IN_PROGRESS: Object.freeze(['PENDING', 'DONE']),
    PENDING: Object.freeze(['IN_PROGRESS']),
    DONE: Object.freeze(['IN_PROGRESS'])
  }),
  PRIORITY: Object.freeze({
    HIGH: Object.freeze({ label: '高', color: '#dc2626' }),
    MIDDLE: Object.freeze({ label: '中', color: '#111827' }),
    LOW: Object.freeze({ label: '低', color: '#6b7280' })
  }),
  NAME_SORT_KEYS: Object.freeze({
    '佐藤 花子': 'sato hanako',
    '田中 次郎': 'tanaka jiro',
    '山田 太郎': 'yamada taro'
  }),
  PROPS: Object.freeze({
    CURRENT_USER_ID: 'GATE1_CURRENT_USER_ID',
    CURRENT_INQUIRY_ID: 'GATE1_CURRENT_INQUIRY_ID',
    FORCE_HISTORY_FAILURE: 'GATE1_FORCE_HISTORY_FAILURE'
  }),
  MESSAGE: Object.freeze({
    INVALID_TRANSITION:
      'このステータスへは変更できません。画面を再読み込みしてください。',
    NO_PERMISSION: 'この操作を行う権限がありません。',
    DONE_ASSIGNEE: '完了済みの問い合わせは担当者を変更できません。',
    INACTIVE_USER: '指定されたユーザーは選択できません。',
    COMMENT_TOO_LONG: 'コメントは 200 文字以内で入力してください。',
    NOT_FOUND: '指定された問い合わせは存在しません。',
    CONCURRENT:
      '他のユーザーが更新しました。画面を再読み込みしてください。',
    ROLLBACK: '履歴登録が失敗したため、本体の更新も元に戻しました。',
    UNEXPECTED: '処理中のエラーです。管理者に確認してください。'
  })
});

/** スプレッドシートを開いたときにメニューを追加します。 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('問い合わせ管理')
    .addItem('1. セットアップ', 'setupGate1')
    .addSeparator()
    .addItem('問い合わせ詳細を開く', 'showInquiryDialog')
    .addItem('操作ユーザーを変更する', 'selectCurrentUser')
    .addItem('問い合わせIDを変更する', 'selectCurrentInquiry')
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu('提出用モック')
        .addItem('正常時を表示', 'showNormalMock')
        .addItem('不正遷移エラーを表示', 'showInvalidTransitionMock')
    )
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu('テスト用')
        .addItem('次の履歴登録を失敗させる', 'enableHistoryFailureTest')
    )
    .addToUi();
}

/** Creates demo data. Existing data in the three app sheets is replaced. */
function setupGate1() {
  var ui = SpreadsheetApp.getUi();
  var answer = ui.alert(
    'Gate1 セットアップ',
    'Reset inquiries、users、inquiry_histories. 実行しますか?',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) {
    return;
  }

  setupGate1Data_();
  ui.alert(
    'セットアップが完了しました。スプレッドシートを再読み込みし、' +
      '「問い合わせ管理 > 問い合わせ詳細を開く」を選択してください。'
  );
}

function setupGate1Data_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Tokyo');

  var inquiriesSheet = resetSheet_(
    ss,
    GATE1.SHEETS.INQUIRIES,
    GATE1.INQUIRY_HEADERS
  );
  var usersSheet = resetSheet_(ss, GATE1.SHEETS.USERS, GATE1.USER_HEADERS);
  var historiesSheet = resetSheet_(
    ss,
    GATE1.SHEETS.HISTORIES,
    GATE1.HISTORY_HEADERS
  );

  var created1 = new Date('2026-07-29T01:15:00.000Z');
  var updated1 = new Date('2026-07-29T01:15:00.000Z');
  var created2 = new Date('2026-07-28T05:30:00.000Z');
  var updated2 = new Date('2026-07-28T07:45:00.000Z');

  inquiriesSheet.getRange(2, 1, 2, GATE1.INQUIRY_HEADERS.length).setValues([
    [
      1001,
      'プリンタが印刷できません',
      '3階の複合機で印刷しようとすると、エラーが表示されます。',
      '山田 太郎',
      'NEW',
      '',
      'MIDDLE',
      created1,
      updated1,
      ''
    ],
    [
      1002,
      'PCが起動しない',
      'PCが起動しません。',
      '田中 次郎',
      'DONE',
      2,
      'HIGH',
      created2,
      updated2,
      updated2
    ]
  ]);

  usersSheet.getRange(2, 1, 4, GATE1.USER_HEADERS.length).setValues([
    [1, '佐藤 花子', 'ADMIN', true],
    [2, '山田 太郎', 'MEMBER', true],
    [3, '田中 次郎', 'MEMBER', true],
    [4, '無効 太郎', 'MEMBER', false]
  ]);

  formatDataSheets_(inquiriesSheet, usersSheet, historiesSheet);

  var props = PropertiesService.getUserProperties();
  props.setProperty(GATE1.PROPS.CURRENT_USER_ID, '1');
  props.setProperty(GATE1.PROPS.CURRENT_INQUIRY_ID, '1001');
  props.deleteProperty(GATE1.PROPS.FORCE_HISTORY_FAILURE);
  SpreadsheetApp.flush();
}

function resetSheet_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  sheet.clear();
  sheet.clearConditionalFormatRules();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.setHiddenGridlines(false);
  return sheet;
}

function formatDataSheets_(inquiriesSheet, usersSheet, historiesSheet) {
  [inquiriesSheet, usersSheet, historiesSheet].forEach(function (sheet) {
    var lastColumn = sheet.getLastColumn();
    sheet
      .getRange(1, 1, 1, lastColumn)
      .setBackground('#1e3a5f')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    sheet.autoResizeColumns(1, lastColumn);
  });

  inquiriesSheet.setColumnWidth(2, 220);
  inquiriesSheet.setColumnWidth(3, 360);
  inquiriesSheet.setColumnWidth(4, 120);
  inquiriesSheet.getRange('H:J').setNumberFormat('yyyy/mm/dd hh:mm:ss');
  usersSheet.setColumnWidth(2, 140);
  historiesSheet.getRange('G:G').setNumberFormat('yyyy/mm/dd hh:mm:ss');
}

/** 通常の問い合わせ詳細ダイアログを表示します。 */
function showInquiryDialog() {
  showInquiryDialog_('', '');
}

/** 提出用の正常時モック。セットアップ直後の問い合わせ1001を表示します。 */
function showNormalMock() {
  PropertiesService.getUserProperties().setProperty(
    GATE1.PROPS.CURRENT_INQUIRY_ID,
    '1001'
  );
  showInquiryDialog_('', '');
}

/**
 * Sends an invalid transition and displays the server response.
 * データは更新されません。
 */
function showInvalidTransitionMock() {
  var vm = getDetailViewModel();
  var invalidStatus = firstInvalidTarget_(vm.inquiry.status);
  var result = changeStatus({
    inquiryId: vm.inquiry.id,
    newStatus: invalidStatus,
    comment: '',
    expectedUpdatedAt: vm.inquiry.updatedAtToken
  });
  showInquiryDialog_(result.message, result.ok ? 'success' : 'error');
}

function firstInvalidTarget_(currentStatus) {
  var all = ['NEW', 'IN_PROGRESS', 'PENDING', 'DONE'];
  var allowed = GATE1.TRANSITIONS[currentStatus] || [];
  for (var i = 0; i < all.length; i += 1) {
    if (all[i] !== currentStatus && allowed.indexOf(all[i]) === -1) {
      return all[i];
    }
  }
  return currentStatus;
}

function showInquiryDialog_(message, type) {
  var html = buildDialogHtml_(message, type)
    .setWidth(820)
    .setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(html, '問い合わせ詳細');
}

/** Changes the demo login user. */
function selectCurrentUser() {
  var users = getAllUsers_().filter(function (user) {
    return user.is_active === true;
  });
  var choices = users
    .map(function (user) {
      return user.id + ': ' + user.name + ' (' + user.role + ')';
    })
    .join('\n');
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    '操作ユーザーを変更する',
    choices + '\n\n使用するユーザーIDを入力してください。',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  var selected = findUserById_(response.getResponseText().trim());
  if (!selected || selected.is_active !== true) {
    ui.alert(GATE1.MESSAGE.INACTIVE_USER);
    return;
  }
  PropertiesService.getUserProperties().setProperty(
    GATE1.PROPS.CURRENT_USER_ID,
    String(selected.id)
  );
  ui.alert('操作ユーザーを「' + selected.name + '」に変更しました。');
}

/** Changes the inquiry shown in the dialog. */
function selectCurrentInquiry() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    '問い合わせIDを変更する',
    '表示する問い合わせIDを入力してください。（1001 など）',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  var inquiryId = response.getResponseText().trim();
  if (!findInquiryById_(inquiryId)) {
    ui.alert(GATE1.MESSAGE.NOT_FOUND);
    return;
  }
  PropertiesService.getUserProperties().setProperty(
    GATE1.PROPS.CURRENT_INQUIRY_ID,
    inquiryId
  );
  showInquiryDialog();
}

/** 次の更新だけ履歴登録を失敗させ、ロールバックを確認できるようにします。 */
function enableHistoryFailureTest() {
  PropertiesService.getUserProperties().setProperty(
    GATE1.PROPS.FORCE_HISTORY_FAILURE,
    '1'
  );
  SpreadsheetApp.getUi().alert(
    '次のステータス変更または担当者変更で履歴登録を意図的に失敗させます。' +
      '問い合わせ本体が元に戻ることを確認してください。'
  );
}

/** Returns a serializable view model for the dialog. */
function getDetailViewModel() {
  var actor = getCurrentActor_();
  var inquiryId = getCurrentInquiryId_();
  var inquiry = findInquiryById_(inquiryId);
  if (!inquiry) {
    throw new Error(GATE1.MESSAGE.NOT_FOUND);
  }

  var assignee = inquiry.assignee_id === ''
    ? null
    : findUserById_(inquiry.assignee_id);
  var activeUsers = getAllUsers_()
    .filter(function (user) {
      return user.is_active === true;
    })
    .sort(function (a, b) {
      return japaneseNameSortKey_(a.name).localeCompare(
        japaneseNameSortKey_(b.name),
        'en'
      );
    })
    .map(function (user) {
      return { id: String(user.id), name: user.name, role: user.role };
    });

  var nextStatuses = (GATE1.TRANSITIONS[inquiry.status] || []).map(
    function (status) {
      return { code: status, label: GATE1.STATUS_LABELS[status] };
    }
  );
  var priority = GATE1.PRIORITY[inquiry.priority] || {
    label: inquiry.priority,
    color: '#111827'
  };
  var memberLocked =
    actor.role === 'MEMBER' && inquiry.assignee_id !== '';
  var doneLocked = inquiry.status === 'DONE';
  var assignmentNotice = '';
  if (memberLocked) {
    assignmentNotice = '担当者の変更は管理者に依頼してください';
  }

  return {
    actor: { id: String(actor.id), name: actor.name, role: actor.role },
    inquiry: {
      id: String(inquiry.id),
      title: inquiry.title,
      body: inquiry.body,
      requesterName: inquiry.requester_name,
      status: inquiry.status,
      statusLabel: GATE1.STATUS_LABELS[inquiry.status],
      assigneeId:
        inquiry.assignee_id === '' ? '' : String(inquiry.assignee_id),
      assigneeName: assignee ? assignee.name : '未割り当て',
      priorityLabel: priority.label,
      priorityColor: priority.color,
      createdAt: formatDateTime_(inquiry.created_at),
      updatedAtToken: timestampToken_(inquiry.updated_at)
    },
    nextStatuses: nextStatuses,
    showStatusSection: nextStatuses.length > 0,
    activeUsers: activeUsers,
    assignment: {
      allowUnassigned: actor.role === 'ADMIN',
      buttonDisabled: memberLocked || doneLocked,
      notice: assignmentNotice,
      doneLocked: doneLocked,
      suggestedAssigneeId:
        actor.role === 'MEMBER' && inquiry.assignee_id === ''
          ? String(actor.id)
          : ''
    },
    histories: getHistoryView_(inquiry.id)
  };
}

/**
 * ステータスを変更します。
 * The actor and current status are always loaded again on the server.
 */
function changeStatus(payload) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    return failure_(GATE1.MESSAGE.CONCURRENT);
  }

  try {
    payload = payload || {};
    var inquiry = findInquiryById_(payload.inquiryId);
    if (!inquiry) {
      return failure_(GATE1.MESSAGE.NOT_FOUND);
    }
    if (String(payload.comment || '').length > 200) {
      return failure_(GATE1.MESSAGE.COMMENT_TOO_LONG);
    }
    if (
      timestampToken_(inquiry.updated_at) !==
      String(payload.expectedUpdatedAt || '')
    ) {
      return failure_(GATE1.MESSAGE.CONCURRENT);
    }

    var newStatus = String(payload.newStatus || '');
    var allowed = GATE1.TRANSITIONS[inquiry.status] || [];
    if (allowed.indexOf(newStatus) === -1) {
      return failure_(GATE1.MESSAGE.INVALID_TRANSITION);
    }

    var actor = getCurrentActor_();
    var now = new Date();
    var newRow = inquiry.__values.slice();
    var statusIndex = GATE1.INQUIRY_HEADERS.indexOf('status');
    var assigneeIndex = GATE1.INQUIRY_HEADERS.indexOf('assignee_id');
    var updatedIndex = GATE1.INQUIRY_HEADERS.indexOf('updated_at');
    var closedIndex = GATE1.INQUIRY_HEADERS.indexOf('closed_at');
    var historyChanges = [];

    newRow[statusIndex] = newStatus;
    newRow[updatedIndex] = now;
    if (newStatus === 'DONE') {
      newRow[closedIndex] = now;
    }
    if (inquiry.status === 'DONE' && newStatus === 'IN_PROGRESS') {
      newRow[closedIndex] = '';
    }
    if (
      inquiry.status === 'NEW' &&
      newStatus === 'IN_PROGRESS' &&
      inquiry.assignee_id === ''
    ) {
      newRow[assigneeIndex] = actor.id;
      historyChanges.push({
        fieldName: 'assignee',
        oldValue: null,
        newValue: String(actor.id)
      });
    }
    historyChanges.push({
      fieldName: 'status',
      oldValue: inquiry.status,
      newValue: newStatus
    });

    writeInquiryAndHistories_(inquiry, newRow, historyChanges, actor.id, now);
    return success_(
      'ステータスを「' + GATE1.STATUS_LABELS[newStatus] + '」に変更しました。'
    );
  } catch (error) {
    console.error(error);
    if (String(error && error.message) === 'FORCED_HISTORY_FAILURE') {
      return failure_(GATE1.MESSAGE.ROLLBACK);
    }
    return failure_(GATE1.MESSAGE.UNEXPECTED);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Changes the assignee after checking R-01 through R-06 on the server.
 */
function changeAssignee(payload) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    return failure_(GATE1.MESSAGE.CONCURRENT);
  }

  try {
    payload = payload || {};
    var inquiry = findInquiryById_(payload.inquiryId);
    if (!inquiry) {
      return failure_(GATE1.MESSAGE.NOT_FOUND);
    }
    if (
      timestampToken_(inquiry.updated_at) !==
      String(payload.expectedUpdatedAt || '')
    ) {
      return failure_(GATE1.MESSAGE.CONCURRENT);
    }
    if (inquiry.status === 'DONE') {
      return failure_(GATE1.MESSAGE.DONE_ASSIGNEE);
    }

    var actor = getCurrentActor_();
    var rawAssigneeId = payload.newAssigneeId;
    var newAssigneeId =
      rawAssigneeId === '' || rawAssigneeId === null ||
      typeof rawAssigneeId === 'undefined'
        ? null
        : String(rawAssigneeId);
    var targetUser = newAssigneeId === null
      ? null
      : findUserById_(newAssigneeId);

    if (newAssigneeId !== null && (!targetUser || targetUser.is_active !== true)) {
      return failure_(GATE1.MESSAGE.INACTIVE_USER);
    }

    if (actor.role === 'MEMBER') {
      var memberCanAssign =
        inquiry.assignee_id === '' && newAssigneeId === String(actor.id);
      if (!memberCanAssign) {
        return failure_(GATE1.MESSAGE.NO_PERMISSION);
      }
    } else if (actor.role === 'ADMIN') {
      if (newAssigneeId === null && inquiry.status !== 'NEW') {
        return failure_(GATE1.MESSAGE.NO_PERMISSION);
      }
    } else {
      return failure_(GATE1.MESSAGE.NO_PERMISSION);
    }

    var currentAssigneeId =
      inquiry.assignee_id === '' ? null : String(inquiry.assignee_id);
    if (currentAssigneeId === newAssigneeId) {
      if (newAssigneeId === null) {
        return success_('担当者を未割り当てに戻しました。');
      }
      return success_('担当者を「' + targetUser.name + '」に変更しました。');
    }

    var now = new Date();
    var newRow = inquiry.__values.slice();
    var assigneeIndex = GATE1.INQUIRY_HEADERS.indexOf('assignee_id');
    var updatedIndex = GATE1.INQUIRY_HEADERS.indexOf('updated_at');
    newRow[assigneeIndex] = newAssigneeId === null ? '' : targetUser.id;
    newRow[updatedIndex] = now;

    writeInquiryAndHistories_(
      inquiry,
      newRow,
      [
        {
          fieldName: 'assignee',
          oldValue: currentAssigneeId,
          newValue: newAssigneeId
        }
      ],
      actor.id,
      now
    );

    if (newAssigneeId === null) {
      return success_('担当者を未割り当てに戻しました。');
    }
    return success_('担当者を「' + targetUser.name + '」に変更しました。');
  } catch (error) {
    console.error(error);
    if (String(error && error.message) === 'FORCED_HISTORY_FAILURE') {
      return failure_(GATE1.MESSAGE.ROLLBACK);
    }
    return failure_(GATE1.MESSAGE.UNEXPECTED);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Writes an inquiry row and its histories as one rollback unit.
 */
function writeInquiryAndHistories_(
  inquiry,
  newRow,
  changes,
  changedBy,
  changedAt
) {
  var inquiriesSheet = getRequiredSheet_(GATE1.SHEETS.INQUIRIES);
  var historiesSheet = getRequiredSheet_(GATE1.SHEETS.HISTORIES);
  var originalRow = inquiry.__values.slice();
  var historyLastRow = historiesSheet.getLastRow();
  var rowsToWrite = buildHistoryRows_(
    changes,
    inquiry.id,
    changedBy,
    changedAt
  );

  try {
    inquiriesSheet
      .getRange(inquiry.__row, 1, 1, GATE1.INQUIRY_HEADERS.length)
      .setValues([newRow]);

    var props = PropertiesService.getUserProperties();
    if (props.getProperty(GATE1.PROPS.FORCE_HISTORY_FAILURE) === '1') {
      props.deleteProperty(GATE1.PROPS.FORCE_HISTORY_FAILURE);
      throw new Error('FORCED_HISTORY_FAILURE');
    }

    if (rowsToWrite.length > 0) {
      historiesSheet
        .getRange(
          historyLastRow + 1,
          1,
          rowsToWrite.length,
          GATE1.HISTORY_HEADERS.length
        )
        .setValues(rowsToWrite);
    }
    SpreadsheetApp.flush();
  } catch (error) {
    inquiriesSheet
      .getRange(inquiry.__row, 1, 1, GATE1.INQUIRY_HEADERS.length)
      .setValues([originalRow]);

    var appendedCount = Math.max(0, historiesSheet.getLastRow() - historyLastRow);
    if (appendedCount > 0) {
      historiesSheet
        .getRange(
          historyLastRow + 1,
          1,
          appendedCount,
          GATE1.HISTORY_HEADERS.length
        )
        .clearContent();
    }
    SpreadsheetApp.flush();
    throw error;
  }
}

function buildHistoryRows_(changes, inquiryId, changedBy, changedAt) {
  var nextId = getNextHistoryId_();
  return changes.map(function (change, index) {
    return [
      nextId + index,
      inquiryId,
      changedBy,
      change.fieldName,
      change.oldValue === null ? '' : change.oldValue,
      change.newValue === null ? '' : change.newValue,
      changedAt
    ];
  });
}

function getNextHistoryId_() {
  var sheet = getRequiredSheet_(GATE1.SHEETS.HISTORIES);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return 1;
  }
  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var maxId = values.reduce(function (max, row) {
    var value = Number(row[0]);
    return Number.isFinite(value) ? Math.max(max, value) : max;
  }, 0);
  return maxId + 1;
}

function getHistoryView_(inquiryId) {
  var rows = readObjects_(
    getRequiredSheet_(GATE1.SHEETS.HISTORIES),
    GATE1.HISTORY_HEADERS
  );
  return rows
    .filter(function (row) {
      return String(row.inquiry_id) === String(inquiryId);
    })
    .sort(function (a, b) {
      var timeDifference = toDate_(b.changed_at) - toDate_(a.changed_at);
      if (timeDifference !== 0) {
        return timeDifference;
      }
      return Number(b.id) - Number(a.id);
    })
    .slice(0, 20)
    .map(function (row) {
      var changer = findUserById_(row.changed_by);
      return {
        changedAt: formatDateTime_(row.changed_at),
        changedBy: changer ? changer.name : String(row.changed_by),
        label: row.field_name === 'status' ? 'ステータス' : '担当者',
        oldValue: displayHistoryValue_(row.field_name, row.old_value),
        newValue: displayHistoryValue_(row.field_name, row.new_value)
      };
    });
}

function displayHistoryValue_(fieldName, value) {
  if (fieldName === 'status') {
    return GATE1.STATUS_LABELS[String(value)] || String(value);
  }
  if (value === '' || value === null) {
    return '未割り当て';
  }
  var user = findUserById_(value);
  return user ? user.name : String(value);
}

function getCurrentActor_() {
  var userId = PropertiesService.getUserProperties().getProperty(
    GATE1.PROPS.CURRENT_USER_ID
  ) || '1';
  var user = findUserById_(userId);
  if (!user || user.is_active !== true) {
    throw new Error(GATE1.MESSAGE.NO_PERMISSION);
  }
  return user;
}

function getCurrentInquiryId_() {
  return PropertiesService.getUserProperties().getProperty(
    GATE1.PROPS.CURRENT_INQUIRY_ID
  ) || '1001';
}

function getAllUsers_() {
  return readObjects_(
    getRequiredSheet_(GATE1.SHEETS.USERS),
    GATE1.USER_HEADERS
  ).map(function (user) {
    user.is_active = normalizeBoolean_(user.is_active);
    return user;
  });
}

function findUserById_(userId) {
  var users = getAllUsers_();
  for (var i = 0; i < users.length; i += 1) {
    if (String(users[i].id) === String(userId)) {
      return users[i];
    }
  }
  return null;
}

function findInquiryById_(inquiryId) {
  var rows = readObjects_(
    getRequiredSheet_(GATE1.SHEETS.INQUIRIES),
    GATE1.INQUIRY_HEADERS
  );
  for (var i = 0; i < rows.length; i += 1) {
    if (String(rows[i].id) === String(inquiryId)) {
      return rows[i];
    }
  }
  return null;
}

function readObjects_(sheet, expectedHeaders) {
  var values = sheet.getDataRange().getValues();
  if (values.length === 0) {
    return [];
  }
  var actualHeaders = values[0].map(String);
  expectedHeaders.forEach(function (header) {
    if (actualHeaders.indexOf(header) === -1) {
      throw new Error(sheet.getName() + ' にカラム「' + header + '」がありません。');
    }
  });

  return values.slice(1).reduce(function (records, row, index) {
    var hasValue = row.some(function (cell) {
      return cell !== '';
    });
    if (!hasValue) {
      return records;
    }
    var record = {};
    actualHeaders.forEach(function (header, columnIndex) {
      record[header] = row[columnIndex];
    });
    record.__row = index + 2;
    record.__values = row.slice(0, expectedHeaders.length);
    records.push(record);
    return records;
  }, []);
}

function getRequiredSheet_(name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) {
    throw new Error(
      'シート「' + name + '」がありません。setupGate1() を実行してください。'
    );
  }
  return sheet;
}

function normalizeBoolean_(value) {
  return value === true || String(value).toLowerCase() === 'true';
}

function japaneseNameSortKey_(name) {
  return GATE1.NAME_SORT_KEYS[name] || String(name);
}

function timestampToken_(value) {
  var date = toDate_(value);
  return Number.isNaN(date.getTime()) ? String(value) : String(date.getTime());
}

function toDate_(value) {
  return value instanceof Date ? value : new Date(value);
}

function formatDateTime_(value) {
  var date = toDate_(value);
  if (Number.isNaN(date.getTime())) {
    return '';
  }
  var timezone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  return Utilities.formatDate(date, timezone, 'yyyy/MM/dd HH:mm');
}

function success_(message) {
  return { ok: true, message: message };
}

function failure_(message) {
  return { ok: false, message: message };
}

/** 画面本体。別HTMLファイルを作らず、このCode.gsだけで動作します。 */
function buildDialogHtml_(message, type) {
  var initialMessage = JSON.stringify(String(message || ''));
  var initialType = JSON.stringify(String(type || ''));
  var html = `<!DOCTYPE html>
<html lang="ja">
<head>
  <base target="_top">
  <meta charset="UTF-8">
  <style>
    :root {
      --navy: #17324d;
      --blue: #2563eb;
      --line: #dbe3ec;
      --muted: #64748b;
      --surface: #f8fafc;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      background: #eef3f8;
      color: #172033;
      font-family: -apple-system, BlinkMacSystemFont, "Noto Sans JP",
        "Yu Gothic", "Meiryo", sans-serif;
      font-size: 14px;
    }
    .app { max-width: 790px; margin: 0 auto; padding: 18px; }
    .header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      margin-bottom: 12px;
    }
    h1 { margin: 0; color: var(--navy); font-size: 22px; }
    .actor {
      padding: 6px 10px;
      border: 1px solid #bfd0e2;
      border-radius: 999px;
      background: #fff;
      color: #36516d;
      font-size: 12px;
    }
    .message {
      display: none;
      margin-bottom: 12px;
      padding: 11px 14px;
      border: 1px solid;
      border-radius: 8px;
      font-weight: 700;
    }
    .message.success { display: block; color: #166534; background: #ecfdf3; border-color: #86efac; }
    .message.error { display: block; color: #b91c1c; background: #fff1f2; border-color: #fda4af; }
    .card {
      margin-bottom: 12px;
      overflow: hidden;
      border: 1px solid var(--line);
      border-radius: 10px;
      background: #fff;
      box-shadow: 0 2px 8px rgba(23, 50, 77, 0.05);
    }
    .card-title {
      padding: 9px 14px;
      border-bottom: 1px solid var(--line);
      background: var(--surface);
      color: var(--navy);
      font-weight: 800;
    }
    .card-body { padding: 12px 14px; }
    .details {
      display: grid;
      grid-template-columns: 90px 1fr 76px 1fr;
      gap: 7px 12px;
      line-height: 1.65;
    }
    .label { color: var(--muted); font-weight: 700; }
    .body-text { grid-column: 2 / 5; white-space: pre-wrap; }
    .form-row { margin-bottom: 10px; }
    .current {
      display: inline-block;
      min-width: 70px;
      padding: 3px 9px;
      border-radius: 999px;
      background: #e8eef6;
      color: var(--navy);
      text-align: center;
      font-weight: 800;
    }
    .options { display: flex; flex-wrap: wrap; gap: 12px; margin-top: 7px; }
    .option {
      display: inline-flex;
      align-items: center;
      gap: 5px;
      padding: 7px 10px;
      border: 1px solid #cbd5e1;
      border-radius: 8px;
      cursor: pointer;
    }
    textarea, select {
      width: 100%;
      border: 1px solid #cbd5e1;
      border-radius: 7px;
      background: #fff;
      color: #172033;
      font: inherit;
    }
    textarea { min-height: 58px; padding: 8px; resize: vertical; }
    select { height: 38px; padding: 0 9px; }
    .counter { color: var(--muted); text-align: right; font-size: 11px; }
    .actions { display: flex; justify-content: flex-end; }
    button {
      min-width: 150px;
      padding: 9px 14px;
      border: 0;
      border-radius: 7px;
      background: var(--blue);
      color: #fff;
      font-weight: 800;
      cursor: pointer;
    }
    button:disabled { background: #a8b3c2; cursor: not-allowed; }
    .notice { margin: 7px 0 0; color: #b45309; font-weight: 700; }
    .done-note { margin: 7px 0 0; color: #b91c1c; font-weight: 700; }
    .history { padding: 9px 0; border-bottom: 1px solid #edf1f5; }
    .history:last-child { border-bottom: 0; }
    .history-meta { color: var(--muted); font-size: 12px; }
    .history-change { margin-top: 3px; font-weight: 700; }
    .empty { padding: 8px 0; color: var(--muted); }
    .loading { padding: 60px 0; color: var(--muted); text-align: center; }
    @media (max-width: 680px) {
      .details { grid-template-columns: 84px 1fr; }
      .details .label:nth-of-type(n) { grid-column: 1; }
      .body-text { grid-column: 2; }
    }
  </style>
</head>
<body>
  <main class="app">
    <div class="header">
      <h1>問い合わせ詳細</h1>
      <div class="actor" id="actor"></div>
    </div>
    <div id="message" class="message"></div>
    <div id="content" class="loading">読み込み中...</div>
  </main>
  <script>
    var initialMessage = ${initialMessage};
    var initialType = ${initialType};
    var currentViewModel = null;

    function escapeHtml(value) {
      return String(value == null ? '' : value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
    }

    function loadView(message, type) {
      document.getElementById('content').innerHTML =
        '<div class="loading">読み込み中...</div>';
      google.script.run
        .withSuccessHandler(function (vm) {
          currentViewModel = vm;
          render(vm);
          var text = message || initialMessage;
          var kind = type || initialType;
          if (text) {
            showMessage(text, kind || 'success');
          }
          initialMessage = '';
          initialType = '';
        })
        .withFailureHandler(function (error) {
          showMessage(error.message || String(error), 'error');
        })
        .getDetailViewModel();
    }

    function render(vm) {
      document.getElementById('actor').textContent =
        '操作ユーザー: ' + vm.actor.name + ' / ' + vm.actor.role;
      var html = renderDetails(vm);
      if (vm.showStatusSection) {
        html += renderStatus(vm);
      }
      html += renderAssignee(vm);
      html += renderHistories(vm);
      document.getElementById('content').innerHTML = html;
      updateCounter();
    }

    function renderDetails(vm) {
      return '<section class="card">' +
        '<div class="card-title">基本情報</div>' +
        '<div class="card-body details">' +
        '<div class="label">件名</div><div>' + escapeHtml(vm.inquiry.title) + '</div>' +
        '<div class="label">依頼者</div><div>' + escapeHtml(vm.inquiry.requesterName) + '</div>' +
        '<div class="label">優先度</div><div style="color:' +
          escapeHtml(vm.inquiry.priorityColor) + ';font-weight:800">' +
          escapeHtml(vm.inquiry.priorityLabel) + '</div>' +
        '<div class="label">登録日</div><div>' + escapeHtml(vm.inquiry.createdAt) + '</div>' +
        '<div class="label">本文</div><div class="body-text">' +
          escapeHtml(vm.inquiry.body) + '</div>' +
        '</div></section>';
    }

    function renderStatus(vm) {
      var options = vm.nextStatuses.map(function (status, index) {
        return '<label class="option"><input type="radio" name="nextStatus" value="' +
          escapeHtml(status.code) + '"' + (index === 0 ? ' checked' : '') + '>' +
          escapeHtml(status.label) + '</label>';
      }).join('');
      return '<section class="card">' +
        '<div class="card-title">▼ ステータス変更</div>' +
        '<div class="card-body">' +
        '<div class="form-row"><span class="label">現在：</span> ' +
          '<span class="current">' + escapeHtml(vm.inquiry.statusLabel) + '</span></div>' +
        '<div class="form-row"><div class="label">変更先：</div>' +
          '<div class="options">' + options + '</div></div>' +
        '<div class="form-row"><label class="label" for="comment">' +
          'コメント（任意・200 字以内）</label>' +
          '<textarea id="comment" oninput="updateCounter()"></textarea>' +
          '<div class="counter"><span id="commentCount">0</span> / 200</div></div>' +
        '<div class="actions"><button id="statusButton" onclick="submitStatus()">' +
          'ステータスを変更</button></div>' +
        '</div></section>';
    }

    function renderAssignee(vm) {
      var options = '';
      if (vm.assignment.allowUnassigned) {
        options += '<option value="">未割り当て</option>';
      }
      options += vm.activeUsers.map(function (user) {
        var selectedId = vm.inquiry.assigneeId || vm.assignment.suggestedAssigneeId;
        var selected = String(user.id) === String(selectedId) ? ' selected' : '';
        return '<option value="' + escapeHtml(user.id) + '"' + selected + '>' +
          escapeHtml(user.name) + '</option>';
      }).join('');
      var notice = vm.assignment.notice
        ? '<p class="notice">' + escapeHtml(vm.assignment.notice) + '</p>'
        : '';
      var doneNote = vm.assignment.doneLocked
        ? '<p class="done-note">完了済みのため担当者は変更できません。</p>'
        : '';
      return '<section class="card">' +
        '<div class="card-title">▼ 担当者割り当て</div>' +
        '<div class="card-body">' +
        '<div class="form-row"><span class="label">現在：</span> ' +
          escapeHtml(vm.inquiry.assigneeName) + '</div>' +
        '<div class="form-row"><label class="label" for="assignee">変更先：</label>' +
          '<select id="assignee"' + (vm.assignment.buttonDisabled ? ' disabled' : '') + '>' +
          options + '</select>' + notice + doneNote + '</div>' +
        '<div class="actions"><button id="assigneeButton" onclick="submitAssignee()"' +
          (vm.assignment.buttonDisabled
            ? ' disabled data-rule-disabled="true"'
            : '') + '>担当者を変更</button></div>' +
        '</div></section>';
    }

    function renderHistories(vm) {
      var rows = vm.histories.length === 0
        ? '<div class="empty">変更履歴はありません</div>'
        : vm.histories.map(function (history) {
            return '<div class="history">' +
              '<div class="history-meta">' + escapeHtml(history.changedAt) +
              '&nbsp;&nbsp;' + escapeHtml(history.changedBy) + '</div>' +
              '<div class="history-change">' + escapeHtml(history.label) + '： ' +
              escapeHtml(history.oldValue) + ' → ' + escapeHtml(history.newValue) + '</div>' +
              '</div>';
          }).join('');
      return '<section class="card">' +
        '<div class="card-title">▼ 変更履歴（新しい順・最大 20 件）</div>' +
        '<div class="card-body">' + rows + '</div></section>';
    }

    function updateCounter() {
      var comment = document.getElementById('comment');
      var count = document.getElementById('commentCount');
      if (comment && count) {
        count.textContent = comment.value.length;
      }
    }

    function setBusy(isBusy) {
      ['statusButton', 'assigneeButton'].forEach(function (id) {
        var button = document.getElementById(id);
        if (button) {
          button.disabled = isBusy || button.hasAttribute('data-rule-disabled');
        }
      });
    }

    function submitStatus() {
      var selected = document.querySelector('input[name="nextStatus"]:checked');
      if (!selected) {
        return;
      }
      setBusy(true);
      google.script.run
        .withSuccessHandler(handleResult)
        .withFailureHandler(handleFailure)
        .changeStatus({
          inquiryId: currentViewModel.inquiry.id,
          newStatus: selected.value,
          comment: document.getElementById('comment').value,
          expectedUpdatedAt: currentViewModel.inquiry.updatedAtToken
        });
    }

    function submitAssignee() {
      var select = document.getElementById('assignee');
      if (!select) {
        return;
      }
      setBusy(true);
      google.script.run
        .withSuccessHandler(handleResult)
        .withFailureHandler(handleFailure)
        .changeAssignee({
          inquiryId: currentViewModel.inquiry.id,
          newAssigneeId: select.value,
          expectedUpdatedAt: currentViewModel.inquiry.updatedAtToken
        });
    }

    function handleResult(result) {
      loadView(result.message, result.ok ? 'success' : 'error');
    }

    function handleFailure(error) {
      setBusy(false);
      showMessage(error.message || String(error), 'error');
    }

    function showMessage(text, type) {
      var box = document.getElementById('message');
      box.className = 'message ' + (type === 'error' ? 'error' : 'success');
      box.textContent = text;
      box.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    loadView('', '');
  </script>
</body>
</html>`;
  return HtmlService.createHtmlOutput(html);
}
