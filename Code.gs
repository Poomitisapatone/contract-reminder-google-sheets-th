const CFG = {
  DATA_SHEET: 'รายการสัญญา',
  TIMELINE_SHEET: 'ปฏิทิน Timeline',
  SETTINGS_SHEET: 'ตั้งค่า',
  AUDIT_SHEET: 'ประวัติระบบ',
  FIRST_DATA_ROW: 6,
  DEFAULT_DATA_ROWS: 500,
  COL: {
    COMPANY: 1, PHONE: 2, ITEM: 3, EMAILS: 4, SIGNED: 5, DAYS: 6,
    DUE: 7, DELIVERED_DATE: 8, DELIVERY_STATUS: 9, DAYS_LEFT: 10,
    STATUS: 11, CALENDAR_LINK: 12, NOTE: 13, SYS_ID: 14,
    EVENT_ID: 15, EVENT_URL: 16, REM10: 17, REM5: 18, REM3: 19,
    REM0: 20, LAST_REMINDER: 21, SYNC_STATUS: 22, LAST_ERROR: 23,
    DELIVERED_NOTICE: 24
  }
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ระบบสัญญา V2')
    .addItem('ตั้งค่าระบบครั้งแรก', 'setupSystem')
    .addSeparator()
    .addSubMenu(ui.createMenu('เพิ่มแถวพร้อมใช้')
      .addItem('เพิ่ม 50 แถว', 'add50ReadyRows')
      .addItem('เพิ่ม 100 แถว', 'add100ReadyRows'))
    .addSeparator()
    .addItem('ทดสอบส่งอีเมล', 'testEmailNotification')
    .addItem('ตรวจและส่งแจ้งเตือนตอนนี้', 'sendContractReminders')
    .addItem('ซิงก์ Google Calendar ทั้งหมด', 'syncAllCalendarEvents')
    .addItem('รีเซ็ตสถานะแจ้งเตือนทั้งหมด', 'resetAllReminderFlags')
    .addToUi();
}

function setupSystem() {
  const ss = SpreadsheetApp.getActive();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  ss.setSpreadsheetTimeZone(String(getSetting_('เขตเวลา') || 'Asia/Bangkok'));
  createTriggers_();
  initializeExistingRows_();
  ensureTimelineCapacity_(getLastDataRow_());
  syncAllCalendarEvents();
  ss.getSheetByName(CFG.SETTINGS_SHEET).getRange('A16').setValue('ระบบพร้อมใช้งาน: ติดตั้ง Trigger แล้ว กรุณาทดสอบส่งอีเมล').setBackground('#D9EAD3').setFontColor('#274E13');
  logAudit_('ติดตั้งระบบ', '', '', 'สร้าง Trigger และเชื่อม Google Calendar', 'สำเร็จ', '');
  SpreadsheetApp.getUi().alert('ตั้งค่าระบบ V2 เรียบร้อยแล้ว\nกรุณาใช้เมนู “ทดสอบส่งอีเมล” เพื่อตรวจสอบอีกครั้ง');
}

function createTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (['sendContractReminders', 'onContractEdit'].includes(trigger.getHandlerFunction())) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  const hour = Number(getSetting_('ชั่วโมงตรวจแจ้งเตือน')) || 8;
  ScriptApp.newTrigger('sendContractReminders').timeBased().everyDays(1).atHour(hour).create();
  ScriptApp.newTrigger('onContractEdit').forSpreadsheet(getSS_()).onEdit().create();
}

function onContractEdit(e) {
  if (!e || !e.range) return;
  const range = e.range;
  const sh = range.getSheet();
  if (sh.getName() !== CFG.DATA_SHEET || range.getLastRow() < CFG.FIRST_DATA_ROW) return;

  const firstRow = Math.max(range.getRow(), CFG.FIRST_DATA_ROW);
  const lastRow = Math.min(range.getLastRow(), getLastDataRow_());
  const firstCol = range.getColumn();
  const lastCol = range.getLastColumn();
  const deliveryStatusEdited = firstCol <= CFG.COL.DELIVERY_STATUS && lastCol >= CFG.COL.DELIVERY_STATUS;

  for (let row = firstRow; row <= lastRow; row++) {
    try {
      ensureRowFormulas_(sh, row);
      const companyCell = sh.getRange(row, CFG.COL.COMPANY);
      const company = String(companyCell.getDisplayValue() || '').trim();
      const item = String(sh.getRange(row, CFG.COL.ITEM).getDisplayValue() || '').trim();

      if (company && !sh.getRange(row, CFG.COL.SYS_ID).getValue()) {
        sh.getRange(row, CFG.COL.SYS_ID).setValue(Utilities.getUuid());
      }

      let deliveryStatus = '';
      if (deliveryStatusEdited) {
        const statusCell = sh.getRange(row, CFG.COL.DELIVERY_STATUS);
        deliveryStatus = String(statusCell.getValue() || '').trim();
        if (!deliveryStatus) {
          deliveryStatus = 'ยังไม่ส่ง';
          statusCell.setValue(deliveryStatus);
        }
        const deliveredDate = sh.getRange(row, CFG.COL.DELIVERED_DATE);
        if (deliveryStatus === 'ส่งแล้ว' && !deliveredDate.getValue()) {
          deliveredDate.setValue(startOfDay_(new Date()));
        }
        if (deliveryStatus === 'ยังไม่ส่ง' && getSetting_('ล้างวันที่ส่งเมื่อกลับเป็นยังไม่ส่ง') === 'เปิด') {
          deliveredDate.clearContent();
        }
        if (deliveryStatus === 'ยังไม่ส่ง') {
          sh.getRange(row, CFG.COL.DELIVERED_NOTICE).setValue(false);
        }
      }

      if (firstCol <= CFG.COL.DUE && lastCol >= CFG.COL.SIGNED) {
        resetReminderFlagsForRow_(sh, row);
      }

      SpreadsheetApp.flush();
      if (getSetting_('ซิงก์ Calendar ทันที') !== 'ปิด' && firstCol <= CFG.COL.NOTE) {
        try {
          syncCalendarRow_(sh, row);
        } catch (calendarError) {
          sh.getRange(row, CFG.COL.SYNC_STATUS).setValue('ซิงก์ Calendar ไม่สำเร็จ');
          sh.getRange(row, CFG.COL.LAST_ERROR).setValue(String(calendarError.message || calendarError));
          logAudit_('ซิงก์ Calendar', company, item, 'แถว ' + row, 'ไม่สำเร็จ', String(calendarError.message || calendarError));
        }
      }
      if (deliveryStatusEdited && deliveryStatus === 'ส่งแล้ว' && getSetting_('แจ้งเตือนเมื่อเปลี่ยนเป็นส่งแล้ว') !== 'ปิด') {
        sendDeliveredNotification_(sh, row);
      }
      logAudit_('แก้ไขข้อมูล', company, item, 'แถว ' + row + ' คอลัมน์ ' + firstCol + '-' + lastCol, 'สำเร็จ', '');
    } catch (error) {
      sh.getRange(row, CFG.COL.SYNC_STATUS).setValue('เกิดข้อผิดพลาด');
      sh.getRange(row, CFG.COL.LAST_ERROR).setValue(String(error.message || error));
      logAudit_('แก้ไขข้อมูล', sh.getRange(row, CFG.COL.COMPANY).getDisplayValue(), sh.getRange(row, CFG.COL.ITEM).getDisplayValue(), 'แถว ' + row, 'ไม่สำเร็จ', String(error.message || error));
    }
  }
}

function initializeExistingRows_() {
  const sh = getSS_().getSheetByName(CFG.DATA_SHEET);
  const lastDataRow = getLastDataRow_();
  for (let row = CFG.FIRST_DATA_ROW; row <= lastDataRow; row++) {
    ensureRowFormulas_(sh, row);
    const company = String(sh.getRange(row, CFG.COL.COMPANY).getDisplayValue() || '').trim();
    if (company && !sh.getRange(row, CFG.COL.SYS_ID).getValue()) {
      sh.getRange(row, CFG.COL.SYS_ID).setValue(Utilities.getUuid());
    }
    if (!sh.getRange(row, CFG.COL.DELIVERY_STATUS).getValue()) {
      sh.getRange(row, CFG.COL.DELIVERY_STATUS).setValue('ยังไม่ส่ง');
    }
    if (sh.getRange(row, CFG.COL.DELIVERED_NOTICE).getValue() === '') {
      sh.getRange(row, CFG.COL.DELIVERED_NOTICE).setValue(false);
    }
  }
}

function ensureRowFormulas_(sh, row) {
  const formulas = buildRowFormulas_(row);
  sh.getRange(row, CFG.COL.DUE).setFormula(formulas.due);
  sh.getRange(row, CFG.COL.DAYS_LEFT).setFormula(formulas.daysLeft);
  sh.getRange(row, CFG.COL.STATUS).setFormula(formulas.status);
  sh.getRange(row, CFG.COL.CALENDAR_LINK).setFormula(formulas.calendarLink);
}

function buildRowFormulas_(row) {
  return {
    due: '=IF(OR(A' + row + '="",E' + row + '="",F' + row + '=""),"",E' + row + '+F' + row + '-IF(\'ตั้งค่า\'!$B$12="นับรวมวันเซ็น",1,0))',
    daysLeft: '=IF(A' + row + '="","",IF(G' + row + '="","",IF(I' + row + '="ส่งแล้ว",0,G' + row + '-TODAY())))',
    status: '=IF(A' + row + '="","",IF(G' + row + '="","ยังไม่ระบุวัน",IF(I' + row + '="ส่งแล้ว",IF(H' + row + '="","ส่งแล้ว",IF(INT(H' + row + ')<INT(G' + row + '),"ส่งก่อนกำหนด "&(INT(G' + row + ')-INT(H' + row + '))&" วัน",IF(INT(H' + row + ')=INT(G' + row + '),"ส่งตรงกำหนด","ส่งล่าช้า "&(INT(H' + row + ')-INT(G' + row + '))&" วัน"))),IFS(G' + row + '<TODAY(),"เกินกำหนด "&(TODAY()-G' + row + ')&" วัน",G' + row + '=TODAY(),"ครบกำหนดวันนี้",G' + row + '-TODAY()<=3,"เร่งดำเนินการ",G' + row + '-TODAY()<=5,"ต้องติดตาม",G' + row + '-TODAY()<=10,"ใกล้ครบกำหนด",TRUE,"อยู่ระหว่างดำเนินการ"))))',
    calendarLink: '=IF(OR(A' + row + '="",P' + row + '=""),"",HYPERLINK(P' + row + ',"เปิดใน Google Calendar"))'
  };
}

function add50ReadyRows() {
  addReadyRows_(50);
}

function add100ReadyRows() {
  addReadyRows_(100);
}

function addReadyRows_(amount) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const ss = getSS_();
    const sh = ss.getSheetByName(CFG.DATA_SHEET);
    const currentCount = getDataRowCount_();
    const currentLastRow = CFG.FIRST_DATA_ROW + currentCount - 1;
    const firstNewRow = currentLastRow + 1;
    const newCount = currentCount + amount;
    const newLastRow = CFG.FIRST_DATA_ROW + newCount - 1;

    if (sh.getMaxRows() < newLastRow) {
      sh.insertRowsAfter(sh.getMaxRows(), newLastRow - sh.getMaxRows());
    }

    const target = sh.getRange(firstNewRow, 1, amount, CFG.COL.DELIVERED_NOTICE);
    const hasExistingData = target.getDisplayValues().some(function(rowValues) {
      return rowValues.some(function(value) { return value !== ''; });
    });
    if (hasExistingData) {
      throw new Error('พบข้อมูลอยู่ในช่วงแถว ' + firstNewRow + '-' + newLastRow + ' จึงยกเลิกเพื่อป้องกันข้อมูลถูกเขียนทับ');
    }

    const template = sh.getRange(currentLastRow, 1, 1, CFG.COL.DELIVERED_NOTICE);
    template.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    template.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);

    const dueFormulas = [];
    const daysLeftFormulas = [];
    const statusFormulas = [];
    const calendarLinkFormulas = [];
    for (let row = firstNewRow; row <= newLastRow; row++) {
      const formulas = buildRowFormulas_(row);
      dueFormulas.push([formulas.due]);
      daysLeftFormulas.push([formulas.daysLeft]);
      statusFormulas.push([formulas.status]);
      calendarLinkFormulas.push([formulas.calendarLink]);
    }

    sh.getRange(firstNewRow, CFG.COL.DUE, amount, 1).setFormulas(dueFormulas);
    sh.getRange(firstNewRow, CFG.COL.DAYS_LEFT, amount, 1).setFormulas(daysLeftFormulas);
    sh.getRange(firstNewRow, CFG.COL.STATUS, amount, 1).setFormulas(statusFormulas);
    sh.getRange(firstNewRow, CFG.COL.CALENDAR_LINK, amount, 1).setFormulas(calendarLinkFormulas);
    sh.getRange(firstNewRow, CFG.COL.DELIVERY_STATUS, amount, 1).setValue('ยังไม่ส่ง');
    sh.getRange(firstNewRow, CFG.COL.REM10, amount, 4).setValue(false);
    sh.getRange(firstNewRow, CFG.COL.DELIVERED_NOTICE, amount, 1).setValue(false);

    extendConditionalFormatting_(sh, currentLastRow, newLastRow);
    resizeDataFilter_(sh, newLastRow);
    ensureTimelineCapacity_(newLastRow);
    setSetting_('จำนวนแถวสูงสุด', newCount);
    logAudit_('เพิ่มแถวพร้อมใช้', '', '', 'เพิ่ม ' + amount + ' แถว / รวม ' + newCount + ' แถว', 'สำเร็จ', '');
    ss.toast('เพิ่มแถวพร้อมใช้ ' + amount + ' แถวแล้ว รวมทั้งหมด ' + newCount + ' แถว', 'ระบบสัญญา V2', 7);
  } catch (error) {
    logAudit_('เพิ่มแถวพร้อมใช้', '', '', 'เพิ่ม ' + amount + ' แถว', 'ไม่สำเร็จ', String(error.message || error));
    SpreadsheetApp.getUi().alert('เพิ่มแถวไม่สำเร็จ\n' + String(error.message || error));
  } finally {
    lock.releaseLock();
  }
}

function ensureTimelineCapacity_(targetLastRow) {
  const timeline = getSS_().getSheetByName(CFG.TIMELINE_SHEET);
  if (!timeline) return;
  if (timeline.getMaxRows() < targetLastRow) {
    timeline.insertRowsAfter(timeline.getMaxRows(), targetLastRow - timeline.getMaxRows());
  }

  const preparedLastRow = timeline.getLastRow();
  if (preparedLastRow < CFG.FIRST_DATA_ROW || preparedLastRow >= targetLastRow) return;
  const columns = timeline.getMaxColumns();
  const template = timeline.getRange(preparedLastRow, 1, 1, columns);
  const target = timeline.getRange(preparedLastRow + 1, 1, targetLastRow - preparedLastRow, columns);
  template.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  template.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  extendConditionalFormatting_(timeline, preparedLastRow, targetLastRow);
}

function extendConditionalFormatting_(sh, oldLastRow, newLastRow) {
  let changed = false;
  const rules = sh.getConditionalFormatRules().map(function(rule) {
    let ruleChanged = false;
    const ranges = rule.getRanges().map(function(range) {
      if (range.getLastRow() === oldLastRow && range.getRow() <= CFG.FIRST_DATA_ROW) {
        ruleChanged = true;
        changed = true;
        return sh.getRange(range.getRow(), range.getColumn(), newLastRow - range.getRow() + 1, range.getNumColumns());
      }
      return range;
    });
    return ruleChanged ? rule.copy().setRanges(ranges).build() : rule;
  });
  if (changed) sh.setConditionalFormatRules(rules);
}

function resizeDataFilter_(sh, newLastRow) {
  const oldFilter = sh.getFilter();
  const criteria = {};
  if (oldFilter) {
    for (let col = 1; col <= CFG.COL.NOTE; col++) {
      const criterion = oldFilter.getColumnFilterCriteria(col);
      if (criterion) criteria[col] = criterion;
    }
    oldFilter.remove();
  }
  const newFilter = sh.getRange(CFG.FIRST_DATA_ROW - 1, 1, newLastRow - CFG.FIRST_DATA_ROW + 2, CFG.COL.NOTE).createFilter();
  Object.keys(criteria).forEach(function(col) {
    newFilter.setColumnFilterCriteria(Number(col), criteria[col]);
  });
}

function sendContractReminders() {
  const sh = getSS_().getSheetByName(CFG.DATA_SHEET);
  const count = getLastDataRow_() - CFG.FIRST_DATA_ROW + 1;
  const rows = sh.getRange(CFG.FIRST_DATA_ROW, 1, count, CFG.COL.LAST_ERROR).getValues();
  const today = startOfDay_(new Date());
  const thresholds = getReminderDays_().sort(function(a, b) { return a - b; });

  rows.forEach(function(values, index) {
    const row = CFG.FIRST_DATA_ROW + index;
    const company = String(values[CFG.COL.COMPANY - 1] || '').trim();
    const item = String(values[CFG.COL.ITEM - 1] || '').trim();
    const due = values[CFG.COL.DUE - 1];
    const delivered = String(values[CFG.COL.DELIVERY_STATUS - 1] || '') === 'ส่งแล้ว';
    if (!company || !(due instanceof Date) || delivered) return;

    try {
      const left = dateDiff_(startOfDay_(due), today);
      const threshold = thresholds.find(function(day) { return left >= 0 && left <= day; });
      const flagCol = getReminderFlagColumn_(threshold);
      if (flagCol && sh.getRange(row, flagCol).getValue() !== true) {
        const result = sendReminder_(values, left, threshold);
        if (result.ok) {
          sh.getRange(row, flagCol).setValue(true);
          sh.getRange(row, CFG.COL.LAST_REMINDER).setValue(new Date());
          sh.getRange(row, CFG.COL.SYNC_STATUS).setValue('ส่งอีเมลแล้ว');
          sh.getRange(row, CFG.COL.LAST_ERROR).clearContent();
          logAudit_('ส่งอีเมลแจ้งเตือน', company, item, 'เหลือ ' + left + ' วัน / ระดับ ' + threshold + ' วัน', 'สำเร็จ', '');
        } else {
          sh.getRange(row, CFG.COL.SYNC_STATUS).setValue('ส่งอีเมลไม่สำเร็จ');
          sh.getRange(row, CFG.COL.LAST_ERROR).setValue(result.error);
          logAudit_('ส่งอีเมลแจ้งเตือน', company, item, 'เหลือ ' + left + ' วัน', 'ไม่สำเร็จ', result.error);
        }
      }
      syncCalendarRow_(sh, row);
    } catch (error) {
      sh.getRange(row, CFG.COL.SYNC_STATUS).setValue('เกิดข้อผิดพลาด');
      sh.getRange(row, CFG.COL.LAST_ERROR).setValue(String(error.message || error));
      logAudit_('ตรวจแจ้งเตือน', company, item, 'แถว ' + row, 'ไม่สำเร็จ', String(error.message || error));
    }
  });
}

function sendReminder_(values, left, threshold) {
  let recipients = String(values[CFG.COL.EMAILS - 1] || '').trim().replace(/;/g, ',');
  recipients = recipients.split(',').map(function(v) { return v.trim(); }).filter(String).join(',');
  if (!recipients) return { ok: false, error: 'ไม่ได้กรอกอีเมลแจ้งเตือน' };

  const company = values[CFG.COL.COMPANY - 1];
  const phone = values[CFG.COL.PHONE - 1];
  const item = values[CFG.COL.ITEM - 1];
  const due = formatDate_(values[CFG.COL.DUE - 1]);
  const when = left === 0 ? 'ครบกำหนดส่งวันนี้' : 'จะครบกำหนดในอีก ' + left + ' วัน';
  const subject = '[แจ้งเตือนสัญญา] ' + company + ' - ' + when;
  const plain = 'แจ้งเตือนสัญญาจัดซื้อจัดจ้าง\n' + when + '\nบริษัท: ' + company + '\nรายการ: ' + item + '\nเบอร์ติดต่อ: ' + phone + '\nวันที่ครบกำหนด: ' + due;
  const html = '<div style="font-family:Arial,sans-serif;max-width:640px">' +
    `<h2 style="color:#17365D">แจ้งเตือนสัญญาจัดซื้อจัดจ้าง</h2>` +
    `<p style="font-size:16px"><b>` + escapeHtml_(when) + `</b></p>` +
    `<table style="border-collapse:collapse;width:100%">` +
    rowHtml_('บริษัท', company) + rowHtml_('รายการ', item) + rowHtml_('เบอร์ติดต่อ', phone) + rowHtml_('วันที่ครบกำหนด', due) +
    `</table><p>กรุณาตรวจสอบและดำเนินการตามกำหนด</p>` +
    `<p style="color:#777;font-size:12px">ระดับการแจ้งเตือน ` + threshold + ` วัน</p></div>`;
  try {
    MailApp.sendEmail({ to: recipients, subject: subject, body: plain, htmlBody: html, name: String(getSetting_('ชื่อผู้ส่ง') || 'ระบบแจ้งเตือนสัญญา') });
    return { ok: true, error: '' };
  } catch (error) {
    return { ok: false, error: String(error.message || error) };
  }
}

function sendDeliveredNotification_(sh, row) {
  const noticeCell = sh.getRange(row, CFG.COL.DELIVERED_NOTICE);
  if (noticeCell.getValue() === true) return { ok: true, skipped: true, error: '' };

  SpreadsheetApp.flush();
  const values = sh.getRange(row, 1, 1, CFG.COL.DELIVERED_NOTICE).getValues()[0];
  const company = String(values[CFG.COL.COMPANY - 1] || '').trim();
  const item = String(values[CFG.COL.ITEM - 1] || '').trim();
  if (!company && !item) return { ok: false, skipped: true, error: 'ยังไม่ได้กรอกข้อมูลสัญญา' };

  let recipients = String(values[CFG.COL.EMAILS - 1] || '').trim().replace(/;/g, ',');
  recipients = recipients.split(',').map(function(v) { return v.trim(); }).filter(String).join(',');
  if (!recipients) {
    const missingEmail = 'ไม่ได้กรอกอีเมลแจ้งเตือน';
    sh.getRange(row, CFG.COL.SYNC_STATUS).setValue('แจ้งส่งแล้วไม่สำเร็จ');
    sh.getRange(row, CFG.COL.LAST_ERROR).setValue(missingEmail);
    logAudit_('แจ้งเปลี่ยนเป็นส่งแล้ว', company, item, 'แถว ' + row, 'ไม่สำเร็จ', missingEmail);
    return { ok: false, skipped: false, error: missingEmail };
  }

  const phone = values[CFG.COL.PHONE - 1];
  const due = formatDate_(values[CFG.COL.DUE - 1]);
  const deliveredDate = formatDate_(values[CFG.COL.DELIVERED_DATE - 1]);
  const automaticStatus = sh.getRange(row, CFG.COL.STATUS).getDisplayValue() || 'ส่งแล้ว';
  const subject = '[แจ้งส่งมอบแล้ว] ' + company + ' - ' + item;
  const plain = 'ยืนยันการส่งมอบสัญญาจัดซื้อจัดจ้าง\nบริษัท: ' + company + '\nรายการ: ' + item + '\nเบอร์ติดต่อ: ' + phone + '\nวันที่ครบกำหนด: ' + due + '\nวันที่ส่งจริง: ' + deliveredDate + '\nสถานะ: ' + automaticStatus;
  const html = '<div style="font-family:Arial,sans-serif;max-width:640px">' +
    `<h2 style="color:#274E13">ยืนยันการส่งมอบเรียบร้อยแล้ว</h2>` +
    `<table style="border-collapse:collapse;width:100%">` +
    rowHtml_('บริษัท', company) + rowHtml_('รายการ', item) + rowHtml_('เบอร์ติดต่อ', phone) +
    rowHtml_('วันที่ครบกำหนด', due) + rowHtml_('วันที่ส่งจริง', deliveredDate) + rowHtml_('สถานะ', automaticStatus) +
    `</table><p style="color:#777;font-size:12px">อีเมลนี้ส่งอัตโนมัติเมื่อเปลี่ยนช่อง “ส่งแล้ว” เป็น “ส่งแล้ว”</p></div>`;

  try {
    MailApp.sendEmail({
      to: recipients,
      subject: subject,
      body: plain,
      htmlBody: html,
      name: String(getSetting_('ชื่อผู้ส่ง') || 'ระบบแจ้งเตือนสัญญา')
    });
    noticeCell.setValue(true);
    sh.getRange(row, CFG.COL.SYNC_STATUS).setValue('ส่งอีเมลยืนยันแล้ว');
    sh.getRange(row, CFG.COL.LAST_ERROR).clearContent();
    logAudit_('แจ้งเปลี่ยนเป็นส่งแล้ว', company, item, 'ส่งไปที่ ' + recipients, 'สำเร็จ', '');
    return { ok: true, skipped: false, error: '' };
  } catch (error) {
    const message = String(error.message || error);
    sh.getRange(row, CFG.COL.SYNC_STATUS).setValue('แจ้งส่งแล้วไม่สำเร็จ');
    sh.getRange(row, CFG.COL.LAST_ERROR).setValue(message);
    logAudit_('แจ้งเปลี่ยนเป็นส่งแล้ว', company, item, 'ส่งไปที่ ' + recipients, 'ไม่สำเร็จ', message);
    return { ok: false, skipped: false, error: message };
  }
}

function testEmailNotification() {
  const to = String(getSetting_('อีเมลทดสอบ') || Session.getEffectiveUser().getEmail() || '').trim();
  if (!to) {
    SpreadsheetApp.getUi().alert('กรุณากรอก “อีเมลทดสอบ” ในแท็บตั้งค่า');
    return;
  }
  try {
    const now = Utilities.formatDate(new Date(), String(getSetting_('เขตเวลา') || 'Asia/Bangkok'), 'dd/MM/yyyy HH:mm:ss');
    MailApp.sendEmail({
      to: to,
      subject: '[ทดสอบระบบ] ระบบแจ้งเตือนสัญญาทำงานปกติ',
      body: 'ทดสอบระบบสำเร็จ เวลา ' + now,
      htmlBody: '<div style="font-family:Arial,sans-serif"><h2 style="color:#17365D">ทดสอบระบบสำเร็จ</h2><p>ระบบสามารถส่งอีเมลแจ้งเตือนได้ตามปกติ</p><p>เวลาทดสอบ: ' + now + '</p></div>',
      name: String(getSetting_('ชื่อผู้ส่ง') || 'ระบบแจ้งเตือนสัญญา')
    });
    logAudit_('ทดสอบอีเมล', '', '', 'ส่งไปที่ ' + to, 'สำเร็จ', '');
    SpreadsheetApp.getUi().alert('ส่งอีเมลทดสอบไปที่ ' + to + ' เรียบร้อยแล้ว');
  } catch (error) {
    logAudit_('ทดสอบอีเมล', '', '', 'ส่งไปที่ ' + to, 'ไม่สำเร็จ', String(error.message || error));
    SpreadsheetApp.getUi().alert('ส่งอีเมลทดสอบไม่สำเร็จ\n' + String(error.message || error));
  }
}

function syncAllCalendarEvents() {
  const sh = getSS_().getSheetByName(CFG.DATA_SHEET);
  const lastDataRow = getLastDataRow_();
  for (let row = CFG.FIRST_DATA_ROW; row <= lastDataRow; row++) {
    const company = String(sh.getRange(row, CFG.COL.COMPANY).getDisplayValue() || '').trim();
    const eventId = String(sh.getRange(row, CFG.COL.EVENT_ID).getValue() || '').trim();
    if (!company && !eventId) continue;
    try {
      syncCalendarRow_(sh, row);
    } catch (error) {
      sh.getRange(row, CFG.COL.SYNC_STATUS).setValue('ซิงก์ไม่สำเร็จ');
      sh.getRange(row, CFG.COL.LAST_ERROR).setValue(String(error.message || error));
      logAudit_('ซิงก์ Calendar', company, sh.getRange(row, CFG.COL.ITEM).getDisplayValue(), 'แถว ' + row, 'ไม่สำเร็จ', String(error.message || error));
    }
  }
  getSS_().toast('ซิงก์ Google Calendar เสร็จแล้ว', 'ระบบสัญญา V2', 5);
}

function syncCalendarRow_(sh, row) {
  SpreadsheetApp.flush();
  const values = sh.getRange(row, 1, 1, CFG.COL.LAST_ERROR).getValues()[0];
  const company = String(values[CFG.COL.COMPANY - 1] || '').trim();
  const item = String(values[CFG.COL.ITEM - 1] || '').trim();
  const due = values[CFG.COL.DUE - 1];
  const calendarId = String(getSetting_('Calendar ID') || 'primary').trim();
  const cal = calendarId === 'primary' ? CalendarApp.getDefaultCalendar() : CalendarApp.getCalendarById(calendarId);
  if (!cal) throw new Error('ไม่พบ Google Calendar: ' + calendarId);

  let eventId = String(values[CFG.COL.EVENT_ID - 1] || '').trim();
  let event = eventId ? cal.getEventById(eventId) : null;
  if (!company || !item || !(due instanceof Date)) {
    if (event) event.deleteEvent();
    sh.getRange(row, CFG.COL.EVENT_ID, 1, 2).clearContent();
    sh.getRange(row, CFG.COL.SYNC_STATUS).setValue(eventId ? 'ลบกิจกรรมแล้ว' : 'รอข้อมูลให้ครบ');
    sh.getRange(row, CFG.COL.LAST_ERROR).clearContent();
    return;
  }

  const delivered = String(values[CFG.COL.DELIVERY_STATUS - 1] || '') === 'ส่งแล้ว';
  const baseTitle = company + ' - ' + item;
  const title = delivered ? '[ส่งแล้ว] ' + baseTitle : '[ครบกำหนด] ' + baseTitle;
  const desc = 'บริษัท: ' + company + '\nเบอร์ติดต่อ: ' + values[CFG.COL.PHONE - 1] + '\nรายการ: ' + item + '\nวันที่เซ็น: ' + formatDate_(values[CFG.COL.SIGNED - 1]) + '\nวันครบกำหนด: ' + formatDate_(due) + '\nสถานะ: ' + sh.getRange(row, CFG.COL.STATUS).getDisplayValue();

  if (event) {
    event.setTitle(title).setAllDayDate(due).setDescription(desc);
  } else {
    event = cal.createAllDayEvent(title, due, { description: desc });
    eventId = event.getId();
    sh.getRange(row, CFG.COL.EVENT_ID).setValue(eventId);
  }

  event.removeAllReminders();
  if (!delivered) {
    getReminderDays_().forEach(function(day) {
      if (day > 0) event.addEmailReminder(day * 24 * 60);
      if (day === 0) event.addPopupReminder(0);
    });
  }
  const link = 'https://calendar.google.com/calendar/u/0/r/search?q=' + encodeURIComponent(title);
  sh.getRange(row, CFG.COL.EVENT_URL).setValue(link);
  sh.getRange(row, CFG.COL.SYNC_STATUS).setValue(delivered ? 'ส่งแล้ว / ปิดแจ้งเตือน' : 'ซิงก์แล้ว');
  sh.getRange(row, CFG.COL.LAST_ERROR).clearContent();
}

function resetAllReminderFlags() {
  const sh = getSS_().getSheetByName(CFG.DATA_SHEET);
  const count = getLastDataRow_() - CFG.FIRST_DATA_ROW + 1;
  sh.getRange(CFG.FIRST_DATA_ROW, CFG.COL.REM10, count, 4).setValue(false);
  sh.getRange(CFG.FIRST_DATA_ROW, CFG.COL.LAST_REMINDER, count, 1).clearContent();
  logAudit_('รีเซ็ตการแจ้งเตือน', '', '', 'ทุกสัญญา', 'สำเร็จ', '');
  getSS_().toast('รีเซ็ตสถานะแจ้งเตือนทั้งหมดแล้ว', 'ระบบสัญญา V2', 5);
}

function resetReminderFlagsForRow_(sh, row) {
  sh.getRange(row, CFG.COL.REM10, 1, 4).setValue(false);
  sh.getRange(row, CFG.COL.LAST_REMINDER).clearContent();
}

function getReminderFlagColumn_(day) {
  const map = { 10: CFG.COL.REM10, 5: CFG.COL.REM5, 3: CFG.COL.REM3, 0: CFG.COL.REM0 };
  return Object.prototype.hasOwnProperty.call(map, day) ? map[day] : null;
}

function getReminderDays_() {
  const supported = [10, 5, 3, 0];
  const raw = String(getSetting_('วันแจ้งเตือน') || '10,5,3,0');
  const parsed = raw.split(',').map(function(v) { return Number(v.trim()); }).filter(function(v) { return supported.includes(v); });
  return parsed.length ? Array.from(new Set(parsed)) : supported;
}

function getSetting_(name) {
  const sh = getSS_().getSheetByName(CFG.SETTINGS_SHEET);
  const values = sh.getRange('A5:B30').getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === name) return values[i][1];
  }
  return '';
}

function setSetting_(name, value) {
  const sh = getSS_().getSheetByName(CFG.SETTINGS_SHEET);
  const values = sh.getRange('A5:A30').getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === name) {
      sh.getRange(i + 5, 2).setValue(value);
      return;
    }
  }
  throw new Error('ไม่พบการตั้งค่า: ' + name);
}

function getDataRowCount_() {
  const configured = Math.floor(Number(getSetting_('จำนวนแถวสูงสุด')));
  return configured > 0 ? configured : CFG.DEFAULT_DATA_ROWS;
}

function getLastDataRow_() {
  return CFG.FIRST_DATA_ROW + getDataRowCount_() - 1;
}

function getSS_() {
  const storedId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (storedId) return SpreadsheetApp.openById(storedId);
  return SpreadsheetApp.getActive();
}

function logAudit_(action, company, item, detail, result, error) {
  const sh = getSS_().getSheetByName(CFG.AUDIT_SHEET);
  if (!sh) return;
  sh.appendRow([new Date(), action, company || '', item || '', detail || '', result || '', error || '']);
}

function rowHtml_(label, value) {
  return '<tr><td style="border:1px solid #ddd;padding:8px;background:#f5f7fa"><b>' + escapeHtml_(label) + '</b></td><td style="border:1px solid #ddd;padding:8px">' + escapeHtml_(value) + '</td></tr>';
}

function formatDate_(value) {
  if (!(value instanceof Date)) return '';
  return Utilities.formatDate(value, String(getSetting_('เขตเวลา') || 'Asia/Bangkok'), 'dd/MM/yyyy');
}

function startOfDay_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function dateDiff_(future, present) {
  return Math.round((future - present) / 86400000);
}

function escapeHtml_(value) {
  return String(value == null ? '' : value).replace(/[&<>"']/g, function(char) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char];
  });
}
