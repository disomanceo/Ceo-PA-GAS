const CONFIG = {
  SHEET_ID: '1q-y4CmaJW5lRzldfpLIm3yGT2HACYJSSu8dim59RLDs',
  FOLDER_ID: '1_8W_-1Tl4JGdKnaK_gDB814Ux0bt_Gia',
  APP_NAME: 'ระบบการประเมิน PA Online',
  VERSION: '0.2.0',
  DEVELOPER: 'ผอ.สุธน พุทธรัตน์',
  DEVELOPER_POSITION: 'ผู้อำนวยการโรงเรียนวัดไผ่มุ้ง'
};

function doGet() {
  ensureSystem_();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(CONFIG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getBootstrap() {
  ensureSystem_();
  return {
    ok: true,
    appName: CONFIG.APP_NAME,
    version: CONFIG.VERSION,
    sheetId: CONFIG.SHEET_ID,
    folderId: CONFIG.FOLDER_ID,
    fiscalYear: String(new Date().getFullYear() + 543),
    templates: getTemplates_()
  };
}

function saveEvaluation(payload) {
  ensureSystem_();
  validatePayload_(payload);

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const recordSheet = ss.getSheetByName('PA_RECORDS');
  const scoreSheet = ss.getSheetByName('PA_SCORES');
  const pa3Sheet = ss.getSheetByName('PA3');

  const recordId = payload.recordId || Utilities.getUuid();
  const now = new Date();
  const existingRow = findRecordRow_(recordSheet, recordId);

  const recordRow = [
    recordId,
    payload.fiscalYear || '',
    payload.teacher?.fullName || '',
    payload.teacher?.position || 'ครู',
    payload.teacher?.academicRank || '',
    payload.teacher?.school || '',
    payload.teacher?.affiliation || '',
    payload.teacher?.salaryRank || '',
    payload.teacher?.salaryAmount || '',
    payload.templateKey || '',
    payload.status || 'DRAFT',
    Number(payload.summary?.average || 0),
    payload.summary?.passed === true ? 'PASS' : payload.summary?.passed === false ? 'FAIL' : '',
    payload.startDate || '',
    payload.endDate || '',
    payload.evaluationDate || '',
    now
  ];

  if (existingRow > 1) {
    recordSheet.getRange(existingRow, 1, 1, recordRow.length).setValues([recordRow]);
  } else {
    recordSheet.appendRow(recordRow);
  }

  deleteRowsByRecordId_(scoreSheet, recordId);
  const scoreRows = [];
  (payload.committees || []).forEach((committee, committeeIndex) => {
    Object.entries(committee.scores || {}).forEach(([itemId, rating]) => {
      scoreRows.push([
        recordId,
        committeeIndex + 1,
        committee.name || '',
        committee.position || '',
        itemId,
        Number(rating || 0),
        Number((committee.itemScores || {})[itemId] || 0),
        committee.total == null ? '' : Number(committee.total),
        committee.note || '',
        now
      ]);
    });
  });
  if (scoreRows.length) {
    scoreSheet.getRange(scoreSheet.getLastRow() + 1, 1, scoreRows.length, scoreRows[0].length).setValues(scoreRows);
  }

  upsertPa3_(pa3Sheet, recordId, payload, now);
  return { ok: true, recordId, savedAt: now.toISOString() };
}

function exportPa3Pdf(payload) {
  validatePayload_(payload);
  if (!payload.committees.every(c => c.complete)) {
    throw new Error('กรุณาให้คะแนนกรรมการทั้ง 3 คนให้ครบก่อนสร้าง PA3 PDF');
  }

  const saved = saveEvaluation(payload);
  payload.recordId = saved.recordId;

  const fileName = buildPdfFileName_(payload);
  const doc = DocumentApp.create('TMP_' + fileName);
  const body = doc.getBody();

  try {
    body.setPageWidth(595.28);
    body.setPageHeight(841.89);
    body.setMarginTop(28);
    body.setMarginBottom(28);
    body.setMarginLeft(28);
    body.setMarginRight(28);

    appendTopCode_(body, 'PA 3/ส');
    appendCentered_(body, 'แบบสรุปผลการประเมินการพัฒนางานตามข้อตกลง (PA)', 12, true);
    appendCentered_(body, 'สำหรับข้าราชการครูและบุคลากรทางการศึกษา', 12, true);
    appendCentered_(body, 'ตำแหน่ง ' + (payload.teacher.position || 'ครู') + academicRankPdf_(payload.teacher.academicRank), 12, true);
    appendCentered_(body, 'ประจำปีงบประมาณ พ.ศ. ' + (payload.fiscalYear || ''), 12, true);

    const period = 'ระหว่างวันที่ ' + thaiDate_(payload.startDate) + ' ถึงวันที่ ' + thaiDate_(payload.endDate);
    appendCentered_(body, '(' + period + ')', 10, false);

    body.appendParagraph('');
    appendText_(body, 'ข้อมูลผู้รับการประเมิน', 10, true);
    appendText_(body,
      'ชื่อ ' + (payload.teacher.fullName || '') +
      '   ตำแหน่ง ' + (payload.teacher.position || 'ครู') +
      '   วิทยฐานะ ' + (payload.teacher.academicRank || 'ไม่มีวิทยฐานะ'), 10, false);
    appendText_(body,
      'สถานศึกษา ' + (payload.teacher.school || '') +
      (payload.teacher.affiliation ? '   สังกัด ' + payload.teacher.affiliation : ''), 10, false);
    appendText_(body,
      'รับเงินเดือนอันดับ ' + (payload.teacher.salaryRank || '-') +
      '   อัตราเงินเดือน ' + (payload.teacher.salaryAmount || '-') + ' บาท', 10, false);

    body.appendParagraph('');
    appendText_(body, 'ผลการประเมิน', 10, true);

    const c = payload.committees;
    const resultTable = body.appendTable([
      ['การประเมินข้อตกลง\nในการพัฒนางาน', 'คะแนน\nเต็ม', 'คนที่ 1', 'คนที่ 2', 'คนที่ 3', 'หมายเหตุ'],
      ['ส่วนที่ 1 ข้อตกลงในการพัฒนางานตามมาตรฐานตำแหน่ง', '60',
        formatScore_(c[0].part1), formatScore_(c[1].part1), formatScore_(c[2].part1),
        'เกณฑ์ผ่านต้องได้คะแนน\nจากกรรมการแต่ละคน\nไม่ต่ำกว่าร้อยละ 70%'],
      ['ส่วนที่ 2 ข้อตกลงในการพัฒนางานที่เสนอเป็นประเด็นท้าทายในการพัฒนาผลลัพธ์การเรียนรู้ของผู้เรียน', '40',
        formatScore_(c[0].part2), formatScore_(c[1].part2), formatScore_(c[2].part2), ''],
      ['รวม', '100', formatScore_(c[0].total), formatScore_(c[1].total), formatScore_(c[2].total), '']
    ]);
    stylePa3Table_(resultTable);

    body.appendParagraph('');
    const pass = payload.summary?.passed === true;
    const fail = payload.summary?.passed === false;
    appendCentered_(body,
      'สรุปผลการประเมินทั้ง 2 ส่วน จากกรรมการ 3 คน   ' +
      (pass ? '☑' : '☐') + ' ผ่านเกณฑ์   ' +
      (fail ? '☑' : '☐') + ' ไม่ผ่านเกณฑ์', 10, true);

    body.appendParagraph('');
    body.appendParagraph('');

    const evaluationDate = thaiDate_(payload.evaluationDate);
    appendSignatureBlock_(body, c[0], true, evaluationDate);

    body.appendParagraph('');
    const signTable = body.appendTable([['', '']]);
    signTable.setBorderWidth(0);
    appendSignatureCell_(signTable.getCell(0,0), c[1], evaluationDate);
    appendSignatureCell_(signTable.getCell(0,1), c[2], evaluationDate);

    doc.saveAndClose();

    const tmp = DriveApp.getFileById(doc.getId());
    const pdfBlob = tmp.getAs(MimeType.PDF).setName(fileName + '.pdf');
    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    const pdfFile = folder.createFile(pdfBlob);
    try {
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (sharingError) {
      console.log('PDF sharing unchanged: ' + sharingError.message);
    }

    const log = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName('PA3_PDF');
    log.appendRow([
      payload.recordId,
      payload.fiscalYear || '',
      payload.teacher.fullName || '',
      pdfFile.getId(),
      pdfFile.getUrl(),
      pdfFile.getName(),
      new Date()
    ]);

    return {
      ok: true,
      recordId: payload.recordId,
      fileId: pdfFile.getId(),
      fileName: pdfFile.getName(),
      url: pdfFile.getUrl()
    };
  } finally {
    try { DriveApp.getFileById(doc.getId()).setTrashed(true); } catch (e) {}
  }
}

function listEvaluations() {
  ensureSystem_();
  const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName('PA_RECORDS');
  if (sheet.getLastRow() <= 1) return [];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  return values.map(r => ({
    recordId: r[0],
    fiscalYear: r[1],
    teacherName: r[2],
    position: r[3],
    academicRank: r[4],
    school: r[5],
    templateKey: r[9],
    status: r[10],
    average: Number(r[11] || 0),
    result: r[12],
    startDate: r[13],
    endDate: r[14],
    evaluationDate: r[15],
    updatedAt: r[16]
  })).reverse();
}

function ensureSystem_() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  ensureSheet_(ss, 'PA_RECORDS', [
    'record_id','fiscal_year','teacher_name','position','academic_rank','school','affiliation',
    'salary_rank','salary_amount','template_key','status','average_score','result',
    'start_date','end_date','evaluation_date','updated_at'
  ]);
  ensureSheet_(ss, 'PA_SCORES', [
    'record_id','committee_no','committee_name','committee_position','item_id','rating','item_score',
    'committee_total','note','updated_at'
  ]);
  ensureSheet_(ss, 'PA3', [
    'record_id','fiscal_year','teacher_name','academic_rank',
    'committee_1','score_1','committee_2','score_2','committee_3','score_3',
    'average_score','result','updated_at'
  ]);
  ensureSheet_(ss, 'PA3_PDF', [
    'record_id','fiscal_year','teacher_name','file_id','file_url','file_name','created_at'
  ]);
  ensureSheet_(ss, 'SETTINGS', ['key','value']);
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,headers.length).setFontWeight('bold');
    sh.autoResizeColumns(1, headers.length);
    return sh;
  }
  const current = sh.getRange(1,1,1,Math.max(sh.getLastColumn(), headers.length)).getValues()[0];
  headers.forEach((h,i) => {
    if (!current[i]) sh.getRange(1,i+1).setValue(h).setFontWeight('bold');
  });
  return sh;
}

function findRecordRow_(sheet, recordId) {
  if (sheet.getLastRow() <= 1) return -1;
  const ids = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues().flat();
  const idx = ids.findIndex(v => String(v) === String(recordId));
  return idx < 0 ? -1 : idx + 2;
}

function deleteRowsByRecordId_(sheet, recordId) {
  if (sheet.getLastRow() <= 1) return;
  const values = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues().flat();
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i]) === String(recordId)) sheet.deleteRow(i + 2);
  }
}

function upsertPa3_(sheet, recordId, payload, now) {
  const committees = payload.committees || [];
  const row = [
    recordId,
    payload.fiscalYear || '',
    payload.teacher?.fullName || '',
    payload.teacher?.academicRank || '',
    committees[0]?.name || '', committees[0]?.total == null ? '' : Number(committees[0].total),
    committees[1]?.name || '', committees[1]?.total == null ? '' : Number(committees[1].total),
    committees[2]?.name || '', committees[2]?.total == null ? '' : Number(committees[2].total),
    Number(payload.summary?.average || 0),
    payload.summary?.passed === true ? 'PASS' : payload.summary?.passed === false ? 'FAIL' : '',
    now
  ];
  const found = findRecordRow_(sheet, recordId);
  if (found > 1) sheet.getRange(found,1,1,row.length).setValues([row]);
  else sheet.appendRow(row);
}

function validatePayload_(payload) {
  if (!payload || !payload.teacher || !String(payload.teacher.fullName || '').trim()) {
    throw new Error('กรุณาระบุชื่อผู้รับการประเมิน');
  }
  if (!payload.templateKey) throw new Error('ไม่พบแบบประเมิน');
  if (!Array.isArray(payload.committees) || payload.committees.length !== 3) {
    throw new Error('ต้องมีกรรมการ 3 คน');
  }
}

function getTemplates_() {
  return {
    'teacher-none': { label:'ครู ไม่มีวิทยฐานะ', expectedLevel:'ปรับประยุกต์' },
    'teacher-skilled': { label:'ครูชำนาญการ', expectedLevel:'แก้ไขปัญหา' },
    'teacher-senior-skilled': { label:'ครูชำนาญการพิเศษ', expectedLevel:'ริเริ่ม พัฒนา' },
    'teacher-expert': { label:'ครูเชี่ยวชาญ', expectedLevel:'คิดค้น ปรับเปลี่ยน' }
  };
}

function buildPdfFileName_(payload) {
  const safe = String(payload.teacher.fullName || 'PA3').replace(/[\\/:*?"<>|]/g, '_');
  return 'PA3_' + (payload.fiscalYear || '') + '_' + safe;
}

function academicRankPdf_(rank) {
  const v = String(rank || '').trim();
  return ' วิทยฐานะ ' + (v || 'ไม่มีวิทยฐานะ');
}

function formatScore_(value) {
  const n = Number(value || 0);
  return n.toFixed(2).replace(/\.00$/, '').replace(/(\.\d)0$/, '$1');
}

function thaiDate_(value) {
  if (!value) return '................................';
  const d = new Date(value + 'T00:00:00+07:00');
  if (isNaN(d.getTime())) return String(value);
  const months = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
    'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  return d.getDate() + ' ' + months[d.getMonth()] + ' ' + (d.getFullYear() + 543);
}

function appendTopCode_(body, text) {
  const p = body.appendParagraph(text);
  p.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  p.editAsText().setFontFamily('Sarabun').setFontSize(10).setBold(true);
}

function appendCentered_(body, text, size, bold) {
  const p = body.appendParagraph(text);
  p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  p.setSpacingAfter(0);
  p.editAsText().setFontFamily('Sarabun').setFontSize(size).setBold(!!bold);
  return p;
}

function appendText_(body, text, size, bold) {
  const p = body.appendParagraph(text);
  p.setSpacingAfter(0);
  p.editAsText().setFontFamily('Sarabun').setFontSize(size).setBold(!!bold);
  return p;
}

function stylePa3Table_(table) {
  table.setBorderWidth(1);
  const widths = [205,50,45,45,45,105];
  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      try { cell.setWidth(widths[c]); } catch (e) {}
      cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
      const text = cell.editAsText();
      text.setFontFamily('Sarabun').setFontSize(9);
      if (r === 0 || r === 3) text.setBold(true);
      const p = cell.getChild(0).asParagraph();
      if (c >= 1 && c <= 4) p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    }
  }
}

function appendSignatureBlock_(body, committee, centered, evaluationDate) {
  const role = committee.committeeRole || 'ประธานกรรมการผู้ประเมิน';
  const lines = [
    '(ลงชื่อ) ........................................................',
    '(' + (committee.name || '........................................................') + ')',
    'ตำแหน่ง ' + (committee.position || '........................................................'),
    role,
    'วันที่ ' + evaluationDate
  ];
  lines.forEach((line,i) => {
    const p = body.appendParagraph(line);
    p.setAlignment(centered ? DocumentApp.HorizontalAlignment.CENTER : DocumentApp.HorizontalAlignment.LEFT);
    p.setSpacingAfter(0);
    p.editAsText().setFontFamily('Sarabun').setFontSize(9).setBold(i === 1);
  });
}

function appendSignatureCell_(cell, committee, evaluationDate) {
  const role = committee.committeeRole || 'กรรมการผู้ประเมิน';
  cell.clear();
  const lines = [
    '(ลงชื่อ) ........................................................',
    '(' + (committee.name || '........................................................') + ')',
    'ตำแหน่ง ' + (committee.position || '........................................................'),
    role,
    'วันที่ ' + evaluationDate
  ];
  lines.forEach((line,i) => {
    const p = cell.appendParagraph(line);
    p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    p.setSpacingAfter(0);
    p.editAsText().setFontFamily('Sarabun').setFontSize(9).setBold(i === 1);
  });
}
