/**
 * 선생님 차량 케어 — 신청(index.html) + 추천서(recommend.html) 저장 서버
 * 사용법: 구글 시트의 [확장 프로그램 > Apps Script]에 붙여넣기
 * ※ 코드를 고친 뒤에는 [배포 > 배포 관리 > 수정 > 새 버전]으로 다시 배포해야 반영됩니다.
 */

const SHEET_NAME = '신청';
const REC_SHEET = '추천서';
const NOTIFY_EMAIL = true;   // 신청·추천서 제출 시 내 메일로 알림 (끄려면 false)

const HEADERS = [
  '접수번호', '접수일시', '진행상태', '담당학생',
  '성함', '연락처',
  '제조사', '차종', '연식', '연료', '엔진·트림', '주행거리(km)', '차대번호',
  '요청항목', '타이어 규격', '타이어 메모', '최근 교체 이력', '증상·요청사항', '비고'
];
const REC_HEADERS = [
  '접수번호', '최종 저장', '조사 학생', '차량', '항목 수', '판단 요약',
  '종합 의견', '검토 상태', '지도교사 의견', '원본 데이터(수정 금지)'
];
const STATUS = ['접수', '학생 배정', '추천서 작성중', '검토중', '전달 완료'];

/* ================= 요청 처리 ================= */

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.type === 'rec') saveRec_(d);
    else saveRequest_(d);
    return json_({ ok: true, id: d.id });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  if (e && e.parameter && e.parameter.list) return json_({ ok: true, list: listRequests_() });
  const id = e && e.parameter && String(e.parameter.id || '').trim().toUpperCase();
  if (!id) return ContentService.createTextOutput('차량 케어 서버 작동 중');
  const request = findRequest_(id);
  const rec = findRec_(id);
  return json_({ ok: !!(request || rec), request: request, rec: rec });
}

/** 처음 한 번 실행해서 시트와 권한을 준비 */
function setup() {
  getSheet_();
  getRecSheet_();
  SpreadsheetApp.getActive().toast('신청·추천서 시트 준비 완료');
}

/* ================= 신청 ================= */

function saveRequest_(d) {
  const s = v => safe_(v || '');
  getSheet_().appendRow([
    s(d.id), new Date(), '접수', '',
    s(d.name), s(d.contact),
    s(d.maker), s(d.model), s(d.year), s(d.fuel), s(d.trim),
    d.km ? Number(d.km) : '', s(d.vin),
    s(d.items), s(d.tireSize), s(d.tireNote), s(d.history), s(d.memo), ''
  ]);
  if (NOTIFY_EMAIL) {
    mail_(`[차량 케어] 새 신청 ${d.id} · ${d.model}`, [
      `접수번호: ${d.id}`,
      `신청자: ${d.name}`,
      `차량: ${d.maker} ${d.model} (${d.year}, ${d.fuel}) ${d.trim || ''}`,
      `주행거리: ${d.km ? Number(d.km).toLocaleString() : '-'} km`,
      `요청항목: ${d.items}`,
      d.tireSize ? `타이어 규격: ${d.tireSize}` : '',
      d.memo ? `요청사항: ${d.memo}` : ''
    ]);
  }
}

/** 접수 목록 (최신순, 성함·연락처 제외) */
function listRequests_() {
  const rows = getSheet_().getDataRange().getValues();
  const c = name => HEADERS.indexOf(name);
  const recRows = getRecSheet_().getDataRange().getValues();
  const recStatus = {};
  for (let i = 1; i < recRows.length; i++) {
    recStatus[String(recRows[i][0]).toUpperCase()] = recRows[i][REC_HEADERS.indexOf('검토 상태')];
  }
  const out = [];
  for (let i = rows.length - 1; i >= 1; i--) {
    const r = rows[i];
    const id = String(r[0]).toUpperCase();
    if (!id) continue;
    const d = r[c('접수일시')];
    out.push({
      id: id,
      date: d instanceof Date ? Utilities.formatDate(d, 'Asia/Seoul', 'MM/dd') : '',
      car: [r[c('제조사')], r[c('차종')], r[c('연식')]].filter(String).join(' '),
      items: String(r[c('요청항목')] || ''),
      status: r[c('진행상태')] || '접수',
      rec: recStatus[id] || ''
    });
  }
  return out;
}

/** 학생에게 보여줄 정보만 반환 (성함·연락처 제외) */
function findRequest_(id) {
  const sh = getSheet_();
  const rows = sh.getDataRange().getValues();
  const c = name => HEADERS.indexOf(name);
  for (let i = rows.length - 1; i >= 1; i--) {
    const r = rows[i];
    if (String(r[0]).toUpperCase() !== id) continue;
    return {
      id: id,
      maker: r[c('제조사')], model: r[c('차종')], year: String(r[c('연식')]),
      fuel: r[c('연료')], trim: r[c('엔진·트림')], km: r[c('주행거리(km)')],
      vin: r[c('차대번호')], items: r[c('요청항목')],
      tireSize: r[c('타이어 규격')], tireNote: r[c('타이어 메모')],
      history: r[c('최근 교체 이력')], memo: r[c('증상·요청사항')]
    };
  }
  return null;
}

/* ================= 추천서 ================= */

function saveRec_(d) {
  const sh = getRecSheet_();
  const id = String(d.id || '').toUpperCase();
  if (!id) throw new Error('접수번호 없음');

  const v = d.vehicle || {};
  const items = d.items || [];
  const review = d.review || {};
  const row = [
    id, new Date(), safe_(d.students || ''),
    safe_([v.maker, v.model, v.year].filter(Boolean).join(' ')),
    items.length,
    safe_(items.map(i => `${i.item}: ${i.decision}`).join(' / ')),
    safe_(d.summary || ''),
    d.mode === 'review' ? (review.status || '검토중') : '검토 대기',
    safe_(review.comment || ''),
    JSON.stringify(d)
  ];

  const ids = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues().map(r => String(r[0]).toUpperCase());
  const at = ids.indexOf(id);
  if (at >= 1) sh.getRange(at + 1, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);

  updateRequestStatus_(id, d);

  if (NOTIFY_EMAIL && d.mode !== 'review') {
    mail_(`[차량 케어] 추천서 제출 ${id} · ${v.model || ''}`, [
      `접수번호: ${id}`,
      `조사 학생: ${d.students}`,
      `차량: ${[v.maker, v.model, v.year].filter(Boolean).join(' ')}`,
      `항목: ${items.map(i => `${i.item}(${i.decision})`).join(', ')}`,
      '',
      '검토 페이지: recommend.html?review 주소로 열어 접수번호를 불러오세요.'
    ]);
  }
}

function findRec_(id) {
  const sh = getRecSheet_();
  const rows = sh.getDataRange().getValues();
  const c = REC_HEADERS.indexOf('원본 데이터(수정 금지)');
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toUpperCase() !== id) continue;
    try {
      const d = JSON.parse(rows[i][c]);
      // 시트에서 직접 고친 검토 상태·의견을 우선 반영
      d.review = {
        status: rows[i][REC_HEADERS.indexOf('검토 상태')] || (d.review && d.review.status) || '검토중',
        comment: rows[i][REC_HEADERS.indexOf('지도교사 의견')] || (d.review && d.review.comment) || ''
      };
      return d;
    } catch (e) { return null; }
  }
  return null;
}

/** 신청 시트의 담당학생·진행상태 자동 갱신 */
function updateRequestStatus_(id, d) {
  const sh = getSheet_();
  const ids = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues().map(r => String(r[0]).toUpperCase());
  const at = ids.indexOf(id);
  if (at < 1) return;
  const r = at + 1;
  if (d.students) sh.getRange(r, 4).setValue(safe_(d.students));
  const cur = sh.getRange(r, 3).getValue();
  if (d.mode !== 'review' && ['접수', '학생 배정', '추천서 작성중'].indexOf(cur) >= 0) {
    sh.getRange(r, 3).setValue('검토중');
  }
}

/* ================= 시트 준비 ================= */

function getSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    styleHeader_(sh, HEADERS);
    sh.getRange('B:B').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange('L:L').setNumberFormat('#,##0');
    sh.getRange('M:M').setNumberFormat('@');
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUS, true).build();
    sh.getRange(2, 3, 999, 1).setDataValidation(rule);
  }
  return sh;
}

function getRecSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(REC_SHEET);
  if (!sh) {
    sh = ss.insertSheet(REC_SHEET);
    styleHeader_(sh, REC_HEADERS);
    sh.getRange('B:B').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.setColumnWidth(REC_HEADERS.length, 80);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['검토 대기', '검토중', '수정 요청', '검토 완료'], true).build();
    sh.getRange(2, REC_HEADERS.indexOf('검토 상태') + 1, 999, 1).setDataValidation(rule);
  }
  return sh;
}

function styleHeader_(sh, headers) {
  sh.appendRow(headers);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
}

/* ================= 유틸 ================= */

function mail_(subject, lines) {
  const me = Session.getEffectiveUser().getEmail();
  if (!me) return;
  MailApp.sendEmail(me, subject, lines.filter(Boolean).concat(['', SpreadsheetApp.getActive().getUrl()]).join('\n'));
}

// 수식 주입 방지 (=, +, -, @로 시작하는 값)
function safe_(v) {
  v = String(v);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
