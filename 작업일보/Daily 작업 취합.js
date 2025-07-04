// 2. 폼 팝업 띄우기
function showFormDialog() {
  const html = HtmlService
    .createHtmlOutputFromFile('SheetSelectionForm')
    .setWidth(400)
    .setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, "시트 및 날짜 선택");
}

// 3. 스프레드시트 내 시트 목록 가져오기
function getSheetNames() {
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .map(s => s.getName());
}

// 4. 시트 생성 및 데이터 복사 실행
function createSheetFromTemplateAndSource(templateName, sourceName, dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 템플릿·소스 시트 존재 확인
  const templateSheet = ss.getSheetByName(templateName);
  if (!templateSheet) {
    ui.alert('템플릿 오류', `"${templateName}" 시트를 찾을 수 없습니다.`, ui.ButtonSet.OK);
    return;
  }
  const sourceSheet = ss.getSheetByName(sourceName);
  if (!sourceSheet) {
    ui.alert('데이터 시트 오류', `"${sourceName}" 시트를 찾을 수 없습니다.`, ui.ButtonSet.OK);
    return;
  }

  // ─── 1) 사전 유효성 검사 ───
  const data    = sourceSheet.getRange("A3:N18").getValues();
  const headers = sourceSheet.getRange("A2:N2").getValues()[0];

  // A열(0)부터 N열(13)까지 검사 대상으로 등록
  const required = [];
  for (let col = 0; col <= 13; col++) {
    required.push({ idx: col, name: headers[col] });
  }

  const issues = [];
  data.forEach((row, i) => {
    const dateValue = row[0];     // A열: 작업일시
    const taskCount = row[1];     // B열: 작업수

    // — B열이 비어 있으면(작업수 미기입) 검증·복사 대상에서 완전히 제외
    if (taskCount === '' || taskCount == null) return;

    // B열이 채워졌으니, A~N 중 하나라도 비어있다면 누락으로 처리
    const missing = required
      .filter(c => row[c.idx] === '' || row[c.idx] == null)
      .map(c => c.name);

    if (missing.length) {
      issues.push({
        taskCount: taskCount,
        missingHeaders: missing
      });
    }
  });

  if (issues.length) {
    // 누락 항목을 번호 붙여 한 번에 보여주고 중단
    const msg = issues.map((itm, idx) =>
      `(${idx + 1})\n` +
      `확인 필요 : ${itm.taskCount}번 작업\n` +
      `수정 및 기입 필요 : ${itm.missingHeaders.join(', ')}`
    ).join('\n\n');

    ui.alert('취합 불가', msg, ui.ButtonSet.OK);
    return;  // 여기서 끝내면 시트 복사 안 함
  }

  // ─── 2) 검사 통과 시 시트 복사 ───
  const sheetName = `(작업)${dateStr}`;
  const existing  = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);

  const newSheet = templateSheet.copyTo(ss).setName(sheetName);
  ss.setActiveSheet(newSheet);

  // ─── 3) 데이터 집계 로직 호출 ───
  copyToTargetSheet(sourceName, sheetName);
}


// 5. 데이터 취합 로직 (사전 검증된 데이터만 넘어옴)
function copyToTargetSheet(sourceSheetName, targetSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  const targetSheet = ss.getSheetByName(targetSheetName);
  const data = sourceSheet.getRange("A3:N18").getValues();

  // 층수 기준 열 맵 (A2, C2, E2...)
  const floorCols = {};
  const secondRow = targetSheet.getRange(2, 1, 1, targetSheet.getLastColumn()).getValues()[0];
  for (let col = 0; col < secondRow.length; col += 2) {
    const floor = secondRow[col];
    if (floor) floorCols[floor] = col + 2;
  }

  const processedKeys = new Set();

  data.forEach(row => {
    const taskCount = row[1];
    // B열이 비어 있으면 스킵
    if (taskCount === '' || taskCount == null) return;

    const [ , , bizP, , , guBun, workName, guGan, floor ] = row;
    const col = floorCols[floor];
    if (!col) return;

    const groupKey = `${floor}_${guBun}_${guGan}_${bizP}`;
    if (processedKeys.has(groupKey)) return;
    processedKeys.add(groupKey);

    const groupRows = data.filter(r =>
      r[2] === bizP && r[5] === guBun &&
      r[7] === guGan && r[8] === floor
    );

    // 집계
    const workNames = new Set();
    const locations = new Set();
    const contents  = new Set();
    const hazards   = new Set();
    const safeties  = new Set();

    groupRows.forEach(r => {
      if (r[6]) workNames.add(r[6]);
      const loc = [r[7], r[8], r[9]].filter(Boolean).join(' ');
      if (loc) locations.add(loc);
      if (r[10]) contents.add(r[10]);
      if (r[12]) hazards.add(r[12]);
      if (r[13]) safeties.add(r[13]);
    });

    const mergedWorkName  = mergeWorkNames(workNames);
    const displayWorkName = `(${guBun}) 배관 ${mergedWorkName}`;

    const values = [
      displayWorkName,
      formatAsLineList(locations),
      formatAsLineList(contents),
      bizP,
      formatAsLineList(hazards),
      formatAsLineList(safeties)
    ];

    // 사내→3행부터, 사외→9행부터
    const startRow = (guBun === '사내') ? 3 : 9;
    values.forEach((v, i) => {
      targetSheet.getRange(startRow + i, col).setValue(v);
    });
  });
}

// 6. 작업명 병합 (중복 시 '설치 및 전환작업'으로 통일)
function mergeWorkNames(set) {
  const items = [...set];
  if (items.length > 1 || items[0] === '설치 및 전환작업') {
    return '설치 및 전환작업';
  }
  return items[0] || '';
}

// 7. 줄바꿈 정리 (기호 없음, 공백 제거)
function formatAsLineList(set) {
  return [...set]
    .map(item => item.trim())
    .filter(item => item !== '')
    .join('\n');
}

function showMoveExpectedForm() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const all = ss.getSheets().map(s => s.getName());

  // "1." 으로 시작하는 시트를 우선순위로
  const priority = all.filter(n => n.startsWith('1.'));
  const others   = all.filter(n => !n.startsWith('1.'));
  const ordered  = [...priority, ...others];

  const tpl = HtmlService.createTemplateFromFile('MoveForm');
  tpl.sheetNames   = ordered;
  tpl.defaultSheet = ordered[0];

  const html = tpl.evaluate()
    .setWidth(350)
    .setHeight(150);
  ui.showModalDialog(html, '당일 확정 작업 시트 선택');
}


/**
 * ② 실제 이동 로직 (기존 moveExpectedToToday 내용, prompt/확인 빼고 sheetName 파라미터로)
 */
function moveExpectedToTodayByName(sheetName) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    ui.alert('오류', `"${sheetName}" 시트를 찾을 수 없습니다.`, ui.ButtonSet.OK);
    return;
  }

  // 1) 헤더·데이터 가져오기
  const headers = sheet.getRange('A2:O2').getValues()[0];
  const data    = sheet.getRange('A3:O18').getValues();

  // 2) 검증
  const issues = [];
  data.forEach((row, idx) => {
    const [aVal, bVal] = row;
    if (aVal && !bVal) {
      issues.push({ taskCount: idx+1,
        missingHeaders: headers.slice(1).filter((_, i) => i!==1)
      });
    } else if (bVal) {
      const missing = headers
        .map((h,i) => (i!==2 && !row[i])? h : null)
        .filter(h=>h);
      if (missing.length) issues.push({ taskCount: idx+1, missingHeaders: missing });
    }
  });
  if (issues.length) {
    const msg = issues.map((it,i) =>
      `(${i+1}) ${it.taskCount}번 작업: ${it.missingHeaders.join(', ')}`
    ).join('\n');
    ui.alert('취합 불가', msg, ui.ButtonSet.OK);
    return;
  }

  // 3) 복사→지우기
  sheet.getRange('A3:O18').copyTo(sheet.getRange('A22:O37'));
  sheet.getRange('A3:B18').clearContent();
  sheet.getRange('D3:O18').clearContent();

  // 4) 새 시트 복사 & 이름 설정
  const dateVal  = sheet.getRange('A22').getDisplayValue().replace(/-/g,'');
  const newName  = `(안심)${dateVal}`;
  const newSheet = sheet.copyTo(ss).setName(newName);

  // 5) 새 시트 1~19행 삭제 & 맨 뒤로 이동
  newSheet.deleteRows(1, 19);
  ss.setActiveSheet(newSheet);
  ss.moveActiveSheet(ss.getNumSheets());
}