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

function moveExpectedToToday() {
  const ui    = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();

  // 1) 예/아니오 확인
  const resp = ui.alert(
    '작업 이동',
    '예상작업을 당일작업으로 옮기시겠습니까?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  // 2) 헤더(A2:N2)와 데이터(A3:N18) 가져오기
  const headers = sheet.getRange('A2:N2').getValues()[0];
  const data    = sheet.getRange('A3:N18').getValues();

  // 3) 검증 이슈 수집
  const issues = [];
  data.forEach((row, idx) => {
    const taskNum = idx + 1;
    const aVal    = row[0];
    const bVal    = row[1];

    // A만 있고 B가 비어 있으면 → B~N(C 제외) 모두 누락
    if (aVal && !bVal) {
      issues.push({
        taskCount: taskNum,
        missingHeaders: headers.slice(1).filter((_, i) => i !== 1)  // slice(1): B~N → filter로 C(인덱스2) 제외
      });
    }
    // B가 채워져 있으면 → A~N(C 제외) 전체 검증
    else if (bVal) {
      const missing = headers
        .map((h, i) => (i !== 2 && !row[i]) ? h : null)
        .filter(h => h);
      if (missing.length) {
        issues.push({ taskCount: taskNum, missingHeaders: missing });
      }
    }
    // A,B 둘 다 없으면 스킵
  });

  // 4) 이슈 있으면 메시지 띄우고 종료
  if (issues.length) {
    const msg = issues.map((itm, i) =>
      `(${i+1})\n` +
      `확인 필요 : ${itm.taskCount}번 작업\n` +
      `수정 및 기입 필요 : ${itm.missingHeaders.join(', ')}`
    ).join('\n\n');
    ui.alert('취합 불가', msg, ui.ButtonSet.OK);
    return;
  }

  // 5) 복사
  sheet.getRange('A3:N18').copyTo(sheet.getRange('A22:N37'));

  // 6) 원본 지우기 (C열은 유지)
  // A,B 열 지우기
sheet.getRange('A3:B18').clearContent();
sheet.getRange('D3:N18').clearContent();
}