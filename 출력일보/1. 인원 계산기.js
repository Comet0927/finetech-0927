const ss = SpreadsheetApp.getActiveSpreadsheet();

/**
 * 설치 트리거: 애드온을 설치하거나
 * 배포된 웹 애플리케이션을 설치할 때 한 번만 호출됩니다.
 */
function onInstall(e) {
  // onOpen에게 이벤트 넘겨서
  // 설치 후 바로 메뉴가 보이도록 설정
  onOpen(e);
}

/**
 * 스프레드시트 열 때마다 호출되는 onOpen
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();

  // 1. 인원 계산기 메뉴
  ui.createMenu('1. 인원 계산기')
    .addItem('시트 복사 및 출력 복사', 'showCombinedForm')
    .addToUi();

}

// ✅ 2. 메뉴에서 호출되는 다이얼로그 열기
function showCombinedForm() {
  const html = HtmlService.createHtmlOutputFromFile('CombinedForm')
    .setWidth(450)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, "출력 확정 및 시트 복사");
}

// ✅ 3. 폼에서 사용하는 시트 목록 제공
function getSheetNames() {
  return ss.getSheets().map(s => s.getName());
}

// 4. 선택된 시트들로 전체 작업 실행
function copySheetWithDualSources(outputSheetName, originalSheetName, customDate) {
  const dateStr = customDate || Utilities.formatDate(new Date(), "Asia/Seoul", "yyyyMMdd");
  const newSheetName = `(출력)${dateStr}`;

  // 1) 동일 이름 시트가 있으면 에러 리턴
  if (ss.getSheetByName(newSheetName)) {
    return {
      success: false,
      message: `이미 "${newSheetName}" 시트가 존재합니다.`
    };
  }

  const original = ss.getSheetByName(originalSheetName);
  const output   = ss.getSheetByName(outputSheetName);

  // 2) 원본 시트 복사 → 새 이름 지정 → 활성화
  const copy = original.copyTo(ss).setName(newSheetName);
  ss.setActiveSheet(copy);

  // 날짜 문자열을 파싱해서 자정으로 설정된 Date 객체 생성
  const year  = Number(dateStr.substring(0, 4));
  const month = Number(dateStr.substring(4, 6)) - 1;
  const day   = Number(dateStr.substring(6, 8));
  const dateObj = new Date(year, month, day);

  // 3) 사본 시트 A1에 날짜(YYYYMMDD)만 입력
  const copycell = copy.getRange("A1");
  copycell.setValue(dateObj);
  copycell.setNumberFormat("yyyyMMdd");

  // 4) 출력 시트에서 필요한 구역 복사
  output.getRange("B17:F26")
        .setValues(output.getRange("L17:P26").getValues());

  // 5) 원본 시트 값 덮어쓰기
  original.getRange("H4:L11").setValues(original.getRange("H4:L11").getValues());
  original.getRange("H13:L87").setValues(original.getRange("H13:L87").getValues());

  // 6) 원본 시트 A3에 "출력인원 확정" 표시
  original.getRange("A3").setValue("출력인원 확정");

  // 7) 원본 → 사본 복사
  copy.getRange("C4:G11").setValues(original.getRange("M4:Q11").getValues());
  copy.getRange("C13:G87").setValues(original.getRange("M13:Q87").getValues());

  // 8) 출력 시트 초기화 및 A1에 날짜만 입력
  output.getRange("B3:U12").clearContent();
  const outputcell = output.getRange("A1");
  outputcell.setValue(dateObj);
  outputcell.setNumberFormat("yyyyMMdd");
  output.getRange("B13:U13").clearContent();

  return {
    success: true,
    message: `"${originalSheetName}" 시트를 "${newSheetName}"로 복사했습니다.\n` +
             `사본과 출력 시트의 A1 셀에 날짜(${dateStr})가 입력되었습니다.`
  };
}

