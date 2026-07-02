/**
 * ============================================================
 * © 2026 GEG화성 (깊이 e끌림). All rights reserved.
 *
 * 본 코드는 「저작권법」의 보호를 받는 저작물입니다.
 * - 복제권(제16조)·공중송신권(제18조)·배포권(제20조)은
 *   저작권자에게 있습니다.
 * - 정상 경로로 받은 이용자라도 코드의 무단 복제·재배포·
 *   재판매·리브랜딩은 허용되지 않습니다.
 * - 무단 이용 시 「저작권법」 제136조(5년 이하 징역 또는
 *   5천만 원 이하 벌금) 및 제125조(손해배상) 적용 대상이
 *   될 수 있습니다.
 * - 이용 문의: bacusiki777@gmail.com, for2102@jimj.kr
 * ============================================================
 */

// 빌드 서명
const _BUILD_SIG = 'GEGHS-DEEPE-2026';

// 출처 확인용 함수
function getBuildInfo() {
  return {
    sig: _BUILD_SIG,
    owner: 'GEG화성 (깊이 e끌림)',
    year: 2026
  };
}

/**
 * 시트를 열 때 상단에 '도형 왕국' 메뉴를 만들어 줍니다.
 * 선생님이 처음 준비할 때 사용하는 항목들입니다.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('도형 왕국')
    .addItem('사용 설명 탭 만들기', 'setupGuideSheet')
    .addItem('학생 명단 탭 만들기', 'setupRosterSheet')
    .addToUi();
}

/**
 * 1. 웹 화면(index.html)을 브라우저에 표시해주는 함수
 * 웹 앱 링크로 학생들이 접속할 때 index.html의 내용으로 웹 브라우저 화면을 그려줍니다.
 * 스프레드시트에 묶인(bound) 스크립트라서, 사본을 만들면 그 사본 시트에 자동으로 연결됩니다.
 *
 * 학생 화면이 다른 주소에 배포된 경우에는, 요청에 담긴 항목(action)에 따라
 * 명단이나 기록 자료를 돌려주는 응답도 함께 처리합니다.
 */
function doGet(e) {
  var params = (e && e.parameter) || {};
  var action = params.action;

  // 학생 화면에서 이 시트의 자료(명단 읽기·기록 남기기)를 요청한 경우
  if (action) {
    var data;
    if (action === 'list') {
      data = { students: getStudentList() };
    } else if (action === 'record') {
      var result = recordMission(params.name, params.mission);
      data = { result: result };
    } else {
      data = { error: '알 수 없는 요청입니다.' };
    }

    var body = JSON.stringify(data);
    var cb = params.callback;
    // 학생 화면이 다른 주소에서도 자료를 안전하게 받을 수 있도록,
    // 요청에 담긴 callback 이름으로 감싸서 돌려줍니다.
    if (cb) {
      return ContentService.createTextOutput(cb + '(' + body + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(body)
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 기본: 앱 화면(index.html)을 그려 줍니다.
  var htmlOutput = HtmlService.createHtmlOutputFromFile('index');

  // 모바일/태블릿 등 기기 화면 비율에 맞추기 위한 설정
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');

  // 브라우저 탭 제목 설정
  htmlOutput.setTitle('도형 왕국의 잃어버린 일지');

  // 외부 프레임 접근 허용 보안 정책 설정
  htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  return htmlOutput;
}

/**
 * 2. '학생명단' 탭에서 명단 읽어오기 함수
 * 웹 페이지가 시작될 때 학생 이름 목록을 가져와 전달합니다.
 * (google.script.run.getStudentList() 로 화면에서 호출되거나, doGet의 명단 요청에 사용됩니다.)
 * '학생명단' 탭은 A열=번호, B열=이름 으로 되어 있습니다.
 */
function getStudentList() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("학생명단");
    if (!sheet) return [];

    var values = sheet.getDataRange().getValues();
    var students = [];

    // 0번째 행(1행)은 헤더이므로 제외하고 1번째 행(2행)부터 이름을 모읍니다.
    for (var i = 1; i < values.length; i++) {
      // B열(이름)을 우선 사용하고, 비어 있으면 A열을 사용합니다.
      var name = (values[i][1] !== undefined && values[i][1] !== '') ? values[i][1] : values[i][0];
      if (name !== undefined && name !== null && String(name).trim() !== '') {
        students.push(String(name).trim());
      }
    }
    return students; // 예: ["학생1", "학생2", ...] 또는 선생님이 바꾼 실제 이름
  } catch (err) {
    return [];
  }
}

/**
 * 3. 학생들의 미션 완료 상태를 '학습기록' 탭에 기록하는 함수
 * (google.script.run.recordMission(name, missionId) 로 화면에서 호출되거나, doGet의 기록 요청에 사용됩니다.)
 * @param {string} name - 학생 이름 (예: "학생1")
 * @param {string} missionId - 미션 고유 ID ('classify', 'prism', 'pyramid', 'pattern', 'net')
 */
function recordMission(name, missionId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("학습기록");

    // 혹시 '학습기록' 탭이 삭제되었거나 없다면 자동으로 만들어 줍니다.
    if (!sheet) {
      sheet = ss.insertSheet("학습기록");
      sheet.appendRow(["타임스탬프", "이름", "분류미션", "각기둥미션", "각뿔미션", "규칙미션", "전개도미션"]);
    }

    var values = sheet.getDataRange().getValues();
    var rowIndex = -1; // 학생의 행 위치를 저장할 변수 (-1은 아직 기록에 없는 상태)

    // 이미 이 이름으로 학습 중인 학생이 있는지 찾습니다 (B열 탐색)
    for (var i = 1; i < values.length; i++) {
      if (values[i][1] === name) { // values[i][1]은 B열(이름) 값입니다.
        rowIndex = i + 1; // 행 번호는 index + 1
        break;
      }
    }

    // 각 미션 ID별로 스프레드시트에 저장할 열의 번호를 설정합니다.
    // C열(3번째 열): 분류 / D열: 각기둥 / E열: 각뿔 / F열: 규칙 / G열: 전개도
    var colMap = {
      'classify': 3,
      'prism': 4,
      'pyramid': 5,
      'pattern': 6,
      'net': 7
    };
    var targetCol = colMap[missionId];
    var now = new Date(); // 현재 시간

    if (rowIndex !== -1) {
      // 1) 기록이 이미 있는 학생: 타임스탬프를 갱신하고, 완료한 미션 칸을 "완료"로 적어줍니다.
      sheet.getRange(rowIndex, 1).setValue(now);
      if (targetCol) {
        sheet.getRange(rowIndex, targetCol).setValue("완료");
      }
    } else {
      // 2) 처음 참여하는 학생: 새 줄을 만들어 이름과 완료한 미션을 기록합니다.
      var newRow = [now, name, "", "", "", "", ""];
      if (targetCol) {
        newRow[targetCol - 1] = "완료";
      }
      sheet.appendRow(newRow);
    }

    return "success";
  } catch (err) {
    return "error: " + err.toString();
  }
}

/**
 * '학생 명단 탭 만들기' 메뉴에서 실행합니다.
 * '학생명단' 탭이 없을 때만 A열=번호, B열=이름 으로 만들고 학생1~학생30을 미리 채웁니다.
 * 이미 있으면(선생님이 실제 이름을 넣어 두었을 수 있으므로) 건드리지 않습니다.
 */
function setupRosterSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var name = "학생명단";

  if (ss.getSheetByName(name)) {
    ui.alert("'학생명단' 탭이 이미 있습니다. 이름 칸을 우리 반 학생 이름으로 바꿔 사용하세요.");
    return;
  }

  var sheet = ss.insertSheet(name);

  var rows = [["번호", "이름"]];
  for (var i = 1; i <= 30; i++) {
    rows.push([i, "학생" + i]);
  }
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);

  // 헤더 서식 (앱 도형 테마색)
  sheet.getRange(1, 1, 1, 2)
    .setFontWeight('bold')
    .setBackground('#fde68a')
    .setHorizontalAlignment('center');
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 220);
  sheet.getRange(1, 1, rows.length, 2).setBorder(true, true, true, true, true, true);
  sheet.setFrozenRows(1);

  ui.alert("'학생명단' 탭을 만들었습니다. 이름 칸을 우리 반 학생 이름으로 바꿔 사용하세요.");
}

/**
 * '사용 설명 탭 만들기' 메뉴에서 실행합니다.
 * 선생님을 위한 '사용 설명' 탭을 첫 번째 위치에 새로 만듭니다.
 * (안내 내용이라 다시 눌러도 최신 내용으로 새로 만듭니다.)
 */
function setupGuideSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = "사용 설명";

  var old = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);
  var sheet = ss.insertSheet(name, 0); // 첫 번째 탭으로

  var title = "도형 왕국 - 사용 설명";
  var intro = "이 탭은 선생님을 위한 안내입니다. 학생 화면에는 나타나지 않습니다.";

  var setupSteps = [
    "이 시트를 선생님 계정으로 사본을 만드세요. (파일 메뉴 > 사본 만들기)",
    "상단의 확장 프로그램 > Apps Script 를 열어 코드가 함께 복사되었는지 확인하세요.",
    "Apps Script 오른쪽 위 배포 > 새 배포 를 누르고 유형을 웹 앱으로 선택하세요.",
    "액세스 권한을 모든 사용자로 설정한 뒤 배포하고, 만들어진 웹 앱 주소(끝이 exec 인 주소)를 복사하세요.",
    "학생용 앱 시작 화면 아래의 '선생님 설정'을 열어 그 주소를 붙여넣고 연결하기를 누르세요.",
    "화면에 생긴 학생용 링크를 복사해 학생들에게 나눠 주세요. 학생은 링크만 누르면 바로 시작합니다."
  ];
  var useSteps = [
    "'학생명단' 탭의 이름 칸을 우리 반 학생 이름으로 바꾸면, 앱의 이름 선택 목록에 그대로 나타납니다.",
    "학생이 미션을 마치면 '학습기록' 탭에 완료가 자동으로 기록됩니다.",
    "아직 시트를 연결하지 않았다면 앱은 체험 모드로 열리며, 결과는 학생 기기에만 임시로 저장됩니다."
  ];

  // 내용을 2차원 배열로 모아 한 번에 입력합니다.
  var rows = [];
  var titleRow, introRow, sectionRows = [];
  rows.push([title]); titleRow = rows.length;
  rows.push([intro]); introRow = rows.length;
  rows.push([""]);
  rows.push(["처음 준비 (한 번만)"]); sectionRows.push(rows.length);
  for (var i = 0; i < setupSteps.length; i++) rows.push([(i + 1) + ". " + setupSteps[i]]);
  rows.push([""]);
  rows.push(["수업에서 사용하기"]); sectionRows.push(rows.length);
  for (var j = 0; j < useSteps.length; j++) rows.push([(j + 1) + ". " + useSteps[j]]);
  rows.push([""]);
  rows.push(["데이터나 설정을 바꿀 때는 앱 화면이 아니라 해당 시트 탭에서 직접 수정하세요. 탭 이름은 코드와 연결되어 있으므로 삭제하거나 변경하지 마세요."]);

  var total = rows.length;
  sheet.getRange(1, 1, total, 1).setValues(rows);

  // 서식 (setValues 이후 한 번에)
  sheet.setColumnWidth(1, 720);
  var all = sheet.getRange(1, 1, total, 1);
  all.setWrap(true).setVerticalAlignment("top");
  all.setBorder(true, true, true, true, false, false);

  sheet.getRange(titleRow, 1).setFontSize(14).setFontWeight("bold").setBackground("#fde68a");
  sheet.getRange(introRow, 1).setFontStyle("italic").setFontColor("#6b7280");
  for (var s = 0; s < sectionRows.length; s++) {
    sheet.getRange(sectionRows[s], 1).setFontWeight("bold").setBackground("#fef3c7");
  }
}
