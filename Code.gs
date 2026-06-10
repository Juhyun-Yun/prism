/**
 * 1. 웹 화면(index.html)을 브라우저에 표시해주는 함수
 * 웹 앱 링크로 학생들이 접속할 때 index.html의 내용으로 웹 브라우저 화면을 그려줍니다.
 * 스프레드시트에 묶인(bound) 스크립트라서, 사본을 만들면 그 사본 시트에 자동으로 연결됩니다.
 */
function doGet(e) {
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
 * 2. '학생명단' 시트에서 명단 읽어오기 함수
 * 웹 페이지가 시작될 때 스프레드시트의 A열에 있는 학생 이름 리스트를 가져와 전달합니다.
 * (google.script.run.getStudentList() 로 화면에서 호출됩니다.)
 */
function getStudentList() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("학생명단");
    if (!sheet) return ["학생명단 시트가 없습니다."];

    var values = sheet.getDataRange().getValues();
    var students = [];

    // 0번째 행(1행)은 '이름' 헤더이므로 제외하고 1번째 행(2행)부터 이름 추가
    for (var i = 1; i < values.length; i++) {
      if (values[i][0]) {
        students.push(values[i][0].toString().trim());
      }
    }
    return students; // 예: ["1번 김도형", "2번 박탐험", ...]
  } catch (err) {
    return ["에러 발생: " + err.toString()];
  }
}

/**
 * 3. 학생들의 미션 완료 상태를 '학습기록' 시트에 기록하는 함수
 * (google.script.run.recordMission(name, missionId) 로 화면에서 호출됩니다.)
 * @param {string} name - 학생 이름 (예: "1번 김도형")
 * @param {string} missionId - 미션 고유 ID ('classify', 'prism', 'pyramid', 'pattern', 'net')
 */
function recordMission(name, missionId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("학습기록");

    // 혹시 '학습기록' 시트가 삭제되었거나 없다면 자동으로 만들어 줍니다.
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
