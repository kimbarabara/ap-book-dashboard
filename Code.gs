/**
 * =====================================================================
 * 학원 교재 관리 대시보드 — Apps Script 백엔드 (Code.gs)
 * =====================================================================
 * 이 파일은 "그대로 붙여넣기"용입니다.
 * 스프레드시트에 연결된 컨테이너 바인딩 스크립트(확장 프로그램 > Apps Script)에
 * 이 코드를 그대로 붙여넣고, index.html 을 별도 HTML 파일로 추가하세요.
 *
 * ▼ 초보자 수정 영역 -----------------------------------------------------
 * 실제 스프레드시트의 "탭 이름"이나 "컬럼(헤더) 이름"이 다르다면
 * 아래 SHEET_NAMES / COLUMNS 값만 맞게 수정하면 됩니다.
 * (컬럼은 순서가 아니라 "헤더 이름"으로 찾으므로 시트 안에서 컬럼 순서를
 *  바꿔도 코드를 고칠 필요가 없습니다.)
 * ------------------------------------------------------------------- */

const SHEET_NAMES = {
  BOOK_DB: '교재DB',
  CLASS_SETTING: '반별교재셋팅',
  USER_PERMISSION: '사용자권한'
};

// 교재DB 탭의 헤더(컬럼) 이름
const BOOK_DB_FIELDS = {
  SUBJECT: '과목',
  BOOK_NAME: '교재명',
  PUBLISHER: '출판사',
  ISBN: 'ISBN',
  PRICE: '청구가격',
  GRADE: '학년',
  TARGET_MONTH: '대상 월'
};

// 반별교재셋팅 탭의 헤더(컬럼) 이름
const CLASS_SETTING_FIELDS = {
  CAMPUS: '관',
  GRADE_CLASS: '학년/반',
  BOOK_NAME: '교재명',
  TARGET_MONTH: '대상 월',
  MEMO: '메모'
};

// 사용자권한 탭의 헤더(컬럼) 이름
const USER_PERMISSION_FIELDS = {
  EMAIL: '이메일',
  ROLE: '구분'
};

const ROLE = {
  ADMIN: '관리자',
  TEACHER: '직원'
};

const COMMON_CAMPUS = '공용';   // '관' 값이 이 값이면 모든 관에서 공통으로 보여야 함
const COMMON_MONTH = '공통';    // '대상 월' 값이 이 값이면 어떤 월을 선택해도 항상 포함
const UNREGISTERED_LABEL = '교재DB 미등록';

/* ===================================================================
 * 공통 유틸
 * =================================================================== */

function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheetOrThrow_(sheetName) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('시트 탭을 찾을 수 없습니다: ' + sheetName);
  }
  return sheet;
}

/**
 * 시트를 [{헤더1: 값, 헤더2: 값, ...}, ...] 형태의 객체 배열로 읽어온다.
 * 1행을 헤더로 취급하며, 완전히 빈 행은 건너뛴다.
 */
function readSheetAsObjects_(sheetName) {
  const sheet = getSheetOrThrow_(sheetName);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];

  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(function (h) { return String(h).trim(); });

  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const isEmpty = row.every(function (cell) { return cell === '' || cell === null; });
    if (isEmpty) continue;

    const obj = {};
    for (let c = 0; c < headers.length; c++) {
      obj[headers[c]] = row[c];
    }
    rows.push(obj);
  }
  return rows;
}

function normalize_(value) {
  return String(value === null || value === undefined ? '' : value).trim();
}

/* ===================================================================
 * 접근 제어 (로직 4)
 * =================================================================== */

function getCurrentUserEmail_() {
  const email = Session.getActiveUser().getEmail();
  return normalize_(email).toLowerCase();
}

/**
 * 사용자권한 시트를 조회해 현재 접속자의 역할을 반환한다.
 * 등록되어 있지 않으면 null.
 */
function getCurrentUserRole_() {
  const email = getCurrentUserEmail_();
  if (!email) return null;

  const users = readSheetAsObjects_(SHEET_NAMES.USER_PERMISSION);
  for (let i = 0; i < users.length; i++) {
    const rowEmail = normalize_(users[i][USER_PERMISSION_FIELDS.EMAIL]).toLowerCase();
    if (rowEmail === email) {
      return normalize_(users[i][USER_PERMISSION_FIELDS.ROLE]);
    }
  }
  return null;
}

/**
 * google.script.run 으로 호출되는 모든 서버 함수는 이 함수로 매번 역할을 재검증한다.
 * (화면에서 숨기는 방식이 아니라 서버에서 검증 — 요구사항 필수 조건)
 */
function requireRegisteredUser_() {
  const email = getCurrentUserEmail_();
  const role = getCurrentUserRole_();
  if (!role || (role !== ROLE.ADMIN && role !== ROLE.TEACHER)) {
    throw new Error('ACCESS_DENIED');
  }
  return { email: email, role: role };
}

function requireAdmin_() {
  const auth = requireRegisteredUser_();
  if (auth.role !== ROLE.ADMIN) {
    throw new Error('ACCESS_DENIED');
  }
  return auth;
}

/* ===================================================================
 * 웹앱 진입점
 * =================================================================== */

function doGet(e) {
  const email = getCurrentUserEmail_();
  const role = getCurrentUserRole_();

  if (!role) {
    const blocked = HtmlService.createHtmlOutput(
      '<html><body style="font-family:sans-serif;text-align:center;padding-top:120px;">' +
      '<h2>접근 권한이 없습니다.</h2>' +
      '<p>관리자에게 문의하세요.</p>' +
      '<p style="color:#888;font-size:13px;">(' + (email || '알 수 없음') + ')</p>' +
      '</body></html>'
    );
    blocked.setTitle('접근 제한');
    return blocked;
  }

  const template = HtmlService.createTemplateFromFile('index');
  template.userEmail = email;
  template.userRole = role;
  return template.evaluate()
    .setTitle('학원 교재 관리 대시보드')
    .addMetaTag('viewport', 'width=1280');
}

/* ===================================================================
 * 클라이언트에서 호출하는 함수들 (google.script.run)
 * 모두 requireRegisteredUser_() 로 역할을 재검증한다.
 * =================================================================== */

/**
 * 앱 로드시 필요한 초기 정보: 로그인 정보 + 드롭다운 옵션 목록
 */
function initApp() {
  const auth = requireRegisteredUser_();
  const settings = readSheetAsObjects_(SHEET_NAMES.CLASS_SETTING);
  const books = readSheetAsObjects_(SHEET_NAMES.BOOK_DB);

  const campusSet = {};
  const gradeSet = {};
  const monthSet = {};
  const subjectSet = {};

  settings.forEach(function (row) {
    const campus = normalize_(row[CLASS_SETTING_FIELDS.CAMPUS]);
    const grade = normalize_(row[CLASS_SETTING_FIELDS.GRADE_CLASS]);
    const month = normalize_(row[CLASS_SETTING_FIELDS.TARGET_MONTH]);

    if (campus && campus !== COMMON_CAMPUS) campusSet[campus] = true;
    if (grade) gradeSet[grade] = true;
    if (month && month !== COMMON_MONTH) monthSet[month] = true;
  });

  books.forEach(function (row) {
    const subject = normalize_(row[BOOK_DB_FIELDS.SUBJECT]);
    if (subject) subjectSet[subject] = true;
  });

  function toSortedArray(obj) {
    return Object.keys(obj).sort(function (a, b) { return a.localeCompare(b, 'ko'); });
  }

  return {
    email: auth.email,
    role: auth.role,
    isAdmin: auth.role === ROLE.ADMIN,
    campuses: toSortedArray(campusSet),
    grades: toSortedArray(gradeSet),
    months: toSortedArray(monthSet),
    subjects: toSortedArray(subjectSet)
  };
}

/**
 * 섹션 A — 교재 낱개 검색 (교재명 또는 ISBN 부분 일치)
 */
function searchBooks(keyword) {
  requireRegisteredUser_();

  const query = normalize_(keyword).toLowerCase();
  if (!query) return [];

  const books = readSheetAsObjects_(SHEET_NAMES.BOOK_DB);
  return books
    .filter(function (row) {
      const name = normalize_(row[BOOK_DB_FIELDS.BOOK_NAME]).toLowerCase();
      const isbn = normalize_(row[BOOK_DB_FIELDS.ISBN]).toLowerCase();
      return name.indexOf(query) !== -1 || isbn.indexOf(query) !== -1;
    })
    .map(function (row) {
      return {
        subject: normalize_(row[BOOK_DB_FIELDS.SUBJECT]),
        bookName: normalize_(row[BOOK_DB_FIELDS.BOOK_NAME]),
        publisher: normalize_(row[BOOK_DB_FIELDS.PUBLISHER]),
        isbn: normalize_(row[BOOK_DB_FIELDS.ISBN]),
        price: Number(row[BOOK_DB_FIELDS.PRICE]) || 0,
        grade: normalize_(row[BOOK_DB_FIELDS.GRADE]),
        targetMonth: normalize_(row[BOOK_DB_FIELDS.TARGET_MONTH])
      };
    });
}

/**
 * 섹션 B — 반별 교재 조회 (로직 1: 필터링 + 조인)
 */
function getClassBooks(campus, grade, month) {
  const auth = requireRegisteredUser_();

  const selectedCampus = normalize_(campus);
  const selectedGrade = normalize_(grade);
  const selectedMonth = normalize_(month);

  const settings = readSheetAsObjects_(SHEET_NAMES.CLASS_SETTING);
  const books = readSheetAsObjects_(SHEET_NAMES.BOOK_DB);

  // 교재명 -> 교재DB 행, 빠른 조인을 위한 맵 (동일 교재명이 여러 개면 첫 항목 사용)
  const bookMap = {};
  books.forEach(function (row) {
    const key = normalize_(row[BOOK_DB_FIELDS.BOOK_NAME]);
    if (key && !bookMap[key]) bookMap[key] = row;
  });

  const filtered = settings.filter(function (row) {
    const rowCampus = normalize_(row[CLASS_SETTING_FIELDS.CAMPUS]);
    const rowGrade = normalize_(row[CLASS_SETTING_FIELDS.GRADE_CLASS]);
    const rowMonth = normalize_(row[CLASS_SETTING_FIELDS.TARGET_MONTH]);

    const campusMatch = rowCampus === selectedCampus || rowCampus === COMMON_CAMPUS;
    const gradeMatch = rowGrade === selectedGrade;
    const monthMatch = rowMonth === COMMON_MONTH || rowMonth === selectedMonth;

    return campusMatch && gradeMatch && monthMatch;
  });

  const result = filtered.map(function (row, index) {
    const bookName = normalize_(row[CLASS_SETTING_FIELDS.BOOK_NAME]);
    const bookInfo = bookMap[bookName];
    const registered = !!bookInfo;

    return {
      no: index + 1,
      bookName: bookName,
      subject: registered ? normalize_(bookInfo[BOOK_DB_FIELDS.SUBJECT]) : UNREGISTERED_LABEL,
      publisher: registered ? normalize_(bookInfo[BOOK_DB_FIELDS.PUBLISHER]) : UNREGISTERED_LABEL,
      isbn: registered ? normalize_(bookInfo[BOOK_DB_FIELDS.ISBN]) : UNREGISTERED_LABEL,
      price: registered ? (Number(bookInfo[BOOK_DB_FIELDS.PRICE]) || 0) : 0,
      priceLabel: registered ? (Number(bookInfo[BOOK_DB_FIELDS.PRICE]) || 0).toLocaleString('ko-KR') + '원' : UNREGISTERED_LABEL,
      targetMonth: normalize_(row[CLASS_SETTING_FIELDS.TARGET_MONTH]),
      registered: registered
    };
  });

  return {
    role: auth.role,
    isAdmin: auth.role === ROLE.ADMIN,
    condition: { campus: selectedCampus, grade: selectedGrade, month: selectedMonth },
    rows: result
  };
}

/**
 * 섹션 C — 관/학년/과목 조회 (월 조건 없이 전체 월 대상, 과목으로 필터링)
 * 교재DB에 등록되지 않은 교재는 과목을 알 수 없으므로 결과에서 제외한다.
 */
function getBooksBySubject(campus, grade, subject) {
  const auth = requireRegisteredUser_();

  const selectedCampus = normalize_(campus);
  const selectedGrade = normalize_(grade);
  const selectedSubject = normalize_(subject);

  const settings = readSheetAsObjects_(SHEET_NAMES.CLASS_SETTING);
  const books = readSheetAsObjects_(SHEET_NAMES.BOOK_DB);

  const bookMap = {};
  books.forEach(function (row) {
    const key = normalize_(row[BOOK_DB_FIELDS.BOOK_NAME]);
    if (key && !bookMap[key]) bookMap[key] = row;
  });

  const filtered = settings.filter(function (row) {
    const rowCampus = normalize_(row[CLASS_SETTING_FIELDS.CAMPUS]);
    const rowGrade = normalize_(row[CLASS_SETTING_FIELDS.GRADE_CLASS]);
    const bookName = normalize_(row[CLASS_SETTING_FIELDS.BOOK_NAME]);
    const bookInfo = bookMap[bookName];

    const campusMatch = rowCampus === selectedCampus || rowCampus === COMMON_CAMPUS;
    const gradeMatch = rowGrade === selectedGrade;
    const subjectMatch = !!bookInfo && normalize_(bookInfo[BOOK_DB_FIELDS.SUBJECT]) === selectedSubject;

    return campusMatch && gradeMatch && subjectMatch;
  });

  const result = filtered.map(function (row, index) {
    const bookName = normalize_(row[CLASS_SETTING_FIELDS.BOOK_NAME]);
    const bookInfo = bookMap[bookName];

    return {
      no: index + 1,
      bookName: bookName,
      subject: normalize_(bookInfo[BOOK_DB_FIELDS.SUBJECT]),
      publisher: normalize_(bookInfo[BOOK_DB_FIELDS.PUBLISHER]),
      isbn: normalize_(bookInfo[BOOK_DB_FIELDS.ISBN]),
      price: Number(bookInfo[BOOK_DB_FIELDS.PRICE]) || 0,
      priceLabel: (Number(bookInfo[BOOK_DB_FIELDS.PRICE]) || 0).toLocaleString('ko-KR') + '원',
      targetMonth: normalize_(row[CLASS_SETTING_FIELDS.TARGET_MONTH]),
      registered: true
    };
  });

  return {
    role: auth.role,
    isAdmin: auth.role === ROLE.ADMIN,
    condition: { campus: selectedCampus, grade: selectedGrade, subject: selectedSubject },
    rows: result
  };
}

/* ===================================================================
 * 샘플 데이터 세팅 (Apps Script 편집기에서 관리자가 직접 1회 실행)
 * 함수 목록에서 setupSampleData 선택 후 "실행" 버튼으로 실행
 * 이미 존재하는 탭은 건드리지 않고, 없을 때만 새로 만든다.
 * =================================================================== */

function setupSampleData() {
  const ss = getSpreadsheet_();
  createSheetIfMissing_(ss, SHEET_NAMES.BOOK_DB, [
    BOOK_DB_FIELDS.SUBJECT, BOOK_DB_FIELDS.BOOK_NAME, BOOK_DB_FIELDS.PUBLISHER,
    BOOK_DB_FIELDS.ISBN, BOOK_DB_FIELDS.PRICE, BOOK_DB_FIELDS.GRADE, BOOK_DB_FIELDS.TARGET_MONTH
  ], [
    ['영어', 'Reading Explorer 4', 'NGL', '9781111222333', 18000, '4학년', COMMON_MONTH],
    ['영어', 'Grammar in Use Basic', 'Cambridge', '9784444555666', 22000, '4학년', COMMON_MONTH],
    ['영어', 'Vocabulary 4000', 'Compass', '9787777888999', 15000, '4학년', '8월'],
    ['수학', '개념원리 5-1', '개념원리', '9781234567890', 17000, '5학년', COMMON_MONTH]
  ]);

  createSheetIfMissing_(ss, SHEET_NAMES.CLASS_SETTING, [
    CLASS_SETTING_FIELDS.CAMPUS, CLASS_SETTING_FIELDS.GRADE_CLASS, CLASS_SETTING_FIELDS.BOOK_NAME,
    CLASS_SETTING_FIELDS.TARGET_MONTH, CLASS_SETTING_FIELDS.MEMO
  ], [
    [COMMON_CAMPUS, '4학년', 'Reading Explorer 4', COMMON_MONTH, '전 관 공통 원서'],
    ['1관', '4학년', 'Grammar in Use Basic', COMMON_MONTH, ''],
    ['1관', '4학년', 'Vocabulary 4000', '8월', ''],
    ['2관', '4학년', 'Grammar in Use Basic', COMMON_MONTH, ''],
    ['2관', '5학년', '개념원리 5-1', COMMON_MONTH, '']
  ]);

  createSheetIfMissing_(ss, SHEET_NAMES.USER_PERMISSION, [
    USER_PERMISSION_FIELDS.EMAIL, USER_PERMISSION_FIELDS.ROLE
  ], [
    ['admin@example.com', ROLE.ADMIN],
    ['teacher@example.com', ROLE.TEACHER]
  ]);
}

function createSheetIfMissing_(ss, sheetName, headers, sampleRows) {
  if (ss.getSheetByName(sheetName)) return;
  const sheet = ss.insertSheet(sheetName);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  if (sampleRows && sampleRows.length) {
    sheet.getRange(2, 1, sampleRows.length, headers.length).setValues(sampleRows);
  }
  sheet.autoResizeColumns(1, headers.length);
}
