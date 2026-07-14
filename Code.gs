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
  USER_PERMISSION: '사용자권한',
  ACCESS_LOG: '접속기록'
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

// 접속기록 탭의 헤더(컬럼) 이름 — 없으면 자동으로 생성됨
const ACCESS_LOG_FIELDS = {
  EMAIL: '이메일',
  ROLE: '역할',
  TIME: '접속일시'
};

/**
 * 3단계 권한
 * - ADMIN(관리자): 청구 관리 기능 전부 + 접속 기록/현재 접속자 확인 + DB 시트 바로가기
 * - STAFF(직원): 청구 관리 기능(체크박스/청구총액/선택분 PDF) — 예전 "관리자" 권한과 동일
 * - VIEWER(뷰어): 조회 + 전체 목록 PDF만 — 예전 "직원" 권한과 동일
 */
const ROLE = {
  ADMIN: '관리자',
  STAFF: '직원',
  VIEWER: '뷰어'
};

// 청구 총액/체크박스/선택분 PDF 등 "관리 기능"을 볼 수 있는 역할
function isPrivilegedRole_(role) {
  return role === ROLE.ADMIN || role === ROLE.STAFF;
}

const COMMON_CAMPUS = '공통';   // '관' 값이 이 값이면 모든 관에서 공통으로 보여야 함
const COMMON_MONTH = '공통';    // '대상 월' 값이 이 값이면 어떤 월을 선택해도 항상 포함
const UNREGISTERED_LABEL = '교재DB 미등록';

// 관 드롭다운에 노출할 관 목록 (이 목록에 없는 값은 시트에 있어도 검색 옵션에 나타나지 않음)
const SELECTABLE_CAMPUSES = ['1관', '2관'];

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

/**
 * 검색/조회 결과 공통 정렬: 1차 과목 가나다순, 2차 사용 월(대상 월) 가나다순
 */
function sortBySubjectThenMonth_(rows) {
  return rows.slice().sort(function (a, b) {
    const subjectCompare = normalize_(a.subject).localeCompare(normalize_(b.subject), 'ko');
    if (subjectCompare !== 0) return subjectCompare;
    return normalize_(a.targetMonth).localeCompare(normalize_(b.targetMonth), 'ko');
  });
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
  if (!role || (role !== ROLE.ADMIN && role !== ROLE.STAFF && role !== ROLE.VIEWER)) {
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

  logAccess_(email, role);

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

/**
 * 접속기록 탭에 접속 로그를 한 줄 남긴다. 탭이 없으면 자동 생성한다.
 * 미등록 이메일의 접속 시도도 남겨서 관리자가 확인할 수 있게 한다.
 */
function logAccess_(email, role) {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(SHEET_NAMES.ACCESS_LOG);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.ACCESS_LOG);
    sheet.getRange(1, 1, 1, 3).setValues([[
      ACCESS_LOG_FIELDS.EMAIL, ACCESS_LOG_FIELDS.ROLE, ACCESS_LOG_FIELDS.TIME
    ]]).setFontWeight('bold');
  }
  sheet.appendRow([email || '(알 수 없음)', role || '미등록', new Date()]);
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

  const gradeSet = {};
  const monthSet = {};
  const subjectSet = {};

  settings.forEach(function (row) {
    const grade = normalize_(row[CLASS_SETTING_FIELDS.GRADE_CLASS]);
    const month = normalize_(row[CLASS_SETTING_FIELDS.TARGET_MONTH]);

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
    isAdmin: isPrivilegedRole_(auth.role),
    isTopAdmin: auth.role === ROLE.ADMIN,
    spreadsheetUrl: auth.role === ROLE.ADMIN ? getSpreadsheet_().getUrl() : null,
    campuses: SELECTABLE_CAMPUSES,
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
  const matched = books
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

  return sortBySubjectThenMonth_(matched);
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

  const joined = filtered.map(function (row) {
    const bookName = normalize_(row[CLASS_SETTING_FIELDS.BOOK_NAME]);
    const bookInfo = bookMap[bookName];
    const registered = !!bookInfo;

    return {
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

  const result = sortBySubjectThenMonth_(joined).map(function (row, index) {
    row.no = index + 1;
    return row;
  });

  return {
    role: auth.role,
    isAdmin: isPrivilegedRole_(auth.role),
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

  const joined = filtered.map(function (row) {
    const bookName = normalize_(row[CLASS_SETTING_FIELDS.BOOK_NAME]);
    const bookInfo = bookMap[bookName];

    return {
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

  const result = sortBySubjectThenMonth_(joined).map(function (row, index) {
    row.no = index + 1;
    return row;
  });

  return {
    role: auth.role,
    isAdmin: isPrivilegedRole_(auth.role),
    condition: { campus: selectedCampus, grade: selectedGrade, subject: selectedSubject },
    rows: result
  };
}

/* ===================================================================
 * 관리자 전용 — 접속 기록 / 현재 접속자 (근사치)
 * =================================================================== */

const ONLINE_CACHE_PREFIX = 'online:';
const ONLINE_CACHE_TTL_SECONDS = 150; // 이 시간 안에 하트비트가 없으면 "접속 중"에서 자동 제외됨

/**
 * 클라이언트가 페이지를 열어둔 동안 주기적으로 호출한다.
 * 실제 웹소켓처럼 완벽한 실시간은 아니고, 마지막 신호 후 TTL 동안만
 * "현재 접속 중"으로 표시되는 근사치 방식이다.
 */
function heartbeat() {
  const auth = requireRegisteredUser_();
  const cache = CacheService.getScriptCache();
  cache.put(ONLINE_CACHE_PREFIX + auth.email, JSON.stringify({
    email: auth.email,
    role: auth.role
  }), ONLINE_CACHE_TTL_SECONDS);
  return true;
}

/**
 * 사용자권한에 등록된 이메일 중 최근 하트비트가 살아있는 사람만 반환한다.
 */
function getOnlineUsers() {
  requireAdmin_();

  const users = readSheetAsObjects_(SHEET_NAMES.USER_PERMISSION);
  const keys = users
    .map(function (row) { return normalize_(row[USER_PERMISSION_FIELDS.EMAIL]).toLowerCase(); })
    .filter(function (email) { return !!email; })
    .map(function (email) { return ONLINE_CACHE_PREFIX + email; });

  if (!keys.length) return [];

  const cache = CacheService.getScriptCache();
  const cached = cache.getAll(keys);
  return Object.keys(cached).map(function (key) {
    return JSON.parse(cached[key]);
  });
}

/**
 * 접속기록 탭의 최근 기록을 최신순으로 반환한다 (관리자 전용).
 */
function getAccessLogs(limit) {
  requireAdmin_();

  const rows = readSheetAsObjects_(SHEET_NAMES.ACCESS_LOG);
  const maxCount = limit || 200;

  return rows
    .map(function (row) {
      const time = row[ACCESS_LOG_FIELDS.TIME];
      return {
        email: normalize_(row[ACCESS_LOG_FIELDS.EMAIL]),
        role: normalize_(row[ACCESS_LOG_FIELDS.ROLE]),
        time: (time instanceof Date) ? Utilities.formatDate(time, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : normalize_(time)
      };
    })
    .reverse()
    .slice(0, maxCount);
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
    ['staff@example.com', ROLE.STAFF],
    ['viewer@example.com', ROLE.VIEWER]
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
