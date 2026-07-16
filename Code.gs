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
  AUDIT_LOG: '감사로그',
  UPDATE_LOG: '업데이트'
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
  ROLE: '구분',
  NAME: '이름',
  REGISTERED_DATE: '등록일'
};

// 감사로그 탭의 헤더(컬럼) 이름
const AUDIT_LOG_FIELDS = {
  TIME: '시각',
  ACTOR: '수행자이메일',
  ACTION: '동작유형',
  TARGET: '대상',
  SUMMARY: '변경내용요약'
};

// 업데이트 탭의 헤더(컬럼) 이름
const UPDATE_LOG_FIELDS = {
  DATE: '날짜',
  CATEGORY: '분류',
  CONTENT: '내용',
  VISIBLE: '표시여부'
};

/**
 * 3단계 권한
 * - ADMIN(관리자): 청구 관리 기능 전부 + 접속 기록/현재 접속자 확인 + DB 시트 바로가기
 *   + 미등록 교재 요청함 처리 + 권한 관리 화면 + 감사 로그 조회
 * - STAFF(직원): 청구 관리 기능(체크박스/청구총액/선택분 PDF) + 미등록 교재 등록 요청
 * - VIEWER(뷰어): 조회 + 전체 목록 PDF(가격 없이) + 미등록 교재 등록 요청
 */
const ROLE = {
  ADMIN: '관리자',
  STAFF: '직원',
  VIEWER: '뷰어'
};

// 청구 총액/체크박스/선택분 PDF/가격 표시 등 "관리 기능"을 볼 수 있는 역할
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
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
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

/**
 * 반 이름(예: '6M/6P')을 '/'로 분리해 각 토큰이 숫자로 "시작"하는 경우에만
 * 그 선행 숫자를 학년으로 추출한다. 알파벳으로 시작하는 토큰(G2, G2.5 등)은
 * 숫자가 있어도 레벨 표기로 간주해 학년을 추출하지 않는다.
 * 예: '6M/6P' -> {'6': true}, 'G2' -> {}, '5S' -> {'5': true}
 */
function extractGradeNumbers_(gradeLabel) {
  const tokens = normalize_(gradeLabel).split('/');
  const grades = {};
  tokens.forEach(function (token) {
    const m = token.trim().match(/^(\d+)/);
    if (m) grades[m[1]] = true;
  });
  return grades;
}

/**
 * 반별교재셋팅에 등록된 반(rowGrade)이 조회 조건으로 선택한 반(selectedGrade)에 해당하는지 판정한다.
 * 1) 완전 일치면 매칭 (기존 동작 유지)
 * 2) 등록반이 'N학년' 형식이면, 선택반에서 추출한 학년 숫자 집합에 N이 있으면 매칭
 *    (선택반이 'G2' 처럼 알파벳으로 시작하는 레벨 표기면 학년이 추출되지 않으므로 매칭되지 않는다)
 */
function gradeMatches_(rowGrade, selectedGrade) {
  const row = normalize_(rowGrade);
  const selected = normalize_(selectedGrade);
  if (row === selected) return true;

  const gradeLabelMatch = row.match(/^(\d+)학년$/);
  if (!gradeLabelMatch) return false;

  const selectedGradeNumbers = extractGradeNumbers_(selected);
  return !!selectedGradeNumbers[gradeLabelMatch[1]];
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
 * 접속 로그를 Apps Script 자체 저장소(PropertiesService)에 남긴다.
 * 스프레드시트에 직접 쓰지 않는 이유: 뷰어 권한만 있는 계정은 시트에 쓸 수 없어서
 * (뷰어=읽기 전용) 시트에 로그를 기록하면 그 계정들의 접속 자체가 에러로 막힌다.
 * PropertiesService는 시트 공유 권한과 무관하게 동작한다.
 * 미등록 이메일의 접속 시도도 남겨서 관리자가 확인할 수 있게 한다.
 */
const ACCESS_LOG_KEY_PREFIX = 'log_';
const MAX_ACCESS_LOGS = 500;

function logAccess_(email, role) {
  try {
    const props = PropertiesService.getScriptProperties();
    const key = ACCESS_LOG_KEY_PREFIX + Utilities.formatDate(new Date(), 'Etc/UTC', "yyyyMMdd'T'HHmmss'Z'") + '_' + Math.random().toString(36).slice(2, 8);
    props.setProperty(key, JSON.stringify({
      email: email || '(알 수 없음)',
      role: role || '미등록',
      time: new Date().toISOString()
    }));
    pruneAccessLogsIfNeeded_(props);
  } catch (e) {
    // 접속 기록 저장에 실패해도 로그인 자체는 막지 않는다.
  }
}

function pruneAccessLogsIfNeeded_(props) {
  const keys = props.getKeys().filter(function (k) { return k.indexOf(ACCESS_LOG_KEY_PREFIX) === 0; });
  if (keys.length <= MAX_ACCESS_LOGS + 50) return;
  keys.sort();
  const toDelete = keys.slice(0, keys.length - MAX_ACCESS_LOGS);
  toDelete.forEach(function (k) { props.deleteProperty(k); });
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

  const result = {
    email: auth.email,
    role: auth.role,
    isAdmin: isPrivilegedRole_(auth.role),
    isTopAdmin: auth.role === ROLE.ADMIN,
    spreadsheetUrl: auth.role === ROLE.ADMIN ? getSpreadsheet_().getUrl() : null,
    campuses: SELECTABLE_CAMPUSES,
    grades: toSortedArray(gradeSet),
    months: toSortedArray(monthSet),
    subjects: toSortedArray(subjectSet),
    rejectedNotices: getMyBookRequestNotices()
  };

  if (result.isTopAdmin) {
    result.pendingRequestCount = getAllBookRequests_().filter(function (r) { return r.status === '대기'; }).length;
  }

  return result;
}

/**
 * 섹션 A — 교재 낱개 검색 (교재명 또는 ISBN 부분 일치)
 */
function searchBooks(keyword) {
  const auth = requireRegisteredUser_();

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
      const item = {
        subject: normalize_(row[BOOK_DB_FIELDS.SUBJECT]),
        bookName: normalize_(row[BOOK_DB_FIELDS.BOOK_NAME]),
        publisher: normalize_(row[BOOK_DB_FIELDS.PUBLISHER]),
        isbn: normalize_(row[BOOK_DB_FIELDS.ISBN]),
        grade: normalize_(row[BOOK_DB_FIELDS.GRADE]),
        targetMonth: normalize_(row[BOOK_DB_FIELDS.TARGET_MONTH])
      };
      if (isPrivilegedRole_(auth.role)) {
        item.price = Number(row[BOOK_DB_FIELDS.PRICE]) || 0;
      }
      return item;
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
    const gradeMatch = gradeMatches_(rowGrade, selectedGrade);
    const monthMatch = rowMonth === COMMON_MONTH || rowMonth === selectedMonth;

    return campusMatch && gradeMatch && monthMatch;
  });

  const showPrice = isPrivilegedRole_(auth.role);
  const joined = filtered.map(function (row) {
    const bookName = normalize_(row[CLASS_SETTING_FIELDS.BOOK_NAME]);
    const bookInfo = bookMap[bookName];
    const registered = !!bookInfo;

    const item = {
      bookName: bookName,
      subject: registered ? normalize_(bookInfo[BOOK_DB_FIELDS.SUBJECT]) : UNREGISTERED_LABEL,
      publisher: registered ? normalize_(bookInfo[BOOK_DB_FIELDS.PUBLISHER]) : UNREGISTERED_LABEL,
      isbn: registered ? normalize_(bookInfo[BOOK_DB_FIELDS.ISBN]) : UNREGISTERED_LABEL,
      targetMonth: normalize_(row[CLASS_SETTING_FIELDS.TARGET_MONTH]),
      registered: registered
    };
    if (showPrice) {
      item.price = registered ? (Number(bookInfo[BOOK_DB_FIELDS.PRICE]) || 0) : 0;
      item.priceLabel = registered ? (Number(bookInfo[BOOK_DB_FIELDS.PRICE]) || 0).toLocaleString('ko-KR') + '원' : UNREGISTERED_LABEL;
    }
    return item;
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
 * 섹션 C — 과목별 교재 조회 (관/학년/과목 기준, 월 조건 없이 전체 월 대상)
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
    const gradeMatch = gradeMatches_(rowGrade, selectedGrade);
    const subjectMatch = !!bookInfo && normalize_(bookInfo[BOOK_DB_FIELDS.SUBJECT]) === selectedSubject;

    return campusMatch && gradeMatch && subjectMatch;
  });

  const showPrice = isPrivilegedRole_(auth.role);
  const joined = filtered.map(function (row) {
    const bookName = normalize_(row[CLASS_SETTING_FIELDS.BOOK_NAME]);
    const bookInfo = bookMap[bookName];

    const item = {
      bookName: bookName,
      subject: normalize_(bookInfo[BOOK_DB_FIELDS.SUBJECT]),
      publisher: normalize_(bookInfo[BOOK_DB_FIELDS.PUBLISHER]),
      isbn: normalize_(bookInfo[BOOK_DB_FIELDS.ISBN]),
      targetMonth: normalize_(row[CLASS_SETTING_FIELDS.TARGET_MONTH]),
      registered: true
    };
    if (showPrice) {
      item.price = Number(bookInfo[BOOK_DB_FIELDS.PRICE]) || 0;
      item.priceLabel = (Number(bookInfo[BOOK_DB_FIELDS.PRICE]) || 0).toLocaleString('ko-KR') + '원';
    }
    return item;
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

/**
 * "이 교재를 쓰는 반" 역조회 (전 권한 공통, 가격 정보 없음)
 */
function getClassesForBook(bookName) {
  requireRegisteredUser_();

  const name = normalize_(bookName);
  if (!name) return [];

  const settings = readSheetAsObjects_(SHEET_NAMES.CLASS_SETTING);
  const matched = settings
    .filter(function (row) { return normalize_(row[CLASS_SETTING_FIELDS.BOOK_NAME]) === name; })
    .map(function (row) {
      return {
        campus: normalize_(row[CLASS_SETTING_FIELDS.CAMPUS]),
        gradeClass: normalize_(row[CLASS_SETTING_FIELDS.GRADE_CLASS]),
        targetMonth: normalize_(row[CLASS_SETTING_FIELDS.TARGET_MONTH])
      };
    });

  return matched.sort(function (a, b) {
    return normalize_(a.campus).localeCompare(normalize_(b.campus), 'ko') ||
      normalize_(a.gradeClass).localeCompare(normalize_(b.gradeClass), 'ko');
  });
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
 * 저장된 접속 로그를 최신순으로 반환한다 (관리자 전용).
 */
function getAccessLogs(limit) {
  requireAdmin_();

  const props = PropertiesService.getScriptProperties();
  const maxCount = limit || 200;
  const keys = props.getKeys()
    .filter(function (k) { return k.indexOf(ACCESS_LOG_KEY_PREFIX) === 0; })
    .sort()
    .reverse()
    .slice(0, maxCount);

  return keys.map(function (key) {
    const entry = JSON.parse(props.getProperty(key));
    return {
      email: entry.email,
      role: entry.role,
      time: formatLogTime_(entry.time)
    };
  });
}

function formatLogTime_(isoString) {
  try {
    return Utilities.formatDate(new Date(isoString), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  } catch (e) {
    return isoString;
  }
}

/* ===================================================================
 * 미등록 교재 등록 요청 (직원·뷰어 → 관리자)
 * -------------------------------------------------------------------
 * 뷰어/직원은 스프레드시트 쓰기 권한이 없으므로(접속 로그와 동일한 이유),
 * 요청 "제출"은 시트가 아니라 PropertiesService에 저장한다.
 * 요청을 "승인"해 실제 교재DB에 반영하는 동작만 관리자가 수행하며,
 * 그 시점에만 스프레드시트 쓰기가 발생한다 (관리자는 편집자 권한 필요).
 * =================================================================== */

const BOOK_REQUEST_KEY_PREFIX = 'bookreq_';
const MAX_BOOK_REQUESTS = 1000;

function getAllBookRequests_() {
  const props = PropertiesService.getScriptProperties();
  const keys = props.getKeys().filter(function (k) { return k.indexOf(BOOK_REQUEST_KEY_PREFIX) === 0; });
  return keys
    .map(function (k) { return JSON.parse(props.getProperty(k)); })
    .sort(function (a, b) { return b.requestedAt.localeCompare(a.requestedAt); });
}

function pruneBookRequestsIfNeeded_(props) {
  const all = getAllBookRequests_();
  if (all.length <= MAX_BOOK_REQUESTS) return;

  const removable = all
    .filter(function (r) { return r.status !== '대기'; })
    .sort(function (a, b) { return a.requestedAt.localeCompare(b.requestedAt); });

  const overBy = all.length - MAX_BOOK_REQUESTS;
  removable.slice(0, overBy).forEach(function (r) {
    props.deleteProperty(r.id);
  });
}

/**
 * 미등록 교재 등록 요청 제출 (직원/뷰어)
 * location: { campus, grade, month } — 발견 위치(유형 A) 또는 사용 예정 반(유형 B, 선택)
 * requestType: 'A'(미등록 표시 기반) 또는 'B'(자유 입력/신규 요청). 기본값 'A'.
 */
function submitBookRequest(bookName, publisher, isbn, memo, location, requestType) {
  const auth = requireRegisteredUser_();

  const name = normalize_(bookName);
  if (!name) throw new Error('교재명을 입력해주세요.');

  const duplicate = getAllBookRequests_().some(function (r) {
    return r.bookName === name && r.status === '대기';
  });
  if (duplicate) {
    throw new Error('이미 요청됨 (처리 대기 중)');
  }

  const props = PropertiesService.getScriptProperties();
  const id = BOOK_REQUEST_KEY_PREFIX + Utilities.formatDate(new Date(), 'Etc/UTC', "yyyyMMdd'T'HHmmss'Z'") + '_' + Math.random().toString(36).slice(2, 8);

  const entry = {
    id: id,
    requestType: requestType === 'B' ? 'B' : 'A',
    bookName: name,
    publisher: normalize_(publisher),
    isbn: normalize_(isbn),
    memo: normalize_(memo),
    requesterEmail: auth.email,
    requestedAt: new Date().toISOString(),
    location: {
      campus: normalize_(location && location.campus),
      grade: normalize_(location && location.grade),
      month: normalize_(location && location.month)
    },
    status: '대기',
    rejectReason: '',
    processedAt: '',
    processedBy: ''
  };

  props.setProperty(id, JSON.stringify(entry));
  pruneBookRequestsIfNeeded_(props);
  return entry;
}

/**
 * 현재 접속자가 낸 요청 중 "대기 중"인 교재명 목록.
 * 화면에서 [등록 요청] 버튼 대신 "이미 요청됨" 배지를 보여줄지 판단하는 데 사용.
 * (같은 교재명이면 누가 요청했든 중복 방지 대상이므로 전체 대기 목록을 반환한다.)
 */
function getPendingRequestBookNames() {
  requireRegisteredUser_();
  return getAllBookRequests_()
    .filter(function (r) { return r.status === '대기'; })
    .map(function (r) { return r.bookName; });
}

/**
 * 현재 접속자 본인이 낸 요청 중 "반려"된 것들 (반려 안내 표시용)
 */
function getMyBookRequestNotices() {
  const auth = requireRegisteredUser_();
  return getAllBookRequests_().filter(function (r) {
    return r.requesterEmail === auth.email && r.status === '반려';
  });
}

/**
 * 요청함 전체 목록 (관리자 전용)
 */
function getBookRequests() {
  requireAdmin_();
  return getAllBookRequests_();
}

function getPendingRequestCount() {
  requireAdmin_();
  return getAllBookRequests_().filter(function (r) { return r.status === '대기'; }).length;
}

/**
 * 요청 승인 → 교재DB 시트에 새 행 추가 + 요청 상태 "처리됨" + 감사 로그 기록 (관리자 전용)
 */
function approveBookRequest(requestId, bookData) {
  const auth = requireAdmin_();

  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(normalize_(requestId));
  if (!raw) throw new Error('요청을 찾을 수 없습니다. 이미 처리되었을 수 있습니다.');
  const entry = JSON.parse(raw);
  if (entry.status !== '대기') throw new Error('이미 처리된 요청입니다.');

  const sheet = getSheetOrThrow_(SHEET_NAMES.BOOK_DB);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h).trim(); });

  const finalBookName = normalize_(bookData && bookData.bookName) || entry.bookName;
  const newRow = headers.map(function (h) {
    switch (h) {
      case BOOK_DB_FIELDS.SUBJECT: return normalize_(bookData && bookData.subject);
      case BOOK_DB_FIELDS.BOOK_NAME: return finalBookName;
      case BOOK_DB_FIELDS.PUBLISHER: return normalize_(bookData && bookData.publisher);
      case BOOK_DB_FIELDS.ISBN: return normalize_(bookData && bookData.isbn);
      case BOOK_DB_FIELDS.PRICE: return Number(bookData && bookData.price) || 0;
      case BOOK_DB_FIELDS.GRADE: return normalize_(bookData && bookData.grade);
      case BOOK_DB_FIELDS.TARGET_MONTH: return normalize_(bookData && bookData.targetMonth);
      default: return '';
    }
  });
  sheet.appendRow(newRow);

  entry.status = '처리됨';
  entry.processedAt = new Date().toISOString();
  entry.processedBy = auth.email;
  props.setProperty(entry.id, JSON.stringify(entry));

  logAudit_(auth.email, '등록', finalBookName, '등록 요청 승인 → 교재DB 등록 (요청ID: ' + entry.id + ')');
  return entry;
}

/**
 * 요청 반려 (관리자 전용)
 */
function rejectBookRequest(requestId, reason) {
  const auth = requireAdmin_();

  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(normalize_(requestId));
  if (!raw) throw new Error('요청을 찾을 수 없습니다. 이미 처리되었을 수 있습니다.');
  const entry = JSON.parse(raw);
  if (entry.status !== '대기') throw new Error('이미 처리된 요청입니다.');

  entry.status = '반려';
  entry.rejectReason = normalize_(reason);
  entry.processedAt = new Date().toISOString();
  entry.processedBy = auth.email;
  props.setProperty(entry.id, JSON.stringify(entry));

  logAudit_(auth.email, '요청반려', entry.bookName, '등록 요청 반려 (사유: ' + entry.rejectReason + ')');
  return entry;
}

/* ===================================================================
 * 권한 관리 화면 (관리자 전용)
 * 사용자권한 시트는 관리자만 쓰기 때문에(관리자는 편집자 권한 보유 전제)
 * 기존과 동일하게 시트에 직접 읽고/쓴다.
 * =================================================================== */

function getUserPermissions() {
  requireAdmin_();
  const rows = readSheetAsObjects_(SHEET_NAMES.USER_PERMISSION);
  return rows.map(function (row) {
    return {
      email: normalize_(row[USER_PERMISSION_FIELDS.EMAIL]),
      role: normalize_(row[USER_PERMISSION_FIELDS.ROLE]),
      name: normalize_(row[USER_PERMISSION_FIELDS.NAME]),
      registeredDate: normalize_(row[USER_PERMISSION_FIELDS.REGISTERED_DATE])
    };
  }).sort(function (a, b) { return a.email.localeCompare(b.email); });
}

function isValidRole_(role) {
  return role === ROLE.ADMIN || role === ROLE.STAFF || role === ROLE.VIEWER;
}

/**
 * 새 계정에 권한 부여 (이메일 입력 → 권한 지정)
 */
function addUserPermission(email, role, name) {
  const auth = requireAdmin_();
  const cleanEmail = normalize_(email).toLowerCase();
  if (!cleanEmail) throw new Error('이메일을 입력해주세요.');
  if (!isValidRole_(role)) throw new Error('올바른 권한 값이 아닙니다.');

  const existing = readSheetAsObjects_(SHEET_NAMES.USER_PERMISSION);
  const already = existing.some(function (row) {
    return normalize_(row[USER_PERMISSION_FIELDS.EMAIL]).toLowerCase() === cleanEmail;
  });
  if (already) throw new Error('이미 등록된 이메일입니다.');

  const sheet = getSheetOrThrow_(SHEET_NAMES.USER_PERMISSION);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h).trim(); });
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const newRow = headers.map(function (h) {
    switch (h) {
      case USER_PERMISSION_FIELDS.EMAIL: return cleanEmail;
      case USER_PERMISSION_FIELDS.ROLE: return role;
      case USER_PERMISSION_FIELDS.NAME: return normalize_(name);
      case USER_PERMISSION_FIELDS.REGISTERED_DATE: return today;
      default: return '';
    }
  });
  sheet.appendRow(newRow);

  logAudit_(auth.email, '권한부여', cleanEmail, '권한 "' + role + '" 부여');
  return true;
}

/**
 * 기존 계정 권한 변경
 */
function updateUserPermissionRole(email, newRole) {
  const auth = requireAdmin_();
  const cleanEmail = normalize_(email).toLowerCase();
  if (!isValidRole_(newRole)) throw new Error('올바른 권한 값이 아닙니다.');

  const sheet = getSheetOrThrow_(SHEET_NAMES.USER_PERMISSION);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(function (h) { return String(h).trim(); });
  const emailCol = headers.indexOf(USER_PERMISSION_FIELDS.EMAIL);
  const roleCol = headers.indexOf(USER_PERMISSION_FIELDS.ROLE);

  for (let r = 1; r < values.length; r++) {
    if (normalize_(values[r][emailCol]).toLowerCase() === cleanEmail) {
      const oldRole = values[r][roleCol];
      sheet.getRange(r + 1, roleCol + 1).setValue(newRole);
      logAudit_(auth.email, '권한변경', cleanEmail, '권한 "' + oldRole + '" → "' + newRole + '"');
      return true;
    }
  }
  throw new Error('등록되지 않은 이메일입니다.');
}

/**
 * 계정 권한 회수(삭제)
 */
function removeUserPermission(email) {
  const auth = requireAdmin_();
  const cleanEmail = normalize_(email).toLowerCase();
  if (cleanEmail === auth.email) {
    throw new Error('본인 계정은 회수할 수 없습니다.');
  }

  const sheet = getSheetOrThrow_(SHEET_NAMES.USER_PERMISSION);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(function (h) { return String(h).trim(); });
  const emailCol = headers.indexOf(USER_PERMISSION_FIELDS.EMAIL);

  for (let r = 1; r < values.length; r++) {
    if (normalize_(values[r][emailCol]).toLowerCase() === cleanEmail) {
      sheet.deleteRow(r + 1);
      logAudit_(auth.email, '권한회수', cleanEmail, '권한 회수(삭제)');
      return true;
    }
  }
  throw new Error('등록되지 않은 이메일입니다.');
}

/* ===================================================================
 * 변경 이력 / 감사 로그 (관리자 조회 전용)
 * 교재DB/권한 변경은 전부 관리자가 수행하는 동작이므로(관리자=편집자 권한 보유)
 * 시트에 직접 append 한다.
 * =================================================================== */

function logAudit_(actorEmail, action, target, summary) {
  try {
    const sheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.AUDIT_LOG);
    if (!sheet) return; // 감사로그 시트가 아직 없으면 조용히 건너뜀 (아래 초기 설정 함수로 생성)
    sheet.appendRow([
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      actorEmail,
      action,
      target,
      summary
    ]);
  } catch (e) {
    // 감사 로그 기록 실패가 원래 하려던 작업(등록/권한변경 등)을 막지 않도록 한다.
  }
}

/**
 * 감사 로그 조회 (관리자 전용). filter = { from: 'yyyy-MM-dd', to: 'yyyy-MM-dd', actor: '이메일 일부' }
 */
function getAuditLogs(filter) {
  requireAdmin_();

  const rows = readSheetAsObjects_(SHEET_NAMES.AUDIT_LOG);
  const fromDate = filter && filter.from ? normalize_(filter.from) : '';
  const toDate = filter && filter.to ? normalize_(filter.to) : '';
  const actorQuery = filter && filter.actor ? normalize_(filter.actor).toLowerCase() : '';

  const result = rows.filter(function (row) {
    const time = normalize_(row[AUDIT_LOG_FIELDS.TIME]);
    const rowActor = normalize_(row[AUDIT_LOG_FIELDS.ACTOR]).toLowerCase();
    if (fromDate && time < fromDate) return false;
    if (toDate && time > toDate + ' 23:59:59') return false;
    if (actorQuery && rowActor.indexOf(actorQuery) === -1) return false;
    return true;
  }).map(function (row) {
    return {
      time: normalize_(row[AUDIT_LOG_FIELDS.TIME]),
      actor: normalize_(row[AUDIT_LOG_FIELDS.ACTOR]),
      action: normalize_(row[AUDIT_LOG_FIELDS.ACTION]),
      target: normalize_(row[AUDIT_LOG_FIELDS.TARGET]),
      summary: normalize_(row[AUDIT_LOG_FIELDS.SUMMARY])
    };
  });

  result.reverse();
  return result;
}

/* ===================================================================
 * 업데이트 공지 (전 권한 공통 조회, 작성은 관리자가 시트에서 직접)
 * =================================================================== */

/**
 * 표시여부가 'Y'인 공지만 최신순으로 반환한다. 업데이트 탭이 아직 없으면 빈 배열.
 */
function getUpdates() {
  requireRegisteredUser_();

  let rows;
  try {
    rows = readSheetAsObjects_(SHEET_NAMES.UPDATE_LOG);
  } catch (e) {
    return [];
  }

  const visible = rows
    .filter(function (row) { return normalize_(row[UPDATE_LOG_FIELDS.VISIBLE]).toUpperCase() === 'Y'; })
    .map(function (row) {
      return {
        date: normalize_(row[UPDATE_LOG_FIELDS.DATE]),
        category: normalize_(row[UPDATE_LOG_FIELDS.CATEGORY]),
        content: normalize_(row[UPDATE_LOG_FIELDS.CONTENT])
      };
    });

  visible.sort(function (a, b) { return b.date.localeCompare(a.date); });
  return visible;
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

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  createSheetIfMissing_(ss, SHEET_NAMES.USER_PERMISSION, [
    USER_PERMISSION_FIELDS.EMAIL, USER_PERMISSION_FIELDS.ROLE,
    USER_PERMISSION_FIELDS.NAME, USER_PERMISSION_FIELDS.REGISTERED_DATE
  ], [
    ['admin@example.com', ROLE.ADMIN, '관리자 예시', today],
    ['staff@example.com', ROLE.STAFF, '직원 예시', today],
    ['viewer@example.com', ROLE.VIEWER, '뷰어 예시', today]
  ]);

  createSheetIfMissing_(ss, SHEET_NAMES.AUDIT_LOG, [
    AUDIT_LOG_FIELDS.TIME, AUDIT_LOG_FIELDS.ACTOR, AUDIT_LOG_FIELDS.ACTION,
    AUDIT_LOG_FIELDS.TARGET, AUDIT_LOG_FIELDS.SUMMARY
  ], []);

  createSheetIfMissing_(ss, SHEET_NAMES.UPDATE_LOG, [
    UPDATE_LOG_FIELDS.DATE, UPDATE_LOG_FIELDS.CATEGORY, UPDATE_LOG_FIELDS.CONTENT, UPDATE_LOG_FIELDS.VISIBLE
  ], [
    [today, '공지', '교재 대시보드에 업데이트 소식 게시판이 추가되었습니다.', 'Y']
  ]);
}

/**
 * 기존에 이미 사용 중이던 스프레드시트를 위한 1회성 마이그레이션.
 * - 사용자권한 탭에 "이름"/"등록일" 컬럼이 없으면 추가하고, 기존 행은 등록일을 오늘 날짜로 채워 넣는다.
 * - 감사로그 탭이 없으면 새로 만든다.
 * Apps Script 편집기 함수 목록에서 runFeatureUpgradeSetup 선택 후 "실행" 버튼으로 1회만 실행하면 된다.
 * 이미 반영되어 있으면 중복 실행해도 안전하다(변경할 것이 없으면 그냥 넘어간다).
 */
function runFeatureUpgradeSetup() {
  const ss = getSpreadsheet_();

  const permSheet = ss.getSheetByName(SHEET_NAMES.USER_PERMISSION);
  if (permSheet) {
    const lastCol = permSheet.getLastColumn();
    const headers = permSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    let nameCol = headers.indexOf(USER_PERMISSION_FIELDS.NAME);
    if (nameCol === -1) {
      nameCol = headers.length;
      permSheet.getRange(1, nameCol + 1).setValue(USER_PERMISSION_FIELDS.NAME).setFontWeight('bold');
      headers.push(USER_PERMISSION_FIELDS.NAME);
    }

    let dateCol = headers.indexOf(USER_PERMISSION_FIELDS.REGISTERED_DATE);
    if (dateCol === -1) {
      dateCol = headers.length;
      permSheet.getRange(1, dateCol + 1).setValue(USER_PERMISSION_FIELDS.REGISTERED_DATE).setFontWeight('bold');
      headers.push(USER_PERMISSION_FIELDS.REGISTERED_DATE);
    }

    const lastRow = permSheet.getLastRow();
    if (lastRow >= 2) {
      const dateRange = permSheet.getRange(2, dateCol + 1, lastRow - 1, 1);
      const dateValues = dateRange.getValues();
      const filled = dateValues.map(function (row) {
        return [row[0] === '' || row[0] === null ? today : row[0]];
      });
      dateRange.setValues(filled);
    }
  }

  if (!ss.getSheetByName(SHEET_NAMES.AUDIT_LOG)) {
    createSheetIfMissing_(ss, SHEET_NAMES.AUDIT_LOG, [
      AUDIT_LOG_FIELDS.TIME, AUDIT_LOG_FIELDS.ACTOR, AUDIT_LOG_FIELDS.ACTION,
      AUDIT_LOG_FIELDS.TARGET, AUDIT_LOG_FIELDS.SUMMARY
    ], []);
  }

  if (!ss.getSheetByName(SHEET_NAMES.UPDATE_LOG)) {
    createSheetIfMissing_(ss, SHEET_NAMES.UPDATE_LOG, [
      UPDATE_LOG_FIELDS.DATE, UPDATE_LOG_FIELDS.CATEGORY, UPDATE_LOG_FIELDS.CONTENT, UPDATE_LOG_FIELDS.VISIBLE
    ], [
      [Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'), '공지', '교재 대시보드에 업데이트 소식 게시판이 추가되었습니다.', 'Y']
    ]);
  }
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
