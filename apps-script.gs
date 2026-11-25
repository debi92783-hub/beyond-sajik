/**
 * 비욘드잉글리쉬 사직캠퍼스 · 상담 신청 → 구글 시트 저장 + 원장님 메일 알림
 *
 * 쓰는 법은 "연동-설정-방법.md" 를 봐 주세요.
 * 시트의 '상태' 칸을 [대기 → 확정 / 취소] 로 바꾸면,
 * '취소' 로 바꾼 시간은 홈페이지 달력에서 다시 예약할 수 있게 열려요.
 *
 * 2026-10-26 추가
 *  - '경로' 칸: 홈페이지 / 네이버. 학원 관리 화면 '상담 관리 > 예약 현황'에 경로별로 보여요.
 *  - 관리 화면용 목록(action=list)은 '시트 관리 키'가 맞을 때만 이름·연락처를 내보내요.
 *    키는 [프로젝트 설정 > 스크립트 속성]에 ADMIN_KEY 로 넣어요. 코드에는 적지 않아요.
 *  - 네이버 예약 알림 메일을 5분마다 읽어 시트에 넣는 importNaverMail (아래 설명 참고).
 */

var SHEET_NAME = '상담신청';
var NOTIFY_EMAIL = 'debi92@naver.com';   // 알림 받을 메일 (비우면 메일 안 보냄)

var HEADERS = [
  '접수시각', '상태', '신청종류', '희망날짜', '희망시간',
  '보호자성함', '연락처', '학생이름', '학년',
  '영어유치원', '타학원경력', '파닉스', '현재실력', '레벨테스트', '문의내용',
  '경로', '외부번호'
];
var STATUS_LIST = ['대기', '확정', '완료', '취소', '확인필요'];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    // 상태 칸 드롭다운
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(STATUS_LIST, true).build();
    sh.getRange('B2:B').setDataValidation(rule);
    // 날짜·시간·전화번호가 숫자로 바뀌지 않도록 글자 형식
    sh.getRange('D:E').setNumberFormat('@');
    sh.getRange('G:G').setNumberFormat('@');
  }
  // 예전에 만든 시트에도 새 칸 제목을 채워요
  if (sh.getRange(1, HEADERS.length).getValue() === '') {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  }
  return sh;
}

function adminOk_(key) {
  var k = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  return !!k && key === k;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 이미 잡힌 예약 목록 (이름·연락처는 내보내지 않아요. 날짜·시간·상태·종류·학년만) */
function bookedItems_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 2, last - 1, HEADERS.length - 1).getDisplayValues(); // 상태, 종류, 날짜, 시간, 성함, 연락처, 학생이름, 학년, … 경로
  var since = Utilities.formatDate(new Date(Date.now() - 62 * 864e5), 'Asia/Seoul', 'yyyy-MM-dd');
  var out = [];
  rows.forEach(function (r) {
    var status = r[0], date = r[2], time = r[3];
    if (status === '취소' || !date || !time) return;
    if (date < since) return;
    out.push({ date: date, time: time, status: status, kind: r[1], grade: r[7], src: r[14] || '홈페이지' });
  });
  return out;
}

/** 지금부터 예약할 수 없는 시간 목록 */
function bookedKeys_(sh) {
  var today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  return bookedItems_(sh).filter(function (b) { return b.date >= today; })
    .map(function (b) { return b.date + ' ' + b.time; });
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'feed') {
    return json_(feed_());
  }
  if (e && e.parameter && e.parameter.action === 'booked') {
    var sh = getSheet_();
    // booked: 홈페이지 달력용, items: 학원 관리 화면 공통 일정표용
    return json_({ booked: bookedKeys_(sh), items: bookedItems_(sh) });
  }
  return json_({ ok: true, msg: '비욘드 사직 상담 접수 서버가 켜져 있어요.' });
}

/** 관리 화면용 전체 목록 (관리 키가 맞을 때만, 최근 62일 + 앞으로) */
function adminRows_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, HEADERS.length).getDisplayValues();
  var since = Utilities.formatDate(new Date(Date.now() - 62 * 864e5), 'Asia/Seoul', 'yyyy-MM-dd');
  var out = [];
  rows.forEach(function (r, i) {
    if (r[3] && r[3] < since) return;
    out.push({ row: i + 2, created: r[0], status: r[1], kind: r[2], date: r[3], time: r[4], name: r[5], phone: r[6],
      student: r[7], grade: r[8], kinder: r[9], academy: r[10], phonics: r[11], level: r[12], test: r[13],
      message: r[14], src: r[15] || '홈페이지' });
  });
  return out;
}

function doPost(e) {
  // 업무 앱 게시판 첨부파일 → 구글 드라이브 (2026-11-25). 파일 올리는 동안 상담 신청이 막히지 않게 잠금 전에 처리해요.
  var up = null;
  try { up = JSON.parse(e.postData.contents); } catch (err0) {}
  if (up && up.action === 'upload') {
    try { return json_(uploadFile_(up)); } catch (err1) { return json_({ ok: false, error: 'drive', detail: String(err1 && err1.message || '').slice(0, 200) }); }
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(10000); // 같은 시간 동시 신청 방지
  try {
    var d = JSON.parse(e.postData.contents);
    // 학원 관리 화면 (관리 키 필요)
    if (d.action === 'list' || d.action === 'status') {
      if (!adminOk_(d.key)) return json_({ ok: false, error: 'key' });
      var ash = getSheet_();
      if (d.action === 'list') return json_({ ok: true, rows: adminRows_(ash) });
      var row = parseInt(d.row, 10);
      if (!(row >= 2 && row <= ash.getLastRow()) || STATUS_LIST.indexOf(d.status) === -1) return json_({ ok: false, error: 'bad' });
      ash.getRange(row, 2).setValue(d.status);
      return json_({ ok: true });
    }
    if (d.website) return json_({ ok: true }); // 스팸 봇
    if (!d.name || !d.phone || !d.grade || !d.date || !d.time) {
      return json_({ ok: false, error: 'missing' });
    }
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d.date) || !/^\d{2}:\d{2}$/.test(d.time)) {
      return json_({ ok: false, error: 'bad_format' });
    }

    var sh = getSheet_();
    if (bookedKeys_(sh).indexOf(d.date + ' ' + d.time) !== -1) {
      return json_({ ok: false, error: 'taken' });
    }

    var clean = function (v) {
      v = String(v || '').slice(0, 1000);
      return /^[=+\-@]/.test(v) ? "'" + v : v; // 수식 주입 방지
    };

    sh.appendRow([
      Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm'),
      '대기',
      clean(d.purpose), d.date, d.time,
      clean(d.name), "'" + clean(d.phone), clean(d.student), clean(d.grade),
      clean(d.q_kinder), clean(d.q_academy), clean(d.q_phonics), clean(d.q_level), clean(d.q_test),
      clean(d.message), '홈페이지', ''
    ]);

    if (NOTIFY_EMAIL) {
      var body = [
        '새 ' + d.purpose + ' 신청이 들어왔어요.',
        '',
        '희망: ' + d.date + ' ' + d.time,
        '보호자: ' + d.name + ' / ' + d.phone,
        '학생: ' + (d.student || '-') + ' / ' + d.grade,
        '영어유치원: ' + (d.q_kinder || '-'),
        '타학원 경력: ' + (d.q_academy || '-'),
        '파닉스: ' + (d.q_phonics || '-'),
        '현재 실력: ' + (d.q_level || '-'),
        '레벨테스트: ' + (d.q_test || '-'),
        '문의: ' + (d.message || '-'),
        '',
        '시트에서 상태를 "확정"으로 바꿔 관리해 주세요.',
        SpreadsheetApp.getActiveSpreadsheet().getUrl()
      ].join('\n');
      MailApp.sendEmail(NOTIFY_EMAIL, '[상담신청] ' + d.date + ' ' + d.time + ' ' + d.name + ' 님', body);
    }

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: 'server' });
  } finally {
    lock.releaseLock();
  }
}

/** 설정 확인용: 편집기에서 이 함수를 한 번 실행하면 시트가 만들어지고 권한 허용 창이 떠요 */
function setup() {
  getSheet_();
  if (NOTIFY_EMAIL) MailApp.getRemainingDailyQuota();
}

/* ============================================================
   네이버 예약 → 시트 자동 등록
   ------------------------------------------------------------
   네이버 예약이 들어오면 오는 '알림 메일'을 이 구글 계정 Gmail에서 읽어요.
   1) 네이버 예약 알림 메일이 이 Gmail로 오게 해요 (네이버 메일이면 자동 전달 설정).
   2) Gmail에서 그 메일에 '네이버예약' 라벨이 자동으로 붙도록 필터를 만들어요.
   3) 편집기에서 setupNaverTrigger 를 한 번 실행하면 5분마다 importNaverMail 이 돌아요.
   ※ 아래 parseNaverMail_ 의 글자 규칙은 견본이에요. 실제 알림 메일 한 통을 받아
     문구(예약일시·예약자·예약번호 표기)에 맞게 고친 뒤 켜요.
     읽지 못한 메일은 버리지 않고 상태 '확인필요'로 넣어 두니 놓치는 예약은 없어요.
   ============================================================ */
var NAVER_QUERY = 'label:네이버예약 is:unread';

function setupNaverTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'importNaverMail') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('importNaverMail').timeBased().everyMinutes(5).create();
}

function importNaverMail() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    var sh = getSheet_();
    GmailApp.search(NAVER_QUERY, 0, 30).forEach(function (thread) {
      thread.getMessages().forEach(function (m) {
        if (!m.isUnread()) return;
        var p = parseNaverMail_(m.getSubject(), m.getPlainBody());
        var at = p.no ? findByExt_(sh, p.no) : 0;
        if (p.cancel) {
          if (at) sh.getRange(at, 2).setValue('취소');
        } else if (!at) {
          sh.appendRow([
            Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm'),
            p.date && p.time ? '확정' : '확인필요',
            p.kind || '상담', p.date || '', p.time || '',
            p.name || '', p.phone || '', '', '',
            '', '', '', '', '',
            p.date && p.time ? (p.memo || '') : '메일 제목: ' + m.getSubject(),
            '네이버', p.no || ''
          ]);
        }
        m.markRead();
      });
    });
  } finally {
    lock.releaseLock();
  }
}

function findByExt_(sh, no) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var col = sh.getRange(2, HEADERS.length, last - 1, 1).getDisplayValues();
  for (var i = 0; i < col.length; i++) if (col[i][0] === String(no)) return i + 2;
  return 0;
}

/** 견본 규칙: 실제 메일을 받으면 여기 글자만 맞춰 고쳐요 */
function parseNaverMail_(subject, body) {
  var text = subject + '\n' + body;
  var line = function (labels) {
    for (var i = 0; i < labels.length; i++) {
      var m = text.match(new RegExp(labels[i] + '\\s*[:：]?\\s*([^\\n]+)'));
      if (m) return m[1].trim();
    }
    return '';
  };
  var out = { cancel: /취소/.test(subject), no: line(['예약번호']), name: line(['예약자명', '예약자', '이용자']),
    memo: line(['요청사항', '요청 사항']), kind: line(['상품명', '상품']) };
  var ph = text.match(/01[016789]-?\d{3,4}-?\d{4}/);
  if (ph) out.phone = ph[0];
  var when = line(['이용일시', '예약일시', '일시']) || text;
  var d = when.match(/(\d{4})[.\-년]\s*(\d{1,2})[.\-월]\s*(\d{1,2})/);
  if (d) out.date = d[1] + '-' + ('0' + d[2]).slice(-2) + '-' + ('0' + d[3]).slice(-2);
  var t = when.match(/(오전|오후)?\s*(\d{1,2})\s*[:시]\s*(\d{2})?/);
  if (d && t) {
    var h = parseInt(t[2], 10);
    if (t[1] === '오후' && h < 12) h += 12;
    if (t[1] === '오전' && h === 12) h = 0;
    out.time = ('0' + h).slice(-2) + ':' + (t[3] || '00');
  }
  return out;
}


/* ============================================================
 * 2026-11-23 추가: 홈페이지 '블로그·인스타 최신 글' (?action=feed)
 *  - 블로그: 네이버 블로그 RSS 를 읽어요. 따로 설정할 것 없어요.
 *  - 인스타: [프로젝트 설정 > 스크립트 속성]에 IG_TOKEN(인스타그램 장기 토큰)을 넣으면 보여요.
 *    토큰이 없으면 인스타 칸은 '인스타에서 보기' 칸으로 보여요.
 *    토큰은 60일마다 연장해야 해서, refreshIgToken 을 '시간 기반 트리거(매주 1번)'로 걸어 두세요.
 *  - 30분 동안 결과를 저장해 두고 다시 써요(네이버·인스타에 너무 자주 묻지 않게).
 * ============================================================ */
var BLOG_RSS = 'https://rss.blog.naver.com/beyondsajik.xml';

function feed_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('feed');
  if (hit) return JSON.parse(hit);
  var out = { blog: blogFeed_(), insta: instaFeed_(), at: new Date().toISOString() };
  cache.put('feed', JSON.stringify(out), 1800);
  return out;
}

function blogFeed_() {
  try {
    var xml = UrlFetchApp.fetch(BLOG_RSS, { muteHttpExceptions: true }).getContentText('UTF-8');
    var items = XmlService.parse(xml).getRootElement().getChild('channel').getChildren('item');
    return items.slice(0, 6).map(function (it) {
      var desc = it.getChildText('description') || '';
      var img = desc.match(/<img[^>]+src=["']([^"']+)["']/i);
      return {
        title: it.getChildText('title'),
        link: (it.getChildText('link') || '').replace(/^http:/, 'https:'),
        date: new Date(it.getChildText('pubDate')).toISOString(),
        thumb: img ? img[1].replace(/^http:/, 'https:') : ''
      };
    });
  } catch (err) { return []; }
}

function instaFeed_() {
  var token = PropertiesService.getScriptProperties().getProperty('IG_TOKEN');
  if (!token) return [];
  try {
    var url = 'https://graph.instagram.com/me/media?fields=caption,media_type,media_url,thumbnail_url,permalink,timestamp&limit=6&access_token=' + encodeURIComponent(token);
    var res = JSON.parse(UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText());
    return (res.data || []).map(function (m) {
      return {
        caption: (m.caption || '').split('\n')[0].slice(0, 60),
        link: m.permalink,
        date: m.timestamp,
        thumb: m.media_type === 'VIDEO' ? (m.thumbnail_url || '') : (m.media_url || '')
      };
    });
  } catch (err) { return []; }
}

/** 인스타 토큰 연장 (매주 1번 트리거로 실행) */
function refreshIgToken() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('IG_TOKEN');
  if (!token) return;
  var res = JSON.parse(UrlFetchApp.fetch('https://graph.instagram.com/refresh_access_token?grant_type=ig_refresh_token&access_token=' + encodeURIComponent(token), { muteHttpExceptions: true }).getContentText());
  if (res.access_token) props.setProperty('IG_TOKEN', res.access_token);
}


/* ============================================================
   업무 앱 게시판 첨부파일 → 구글 드라이브 (2026-11-25 추가)
   ------------------------------------------------------------
   - 업무 앱(beyond-sajik.web.app)에 로그인한 선생님·원장님만 올릴 수 있어요.
     앱이 보내 준 로그인 표(idToken)를 Firebase에 물어봐서 학원 계정이 맞는지 확인해요.
   - 파일은 이 구글 계정 드라이브의 '비욘드 사직 업무앱 첨부파일' 폴더에
     게시판 분류별 하위 폴더(예: 게시판_전달사항)로 저장돼요.
   - 링크를 아는 사람만 볼 수 있게(보기 전용) 열어 두고, 링크는 로그인해야 보이는 게시판 글에만 붙어요.
   - 앱에서 글이나 첨부를 지워도 드라이브 파일은 지우지 않아요(원장님 승인 규칙).
     정리는 원장님이 드라이브에서 직접 해 주세요.
   - 처음 한 번: 편집기에서 setupDrive 를 실행해 드라이브 권한을 허용해 주세요.
   ============================================================ */
var FIREBASE_API_KEY = 'AIzaSyDux5xS0ywPVlgNPyeiOIOQU2iXvjFNKo0'; // 업무 앱 화면에도 들어 있는 공개용 키예요(비밀 아님)
var STAFF_DOMAIN = '@beyond-sajik.app';
var UPLOAD_ROOT = '비욘드 사직 업무앱 첨부파일';
var UPLOAD_MAX = 30 * 1024 * 1024;

function staffOk_(idToken) {
  if (!idToken) return null;
  var res = UrlFetchApp.fetch('https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + FIREBASE_API_KEY, {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ idToken: idToken })
  });
  if (res.getResponseCode() !== 200) return null;
  var u = (JSON.parse(res.getContentText()).users || [])[0];
  if (!u || u.disabled || !u.email || u.email.slice(-STAFF_DOMAIN.length) !== STAFF_DOMAIN) return null;
  return u.email.slice(0, u.email.indexOf('.g'));
}

function uploadFolder_(sub) {
  var props = PropertiesService.getScriptProperties();
  var root = null, id = props.getProperty('UPLOAD_FOLDER_ID');
  if (id) { try { root = DriveApp.getFolderById(id); if (root.isTrashed()) root = null; } catch (e) { root = null; } }
  if (!root) {
    var it = DriveApp.getFoldersByName(UPLOAD_ROOT);
    root = it.hasNext() ? it.next() : DriveApp.createFolder(UPLOAD_ROOT);
    props.setProperty('UPLOAD_FOLDER_ID', root.getId());
  }
  sub = String(sub || '').replace(/[\\/:*?"<>|]/g, '').slice(0, 40);
  if (!sub) return root;
  var s = root.getFoldersByName(sub);
  return s.hasNext() ? s.next() : root.createFolder(sub);
}

function uploadFile_(d) {
  var who = staffOk_(d.idToken);
  if (!who) return { ok: false, error: 'login' };
  if (!d.data || !d.name) return { ok: false, error: 'missing' };
  var bytes = Utilities.base64Decode(d.data);
  if (bytes.length > UPLOAD_MAX) return { ok: false, error: 'big' };
  var name = String(d.name).replace(/[\\/:*?"<>|]/g, '_').slice(0, 120);
  var blob = Utilities.newBlob(bytes, String(d.type || 'application/octet-stream'), name);
  var f = uploadFolder_(d.folder).createFile(blob);
  f.setDescription('업무 앱 게시판 첨부 · 올린 사람: ' + who + ' · ' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm') + (d.title ? ' · 글: ' + String(d.title).slice(0, 80) : ''));
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return { ok: true, id: f.getId(), url: 'https://drive.google.com/file/d/' + f.getId() + '/view' };
}

/** 처음 한 번 실행: 드라이브 권한 허용 창이 뜨고, 첨부 폴더가 만들어져요 */
function setupDrive() {
  var f = uploadFolder_('');
  Logger.log('첨부 폴더: ' + f.getUrl());
}
