/**
 * 비욘드잉글리쉬 사직캠퍼스 · 상담 신청 → 구글 시트 저장 + 원장님 메일 알림
 *
 * 쓰는 법은 "연동-설정-방법.md" 를 봐 주세요.
 * 시트의 '상태' 칸을 [대기 → 확정 / 취소] 로 바꾸면,
 * '취소' 로 바꾼 시간은 홈페이지 달력에서 다시 예약할 수 있게 열려요.
 */

var SHEET_NAME = '상담신청';
var NOTIFY_EMAIL = 'debi92@naver.com';   // 알림 받을 메일 (비우면 메일 안 보냄)

var HEADERS = [
  '접수시각', '상태', '신청종류', '희망날짜', '희망시간',
  '보호자성함', '연락처', '학생이름', '학년',
  '영어유치원', '타학원경력', '파닉스', '현재실력', '레벨테스트', '문의내용'
];

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
      .requireValueInList(['대기', '확정', '완료', '취소'], true).build();
    sh.getRange('B2:B').setDataValidation(rule);
    // 날짜·시간·전화번호가 숫자로 바뀌지 않도록 글자 형식
    sh.getRange('D:E').setNumberFormat('@');
    sh.getRange('G:G').setNumberFormat('@');
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 이미 잡힌 예약 시간 목록 (이름·연락처는 내보내지 않아요) */
function bookedKeys_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 2, last - 1, 4).getDisplayValues(); // 상태, 종류, 날짜, 시간
  var today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  var out = [];
  rows.forEach(function (r) {
    var status = r[0], date = r[2], time = r[3];
    if (status === '취소' || !date || !time) return;
    if (date < today) return;
    out.push(date + ' ' + time);
  });
  return out;
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'booked') {
    return json_({ booked: bookedKeys_(getSheet_()) });
  }
  return json_({ ok: true, msg: '비욘드 사직 상담 접수 서버가 켜져 있어요.' });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000); // 같은 시간 동시 신청 방지
  try {
    var d = JSON.parse(e.postData.contents);
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
      clean(d.name), clean(d.phone), clean(d.student), clean(d.grade),
      clean(d.q_kinder), clean(d.q_academy), clean(d.q_phonics), clean(d.q_level), clean(d.q_test),
      clean(d.message)
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
