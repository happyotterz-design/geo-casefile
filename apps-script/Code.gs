/**
 * 태종대 지오 사건파일 · 학생 결과 수집 창구
 *
 * 사용법
 * 1. 새 구글 시트를 만들고 [확장 프로그램 → Apps Script]를 연다.
 * 2. 이 코드를 Code.gs에 통째로 붙여넣고 저장한다.
 * 3. 위쪽 함수 목록에서 setup 을 고르고 [실행] → 권한 허용.
 *    (시트에 「사진」「결과」「방명록」 탭과 드라이브 폴더가 만들어진다)
 *    방명록에서 숨기고 싶은 글은 「숨기기(X)」 칸에 X 를 적으면 수첩에서 안 보인다.
 *
 * 코드를 고친 뒤에는 [배포 → 배포 관리 → 연필 → 버전: 새 버전 → 배포] 해야 반영된다(주소는 그대로).
 * 4. [배포 → 새 배포 → 유형: 웹 앱]
 *    - 다음 사용자 인증 정보로 실행: 나
 *    - 액세스 권한이 있는 사용자: 모든 사용자
 *    → 배포 후 나오는 웹 앱 URL(https://script.google.com/macros/s/…/exec)을 수사 수첩 설정에 넣는다.
 */

const FOLDER_NAME = '태종대 지오 사건파일 · 학생 사진';
const PHOTO_HEAD = ['받은 시각', '수사팀', '거점', '선 개수', '사진 보기', '파일 ID'];
const GUEST_HEAD = ['받은 시각', '수사팀', '종류', '글', '위도', '경도', '사진 보기', '파일 ID', '숨기기(X)', '글 ID', '삭제 키'];
const RESULT_HEAD = ['받은 시각', '수사팀', '보낸 시점', '총점', '등급', '푼 핀', '힌트 사용', '정답 보기',
                     '지목한 범인', '고른 동기', '고른 증거', '보고서 점수', '보고서 제출 횟수', '관찰한 둔덕', '결론', '공청회 메모'];

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet_(ss, '사진', PHOTO_HEAD);
  sheet_(ss, '결과', RESULT_HEAD);
  sheet_(ss, '방명록', GUEST_HEAD);
  const folder = folder_();
  Logger.log('준비 완료. 사진 폴더: ' + folder.getUrl());
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'guestbook') {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('방명록');
    if (!sh || sh.getLastRow() < 2) return json_({ ok: true, items: [] });
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, GUEST_HEAD.length).getValues();
    const items = rows
      .filter(function (r) { return String(r[8]).trim().toUpperCase().indexOf('X') !== 0; })
      .map(function (r) {
        return { at: r[0] instanceof Date ? r[0].getTime() : r[0], team: r[1], kind: r[2], text: r[3],
                 lat: r[4] === '' ? null : Number(r[4]), lng: r[5] === '' ? null : Number(r[5]), img: r[7] || '', id: r[9] || '' };
      })
      .reverse().slice(0, 150);
    return json_({ ok: true, items: items });
  }
  return json_({ ok: true, msg: '태종대 지오 사건파일 수집 창구가 켜져 있어요.' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const now = new Date();
    const team = clean_(d.team) || '이름 없는 수사팀';

    if (d.type === 'photo') {
      const b64 = String(d.image || '').replace(/^data:image\/\w+;base64,/, '');
      if (!b64) return json_({ ok: false, error: '사진이 비어 있어요' });
      const stamp = Utilities.formatDate(now, 'Asia/Seoul', 'MMdd_HHmm');
      const name = fileSafe_(team) + '_' + fileSafe_(d.site) + '_' + stamp + '.jpg';
      const file = folder_().createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', name));
      sheet_(ss, '사진', PHOTO_HEAD).appendRow([now, team, clean_(d.siteName || d.site), Number(d.lines) || 0, file.getUrl(), file.getId()]);
      return json_({ ok: true, url: file.getUrl() });
    }

    if (d.type === 'guest') {
      const text = clean_(d.text, 300);
      if (!text) return json_({ ok: false, error: '글이 비어 있어요' });
      let url = '', id = '';
      const g64 = String(d.image || '').replace(/^data:image\/\w+;base64,/, '');
      if (g64) {
        const gname = '방명록_' + fileSafe_(team) + '_' + Utilities.formatDate(now, 'Asia/Seoul', 'MMdd_HHmmss') + '.jpg';
        const gf = folder_().createFile(Utilities.newBlob(Utilities.base64Decode(g64), 'image/jpeg', gname));
        try { gf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (err) { /* 공유가 막힌 계정이면 사진만 안 보인다 */ }
        url = gf.getUrl(); id = gf.getId();
      }
      const lat = Number(d.lat), lng = Number(d.lng);
      sheet_(ss, '방명록', GUEST_HEAD).appendRow([now, team, d.kind === '질문' ? '질문' : '방명록', text,
        d.lat != null && isFinite(lat) ? lat : '', d.lng != null && isFinite(lng) ? lng : '', url, id, '',
        clean_(d.gid, 40), clean_(d.key, 60)]);
      return json_({ ok: true });
    }

    if (d.type === 'guestDel') {
      const sh = ss.getSheetByName('방명록');
      if (!sh || sh.getLastRow() < 2) return json_({ ok: false, error: '글을 찾지 못했어요' });
      const rows = sh.getRange(2, 1, sh.getLastRow() - 1, GUEST_HEAD.length).getValues();
      for (let i = 0; i < rows.length; i++) {
        if (String(rows[i][9]) === String(d.id)) {
          if (!d.key || String(rows[i][10]) !== String(d.key)) return json_({ ok: false, error: '내가 쓴 글만 지울 수 있어요' });
          sh.getRange(i + 2, 9).setValue('X(학생 삭제)');
          return json_({ ok: true });
        }
      }
      return json_({ ok: false, error: '글을 찾지 못했어요' });
    }

    if (d.type === 'result') {
      sheet_(ss, '결과', RESULT_HEAD).appendRow([
        now, team, clean_(d.kind), d.score, clean_(d.grade), clean_(d.solved), d.hints, d.reveals,
        clean_(d.suspect), clean_(d.motive), (d.evidence || []).map(clean_).join(', '),
        clean_(d.report), d.reportAttempts || '', clean_(d.observed), clean_(d.conclusion, 2000), clean_(d.hearing, 2000)
      ]);
      return json_({ ok: true });
    }

    return json_({ ok: false, error: '알 수 없는 요청이에요' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function sheet_(ss, name, head) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(head);
    sh.setFrozenRows(1);
  }
  if (sh.getLastColumn() < head.length) sh.getRange(1, 1, 1, head.length).setValues([head]);
  sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#e9efeb');
  return sh;
}

function folder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* 폴더가 지워졌으면 새로 만든다 */ }
  }
  const f = DriveApp.createFolder(FOLDER_NAME);
  props.setProperty('FOLDER_ID', f.getId());
  return f;
}

function clean_(v, max) {
  // 시트 수식으로 해석되지 않도록 앞의 =, +, -, @ 를 막고 길이를 제한한다
  let s = String(v == null ? '' : v).trim();
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s.slice(0, max || 200);
}

function fileSafe_(v) {
  return String(v == null ? '' : v).replace(/[\\/:*?"<>|]/g, ' ').trim().slice(0, 40) || '이름없음';
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
