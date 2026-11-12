/**
 * 김정수필름 테스트 스크리닝 — 예매·설문 시트 연결
 * 구글 시트 > 확장 프로그램 > Apps Script 에 이 파일 전체를 붙여넣고
 * setup() 을 한 번 실행한 뒤, 배포 > 새 배포 > 웹 앱 (실행: 나, 액세스: 모든 사용자) 으로 배포하세요.
 */

const SHEET_BOOK = '예매';
const SHEET_INVITE = '초대명단';
const SHEET_SET = '설정';
const SHEET_SURVEY = '관람후기';

const INVITES = ['김정수','이하림','김유엘','김준형','편수창','심애정','이채훈','이승현','최결','김강태','김조은','박재한','범수'];

function setup() {
  const ss = SpreadsheetApp.getActive();
  const mk = (name, header) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(header);
    sh.setFrozenRows(1);
    return sh;
  };
  mk(SHEET_BOOK, ['예매번호','시각','구분','좌석','이름','연락처','알게 된 경로','기대하는 점','개인정보 동의','취소']);
  const inv = mk(SHEET_INVITE, ['이름','좌석','선택 시각']);
  if (inv.getLastRow() === 1) INVITES.forEach(n => inv.appendRow([n, '', '']));
  const set = mk(SHEET_SET, ['항목','값','설명']);
  if (set.getLastRow() === 1) {
    set.appendRow(['일반 선착순 좌석 수', 13, '26석 - 초대 인원. 바뀌면 숫자만 고치세요']);
    set.appendRow(['예매 열기', '예', '"아니오"로 바꾸면 예매가 닫힙니다']);
    set.appendRow(['막아 둘 좌석', '', '쉼표로 구분 (예: A1,A2) — 스태프·장비 자리']);
  }
  mk(SHEET_SURVEY, ['시각','이름','골디락스존 프로젝트 별점','골디락스존 한마디','잡아먹히거나 먼저 씹어라 별점','잡아먹히거나 한마디','기름도둑 별점','기름도둑 한마디','가장 좋았던 작품','기억에 남는 장면','헷갈린 부분','추천 점수(0-10)','세계관 다음 이야기','하고 싶은 말','후기 인용 동의']);
}

function settings_() {
  const v = SpreadsheetApp.getActive().getSheetByName(SHEET_SET).getRange(2, 2, 3, 1).getValues();
  return {
    cap: Number(v[0][0]) || 0,
    open: String(v[1][0]).trim() !== '아니오',
    blocked: String(v[2][0] || '').split(',').map(s => s.trim().toUpperCase()).filter(Boolean)
  };
}

function bookings_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_BOOK);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 10).getValues().filter(r => !r[9]);
}

function invites_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_INVITE);
  const n = sh.getLastRow() - 1;
  return n < 1 ? [] : sh.getRange(2, 1, n, 2).getValues().map((r, i) => ({ row: i + 2, name: String(r[0]).trim(), seat: String(r[1]).trim() }));
}

function status_() {
  const s = settings_();
  const b = bookings_();
  const inv = invites_();
  const taken = b.map(r => String(r[3])).concat(inv.filter(i => i.seat).map(i => i.seat)).concat(s.blocked);
  const pub = b.filter(r => r[2] === '일반').length;
  return { ok: true, open: s.open, taken: taken, publicLeft: Math.max(0, s.cap - pub), cap: s.cap, invitePending: inv.filter(i => !i.seat).length };
}

const norm_ = s => String(s || '').replace(/\s/g, '').replace(/(배우|님|감독|조감독)$/g, '');
const out_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function doGet(e) {
  return out_(status_());
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    const ss = SpreadsheetApp.getActive();
    if (d.type === 'survey') {
      ss.getSheetByName(SHEET_SURVEY).appendRow([new Date(), d.name || '', d.r1, d.c1, d.r2, d.c2, d.r3, d.c3, d.best, d.scene, d.confuse, d.nps, d.next, d.msg, d.quote ? '예' : '아니오']);
      return out_({ ok: true });
    }
    const st = status_();
    if (!st.open) return out_({ ok: false, code: 'closed', msg: '예매가 마감되었습니다.' });
    const seat = String(d.seat || '').toUpperCase();
    if (!/^[A-Z]\d{1,2}$/.test(seat)) return out_({ ok: false, code: 'seat', msg: '좌석을 다시 골라 주세요.' });
    if (st.taken.indexOf(seat) >= 0) return out_({ ok: false, code: 'taken', msg: '방금 다른 분이 고른 좌석입니다. 다른 좌석을 골라 주세요.', taken: st.taken });

    if (d.type === 'invite') {
      const me = invites_().find(i => norm_(i.name) === norm_(d.name));
      if (!me) return out_({ ok: false, code: 'noinvite', msg: '초대 명단에서 이름을 찾지 못했습니다. 감독에게 문의해 주세요.' });
      if (me.seat) return out_({ ok: false, code: 'already', msg: '이미 ' + me.seat + ' 좌석을 고르셨습니다.', seat: me.seat });
      const sh = ss.getSheetByName(SHEET_INVITE);
      sh.getRange(me.row, 2, 1, 2).setValues([[seat, new Date()]]);
      return out_({ ok: true, seat: seat, name: me.name, no: 'INV-' + seat });
    }

    if (d.type === 'book') {
      if (st.publicLeft <= 0) return out_({ ok: false, code: 'soldout', msg: '선착순 좌석이 모두 찼습니다.' });
      const phone = String(d.phone || '').replace(/\D/g, '');
      if (!d.name || phone.length < 10) return out_({ ok: false, code: 'info', msg: '이름과 연락처를 확인해 주세요.' });
      if (bookings_().some(r => String(r[5]).replace(/\D/g, '') === phone)) return out_({ ok: false, code: 'dup', msg: '이 연락처로 이미 예매되어 있습니다.' });
      const no = 'KJF-' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'MMdd') + '-' + seat;
      ss.getSheetByName(SHEET_BOOK).appendRow([no, new Date(), '일반', seat, d.name, "'" + d.phone, d.from || '', d.expect || '', d.agree ? '예' : '아니오', '']);
      return out_({ ok: true, seat: seat, name: d.name, no: no });
    }
    return out_({ ok: false, code: 'type', msg: '잘못된 요청입니다.' });
  } catch (err) {
    return out_({ ok: false, code: 'error', msg: '잠시 후 다시 시도해 주세요.' });
  } finally {
    lock.releaseLock();
  }
}
