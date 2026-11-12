/**
 * 김정수필름 테스트 스크리닝 — 예매·설문 시트 연결
 * 구글 시트 > 확장 프로그램 > Apps Script 에 이 파일 전체를 붙여넣고
 * setup() 을 한 번 실행한 뒤, 배포 > 새 배포 > 웹 앱 (실행: 나, 액세스: 모든 사용자) 으로 배포하세요.
 */

const SHEET_BOOK = '예매';
const SHEET_INVITE = '초대명단';
const SHEET_SET = '설정';
const SHEET_SURVEY = '관람후기';

const INVITES = ['김정수','이하림','김유엘','김준형','편수창','심애정','이채훈','이승현','최결','김강태','김조은','박재한','박범수'];
/* 라이카시네마 2관 26석 (예매 페이지 LAYOUT 과 같아야 함) */
const ROWS = { A: 8, B: 9, C: 9 };
const SURVEY_HEADER = ['시각','구분','예매번호/이름','골디락스존 프로젝트 별점','골디락스존 한마디','잡아먹히거나 먼저 씹어라 별점','잡아먹히거나 한마디','기름도둑 별점','기름도둑 한마디','가장 좋았던 작품','기억에 남는 장면','헷갈린 부분','추천 점수(0-10)','세계관 연결','하고 싶은 말','후기 인용 동의','적은 이름'];

function setup() {
  const ss = SpreadsheetApp.getActive();
  const mk = (name, header) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(header);
    sh.setFrozenRows(1);
    return sh;
  };
  const b = mk(SHEET_BOOK, ['예매번호','시각','구분','좌석','이름','연락처','알게 된 경로','기대하는 점','개인정보 동의','취소']);
  b.getRange('F:F').setNumberFormat('@');
  const inv = mk(SHEET_INVITE, ['이름','좌석','선택 시각']);
  if (inv.getLastRow() === 1) INVITES.forEach(n => inv.appendRow([n, '', '']));
  const set = mk(SHEET_SET, ['항목','값','설명']);
  if (set.getLastRow() === 1) {
    set.appendRow(['일반 선착순 좌석 수', 13, '26석 - 초대 인원. 바뀌면 숫자만 고치세요']);
    set.appendRow(['예매 열기', '예', '"아니오"로 바꾸면 예매가 닫힙니다']);
    set.appendRow(['막아 둘 좌석', '', '쉼표로 구분 (예: A1,A2) — 스태프·장비 자리']);
  }
  if (set.getLastRow() < 5) set.appendRow(['설문 열기', '예', '"아니오"로 바꾸면 관람 후 설문이 닫힙니다']);
  const sv = mk(SHEET_SURVEY, SURVEY_HEADER);
  const head = sv.getRange(1, 1, 1, SURVEY_HEADER.length).getValues()[0].join('|');
  if (head !== SURVEY_HEADER.join('|')) {
    /* 예전 설문 양식(시험 응답)이면 비우고 새 머리줄로 바꿈 */
    if (sv.getLastRow() > 1) sv.deleteRows(2, sv.getLastRow() - 1);
    sv.getRange(1, 1, 1, sv.getMaxColumns()).clearContent();
    sv.getRange(1, 1, 1, SURVEY_HEADER.length).setValues([SURVEY_HEADER]);
  }
}

/* ---------- 읽기 ---------- */
function settings_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SET);
  const v = sh.getRange(2, 2, 4, 1).getValues();
  const total = Object.keys(ROWS).reduce((a, k) => a + ROWS[k], 0);
  const blocked = String(v[2][0] || '').split(/[,\s]+/).map(s => s.trim().toUpperCase()).filter(Boolean);
  return {
    /* 일반 좌석 = 26석 - 초대 인원 - 막아 둔 좌석 (자동) */
    cap: Math.max(0, total - invites_().length - blocked.length),
    open: String(v[1][0]).trim() !== '아니오',
    blocked: blocked,
    survey: String(v[3][0]).trim() !== '아니오'
  };
}
function bookRows_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_BOOK);
  const n = sh.getLastRow() - 1;
  const v = n < 1 ? [] : sh.getRange(2, 1, n, 10).getValues();
  return v.map((r, i) => ({ row: i + 2, no: String(r[0]).trim().toUpperCase(), kind: String(r[2]), seat: String(r[3]).trim().toUpperCase(), name: String(r[4]), phone: digits_(r[5]), cancel: String(r[9]).trim() }));
}
function invites_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_INVITE);
  const n = sh.getLastRow() - 1;
  return n < 1 ? [] : sh.getRange(2, 1, n, 2).getValues().map((r, i) => ({ row: i + 2, name: String(r[0]).trim(), seat: String(r[1]).trim().toUpperCase() })).filter(i => i.name);
}
function status_(exceptSeat) {
  const s = settings_();
  const live = bookRows_().filter(r => !r.cancel && r.seat);
  const inv = invites_();
  const taken = live.map(r => r.seat).concat(inv.filter(i => i.seat).map(i => i.seat)).concat(s.blocked)
    .filter((x, i, a) => x && a.indexOf(x) === i && x !== exceptSeat);
  const pub = live.filter(r => r.kind === '일반').length;
  return { ok: true, open: s.open, survey: s.survey, taken: taken, publicLeft: Math.max(0, s.cap - pub), cap: s.cap, invitePending: inv.filter(i => !i.seat).length };
}

/* ---------- 도우미 ---------- */
const digits_ = s => String(s || '').replace(/\D/g, '');
const norm_ = s => String(s || '').replace(/\s/g, '').replace(/(조감독님|감독님|배우님|조감독|감독|배우|님)$/g, '');
const clean_ = (s, max) => { let t = String(s == null ? '' : s).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max || 300); if (/^[=+\-@]/.test(t)) t = "'" + t; return t; };
const validSeat_ = s => { const m = /^([A-Z])(\d{1,2})$/.exec(s); return !!m && ROWS[m[1]] && +m[2] >= 1 && +m[2] <= ROWS[m[1]]; };
const validPhone_ = p => /^010\d{8}$/.test(p);
const fmtP_ = p => String(p || '').replace(/^(\d{3})(\d{4})(\d{4})$/, '$1-$2-$3');
const NOTIFY_TO = 'littledirectorsoo@gmail.com';
function notify_(subject, lines) {
  try {
    const st = status_();
    MailApp.sendEmail(NOTIFY_TO, '[스크리닝] ' + subject, lines.concat(['', '남은 일반석 ' + st.publicLeft + '석 · 초대 미선택 ' + st.invitePending + '명', 'https://docs.google.com/spreadsheets/d/' + SpreadsheetApp.getActive().getId()]).join('\n'));
  } catch (e) {}
}
function testNotify() { notify_('알림 시험', ['예매 알림이 이렇게 옵니다.']); }
const out_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const newNo_ = () => { const ex = bookRows_().map(r => r.no); let no; do { no = 'KJF-' + String(Math.floor(1000 + Math.random() * 9000)); } while (ex.indexOf(no) >= 0); return no; };

/* 설문 자격 확인: 일반 = 예매번호+연락처, 초대 = 이름(좌석을 고른 사람) */
function whoAmI_(d) {
  if (d.kind === 'invite') {
    const me = invites_().find(i => norm_(i.name) === norm_(d.name));
    if (!me) return { ok: false, msg: '초대 명단에서 이름을 찾지 못했습니다.' };
    return { ok: true, kind: '초대', key: me.name, name: me.name };
  }
  const no = String(d.no || '').trim().toUpperCase(), phone = digits_(d.phone);
  const r = bookRows_().find(x => x.no === no && x.phone === phone && !x.cancel);
  if (!r) return { ok: false, msg: '예매번호와 연락처가 맞는 예매를 찾지 못했습니다. 예매 완료 화면의 번호를 확인해 주세요.' };
  return { ok: true, kind: '일반', key: r.no, name: r.name };
}
function surveyed_(key) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SURVEY);
  const n = sh.getLastRow() - 1;
  if (n < 1) return false;
  return sh.getRange(2, 3, n, 1).getValues().some(r => String(r[0]) === key);
}

/* ---------- 요청 ---------- */
function doGet(e) {
  try { return out_(status_()); } catch (err) { return out_({ ok: false, code: 'error', msg: '잠시 후 다시 시도해 주세요.' }); }
}

function doPost(e) {
  let d;
  try { d = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return out_({ ok: false, code: 'bad', msg: '잘못된 요청입니다.' }); }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return out_({ ok: false, code: 'busy', msg: '지금 예매가 몰려 있습니다. 잠시 후 다시 눌러 주세요.' });
  try {
    const ss = SpreadsheetApp.getActive();
    const s = settings_();

    if (d.type === 'auth' || d.type === 'survey') {
      if (!s.survey) return out_({ ok: false, code: 'surveyclosed', msg: '설문이 아직 열리지 않았거나 마감되었습니다.' });
      const me = whoAmI_(d);
      if (!me.ok) return out_({ ok: false, code: 'noauth', msg: me.msg });
      if (surveyed_(me.key)) return out_({ ok: false, code: 'done', msg: '이미 설문을 보내 주셨습니다. 감사합니다.' });
      if (d.type === 'auth') return out_({ ok: true, name: me.name });
      const r = n => { const v = Number(n); return v >= 1 && v <= 5 ? v : ''; };
      const nps = d.nps === '' || d.nps == null ? '' : Math.max(0, Math.min(10, Number(d.nps)));
      ss.getSheetByName(SHEET_SURVEY).appendRow([new Date(), me.kind, me.key, r(d.r1), clean_(d.c1), r(d.r2), clean_(d.c2), r(d.r3), clean_(d.c3), clean_(d.best, 40), clean_(d.scene), clean_(d.confuse), nps, clean_(d.next, 40), clean_(d.msg, 1000), d.quote ? '예' : '아니오', clean_(d.name, 40)]);
      return out_({ ok: true });
    }

    if (d.type === 'cancel' && d.name && !d.no) {
      const me = invites_().find(i => norm_(i.name) === norm_(d.name));
      if (!me) return out_({ ok: false, code: 'noinvite', msg: '초대 명단에서 이름을 찾지 못했습니다.' });
      if (!me.seat) return out_({ ok: false, code: 'already', msg: '선택된 초대석이 없습니다.' });
      ss.getSheetByName(SHEET_INVITE).getRange(me.row, 2, 1, 2).setValues([['', '']]);
      notify_('초대석 취소 · ' + me.seat + ' ' + me.name, ['초대 손님: ' + me.name, '좌석 ' + me.seat + ' 이(가) 비워졌습니다.']);
      return out_({ ok: true, seat: me.seat });
    }
    if (d.type === 'cancel') {
      const no = String(d.no || '').trim().toUpperCase(), phone = digits_(d.phone);
      const r = bookRows_().find(x => x.no === no && x.phone === phone);
      if (!r) return out_({ ok: false, code: 'nofind', msg: '예매번호와 연락처가 맞는 예매를 찾지 못했습니다.' });
      if (r.cancel) return out_({ ok: false, code: 'already', msg: '이미 취소된 예매입니다.' });
      ss.getSheetByName(SHEET_BOOK).getRange(r.row, 10).setValue('관객 취소 ' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'MM/dd HH:mm'));
      notify_('예매 취소 · ' + r.seat + ' ' + r.name, ['취소: ' + r.name + ' (' + r.no + ')', '연락처: ' + fmtP_(r.phone), '좌석 ' + r.seat + ' 이(가) 다시 열렸습니다.']);
      return out_({ ok: true, seat: r.seat });
    }

    if (!s.open) return out_({ ok: false, code: 'closed', msg: '예매가 마감되었습니다.' });
    const seat = String(d.seat || '').trim().toUpperCase();
    if (!validSeat_(seat)) return out_({ ok: false, code: 'seat', msg: '없는 좌석입니다. 좌석을 다시 골라 주세요.' });

    if (d.type === 'invite') {
      const me = invites_().find(i => norm_(i.name) === norm_(d.name));
      if (!me) return out_({ ok: false, code: 'noinvite', msg: '초대 명단에서 이름을 찾지 못했습니다. 이름만(예: 홍길동) 적어 주시거나 감독에게 문의해 주세요.' });
      const st = status_(me.seat);
      if (st.taken.indexOf(seat) >= 0) return out_({ ok: false, code: 'taken', msg: '방금 다른 분이 고른 좌석입니다. 다른 좌석을 골라 주세요.', taken: status_().taken });
      ss.getSheetByName(SHEET_INVITE).getRange(me.row, 2, 1, 2).setValues([[seat, new Date()]]);
      notify_('초대석 ' + (me.seat && me.seat !== seat ? '변경' : '선택') + ' · ' + seat + ' ' + me.name, ['초대 손님: ' + me.name, '좌석: ' + (me.seat && me.seat !== seat ? me.seat + ' → ' : '') + seat]);
      return out_({ ok: true, seat: seat, name: me.name, no: '초대석', changed: me.seat && me.seat !== seat ? me.seat : '' });
    }

    if (d.type === 'book') {
      const phone = digits_(d.phone), name = clean_(d.name, 40);
      if (!name) return out_({ ok: false, code: 'info', msg: '이름을 적어 주세요.' });
      if (!validPhone_(phone)) return out_({ ok: false, code: 'info', msg: '휴대폰 번호는 010으로 시작하는 11자리로 적어 주세요. (예: 010-1234-5678)' });
      if (!d.agree) return out_({ ok: false, code: 'agree', msg: '개인정보 수집·이용에 동의해 주셔야 예매할 수 있습니다.' });
      const mine = bookRows_().find(r => r.phone === phone && !r.cancel);
      const st = status_(mine ? mine.seat : '');
      if (st.taken.indexOf(seat) >= 0) return out_({ ok: false, code: 'taken', msg: '방금 다른 분이 고른 좌석입니다. 다른 좌석을 골라 주세요.', taken: status_().taken });
      const sh = ss.getSheetByName(SHEET_BOOK);
      if (mine) {
        sh.getRange(mine.row, 4).setValue(seat);
        notify_('좌석 변경 · ' + seat + ' ' + mine.name, ['예매자: ' + mine.name + ' (' + mine.no + ')', '연락처: ' + fmtP_(mine.phone), '좌석: ' + mine.seat + ' → ' + seat]);
        return out_({ ok: true, seat: seat, name: mine.name, no: mine.no, changed: mine.seat !== seat ? mine.seat : '' });
      }
      if (st.publicLeft <= 0) return out_({ ok: false, code: 'soldout', msg: '선착순 좌석이 모두 찼습니다. 취소석이 생기면 다시 열립니다.' });
      const no = newNo_();
      sh.appendRow([no, new Date(), '일반', seat, name, phone.replace(/^(\d{3})(\d{3,4})(\d{4})$/, '$1-$2-$3'), clean_(d.from, 40), clean_(d.expect), '예', '']);
      notify_('새 예매 · ' + seat + ' ' + name, ['이름: ' + name, '연락처: ' + fmtP_(phone), '좌석: ' + seat, '예매번호: ' + no, '알게 된 경로: ' + clean_(d.from, 40)]);
      return out_({ ok: true, seat: seat, name: name, no: no });
    }
    return out_({ ok: false, code: 'type', msg: '잘못된 요청입니다.' });
  } catch (err) {
    return out_({ ok: false, code: 'error', msg: '잠시 후 다시 시도해 주세요.' });
  } finally {
    lock.releaseLock();
  }
}

/* 상영회 전 초기화: 예매·초대 좌석·설문 응답을 모두 비움 (편집기에서 직접 실행할 때만) */
function resetAll() {
  const ss = SpreadsheetApp.getActive();
  [SHEET_BOOK, SHEET_SURVEY].forEach(n => { const sh = ss.getSheetByName(n); if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1); });
  /* 초대명단은 그대로 둠 */
  ss.getSheetByName(SHEET_SET).getRange(2, 1, 1, 3).setValues([['일반 선착순 좌석 수', '자동', '26석 - 초대명단 인원 - 막아 둔 좌석. 초대명단에 이름을 넣고 빼면 저절로 바뀝니다']]);
}
