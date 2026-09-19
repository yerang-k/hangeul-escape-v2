/**
 * 한글날 방탈출 — 학년별 순위 집계
 * 구글폼 "응답 스프레드시트"에 붙이는 Apps Script.
 *
 * [설치]
 *  1) 게임 완주 기록이 쌓이는 구글폼의 "응답 스프레드시트"를 연다.
 *  2) 상단 메뉴: 확장 프로그램 → Apps Script.
 *  3) 기본 코드를 지우고 이 파일 내용을 통째로 붙여넣고 저장(💾).
 *  4) 스프레드시트로 돌아와 새로고침(F5) → 상단에 '🏆 순위' 메뉴가 생김.
 *  5) '🏆 순위 → 학년별 순위 계산' 클릭. 처음엔 권한 승인 창이 뜨니 허용.
 *  → '1학년 순위 / 2학년 순위 / 3학년 순위' 시트가 만들어지고, 상위 10명이 강조됨.
 *  언제든 다시 눌러 새로 집계 가능(행사 중 중간 집계도 OK).
 */

// ── 설정(필요하면 여기만 바꾸면 됨) ──
var TOP_N = 10;            // 학년별로 강조할 상위 인원 수
var RESPONSE_SHEET = '';   // 응답 시트 이름. 비우면 자동으로 폼 응답 시트를 찾음.
// 열 제목 후보. 폼 질문 제목이 아래와 다르면 단어를 추가하면 됨(못 찾으면 데이터 모양으로 자동 탐지).
var HDR = {
  sid:  ['학번', 'student', 'id'],
  name: ['이름', '성명', 'name'],
  dur:  ['소요', '시간', '기록', 'duration', '걸린']
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏆 순위')
    .addItem('학년별 순위 계산', 'buildRanking')
    .addToUi();
}

/** "12분 34초" → 754(초). 숫자만 있으면 초로 간주. */
function parseDurationSec(v) {
  if (v === null || v === '') return null;
  if (typeof v === 'number' && isFinite(v)) return Math.round(v);
  var s = String(v);
  var m = s.match(/(\d+)\s*분/);
  var sec = s.match(/(\d+)\s*초/);
  if (m || sec) return (m ? parseInt(m[1], 10) : 0) * 60 + (sec ? parseInt(sec[1], 10) : 0);
  var n = s.match(/\d+/);
  return n ? parseInt(n[0], 10) : null;
}

/** 학번(예 10305) → {grade:'1', cls:'03', num:'05'} */
function parseSid(v) {
  var d = String(v).replace(/\D/g, '');
  if (d.length === 5) return { grade: d.charAt(0), cls: d.substr(1, 2), num: d.substr(3, 2), sid: d };
  if (d.length >= 1) return { grade: d.charAt(0), cls: '', num: '', sid: d };
  return null;
}

function fmtSec(sec) {
  return Math.floor(sec / 60) + '분 ' + (sec % 60) + '초';
}

/** 열 위치 찾기: 제목 후보 우선, 못 찾으면 데이터 패턴으로 자동 탐지. */
function findCols(headers, rows) {
  function byHeader(cands) {
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i]).toLowerCase();
      for (var j = 0; j < cands.length; j++) {
        if (h.indexOf(cands[j].toLowerCase()) >= 0) return i;
      }
    }
    return -1;
  }
  function detect(test) {
    var best = -1, bestHit = 0;
    for (var c = 0; c < headers.length; c++) {
      var hit = 0, seen = 0;
      for (var r = 0; r < rows.length && seen < 30; r++) {
        var val = rows[r][c];
        if (val === '' || val === null) continue;
        seen++;
        if (test(val)) hit++;
      }
      if (seen > 0 && hit > bestHit) { bestHit = hit; best = c; }
    }
    return best;
  }
  var sid = byHeader(HDR.sid);
  var name = byHeader(HDR.name);
  var dur = byHeader(HDR.dur);
  if (sid < 0) sid = detect(function (v) { return /^\D*\d{5}\D*$/.test(String(v)); });          // 5자리 학번
  if (dur < 0) dur = detect(function (v) { return /\d+\s*분/.test(String(v)); });                 // "N분"
  return { sid: sid, name: name, dur: dur };
}

function buildRanking() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // 1) 응답 시트 찾기
  var sh = RESPONSE_SHEET ? ss.getSheetByName(RESPONSE_SHEET) : null;
  if (!sh) {
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var a1 = String(sheets[i].getRange(1, 1).getValue());
      if (a1.indexOf('타임스탬프') >= 0 || a1.toLowerCase().indexOf('timestamp') >= 0) { sh = sheets[i]; break; }
    }
    if (!sh) sh = sheets[0];
  }
  var values = sh.getDataRange().getValues();
  if (values.length < 2) { ui.alert('응답 데이터가 없습니다. (제출된 완주 기록이 아직 없음)'); return; }

  var headers = values[0];
  var rows = values.slice(1);
  var col = findCols(headers, rows);
  if (col.sid < 0 || col.dur < 0) {
    ui.alert('열을 자동으로 찾지 못했습니다.\n\n찾은 열 제목:\n' + headers.join(' | ') +
      '\n\n스크립트 상단 HDR 설정에서 학번/소요시간 열 제목 단어를 추가한 뒤 다시 실행하세요.');
    return;
  }

  // 2) 학번별 최고기록만 남기기(중복·재도전 시 가장 빠른 기록)
  var best = {}; // sid -> {grade,cls,num,name,sec,ts}
  for (var r = 0; r < rows.length; r++) {
    var row = rows[r];
    var p = parseSid(row[col.sid]);
    var sec = parseDurationSec(row[col.dur]);
    if (!p || sec === null) continue;
    var name = (col.name >= 0 ? String(row[col.name] || '') : '').trim();
    var ts = row[0] instanceof Date ? row[0].getTime() : 0; // A열 타임스탬프
    var cur = best[p.sid];
    if (!cur || sec < cur.sec) best[p.sid] = { grade: p.grade, cls: p.cls, num: p.num, name: name, sec: sec, ts: ts };
  }

  // 3) 학년별로 묶어 정렬(빠른 순, 동점이면 먼저 완주한 순)
  var byGrade = {};
  Object.keys(best).forEach(function (sid) {
    var b = best[sid];
    (byGrade[b.grade] = byGrade[b.grade] || []).push(b);
  });

  var grades = Object.keys(byGrade).sort();
  if (!grades.length) { ui.alert('집계할 유효한 기록이 없습니다. 학번/소요시간 열 값을 확인하세요.'); return; }

  grades.forEach(function (g) {
    var list = byGrade[g].sort(function (a, b) { return a.sec - b.sec || a.ts - b.ts; });
    var out = ss.getSheetByName(g + '학년 순위') || ss.insertSheet(g + '학년 순위');
    out.clear();
    var head = ['순위', '이름', '학번', '반', '번호', '소요시간', '기록(초)', '완주시각'];
    var data = [head];
    list.forEach(function (b, idx) {
      data.push([
        idx + 1, b.name, b.sid, b.cls, b.num, fmtSec(b.sec), b.sec,
        b.ts ? new Date(b.ts) : ''
      ]);
    });
    out.getRange(1, 1, data.length, head.length).setValues(data);
    out.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#241a10').setFontColor('#e0b357');
    // 상위 TOP_N 강조
    var hi = Math.min(TOP_N, list.length);
    if (hi > 0) out.getRange(2, 1, hi, head.length).setBackground('#fff4d6').setFontWeight('bold');
    out.setFrozenRows(1);
    out.autoResizeColumns(1, head.length);
    // 상위 10명 아래에 구분선(빨간 상단 테두리)
    if (list.length > TOP_N) out.getRange(2 + TOP_N, 1, 1, head.length).setBorder(true, null, null, null, null, null, '#e06a52', SpreadsheetApp.BorderStyle.SOLID_THICK);
  });

  ui.alert('집계 완료 ✅\n\n' + grades.map(function (g) { return g + '학년: ' + byGrade[g].length + '명'; }).join('\n') +
    '\n\n각 "N학년 순위" 시트에서 상위 ' + TOP_N + '명이 노란색으로 강조됐습니다.');
}

/** 자체 점검(에디터에서 이 함수만 실행하면 로직 검증) */
function demo() {
  if (parseDurationSec('12분 34초') !== 754) throw '분초 파싱 오류';
  if (parseDurationSec('0분 45초') !== 45) throw '초 파싱 오류';
  if (parseDurationSec('90') !== 90) throw '숫자 파싱 오류';
  var p = parseSid('10305');
  if (!(p.grade === '1' && p.cls === '03' && p.num === '05')) throw '학번 파싱 오류';
  if (parseSid('31048').grade !== '3') throw '학년 추출 오류';
  Logger.log('자체 점검 통과');
}
