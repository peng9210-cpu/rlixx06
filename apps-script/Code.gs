// 갈릭스 라이브 아카이브 · 데이터 서버
// 화면은 GitHub Pages(SITE_URL)에서 뜨고, 이 스크립트는 구글 시트 데이터만 주고받아.
// 배포할 때는 "배포 관리 → 기존 배포 수정(연필) → 새 버전"으로 해야 주소가 안 바뀌어.

const SITE_URL = 'https://peng9210-cpu.github.io/rlixx06/';
const IDS = {
  shows: '1WOiUUa38pqI-mR2WCdC193IW-YqJfMVDcT8ZIKgy7HQ',
  sets: '1l14Eio-Jv3qfV_kS1HLt5by6rFXd3Iq-FFtOqSTEwr0',
  songs: '1cC56qtFB8HLMkzFU24AGsyOojEUfl2xY30_6oddhPyk',
  reviews: '1BR_FnAfqRMusLf3S9MWr-AQKjaM-jT7TrGZjD2kKeY4',
  visits: '1EGtSMS7iU4y4ayTCB0uy1xDMs2cXihOfNsKjrA3uCxA',
  messages: '1R2jb3gxSTo_SbRHIC0RMC7Kso8om-CttAJ0INTIsGms'
};
// 사이트에서 부를 수 있는 함수 목록
const API = {
  logVisit: logVisit, addMessage: addMessage, deleteMessage: deleteMessage,
  saveShow: saveShow, deleteShow: deleteShow, saveSong: saveSong, deleteSong: deleteSong,
  saveSet: saveSet, saveReview: saveReview, deleteReview: deleteReview
};

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.api === 'data') {
    if (p.fresh) CacheService.getScriptCache().remove('data');
    try { return json_({ok: true, result: getData()}); } catch (err) { return json_({ok: false, error: err.message}); }
  }
  if (p.api === 'visits') {
    try { return json_({ok: true, result: getVisits()}); } catch (err) { return json_({ok: false, error: err.message}); }
  }
  // 예전 주소로 들어온 사람은 새 사이트로 안내
  const html = "<!doctype html><meta name='viewport' content='width=device-width,initial-scale=1'>" +
    "<style>body{margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;background:#F6F1E9;font-family:system-ui,sans-serif;color:#211C17;text-align:center}" +
    "a{display:inline-block;margin-top:18px;background:#211C17;color:#FFF8EC;text-decoration:none;padding:14px 26px;border-radius:99px;font-weight:700}</style>" +
    "<div><div style='font-size:44px'>🧄</div><h2 style='margin:10px 0 6px'>갈릭스 아카이브가 이사했어</h2>" +
    "<div style='color:#8C8378;font-size:14px'>더 빨라진 새 주소로 와줘</div>" +
    "<a href='" + SITE_URL + "' target='_top'>새 아카이브로 가기 →</a></div>";
  return HtmlService.createHtmlOutput(html).setTitle('갈릭스 라이브 아카이브')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = API[body.fn];
    if (!fn) throw new Error('없는 기능이야: ' + body.fn);
    return json_({ok: true, result: fn.apply(null, body.args || [])});
  } catch (err) {
    return json_({ok: false, error: err.message});
  }
}

function logVisit(dev, first) {
  try {
    const sh = SpreadsheetApp.openById(IDS.visits).getSheets()[0];
    const now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
    sh.appendRow(["'" + now, String(dev || '').slice(0, 12), first ? 'O' : '']);
  } catch (err) {}
  return true;
}

function getVisits() {
  const sh = SpreadsheetApp.openById(IDS.visits).getSheets()[0];
  const rows = sh.getDataRange().getDisplayValues().slice(1).filter(function (r) { return r[1]; });
  const today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  const all = {}, td = {};
  rows.forEach(function (r) {
    all[r[1]] = true;
    if (String(r[0]).indexOf(today) === 0) td[r[1]] = true;
  });
  return {total: rows.length, people: Object.keys(all).length,
    todayPeople: Object.keys(td).length, last: rows.length ? rows[rows.length - 1][0] : ''};
}

function open_(key) {
  const ss = SpreadsheetApp.openById(IDS[key]);
  const sh = ss.getSheets()[0];
  const vals = sh.getDataRange().getValues();
  return {sh: sh, tz: ss.getSpreadsheetTimeZone(), vals: vals,
    head: vals[0].map(function (x) { return String(x).trim(); })};
}

function norm_(v, tz) {
  return (v instanceof Date) ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v).trim();
}
function txt_(s) { return s ? "'" + s : ''; }
function ex_(t, i) { return t.vals[i].join('').indexOf('(예시)') >= 0; }
function cell_(t, i, name) {
  const j = t.head.indexOf(name);
  return j < 0 ? '' : norm_(t.vals[i][j], t.tz);
}
function setCell_(t, i, name, v) {
  const j = t.head.indexOf(name);
  if (j >= 0) t.sh.getRange(i + 1, j + 1).setValue(v);
}
function rowFrom_(head, o) {
  return head.map(function (h) {
    const v = o[h];
    if (v === undefined || v === null) return '';
    if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v;
    return v;
  });
}
function removeWhere_(t, fn) {
  for (let i = t.vals.length - 1; i >= 1; i--) if (fn(i)) t.sh.deleteRow(i + 1);
}
function sameShow_(t, i, date, round) {
  return cell_(t, i, '날짜') === date && cell_(t, i, '회차') === round;
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally {
    CacheService.getScriptCache().remove('data');
    lock.releaseLock();
  }
}

function readRows_(key) {
  const t = open_(key);
  const disp = t.sh.getDataRange().getDisplayValues();
  const out = [];
  for (let i = 1; i < t.vals.length; i++) {
    const joined = disp[i].join('');
    if (joined.trim() === '' || joined.indexOf('(예시)') >= 0) continue;
    const o = {};
    t.head.forEach(function (h, j) {
      o[h] = (h === '시간') ? String(disp[i][j]).trim() : norm_(t.vals[i][j], t.tz);
    });
    out.push(o);
  }
  return out;
}

function getData() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('data');
  if (hit) return JSON.parse(hit);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const again = cache.get('data');
    if (again) return JSON.parse(again);
    const d = {shows: readRows_('shows'), sets: readRows_('sets'),
      songs: readRows_('songs'), reviews: readRows_('reviews'), messages: readRows_('messages')};
    try { cache.put('data', JSON.stringify(d), 21600); } catch (err) {}
    return d;
  } finally { lock.releaseLock(); }
}

function addMessage(name, text, time) {
  return withLock_(function () {
    text = String(text || '').trim().slice(0, 300);
    if (!text) throw new Error('한마디를 적어줘');
    const tm = /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(String(time)) ? String(time)
      : Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
    const t = open_('messages');
    t.sh.appendRow(rowFrom_(t.head, {'시간': "'" + tm, '이름': String(name || '').slice(0, 20), '내용': text}));
    return true;
  });
}

function deleteMessage(time, name, text) {
  return withLock_(function () {
    const t = open_('messages');
    const disp = t.sh.getDataRange().getDisplayValues();
    const ti = t.head.indexOf('시간'), ni = t.head.indexOf('이름'), ci = t.head.indexOf('내용');
    for (let i = disp.length - 1; i >= 1; i--) {
      if (disp[i][ti] === time && disp[i][ni] === (name || '') && disp[i][ci] === text) {
        t.sh.deleteRow(i + 1);
        break;
      }
    }
    return true;
  });
}

function saveShow(orig, s) {
  return withLock_(function () {
    const t = open_('shows');
    const round = s['회차'] || '';
    const nk = s['날짜'] + '|' + round;
    let at = -1;
    for (let i = 1; i < t.vals.length; i++) {
      if (ex_(t, i)) continue;
      const k = cell_(t, i, '날짜') + '|' + cell_(t, i, '회차');
      if (orig && k === orig) { if (at < 0) at = i; }
      else if (k === nk) throw new Error('같은 날짜·회차 공연이 이미 있어');
    }
    const o = Object.assign({}, s, {'날짜': txt_(s['날짜']), '시간': txt_(s['시간']), '예매 오픈일': txt_(s['예매 오픈일'])});
    const row = rowFrom_(t.head, o);
    if (at > 0) t.sh.getRange(at + 1, 1, 1, row.length).setValues([row]);
    else t.sh.appendRow(row);
    if (orig && orig !== nk) {
      const p = orig.split('|');
      ['sets', 'reviews'].forEach(function (name) {
        const u = open_(name);
        for (let i = 1; i < u.vals.length; i++) {
          if (sameShow_(u, i, p[0], p[1] || '')) {
            setCell_(u, i, '날짜', txt_(s['날짜']));
            setCell_(u, i, '회차', round);
          }
        }
      });
    }
    return true;
  });
}

function deleteShow(key) {
  return withLock_(function () {
    const p = key.split('|');
    const r = p[1] || '';
    const t = open_('shows');
    const hits = [];
    for (let i = 1; i < t.vals.length; i++) {
      if (!ex_(t, i) && sameShow_(t, i, p[0], r)) hits.push(i);
    }
    if (!hits.length) return true;
    t.sh.deleteRow(hits[hits.length - 1] + 1);
    if (hits.length === 1) {
      ['sets', 'reviews'].forEach(function (name) {
        const u = open_(name);
        removeWhere_(u, function (i) { return sameShow_(u, i, p[0], r); });
      });
    }
    return true;
  });
}

function saveSong(orig, s) {
  return withLock_(function () {
    const t = open_('songs');
    let at = -1;
    for (let i = 1; i < t.vals.length; i++) {
      if (ex_(t, i)) continue;
      const n = cell_(t, i, '곡명');
      if (orig && n === orig) { if (at < 0) at = i; }
      else if (n === s['곡명']) throw new Error('이미 있는 곡이야');
    }
    const row = rowFrom_(t.head, Object.assign({}, s, {'발매일': txt_(s['발매일'])}));
    if (at > 0) t.sh.getRange(at + 1, 1, 1, row.length).setValues([row]);
    else t.sh.appendRow(row);
    if (orig && orig !== s['곡명']) {
      const u = open_('sets');
      for (let i = 1; i < u.vals.length; i++) {
        if (cell_(u, i, '곡명') === orig) setCell_(u, i, '곡명', s['곡명']);
      }
    }
    return true;
  });
}

function deleteSong(name) {
  return withLock_(function () {
    const t = open_('songs');
    for (let i = t.vals.length - 1; i >= 1; i--) {
      if (!ex_(t, i) && cell_(t, i, '곡명') === name) { t.sh.deleteRow(i + 1); break; }
    }
    return true;
  });
}

function saveSet(date, round, items) {
  return withLock_(function () {
    const t = open_('sets');
    removeWhere_(t, function (i) { return sameShow_(t, i, date, round); });
    items.forEach(function (it, k) {
      t.sh.appendRow(rowFrom_(t.head, {'날짜': txt_(date), '회차': round, '순서': k + 1,
        '곡명': it.name, '앵콜': it.encore ? 'O' : 'X', '메모': it.memo || ''}));
    });
    const g = open_('songs');
    const have = {};
    for (let i = 1; i < g.vals.length; i++) have[cell_(g, i, '곡명')] = true;
    items.forEach(function (it) {
      if (!have[it.name]) { have[it.name] = true; g.sh.appendRow(rowFrom_(g.head, {'곡명': it.name})); }
    });
    return true;
  });
}

function saveReview(date, round, name, text, origName) {
  return withLock_(function () {
    if (!name) throw new Error('이름을 적어줘');
    const t = open_('reviews');
    removeWhere_(t, function (i) {
      const n = cell_(t, i, '이름');
      return sameShow_(t, i, date, round) && (n === name || (origName && n === origName));
    });
    t.sh.appendRow(rowFrom_(t.head, {'날짜': txt_(date), '회차': round, '이름': name, '후기': text || ''}));
    return true;
  });
}

function deleteReview(date, round, name) {
  return withLock_(function () {
    const t = open_('reviews');
    removeWhere_(t, function (i) { return sameShow_(t, i, date, round) && cell_(t, i, '이름') === name; });
    return true;
  });
}
