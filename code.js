const SS_ID = '1OsE77i0X-5IIRUP83KuTlz-dtyQQIcg_Tai56nVq8MY';
const SHEET_ITEMS = 'items_db';
const SHEET_STUDENTS = 'students';
const SHEET_RESULTS = 'results';
const SHEET_SESSIONS = 'sessions';
const TEACHER_PIN = '8478';

// ★ RolesBuilders 出力シート名（追加：挙動は変えず、読むだけ）
const SHEET_ROLES_JP = 'items_roles_jp';
const SHEET_ROLES_EN = 'items_roles_en';

// ★【追加】ダミーワード用シート（A列 word / 1行目ヘッダーOK）
const SHEET_DUMMY_WORDS = 'dummy_words';

// ★【追加】unit別コイン管理（studentsに coin_u1..u8 を作る前提）
const MAX_UNITS = 8;

// ★ ユニットごとのボス画像（Drive のファイルID）
const BOSS_IMAGE_IDS = {
  '1': '19UnMJmAJ1QzUtUdgAcElHtO9kY8wthF7', // Unit1 じこしょうかいボス
  // 2以降を増やすならここに '2': 'xxxx', みたいに追加
};

// ---------- Router ----------
function doGet(e) {
  const p = (e && e.parameter) || {};

  // ★ 画像専用ルート
  // 例: .../exec?bossImg=1&unit=1
  if (p.bossImg === '1') {
    const unit = p.unit ? String(p.unit) : '1';
    return serveBossImage_(unit);
  }

  // ✅【追加】先生ログインページ（admin=1なしで入れる）
  // 例: .../exec?page=teacher_login
  if (p.page === 'teacher_login') {
    return HtmlService.createHtmlOutputFromFile('teacher_login')
      .setTitle('語順でGO！ | 教師ログイン');
  }

  // 先生ダッシュボード（既存どおり）
  if (p.admin === '1') {
    return HtmlService.createHtmlOutputFromFile('teacher')
      .setTitle('語順でGO！ | 先生ダッシュボード');
  }

  // ゲーム用サブページ
  if (p.page && p.t) {
    let fileName = null;
    switch (String(p.page)) {
      case 'game_practice':
        fileName = 'game_practice';
        break;
      case 'game_challenge':
        fileName = 'game_challenge';
        break;
      case 'game_focus':
        fileName = 'game_focus';
        break;
      case 'review_menu':
        fileName = 'review_menu';
        break;
      case 'game_stamp':
        fileName = 'game_stamp';
        break;
      case 'game_boss':
        fileName = 'game_boss';
        break;
      case 'game_menu':        // ★ 復習メニュー(スタンプ/ボス選択用)
        fileName = 'game_menu';
        break;
      case 'game_review':      // ★ 1文復習
        fileName = 'game_review';
        break;

      // ✅【ここだけ追加】チャレンジ2択メニュー
      case 'challenge_menu':
        fileName = 'challenge_menu';
        break;
    }
    if (fileName) {
      const tplSub = HtmlService.createTemplateFromFile(fileName);
      tplSub.token = String(p.t);
      // ★ unit / id / from も必ず渡す
      tplSub.unit = p.unit ? String(p.unit) : '';
      tplSub.id   = p.id   ? String(p.id)   : '';
      tplSub.from = p.from ? String(p.from) : '';

      return tplSub.evaluate().setTitle('語順でGO！');
    }
  }

  // トークン付きアクセス（URL は今まで通り ?t=xxx）
  // Home.html（ホーム画面）を表示
  if (p.t) {
    const tpl = HtmlService.createTemplateFromFile('Home');
    tpl.token = String(p.t);
    return tpl.evaluate().setTitle('語順でGO！ | ホーム');
  }

  // デフォルト：ログイン画面
  return HtmlService.createHtmlOutputFromFile('login')
    .setTitle('語順でGO！ | ログイン');
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// ✅【追加】白画面防止：/exec のURLをJS側に返す
function getExecBaseUrl() {
  return ScriptApp.getService().getUrl();
}

// ---------- Helpers ----------
function book() {
  return SpreadsheetApp.openById(SS_ID);
}
function sh(name) {
  return book().getSheetByName(name);
}
function ensureSheet(name, headers) {
  let ws = sh(name);
  if (!ws) ws = book().insertSheet(name);
  if (ws.getLastRow() === 0 && headers && headers.length) {
    ws.appendRow(headers);
  }
  return ws;
}
function nowDate() {
  return new Date();
}

// ★ 数字文字を半角にそろえる（全角だった場合の保険）
function normalizeDigits(str) {
  return String(str || '').replace(/[０-９]/g, c =>
    String.fromCharCode(c.charCodeAt(0) - 0xFEE0)
  );
}

// ★ フルかなから「名」だけ取り出す
function extractGivenKana_(fullKana) {
  const s = String(fullKana || '').trim();
  if (!s) return '';
  const parts = s.split(/\s+/);
  if (parts.length >= 2) {
    return parts[parts.length - 1];
  }
  return parts[0];
}

// ★ ふりがな（ひらがな）→ ヘボン式っぽいローマ字（名前用）に変換
function kanaToHepburnName_(kana) {
  kana = String(kana || '').trim();
  if (!kana) return '';

  const digraphMap = {
    'きゃ': 'kya', 'きゅ': 'kyu', 'きょ': 'kyo',
    'ぎゃ': 'gya', 'ぎゅ': 'gyu', 'ぎょ': 'gyo',
    'しゃ': 'sha', 'しゅ': 'shu', 'しょ': 'sho',
    'じゃ': 'ja',  'じゅ': 'ju',  'じょ': 'jo',
    'ちゃ': 'cha', 'ちゅ': 'chu', 'ちょ': 'cho',
    'にゃ': 'nya', 'にゅ': 'nyu', 'にょ': 'nyo',
    'ひゃ': 'hya', 'ひゅ': 'hyu', 'ひょ': 'hyo',
    'びゃ': 'bya', 'びゅ': 'byu', 'びょ': 'byo',
    'ぴゃ': 'pya', 'ぴゅ': 'pyu', 'ぴょ': 'pyo',
    'みゃ': 'mya', 'みゅ': 'myu', 'みょ': 'myo',
    'りゃ': 'rya', 'りゅ': 'ryu', 'りょ': 'ryo'
  };

  const monoMap = {
    'あ': 'a',  'い': 'i',  'う': 'u',  'え': 'e',  'お': 'o',
    'か': 'ka', 'き': 'ki', 'く': 'ku', 'け': 'ke', 'こ': 'ko',
    'さ': 'sa', 'し': 'shi','す': 'su', 'せ': 'se', 'そ': 'so',
    'た': 'ta', 'ち': 'chi','つ': 'tsu','て': 'te', 'と': 'to',
    'な': 'na', 'に': 'ni', 'ぬ': 'nu', 'ね': 'ne', 'の': 'no',
    'は': 'ha', 'ひ': 'hi', 'ふ': 'fu', 'へ': 'he', 'ほ': 'ho',
    'ま': 'ma', 'み': 'mi', 'む': 'mu', 'め': 'me', 'も': 'mo',
    'や': 'ya', 'ゆ': 'yu', 'よ': 'yo',
    'ら': 'ra', 'り': 'ri', 'る': 'ru', 'れ': 're', 'ろ': 'ro',
    'わ': 'wa', 'を': 'o',  'ん': 'n',
    'が': 'ga', 'ぎ': 'gi', 'ぐ': 'gu', 'げ': 'ge', 'ご': 'go',
    'ざ': 'za', 'じ': 'ji', 'ず': 'zu', 'ぜ': 'ze', 'ぞ': 'zo',
    'だ': 'da', 'ぢ': 'ji', 'づ': 'zu', 'で': 'de', 'ど': 'do',
    'ば': 'ba', 'び': 'bi', 'ぶ': 'bu', 'べ': 'be', 'ぼ': 'bo',
    'ぱ': 'pa', 'ぴ': 'pi', 'ぷ': 'pu', 'ぺ': 'pe', 'ぽ': 'po',
    'ぁ': 'a',  'ぃ': 'i',  'ぅ': 'u',  'ぇ': 'e',  'ぉ': 'o',
    'ゃ': 'ya', 'ゅ': 'yu', 'ょ': 'yo',
    'っ': ''
  };

  let result = '';
  let i = 0;
  let sokuon = false;

  while (i < kana.length) {
    const ch = kana[i];

    if (ch === 'っ') {
      sokuon = true;
      i++;
      continue;
    }

    const pair = kana.substring(i, i + 2);
    let roma = '';
    if (digraphMap[pair]) {
      roma = digraphMap[pair];
      i += 2;
    } else {
      roma = monoMap[ch] || '';
      i += 1;
    }

    if (sokuon && roma) {
      const c = roma.charAt(0);
      result += c + roma;
      sokuon = false;
    } else {
      result += roma;
    }
  }

  result = result
    .replace(/aa/g, 'a')
    .replace(/ii/g, 'i')
    .replace(/uu/g, 'u')
    .replace(/ee/g, 'e')
    .replace(/oo/g, 'o')
    .replace(/ou/g, 'o');

  if (!result) return '';
  return result.charAt(0).toUpperCase() + result.slice(1);
}

// --------------------
// ★【追加】Unit Coins / Challenge Unlock（students に保存）ユーティリティ
// --------------------

// unit を 1..MAX_UNITS に正規化（不正はエラー）
function _normalizeUnit_(unit) {
  const u = Number(unit);
  if (!Number.isFinite(u) || u < 1 || u > MAX_UNITS) {
    throw new Error('unit が不正です（1〜' + MAX_UNITS + '）: ' + unit);
  }
  return u;
}

// students の head に列が無ければ追加する（戻り値は index）
function _ensureStudentsColumn_(stuSheet, head, key) {
  let idx = head.indexOf(key);
  if (idx < 0) {
    const newCol = head.length + 1;
    stuSheet.getRange(1, newCol).setValue(key);
    head.push(key);
    idx = head.length - 1;
  }
  return idx;
}

// coin_u1..u8 / unlock_challenge_u1..u8 を保証
function _ensureStudentsCoinsAndUnlockColumns_(stuSheet, head) {
  const coinIdxByUnit = {};
  const unlockIdxByUnit = {};

  for (let u = 1; u <= MAX_UNITS; u++) {
    coinIdxByUnit[u] = _ensureStudentsColumn_(stuSheet, head, 'coin_u' + u);
    unlockIdxByUnit[u] = _ensureStudentsColumn_(stuSheet, head, 'unlock_challenge_u' + u);
  }
  return { coinIdxByUnit, unlockIdxByUnit };
}

// nickname（student_id）から students 行を探す（なければエラー）
function _findStudentRowByIdOrThrow_(stuSheet, head, studentId) {
  const nIdx = head.indexOf('nickname');
  if (nIdx < 0) throw new Error('students: nickname 列がありません');

  const target = String(studentId || '').trim();
  if (!target) throw new Error('students: studentId が空です');

  const lastRow = stuSheet.getLastRow();
  if (lastRow < 2) throw new Error('students: データがありません');

  const vals = stuSheet.getRange(1, 1, lastRow, head.length).getValues();
  for (let r = 1; r < vals.length; r++) {
    if (String(vals[r][nIdx] || '').trim() === target) {
      return { rowIndex: r + 1, rowValues: vals[r] };
    }
  }
  throw new Error('students: 対象の生徒が見つかりません: ' + target);
}

// ★【追加】points廃止に伴い：student の行番号だけ欲しい時用
function _getStudentRowIndexById_(studentId) {
  const stuSheet = sh(SHEET_STUDENTS);
  if (!stuSheet || stuSheet.getLastRow() < 2) return null;

  const vals = stuSheet.getDataRange().getValues();
  const head = vals[0] || [];
  const nIdx = head.indexOf('nickname');
  if (nIdx < 0) return null;

  const target = String(studentId || '').trim();
  if (!target) return null;

  for (let r = 1; r < vals.length; r++) {
    if (String(vals[r][nIdx] || '').trim() === target) {
      return r + 1;
    }
  }
  return null;
}

// --------------------
// ★【追加】(A) コイン加算：練習のミス回数 → unit別 coin に加算
// addCoinsByPractice(token, unit, missCount)
// --------------------
function addCoinsByPractice(token, unit, missCount) {
  const stu = getStudentByToken(token); // 既存の token 検証を流用
  const u = _normalizeUnit_(unit);

  const m = Number(missCount);
  const miss = Number.isFinite(m) && m >= 0 ? Math.floor(m) : 0;

  let added = 0;
  if (miss === 0) added = 5;
  else if (miss === 1) added = 3;
  else if (miss === 2) added = 1;
  else added = 0;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const stuSheet = ensureSheet(
      SHEET_STUDENTS,
      ['nickname', 'full_name', 'full_name_kana', 'first_name_en', 'login_pass', 'enabled']
    );

    // head を最新化（ensureで列追加するため）
    const head = stuSheet.getRange(1, 1, 1, stuSheet.getLastColumn()).getValues()[0] || [];
    const { coinIdxByUnit } = _ensureStudentsCoinsAndUnlockColumns_(stuSheet, head);

    const found = _findStudentRowByIdOrThrow_(stuSheet, head, stu.student_id);
    const rowIndex = found.rowIndex;

    const coinColIdx0 = coinIdxByUnit[u]; // 0-based
    const coinCell = stuSheet.getRange(rowIndex, coinColIdx0 + 1);

    const current = Number(coinCell.getValue() || 0) || 0;
    const next = current + added;
    if (added !== 0) coinCell.setValue(next);

    return {
      ok: true,
      unit: u,
      addedCoins: added,
      totalCoins: next,
      student_id: stu.student_id
    };
  } finally {
    lock.releaseLock();
  }
}

// --------------------
// ★【追加】(B) 解放（消費）：unit別 coin を消費して unlock を立てる
// unlockChallenge(token, unit, cost?)
// --------------------
function unlockChallenge(token, unit, cost) {
  const stu = getStudentByToken(token);
  const u = _normalizeUnit_(unit);

  const c = Number(cost);
  const need = Number.isFinite(c) && c > 0 ? Math.floor(c) : 10;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const stuSheet = ensureSheet(
      SHEET_STUDENTS,
      ['nickname', 'full_name', 'full_name_kana', 'first_name_en', 'login_pass', 'enabled']
    );

    const head = stuSheet.getRange(1, 1, 1, stuSheet.getLastColumn()).getValues()[0] || [];
    const { coinIdxByUnit, unlockIdxByUnit } = _ensureStudentsCoinsAndUnlockColumns_(stuSheet, head);

    const found = _findStudentRowByIdOrThrow_(stuSheet, head, stu.student_id);
    const rowIndex = found.rowIndex;

    const coinColIdx0 = coinIdxByUnit[u];
    const unlockColIdx0 = unlockIdxByUnit[u];

    const coinCell = stuSheet.getRange(rowIndex, coinColIdx0 + 1);
    const unlockCell = stuSheet.getRange(rowIndex, unlockColIdx0 + 1);

    const already = String(unlockCell.getValue() || '').trim() === '1'
      || String(unlockCell.getValue() || '').toLowerCase() === 'true';

    const current = Number(coinCell.getValue() || 0) || 0;

    if (already) {
      return {
        ok: true,
        unit: u,
        unlocked: true,
        alreadyUnlocked: true,
        required: need,
        totalCoins: current
      };
    }

    if (current < need) {
      return {
        ok: false,
        unit: u,
        unlocked: false,
        reason: 'notEnough',
        required: need,
        totalCoins: current
      };
    }

    const next = current - need;
    coinCell.setValue(next);
    unlockCell.setValue(1);

    return {
      ok: true,
      unit: u,
      unlocked: true,
      alreadyUnlocked: false,
      required: need,
      totalCoins: next
    };
  } finally {
    lock.releaseLock();
  }
}

// --------------------
// ★【追加】(C) 解放チェック：unit別 unlock を返す
// isChallengeUnlocked(token, unit)
// --------------------
function isChallengeUnlocked(token, unit) {
  const stu = getStudentByToken(token);
  const u = _normalizeUnit_(unit);

  const stuSheet = sh(SHEET_STUDENTS);
  if (!stuSheet || stuSheet.getLastRow() < 2) {
    return { ok: true, unit: u, unlocked: false, totalCoins: 0 };
  }

  const head = stuSheet.getRange(1, 1, 1, stuSheet.getLastColumn()).getValues()[0] || [];
  const { coinIdxByUnit, unlockIdxByUnit } = _ensureStudentsCoinsAndUnlockColumns_(stuSheet, head);

  const found = _findStudentRowByIdOrThrow_(stuSheet, head, stu.student_id);
  const rowIndex = found.rowIndex;

  const coinColIdx0 = coinIdxByUnit[u];
  const unlockColIdx0 = unlockIdxByUnit[u];

  const coinCell = stuSheet.getRange(rowIndex, coinColIdx0 + 1);
  const unlockCell = stuSheet.getRange(rowIndex, unlockColIdx0 + 1);

  const current = Number(coinCell.getValue() || 0) || 0;
  const unlocked = String(unlockCell.getValue() || '').trim() === '1'
    || String(unlockCell.getValue() || '').toLowerCase() === 'true';

  return { ok: true, unit: u, unlocked, totalCoins: current };
}

// --------------------
// ★【追加】Home 用API：coinsByUnit / unlockByUnit をまとめて返す
// --------------------
function getHomeStatus(token) {
  const stu = getStudentByToken(token);
  const baseUrl = ScriptApp.getService().getUrl();

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const stuSheet = ensureSheet(
      SHEET_STUDENTS,
      ['nickname', 'full_name', 'full_name_kana', 'first_name_en', 'login_pass', 'enabled']
    );

    const head = stuSheet.getRange(1, 1, 1, stuSheet.getLastColumn()).getValues()[0] || [];
    const { coinIdxByUnit, unlockIdxByUnit } = _ensureStudentsCoinsAndUnlockColumns_(stuSheet, head);

    const found = _findStudentRowByIdOrThrow_(stuSheet, head, stu.student_id);
    const rowIndex = found.rowIndex;

    const row = stuSheet.getRange(rowIndex, 1, 1, head.length).getValues()[0] || [];

    const coinsByUnit = {};
    const unlockByUnit = {};
    for (let u = 1; u <= MAX_UNITS; u++) {
      const coin = Number(row[coinIdxByUnit[u]] || 0) || 0;

      const rawUnlock = String(row[unlockIdxByUnit[u]] || '').trim();
      const unlocked = rawUnlock === '1' || rawUnlock.toLowerCase() === 'true';

      coinsByUnit[u] = coin;
      unlockByUnit[u] = unlocked;
    }

    const detail = _getStudentDetailByToken_(token) || {};

    return {
      ok: true,
      token: String(token || ''),
      baseUrl,
      student: {
        id: String(stu.student_id || ''),
        name: String(stu.name || '')
      },
      coinsByUnit,
      unlockByUnit,
      fullNameKana: String(detail.fullNameKana || ''),
      firstNameEn: String(detail.firstNameEn || '')
    };
  } finally {
    lock.releaseLock();
  }
}

// --------------------
// ★【追加】Challenge（ダミー1語）用ユーティリティ
// --------------------
const DUMMY_STOP_WORDS_ = new Set([
  'a','an','the','to','of','in','on','at','for','from','with','and','or','but',
  'am','is','are','was','were','be','been','being',
  'do','does','did','can','will','would','should','could','may','might','must',
  'i','you','he','she','it','we','they','me','him','her','us','them',
  'my','your','his','her','its','our','their'
]);

function _normWord_(w) {
  return String(w || '')
    .trim()
    .toLowerCase()
    .replace(/^[^a-z0-9']+|[^a-z0-9']+$/gi, '');
}

function _getDummyWordList_() {
  const ws = sh(SHEET_DUMMY_WORDS);
  if (!ws || ws.getLastRow() < 1) return [];

  const vals = ws.getRange(1, 1, ws.getLastRow(), 1).getValues();
  const list = [];

  for (let r = 0; r < vals.length; r++) {
    const raw = String(vals[r][0] || '').trim();
    if (!raw) continue;
    if (r === 0 && _normWord_(raw) === 'word') continue;
    list.push(raw);
  }
  return list;
}

function _pickDummyWord_(targetWords, token) {
  const all = _getDummyWordList_();
  if (!all.length) throw new Error('dummy_words シートが空です（A列に単語を入れてね）');

  const targetSet = new Set((targetWords || []).map(_normWord_).filter(Boolean));

  let recent = [];
  if (token) {
    try {
      const cache = CacheService.getScriptCache();
      const raw = cache.get('go_dummy_recent_' + String(token));
      if (raw) {
        recent = JSON.parse(raw) || [];
        if (!Array.isArray(recent)) recent = [];
      }
    } catch (e) {
      recent = [];
    }
  }
  const recentSet = new Set(recent.map(_normWord_).filter(Boolean));

  const candidates = [];
  for (const w of all) {
    const nw = _normWord_(w);
    if (!nw) continue;
    if (DUMMY_STOP_WORDS_.has(nw)) continue;
    if (targetSet.has(nw)) continue;
    if (recentSet.has(nw)) continue;
    candidates.push(w);
  }

  let finalList = candidates;
  if (!finalList.length) {
    finalList = all.filter(w => {
      const nw = _normWord_(w);
      return nw && !DUMMY_STOP_WORDS_.has(nw) && !targetSet.has(nw);
    });
  }

  if (!finalList.length) {
    throw new Error('ダミー候補がありません（dummy_words を増やすか、stopwords/被り条件を見直してね）');
  }

  const picked = finalList[Math.floor(Math.random() * finalList.length)];

  if (token) {
    try {
      const cache = CacheService.getScriptCache();
      const next = [picked, ...recent].slice(0, 3);
      cache.put('go_dummy_recent_' + String(token), JSON.stringify(next), 60 * 30);
    } catch (e) {}
  }

  return picked;
}

// ---------- Students / Sessions ----------
function _issueToken_(studentKey, displayName) {
  const token = Utilities.getUuid().replace(/-/g, '').slice(0, 24);
  const sess = ensureSheet(SHEET_SESSIONS, ['token', 'student_id', 'name', 'created_at', 'active']);
  sess.appendRow([token, studentKey, displayName, nowDate(), true]);

  const baseUrl = ScriptApp.getService().getUrl();

  return {
    token,
    baseUrl,
    student: { id: studentKey, name: displayName }
  };
}

function createSession(mode, studentId, name, clazz, password) {
  const m = String(mode || 'login').trim();

  const nickname     = String(studentId || '').trim();
  const fullName     = String(name || '').trim();
  const fullNameKana = String(clazz || '').trim();

  const pass = normalizeDigits(String(password || ''))
    .replace(/[^0-9]/g, '');

  const s = ensureSheet(
    SHEET_STUDENTS,
    ['nickname', 'full_name', 'full_name_kana', 'first_name_en', 'login_pass', 'enabled']
  );
  const values = s.getDataRange().getValues();
  const head = values[0] || [];

  let lastCol = head.length;

  function ensureHeader_(key) {
    let idx = head.indexOf(key);
    if (idx < 0) {
      lastCol += 1;
      s.getRange(1, lastCol).setValue(key);
      head[lastCol - 1] = key;
      idx = lastCol - 1;
    }
    return idx;
  }

  const idx = {
    nickname:     ensureHeader_('nickname'),
    fullName:     ensureHeader_('full_name'),
    fullNameKana: ensureHeader_('full_name_kana'),
    firstNameEn:  ensureHeader_('first_name_en'),
    pass:         ensureHeader_('login_pass'),
    enabled:      ensureHeader_('enabled')
  };

  if (!/^\d{4}$/.test(pass)) {
    throw new Error('パスワードは数字4ケタで入力してください。（例：0123）');
  }

  let foundRow = -1;
  for (let r = 1; r < values.length; r++) {
    const rowNick = String(values[r][idx.nickname] || '').trim();
    if (rowNick === nickname) {
      foundRow = r + 1;
      break;
    }
  }

  if (m === 'register') {
    if (!fullName) throw new Error('本名（フルネーム）を入力してください。');
    if (!fullNameKana) throw new Error('ふりがな（フルネーム）を入力してください。');
    if (!nickname) throw new Error('ニックネームを入力してください。');

    if (foundRow !== -1) {
      const row = s.getRange(foundRow, 1, 1, head.length).getValues()[0];
      const enabled = String(row[idx.enabled]).toLowerCase() !== 'false';
      if (!enabled) throw new Error('このニックネームは先生が停止しています。先生に確認してください。');
      throw new Error('このニックネームはすでに使われています。ちがうニックネームにしてね。');
    }

    const givenKana   = extractGivenKana_(fullNameKana);
    const firstNameEn = kanaToHepburnName_(givenKana);

    const row = new Array(head.length).fill('');
    row[idx.nickname]     = nickname;
    row[idx.fullName]     = fullName;
    row[idx.fullNameKana] = fullNameKana;
    row[idx.firstNameEn]  = firstNameEn;
    row[idx.pass]         = "'" + pass;
    row[idx.enabled]      = true;

    s.appendRow(row);

    return _issueToken_(nickname, fullName);
  }

  if (!nickname) throw new Error('ニックネームを入力してください。');
  if (foundRow === -1) throw new Error('このニックネームは登録されていません。先生に確認してください。');

  const row = s.getRange(foundRow, 1, 1, head.length).getValues()[0];
  const enabled = String(row[idx.enabled]).toLowerCase() !== 'false';
  if (!enabled) throw new Error('このアカウントは先生が停止しています。');

  const rowFullName = String(row[idx.fullName] || '').trim();
  const rowPass = normalizeDigits(String(row[idx.pass] || ''))
    .replace(/[^0-9]/g, '');

  if (rowPass !== pass) throw new Error('パスワードがちがいます。もう一度入力してください。');

  const sess = ensureSheet(SHEET_SESSIONS, ['token', 'student_id', 'name', 'created_at', 'active']);
  const sVals = sess.getDataRange().getValues();
  const sHead = sVals[0] || [];
  const tIdx  = sHead.indexOf('token');
  const sidIdx = sHead.indexOf('student_id');
  const actIdx = sHead.indexOf('active');

  let reuseToken = null;
  for (let i = sVals.length - 1; i >= 1; i--) {
    const r = sVals[i];
    const sid = String(r[sidIdx] || '').trim();
    const active = (actIdx < 0) ? true : String(r[actIdx]).toLowerCase() !== 'false';
    if (sid === nickname && active) {
      reuseToken = String(r[tIdx] || '').trim();
      break;
    }
  }

  if (reuseToken) {
    const baseUrl = ScriptApp.getService().getUrl();
    return {
      token: reuseToken,
      baseUrl,
      student: { id: nickname, name: rowFullName }
    };
  }

  return _issueToken_(nickname, rowFullName);
}

function _getStudentDetailByToken_(token) {
  token = String(token || '').trim();
  if (!token) return null;

  const sess = sh(SHEET_SESSIONS);
  if (!sess || sess.getLastRow() < 2) return null;

  const sVals = sess.getDataRange().getValues();
  const sHead = sVals[0] || [];
  const tIdx  = sHead.indexOf('token');
  const sidIdx = sHead.indexOf('student_id');
  if (tIdx < 0 || sidIdx < 0) return null;

  let nickname = '';
  for (let r = 1; r < sVals.length; r++) {
    const row = sVals[r];
    if (String(row[tIdx] || '').trim() === token) {
      nickname = String(row[sidIdx] || '').trim();
      break;
    }
  }
  if (!nickname) return null;

  const stuSheet = sh(SHEET_STUDENTS);
  if (!stuSheet || stuSheet.getLastRow() < 2) return null;

  const v = stuSheet.getDataRange().getValues();
  const h = v[0] || [];
  const nIdx = h.indexOf('nickname');
  const kanaIdx = h.indexOf('full_name_kana');
  const firstIdx = h.indexOf('first_name_en');
  if (nIdx < 0) return null;

  for (let r = 1; r < v.length; r++) {
    const row = v[r];
    if (String(row[nIdx] || '').trim() === nickname) {
      return {
        nickname,
        fullNameKana: kanaIdx >= 0 ? String(row[kanaIdx] || '').trim() : '',
        firstNameEn: firstIdx >= 0 ? String(row[firstIdx] || '').trim() : ''
      };
    }
  }
  return null;
}

function getStudentByToken(token) {
  const sess = sh(SHEET_SESSIONS);
  if (!sess) throw new Error('sessions シートがありません');
  const values = sess.getDataRange().getValues();
  if (!values.length) throw new Error('sessions シートが空です');

  const head = values[0];
  const tIdx = head.indexOf('token');
  const idIdx = head.indexOf('student_id');
  const nIdx = head.indexOf('name');
  const aIdx = head.indexOf('active');

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (row[tIdx] === token && (aIdx < 0 || String(row[aIdx]).toLowerCase() !== 'false')) {
      const sid = row[idIdx];
      const rowIndex = _getStudentRowIndexById_(sid);

      return {
        student_id: sid,
        name: nIdx >= 0 ? row[nIdx] : '',
        studentRow: rowIndex
      };
    }
  }
  throw new Error('無効なトークンです。ログインし直してください。');
}

// ---------- Items ----------
function _getHeaderIndexMap_(header) {
  const map = {};
  header.forEach((v, i) => { map[String(v).trim()] = i; });

  const correctKey =
    ('correct_sentence' in map) ? 'correct_sentence' :
    ('coreect_sentence' in map) ? 'coreect_sentence' : null;

  if (!('id' in map)) throw new Error('items_db: id 列がありません');
  if (!('prompt' in map)) throw new Error('items_db: prompt 列がありません');
  if (!correctKey)       throw new Error('items_db: correct_sentence（または coreect_sentence）がありません');

  return {
    id: map.id,
    prompt: map.prompt,
    correct: map[correctKey],
    ok: map.audio_ok_url,
    ng: map.audio_ng_url,
    unit: (typeof map.unit === 'number')
      ? map.unit
      : (typeof map.Unit === 'number' ? map.Unit : null)
  };
}

/**
 * ★修正（最小）：roles シートから id の行を拾って roles オブジェクトにする
 */
function _getRolesById_(sheetName, id) {
  const ws = sh(sheetName);
  if (!ws || ws.getLastRow() < 2) return null;

  const values = ws.getDataRange().getValues();
  const head = values[0] || [];
  if (!head.length) return null;

  const col = {};
  head.forEach((h, i) => {
    const key = String(h || '').trim().toLowerCase();
    if (key) col[key] = i;
  });

  const idIdx  = col['id'];
  if (idIdx == null) return null;

  const whIdx  = col['wh'];
  const sIdx   = col['s'];
  const auxIdx = col['aux'];
  const vIdx   = col['v'];
  const ocIdx  =
    (col['o/c'] != null) ? col['o/c'] :
    (col['oc']  != null) ? col['oc']  :
    (col['o_c'] != null) ? col['o_c'] : null;
  const advIdx = col['adv'];

  const targetId = String(id).trim();
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (String(row[idIdx] || '').trim() !== targetId) continue;

    return {
      Wh:  whIdx  != null ? String(row[whIdx]  || '') : '',
      S:   sIdx   != null ? String(row[sIdx]   || '') : '',
      Aux: auxIdx != null ? String(row[auxIdx] || '') : '',
      V:   vIdx   != null ? String(row[vIdx]   || '') : '',
      OC:  ocIdx  != null ? String(row[ocIdx]  || '') : '',
      Adv: advIdx != null ? String(row[advIdx] || '') : ''
    };
  }
  return null;
}

function getItemById(id, token) {
  const ws = sh(SHEET_ITEMS);
  if (!ws) throw new Error('items_db シートがありません');
  const values = ws.getDataRange().getValues();
  if (!values.length) throw new Error('items_db シートが空です');
  const idx = _getHeaderIndexMap_(values[0]);

  let stuDetail = null;
  if (token) {
    try {
      stuDetail = _getStudentDetailByToken_(token);
    } catch (e) {
      stuDetail = null;
    }
  }

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (String(row[idx.id]) === String(id)) {
      let prompt = String(row[idx.prompt] || '');
      const sentence = String(row[idx.correct] || '').trim();
      if (!sentence) throw new Error('このIDの correct_sentence が空です: ' + id);

      if (stuDetail && prompt.indexOf('〇〇') !== -1) {
        const givenKana = extractGivenKana_(stuDetail.fullNameKana || '');
        const dispName = givenKana || (stuDetail.fullNameKana || '');
        if (dispName) {
          prompt = prompt.replace('〇〇', dispName);
        }
      }

      let words = sentence.split(/\s+/);
      if (stuDetail && stuDetail.firstNameEn) {
        const fn = stuDetail.firstNameEn;
        const trimmed = sentence.replace(/\s+/g, ' ').trim();
        if (trimmed === 'I am') {
          words = ['I', 'am', fn];
        }
      }

      const unitVal = idx.unit != null ? Number(row[idx.unit] || 0) || 1 : 1;

      const rolesJp = _getRolesById_(SHEET_ROLES_JP, id);
      const rolesEn = _getRolesById_(SHEET_ROLES_EN, id);

      return {
        id: String(id),
        prompt,
        words,
        audioOk: String((idx.ok != null ? row[idx.ok] : '') || ''),
        audioNg: String((idx.ng != null ? row[idx.ng] : '') || ''),
        unit: unitVal,
        total: Math.max(ws.getLastRow() - 1, 0),

        rolesJp: rolesJp,
        rolesEn: rolesEn,
        roles_jp: rolesJp,
        roles_en: rolesEn,

        roles_Jp: rolesJp,
        roles_En: rolesEn,
        jpRoles: rolesJp,
        enRoles: rolesEn
      };
    }
  }
  throw new Error('指定IDが見つかりません: ' + id);
}

function listPracticeItemIds(unit) {
  const unitNum = unit != null && unit !== '' ? (Number(unit) || 0) : 0;

  const ws = sh(SHEET_ITEMS);
  if (!ws || ws.getLastRow() < 2) return [];

  const values = ws.getDataRange().getValues();
  const idx = _getHeaderIndexMap_(values[0]);

  const ids = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const id = String(row[idx.id] || '').trim();
    if (!id) continue;

    const unitVal = idx.unit != null ? Number(row[idx.unit] || 0) || 1 : 1;
    if (unitNum && unitVal !== unitNum) continue;

    ids.push(id);
  }
  return ids;
}

function getRandomPracticeItem(token, unit, excludeIds) {
  token = String(token || '').trim();
  const unitNum = Number(unit || 0) || 0;

  const ws = sh(SHEET_ITEMS);
  if (!ws || ws.getLastRow() < 2) {
    throw new Error('items_db シートがありません');
  }

  const values = ws.getDataRange().getValues();
  const idx = _getHeaderIndexMap_(values[0]);

  const excludeSet = new Set();
  if (Array.isArray(excludeIds)) {
    excludeIds.forEach(id => excludeSet.add(String(id)));
  } else if (excludeIds != null && excludeIds !== '') {
    String(excludeIds).split(',').forEach(id => {
      const trimmed = String(id || '').trim();
      if (trimmed) excludeSet.add(trimmed);
    });
  }

  const buildCandidates = (ignoreExclude) => {
    const list = [];
    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const id = String(row[idx.id] || '').trim();
      if (!id) continue;

      const unitVal = idx.unit != null ? Number(row[idx.unit] || 0) || 1 : 1;
      if (unitNum && unitVal !== unitNum) continue;

      if (!ignoreExclude && excludeSet.size && excludeSet.has(id)) continue;

      list.push(id);
    }
    return list;
  };

  let candidates = buildCandidates(false);
  if (!candidates.length) {
    candidates = buildCandidates(true);
  }

  if (!candidates.length) {
    return null;
  }

  const randomId = candidates[Math.floor(Math.random() * candidates.length)];
  return getItemById(randomId, token);
}

function checkOrder(id, arrangedWords, token) {
  const item = getItemById(id, token);

  const target = (item.words || []).map(w => String(w));
  const answer = (arrangedWords || []).map(w => String(w));

  let ok = false;

  if (
    target.length === 2 &&
    target[0] === 'I' &&
    target[1] === 'am' &&
    answer.length === 3 &&
    answer[0] === 'I' &&
    answer[1] === 'am'
  ) {
    ok = true;
  } else {
    ok =
      target.length === answer.length &&
      target.every((w, i) => w === answer[i]);
  }

  return { ok };
}

// --------------------
// ★【追加】Challenge API（ダミー1語）
// --------------------
function getChallengeItemById(id, token) {
  const item = getItemById(id, token);

  const dummyWord = _pickDummyWord_(item.words || [], token);
  const wordsWithDummy = _shuffle_([...(item.words || []), dummyWord]);

  const correctSentence = (item.words || []).join(' ').trim();

  return {
    ...item,
    correct_sentence: correctSentence,
    dummyWord,
    wordsWithDummy
  };
}

function checkChallengeOrder(id, arrangedWords, dummyWord, token) {
  const item = getItemById(id, token);
  const target = (item.words || []).map(w => String(w));
  const answer = (arrangedWords || []).map(w => String(w));

  const dummy = String(dummyWord || '').trim();
  const dummyUsed = !!dummy && answer.includes(dummy);

  if (dummyUsed) {
    return { ok: false, reason: 'dummyUsed', dummyUsed };
  }

  const ok =
    target.length === answer.length &&
    target.every((w, i) => w === answer[i]);

  return { ok, reason: ok ? 'ok' : 'mismatch', dummyUsed: false };
}

function recordChallengeResult(token, itemId, ok, timeMs, dummyWord, dummyUsed, reason) {
  const stu = getStudentByToken(token);

  let unitVal = 1;
  try {
    const item = getItemById(itemId);
    unitVal = Number(item.unit || 1) || 1;
  } catch (e) {
    unitVal = 1;
  }

  const r = ensureSheet(
    SHEET_RESULTS,
    ['timestamp','token','student_id','name','item_id','correct','time_ms','mode','unit']
  );

  const lastCol = r.getLastColumn();
  const head = r.getRange(1, 1, 1, lastCol).getValues()[0] || [];

  function ensureHeader_(key) {
    let idx = head.indexOf(key);
    if (idx < 0) {
      const newCol = head.length + 1;
      r.getRange(1, newCol).setValue(key);
      head.push(key);
      idx = head.length - 1;
    }
    return idx;
  }

  const modeIdx = ensureHeader_('mode');
  const unitIdx = ensureHeader_('unit');

  const dummyIdx = ensureHeader_('dummy_word');
  const usedIdx  = ensureHeader_('dummy_used');
  const rsnIdx   = ensureHeader_('challenge_reason');

  const row = [
    nowDate(),
    token,
    stu.student_id,
    stu.name,
    itemId,
    ok ? 1 : 0,
    Number(timeMs) || 0
  ];

  const maxIdx = Math.max(modeIdx, unitIdx, dummyIdx, usedIdx, rsnIdx);
  if (row.length <= maxIdx) {
    for (let i = row.length; i <= maxIdx; i++) row[i] = '';
  }

  row[modeIdx] = 'challenge';
  row[unitIdx] = unitVal;

  row[dummyIdx] = String(dummyWord || '').trim();
  row[usedIdx]  = dummyUsed ? 1 : 0;
  row[rsnIdx]   = String(reason || '').trim();

  r.appendRow(row);
  return { saved: true };
}

function _shuffle_(a) {
  const arr = Array.isArray(a) ? a.slice() : [];
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}

function listFocusSentences(token, unit) {
  const ws = sh(SHEET_ITEMS);
  if (!ws) throw new Error('items_db シートがありません');

  const lastRow = ws.getLastRow();
  if (lastRow < 2) return [];

  const values = ws.getDataRange().getValues();
  const idx = _getHeaderIndexMap_(values[0]);

  const unitNum = (unit != null && unit !== '') ? (Number(unit) || 0) : 0;

  let stuDetail = null;
  if (token) {
    try {
      stuDetail = _getStudentDetailByToken_(token);
    } catch (e) {
      stuDetail = null;
    }
  }

  const LIMIT = 100;
  const result = [];

  for (let r = 1; r < values.length; r++) {
    const row = values[r];

    const id = row[idx.id];
    if (!id) continue;

    const unitVal = idx.unit != null ? (Number(row[idx.unit] || 0) || 1) : 1;
    if (unitNum && unitVal !== unitNum) continue;

    let jp = String(row[idx.prompt] || '').trim();
    if (!jp) continue;

    if (stuDetail && jp.indexOf('〇〇') !== -1) {
      const givenKana = extractGivenKana_(stuDetail.fullNameKana || '');
      const dispName = givenKana || (stuDetail.fullNameKana || '');
      if (dispName) jp = jp.replace('〇〇', dispName);
    }

    let enText = String(row[idx.correct] || '').trim();
    if (stuDetail && stuDetail.firstNameEn) {
      const trimmed = enText.replace(/\s+/g, ' ').trim();
      if (trimmed === 'I am') {
        enText = 'I am ' + stuDetail.firstNameEn;
      }
    }

    result.push({
      id: String(id),
      text: jp,
      unit: unitVal,
      enText: String(enText || '')
    });

    if (result.length >= LIMIT) break;
  }

  return result;
}

// ---------- Results ----------
function recordResult(token, itemId, ok, timeMs, mode, unit, attemptCount) {
  const stu = getStudentByToken(token);

  let unitVal = 1;
  const unitArg = Number(unit);
  if (Number.isFinite(unitArg) && unitArg > 0) {
    unitVal = unitArg;
  } else {
    try {
      const item = getItemById(itemId);
      unitVal = Number(item.unit || 1) || 1;
    } catch (e) {
      unitVal = 1;
    }
  }

  const r = ensureSheet(
    SHEET_RESULTS,
    ['timestamp','token','student_id','name','item_id','correct','time_ms','mode','unit']
  );

  const lastCol = r.getLastColumn();
  const head = r.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  let modeIdx = head.indexOf('mode');
  let unitIdx = head.indexOf('unit');

  if (modeIdx < 0) {
    modeIdx = head.length;
    r.getRange(1, modeIdx + 1).setValue('mode');
    head[modeIdx] = 'mode';
  }
  if (unitIdx < 0) {
    unitIdx = head.length;
    r.getRange(1, unitIdx + 1).setValue('unit');
    head[unitIdx] = 'unit';
  }

  const modeStr = String(mode || '').trim().toLowerCase();

  const row = [
    nowDate(),
    token,
    stu.student_id,
    stu.name,
    itemId,
    ok ? 1 : 0,
    Number(timeMs) || 0
  ];

  const maxIdx = Math.max(modeIdx, unitIdx);
  if (row.length <= maxIdx) {
    for (let i = row.length; i <= maxIdx; i++) {
      row[i] = '';
    }
  }
  row[modeIdx] = modeStr;
  row[unitIdx] = unitVal;

  r.appendRow(row);

  return { saved: true };
}

// ---------- 復習用 共通ロジック ----------
function _getWrongStatsByToken(token, unit) {
  const ws = sh(SHEET_RESULTS);
  if (!ws || ws.getLastRow() < 2) return {};

  const values = ws.getDataRange().getValues();
  const head = values[0] || [];

  const tokenIdx   = head.indexOf('token');
  const itemIdx    = head.indexOf('item_id');
  const correctIdx = head.indexOf('correct');
  const modeIdx    = head.indexOf('mode');
  const unitIdx    = head.indexOf('unit');

  if (tokenIdx < 0 || itemIdx < 0 || correctIdx < 0) return {};

  const targetUnit = unit ? Number(unit) || 0 : 0;

  const stats = {};
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (String(row[tokenIdx]) !== String(token)) continue;

    if (targetUnit && unitIdx >= 0) {
      const rowUnit = Number(row[unitIdx] || 0) || 0;
      if (rowUnit !== targetUnit) continue;
    }

    const itemId = String(row[itemIdx] || '').trim();
    if (!itemId) continue;

    const correct = Number(row[correctIdx] || 0) === 1;
    const mode = modeIdx >= 0
      ? String(row[modeIdx] || '').trim().toLowerCase()
      : '';

    if (!stats[itemId]) {
      stats[itemId] = { wrongCount: 0, stampCorrectCount: 0 };
    }
    const st = stats[itemId];

    if (!correct) {
      st.wrongCount += 1;
    } else if (mode === 'stamp') {
      st.stampCorrectCount += 1;
    }
  }

  return stats;
}

function getReviewList(token, unit) {
  token = String(token || '').trim();
  if (!token) return [];
  const unitNum = (unit != null && unit !== '') ? (Number(unit) || 0) : 0;

  const stats = _getWrongStatsByToken(token, unitNum || null);
  return Object.keys(stats).filter(id => stats[id].wrongCount > 0);
}

// ---------- スタンプラリー用 API ----------
function getStampItems(token, unit) {
  token = String(token || '').trim();
  const unitNum = (unit != null && unit !== '') ? (Number(unit) || 0) : 0;

  if (!token) return { items: [] };

  const stuDetail = _getStudentDetailByToken_(token);

  const stats = _getWrongStatsByToken(token, unitNum || null);

  const targetIds = Object.keys(stats).filter(id => stats[id].wrongCount > 0);
  if (!targetIds.length) {
    return { items: [] };
  }

  const ws = sh(SHEET_ITEMS);
  if (!ws || ws.getLastRow() < 2) return { items: [] };

  const values = ws.getDataRange().getValues();
  const idx = _getHeaderIndexMap_(values[0]);

  const result = [];
  const idSet = new Set(targetIds);

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const id = String(row[idx.id] || '').trim();
    if (!id || !idSet.has(id)) continue;

    const sentence = String(row[idx.correct] || '').trim();
    if (!sentence) continue;

    const unitVal = idx.unit != null ? Number(row[idx.unit] || 0) || 1 : 1;
    if (unitNum && unitVal !== unitNum) continue;

    const words = sentence.split(/\s+/);
    const prompt = String(row[idx.prompt] || '');
    const audioOk = String((idx.ok != null ? row[idx.ok] : '') || '');
    const audioNg = String((idx.ng != null ? row[idx.ng] : '') || '');

    const st = stats[id] || { wrongCount: 0, stampCorrectCount: 0 };

    const item = {
      id,
      prompt,
      words,
      audioOk,
      audioNg,
      reviewCount: st.stampCorrectCount,
      wrongCount: st.wrongCount,
      correct_sentence: sentence,
      unit: unitVal
    };

    if (stuDetail) {
      item.full_name_kana = stuDetail.fullNameKana || '';
      item.first_name_en  = stuDetail.firstNameEn  || '';
    }

    result.push(item);
  }

  return { items: result };
}

// ---------- ボス討伐用 API ----------
function getBossItems(token, unit) {
  token = String(token || '').trim();
  const unitNum = Number(unit || 0) || 0;

  if (!token) return { items: [] };

  const stuDetail = _getStudentDetailByToken_(token);

  const stats = _getWrongStatsByToken(token, unitNum || null);

  const targetIds = Object.keys(stats).filter(id => stats[id].wrongCount > 0);
  if (!targetIds.length) {
    return { items: [] };
  }
  const idSet = new Set(targetIds);

  const ws = sh(SHEET_ITEMS);
  if (!ws || ws.getLastRow() < 2) return { items: [] };

  const values = ws.getDataRange().getValues();
  const idx = _getHeaderIndexMap_(values[0]);

  const items = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const id = String(row[idx.id] || '').trim();
    if (!id || !idSet.has(id)) continue;

    const sentence = String(row[idx.correct] || '').trim();
    if (!sentence) continue;

    const unitVal = idx.unit != null ? Number(row[idx.unit] || 0) || 1 : 1;
    if (unitNum && unitVal !== unitNum) continue;

    const words = sentence.split(/\s+/);
    const prompt = String(row[idx.prompt] || '');
    const audioOk = String((idx.ok != null ? row[idx.ok] : '') || '');
    const audioNg = String((idx.ng != null ? row[idx.ng] : '') || '');

    const item = {
      id,
      prompt,
      words,
      audioOk,
      audioNg,
      unit: unitVal,
      correct_sentence: sentence
    };

    if (stuDetail) {
      item.full_name_kana = stuDetail.fullNameKana || '';
      item.first_name_en  = stuDetail.firstNameEn  || '';
    }

    items.push(item);
  }

  return { items };
}

// ---------- Dashboard utils ----------
function setupDashboard() {
  const db = ensureSheet('dashboard', []);
  db.clear();

  db.getRange('A1').setValue('ダッシュボード');
  db.getRange('A3').setValue('表示期間（日数）');
  db.getRange('B3').setValue(7);
  db.getRange('A4').setValue('（結果・生徒別 集計）');

  db.getRange('A6').setFormula(
    `=ARRAYFORMULA(
      IFERROR(
        LET(
          rng, FILTER(results!A2:H, results!A2:A >= TODAY()-$B$3),
          q,   QUERY(rng,
            "select Col3, Col4, count(Col1), sum(Col6), avg(Col7), max(Col1)
             group by Col3, Col4
             label count(Col1) 'Attempts', sum(Col6) 'Correct', avg(Col7) 'Avg_ms', max(Col1) 'Last'", 0),
          hdr, {"student_id","name","Attempts","Correct","Accuracy","Avg_ms","Last"},
          data, IF(ROWS(q)=0,, { INDEX(q,,1), INDEX(q,,2), INDEX(q,,3), INDEX(q,,4),
                                 IF(INDEX(q,,3)>0, INDEX(q,,4)/INDEX(q,,3), ),
                                 INDEX(q,,5), INDEX(q,,6) } ),
          VSTACK(hdr, data)
        ),
      "")
    )`
  );

  db.getRange('A20').setFormula(
    `=ARRAYFORMULA(
      QUERY(
        FILTER(results!A2:H, results!A2:A >= TODAY()-$B$3),
        "select Col5, count(Col1), sum(Col6), sum(Col6)/count(Col1)
         group by Col5
         label count(Col1) 'Attempts', sum(Col6) 'Correct', sum(Col6)/count(Col1) 'Accuracy'", 0)
    )`
  );
}

function verifyTeacher(pin) {
  return { ok: pin === TEACHER_PIN };
}

function listSummary(days) {
  const r = sh(SHEET_RESULTS);
  if (!r || r.getLastRow() < 2) return [];
  const vals = r.getDataRange().getValues();
  const head = vals[0];
  const tIdx = head.indexOf('timestamp');
  const idIdx = head.indexOf('student_id');
  const nIdx = head.indexOf('name');
  const cIdx = head.indexOf('correct');
  const msIdx = head.indexOf('time_ms');

  const since = days && days > 0
    ? new Date(Date.now() - days * 24 * 3600 * 1000)
    : null;

  const sum = {};
  for (let i = 1; i < vals.length; i++) {
    const row = vals[i];
    const ts = row[tIdx] instanceof Date ? row[tIdx] : new Date(row[tIdx]);
    if (since && ts < since) continue;

    const key = String(row[idIdx] || '');
    if (!sum[key]) {
      sum[key] = {
        student_id: key,
        name: row[nIdx] || '',
        attempts: 0,
        correct: 0,
        avg_ms: 0,
        last: ts
      };
    }
    const s = sum[key];
    s.attempts += 1;
    if (Number(row[cIdx]) === 1) s.correct += 1;
    const ms = Number(row[msIdx] || 0);
    s.avg_ms += ms;
    if (ts > s.last) s.last = ts;
  }

  return Object.values(sum)
    .map(s => ({
      ...s,
      avg_ms: s.attempts ? Math.round(s.avg_ms / s.attempts) : 0
    }))
    .sort((a, b) => b.correct - a.correct || a.avg_ms - b.avg_ms);
}

// ---------- ボス画像配信（GAS → <img> 直読み） ----------
function serveBossImage_(unit) {
  const id = BOSS_IMAGE_IDS[unit] || BOSS_IMAGE_IDS['1'];
  if (!id) {
    return ContentService
      .createBinaryOutput(Utilities.newBlob(''))
      .setMimeType(ContentService.MimeType.PNG);
  }

  const file = DriveApp.getFileById(id);
  const blob = file.getBlob();

  return ContentService
    .createBinaryOutput(blob)
    .setMimeType(ContentService.MimeType.PNG);
}
