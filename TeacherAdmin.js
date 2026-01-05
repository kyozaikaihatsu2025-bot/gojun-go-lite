/**
 * TeacherAdmin.gs
 * 先生ログイン、設定保存、ダッシュボード集計、生徒管理 を提供
 * 既存の Code.gs を壊さないため、doGet はここでは作らない
 *
 * 前提（既存にある想定）:
 * - SS_ID（スプレID）
 * - SHEET_STUDENTS = 'students'
 * - SHEET_RESULTS  = 'results'
 * - SHEET_SESSIONS = 'sessions'
 * - TEACHER_PIN（先生PIN）
 * - MAX_UNITS（なければ 8 扱いにする）
 */

// ---------- 設定/セッション（ScriptProperties） ----------
const ADMIN_TTL_MS = 12 * 60 * 60 * 1000; // 12時間
const PROP_ADMIN_PREFIX = 'GO_ADMIN_SESS_';
const PROP_SETTINGS_JSON = 'GO_ADMIN_SETTINGS_JSON';

// ---------- 共通：/exec baseUrl を返す（白画面防止の核） ----------
function getExecBaseUrl() {
  const url = ScriptApp.getService().getUrl();
  return String(url).split('?')[0]; // .../exec
}

// ---------- 先生ログイン ----------
function teacherVerify(pinRaw) {
  const baseUrl = getExecBaseUrl();

  const pin = String(pinRaw ?? '').replace(/[^\d]/g, '');
  if (pin.length !== 4) {
    return { ok: false, message: 'PINは数字4ケタだよ', baseUrl };
  }

  // TEACHER_PIN が未定義なら、原因が分かるように落とす
  const teacherPin = (typeof TEACHER_PIN !== 'undefined') ? String(TEACHER_PIN) : null;
  if (!teacherPin) {
    return { ok: false, message: 'TEACHER_PIN が Code.gs に無いみたい', baseUrl };
  }

  if (pin !== teacherPin) {
    return { ok: false, message: 'PINがちがうよ', baseUrl };
  }

  const adminKey = Utilities.getUuid();
  const now = Date.now();

  const props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_ADMIN_PREFIX + adminKey, String(now));

  // 初回はデフォ設定を用意
  _ensureDefaultSettings_();

  return {
    ok: true,
    adminKey,
    expiresAt: now + ADMIN_TTL_MS,
    baseUrl,
  };
}

function teacherLogout(adminKey) {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(PROP_ADMIN_PREFIX + String(adminKey || ''));
  return { ok: true };
}

function _requireAdmin_(adminKey) {
  const key = String(adminKey || '');
  if (!key) throw new Error('adminKey がありません');

  const props = PropertiesService.getScriptProperties();
  const ts = Number(props.getProperty(PROP_ADMIN_PREFIX + key) || 0);
  if (!ts) throw new Error('ログイン期限が切れたよ（もう一回ログインしてね）');

  if (Date.now() - ts > ADMIN_TTL_MS) {
    props.deleteProperty(PROP_ADMIN_PREFIX + key);
    throw new Error('ログイン期限が切れたよ（もう一回ログインしてね）');
  }

  // 生存確認（更新）
  props.setProperty(PROP_ADMIN_PREFIX + key, String(Date.now()));
}

// ---------- 設定（先生がいじるやつ） ----------
function teacherGetSettings(adminKey) {
  _requireAdmin_(adminKey);
  _ensureDefaultSettings_();
  const props = PropertiesService.getScriptProperties();
  return JSON.parse(props.getProperty(PROP_SETTINGS_JSON));
}

function teacherSaveSettings(adminKey, settings) {
  _requireAdmin_(adminKey);

  // 変な値が入らないように軽く整形
  const normalized = _normalizeSettings_(settings);

  const props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_SETTINGS_JSON, JSON.stringify(normalized));

  return { ok: true, settings: normalized };
}

function teacherResetSettings(adminKey) {
  _requireAdmin_(adminKey);
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(PROP_SETTINGS_JSON);
  _ensureDefaultSettings_();
  return { ok: true, settings: JSON.parse(props.getProperty(PROP_SETTINGS_JSON)) };
}

function _ensureDefaultSettings_() {
  const props = PropertiesService.getScriptProperties();
  const cur = props.getProperty(PROP_SETTINGS_JSON);
  if (cur) return;

  const units = _getMaxUnits_();
  const bossCostByUnit = {};
  const goalsByUnit = {};
  for (let u = 1; u <= units; u++) {
    bossCostByUnit[String(u)] = (u === 1) ? 10 : 15; // 例：Unit1=10、他=15
    goalsByUnit[String(u)] = 20; // 例：Unitの目標
  }

  const defaults = {
    version: 1,
    updatedAt: Date.now(),

    // A. 音声スピード
    tts: { normal: 1.0, slow: 0.5 },

    // 生徒ごとの上書き（例：{"studentId": {normal:0.9, slow:0.6}}）
    ttsOverridesByStudentId: {},

    // B. ボス解放コスト（unit別）
    bossCostByUnit,

    // C. コイン付与ルール（mode別）
    coinRulesByMode: {
      practice: { miss0: 5, miss1: 3, miss2: 1, else: 0 },
      challenge: { miss0: 5, miss1: 3, miss2: 1, else: 0 },
      boss: { miss0: 5, miss1: 3, miss2: 1, else: 0 },
    },

    // D. スタンプ★閾値
    stampThresholds: { star1: 2, star2: 5, star3: 8 },

    // 学習モードON/OFF
    featureToggles: { challenge: true, boss: true, stamp: true },

    // 目標設定（unit別）
    goalsByUnit,
  };

  props.setProperty(PROP_SETTINGS_JSON, JSON.stringify(defaults));
}

function _normalizeSettings_(s) {
  const units = _getMaxUnits_();

  const num = (v, def) => {
    const n = Number(v);
    return Number.isFinite(n) ? n : def;
  };

  const clamp = (v, min, max, def) => {
    const n = num(v, def);
    return Math.min(max, Math.max(min, n));
  };

  const out = {
    version: 1,
    updatedAt: Date.now(),
    tts: {
      normal: clamp(s?.tts?.normal, 0.5, 2.0, 1.0),
      slow: clamp(s?.tts?.slow, 0.3, 1.5, 0.5),
    },
    ttsOverridesByStudentId: (s?.ttsOverridesByStudentId && typeof s.ttsOverridesByStudentId === 'object')
      ? s.ttsOverridesByStudentId
      : {},
    bossCostByUnit: {},
    coinRulesByMode: {},
    stampThresholds: {
      star1: Math.max(0, Math.floor(num(s?.stampThresholds?.star1, 2))),
      star2: Math.max(0, Math.floor(num(s?.stampThresholds?.star2, 5))),
      star3: Math.max(0, Math.floor(num(s?.stampThresholds?.star3, 8))),
    },
    featureToggles: {
      challenge: !!s?.featureToggles?.challenge,
      boss: !!s?.featureToggles?.boss,
      stamp: !!s?.featureToggles?.stamp,
    },
    goalsByUnit: {},
  };

  // bossCostByUnit / goalsByUnit は unit 1..MAX を強制
  for (let u = 1; u <= units; u++) {
    out.bossCostByUnit[String(u)] = Math.max(0, Math.floor(num(s?.bossCostByUnit?.[String(u)], 10)));
    out.goalsByUnit[String(u)] = Math.max(0, Math.floor(num(s?.goalsByUnit?.[String(u)], 20)));
  }

  // coinRulesByMode は最低3つ用意（他もあってOK）
  const modes = ['practice', 'challenge', 'boss', 'focus', 'review', 'stamp'];
  for (const m of modes) {
    const r = s?.coinRulesByMode?.[m] || {};
    out.coinRulesByMode[m] = {
      miss0: Math.floor(num(r.miss0, (m === 'practice' || m === 'challenge' || m === 'boss') ? 5 : 0)),
      miss1: Math.floor(num(r.miss1, (m === 'practice' || m === 'challenge' || m === 'boss') ? 3 : 0)),
      miss2: Math.floor(num(r.miss2, (m === 'practice' || m === 'challenge' || m === 'boss') ? 1 : 0)),
      else: Math.floor(num(r.else, 0)),
    };
  }

  return out;
}

function _getMaxUnits_() {
  const units = (typeof MAX_UNITS !== 'undefined') ? Number(MAX_UNITS) : 8;
  return Number.isFinite(units) && units > 0 ? units : 8;
}

// ---------- 生徒一覧＆管理 ----------
function teacherListStudents(adminKey) {
  _requireAdmin_(adminKey);
  const ss = _openSs_();
  const sh = ss.getSheetByName('students');
  if (!sh) throw new Error('students シートが見つからない');

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(v => String(v || '').trim());
  const idx = _headerIndexMap_(headers);

  const students = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const studentId = _pickCell_(row, idx, ['student_id', 'id'], '');
    if (!studentId) continue;

    const name = _pickCell_(row, idx, ['name', 'full_name', 'nickname'], '');
    const kana = _pickCell_(row, idx, ['full_name_kana', 'kana'], '');
    const enabledRaw = _pickCell_(row, idx, ['enabled'], true);

    const enabled = (String(enabledRaw).toLowerCase() === 'true' || enabledRaw === true || enabledRaw === 1 || String(enabledRaw) === '1');

    // coin_u1.. / unlock_challenge_u1.. は無ければ 0/false
    const units = _getMaxUnits_();
    const coinsByUnit = {};
    const unlockByUnit = {};
    for (let u = 1; u <= units; u++) {
      coinsByUnit[u] = Number(_pickCell_(row, idx, [`coin_u${u}`], 0)) || 0;
      unlockByUnit[u] = String(_pickCell_(row, idx, [`unlock_challenge_u${u}`], '0')) === '1'
        || String(_pickCell_(row, idx, [`unlock_challenge_u${u}`], 'false')).toLowerCase() === 'true'
        || _pickCell_(row, idx, [`unlock_challenge_u${u}`], false) === true;
    }

    students.push({ studentId: String(studentId), name: String(name), kana: String(kana), enabled, coinsByUnit, unlockByUnit });
  }
  return students;
}

function teacherUpdateStudent(adminKey, patch) {
  _requireAdmin_(adminKey);

  const studentId = String(patch?.studentId || '');
  if (!studentId) throw new Error('studentId がない');

  const ss = _openSs_();
  const sh = ss.getSheetByName('students');
  if (!sh) throw new Error('students シートが見つからない');

  const range = sh.getDataRange();
  const values = range.getValues();
  const headers = values[0].map(v => String(v || '').trim());
  const idx = _headerIndexMap_(headers);

  const idCol = idx['student_id'] ?? idx['id'];
  if (idCol == null) throw new Error('studentsに student_id か id 列が必要');

  let targetRow = -1;
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idCol]) === studentId) { targetRow = r + 1; break; } // 1-index row
  }
  if (targetRow === -1) throw new Error('その studentId が見つからない');

  // enabled
  if (patch?.enabled != null && idx['enabled'] != null) {
    sh.getRange(targetRow, idx['enabled'] + 1).setValue(!!patch.enabled);
  }

  // ふりがな修正（full_name_kana）
  if (patch?.full_name_kana != null) {
    const col = idx['full_name_kana'] ?? idx['kana'];
    if (col != null) sh.getRange(targetRow, col + 1).setValue(String(patch.full_name_kana));
  }

  // 名前修正（full_name）
  if (patch?.full_name != null) {
    const col = idx['full_name'] ?? idx['name'];
    if (col != null) sh.getRange(targetRow, col + 1).setValue(String(patch.full_name));
  }

  // PWリセット（4ケタ）
  if (patch?.login_pass != null) {
    const col = idx['login_pass'] ?? idx['pass'] ?? idx['password'];
    if (col != null) {
      const pass = String(patch.login_pass).replace(/[^\d]/g, '');
      if (pass.length !== 4) throw new Error('PWは数字4ケタにしてね');
      sh.getRange(targetRow, col + 1).setValue(pass);
    }
  }

  return { ok: true };
}

function teacherSaveStudentTtsOverride(adminKey, studentId, normal, slow) {
  _requireAdmin_(adminKey);
  _ensureDefaultSettings_();

  const props = PropertiesService.getScriptProperties();
  const s = JSON.parse(props.getProperty(PROP_SETTINGS_JSON));

  const n = Number(normal);
  const sl = Number(slow);
  if (!Number.isFinite(n) || !Number.isFinite(sl)) throw new Error('数値がおかしいよ');

  s.ttsOverridesByStudentId = s.ttsOverridesByStudentId || {};
  s.ttsOverridesByStudentId[String(studentId)] = {
    normal: Math.min(2.0, Math.max(0.5, n)),
    slow: Math.min(1.5, Math.max(0.3, sl)),
  };

  props.setProperty(PROP_SETTINGS_JSON, JSON.stringify(s));
  return { ok: true };
}

// ---------- ダッシュボード：クラス傾向 ----------
function teacherGetClassTrends(adminKey) {
  _requireAdmin_(adminKey);

  const ss = _openSs_();
  const stSh = ss.getSheetByName('students');
  const rsSh = ss.getSheetByName('results');
  if (!stSh) throw new Error('students シートが見つからない');
  if (!rsSh) throw new Error('results シートが見つからない');

  const units = _getMaxUnits_();

  // 平均コイン（studentsから）
  const stValues = stSh.getDataRange().getValues();
  const stHeaders = stValues[0].map(v => String(v || '').trim());
  const stIdx = _headerIndexMap_(stHeaders);

  const sumCoins = Array(units + 1).fill(0);
  const cnt = Array(units + 1).fill(0);

  for (let r = 1; r < stValues.length; r++) {
    const row = stValues[r];
    for (let u = 1; u <= units; u++) {
      const c = Number(_pickCell_(row, stIdx, [`coin_u${u}`], 0)) || 0;
      sumCoins[u] += c;
      cnt[u] += 1;
    }
  }

  const avgCoinsByUnit = {};
  for (let u = 1; u <= units; u++) {
    avgCoinsByUnit[u] = cnt[u] ? Math.round((sumCoins[u] / cnt[u]) * 10) / 10 : 0;
  }

  // 平均正答率＆つまずきTOP10（resultsから）
  const rsValues = rsSh.getDataRange().getValues();
  const rsHeaders = rsValues[0].map(v => String(v || '').trim());
  const rsIdx = _headerIndexMap_(rsHeaders);

  const colUnit = rsIdx['unit'];
  const colCorrect = rsIdx['correct'];
  const colItem = rsIdx['item_id'];
  const colTime = rsIdx['timestamp'];

  if (colUnit == null || colCorrect == null || colItem == null) {
    throw new Error('results に unit / correct / item_id が必要');
  }

  const totalByUnit = Array(units + 1).fill(0);
  const correctByUnit = Array(units + 1).fill(0);

  const wrongCountByItem = new Map();

  const now = Date.now();
  const days30 = 30 * 24 * 60 * 60 * 1000;

  for (let r = 1; r < rsValues.length; r++) {
    const row = rsValues[r];
    const u = Number(row[colUnit]) || 0;
    if (u < 1 || u > units) continue;

    totalByUnit[u] += 1;

    const correctVal = row[colCorrect];
    const correct = (correctVal === true) || String(correctVal).toLowerCase() === 'true' || String(correctVal) === '1';
    if (correct) correctByUnit[u] += 1;

    // つまずきTOP：直近30日＆不正解のみ
    const ts = _asTimeMs_(row[colTime]);
    if (ts && (now - ts) <= days30 && !correct) {
      const itemId = String(row[colItem] ?? '');
      if (itemId) wrongCountByItem.set(itemId, (wrongCountByItem.get(itemId) || 0) + 1);
    }
  }

  const avgAccuracyByUnit = {};
  for (let u = 1; u <= units; u++) {
    avgAccuracyByUnit[u] = totalByUnit[u] ? Math.round((correctByUnit[u] / totalByUnit[u]) * 1000) / 10 : 0;
  }

  const topWrongItems = [...wrongCountByItem.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10)
    .map(([itemId, wrongCount]) => ({ itemId, wrongCount }));

  return { avgCoinsByUnit, avgAccuracyByUnit, topWrongItems };
}

// ---------- ダッシュボード：生徒個別 ----------
function teacherGetStudentDetail(adminKey, studentId) {
  _requireAdmin_(adminKey);

  const ss = _openSs_();
  const rsSh = ss.getSheetByName('results');
  if (!rsSh) throw new Error('results シートが見つからない');

  const values = rsSh.getDataRange().getValues();
  if (values.length < 2) return { studentId, byUnit: {}, byMode: {} };

  const headers = values[0].map(v => String(v || '').trim());
  const idx = _headerIndexMap_(headers);

  const colStudent = idx['student_id'];
  const colUnit = idx['unit'];
  const colCorrect = idx['correct'];
  const colMode = idx['mode'];
  const colItem = idx['item_id'];

  if (colStudent == null || colUnit == null || colCorrect == null || colMode == null) {
    throw new Error('results に student_id / unit / correct / mode が必要');
  }

  const units = _getMaxUnits_();
  const byUnit = {};
  for (let u = 1; u <= units; u++) {
    byUnit[u] = { attempts: 0, correct: 0, wrong: 0 };
  }

  const byMode = {};
  const wrongByItem = new Map();

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (String(row[colStudent]) !== String(studentId)) continue;

    const u = Number(row[colUnit]) || 0;
    if (u >= 1 && u <= units) {
      byUnit[u].attempts += 1;
    }

    const correctVal = row[colCorrect];
    const correct = (correctVal === true) || String(correctVal).toLowerCase() === 'true' || String(correctVal) === '1';
    if (u >= 1 && u <= units) {
      if (correct) byUnit[u].correct += 1;
      else byUnit[u].wrong += 1;
    }

    const mode = String(row[colMode] || '');
    byMode[mode] = (byMode[mode] || 0) + 1;

    if (!correct) {
      const itemId = String(row[idx['item_id']] ?? '');
      if (itemId) wrongByItem.set(itemId, (wrongByItem.get(itemId) || 0) + 1);
    }
  }

  const topWrongItems = [...wrongByItem.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10)
    .map(([itemId, wrongCount]) => ({ itemId, wrongCount }));

  return { studentId: String(studentId), byUnit, byMode, topWrongItems };
}

// ---------- 生徒側：設定を返す（各ゲーム画面が今後使える） ----------
function getAppSettingsForToken(token) {
  _ensureDefaultSettings_();
  const props = PropertiesService.getScriptProperties();
  const s = JSON.parse(props.getProperty(PROP_SETTINGS_JSON));

  const student = _getStudentByTokenSafe_(token);
  const studentId = student?.student_id || student?.id || '';

  const override = (studentId && s.ttsOverridesByStudentId && s.ttsOverridesByStudentId[String(studentId)])
    ? s.ttsOverridesByStudentId[String(studentId)]
    : null;

  return {
    tts: override ? override : s.tts,
    featureToggles: s.featureToggles,
    bossCostByUnit: s.bossCostByUnit,
    coinRulesByMode: s.coinRulesByMode,
    stampThresholds: s.stampThresholds,
    goalsByUnit: s.goalsByUnit,
    updatedAt: s.updatedAt,
  };
}

// ---------- 内部ユーティリティ ----------
function _openSs_() {
  const id = (typeof SS_ID !== 'undefined') ? SS_ID : null;
  if (!id) throw new Error('SS_ID が Code.gs に無いみたい');
  return SpreadsheetApp.openById(id);
}

function _headerIndexMap_(headers) {
  const map = {};
  headers.forEach((h, i) => { if (h) map[String(h)] = i; });
  return map;
}

function _pickCell_(row, idxMap, keys, def) {
  for (const k of keys) {
    if (idxMap[k] != null) return row[idxMap[k]];
  }
  return def;
}

function _asTimeMs_(v) {
  if (!v) return 0;
  if (v instanceof Date) return v.getTime();
  const n = Number(v);
  if (Number.isFinite(n) && n > 1000000000) return n; // たぶんms
  const d = new Date(v);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

function _getStudentByTokenSafe_(token) {
  try {
    // 既存関数があるならそれを使う
    if (typeof getStudentByToken === 'function') return getStudentByToken(token);
  } catch (e) {}

  // なければ sessions シートから拾う（最低限）
  const ss = _openSs_();
  const sessSh = ss.getSheetByName('sessions');
  const stuSh = ss.getSheetByName('students');
  if (!sessSh || !stuSh) return null;

  const sVals = sessSh.getDataRange().getValues();
  if (sVals.length < 2) return null;
  const sHead = sVals[0].map(v => String(v || '').trim());
  const sIdx = _headerIndexMap_(sHead);

  const colToken = sIdx['token'];
  const colStudentId = sIdx['student_id'];
  if (colToken == null || colStudentId == null) return null;

  let studentId = '';
  for (let r = 1; r < sVals.length; r++) {
    if (String(sVals[r][colToken]) === String(token)) {
      studentId = String(sVals[r][colStudentId] || '');
      break;
    }
  }
  if (!studentId) return null;

  const stVals = stuSh.getDataRange().getValues();
  if (stVals.length < 2) return null;
  const stHead = stVals[0].map(v => String(v || '').trim());
  const stIdx = _headerIndexMap_(stHead);
  const colId = stIdx['student_id'] ?? stIdx['id'];
  if (colId == null) return null;

  for (let r = 1; r < stVals.length; r++) {
    if (String(stVals[r][colId]) === studentId) {
      return { student_id: studentId, row: r + 1 };
    }
  }
  return null;
}
