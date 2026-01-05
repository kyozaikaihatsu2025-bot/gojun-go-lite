/**
 * items_db の
 * A列: id
 * B列: 日本語（prompt）
 * C列: 英語（correct_sentence）
 * を Wh / S / Aux / V / O/C / Adv にざっくり分けて別シートに書き出す（ID列つき）
 */
function buildItemsRoles() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const src = ss.getSheetByName(SHEET_ITEMS);
  if (!src) throw new Error('items_db シートが見つかりません');

  const values = src.getDataRange().getValues();
  if (values.length <= 1) return;

  const jpSheet = ss.getSheetByName(SHEET_ROLES_JP) || ss.insertSheet(SHEET_ROLES_JP);
  const enSheet = ss.getSheetByName(SHEET_ROLES_EN) || ss.insertSheet(SHEET_ROLES_EN);
  jpSheet.clearContents();
  enSheet.clearContents();

  const header = [['id', 'Wh', 'S', 'Aux', 'V', 'O/C', 'Adv']];
  jpSheet.getRange(1, 1, 1, 7).setValues(header);
  enSheet.getRange(1, 1, 1, 7).setValues(header);

  const jpRows = [];
  const enRows = [];

  for (let r = 1; r < values.length; r++) {
    const id = String(values[r][0] || '').trim();
    if (!id) continue;

    const jp = String(values[r][1] || '').trim();
    const en = String(values[r][2] || '').trim();

    // 日本語の「に」は英語ヒントも使う
    const jpParts = splitJapaneseRoles(jp, en);
    const enParts = splitEnglishRoles(en);

    jpRows.push([id, jpParts.Wh, jpParts.S, jpParts.Aux, jpParts.V, jpParts.OC, jpParts.Adv]);
    enRows.push([id, enParts.Wh, enParts.S, enParts.Aux, enParts.V, enParts.OC, enParts.Adv]);
  }

  if (jpRows.length) jpSheet.getRange(2, 1, jpRows.length, 7).setValues(jpRows);
  if (enRows.length) enSheet.getRange(2, 1, enRows.length, 7).setValues(enRows);
}

/* ========= 日本語側 ========= */
function splitJapaneseRoles(text, enText) {
  let Wh = '', S = '', Aux = '', V = '', OC = '', Adv = '';
  if (!text) return { Wh, S, Aux, V, OC, Adv };

  let t = String(text).trim();

  // カタカナの「ー」「・」も 1語として扱うために含める
  const JP_WORD = '[一-龥ぁ-んァ-ンー・A-Za-z0-9０-９]+';

  // --- Wh 抜き出し ---
  const whList = ['どうして', 'なぜ', 'どちら', 'どれ', 'だれ', '誰', 'どこ', 'いつ', 'どう', '何', 'なん'];

  const findWhIndex = (str, w) => {
    if (w !== 'いつ') return str.indexOf(w);
    // 「いつ」でも、直後が「も」「で」なら疑問詞じゃない（いつも/いつでも等）
    let from = 0;
    while (true) {
      const idx = str.indexOf('いつ', from);
      if (idx < 0) return -1;
      const next = str.charAt(idx + 2);
      if (next === 'も' || next === 'で') {
        from = idx + 1;
        continue;
      }
      return idx;
    }
  };

  const takeParticleAfterWh = (base, idx, wLen) => {
    const rest = base.slice(idx + wLen);

    // 2文字助詞
    if (rest.startsWith('から')) return { token: 'から', len: 2 };
    if (rest.startsWith('まで')) return { token: 'まで', len: 2 };

    // 1文字助詞
    const c = rest.charAt(0);
    const one = ['を', 'に', 'へ', 'で', 'と'];
    if (one.includes(c)) {
      // なんですか / どれですか 等の「で」は助詞じゃない（=「です」の一部）
      if (c === 'で' && (rest.startsWith('です') || rest.startsWith('でした') || rest.startsWith('でし'))) {
        return { token: '', len: 0 };
      }
      return { token: c, len: 1 };
    }
    return { token: '', len: 0 };
  };

  while (true) {
    let foundWord = null;
    let foundIdx = -1;
    for (const w of whList) {
      const idx = findWhIndex(t, w);
      if (idx >= 0 && (foundIdx < 0 || idx < foundIdx)) {
        foundIdx = idx;
        foundWord = w;
      }
    }
    if (foundIdx < 0) break;

    // 何を / どこに など「疑問詞＋助詞」をひとかたまりで Wh に入れる
    const p = takeParticleAfterWh(t, foundIdx, foundWord.length);
    const chunk = foundWord + p.token;
    Wh += (Wh ? '／' : '') + chunk;
    t = t.slice(0, foundIdx) + t.slice(foundIdx + foundWord.length + p.len);
  }

  if (t.endsWith('か')) {
    Wh += (Wh ? '／' : '') + 'か';
    t = t.slice(0, -1);
  }

  // 「たくさん/いっぱい」は Adv にせず O/C（名詞・形容詞側）へ
  const qtyList = ['たくさん', 'いっぱい'];
  for (const q of qtyList) {
    let from = 0;
    while (true) {
      const idx = t.indexOf(q, from);
      if (idx < 0) break;
      const nextChar = t.charAt(idx + q.length);
      const chunk = (nextChar === 'の') ? (q + 'の') : q;
      OC += chunk;
      t = t.slice(0, idx) + t.slice(idx + chunk.length);
      from = idx;
    }
  }

  // --- 典型的な副詞をざっくり除去 ---
  const advList = [
    'とても', 'すごく', 'かなり', '少し', 'すこし', 'あまり', 'あんまり', '全然', 'ぜんぜん',
    'よく', 'いつも', 'ふだん', '普段', 'たいてい', '時々', 'ときどき',
    '早く', 'はやく', 'ゆっくり', '上手に', 'じょうずに', '下手に', 'へたに', '速く', '高く',
    '一生懸命', 'いっしょうけんめい',
    '毎日', 'まいにち', '毎週', 'まいしゅう', '毎月', 'まいつき', '毎年', 'まいとし',
    '昨日', 'きのう', '今日', 'きょう', '明日', 'あした',
    '今', 'いま', 'さっき', 'あとで', 'もうすぐ', 'すぐ',
    'ぜったい', '絶対', 'たぶん', '多分',
    'ずっと', 'もっと', 'また', 'もう', 'まだ',
    'みんなで', '一緒に', 'いっしょに',
    '静かに', 'しずかに', '元気に', 'げんきに',
    '丁寧に', 'ていねいに', '簡単に', 'かんたんに'
  ];

  // 『買いました/手伝いました』の中の「いま」を副詞にしない（境界チェック強化）
  for (const a of advList) {
    if (a === 'いま') {
      let from = 0;
      while (true) {
        const idx = t.indexOf(a, from);
        if (idx < 0) break;

        const nextChar = t.charAt(idx + a.length);
        const next3 = t.slice(idx, idx + 3);
        const next4 = t.slice(idx, idx + 4);

        const isPoliteEndingPart =
          (nextChar === 'し' || nextChar === 'す' || nextChar === 'せ') ||
          (next3 === 'います') ||
          (next4 === 'いました') ||
          (next4 === 'いません');

        if (isPoliteEndingPart) {
          from = idx + 1;
          continue;
        }

        Adv += (Adv ? '／' : '') + a;
        t = t.slice(0, idx) + t.slice(idx + a.length);
        from = idx;
      }
      continue;
    }

    while (true) {
      const idx = t.indexOf(a);
      if (idx < 0) break;
      Adv += (Adv ? '／' : '') + a;
      t = t.slice(0, idx) + t.slice(idx + a.length);
    }
  }

  // --- 主語：最初の「は」「が」まで ---
  let idxHa = t.indexOf('は');
  let idxGa = t.indexOf('が');
  let subjIdx = -1;
  if (idxHa >= 0 && idxGa >= 0) subjIdx = Math.min(idxHa, idxGa);
  else subjIdx = (idxHa >= 0 ? idxHa : idxGa);

  if (subjIdx >= 0) {
    S = t.slice(0, subjIdx + 1);
    t = t.slice(subjIdx + 1);
  }

  // ✅役に立ちます：役に立ち=O/C、ます=V
  {
    const m = t.match(/(役に立ち)(ませんでした|ません|ました|ます)/);
    if (m) {
      OC += (OC ? '／' : '') + m[1];
      V = m[2];
      const rest = t.replace(m[0], '');
      if (rest) OC += rest;
      return { Wh, S, Aux, V, OC, Adv };
    }
  }

  // 英語の「動詞の直後」から、日本語の「〜に」を Adv/O/C どっち寄せか判断する
  const niHint = _classifyNiByEnglish_(enText);

  // 〇〇のあとに / 〇〇のあとで を 1まとまりで Adv
  const afterPhrasePatterns = [
    new RegExp(JP_WORD + 'のあとに', 'g'),
    new RegExp(JP_WORD + 'の後に', 'g'),
    new RegExp(JP_WORD + 'のあとで', 'g'),
    new RegExp(JP_WORD + 'の後で', 'g')
  ];
  for (const pat of afterPhrasePatterns) {
    t = t.replace(pat, (match) => {
      Adv += (Adv ? '／' : '') + match;
      return '';
    });
  }

  // --- 場所・出身などの副詞句（名詞＋助詞） ---
  const locPhrasePatterns = [
    new RegExp(JP_WORD + '出身', 'g'),
    new RegExp(JP_WORD + 'の' + JP_WORD + 'で(?!し|す|き)', 'g'),
    new RegExp(JP_WORD + 'で(?!し|す|き)', 'g'),
    new RegExp(JP_WORD + 'へ', 'g'),
    new RegExp(JP_WORD + 'から', 'g'),
    new RegExp(JP_WORD + 'まで', 'g')
  ];
  for (const pat of locPhrasePatterns) {
    t = t.replace(pat, (match) => {
      Adv += (Adv ? '／' : '') + match;
      return '';
    });
  }

  // 時間・季節などは確実に Adv
  const timeNiPatterns = [
    /[0-9０-９一二三四五六七八九十]+時[0-9０-９一二三四五六七八九十]+分に/g,
    /[0-9０-９一二三四五六七八九十]+時半に/g,
    /[0-9０-９一二三四五六七八九十]+(時|分|日|月|年)に/g,
    new RegExp(JP_WORD + '日に', 'g'),
    /(月曜日|火曜日|水曜日|木曜日|金曜日|土曜日|日曜日)に/g,
    /(きょう|今日|あした|明日|きのう|昨日|元日)に/g,
    /(朝|昼|夜|夕方|午前|午後)に/g,
    /(春|夏|秋|冬)に/g,
    /(今週|来週|先週|今月|来月|先月|今年|来年|去年)に/g
  ];
  for (const pat of timeNiPatterns) {
    t = t.replace(pat, (match) => {
      Adv += (Adv ? '／' : '') + match;
      return '';
    });
  }

  // 名詞＋に の振り分け
  const niPattern = new RegExp(JP_WORD + 'に', 'g');
  t = t.replace(niPattern, (match, offset, whole) => {
    const rest = whole.slice(offset + match.length);
    const hasWoAhead = rest.includes('を');

    if (niHint === 'adv' || hasWoAhead) {
      Adv += (Adv ? '／' : '') + match;
    } else {
      OC += (OC ? '／' : '') + match;
    }
    return '';
  });

  const withPhrasePatterns = [
    /(父|母|お父さん|お母さん|友達|友だち|家族|弟|妹|兄|姉|先生|みんな|犬|猫)と/g
  ];
  for (const pat of withPhrasePatterns) {
    t = t.replace(pat, (match) => {
      Adv += (Adv ? '／' : '') + match;
      return '';
    });
  }

  // 継続時間
  const durationPatterns = [
    /[0-9０-９一二三四五六七八九十]+(分間|分)/g,
    /[0-9０-９一二三四五六七八九十]+(時間半|時間)/g,
    /[0-9０-９一二三四五六七八九十]+(日間|日)/g,
    /[0-9０-９一二三四五六七八九十]+(週間|週)/g,
    /[0-9０-９一二三四五六七八九十]+(か月間|ヶ月間|か月|ヶ月)/g,
    /[0-9０-９一二三四五六七八九十]+(年間|年)(?!生)/g
  ];
  for (const pat of durationPatterns) {
    t = t.replace(pat, (match, _g1, offset, whole) => {
      const next = whole.charAt(offset + match.length);
      if (next === 'を') return match;
      Adv += (Adv ? '／' : '') + match;
      return '';
    });
  }

  // 「です/でした/ます/ました」検出
  const candidates = [
    { token: 'でした', type: 'desu', len: 3, idx: t.indexOf('でした') },
    { token: 'です', type: 'desu', len: 2, idx: t.indexOf('です') },
    { token: 'ました', type: 'masu', len: 3, idx: t.indexOf('ました') },
    { token: 'ます', type: 'masu', len: 2, idx: t.indexOf('ます') }
  ];

  let verbIdx = -1;
  let verbType = '';
  let verbToken = '';
  let verbLen = 0;

  for (const c of candidates) {
    if (c.idx < 0) continue;
    if (verbIdx < 0 || c.idx < verbIdx) {
      verbIdx = c.idx;
      verbType = c.type;
      verbToken = c.token;
      verbLen = c.len;
    }
  }

  if (verbIdx < 0) {
    OC += t;
    return { Wh, S, Aux, V, OC, Adv };
  }

  const before = t.slice(0, verbIdx);
  const after = t.slice(verbIdx + verbLen);

  // --- 「△△することができます」 ---
  const kotoDekimasu = 'ことができます';
  const idxKD = t.indexOf(kotoDekimasu);
  if (idxKD >= 0) {
    const beforeKD = t.slice(0, idxKD);
    const woIdx = beforeKD.lastIndexOf('を');
    if (woIdx >= 0) {
      OC += beforeKD.slice(0, woIdx + 1);
      V = beforeKD.slice(woIdx + 1) + 'ことが';
    } else {
      V = t.slice(0, idxKD + 'ことが'.length);
    }
    Aux = 'できます';
    const tail = t.slice(idxKD + kotoDekimasu.length);
    if (tail) OC += tail;
    return { Wh, S, Aux, V, OC, Adv };
  }

  // --- 「〜できます」 ---
  const idxDekimasu = t.indexOf('できます');
  if (idxDekimasu >= 0) {
    const beforeDek = t.slice(0, idxDekimasu);
    const lastWo = beforeDek.lastIndexOf('を');
    if (lastWo >= 0) {
      OC += beforeDek.slice(0, lastWo + 1);
      V = beforeDek.slice(lastWo + 1);
    } else {
      V = beforeDek;
    }
    Aux = 'できます';
    const tail = t.slice(idxDekimasu + 4);
    if (tail) OC += tail;
    return { Wh, S, Aux, V, OC, Adv };
  }

  // --- 「〜です／でした」系 ---
  if (verbType === 'desu') {
    // 〜したいです / 〜見たいです は「たい＋です」までを動詞(V)にする
    const taiEndingRe = /(たい|たかった|たくない|たくなかった)$/;
    if (taiEndingRe.test(before)) {
      const lastWo = before.lastIndexOf('を');
      if (lastWo >= 0) {
        OC += before.slice(0, lastWo + 1);
        V = before.slice(lastWo + 1) + verbToken;
      } else {
        V = before + verbToken;
      }
      if (after) OC += after;
      return { Wh, S, Aux, V, OC, Adv };
    }

    // 幸せでした → OC=幸せ, V=でした
    const happinessWords = ['幸せ', 'しあわせ'];
    for (const hw of happinessWords) {
      const idx = before.lastIndexOf(hw);
      if (idx >= 0 && idx + hw.length === before.length) {
        const head = before.slice(0, idx);
        if (head) OC += head;
        OC += hw;
        V = verbToken;
        if (after) OC += after;
        return { Wh, S, Aux, V, OC, Adv };
      }
    }

    // ✅ここが今回の修正：好き/大好き は「好きです」丸ごと V にする
    // 根拠：この教材では like/love 相当として動詞枠に入れたほうが分かりやすい
    const likeWords = ['大好き', 'だいすき', '好き', 'すき', '大嫌い', 'だいきらい', '嫌い', 'きらい'];
    let foundLike = '';
    for (const w of likeWords) {
      if (before.endsWith(w) && w.length > foundLike.length) foundLike = w;
    }
    if (foundLike) {
      const cut = before.length - foundLike.length;
      const npPart = before.slice(0, cut);
      if (npPart) OC += npPart;   // 例：ネコが
      V = foundLike + verbToken;  // 例：好きです / 大好きです
      if (after) OC += after;
      return { Wh, S, Aux, V, OC, Adv };
    }

    const lexList = [
      '好き', 'すき', '大好き', 'だいすき', '上手', 'じょうず', '下手', 'へた', '嫌い', 'きらい', '得意', 'とくい',
      '元気', 'げんき', '静か', 'しずか', 'きれい', '有名', 'ゆうめい', '親切', 'しんせつ',
      '大切', 'たいせつ', '大事', 'だいじ', '便利', 'べんり', '不便', 'ふべん',
      '安全', 'あんぜん', '危険', 'きけん', '簡単', 'かんたん', '難しい', 'むずかしい',
      '大変', 'たいへん', '暇', 'ひま', '忙しい', 'いそがしい',
      '楽しい', 'たのしい', 'うれしい', 'かなしい', 'こわい',
      '暑い', 'あつい', '寒い', 'さむい', '熱い', 'あつい', '冷たい', 'つめたい',
      '大きい', 'おおきい', '小さい', 'ちいさい', '長い', 'ながい', '短い', 'みじかい',
      '高い', 'たかい', '低い', 'ひくい', '新しい', 'あたらしい', '古い', 'ふるい',
      '多い', 'おおい', '少ない', 'すくない', '早い', 'はやい', '遅い', 'おそい',
      '幸せ', 'しあわせ',
      'あたたかい'
    ];

    let foundAdj = '';
    for (const w of lexList) {
      if (before.endsWith(w) && w.length > foundAdj.length) {
        foundAdj = w;
      }
    }

    // 形容詞＋です は「形容詞は O/C」「です/でした は V」に分ける（従来どおり）
    if (foundAdj) {
      const cut = before.length - foundAdj.length;
      const npPart = before.slice(0, cut);
      if (npPart) OC += npPart;
      OC += foundAdj;
      V = verbToken;
      if (after) OC += after;
      return { Wh, S, Aux, V, OC, Adv };
    }

    if (before) OC += before;
    V = verbToken;
    if (after) OC += after;
    return { Wh, S, Aux, V, OC, Adv };
  }

  // --- 「〜ます／ました」系 ---
  if (verbType === 'masu') {
    const lastWo = before.lastIndexOf('を');
    if (lastWo >= 0) {
      OC += before.slice(0, lastWo + 1);
      V = before.slice(lastWo + 1) + verbToken;
    } else {
      const lastGa = before.lastIndexOf('が');
      const lastHa2 = before.lastIndexOf('は');
      const cutIdx = Math.max(lastGa, lastHa2);

      if (cutIdx > 0 && cutIdx < before.length - 1) {
        OC += before.slice(0, cutIdx + 1);
        V = before.slice(cutIdx + 1) + verbToken;
      } else {
        V = before + verbToken;
      }
    }
    if (after) OC += after;
    return { Wh, S, Aux, V, OC, Adv };
  }

  OC += t;
  return { Wh, S, Aux, V, OC, Adv };
}

/* ========= 英語側 ========= */
function splitEnglishRoles(text) {
  let Wh = '', S = '', Aux = '', V = '', OC = '', Adv = '';
  if (!text) return { Wh, S, Aux, V, OC, Adv };

  const words = text.trim().split(/\s+/);
  if (!words.length) return { Wh, S, Aux, V, OC, Adv };

  const roles = computeRolesForItems(words);
  const clean = words.map(w => String(w || '').toLowerCase().replace(/[^a-z]/g, ''));
  const n = clean.length;

  const wantIdx = clean.indexOf('want');

  const prepSet = new Set([
    'in', 'on', 'at', 'from', 'to', 'for', 'about', 'of',
    'into', 'onto', 'before', 'after', 'around', 'between', 'behind',
    'under', 'over', 'near', 'with'
  ]);

  for (let i = 0; i < n; i++) {
    if (roles[i] === 'verb') continue; // want to の to を Adv で上書きしない

    const w = clean[i];
    if (!prepSet.has(w)) continue;

    // want の直後の to だけは「不定詞to」なので前置詞扱いしない
    if (w === 'to' && wantIdx >= 0 && i === wantIdx + 1) continue;

    roles[i] = 'adv';
    for (let j = i + 1; j < n; j++) {
      if (roles[j] === 'verb') break;
      if (roles[j] === 'obj') roles[j] = 'adv';
      else break;
    }
  }

  const numWordSet = new Set(['one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten', 'half']);
  const unitSet = new Set(['minute', 'minutes', 'hour', 'hours', 'day', 'days', 'week', 'weeks', 'month', 'months', 'year', 'years']);

  for (let i = 0; i < n; i++) {
    if (clean[i] !== 'for') continue;
    if (roles[i] === 'verb') continue;
    roles[i] = 'adv';

    for (let j = i + 1; j < n; j++) {
      if (roles[j] === 'verb') break;

      const raw = String(words[j] || '');
      const cj = clean[j];
      const isNumRaw = /^[0-9]+$/.test(raw);
      const isNumCleanEmpty = (cj === '' && /[0-9]/.test(raw));
      const isNumWord = numWordSet.has(cj);
      const isUnit = unitSet.has(cj);
      if (!(isNumRaw || isNumCleanEmpty || isNumWord || isUnit)) break;

      if (roles[j] === 'obj' || roles[j] === 'adv') roles[j] = 'adv';
      else break;
    }
  }

  const whArr = [], sArr = [], auxArr = [], vArr = [], ocArr = [], advArr = [];
  words.forEach((w, i) => {
    const role = roles[i];
    if (role === 'wh') whArr.push(w);
    else if (role === 'subj') sArr.push(w);
    else if (role === 'aux') auxArr.push(w);
    else if (role === 'verb') vArr.push(w);
    else if (role === 'adv') advArr.push(w);
    else ocArr.push(w);
  });

  Wh = whArr.join(' ');
  S = sArr.join(' ');
  Aux = auxArr.join(' ');
  V = vArr.join(' ');
  OC = ocArr.join(' ');
  Adv = advArr.join(' ');

  return { Wh, S, Aux, V, OC, Adv };
}

function computeRolesForItems(words) {
  const clean = words.map(w => String(w || '').toLowerCase().replace(/[^a-z]/g, ''));
  const n = clean.length;
  const roles = new Array(n).fill(null);

  const whSet = new Set(['what', 'where', 'when', 'who', 'which', 'how', 'why']);
  const auxSet = new Set(['can', 'will', 'would', 'could', 'shall', 'should', 'might', 'must', 'do', 'does', 'did']);
  const beSet = new Set(['am', 'is', 'are', 'was', 'were', 'be', 'been', 'being']);

  const irregularPastSet = new Set([
    'was', 'were',
    'did', 'had',
    'went', 'came',
    'saw', 'ate', 'drank',
    'got', 'gave', 'took', 'made',
    'said', 'told', 'heard',
    'read', 'wrote', 'spoke',
    'ran', 'swam', 'sang',
    'slept', 'sat', 'stood',
    'met', 'bought', 'brought',
    'thought', 'taught', 'caught',
    'knew', 'found', 'left',
    'kept', 'lost', 'paid', 'sent',
    'wore', 'won', 'understood',
    'fell', 'felt', 'held', 'cut', 'put',
    'drew', 'drove', 'flew', 'forgot',
    'grew', 'became', 'began', 'broke', 'chose', 'fed', 'led'
  ]);

  const verbSet = new Set([
    'do', 'check',
    'like', 'play', 'go', 'draw', 'eat', 'drink', 'live', 'study', 'know',
    'think', 'want', 'make', 'swim', 'run', 'cook', 'write', 'speak', 'walk',
    'have', 'has', 'had', 'love', 'help', 'sing', 'read',
    'see', 'look', 'watch', 'hear', 'listen', 'say', 'tell', 'talk', 'speak',
    'come', 'get', 'give', 'take', 'use', 'need', 'open', 'close', 'put', 'bring',
    'buy', 'sell', 'pick', 'visit', 'try', 'wear', 'throw', 'wash', 'clean',
    'build', 'find', 'meet', 'call', 'work', 'learn',
    'ride', 'drive', 'travel', 'stay', 'start', 'finish', 'end', 'begin',
    'join', 'practice', 'enjoy', 'forget', 'pack', 'brush'
  ]);

  const advSet = new Set([
    'well', 'fast', 'slowly', 'quickly', 'usually', 'always', 'often', 'sometimes', 'never',
    'here', 'there', 'today', 'yesterday', 'tomorrow', 'now', 'late', 'early', 'hard', 'high', 'very', 'much',
    'really', 'too', 'also', 'again', 'together', 'soon', 'still', 'then', 'only'
  ]);

  const lyNotAdverbSet = new Set(['family']);

  const monthSet = new Set([
    'january', 'february', 'march', 'april', 'may', 'june', 'july', 'august',
    'september', 'october', 'november', 'december'
  ]);

  const subjStarterSet = new Set(['i', 'you', 'he', 'she', 'it', 'we', 'they', 'this', 'that', 'these', 'those']);

  const whIdx = clean.findIndex(w => whSet.has(w));
  if (whIdx >= 0) {
    roles[whIdx] = 'wh';
    if (clean[whIdx] === 'what' && whIdx + 1 < n && !roles[whIdx + 1]) {
      roles[whIdx + 1] = 'obj';
    }
  }

  const isMainDo = (i) => {
    const w = clean[i];
    if (w !== 'do' && w !== 'does' && w !== 'did') return false;
    if (i === 0) return false;
    if (i + 1 >= n) return true;
    const next = clean[i + 1];
    if (subjStarterSet.has(next)) return false;
    if (next === 'not') return false;
    if (beSet.has(next) || verbSet.has(next) || next === 'to') return false;
    return true;
  };

  const auxIdxs = [];
  clean.forEach((w, i) => {
    if (!auxSet.has(w)) return;
    if ((w === 'do' || w === 'does' || w === 'did') && isMainDo(i)) return;
    roles[i] = 'aux';
    auxIdxs.push(i);
  });

  let wantToStart = -1;
  for (let i = 0; i < n - 2; i++) {
    if (clean[i] === 'want' && clean[i + 1] === 'to' && !roles[i] && !roles[i + 1]) {
      wantToStart = i;
      break;
    }
  }

  let verbIdx = -1;

  const hasBeAfter = (startIdx) => {
    for (let k = startIdx + 1; k < n; k++) {
      if (beSet.has(clean[k])) return true;
    }
    return false;
  };

  const verbScore = (i) => {
    const w = clean[i];
    if (!w) return 0;
    if (wantToStart >= 0 && i === wantToStart) return 3;
    if (irregularPastSet.has(w)) return 2;
    if (w.endsWith('ed') && w.length > 3) return 2;
    if (beSet.has(w)) return 1;
    if (verbSet.has(w)) return 1;
    return 0;
  };

  let bestScore = 0;
  let bestIdx = -1;

  for (let i = 0; i < n; i++) {
    if (roles[i]) continue;
    const score = verbScore(i);
    if (score <= 0) continue;

    if (score > bestScore) {
      bestScore = score;
      bestIdx = i;
    } else if (score === bestScore && bestIdx >= 0 && i < bestIdx) {
      bestIdx = i;
    }
  }

  if (bestIdx >= 0) {
    const w = clean[bestIdx];

    if (wantToStart >= 0 && bestIdx === wantToStart) {
      roles[bestIdx] = 'verb';
      roles[bestIdx + 1] = 'verb';

      const k = bestIdx + 2;
      if (k < n && !roles[k]) {
        const wk = clean[k];
        const isEdVerb = wk.endsWith('ed') && wk.length > 3;
        if (beSet.has(wk) || verbSet.has(wk) || irregularPastSet.has(wk) || isEdVerb) {
          roles[k] = 'verb';
        }
      }
      verbIdx = bestIdx;
    } else {
      const isEdVerbCandidate = w.endsWith('ed') && w.length > 3;
      const shouldTreatEdAsVerb = isEdVerbCandidate && !(bestIdx === 0 && hasBeAfter(0));

      if (!isEdVerbCandidate || shouldTreatEdAsVerb || irregularPastSet.has(w) || beSet.has(w) || verbSet.has(w)) {
        verbIdx = bestIdx;
        roles[verbIdx] = 'verb';
      }
    }
  }

  if (verbIdx < 0 && n >= 2) {
    const w0 = clean[0];
    const w1 = clean[1];
    const w1Raw = String(words[1] || '');
    const looksLikeSubjectStart = subjStarterSet.has(w0) || /^[A-Z]/.test(String(words[0] || ''));
    const looksLikeWord = /^[A-Za-z]/.test(w1Raw);

    if (looksLikeSubjectStart && looksLikeWord && !auxSet.has(w1) && !whSet.has(w1)) {
      const prev = clean[0];
      const safeIngVerb = w1.endsWith('ing') && (beSet.has(prev) || auxSet.has(prev) || prev === 'to');
      if (!w1.endsWith('ing') || safeIngVerb) {
        verbIdx = 1;
        roles[1] = 'verb';
      }
    }
  }

  for (let i = 0; i < n; i++) {
    if (roles[i]) continue;
    const w = clean[i];
    const original = words[i];
    const isMonth = monthSet.has(w) && /^[A-Z]/.test(original);
    if (isMonth) continue;

    const isLyAdverb = w.endsWith('ly') && w.length > 2 && !lyNotAdverbSet.has(w);
    if (advSet.has(w) || isLyAdverb) roles[i] = 'adv';
  }

  let boundary = -1;
  if (verbIdx >= 0) boundary = verbIdx;
  else if (auxIdxs.length) boundary = auxIdxs[0];

  if (boundary >= 0) {
    for (let i = 0; i < boundary; i++) {
      if (!roles[i]) roles[i] = 'subj';
    }
  }

  for (let i = 0; i < n; i++) {
    if (!roles[i]) roles[i] = 'obj';
  }

  return roles;
}

/* ========= 追加ヘルパー ========= */
function _classifyNiByEnglish_(enText) {
  const s = String(enText || '').trim();
  if (!s) return '';

  const words = s.split(/\s+/);
  const clean = words.map(w => String(w || '').toLowerCase().replace(/[^a-z]/g, ''));
  const roles = computeRolesForItems(words);
  const n = clean.length;

  const beSet = new Set(['am', 'is', 'are', 'was', 'were', 'be', 'been', 'being']);
  const prepSet = new Set([
    'in', 'on', 'at', 'from', 'to', 'for', 'about', 'of',
    'into', 'onto', 'before', 'after', 'around', 'between', 'behind',
    'under', 'over', 'near', 'with'
  ]);

  const exceptionPairs = new Set(['talk|to', 'listen|to', 'look|at', 'wait|for']);

  let vIdx = -1;
  for (let i = 0; i < n; i++) {
    if (roles[i] === 'verb' && clean[i] && clean[i] !== 'to') {
      vIdx = i;
      break;
    }
  }
  if (vIdx < 0) return '';

  if (clean[vIdx] === 'want') {
    const toIdx = vIdx + 1;
    if (toIdx < n && clean[toIdx] === 'to' && roles[toIdx] === 'verb') {
      for (let k = toIdx + 1; k < n; k++) {
        if (roles[k] === 'verb' && clean[k] && clean[k] !== 'to') {
          vIdx = k;
          break;
        }
      }
    }
  }

  const v = clean[vIdx];
  if (beSet.has(v)) return 'oc';

  let j = vIdx + 1;
  while (j < n && !clean[j]) j++;
  if (j >= n) return '';

  if (clean[j] === 'not') {
    j++;
    while (j < n && !clean[j]) j++;
    if (j >= n) return '';
  }

  const w = clean[j];

  if (prepSet.has(w)) {
    if (w === 'to') {
      const k = j + 1;
      if (k < n && roles[k] === 'verb') return '';
    }
    const key = `${v}|${w}`;
    if (exceptionPairs.has(key)) return 'oc';
    return 'adv';
  }

  if (roles[j] === 'obj') return 'oc';
  return '';
}
