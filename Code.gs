/**
 * 血圧記録アプリ（手入力版）
 *
 * 【セットアップ】
 * 設定値はコードに書かず、すべて「⚙ プロジェクトの設定 → スクリプト プロパティ」に置く。
 * こうしておくとコードを貼り替えても設定が消えない。
 *
 *   SHEET_ID    スプレッドシートURLの /d/ と /edit の間の文字列
 *   GEMINI_KEY  https://aistudio.google.com/apikey で発行したキー
 *
 * 設定できているかは checkSetup() を実行して確認できる。
 *
 * ※ 後から撮影＋自動読み取り機能を足せる構造にしてあります
 */

const SHEET_NAME = '血圧記録';

/** スクリプトプロパティから設定値を読む（未設定なら分かりやすく落とす） */
function prop_(name) {
  const v = PropertiesService.getScriptProperties().getProperty(name);
  if (!v) {
    throw new Error(name + ' が未設定です。⚙ プロジェクトの設定 → スクリプト プロパティ に追加してください');
  }
  return v;
}

/** 設定が揃っているかの確認用。エディタでこれを実行してログを見る */
function checkSetup() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('SHEET_ID');
  const key = p.getProperty('GEMINI_KEY');

  Logger.log('SHEET_ID   : ' + (id ? 'OK（…' + id.slice(-6) + '）' : '未設定'));
  Logger.log('GEMINI_KEY : ' + (key ? 'OK（…' + key.slice(-4) + '）' : '未設定'));

  if (!id) return;

  try {
    Logger.log('シート名   : ' + SpreadsheetApp.openById(id).getName() + ' を開けました');
  } catch (e) {
    Logger.log('シートを開けません。IDが違う可能性があります');
    return;
  }

  // 「基本情報」シートが無ければここで作る。AIに毎回渡す前提情報の置き場所
  getProfileSheet_();
  const profile = getProfileText_();
  Logger.log('基本情報   : ' + (profile
    ? '以下をAIに渡しています\n' + profile
    : '「基本情報」シートを用意しました。書きたい項目だけ埋めてください（空欄はAIに渡りません）'));
}

/**
 * ウェブアプリの入口
 * ・通常URL          → 入力画面（index）
 * ・?page=viewer 付き → グラフと直近の傾向（viewer）
 * ・?page=report 付き → 月次まとめと質問コーナー（report）
 */
function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) || '';
  var file  = page === 'viewer' ? 'viewer'
            : page === 'report' ? 'report'
            : page === 'print'  ? 'print'
            : 'index';
  var title = page === 'viewer' ? '血圧 HIGH VOLTAGE'
            : page === 'report' ? '血圧 レポート'
            : page === 'print'  ? '血圧記録（印刷用）'
            : '血圧記録';

  return HtmlService.createHtmlOutputFromFile(file)
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

/** 現在のウェブアプリURL（入力画面⇄ビューアのリンク用） */
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

// ===== AI分析コメント（Gemini API） =====
// APIキーは「プロジェクトの設定 → スクリプト プロパティ」に GEMINI_KEY として登録する。
// ※ HTML側には絶対に書かないこと（ページを開いた人に盗まれる）
// モデルは時々提供終了になる。404が出たら listModels() を実行して使える名前を確認する
const GEMINI_MODEL = 'gemini-3.6-flash';
const INSIGHT_CACHE_KEY = 'INSIGHT_CACHE';
const INSIGHT_INTERVAL_DAYS = 3;   // 傾向カードを作り直す間隔
const MONTHLY_SHEET = '月次まとめ';
const PROFILE_SHEET = '基本情報';

// =AI() で運用していたときの置き場所（APIが使えないときの予備）
const INSIGHT_SHEET = 'シート3';
const INSIGHT_CELL = 'F2';

/**
 * ビューアの「傾向」カード用のコメント。
 * 3日ごとに作り直す（毎回だと変化が出ず、APIも無駄に叩くため）。
 * ただし記録が1件も増えていなければ、3日経っても作り直さない
 */
function getInsight() {
  try {
    const days = getChartData().days || [];
    if (days.length < 3) return '';   // データが少ないうちは出さない

    // 基本情報を書き換えたときも作り直せるよう、その長さも判定に含める
    const sig = getOrCreateSheet_().getLastRow() + ':' + days[days.length - 1].date +
                ':p' + getProfileText_().length;
    const props = PropertiesService.getScriptProperties();
    const cached = JSON.parse(props.getProperty(INSIGHT_CACHE_KEY) || 'null');

    if (cached) {
      if (cached.sig === sig) return cached.text;              // 記録に変化なし
      if (daysSince_(cached.date) < INSIGHT_INTERVAL_DAYS) {   // 変化はあるが、まだ間隔内
        return cached.text;
      }
    }

    const text = askGemini_(buildInsightPrompt_(days));
    if (text) {
      props.setProperty(INSIGHT_CACHE_KEY, JSON.stringify({
        sig: sig, text: text, date: todayStr_()
      }));
      return text;
    }
    // 生成に失敗（上限・通信エラー）→ 前回の文 → =AI()のセル の順で拾う
    return cached ? cached.text : readSheetInsight_();
  } catch (e) {
    return readSheetInsight_();
  }
}

/**
 * セルの値を Date にする。読めなければ null。
 * 手入力した行は Date、アプリが書いた行は文字列、と型が混在するため必ずここを通す
 */
function toDate_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  if (!v) return null;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

/** 時刻セル → 'HH:mm' の文字列。Dateのまま返すと画面へ渡せない */
function fmtTime_(v) {
  if (v instanceof Date) {
    return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, 'Asia/Tokyo', 'HH:mm');
  }
  return v ? String(v) : '';
}

/** 今日の日付（yyyy/MM/dd・日本時間） */
function todayStr_() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
}

/** yyyy/MM/dd から今日までの経過日数。日付が無ければ十分大きい値を返す */
function daysSince_(dateStr) {
  if (!dateStr) return 9999;
  const then = new Date(dateStr + ' 00:00:00');
  if (isNaN(then.getTime())) return 9999;
  const now = new Date(todayStr_() + ' 00:00:00');
  return Math.floor((now - then) / 86400000);
}

// ===== AIに毎回渡す前提情報 =====
// Gemini は呼び出しごとに記憶がリセットされる。前回の内容も本人のことも覚えていないので、
// 知っていてほしいことは毎回プロンプトに入れる必要がある。

/** 基本情報シートを取得（なければ記入欄つきで作成） */
function getProfileSheet_() {
  const ss = SpreadsheetApp.openById(prop_('SHEET_ID'));
  let sh = ss.getSheetByName(PROFILE_SHEET);
  if (!sh) {
    sh = ss.insertSheet(PROFILE_SHEET);
    sh.appendRow(['項目', '内容']);
    sh.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#f1f3f4');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 170);
    sh.setColumnWidth(2, 430);
    // 内容が空の行はAIに渡さないので、書きたいものだけ埋めればよい
    [['年齢', ''], ['性別', ''], ['服薬', ''], ['通院', ''],
     ['目標', ''], ['気をつけていること', ''], ['その他', ''],
     // メモ欄から拾う言葉。ここだけは初期値を入れておく（AIには渡らない設定値）
     [EX_WORDS_ROW, DEFAULT_EX_WORDS],
     [COND_WORDS_ROW, DEFAULT_COND_WORDS]]
      .forEach(function (r) { sh.appendRow(r); });
  }
  return sh;
}

// メモ欄から拾う言葉。どちらも基本情報シートで差し替えられる。
// メモ本文はAIに送らず、これらに一致したかどうかだけを渡す
const EX_WORDS_ROW = '運動キーワード';
const DEFAULT_EX_WORDS = '筋トレ,バイク,ハイキング,散歩,ウォーキング,ジム,ランニング,水泳,ヨガ';

const COND_WORDS_ROW = '体調キーワード';
const DEFAULT_COND_WORDS = '風邪,発熱,寝不足,ストレス,頭痛,1回のみ';

/** 設定用の行。値ではなく設定なので、基本情報としてAIに渡さない */
const SETTING_ROWS = [EX_WORDS_ROW, COND_WORDS_ROW];

/** 基本情報を「項目: 内容」の行にする。未記入なら空文字 */
function getProfileText_() {
  try {
    const sh = getProfileSheet_();
    if (sh.getLastRow() < 2) return '';
    return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
      .filter(function (r) { return r[0] && r[1] !== '' && r[1] !== null; })
      .filter(function (r) { return SETTING_ROWS.indexOf(String(r[0]).trim()) < 0; })   // 設定値はAIに渡さない
      .map(function (r) { return String(r[0]).trim() + ': ' + String(r[1]).trim(); })
      .join('\n');
  } catch (e) {
    return '';
  }
}

/**
 * 基本情報シートの設定行から、キーワードの一覧を読む。
 * 未記入なら既定値。実行のたびにシートを読まないようキャッシュする
 */
var WORDS_CACHE = {};
function wordsFor_(rowName, fallback) {
  if (WORDS_CACHE[rowName]) return WORDS_CACHE[rowName];

  let custom = '';
  try {
    const sh = getProfileSheet_();
    if (sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
        if (String(r[0]).trim() === rowName && r[1]) custom = String(r[1]);
      });
    }
  } catch (e) {}

  WORDS_CACHE[rowName] = (custom || fallback)
    .split(/[,、\s]+/)
    .map(function (s) { return s.trim(); })
    .filter(function (s) { return s; });
  return WORDS_CACHE[rowName];
}

/** 過去の月次まとめを古い順に n か月分。長期の流れを踏まえさせるために渡す */
function pastMonthlyText_(n) {
  try {
    const rows = readMonthlyRows_();
    if (!rows.length) return '';
    return rows.slice(-(n || 3)).map(function (m) {
      return '【' + m.label + '】朝平均 ' + (m.morning || '–') + ' / 夜平均 ' + (m.night || '–') + '\n' + m.text;
    }).join('\n\n');
  } catch (e) {
    return '';
  }
}

/** どのプロンプトの冒頭にも付ける共通の前提 */
function contextBlock_() {
  const parts = [];

  const profile = getProfileText_();
  if (profile) parts.push('【この人の基本情報】\n' + profile);

  const past = pastMonthlyText_(3);
  if (past) parts.push('【これまでの月ごとの振り返り】\n' + past);

  return parts.length ? parts.join('\n\n') + '\n\n' : '';
}

/** 基本情報を扱ううえでの共通ルール。薬や受診の判断に踏み込ませない */
const SAFETY_RULES = [
  '- 診断はしない。薬の増減・中止・受診の要否には一切触れない',
  '- 基本情報は「読み手の状況を踏まえる」ためだけに使い、医学的な助言はしない',
  '- 運動・睡眠と血圧の関係は、数日程度の一致では断定しない。',
  '  日数が足りないときは「3週間ほど」のように丸めた目安を言う。',
  '  計算して出した数ではないので、「20日分」のような細かい数字は使わない'
];

/**
 * 書き方の共通ルール。
 * 指示しないと「詳細が分からないので何とも言えませんが」のような前置きや、
 * 「医師にご相談ください」といった定型の免責文で字数を潰してしまう
 */
const STYLE_RULES = [
  '- 結論から書く。前置き・断り書き・言い訳は一切不要',
  '- 「続けていきましょう」のような激励や呼びかけで締めない',
  '- 「詳細が分からないので何とも言えませんが」「データからは判断できかねますが」のような',
  '  ためらいの文、および自分の限界についての説明は書かない',
  '- 「医師にご相談ください」などの定型文も書かない（画面に注記を別途出している）'
];

/** データの読み方。書式を説明しておかないと数値の意味を取り違える */
const DATA_LEGEND =
  '単位はmmHg。「7/1(水) 朝 07:20 128/86(脈72)」は7月1日水曜の7時20分に測って' +
  '収縮期128・拡張期86・脈拍72だった、という意味。' +
  '曜日は記載どおりで、自分で計算し直さないこと。' +
  '朝は起床後、夜は就寝前の測定。血圧は時間帯で変動するので、日によって測定時刻が大きく違うときは、' +
  '数値の違いをその差で説明できないか先に検討すること。' +
  '深睡眠はその夜の深い睡眠時間（分）。' +
  '「運動: 筋トレ」はその日に行った運動。記載がない日は、していないか記録し忘れのどちらか。' +
  '「体調: 風邪」はその日の特記事項。数値が普段と違う理由になりうるので、' +
  '傾向を語るときはその日を分けて扱い、平常時の傾向と混ぜないこと。';

// 家庭血圧の目安（mmHg）。この値を超えた日数を数えるのに使う
const HOME_SYS = 135;
const HOME_DIA = 85;

/**
 * 集計はコードで行い、結果だけをAIに渡す。
 * 90日分の記録を丸ごと渡して数えさせると、数え間違いが本文に混ざっても
 * 見た目が正しい数字と変わらないため、誰も気づけない。
 * ここで出した数字を「そのまま使う」と指示して、本文の数値を検証できるようにする。
 *
 * 渡す days は、記録一覧に載せるのと同じ配列であること。
 * 別の範囲で集計すると、集計と一覧が食い違う
 */
function buildStatsBlock_(days) {
  if (!days || !days.length) return '';

  const avg = function (a) {
    return a.length ? Math.round(a.reduce(function (x, y) { return x + y; }, 0) / a.length) : null;
  };

  /** 朝または夜の集計。数値の無い日は除いて数える */
  const slotLine = function (slot, name) {
    const rec = days.filter(function (d) { return d[slot].sys && d[slot].dia; });
    if (!rec.length) return null;

    const sys = rec.map(function (d) { return d[slot].sys; });
    const dia = rec.map(function (d) { return d[slot].dia; });
    const hi = rec.reduce(function (a, b) { return b[slot].sys > a[slot].sys ? b : a; });
    const lo = rec.reduce(function (a, b) { return b[slot].sys < a[slot].sys ? b : a; });

    return name + ': ' + rec.length + '日分' +
      ' / 平均 ' + avg(sys) + '/' + avg(dia) +
      ' / 収縮期' + HOME_SYS + '超え ' + sys.filter(function (v) { return v > HOME_SYS; }).length + '日' +
      ' / 拡張期' + HOME_DIA + '超え ' + dia.filter(function (v) { return v > HOME_DIA; }).length + '日' +
      ' / 最高 ' + hi.label + ' ' + hi[slot].sys + '/' + hi[slot].dia +
      ' / 最低 ' + lo.label + ' ' + lo[slot].sys + '/' + lo[slot].dia;
  };

  const lines = [
    '期間: ' + days[0].label + '〜' + days[days.length - 1].label + '（記録のある日 ' + days.length + '日）',
    slotLine('morning', '朝'),
    slotLine('night', '夜')
  ];

  const deep = days.filter(function (d) { return d.deep != null; }).map(function (d) { return d.deep; });
  if (deep.length) {
    lines.push('深睡眠: ' + deep.length + '日分 / 平均 ' + avg(deep) + '分' +
               ' / 最短 ' + Math.min.apply(null, deep) + '分 / 最長 ' + Math.max.apply(null, deep) + '分');
  }

  const ex = days.filter(function (d) { return d.exercise && d.exercise.length; });
  if (ex.length) {
    lines.push('運動の記録があった日: ' + ex.length + '日（' +
               ex.map(function (d) { return d.label; }).join('・') + '）');
  }

  const cond = days.filter(function (d) { return d.condition && d.condition.length; });
  if (cond.length) {
    lines.push('体調の記録があった日: ' + cond.map(function (d) {
      return d.label + '(' + d.condition.join('・') + ')';
    }).join('・'));
  }

  return [
    '【集計済みの数値】ここの数字をそのまま使うこと。記録を自分で数え直さないこと。',
    '同じ時間帯に複数回測った日は平均して1件にしてあるので、単位は「回」ではなく「日」。'
  ].concat(lines.filter(function (l) { return l; })).join('\n');
}

/**
 * この記録で増やせる項目の一覧。
 * 渡さないと「食事やストレスの記録があれば」のように、
 * このアプリでは残しようのないものを勧めてしまう。
 * キーワードは基本情報シートで差し替えられるので、そのときの中身を読んで渡す
 */
function recordableBlock_() {
  return [
    '【この記録で増やせる項目】材料が足りないと書くときは、この中から挙げること。ここに無いものは勧めない',
    '- 朝と夜の血圧・脈拍（測定時刻をそろえると、時間帯による差と混ざらなくなる）',
    '- 深睡眠の時間（分）',
    '- 運動: メモに書くと拾える言葉 … ' + wordsFor_(EX_WORDS_ROW, DEFAULT_EX_WORDS).join('・'),
    '- 体調: メモに書くと拾える言葉 … ' + wordsFor_(COND_WORDS_ROW, DEFAULT_COND_WORDS).join('・')
  ].join('\n');
}

/** 送るのは日付と数値だけ。メモ欄の自由記述は含めない */
function buildInsightPrompt_(days) {
  const lines = days.slice(-30).map(function (d) {
    const parts = [d.label + '(' + d.wd + ')'];
    if (d.morning.sys) parts.push('朝' + (d.morning.time ? ' ' + d.morning.time : '') + ' ' + d.morning.sys + '/' + d.morning.dia + (d.morning.pul ? '(脈' + d.morning.pul + ')' : ''));
    if (d.night.sys)   parts.push('夜' + (d.night.time ? ' ' + d.night.time : '') + ' ' + d.night.sys + '/' + d.night.dia + (d.night.pul ? '(脈' + d.night.pul + ')' : ''));
    if (d.deep != null) parts.push('深睡眠 ' + d.deep + '分');
    if (d.exercise && d.exercise.length) parts.push('運動: ' + d.exercise.join('・'));
    if (d.condition && d.condition.length) parts.push('体調: ' + d.condition.join('・'));
    return parts.join(' / ');
  });

  return [
    'あなたは家庭血圧の記録を読み解いて、本人に分かりやすく伝える役です。',
    '',
    contextBlock_() +
    '【直近の記録】家庭で測った血圧です。' + DATA_LEGEND,
    '',
    lines.join('\n'),
    '',
    '直近の記録から読み取れる傾向を、日本語で2〜3文にまとめてください。',
    '条件:',
    '- 全体で150文字以内。必ず文を言い切って終える',
    '- 月ごとの振り返りがある場合は、それと比べて良くなったか・変わらないかに触れる',
    '- 家庭血圧の目安（収縮期135 / 拡張期85）と比べてどうかに触れる',
    '- 朝と夜の差、睡眠時間との関係で目立つ点があれば書く',
    '- 「〜の傾向が見えます」程度の書き方にとどめる',
    '- 最後の一文まで、記録から読み取れた事実で埋める'
  ].concat(STYLE_RULES).concat(SAFETY_RULES).concat([
    '- 見出しは不要。本文だけを返す'
  ]).join('\n');
}

/** Gemini API を呼ぶ。失敗時は空文字（呼び出し側で握る） */
function askGemini_(prompt) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
  if (!key) return '';

  const res = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent',
    {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': key },
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        // 思考トークンも maxOutputTokens に含まれる。
        // 3.6-flash は既定が medium で思考が長いため、本文が切れないよう枠を広く取る
        generationConfig: { temperature: 0.4, maxOutputTokens: 4000 }
      }),
      muteHttpExceptions: true
    }
  );

  if (res.getResponseCode() !== 200) {
    Logger.log('Gemini APIエラー ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 400));
    return '';   // 429（無料枠の上限）もここで落とす
  }

  const json = JSON.parse(res.getContentText());
  const cand = json.candidates && json.candidates[0];

  // 枠切れで尻切れになった文は採用しない（呼び出し側が前回の文にフォールバックする）
  if (cand && cand.finishReason === 'MAX_TOKENS') {
    Logger.log('出力枠オーバー（MAX_TOKENS）。maxOutputTokens を増やしてください');
    return '';
  }

  const part = cand && cand.content && cand.content.parts && cand.content.parts[0];
  const text = (part && part.text) ? part.text.trim() : '';
  if (!text) {
    Logger.log('本文が空です。finishReason=' + (cand && cand.finishReason) +
               ' / 応答: ' + res.getContentText().slice(0, 400));
  }
  return text;
}

/**
 * 動作確認用：エディタでこの関数を実行し、下に出る「実行ログ」を見る
 * getInsight() はエラーを握りつぶす作りなので、切り分けはこちらで行う
 */
function testGemini() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
  if (!key) {
    Logger.log('NG: スクリプトプロパティに GEMINI_KEY がありません（名前のスペルを確認）');
    return;
  }
  Logger.log('キー: 登録あり（末尾 …' + key.slice(-4) + '）');

  const res = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent',
    {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': key },
      payload: JSON.stringify({ contents: [{ parts: [{ text: '「テスト成功」とだけ返して' }] }] }),
      muteHttpExceptions: true
    }
  );

  Logger.log('HTTPステータス: ' + res.getResponseCode() + '（200なら成功）');
  Logger.log(res.getContentText().slice(0, 600));
}

/**
 * 「最近の記録」が読み込めないときの切り分け用。
 * エディタで実行して、ログにデータが出ればサーバー側は正常（＝画面側の問題）、
 * ここでエラーになればシートの中身が原因
 */
function testInitialData() {
  const sheet = getOrCreateSheet_();
  Logger.log('最終行: ' + sheet.getLastRow());

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const n = Math.min(5, lastRow - 1);
    sheet.getRange(lastRow - n + 1, 1, n, 6).getValues().forEach(function (r, i) {
      // 日付セルの型を確認する（手入力した行が文字列になっていると落ちる）
      Logger.log('行' + (i + 1) + ' 日付=[' + r[0] + '] 型=' + (r[0] instanceof Date ? 'Date' : typeof r[0]) +
                 ' / 時刻=[' + r[1] + '] / 時間帯=[' + r[2] + ']');
    });
  }

  Logger.log('--- getInitialData() の結果 ---');
  Logger.log(JSON.stringify(getInitialData()));
}

/**
 * 保存済みのコメントを捨てて、今すぐ作り直す。
 * プロンプトやモデルを変えた直後は、記録が増えるまで古い文が出続けるのでこれを実行する
 */
function refreshInsight() {
  PropertiesService.getScriptProperties().deleteProperty(INSIGHT_CACHE_KEY);
  Logger.log(getInsight() || '(生成できませんでした)');
}

// ===== 月次まとめ =====
// 終わった月ごとに1本ずつコメントを作り、シートに貯めていく。
// 一度作ったら作り直さないので、過去の月はいつでも同じ文を読み返せる。

/** 'yyyy/MM/dd' → 'yyyy-MM' */
function monthKeyOf_(dateStr) { return dateStr.slice(0, 7).replace('/', '-'); }

/** 'yyyy-MM' → '2026年7月' */
function monthLabelOf_(key) {
  const p = key.split('-');
  return p[0] + '年' + Number(p[1]) + '月';
}

/** 'yyyy-MM' → '7月'（タブ用の短い表記） */
function monthShortOf_(key) {
  return Number(key.split('-')[1]) + '月';
}

/** A列の値を 'yyyy-MM' に戻す。シートが日付に変換していても拾えるようにする */
function monthKeyFromCell_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM');
  return String(v || '');
}

/** 月次まとめシートを取得（なければ作成） */
function getMonthlySheet_() {
  const ss = SpreadsheetApp.openById(prop_('SHEET_ID'));
  let sh = ss.getSheetByName(MONTHLY_SHEET);
  if (!sh) {
    sh = ss.insertSheet(MONTHLY_SHEET);
    sh.appendRow(['月', '表示名', '日数', '朝平均', '夜平均', 'まとめ', '生成日時']);
    sh.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#f1f3f4');
    sh.setFrozenRows(1);
    sh.getRange('A:B').setNumberFormat('@');   // '2026-07' '2026年7月' を日付に変換させない
    sh.setColumnWidth(6, 460);
  }
  return sh;
}

/** その月の平均値をまとめる */
function monthStats_(list) {
  function avgOf(pick) {
    const v = list.map(pick).filter(function (x) { return x; });
    if (!v.length) return null;
    return Math.round(v.reduce(function (a, b) { return a + b; }, 0) / v.length);
  }
  return {
    days: list.length,
    mSys: avgOf(function (d) { return d.morning.sys; }),
    mDia: avgOf(function (d) { return d.morning.dia; }),
    nSys: avgOf(function (d) { return d.night.sys; }),
    nDia: avgOf(function (d) { return d.night.dia; }),
    deep: avgOf(function (d) { return d.deep; })
  };
}

function buildMonthlyPrompt_(label, list, st) {
  const lines = list.map(function (d) {
    const parts = [d.label + '(' + d.wd + ')'];
    if (d.morning.sys) parts.push('朝' + (d.morning.time ? ' ' + d.morning.time : '') + ' ' + d.morning.sys + '/' + d.morning.dia);
    if (d.night.sys)   parts.push('夜' + (d.night.time ? ' ' + d.night.time : '') + ' ' + d.night.sys + '/' + d.night.dia);
    if (d.deep != null) parts.push('深睡眠 ' + d.deep + '分');
    if (d.exercise && d.exercise.length) parts.push('運動: ' + d.exercise.join('・'));
    if (d.condition && d.condition.length) parts.push('体調: ' + d.condition.join('・'));
    return parts.join(' / ');
  });

  return [
    'あなたは家庭血圧の記録を読み解いて、本人に分かりやすく伝える役です。',
    '',
    contextBlock_() +
    '【' + label + 'の記録】家庭で測った血圧です。' + DATA_LEGEND,
    '',
    lines.join('\n'),
    '',
    '月間平均: 朝 ' + st.mSys + '/' + st.mDia + '、夜 ' + st.nSys + '/' + st.nDia +
      (st.deep != null ? '、深睡眠 ' + st.deep + '分' : ''),
    '',
    label + 'を振り返るまとめを、日本語で3〜4文で書いてください。',
    '条件:',
    '- 全体で200文字以内。必ず文を言い切って終える',
    '- 前の月の振り返りがある場合は、そこからの変化に必ず触れる',
    '- 月の前半と後半で変化があれば触れる',
    '- 家庭血圧の目安（収縮期135 / 拡張期85）と比べてどうだったかに触れる',
    '- 「〜の傾向でした」程度の書き方にとどめる',
    '- 最後の一文まで、記録から読み取れた事実で埋める'
  ].concat(STYLE_RULES).concat(SAFETY_RULES).concat([
    '- 見出しは不要。本文だけを返す'
  ]).join('\n');
}

/**
 * まだ作っていない「終わった月」のまとめを生成してシートに追記する。
 * 今月はまだ途中なので対象外。毎月1日のトリガーからも呼ばれる
 */
function buildMonthlySummaries(maxCount) {
  // トリガー経由だと引数にイベントオブジェクトが渡るので、数値のときだけ採用する
  const limit = (typeof maxCount === 'number') ? maxCount : 3;

  const days = getChartData().days || [];
  if (!days.length) return 0;

  const thisMonth = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
  const sh = getMonthlySheet_();

  // A列が日付型に変換されていても 'yyyy-MM' に戻してから照合する。
  // ここを String() で比較すると毎回「未作成」と判定され、開くたびに行が増える
  const done = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
      .forEach(function (r) { done[monthKeyFromCell_(r[0])] = true; });
  }

  const byMonth = {};
  const order = [];
  days.forEach(function (d) {
    const k = monthKeyOf_(d.date);
    if (!byMonth[k]) { byMonth[k] = []; order.push(k); }
    byMonth[k].push(d);
  });

  let made = 0;
  for (let i = 0; i < order.length && made < limit; i++) {
    const k = order[i];
    if (k === thisMonth || done[k]) continue;   // 今月・作成済みは飛ばす
    if (byMonth[k].length < 3) continue;        // 記録が3日未満の月は作らない

    const st = monthStats_(byMonth[k]);
    const text = askGemini_(buildMonthlyPrompt_(monthLabelOf_(k), byMonth[k], st));
    if (!text) continue;

    sh.appendRow([
      k,
      monthLabelOf_(k),
      st.days,
      st.mSys ? st.mSys + '/' + st.mDia : '',
      st.nSys ? st.nSys + '/' + st.nDia : '',
      text,
      Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')
    ]);
    // 書いた直後に文字列として固定する（日付に化けると次回の照合が外れる）
    sh.getRange(sh.getLastRow(), 1, 1, 2)
      .setNumberFormat('@')
      .setValues([[k, monthLabelOf_(k)]]);
    made++;
  }
  return made;
}

/**
 * 月次まとめシートを読むだけ（生成はしない）。古い月が先頭。
 * プロンプトへ渡すときにも使うので、生成処理とは必ず分けておく
 */
function readMonthlyRows_() {
  const sh = getMonthlySheet_();
  if (sh.getLastRow() < 2) return [];

  // 表示名はシートのB列を使わずコードで組み立てる。
  // '2026年7月' はシートが日付として解釈してしまい、Date型で返ってくることがあるため
  return sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues()
    .map(function (r) {
      const key = monthKeyFromCell_(r[0]);
      return {
        key: key,
        label: monthLabelOf_(key),
        short: monthShortOf_(key),
        days: r[2] ? Number(r[2]) : 0,
        morning: r[3] ? String(r[3]) : '',
        night: r[4] ? String(r[4]) : '',
        text: r[5] ? String(r[5]) : ''
      };
    })
    .filter(function (m) { return m.text; });
}

/** レポートページの「月次まとめ」カード用。新しい月が先頭 */
function getMonthlySummaries() {
  try {
    buildMonthlySummaries(2);   // 未生成の月があれば作る（表示が重くならないよう2件まで）
    return readMonthlyRows_().reverse();
  } catch (e) {
    return [];
  }
}

/**
 * 月次まとめの重複行を掃除する。同じ月は「最初の1行」だけ残す。
 * 重複が出たあとに一度実行すればよい
 */
function dedupeMonthlySummaries() {
  const sh = getMonthlySheet_();
  if (sh.getLastRow() < 3) { Logger.log('重複はありません'); return; }

  const seen = {};
  const toDelete = [];
  sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r, i) {
    const k = monthKeyFromCell_(r[0]);
    if (seen[k]) toDelete.push(i + 2);   // データは2行目から
    else seen[k] = true;
  });

  // 下から消さないと行番号がずれる
  toDelete.reverse().forEach(function (row) { sh.deleteRow(row); });
  Logger.log(toDelete.length + '行を削除しました（残り ' + (sh.getLastRow() - 1) + 'か月分）');
}

/**
 * 直前のやりとり1往復ぶん。
 * 「じゃあ何日あればいい？」のような、前の回答を受けた質問を読めるようにする。
 * 画面側が持っているものを受け取るだけで、サーバーには残さない
 */
function prevBlock_(prev) {
  if (!prev) return '';
  const pq = String(prev.q || '').trim().slice(0, 200);
  const pa = String(prev.a || '').trim().slice(0, 400);
  if (!pq || !pa) return '';

  return [
    '【直前のやりとり】今の質問がこれを受けたものかを読み取るためだけに使う。',
    'ここにある数値は前回の出力であって記録ではない。',
    '数値は必ず【集計済みの数値】から取り直すこと。前回の回答をなぞらない',
    'Q: ' + pq,
    'A: ' + pa
  ].join('\n');
}

/**
 * 質問コーナー：自分の記録について自由に質問する。
 * 送るのは日付と数値だけ。メモ欄の文章は含めない。
 * prev は直前のやりとり {q, a}（画面が持っている。無ければ省略可）
 */
function askAboutData(question, prev) {
  const q = String(question || '').trim();
  if (!q) return '';
  if (q.length > 200) return '質問が長すぎます。200文字以内でお願いします。';

  try {
    const days = getChartData().days || [];
    if (days.length < 3) return 'まだ記録が少ないため答えられません。3日分以上たまってから試してください。';

    // 一覧と集計は必ず同じ範囲から作る（食い違うと本文の数字が信用できなくなる）
    const recent = days.slice(-90);
    const lines = recent.map(function (d) {
      const parts = [d.label + '(' + d.wd + ')'];
      if (d.morning.sys) parts.push('朝' + (d.morning.time ? ' ' + d.morning.time : '') + ' ' + d.morning.sys + '/' + d.morning.dia + (d.morning.pul ? '(脈' + d.morning.pul + ')' : ''));
      if (d.night.sys)   parts.push('夜' + (d.night.time ? ' ' + d.night.time : '') + ' ' + d.night.sys + '/' + d.night.dia + (d.night.pul ? '(脈' + d.night.pul + ')' : ''));
      if (d.deep != null) parts.push('深睡眠 ' + d.deep + '分');
      if (d.exercise && d.exercise.length) parts.push('運動: ' + d.exercise.join('・'));
      if (d.condition && d.condition.length) parts.push('体調: ' + d.condition.join('・'));
      return parts.join(' / ');
    });

    const text = askGemini_([
      'あなたは家庭血圧の記録を読み解いて、本人に分かりやすく伝える役です。',
      '',
      contextBlock_() +
      '【記録】家庭で測った血圧です。' + DATA_LEGEND,
      '',
      lines.join('\n'),
      '',
      buildStatsBlock_(recent),
      '',
      recordableBlock_(),
      '',
      prevBlock_(prev),
      '',
      '【質問】' + q,
      '',
      '答え方:',
      '- 最初の一文で、質問に直接答える。統計や背景を書くのはそのあと',
      '- 根拠として、上の記録から日付と数値を挙げる',
      '- 日数・平均・最高最低は【集計済みの数値】から引く。自分で数えない',
      '- 記録の数値では答えようのない質問なら、最初の一文でそう言い切り、',
      '  そのあとに次の2つを書いて終える。質問と関係のない統計を並べて埋めない',
      '  (1) どのくらいの期間があれば見えてくるか。「3週間ほど」のように丸めた目安でよい',
      '  (2)【この記録で増やせる項目】のうち、何を足すと判断しやすくなるか'
    ].concat(STYLE_RULES).concat(SAFETY_RULES).concat([
      '- 350文字以内。見出しは不要'
    ]).join('\n'));

    return text || '答えを作れませんでした。もう一度試してください。';
  } catch (e) {
    return 'エラーが起きました: ' + e.message;
  }
}

/**
 * 毎月1日に前月分のまとめを自動生成するトリガーを仕掛ける。
 * 一度だけ実行すればよい（ビューアを開いたときにも補完されるので必須ではない）
 */
function installMonthlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'buildMonthlySummaries') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('buildMonthlySummaries').timeBased().onMonthDay(1).atHour(6).create();
  Logger.log('毎月1日の6時台に、前月分のまとめを自動生成します');
}

/**
 * 「このモデルは提供終了です」(404) が出たときに実行する。
 * 今このキーで使えるモデル名が実行ログに並ぶので、GEMINI_MODEL を書き換える
 */
function listModels() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
  const res = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models',
    { headers: { 'x-goog-api-key': key }, muteHttpExceptions: true }
  );
  const models = (JSON.parse(res.getContentText()).models) || [];
  models
    .filter(function (m) { return (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0; })
    .forEach(function (m) { Logger.log(m.name.replace('models/', '')); });
}

/** 予備：=AI() が書き込んだセルを読む */
function readSheetInsight_() {
  try {
    const sh = SpreadsheetApp.openById(prop_('SHEET_ID')).getSheetByName(INSIGHT_SHEET);
    if (!sh) return '';
    const v = sh.getRange(INSIGHT_CELL).getValue();
    return (typeof v === 'string' && v.trim()) ? v.trim() : '';
  } catch (e) {
    return '';
  }
}

/**
 * 記録を保存する
 * @param {Object} data - {systolic, diastolic, pulse, deepSleep, memo}
 */
function saveRecord(data) {
  const sheet = getOrCreateSheet_();
  const now = new Date();

  const sys = Number(data.systolic);
  const dia = Number(data.diastolic);
  const pul = data.pulse ? Number(data.pulse) : '';
  const deep = data.deepSleep ? Number(data.deepSleep) : '';   // 深睡眠（分・朝1回）

  // 入力ミス防止の簡易チェック
  if (!sys || !dia) throw new Error('収縮期と拡張期を入力してください');
  if (sys < 50 || sys > 300) throw new Error('収縮期の値を確認してください');
  if (dia < 30 || dia > 200) throw new Error('拡張期の値を確認してください');
  if (deep !== '' && (deep < 0 || deep > 1000)) throw new Error('深睡眠(分)の値を確認してください');

  sheet.appendRow([
    Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd'),
    Utilities.formatDate(now, 'Asia/Tokyo', 'HH:mm'),
    getTimeSlot_(now),
    sys,
    dia,
    pul,
    data.memo || '',
    deep
  ]);

  return {
    ok: true,
    savedAt: Utilities.formatDate(now, 'Asia/Tokyo', 'M月d日 HH:mm')
  };
}

/** 時間帯を自動判定（朝・昼・夜） */
function getTimeSlot_(date) {
  const h = Number(Utilities.formatDate(date, 'Asia/Tokyo', 'H'));
  if (h < 11) return '朝';
  if (h < 17) return '昼';
  return '夜';
}

/** シートを取得（なければ作成） */
function getOrCreateSheet_() {
  const ss = SpreadsheetApp.openById(prop_('SHEET_ID'));
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['日付', '時刻', '時間帯', '収縮期', '拡張期', '脈拍', 'メモ', '深睡眠']);
    sheet.getRange(1, 1, 1, 8)
      .setFontWeight('bold')
      .setBackground('#f1f3f4');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 95);
    sheet.setColumnWidth(7, 220);
  } else if (sheet.getRange(1, 8).getValue() !== '深睡眠') {
    // 既存シートに深睡眠列（H列）を後付け
    sheet.getRange(1, 8).setValue('深睡眠').setFontWeight('bold').setBackground('#f1f3f4');
  }
  return sheet;
}

/**
 * ビューア用データ：日ごとに「朝の平均」「夜の平均」を算出して返す
 * （朝2回・夜2回測った値をその日の平均にまとめる）
 */
function getChartData() {
  const sheet = getOrCreateSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { days: [] };

  const values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const map = {};
  const order = [];

  values.forEach(function (r) {
    const dt = toDate_(r[0]);
    if (!dt) return;   // 日付として読めない行は飛ばす
    const date = Utilities.formatDate(dt, 'Asia/Tokyo', 'yyyy/MM/dd');
    const slot = r[2] ? String(r[2]).trim() : '';

    if (!map[date]) {
      // t は測定時刻。同じ「朝」でも測る時間が2時間ずれれば数値は変わるので、
      // 集約時に捨てずに持っておく
      map[date] = {
        '朝': { sys: [], dia: [], pul: [], t: [] },
        '夜': { sys: [], dia: [], pul: [], t: [] },
        deep: null, ex: {}, cond: {}
      };
      order.push(date);
    }

    // 深睡眠（分・その日1回）: どの行に入っていても拾う
    const deep = r[7] !== '' && r[7] !== null ? Number(r[7]) : null;
    if (deep !== null && !isNaN(deep)) map[date].deep = deep;

    // メモ欄からはキーワードの一致だけを拾う。メモ本文そのものはどこにも渡さない
    const memo = r[6] ? String(r[6]) : '';
    if (memo) {
      wordsFor_(EX_WORDS_ROW, DEFAULT_EX_WORDS).forEach(function (w) {
        if (memo.indexOf(w) >= 0) map[date].ex[w] = true;
      });
      wordsFor_(COND_WORDS_ROW, DEFAULT_COND_WORDS).forEach(function (w) {
        if (memo.indexOf(w) >= 0) map[date].cond[w] = true;
      });
    }

    if (slot !== '朝' && slot !== '夜') return;
    const sys = Number(r[3]);
    const dia = Number(r[4]);
    const pul = r[5] ? Number(r[5]) : null;
    const t = fmtTime_(r[1]);
    if (sys) map[date][slot].sys.push(sys);
    if (dia) map[date][slot].dia.push(dia);
    if (pul) map[date][slot].pul.push(pul);
    if (sys && t) map[date][slot].t.push(t);
  });

  function avg(a) {
    if (!a || !a.length) return null;
    const sum = a.reduce(function (x, y) { return x + y; }, 0);
    return Math.round(sum / a.length);
  }

  /** その時間帯で最初に測った時刻。'HH:mm' なので文字列のまま並べ替えられる */
  function firstTime(a) {
    return (a && a.length) ? a.slice().sort()[0] : '';
  }

  // 曜日はAIに渡すために持たせる。渡さないとカレンダーを推測させることになり、
  // 「週末は高い」のような検証できない話を作られてしまう
  const WD = ['日', '月', '火', '水', '木', '金', '土'];

  const days = order.map(function (d) {
    return {
      date: d,
      label: Utilities.formatDate(toDate_(d), 'Asia/Tokyo', 'M/d'),
      wd: WD[new Date(d + ' 00:00:00').getDay()],
      morning: { sys: avg(map[d]['朝'].sys), dia: avg(map[d]['朝'].dia), pul: avg(map[d]['朝'].pul), time: firstTime(map[d]['朝'].t) },
      night:   { sys: avg(map[d]['夜'].sys), dia: avg(map[d]['夜'].dia), pul: avg(map[d]['夜'].pul), time: firstTime(map[d]['夜'].t) },
      deep:    map[d].deep,
      exercise: Object.keys(map[d].ex),
      condition: Object.keys(map[d].cond)
    };
  });

  return { days: days };
}

// ===== 連続記録日数 =====
// 節目までの残りが見えていると、その日の行動が変わる。数字そのものより「あと何日」が効く
const STREAK_GOALS = [3, 7, 14, 30, 60, 100, 180, 365];

// 登山や旅行では血圧計を持って行けない。これ以下の空白は途切れとみなさない。
// 「旅行に行くと記録が終わる」仕様だと、続ける仕掛けとして本末転倒になる
const STREAK_MAX_GAP = 3;

/**
 * 連続記録日数と、次の節目までの残り。
 * ・今日まだ記録していなくても途切れ扱いにしない（朝いちで0日と出して意欲を削がない）
 * ・3日までの空白は飛び越えて数える（数えるのは実際に記録した日数）
 */
function streakInfo_() {
  const empty = { days: 0, next: STREAK_GOALS[0], remain: STREAK_GOALS[0], todayDone: false };

  const sheet = getOrCreateSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return empty;

  // 記録のある日を集める（1日に何回測っても1日と数える）
  const seen = {};
  sheet.getRange(2, 1, lastRow - 1, 1).getValues().forEach(function (r) {
    const d = toDate_(r[0]);
    if (d) seen[Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd')] = true;
  });

  const ONE_DAY = 86400000;
  const today = todayStr_();
  const todayDone = !!seen[today];

  let cur = new Date(today + ' 00:00:00');
  if (!todayDone) cur = new Date(cur.getTime() - ONE_DAY);   // 今日の分はこれから測ればよい

  let days = 0;
  let gap = 0;
  for (let i = 0; i < 3000; i++) {
    if (seen[Utilities.formatDate(cur, 'Asia/Tokyo', 'yyyy/MM/dd')]) {
      days++;
      gap = 0;
    } else {
      gap++;
      if (gap > STREAK_MAX_GAP) break;   // 4日以上あいたらそこで途切れ
    }
    cur = new Date(cur.getTime() - ONE_DAY);
  }

  // 今日まだなら、今日測ったときの日数を基準に次の節目を出す
  const shown = todayDone ? days : days + 1;
  let next = null;
  for (let i = 0; i < STREAK_GOALS.length; i++) {
    if (STREAK_GOALS[i] >= shown) { next = STREAK_GOALS[i]; break; }
  }

  return {
    days: days,
    next: next,
    remain: next ? Math.max(0, next - days) : 0,
    todayDone: todayDone
  };
}

// ===== 目標の達成状況（テーマ解放の条件） =====
// 「直近30日の8割が基準値以下」。1〜2日の乱れでは壊れないので、
// 数値が悪そうな日に測るのをためらわずに済む。
// 連続記録と違って外れることもあるため、毎日「今日達成か」の緊張感が続く
const GOAL_SYS = 135;      // 家庭血圧の目安
const GOAL_DIA = 85;
const GOAL_WINDOW = 30;    // 記録のある直近30日で見る（休んだ日は数に入れない）
const GOAL_RATIO = 0.8;
const GOAL_MIN_DAYS = 20;  // これ未満は判定しない（少ない日数だと簡単に達成できてしまう）

function goalInfo_() {
  const win = (getChartData().days || []).slice(-GOAL_WINDOW);

  let ok = 0;
  let counted = 0;
  win.forEach(function (d) {
    const sys = [], dia = [];
    ['morning', 'night'].forEach(function (k) {
      if (d[k] && d[k].sys) { sys.push(d[k].sys); dia.push(d[k].dia); }
    });
    if (!sys.length) return;
    counted++;

    function avg(a) { return Math.round(a.reduce(function (x, y) { return x + y; }, 0) / a.length); }
    if (avg(sys) < GOAL_SYS && avg(dia) < GOAL_DIA) ok++;
  });

  const need = Math.ceil(counted * GOAL_RATIO);
  return {
    ok: ok,
    total: counted,
    need: need,
    remain: Math.max(0, need - ok),
    ready: counted >= GOAL_MIN_DAYS,
    achieved: counted >= GOAL_MIN_DAYS && ok >= need
  };
}

/**
 * 印刷用ページのデータ。診察に持っていく1枚を作るために使う。
 * 期間は暦月ではなく「今日から遡ってnか月」。診察日は月末月初に揃わないため
 */
function getPrintData(months) {
  const n = (Number(months) === 3) ? 3 : 1;

  const to = new Date(todayStr_() + ' 00:00:00');
  const from = new Date(to.getTime());
  const keepDay = from.getDate();
  from.setMonth(from.getMonth() - n);
  if (from.getDate() !== keepDay) from.setDate(0);   // 2/31 のような繰り上がりを月末に直す
  from.setDate(from.getDate() + 1);

  const fromStr = Utilities.formatDate(from, 'Asia/Tokyo', 'yyyy/MM/dd');
  const days = (getChartData().days || []).filter(function (d) { return d.date >= fromStr; });

  // 日ごとの平均で「目安を超えた日」を数える（1日に複数回測っていても1日と数える）
  const mS = [], mD = [], nS = [], nD = [];
  let over = 0;
  days.forEach(function (d) {
    if (d.morning.sys) { mS.push(d.morning.sys); mD.push(d.morning.dia); }
    if (d.night.sys)   { nS.push(d.night.sys);   nD.push(d.night.dia); }

    const s = [], t = [];
    if (d.morning.sys) { s.push(d.morning.sys); t.push(d.morning.dia); }
    if (d.night.sys)   { s.push(d.night.sys);   t.push(d.night.dia); }
    if (!s.length) return;
    if (mean_(s) >= GOAL_SYS || mean_(t) >= GOAL_DIA) over++;
  });

  return {
    label: fmtJp_(from) + ' 〜 ' + fmtJp_(to),
    months: n,
    days: days,
    stats: {
      recorded: days.length,
      over: over,
      mSys: mean_(mS), mDia: mean_(mD),
      nSys: mean_(nS), nDia: mean_(nD),
      goalSys: GOAL_SYS, goalDia: GOAL_DIA
    }
  };
}

/** 平均（空なら null） */
function mean_(a) {
  if (!a || !a.length) return null;
  return Math.round(a.reduce(function (x, y) { return x + y; }, 0) / a.length);
}

/** Date → '2026年7月21日' */
function fmtJp_(d) {
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy年M月d日');
}

/** 起動時に渡す情報（前回値と直近履歴） */
function getInitialData() {
  const sheet = getOrCreateSheet_();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return { last: null, history: [], streak: streakInfo_(), goal: goalInfo_() };
  }

  const n = Math.min(5, lastRow - 1);
  const values = sheet.getRange(lastRow - n + 1, 1, n, 6).getValues();

  // 画面へ渡すのは文字列と数値だけにする。
  // Dateのまま返すとシリアライズに失敗し、戻り値が丸ごと null になる
  const rows = values.reverse().map(function (r) {
    const d = toDate_(r[0]);
    return {
      date: d ? Utilities.formatDate(d, 'Asia/Tokyo', 'M/d') : '',
      time: fmtTime_(r[1]),
      slot: r[2] ? String(r[2]) : '',
      systolic: r[3] ? Number(r[3]) : '',
      diastolic: r[4] ? Number(r[4]) : '',
      pulse: r[5] ? Number(r[5]) : ''
    };
  });

  return { last: rows[0], history: rows, streak: streakInfo_(), goal: goalInfo_() };
}
