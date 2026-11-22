/**
 * 釣りログ ― セットアップ
 *
 * 空のGoogleスプレッドシートで setupAll を1回実行すると、
 * 釣行・計測・計測集計・水深・場所マスター・タックルマスタ・選択肢 の7タブが組み上がる。
 *
 * すでにあるタブのデータは消さない（見出し・数式・書式・ドロップダウンだけ整える）。
 * 計測集計はデータを持たないので、毎回作り直す。
 */

// ===== 列の配置（weather.gs もこれを見る） =====
const LOG_HEADERS = ['ID', '日付', '場所', 'ポイント', '潮色', '潮流向き', '潮流強さ',
  '風向', '風速', '波', '釣法', '開始時間', '終了時間', 'ベイト', '感想'];

// 計測：A〜I は手で入れる、J〜O は自動
const MEAS_INPUT = ['ID', '日付', '場所', 'ロッド', 'リール', 'ルアー', '投目', '飛距離m', '備考'];
const MEAS_AUTO  = ['硬さ', 'スプール径mm', '径帯', '重量g', '重量帯', '形状'];

const DEPTH_HEADERS = ['日付', '場所', '地点', '沖距離m', '水深m下限', '水深m上限', '測り方', '備考'];
const PLACE_HEADERS = ['場所', '座標', 'メモ'];

const WEATHER_COL = 16;   // 釣行・計測とも P列から気象8列

// 計測シートで集計に使う列
const M = { rod: 'D', reel: 'E', dist: 'H', hard: 'J', dia: 'L', band: 'N', shape: 'O' };

// ===== 初期値 =====
const DEFAULT_BANDS = [5, 10, 15, 20, 30, 40];   // 重量帯。下限は隣との中間を自動で出す
const DEFAULT_DIA   = [[0, '〜30mm'], [31, '31-32mm'], [33, '33-34mm'], [35, '35-36mm'], [37, '37-38mm'], [39, '39mm〜']];
const DEFAULT_HARD  = ['UL', 'L', 'ML', 'M', 'MH', 'H', 'XH'];
const DEFAULT_SHAPE = ['ミノー', 'シンペン', 'バイブ', 'メタルジグ', 'ジグヘッド', 'ワーム', 'クランク', 'トップ', 'エギ', 'オモリ'];
const SAMPLE_RODS   = [['（例）8.6ft L シーバス', 'L'], ['（例）6.10ft MH ベイト', 'MH'], ['（例）9.6ft M ショア', 'M']];
const SAMPLE_REELS  = [['（例）小径ベイト 32mm', 32], ['（例）標準ベイト 34mm', 34], ['（例）遠投ベイト 38mm', 38]];
const SAMPLE_LURES  = [['（例）ミノー 78mm 10g', 10, 'ミノー'], ['（例）シンペン 80mm 12g', 12, 'シンペン'],
  ['（例）バイブ 20g', 20, 'バイブ'], ['（例）メタルジグ 30g', 30, 'メタルジグ']];

const CHOICES = {
  '潮色':     ['超クリア', 'クリア', 'ささ濁り', '弱濁り', '濁り', '白濁り'],
  '潮流向き': ['右→左', '左→右', '停滞', '二枚潮'],
  '潮流強さ': ['早い', '普通', 'ゆっくり', '止まり'],
  '風向':     ['無風', '北', '北東', '東', '南東', '南', '南西', '西', '北西'],
  '風速':     ['無風', '微風', '弱い', '普通', '強い', '強風'],
  '波':       ['べた凪', '凪', 'さざ波', '小さい', '風波', 'うねり', '大うねり'],
  '釣法':     ['ルアー', 'エギ', 'フカセ', 'ジギング', '投げ', 'サビキ'],
  'ベイト':   ['なし', 'あり', 'イワシ', 'アジ', 'ボラ', '小魚（種類不明）'],
  '測り方':   ['ディーパー', 'カウント', '目測'],
};

// ===== 色 =====
const C_HEAD = '#2F5597', C_AUTO = '#EDEDED', C_PICK = '#FFF2CC';
const C_RED = '#F8696B', C_YEL = '#FFEB84', C_GRN = '#63BE7B';

// ---------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('釣りログ')
    .addItem('初期セットアップ', 'setupAll')
    .addItem('計測集計を作り直す', 'buildSummary')
    .addSeparator()
    .addItem('気象データを取得', 'fillWeather')
    .addItem('毎週自動で気象を取る', 'installWeeklyTrigger')
    .addToUi();
}

function setupAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  buildChoices_(ss);
  buildPlaces_(ss);
  buildTackle_(ss);
  buildLog_(ss);
  buildMeasure_(ss);
  buildDepth_(ss);
  buildSummary();
  arrangeSheets_(ss);
  ss.toast('セットアップが終わりました。まずタックルマスタと場所マスターを自分の道具・釣り場に書き換えてください。', '釣りログ', 20);
}

// ---------------------------------------------------------------
// 各シート

function buildChoices_(ss) {
  const { sh, created } = ensureSheet_(ss, '選択肢');
  const keys = Object.keys(CHOICES);
  if (created) {
    const rows = Math.max.apply(null, keys.map(k => CHOICES[k].length));
    const grid = [keys];
    for (let i = 0; i < rows; i++) grid.push(keys.map(k => CHOICES[k][i] || ''));
    sh.getRange(1, 1, grid.length, keys.length).setValues(grid);
  }
  styleHeader_(sh, keys.length);
  sh.setColumnWidths(1, keys.length, 110);
}

function buildPlaces_(ss) {
  const { sh, created } = ensureSheet_(ss, '場所マスター');
  if (created) sh.getRange(1, 1, 1, PLACE_HEADERS.length).setValues([PLACE_HEADERS]);
  styleHeader_(sh, PLACE_HEADERS.length);
  sh.setColumnWidth(1, 140).setColumnWidth(2, 280).setColumnWidth(3, 300);
  if (created) {
    sh.getRange(2, 1, 1, 3).setValues([['（例）〇〇漁港', '36.9000, 137.4000',
      'Googleマップで右クリック→座標をコピー、をそのまま貼る']]);
  }
}

function buildTackle_(ss) {
  const { sh, created } = ensureSheet_(ss, 'タックルマスタ');
  const heads = {
    1: 'ロッド', 2: '硬さ', 4: 'リール', 5: 'スプール径mm', 7: 'ルアー', 8: '重量g', 9: '形状',
    11: '重量帯の下限（自動）', 12: '重量帯', 14: '径帯の下限', 15: '径帯', 17: '硬さの並び', 19: '形状の並び',
  };
  Object.keys(heads).forEach(c => sh.getRange(1, Number(c)).setValue(heads[c]));
  if (created) {
    put_(sh, 2, 1, SAMPLE_RODS);
    put_(sh, 2, 4, SAMPLE_REELS);
    put_(sh, 2, 7, SAMPLE_LURES);
    put_(sh, 2, 12, DEFAULT_BANDS.map(v => [v]));
    put_(sh, 2, 14, DEFAULT_DIA);
    put_(sh, 2, 17, DEFAULT_HARD.map(v => [v]));
    put_(sh, 2, 19, DEFAULT_SHAPE.map(v => [v]));
  }
  sh.setFrozenRows(1);
  Object.keys(heads).forEach(c => styleHeaderCell_(sh.getRange(1, Number(c))));
  [1, 4, 7].forEach(c => sh.setColumnWidth(c, 220));
  [3, 6, 10, 13, 16, 18].forEach(c => sh.setColumnWidth(c, 24));
  // 硬さ・形状は「並び」から選ぶ
  setListFromRange_(sh.getRange('B2:B200'), sh.getRange('Q2:Q30'));
  setListFromRange_(sh.getRange('I2:I300'), sh.getRange('S2:S30'));
  sh.getRange('U1').setValue('ロッドやリールを買い替えたら行を書き換える。集計は硬さ・径帯・形状で組んであるので、製品が替わっても表の形は変わらない。')
    .setFontSize(9).setFontStyle('italic');
  // 重量帯の下限は自動：最初の帯は0、あとは1つ前の帯との中間（10と15なら12.5）。
  // 帯はL列に重さを並べるだけでよい
  const lower = [];
  for (let r = 2; r <= 60; r++) {
    lower.push([r === 2 ? '=IF(L2="","",0)' : `=IF(L${r}="","",(L${r - 1}+L${r})/2)`]);
  }
  sh.getRange('K2:K60').setFormulas(lower).setBackground(C_AUTO);
  sh.getRange('U2').setValue('重量帯はL列に重さを小さい順に並べるだけ。境目（K列）は隣の帯との中間を自動で出す。18gは20gの列に入る。帯を変えたら「計測集計を作り直す」。')
    .setFontSize(9).setFontStyle('italic');
}

function buildLog_(ss) {
  const { sh, created } = ensureSheet_(ss, '釣行');
  if (created) sh.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
  styleHeader_(sh, LOG_HEADERS.length);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(3);
  sh.getRange('B2:B').setNumberFormat('yyyy/mm/dd');
  sh.getRange('L2:M').setNumberFormat('HH:mm');
  sh.setColumnWidth(15, 420);
  sh.getRange('O2:O').setWrap(true);

  setListFromRange_(sh.getRange('C2:C'), placesRange_(ss));
  const map = { E: '潮色', F: '潮流向き', G: '潮流強さ', H: '風向', I: '風速', J: '波', K: '釣法', N: 'ベイト' };
  Object.keys(map).forEach(col => setListFromRange_(sh.getRange(col + '2:' + col), choiceRange_(ss, map[col])));
}

function buildMeasure_(ss) {
  const { sh, created } = ensureSheet_(ss, '計測');
  if (created) sh.getRange(1, 1, 1, MEAS_INPUT.length).setValues([MEAS_INPUT]);

  // 自動の列は見出しごと1セルの配列数式。行を足しても数式を伸ばす必要がない
  const T = 'タックルマスタ';
  const auto = [
    `={"硬さ";ARRAYFORMULA(IF(D2:D="","",IFERROR(VLOOKUP(D2:D,${T}!$A$2:$B,2,FALSE),"")))}`,
    `={"スプール径mm";ARRAYFORMULA(IF(E2:E="","",IFERROR(VLOOKUP(E2:E,${T}!$D$2:$E,2,FALSE),"")))}`,
    `={"径帯";ARRAYFORMULA(IF(K2:K="","",IFERROR(LOOKUP(K2:K,FILTER(${T}!$N$2:$N,${T}!$N$2:$N<>""),FILTER(${T}!$O$2:$O,${T}!$N$2:$N<>"")),"")))}`,
    `={"重量g";ARRAYFORMULA(IF(F2:F="","",IFERROR(VLOOKUP(F2:F,${T}!$G$2:$H,2,FALSE),"")))}`,
    `={"重量帯";ARRAYFORMULA(IF(M2:M="","",IFERROR(LOOKUP(M2:M,FILTER(${T}!$K$2:$K,${T}!$K$2:$K<>""),FILTER(${T}!$L$2:$L,${T}!$K$2:$K<>"")),"")))}`,
    `={"形状";ARRAYFORMULA(IF(F2:F="","",IFERROR(VLOOKUP(F2:F,${T}!$G$2:$I,3,FALSE),"")))}`,
  ];
  // 自動列に手で何か入っていると配列数式が展開できないので、2行目以降を空にする
  sh.getRange(2, 10, Math.max(1, sh.getMaxRows() - 1), MEAS_AUTO.length).clearContent();
  sh.getRange(1, 10, 1, auto.length).setFormulas([auto]);

  styleHeader_(sh, MEAS_INPUT.length + MEAS_AUTO.length);
  sh.getRange(1, 10, 1, MEAS_AUTO.length).setBackground('#7F7F7F');
  sh.getRange(2, 10, sh.getMaxRows() - 1, MEAS_AUTO.length).setBackground(C_AUTO);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(3);
  sh.getRange('B2:B').setNumberFormat('yyyy/mm/dd');
  sh.setColumnWidth(4, 160).setColumnWidth(5, 160).setColumnWidth(6, 180).setColumnWidth(9, 260);

  setListFromRange_(sh.getRange('C2:C'), placesRange_(ss));
  const tm = ss.getSheetByName(T);
  setListFromRange_(sh.getRange('D2:D'), tm.getRange('A2:A200'));
  setListFromRange_(sh.getRange('E2:E'), tm.getRange('D2:D200'));
  setListFromRange_(sh.getRange('F2:F'), tm.getRange('G2:G300'));
}

function buildDepth_(ss) {
  const { sh, created } = ensureSheet_(ss, '水深');
  if (created) sh.getRange(1, 1, 1, DEPTH_HEADERS.length).setValues([DEPTH_HEADERS]);
  styleHeader_(sh, DEPTH_HEADERS.length);
  sh.setFrozenRows(1);
  sh.getRange('A2:A').setNumberFormat('yyyy/mm/dd');
  sh.setColumnWidth(3, 200).setColumnWidth(8, 360);
  setListFromRange_(sh.getRange('B2:B'), placesRange_(ss));
  setListFromRange_(sh.getRange('G2:G'), choiceRange_(ss, '測り方'));
}

// ---------------------------------------------------------------
// 計測集計（断面A・B・C）

// 重量帯の列数は、タックルマスタに書いてある帯の数で決まる（buildSummary の中で数える）。
// 帯を増やしたら「釣りログ → 計測集計を作り直す」で列が増える。
const SUM_ROWS = 12;          // 硬さ・径帯・形状を最大いくつ並べるか

function buildSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('計測集計');
  const kept = sh ? readSelections_(sh) : {};   // 作り直す前の「対象ロッド／リール」を覚えておく
  if (sh) {
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();   // 列数が変わると結合位置がずれるため
    sh.clear();
    sh.clearConditionalFormatRules();
    sh.getDataRange().clearDataValidations();
  } else {
    sh = ss.insertSheet('計測集計');
  }
  const tm0 = ss.getSheetByName('タックルマスタ');
  const SUM_WEIGHT_COLS = Math.max(1, tm0.getRange('L2:L100').getValues().filter(r => r[0] !== '').length);
  const SUM_BLOCK_W = 1 + SUM_WEIGHT_COLS + 1;   // ラベル列＋重量列＋すき間

  // 3つの断面を横に並べるので、新しいシートの26列では足りないことがある
  const needCols = 3 * SUM_BLOCK_W;
  if (sh.getMaxColumns() < needCols) sh.insertColumnsAfter(sh.getMaxColumns(), needCols - sh.getMaxColumns());

  const tm = ss.getSheetByName('タックルマスタ');
  const firstRod  = tm.getRange('A2').getValue();
  const firstReel = tm.getRange('D2').getValue();

  const blocks = [
    { title: '断面A：リールを固定して、ロッドの硬さ × ルアー重量',
      sel: [['対象リール', M.reel, firstReel, 'D2:D200']],
      rowLabel: '硬さ', rowCol: M.hard, labelSrc: 'Q' },
    { title: '断面B：ロッドを固定して、スプール径 × ルアー重量',
      sel: [['対象ロッド', M.rod, firstRod, 'A2:A200']],
      rowLabel: '径帯', rowCol: M.dia, labelSrc: 'O' },
    { title: '断面C：ロッドとリールを固定して、ルアーの形状 × 重量',
      sel: [['対象ロッド', M.rod, firstRod, 'A2:A200'], ['対象リール', M.reel, firstReel, 'D2:D200']],
      rowLabel: '形状', rowCol: M.shape, labelSrc: 'S' },
  ];
  const tables = [
    { title: '飛距離の平均（m）', kind: 'avg', fmt: '0.0', grad: 'up' },
    { title: '投数',             kind: 'cnt', fmt: '0',   grad: null },
    { title: 'ばらつき（標準偏差 m）', kind: 'sd', fmt: '0.0', grad: 'down' },
    { title: '最大飛距離（m）・参考', kind: 'max', fmt: '0.0', grad: 'up' },
  ];

  blocks.forEach((b, bi) => {
    const c0 = 1 + bi * SUM_BLOCK_W;
    sh.getRange(1, c0).setValue(b.title).setFontSize(13).setFontWeight('bold');

    // 固定する条件
    const selCells = b.sel.map((s, i) => {
      const r = 3 + i;
      sh.getRange(r, c0).setValue(s[0]);
      const cell = sh.getRange(r, c0 + 1, 1, 3).merge();
      const keep = kept[b.title + '|' + s[0]];
      cell.setValue(keep !== undefined && keep !== '' ? keep : s[2]).setBackground(C_PICK).setFontWeight('bold').setHorizontalAlignment('center');
      setListFromRange_(sh.getRange(r, c0 + 1), tm.getRange(s[3]));
      return { measCol: s[1], ref: '$' + col_(c0 + 1) + '$' + r };
    });

    let top = 6;
    tables.forEach(t => {
      sh.getRange(top, c0).setValue(t.title).setFontWeight('bold').setFontColor(C_HEAD);
      const h = top + 1;
      // 見出し行：ラベル＋重量帯（タックルマスタから）
      const head = [b.rowLabel];
      for (let j = 0; j < SUM_WEIGHT_COLS; j++) head.push(`=IF(タックルマスタ!$L${j + 2}="","",タックルマスタ!$L${j + 2})`);
      sh.getRange(h, c0, 1, head.length).setFormulas([head.map(v => String(v).charAt(0) === '=' ? v : '')]);
      sh.getRange(h, c0).setValue(b.rowLabel);
      sh.getRange(h, c0, 1, head.length).setBackground(C_HEAD).setFontColor('#FFFFFF')
        .setFontWeight('bold').setHorizontalAlignment('center');
      sh.getRange(h, c0 + 1, 1, SUM_WEIGHT_COLS).setNumberFormat('0"g"');

      // 中身
      const body = [];
      for (let i = 0; i < SUM_ROWS; i++) {
        const r = h + 1 + i;
        const lab = '$' + col_(c0) + r;
        const line = [`=IF(タックルマスタ!$${b.labelSrc}${i + 2}="","",タックルマスタ!$${b.labelSrc}${i + 2})`];
        for (let j = 0; j < SUM_WEIGHT_COLS; j++) {
          const w = col_(c0 + 1 + j) + '$' + h;
          line.push(cellFormula_(t.kind, selCells, b.rowCol, lab, w));
        }
        body.push(line);
      }
      const bodyRange = sh.getRange(h + 1, c0, SUM_ROWS, 1 + SUM_WEIGHT_COLS);
      bodyRange.setFormulas(body).setHorizontalAlignment('center');
      sh.getRange(h + 1, c0 + 1, SUM_ROWS, SUM_WEIGHT_COLS).setNumberFormat(t.fmt);
      sh.getRange(h, c0, SUM_ROWS + 1, 1 + SUM_WEIGHT_COLS)
        .setBorder(true, true, true, true, true, true, '#BFBFBF', SpreadsheetApp.BorderStyle.SOLID);
      if (t.grad) addGradient_(sh, sh.getRange(h + 1, c0 + 1, SUM_ROWS, SUM_WEIGHT_COLS), t.grad);

      top = h + SUM_ROWS + 2;
    });

    sh.setColumnWidth(c0, 120);
    sh.setColumnWidths(c0 + 1, SUM_WEIGHT_COLS, 62);
    sh.setColumnWidth(c0 + SUM_WEIGHT_COLS + 1, 24);
  });

  const notesTop = 6 + tables.length * (SUM_ROWS + 3) + 1;
  [
    '黄色のセルを切り替えると、その条件の面が出る。',
    '見るのは「平均」。バックラッシュせずに毎回出せる距離。最大は条件が揃ったときの参考値。',
    'ばらつきが平均の1割以内なら、その平均は信用していい。投数が5未満のマスはあてにならない。',
    '行は製品名ではなく硬さ・径帯・形状。道具を買い替えても表の形は変わらない。',
  ].forEach((t, i) => sh.getRange(notesTop + i, 1).setValue(t).setFontSize(9).setFontStyle('italic'));
}

/** 計測集計の1行目の見出しと、その下の「対象〇〇」の選択値を拾う */
function readSelections_(sh) {
  const out = {};
  const w = sh.getLastColumn();
  if (w < 1) return out;
  const vals = sh.getRange(1, 1, 5, w).getValues();
  for (let c = 0; c < w; c++) {
    const title = String(vals[0][c] || '');
    if (title.indexOf('断面') !== 0) continue;
    for (let r = 2; r <= 4; r++) {
      const label = String(vals[r][c] || '');
      if (label.indexOf('対象') === 0) out[title + '|' + label] = vals[r][c + 1];
    }
  }
  return out;
}

/** 集計の1マス分の数式 */
function cellFormula_(kind, sels, rowCol, labRef, wRef) {
  const R = c => `計測!$${c}$2:$${c}`;
  const guard = `OR(${labRef}="",${wRef}="")`;
  if (kind === 'avg' || kind === 'cnt') {
    const crit = sels.map(s => `${R(s.measCol)},${s.ref}`)
      .concat([`${R(rowCol)},${labRef}`, `${R(M.band)},${wRef}`]).join(',');
    if (kind === 'avg') return `=IF(${guard},"",IFERROR(AVERAGEIFS(${R(M.dist)},${crit}),""))`;
    return `=IF(${guard},"",COUNTIFS(${crit},${R(M.dist)},"<>"))`;
  }
  const cond = sels.map(s => `${R(s.measCol)}=${s.ref}`)
    .concat([`${R(rowCol)}=${labRef}`, `${R(M.band)}=${wRef}`, `${R(M.dist)}<>""`]).join(',');
  const fn = kind === 'sd' ? 'STDEV' : 'MAX';
  return `=IF(${guard},"",IFERROR(${fn}(FILTER(${R(M.dist)},${cond})),""))`;
}

function addGradient_(sh, range, dir) {
  const lo = dir === 'up' ? C_RED : C_GRN;
  const hi = dir === 'up' ? C_GRN : C_RED;
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint(lo)
    .setGradientMidpointWithValue(C_YEL, SpreadsheetApp.InterpolationType.PERCENTILE, '50')
    .setGradientMaxpoint(hi)
    .setRanges([range]).build();
  const rules = sh.getConditionalFormatRules();
  rules.push(rule);
  sh.setConditionalFormatRules(rules);
}

// ---------------------------------------------------------------
// 共通

function ensureSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (sh) return { sh: sh, created: false };
  sh = ss.insertSheet(name);
  return { sh: sh, created: true };
}

function arrangeSheets_(ss) {
  const order = ['釣行', '計測', '計測集計', '水深', '場所マスター', 'タックルマスタ', '選択肢'];
  order.forEach((n, i) => {
    const sh = ss.getSheetByName(n);
    if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1); }
  });
  // 最初からある空のシートを片付ける
  ['シート1', 'Sheet1'].forEach(n => {
    const sh = ss.getSheetByName(n);
    if (sh && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
  ss.setActiveSheet(ss.getSheetByName('釣行'));
}

function styleHeader_(sh, n) {
  styleHeaderCell_(sh.getRange(1, 1, 1, n));
}

function styleHeaderCell_(range) {
  range.setBackground(C_HEAD).setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
}

function put_(sh, row, col, rows) {
  if (rows.length) sh.getRange(row, col, rows.length, rows[0].length).setValues(rows);
}

function setListFromRange_(target, source) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(source, true)
    .setAllowInvalid(true)          // リストに無い言葉も書ける。表記ゆれ防止の補助として使う
    .build();
  target.setDataValidation(rule);
}

function placesRange_(ss) {
  return ss.getSheetByName('場所マスター').getRange('A2:A300');
}

function choiceRange_(ss, key) {
  const sh = ss.getSheetByName('選択肢');
  const keys = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const i = keys.indexOf(key);
  if (i < 0) throw new Error('選択肢シートに「' + key + '」の列がありません');
  return sh.getRange(2, i + 1, 100, 1);
}

function col_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
  return s;
}
