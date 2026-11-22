/**
 * 釣りログ ― 気象データの自動取得
 *
 * 釣行・計測シートの各行に、その日その場所の前々日・前日・当日の雨量、
 * 気圧、気温、最大風速を入れる。データ元は Open-Meteo の過去データAPI（無料・登録不要）。
 *
 * 場所の座標は「場所マスター」の「座標」列から読む。
 * Googleマップで右クリック→座標をコピーした "36.8419, 137.4198" をそのまま貼ればよい。
 * 緯度と経度の順番が逆でも自動で直す。
 *
 * 過去データは5日ほど遅れて確定する。当日や直近の行は「未確定」として飛ばし、
 * 後日もう一度実行すれば埋まる。埋まった行は取り直さない。
 */

// ===== 設定 =====
const SHEET_PLACE  = '場所マスター';

// 気象を入れるシート。列は setup.gs の WEATHER_COL（P列）。
// ファイルの読み込み順に左右されないよう、関数の中で組み立てる
function targets_() {
  return [
    { sheet: '釣行', dateCol: 2, placeCol: 3, outCol: WEATHER_COL },
    { sheet: '計測', dateCol: 2, placeCol: 3, outCol: WEATHER_COL },
  ];
}
const TIMEZONE     = 'Asia/Tokyo';
const MAX_PER_RUN  = 40;   // 1回の実行で取りに行く最大件数（APIが混んでいるとき用）

const OUT_HEADERS = [
  '前々日雨量mm',
  '前日雨量mm',
  '当日雨量mm',
  '当日平均気圧hPa',
  '当日最高気温',
  '当日最低気温',
  '当日最大風速m/s',
  '気象取得日時',
];

function fillWeather() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const badPlaces = [];
  const places = loadPlaces_(ss, badPlaces);
  if (Object.keys(places).length === 0) {
    alert_('「' + SHEET_PLACE + '」に場所と座標を入れてください\n'
         + (badPlaces.length ? '読めなかった行：' + badPlaces.join('、') : ''));
    return;
  }

  const cache = {};                       // 場所＋日付でまとめて1回だけ取る
  let fetched = 0;
  const report = [];

  for (const tg of targets_()) {
    const sh = ss.getSheetByName(tg.sheet);
    if (!sh) continue;

    ensureHeaders_(sh, tg.outCol);

    const lastRow = sh.getLastRow();
    if (lastRow < 2) continue;

    const width = Math.max(tg.outCol + OUT_HEADERS.length - 1, sh.getLastColumn());
    const rows  = sh.getRange(2, 1, lastRow - 1, width).getValues();

    let filled = 0, skipped = 0, notReady = 0, remaining = 0;
    const errors = [], noPlace = [];

    for (let i = 0; i < rows.length; i++) {
      const row    = rows[i];
      const rowNum = i + 2;

      const date  = toDate_(row[tg.dateCol - 1]);
      const place = String(row[tg.placeCol - 1] || '').trim();
      if (!date || !place) continue;

      const cur = row[tg.outCol - 1];
      if (cur !== '' && cur !== null && cur !== undefined) { skipped++; continue; }

      const pos = places[place];
      if (!pos) { if (noPlace.indexOf(place) < 0) noPlace.push(place); continue; }

      const key = place + '|' + fmt_(date);
      let w = cache[key];
      if (w === undefined) {
        if (fetched >= MAX_PER_RUN) { remaining++; continue; }
        w = fetchWeather_(pos.lat, pos.lng, date);
        cache[key] = w;
        fetched++;
        Utilities.sleep(200);
      }
      if (w && w.error) { if (errors.indexOf(w.error) < 0) errors.push(w.error); continue; }
      if (!w) { notReady++; continue; }

      sh.getRange(rowNum, tg.outCol, 1, OUT_HEADERS.length).setValues([[
        w.rain2, w.rain1, w.rain0, w.pressure, w.tmax, w.tmin, w.wind,
        Utilities.formatDate(new Date(), TIMEZONE, 'yyyy/MM/dd HH:mm'),
      ]]);
      filled++;
    }

    let line = tg.sheet + '：' + filled + '件';
    if (skipped)  line += '／取得済み' + skipped;
    if (notReady) line += '／未確定' + notReady;
    if (remaining) line += '／次回まわし' + remaining;
    if (noPlace.length) line += '／場所マスターに無い：' + noPlace.join('、');
    if (errors.length)  line += '\n  エラー：' + errors.join(' / ');
    report.push(line);
  }

  if (badPlaces.length) report.push('座標が読めない場所：' + badPlaces.join('、'));
  alert_(report.join('\n'));
}

// ===== 以下、内部処理 =====

function ensureHeaders_(sh, outCol) {
  const cur = sh.getRange(1, outCol, 1, OUT_HEADERS.length).getValues()[0];
  if (cur.join('') === OUT_HEADERS.join('')) return;
  sh.getRange(1, outCol, 1, OUT_HEADERS.length)
    .setValues([OUT_HEADERS])
    .setFontWeight('bold');
}

/**
 * 場所マスターを {場所名: {lat, lng}} にする。
 * 「座標」列（Googleマップからコピーした "36.8419, 137.4198" をそのまま貼る）を優先。
 * 無ければ「緯度」「経度」の2列を見る。どちらも順番の取り違えは自動で直す。
 */
function loadPlaces_(ss, badPlaces) {
  badPlaces = badPlaces || [];
  const sh = ss.getSheetByName(SHEET_PLACE);
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;

  const vals = sh.getRange(1, 1, sh.getLastRow(), Math.max(2, sh.getLastColumn())).getValues();
  const head = vals[0].map(v => String(v).trim());
  const iName  = head.indexOf('場所');
  const iCoord = head.indexOf('座標');
  const iLat   = head.indexOf('緯度');
  const iLng   = head.indexOf('経度');
  if (iName < 0) return out;

  for (let r = 1; r < vals.length; r++) {
    const name = String(vals[r][iName] || '').trim();
    if (!name) continue;

    // 行のどこかに座標があれば拾う。座標列 → 緯度/経度列 → 行全体の数値、の順で探す
    let nums = [];
    if (iCoord >= 0) nums = pickNumbers_(vals[r][iCoord]);
    if (nums.length < 2 && iLat >= 0 && iLng >= 0) {
      nums = pickNumbers_(vals[r][iLat]).concat(pickNumbers_(vals[r][iLng]));
    }
    if (nums.length < 2) {
      // 見出しがどうであれ、その行にある数値2つを座標とみなす（列ずれ対策）
      for (let c = 0; c < vals[r].length && nums.length < 2; c++) {
        if (c === iName) continue;
        nums = nums.concat(pickNumbers_(vals[r][c]));
      }
    }
    if (nums.length < 2) { badPlaces.push(name); continue; }

    const pos = assignLatLng_(nums[0], nums[1]);
    if (!pos) { badPlaces.push(name); continue; }
    out[name] = pos;
  }
  return out;
}

/** セルの中身から有限の数値だけを取り出す。null・空文字・文字列まじりも扱える */
function pickNumbers_(v) {
  if (v === null || v === undefined || v === '') return [];
  if (typeof v === 'number') return isFinite(v) ? [v] : [];
  const m = String(v).match(/-?\d+(\.\d+)?/g);
  if (!m) return [];
  return m.map(Number).filter(x => isFinite(x));
}

/**
 * 2つの数値のどちらが緯度でどちらが経度かを決める。
 * 日本の範囲（緯度20〜46、経度122〜154）で判定し、外れていれば ±90 で判定する。
 */
function assignLatLng_(a, b) {
  if (typeof a !== 'number' || typeof b !== 'number' || !isFinite(a) || !isFinite(b)) return null;
  if (a === 0 && b === 0) return null;
  const inJpLat = v => v >= 20 && v <= 46;
  const inJpLng = v => v >= 122 && v <= 154;

  if (inJpLat(a) && inJpLng(b)) return { lat: a, lng: b };
  if (inJpLat(b) && inJpLng(a)) return { lat: b, lng: a };   // 入れ替わっていた

  // 日本国外。緯度は±90までしかない、という一般則で判定する
  if (Math.abs(a) <= 90 && Math.abs(b) <= 180 && Math.abs(b) > 90) return { lat: a, lng: b };
  if (Math.abs(b) <= 90 && Math.abs(a) <= 180 && Math.abs(a) > 90) return { lat: b, lng: a };
  if (Math.abs(a) <= 90 && Math.abs(b) <= 90) return { lat: a, lng: b };   // 判別不能。並び順を信じる
  return null;
}

/** 当日・前日・前々日の気象を取る。当日分がまだ無ければ null */
function fetchWeather_(lat, lng, date) {
  const d0 = fmt_(date);
  const d2 = fmt_(addDays_(date, -2));

  const url = 'https://archive-api.open-meteo.com/v1/archive'
    + '?latitude='  + lat
    + '&longitude=' + lng
    + '&start_date=' + d2
    + '&end_date='   + d0
    + '&daily=precipitation_sum,temperature_2m_max,temperature_2m_min,wind_speed_10m_max'
    + '&hourly=pressure_msl'
    + '&wind_speed_unit=ms'
    + '&timezone=' + encodeURIComponent(TIMEZONE);

  let json;
  try {
    const res  = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const body = res.getContentText();
    if (res.getResponseCode() !== 200) {
      Logger.log('HTTP ' + res.getResponseCode() + ' : ' + body);
      let reason = body;
      try { reason = JSON.parse(body).reason || body; } catch (e2) {}
      return { error: 'HTTP ' + res.getResponseCode() + '（緯度' + lat + ' 経度' + lng + '）: ' + reason };
    }
    json = JSON.parse(body);
  } catch (e) {
    Logger.log('取得失敗: ' + e);
    return { error: String(e) };
  }

  const daily = json.daily;
  if (!daily || !daily.time || daily.time.length < 3) {
    return { error: '日別データが返ってきません（緯度' + lat + ' 経度' + lng + '）' };
  }

  const i0 = daily.time.indexOf(d0);           // 当日の位置
  if (i0 < 0) return { error: d0 + ' の行が返ってきません' };
  const rain0 = daily.precipitation_sum[i0];
  if (rain0 === null || rain0 === undefined) return null;   // まだ確定していない

  return {
    rain2:    num_(daily.precipitation_sum[i0 - 2]),
    rain1:    num_(daily.precipitation_sum[i0 - 1]),
    rain0:    num_(rain0),
    tmax:     num_(daily.temperature_2m_max[i0]),
    tmin:     num_(daily.temperature_2m_min[i0]),
    wind:     num_(daily.wind_speed_10m_max[i0]),
    pressure: meanPressure_(json.hourly, d0),
  };
}

/** 当日24時間分の海面気圧の平均 */
function meanPressure_(hourly, dayStr) {
  if (!hourly || !hourly.time) return '';
  let sum = 0, n = 0;
  for (let i = 0; i < hourly.time.length; i++) {
    if (hourly.time[i].indexOf(dayStr) !== 0) continue;
    const v = hourly.pressure_msl[i];
    if (v === null || v === undefined) continue;
    sum += v; n++;
  }
  return n ? Math.round(sum / n * 10) / 10 : '';
}

function toDate_(v) {
  if (v instanceof Date) return v;
  const s = String(v || '').trim();
  if (!s) return null;
  const m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function addDays_(d, n) {
  const x = new Date(d.getTime());
  x.setDate(x.getDate() + n);
  return x;
}

function fmt_(d) {
  return Utilities.formatDate(d, TIMEZONE, 'yyyy-MM-dd');
}

function num_(v) {
  return (v === null || v === undefined) ? '' : v;
}

/**
 * 画面右下に出して知らせる。alertと違い、押されるまで待たないので
 * 実行が止まったように見えることがない。
 */
function alert_(msg) {
  Logger.log(msg);
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, '釣りログ', 30);
  } catch (e) {}   // トリガー実行時はUIが無い
}

/**
 * 週1回まとめて取りたい場合。一度だけ実行すればトリガーが登録される。
 * （毎週月曜の朝に、まだ埋まっていない行を取りにいく）
 */
function installWeeklyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'fillWeather') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('fillWeather')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(7)
    .create();
  alert_('毎週月曜7時に自動取得するよう設定しました');
}
