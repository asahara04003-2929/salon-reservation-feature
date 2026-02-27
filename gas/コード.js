/** =========================
 *  CONFIG
 * ========================= */
// スプレットシートID
const SPREADSHEET_ID = '1gBiXUNrYO_kloeI4zIiEqlGiZn8G2BkSs3VcIAYeXjU';

// シート名定義
const SHEET_PLAN = 'PLAN_MASTER';
const SHEET_USERS = 'USERS';
const SHEET_RESERVATIONS = 'RESERVATIONS';
const SHEET_CALENDAR = 'CALENDAR'; // ←あなたのカレンダー表示用シート名に合わせて変えてOK
const SHEET_TODAY = "TODAY";
const SHEET_TODAY_GANTT = "TODAY_GANTT_CHART";

// CALENDARの描画位置定義
const CAL_THIS_START_ROW = 4;   // 今月カレンダー本体の開始行（B4）
const CAL_NEXT_START_ROW = 12;  // ★来月カレンダー本体の開始行（B??）←ここだけ調整してね
const CAL_THIS_TITLE_CELL = "B2";   // 今月タイトルを出すセル（左上）
const CAL_NEXT_TITLE_CELL = "B10";  // 来月タイトルを出すセル（左上）←ここだけ調整してね
const CAL_START_COL = 2;        // B列
const CAL_COLS = 7;             // 日〜土

// 日本の祝日取得用の定義
const HOLIDAY_CALENDAR_ID = "ja.japanese#holiday@group.v.calendar.google.com"; // 日本の祝日

/**
 * メニュー追加（任意）
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('自作マクロ')
    .addSeparator()
    .addItem('本日予約一覧更新', 'refreshTodayReservations') // ★追加
    .addSeparator()
    .addItem('本日ガントチャート更新', 'renderTodayGanttChart') // ★追加
    .addSeparator()
    .addItem('予約カレンダー更新', 'renderReservationCalendar') // ★追加
    .addSeparator()
    .addItem('祝日情報取得', 'syncJapaneseHolidaysToBlackouts') // ★追加
    .addToUi();
}

/**
 * C1(YYYY/MM) を見て、RESERVATIONS を月間カレンダーに描画
 */
/**
 * 今月 + 来月 を RESERVATIONS からカレンダーに描画（C1は使わない）
 */
function renderReservationCalendar() {
  const calSh = sh_(SHEET_CALENDAR);
  const resSh = sh_(SHEET_RESERVATIONS);
  if (!calSh) throw new Error(`Sheet not found: ${SHEET_CALENDAR}`);
  if (!resSh) throw new Error(`Sheet not found: ${SHEET_RESERVATIONS}`);

  // JST基準で「今月」を決める
  const tz = "Asia/Tokyo";
  const now = new Date();
  const y = Number(Utilities.formatDate(now, tz, "yyyy"));
  const m = Number(Utilities.formatDate(now, tz, "MM")); // 1-12

  // 今月・来月を描画
  renderOneMonthCalendar_(calSh, resSh, y, m, CAL_THIS_START_ROW, CAL_THIS_TITLE_CELL);
  const next = addMonth_(y, m, 1); // {year, month}
  renderOneMonthCalendar_(calSh, resSh, next.year, next.month, CAL_NEXT_START_ROW, CAL_NEXT_TITLE_CELL);
}

/**
 * 指定年月(1-12)の予約を、指定開始行のカレンダー枠へ描画
 * startRow は「日付が入るマスの開始行（B列）」を渡す
 */
function renderOneMonthCalendar_(calSh, resSh, year, month, startRow, titleCellA1) {
  // ★タイトル（YYYY/MM）
  const ym = `${year}/${String(month).padStart(2, "0")}`;
  if (titleCellA1) calSh.getRange(titleCellA1).setValue(ym);

  // その月の1日・末日
  const firstDay = new Date(year, month - 1, 1, 0, 0, 0, 0);
  const lastDay  = new Date(year, month, 0, 23, 59, 59, 999);
  const daysInMonth = new Date(year, month, 0).getDate();

  // 日曜始まり
  const offset = firstDay.getDay(); // 0=日..6=土
  const totalCells = offset + daysInMonth;
  const weeks = Math.ceil(totalCells / 7); // 5 or 6

  // 表示範囲クリア（B?:H?）
  const clearRange = calSh.getRange(startRow, CAL_START_COL, weeks, CAL_COLS);
  clearRange.clearContent();
  clearRange.setVerticalAlignment("top");
  clearRange.setWrap(true);

  // ===== 予約データ取得（必要列だけ） =====
  const lastRow = resSh.getLastRow();
  if (lastRow < 2) return;

  const header = resSh.getRange(1, 1, 1, resSh.getLastColumn()).getValues()[0].map(String);
  const ridx = indexMap_(header);
  requiredCols_(ridx, ["reserved_start", "reserved_end", "status", "line_user_id"]);

  const n = lastRow - 1;

  // 必須列
  const colStatus = resSh.getRange(2, ridx.status + 1, n, 1).getValues();
  const colStart  = resSh.getRange(2, ridx.reserved_start + 1, n, 1).getValues();
  const colLineId = resSh.getRange(2, ridx.line_user_id + 1, n, 1).getValues();

  // 任意列（スナップショット優先）
  const hasNameSnap = ridx.name_snapshot !== undefined;
  const colNameSnap = hasNameSnap ? resSh.getRange(2, ridx.name_snapshot + 1, n, 1).getValues() : null;

  const hasPlanNamesSnap = ridx.plan_names_snapshot !== undefined;
  const colPlanNamesSnap = hasPlanNamesSnap ? resSh.getRange(2, ridx.plan_names_snapshot + 1, n, 1).getValues() : null;

  const hasPlanNameSnap = ridx.plan_name_snapshot !== undefined;
  const colPlanNameSnap = (!hasPlanNamesSnap && hasPlanNameSnap)
    ? resSh.getRange(2, ridx.plan_name_snapshot + 1, n, 1).getValues()
    : null;

  // USERS名寄せ（既存どおり）
  const userNameByLineId = buildUserNameMap_();

  /** @type {Record<string, string[]>} */
  const byDay = {};

  for (let i = 0; i < n; i++) {
    const status = String(colStatus[i][0] || "");
    if (status !== "CONFIRMED") continue;

    const start = coerceToDate_(colStart[i][0]);
    if (!start) continue;

    // 当月外は除外（開始日で判定）
    if (start < firstDay || start > lastDay) continue;

    const dayKey = formatYmd_(start); // YYYY-MM-DD
    const hhmm = formatHm_(start);

    const lineUserId = String(colLineId[i][0] || "").trim();

    const nameSnap = hasNameSnap ? String(colNameSnap[i][0] || "").trim() : "";
    const customer = nameSnap
      ? nameSnap
      : (userNameByLineId[lineUserId] || lineUserId || "（不明）");

    let planName = "";
    if (hasPlanNamesSnap) {
      planName = String(colPlanNamesSnap[i][0] || "").trim();
    } else if (colPlanNameSnap) {
      planName = String(colPlanNameSnap[i][0] || "").trim();
    }

    const text = `${hhmm} ${customer}${planName ? " " + planName : ""}`.trim();

    if (!byDay[dayKey]) byDay[dayKey] = [];
    byDay[dayKey].push(text);
  }

  // 時刻順
  Object.keys(byDay).forEach(k => byDay[k].sort());

  // 描画
  for (let d = 1; d <= daysInMonth; d++) {
    const dateObj = new Date(year, month - 1, d, 0, 0, 0, 0);
    const cellIndex = offset + (d - 1);
    const weekIndex = Math.floor(cellIndex / 7);
    const colIndex = cellIndex % 7;

    const row = startRow + weekIndex;
    const col = CAL_START_COL + colIndex;

    const dayKey = formatYmd_(dateObj);
    const lines = byDay[dayKey] || [];

    const cellText = lines.length ? `${d}\n${lines.join("\n")}` : `${d}`;
    calSh.getRange(row, col).setValue(cellText);
  }
}

/** year/month(1-12) に月加算して返す */
function addMonth_(year, month, add) {
  const d = new Date(year, month - 1 + add, 1);
  return { year: d.getFullYear(), month: d.getMonth() + 1 };
}


/***************
 * helpers
 ***************/
function parseYearMonth_(ym) {
  const m = /^(\d{4})\/(\d{1,2})$/.exec(ym);
  if (!m) throw new Error('C1は YYYY/MM 形式で入力してください（例: 2025/01）');
  const year = Number(m[1]);
  const month = Number(m[2]);
  if (!Number.isFinite(year) || !Number.isFinite(month) || month < 1 || month > 12) {
    throw new Error('C1の年月が不正です');
  }
  return { year, month };
}


function requiredCols_(idx, cols) {
  cols.forEach(c => {
    if (idx[c] === undefined) throw new Error(`RESERVATIONSに必要な列がありません: ${c}`);
  });
}


function formatHm_(d) {
  const hh = String(d.getHours()).padStart(2, '0');
  const mm = String(d.getMinutes()).padStart(2, '0');
  return `${hh}:${mm}`;
}

function buildUserNameMap_() {
  const sh = sh_(SHEET_USERS);
  if (!sh) return {};

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return {};

  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const idx = indexMap_(header);
  if (idx.line_user_id === undefined || idx.name === undefined) return {};

  const n = lastRow - 1;

  // 必要列だけ取得
  const colLine = sh.getRange(2, idx.line_user_id + 1, n, 1).getValues();
  const colName = sh.getRange(2, idx.name + 1, n, 1).getValues();

  const map = {};
  for (let i = 0; i < n; i++) {
    const lineId = String(colLine[i][0] || '').trim();
    const name   = String(colName[i][0] || '').trim();
    if (lineId && name) map[lineId] = name;
  }
  return map;
}


/** =========================
 *  HTTP Entry
 * ========================= */
function doGet(e) {
  try {
    const action = (e.parameter.action || '').trim();
    if (!action) return json_({ ok: false, error: 'MISSING_ACTION' });

    switch (action) {
      case 'plans':
        return cachedJson_(
          `plans:${cacheVer_()}`,     // バスター込み
          60 * 30,                   // 30分
          () => ({ ok: true, plans: listPlans_() })
        );

      case 'me': {
        const lineUserId = reqParam_(e, 'line_user_id');
        return cachedJson_(
          `me:${cacheVer_()}:${lineUserId}`,
          60 * 5, // 5分
          () => {
            const user = getUserByLineId_(lineUserId);
            return ({ ok: true, exists: !!user, user: user || null });
          }
        );
      }

      case 'my_reservations': {
        const lineUserId = reqParam_(e, 'line_user_id');
        const status = (e.parameter.status || 'CONFIRMED').trim();
        return cachedJson_(
          `myres:${cacheVer_()}:${lineUserId}:${status}`,
          30, // 30秒
          () => ({ ok: true, reservations: listReservationsByUser_(lineUserId, status) })
        );
      }

      case 'availability': {
        const date = reqParam_(e, 'date');
        const planId = reqParam_(e, 'plan_id');
        return cachedJson_(
          `av:${cacheVer_()}:${date}:${planId}`,
          20, // 20秒
          () => ({ ok:true, ...getAvailability_(date, planId) })
        );
      }

      case 'availability_range': {
        const from = reqParam_(e, 'from');
        const days = Number((e.parameter.days || '7').toString().trim());
        const durationMinParam = (e.parameter.duration_min || '').toString().trim();
        const planId = (e.parameter.plan_id || '').toString().trim();
        const durationMin = durationMinParam ? Number(durationMinParam) : null;

        return cachedJson_(
          `avr:${cacheVer_()}:${from}:${days}:${durationMin||''}:${planId||''}`,
          20, // 20秒
          () => ({ ok:true, ...getAvailabilityRangeByDuration_(from, days, planId, durationMin) })
        );
      }

      case 'availability_range_materials': {
        const from = reqParam_(e, 'from');
        const days = Number((e.parameter.days || '7').toString().trim());
        const durationMinParam = (e.parameter.duration_min || '').toString().trim();
        const planId = (e.parameter.plan_id || '').toString().trim();
        const durationMin = durationMinParam ? Number(durationMinParam) : null;

        return cachedJson_(
          `arm:${cacheVer_()}:${from}:${days}:${durationMin||''}:${planId||''}`,
          20, // 20秒
          () => ({ ok:true, ...getAvailabilityRangeMaterialsByDuration_(from, days, planId, durationMin) })
        );
      }

      default:
        return json_({ ok: false, error: 'UNKNOWN_ACTION', action });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err), stack: err && err.stack ? String(err.stack) : null });
  }
}

function doPost(e) {
  try {
    const body = e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
    const action = (body.action || '').trim();
    if (!action) return json_({ ok: false, error: 'MISSING_ACTION' });

    switch (action) {
      case 'users_upsert': {
        const user = upsertUser_(body);
        bumpCacheVer_();                 // ★追加（ユーザー情報キャッシュ破棄）
        return json_({ ok: true, user });
      }

      case 'reserve': {
        const result = createReservation_(body);
        bumpCacheVer_();

        try { reservationDetailsUpdate(); } 
        catch (err) { console.error('[reservationDetailsUpdate] failed:', err); }

        return json_({ ok: true, reservation: result });
      }

      case 'cancel': {
        const result = cancelReservation_(body);
        bumpCacheVer_();

        try { reservationDetailsUpdate(); } 
        catch (err) { console.error('[reservationDetailsUpdate] failed:', err); }

        return json_({ ok: true, canceled: result });
      }

      default:
        return json_({ ok: false, error: 'UNKNOWN_ACTION', action });
    }
  } catch (err) {
    return json_({
      ok: false,
      error: String(err && err.message ? err.message : err),
      admin_phone: err && err.admin_phone ? String(err.admin_phone) : "",
      allowed_genders: err && err.allowed_genders ? err.allowed_genders : null, // ★追加
      stack: err && err.stack ? String(err.stack) : null
    });
  }
}

/** =========================
 *  Core: Plans
 * ========================= */
/********************************************************************
 * ✅ GAS側（PLAN_MASTER に descriptions 列を追加して返す）
 * 1) PLAN_MASTER シートに descriptions 列（ヘッダ名：descriptions）を追加
 * 2) 下の listPlans_ と getPlanById_ をコピペで差し替え
 ********************************************************************/

function listPlans_() {
  const sh = sh_(SHEET_PLAN);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const header = values[0];
  const idx = indexMap_(header);

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const isActive = String(row[idx.is_active] ?? '').toUpperCase() === 'TRUE';
    if (!isActive) continue;

    // descriptions（description でも拾う）
    const descIdx = (idx.descriptions !== undefined) ? idx.descriptions
                  : (idx.description !== undefined) ? idx.description
                  : undefined;
    const desc = (descIdx !== undefined) ? String(row[descIdx] ?? '') : '';

    // ★ order（未入力は最後尾）
    const ordRaw = (idx.order !== undefined) ? row[idx.order] : null;
    const ordNum = Number(ordRaw);
    const ord = Number.isFinite(ordNum) ? ordNum : 999999;

    out.push({
      plan_id: String(row[idx.plan_id]),
      plan_name: String(row[idx.plan_name]),
      duration_min: Number(row[idx.duration_min]),
      price: Number(row[idx.price]),
      descriptions: desc,
      order: ord, // ★返却にも含める（フロントで使いたければ）
    });
  }

  // ★ order 昇順 → plan_id 昇順（同順時の安定化）
  out.sort((a, b) => (a.order - b.order) || a.plan_id.localeCompare(b.plan_id));

  return out;
}

let PLAN_MAP_CACHE_ = null;
function getPlanMap_(){
  if (PLAN_MAP_CACHE_) return PLAN_MAP_CACHE_;
  const sh = sh_(SHEET_PLAN);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return (PLAN_MAP_CACHE_ = new Map());
  const header = values[0].map(String);
  const idx = indexMap_(header);
  const map = new Map();
  for (let r=1; r<values.length; r++){
    const row = values[r];
    const plan_id = String(row[idx.plan_id]||'').trim();
    if (!plan_id) continue;
    const isActive = String(row[idx.is_active] ?? '').toUpperCase() === 'TRUE';
    const ord = Number(row[idx.order]); 
    map.set(plan_id, {
      plan_id,
      plan_name: String(row[idx.plan_name]||''),
      duration_min: Number(row[idx.duration_min]||0),
      price: Number(row[idx.price]||0),
      is_active: isActive,
      order: Number.isFinite(ord) ? ord : 999999,
      descriptions: idx.descriptions!==undefined ? String(row[idx.descriptions]||'') : ''
    });
  }
  return (PLAN_MAP_CACHE_ = map);
}

function getPlanById_(planId){
  return getPlanMap_().get(String(planId)) || null;
}

/** =========================
 *  Core: Users
 * ========================= */
function getUserByLineId_(lineUserId) {
  const sh = sh_(SHEET_USERS);
  if (!sh) return null;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const idx = indexMap_(header);
  if (idx.line_user_id === undefined) return null;

  const n = lastRow - 1;

  // まず line_user_id 列だけ取得して一致行を探す（1列だけなので速い）
  const colLine = sh.getRange(2, idx.line_user_id + 1, n, 1).getValues();
  let hitRow = -1; // 0-based in data area
  for (let i = 0; i < n; i++) {
    if (String(colLine[i][0]) === String(lineUserId)) { hitRow = i; break; }
  }
  if (hitRow === -1) return null;

  // ヒットした行だけ “その行全列” を取得（1行なので軽い）
  const row = sh.getRange(hitRow + 2, 1, 1, header.length).getValues()[0];
  return rowToObj_(header, row);
}

function upsertUser_(body) {
  const lineUserId = reqBody_(body, 'line_user_id');
  const nickName = (body.nick_name || "").toString().trim();
  const name = reqBody_(body, 'name');
  const kana = reqBody_(body, 'kana');
  const birthday = (body.birthday || "").toString().trim();
  const gender = reqBody_(body, 'gender');

  // allowed_gender チェック
  const allowed = getAllowedGenders_();
  if (!isGenderAllowed_(gender, allowed)) {
    const err = new Error("GENDER_NOT_ALLOWED");
    err.allowed_genders = allowed;
    throw err;
  }

  // phone は数字だけにして文字列化（先頭0維持）
  const phoneRaw = body.phone;
  const phone = String(phoneRaw ?? '').replace(/[^0-9]/g, '');
  const email = (body.email || '').trim();

  const now = new Date();
  const sh = sh_(SHEET_USERS);
  if (!sh) throw new Error(`Sheet not found: ${SHEET_USERS}`);

  const lastRow = sh.getLastRow();

  // header（1行だけ）
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const idx = indexMap_(header);
  if (idx.line_user_id === undefined) throw new Error("USERS_MISSING_COLUMN_line_user_id");

  // phone列の表示形式をテキストに固定（予防）
  ensureUsersPhoneTextColumn_(sh, idx);

  const lock = LockService.getScriptLock();
  lock.tryLock(15000);

  try {
    // 既存行検索：line_user_id列だけ読む（全件 rowToObj しない）
    let targetRow = -1; // 1-based
    if (lastRow >= 2) {
      const n = lastRow - 1;
      const colLine = sh.getRange(2, idx.line_user_id + 1, n, 1).getValues();
      for (let i = 0; i < n; i++) {
        if (String(colLine[i][0]) === String(lineUserId)) {
          targetRow = i + 2;
          break;
        }
      }
    }

    // 書きたい値（列が無い場合もあるので idx を見て反映）
    const record = {
      line_user_id: lineUserId,
      nick_name: nickName,
      name,
      kana,
      gender,
      birthday,
      phone,  // 書き込み時は ' を付ける
      email,
      is_active: true,
      updated_at: now,
    };

    if (targetRow === -1) {
      // ===== insert =====
      record.created_at = now;

      const newRow = header.map(h => {
        const key = String(h || '').trim();
        switch (key) {
          case 'line_user_id': return record.line_user_id;
          case 'nick_name':    return record.nick_name;
          case 'name':         return record.name;
          case 'kana':         return record.kana;
          case 'gender':       return record.gender;
          case 'birthday':     return record.birthday;
          case 'phone':        return record.phone ? `'${record.phone}` : '';
          case 'email':        return record.email;
          case 'created_at':   return record.created_at;
          case 'updated_at':   return record.updated_at;
          case 'is_active':    return record.is_active;
          default:             return '';
        }
      });

      sh.appendRow(newRow);

      // phoneセルもテキスト書式
      if (idx.phone !== undefined) {
        const r = sh.getLastRow();
        sh.getRange(r, idx.phone + 1).setNumberFormat('@');
      }

    } else {
      // ===== update =====
      // 対象行1行だけ読む（ここはOK：1行だけ）
      const row = sh.getRange(targetRow, 1, 1, header.length).getValues()[0];

      // created_at は維持
      // 各列が存在する場合だけ上書き
      if (idx.nick_name !== undefined) row[idx.nick_name] = record.nick_name;
      if (idx.name !== undefined)      row[idx.name] = record.name;
      if (idx.kana !== undefined)      row[idx.kana] = record.kana;
      if (idx.gender !== undefined)    row[idx.gender] = record.gender;
      if (idx.birthday !== undefined)  row[idx.birthday] = record.birthday;
      if (idx.email !== undefined)     row[idx.email] = record.email;
      if (idx.is_active !== undefined) row[idx.is_active] = record.is_active;
      if (idx.updated_at !== undefined)row[idx.updated_at] = record.updated_at;

      // phone は「'付き + テキスト書式」で確定
      if (idx.phone !== undefined) {
        row[idx.phone] = record.phone ? `'${record.phone}` : '';
      }

      // 1行だけ一括書き戻し
      sh.getRange(targetRow, 1, 1, header.length).setValues([row]);

      // phone列の表示形式を念押し
      if (idx.phone !== undefined) {
        sh.getRange(targetRow, idx.phone + 1).setNumberFormat('@');
      }
    }

    return getUserByLineId_(lineUserId);

  } finally {
    lock.releaseLock();
  }
}

/**
 * USERSシートの phone 列を「プレーンテキスト」に固定する
 * - 列全体を @ にする（これが一番確実）
 */
function ensureUsersPhoneTextColumn_(usersSheet, idx) {
  if (!usersSheet) return;
  if (!idx || idx.phone === undefined) return;

  const col = idx.phone + 1; // 1-based
  usersSheet.getRange(1, col, usersSheet.getMaxRows(), 1).setNumberFormat('@');
}

/** =========================
 *  Core: Reservations
 * ========================= */
function listReservationsByUser_(lineUserId, status) {
  const sh = sh_(SHEET_RESERVATIONS);
  if (!sh) throw new Error(`Sheet not found: ${SHEET_RESERVATIONS}`);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const idx = indexMap_(header);

  requiredCols_(idx, ['line_user_id', 'status', 'reserved_start', 'reserved_end', 'reservation_id']);

  const n = lastRow - 1;

  // 必要列だけ
  const colLine   = sh.getRange(2, idx.line_user_id + 1, n, 1).getValues();
  const colStatus = sh.getRange(2, idx.status + 1, n, 1).getValues();
  const colStart  = sh.getRange(2, idx.reserved_start + 1, n, 1).getValues();
  const colEnd    = sh.getRange(2, idx.reserved_end + 1, n, 1).getValues();
  const colRid    = sh.getRange(2, idx.reservation_id + 1, n, 1).getValues();

  // プラン名（スナップショット優先）
  const hasPlanNames = idx.plan_names_snapshot !== undefined;
  const colPlanNames = hasPlanNames ? sh.getRange(2, idx.plan_names_snapshot + 1, n, 1).getValues() : null;

  // ★ 追加：所要時間・金額（スナップショット）
  const hasDurSnap = idx.duration_min_snapshot !== undefined;
  const colDurSnap = hasDurSnap ? sh.getRange(2, idx.duration_min_snapshot + 1, n, 1).getValues() : null;

  const hasPriceSnap = idx.price_snapshot !== undefined;
  const colPriceSnap = hasPriceSnap ? sh.getRange(2, idx.price_snapshot + 1, n, 1).getValues() : null;

  // （任意）cancel_token も返したいなら
  const hasCancelToken = idx.cancel_token !== undefined;
  const colCancelToken = hasCancelToken ? sh.getRange(2, idx.cancel_token + 1, n, 1).getValues() : null;

  const out = [];
  for (let i = 0; i < n; i++) {
    if (String(colLine[i][0]) !== String(lineUserId)) continue;
    if (status && String(colStatus[i][0]) !== String(status)) continue;

    const start = colStart[i][0];
    const end   = colEnd[i][0];

    // フロントで使う“表示プラン名”は snapshot を優先
    const planNames = hasPlanNames ? String(colPlanNames[i][0] || '').trim() : '';

    // ★ NaN回避：数値化できない時は null にする
    const durRaw = hasDurSnap ? colDurSnap[i][0] : null;
    const priceRaw = hasPriceSnap ? colPriceSnap[i][0] : null;

    const dur = (durRaw === null || durRaw === '' || durRaw === undefined) ? null : Number(durRaw);
    const price = (priceRaw === null || priceRaw === '' || priceRaw === undefined) ? null : Number(priceRaw);

    out.push({
      reservation_id: String(colRid[i][0] || ''),
      status: String(colStatus[i][0] || ''),
      reserved_start: start,
      reserved_end: end,

      // ★ ここが重要
      plan_name: planNames,        // 既存UIが plan_name 参照ならここに入れる
      duration_min: Number.isFinite(dur) ? dur : null,
      price: Number.isFinite(price) ? price : null,

      cancel_token: hasCancelToken ? String(colCancelToken[i][0] || '') : ''
    });
  }

  out.sort((a, b) => new Date(a.reserved_start) - new Date(b.reserved_start));
  return out;
}

function createReservation_(body) {
  const lineUserId = reqBody_(body, 'line_user_id');

  // ★追加：ペナルティならここで予約不可
  assertNotPenalized_(lineUserId);

  // ✅追加：note（任意）
  const note = (body.note ?? "").toString().trim().slice(0, 500);

  // ★複数対応：plan_ids（配列）優先、無ければ旧plan_id互換
  let planIds = body.plan_ids;
  if (Array.isArray(planIds) && planIds.length > 0) {
    planIds = planIds.map(x => String(x).trim()).filter(Boolean);
  } else {
    const legacyPlanId = (body.plan_id ?? '').toString().trim();
    if (!legacyPlanId) throw new Error('MISSING_BODY_plan_id'); // 旧互換のため残す
    planIds = [legacyPlanId];
  }

  const startIso = reqBody_(body, 'start_at');
  const startAt = new Date(startIso);
  startAt.setSeconds(0, 0);
  
  if (isNaN(startAt.getTime())) throw new Error('INVALID_START_AT');

  // ユーザー登録必須
  const user = getUserByLineId_(lineUserId);
  if (!user) throw new Error('USER_NOT_REGISTERED');

  // ★プランをまとめて読み込んで合算（orderで並べ替え）
  let plans = planIds.map(pid => {
    const p = getPlanById_(pid);
    if (!p || !p.is_active) throw new Error('PLAN_NOT_FOUND_OR_INACTIVE');
    return p;
  });

  console.log("[reserve] input planIds:", JSON.stringify(planIds));
  console.log("[reserve] loaded plans:", JSON.stringify(plans.map(p => ({
    plan_id: p.plan_id, order: p.order, name: p.plan_name
  }))));

  // ★ order 昇順 → plan_id 昇順（同順時安定化）
  plans.sort((a, b) => (Number(a.order || 999999) - Number(b.order || 999999)) || String(a.plan_id).localeCompare(String(b.plan_id)));

  console.log("[reserve] sorted plans:", JSON.stringify(plans.map(p => ({
    plan_id: p.plan_id, order: p.order, name: p.plan_name
  }))));

  // ★ 並べ替え後の planIds も作り直す（snapshot用）
  planIds = plans.map(p => String(p.plan_id));

  const totalDuration = plans.reduce((a, p) => a + Number(p.duration_min), 0);
  const totalPrice = plans.reduce((a, p) => a + Number(p.price), 0);
  const priceStr = Number(totalPrice).toLocaleString("ja-JP");

  // ★ order昇順の名前で連結
  const planNames = plans.map(p => p.plan_name).join(' + ');

  const endAt = new Date(startAt.getTime() + totalDuration * 60 * 1000);
  endAt.setSeconds(0, 0);

  // 同時予約なし：重複チェックはロック下で
  const lock = LockService.getScriptLock();
  lock.tryLock(15000);

  try {
    if (hasConflict_(startAt, endAt)) throw new Error('TIME_SLOT_TAKEN');

    const sh = sh_(SHEET_RESERVATIONS);
    const values = sh.getDataRange().getValues();
    const header = values[0].map(String);
    const idx = indexMap_(header);

    const now = new Date();
    const reservationId = genReservationId_();
    const cancelToken = genToken_();

    // ★スナップショット（マスター更新の影響を受けない）
    const rowObj = {
      reservation_id: reservationId,
      line_user_id: lineUserId,
      status: 'CONFIRMED',
      reserved_start: startAt,
      reserved_end: endAt,

      // 互換：単数列が残ってる場合に備える（最初のプランを入れておく）
      plan_id: planIds[0],

      // スナップショット（推奨列）
      plan_ids_snapshot: planIds.join(','),
      plan_names_snapshot: planNames,
      duration_min_snapshot: totalDuration,
      price_snapshot: totalPrice,
      note: note, // ✅追加
      created_at: now,
      canceled_at: '',
      cancel_token: cancelToken,
    };

    // ヘッダーに存在する列だけ書く（列追加前でも落ちないように）
    const row = header.map(h => (h in rowObj ? rowObj[h] : ''));
    sh.appendRow(row);

    const adminLineUserId = getAdminLineUserId_();
    const isAdminSelfReserve = adminLineUserId && String(lineUserId).trim() === adminLineUserId;

    // ★ 管理者メール送信（CONFIGに admin_emails がある時だけ）
    // ※ただし admin本人の予約なら送らない
    if (!isAdminSelfReserve) {
      try {
        const reservation = {
          reservation_id: reservationId,
          reserved_start: startAt,
          reserved_end: endAt,
        };

        // note も渡す（下の「3)」も反映する場合）
        sendAdminMailOnReserve_(reservation, user, planNames, priceStr, note);
      } catch (mailErr) {
        console.error("admin mail failed:", mailErr);
      }
    } else {
      console.log("[reserve] skip notifications (admin self reserve)");
    }

    // ★ 予約確定をユーザーへPush通知（失敗しても予約は成功扱い）
    // ※ただし admin本人の予約なら送らない
    if (!isAdminSelfReserve) {
      try {
        const tz = "Asia/Tokyo";
        const startStr = Utilities.formatDate(startAt, tz, "yyyy/MM/dd HH:mm");
        const endStr   = Utilities.formatDate(endAt, tz, "HH:mm");

        pushLineMessage_(lineUserId,
          "✅ 予約が確定しました\n" +
          `日時：${startStr} - ${endStr}\n` +
          `プラン：${planNames}\n` +
          `料金：${priceStr}円\n` +
          `予約ID：${reservationId}`
        );
      } catch (pushErr) {
        console.error("push on reserve failed:", pushErr);
      }
    }

    return {
      reservation_id: reservationId,
      cancel_token: cancelToken,
      reserved_start: startAt,
      reserved_end: endAt,
      plan_names: planNames,
      duration_min: totalDuration,
      price: totalPrice,
      plan_ids: planIds,
    };
  } finally {
    lock.releaseLock();
  }
}

function cancelReservation_(body) {
  const cancelToken = reqBody_(body, 'cancel_token');

  const lock = LockService.getScriptLock();
  lock.tryLock(15000);

  try {
    const sh = sh_(SHEET_RESERVATIONS);
    const values = sh.getDataRange().getValues();
    const header = values[0].map(String);
    const idx = indexMap_(header);

    // ★ idx を作った後に必須列チェック
    requiredCols_(idx, ['cancel_token', 'status', 'canceled_at', 'reservation_id', 'line_user_id', 'reserved_start', 'reserved_end']);


    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      if (String(row[idx.cancel_token]) !== String(cancelToken)) continue;

      const status = String(row[idx.status]);
      if (status === 'CANCELED') {
        return { reservation_id: String(row[idx.reservation_id]), status: 'CANCELED' };
      }

      const rowNo = r + 1;

      // 対象予約の情報を先に取得（push用）
      const lineUserId = (idx.line_user_id !== undefined) ? String(row[idx.line_user_id] || '').trim() : '';
      const reservedStart = (idx.reserved_start !== undefined) ? coerceToDate_(row[idx.reserved_start]) : null;
      const reservedEnd   = (idx.reserved_end !== undefined) ? coerceToDate_(row[idx.reserved_end]) : null;
      const tz = "Asia/Tokyo";
      const todayKey = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
      const resvDayKey = reservedStart ? Utilities.formatDate(reservedStart, tz, "yyyy-MM-dd") : "";

      if (resvDayKey && resvDayKey === todayKey) {
        const phone = getAdminPhone_();
        const err = new Error("SAME_DAY_CANCEL_NOT_ALLOWED");
        err.admin_phone = phone;
        throw err;
      }

      const planNames =
        (idx.plan_names_snapshot !== undefined && String(row[idx.plan_names_snapshot] || '').trim())
          ? String(row[idx.plan_names_snapshot]).trim()
          : (idx.plan_name_snapshot !== undefined ? String(row[idx.plan_name_snapshot] || '').trim() : '');

      const reservationId = String(row[idx.reservation_id] || '');

      // キャンセル更新
      sh.getRange(rowNo, idx.status + 1).setValue('CANCELED');
      sh.getRange(rowNo, idx.canceled_at + 1).setValue(new Date());

      // ★ ユーザーへPush通知（失敗してもキャンセルは成功扱い）
      try {
        if (lineUserId) {
          const tz = "Asia/Tokyo";
          const startStr = reservedStart ? Utilities.formatDate(reservedStart, tz, "yyyy/MM/dd HH:mm") : "";
          const endStr   = reservedEnd ? Utilities.formatDate(reservedEnd, tz, "HH:mm") : "";

          pushLineMessage_(lineUserId,
            "✅ 予約をキャンセルしました\n" +
            (startStr ? `日時：${startStr}${endStr ? " - " + endStr : ""}\n` : "") +
            (planNames ? `プラン：${planNames}\n` : "") +
            (reservationId ? `予約ID：${reservationId}` : "")
          );
        }
      } catch (pushErr) {
        console.error("push on cancel failed:", pushErr);
      }

      // 管理者へメール送信
      try {
        // user を取得（USERSから）
        const user = getUserByLineId_(lineUserId) || {};
        // 料金スナップショット
        const priceSnap = (idx.price_snapshot !== undefined) ? Number(row[idx.price_snapshot] || 0) : 0;
        const priceStr = Number(priceSnap).toLocaleString("ja-JP");

        const reservation = {
          reservation_id: reservationId,
          reserved_start: reservedStart,
          reserved_end: reservedEnd,
        };

        sendAdminMailOnCancel_(reservation, user, planNames, priceStr);
      } catch (mailErr) {
        console.error("admin cancel mail failed:", mailErr);
      }

      return { reservation_id: reservationId, status: 'CANCELED' };
    }
    throw new Error('CANCEL_TOKEN_NOT_FOUND');
  } finally {
    lock.releaseLock();
  }
}

/** =========================
 *  Conflict Check (minimal)
 * ========================= */
function hasConflict_(startAt, endAt) {
  const sh = sh_(SHEET_RESERVATIONS);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return false;

  const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
  const idx = indexMap_(header);
  requiredCols_(idx, ['status','reserved_start','reserved_end']);

  const n = lastRow - 1;
  const colStatus = sh.getRange(2, idx.status+1, n, 1).getValues();
  const colStart  = sh.getRange(2, idx.reserved_start+1, n, 1).getValues();
  const colEnd    = sh.getRange(2, idx.reserved_end+1, n, 1).getValues();

  const a0 = startAt.getTime();
  const a1 = endAt.getTime();

  for (let i=0; i<n; i++){
    if (String(colStatus[i][0]).trim() !== 'CONFIRMED') continue;
    const s = coerceToDate_(colStart[i][0]);
    const e = coerceToDate_(colEnd[i][0]);
    if (!s || !e) continue;

    if (a0 < e.getTime() && a1 > s.getTime()) return true;
  }
  return false;
}

/** =========================
 *  Helpers
 * ========================= */
let SS_CACHE_ = null;
function ss_(){
  if (SS_CACHE_) return SS_CACHE_;
  return (SS_CACHE_ = SpreadsheetApp.openById(SPREADSHEET_ID));
}

const SHEETS_ = {};
function sh_(name){
  const cached = SHEETS_[name];
  if (cached) return cached;                 // シートオブジェクトは truthy
  const sh = ss_().getSheetByName(name);
  if (sh) SHEETS_[name] = sh;                // null はキャッシュしない
  return sh;
}
function shFresh_(name){
  const sh = ss_().getSheetByName(name);
  if (sh) SHEETS_[name] = sh;
  return sh;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function reqParam_(e, key) {
  const v = (e.parameter[key] || '').trim();
  if (!v) throw new Error(`MISSING_PARAM_${key}`);
  return v;
}

/** =========================
 *  Cache helpers
 * ========================= */
function cacheVer_(){
  const p = PropertiesService.getScriptProperties();
  return p.getProperty('CACHE_VER') || '1';
}
function bumpCacheVer_(){
  const p = PropertiesService.getScriptProperties();
  const v = Number(p.getProperty('CACHE_VER') || '1') + 1;
  p.setProperty('CACHE_VER', String(v));
}

function cachedJson_(key, ttlSec, computeObjFn){
  const cache = CacheService.getScriptCache();
  const hit = cache.get(key);
  if (hit) {
    return ContentService.createTextOutput(hit)
      .setMimeType(ContentService.MimeType.JSON);
  }
  const obj = computeObjFn();
  const text = JSON.stringify(obj);
  cache.put(key, text, ttlSec);
  return ContentService.createTextOutput(text)
    .setMimeType(ContentService.MimeType.JSON);
}


function reqBody_(body, key) {
  const v = (body[key] ?? '').toString().trim();
  if (!v) throw new Error(`MISSING_BODY_${key}`);
  return v;
}

function indexMap_(header) {
  const m = {};
  header.forEach((h, i) => {
    const key = String(h ?? "")
      .trim()
      .replace(/[（(].*$/, "")   // 括弧注釈を削る：descriptions（説明）→descriptions
      .trim()
      .toLowerCase();           // 大文字小文字を吸収
    if (key) m[key] = i;
  });
  return m;
}

function rowToObj_(header, row) {
  const obj = {};
  header.forEach((h, i) => obj[h] = row[i]);
  return obj;
}

function genReservationId_() {
  const d = new Date();
  const y = d.getFullYear();
  const mo = String(d.getMonth() + 1).padStart(2, '0');
  const da = String(d.getDate()).padStart(2, '0');
  const hh = String(d.getHours()).padStart(2, '0');
  const mi = String(d.getMinutes()).padStart(2, '0');
  const ss = String(d.getSeconds()).padStart(2, '0');
  const rand = Math.floor(Math.random() * 10000).toString().padStart(4, '0');
  return `R-${y}${mo}${da}-${hh}${mi}${ss}-${rand}`;
}

function genToken_() {
  // 簡易UUID風
  return Utilities.getUuid();
}


/**
 * USERS.penalty_flg が TRUE なら予約不可
 * - admin_phone は CONFIG.admin_phone を返す（既存 getAdminPhone_ を使用）
 * - doPost catch が err.admin_phone を返す仕様に合わせる
 */
function assertNotPenalized_(lineUserId) {
  const user = getUserByLineId_(lineUserId); // 既存（USERSを読む）
  if (!user) return; // 未登録は既存仕様（USER_NOT_REGISTERED）で落ちるのでここでは何もしない

  const penalty = String(user.penalty_flg ?? "").toUpperCase() === "TRUE";
  if (penalty) {
    const err = new Error("PENALTY");
    err.admin_phone = getAdminPhone_(); // 既存（CONFIGから取得）
    throw err;
  }
}


/** =========================
 *  Availability (SLOTS-based)
 *  - date: "YYYY-MM-DD"
 *  - plan_id: "P001"
 *  Returns: { available: [ISO...], slot_source_hint: string }
 * ========================= */
function getAvailability_(dateYmd, planId) {
  const plan = getPlanById_(planId);
  if (!plan || !plan.is_active) throw new Error('PLAN_NOT_FOUND_OR_INACTIVE');

  const durationMin = Number(plan.duration_min);
  const fromStart = parseYmdAsLocalDate_(dateYmd);
  const rangeEnd = new Date(fromStart.getTime() + 24 * 60 * 60 * 1000);

  // materials（1日分）
  const mats = getAvailabilityRangeMaterialsByDuration_(dateYmd, 1, planId, durationMin);

  // mats を元にサーバで available を作る（フロントで作るなら返さなくてもOK）
  const granMin = mats.granularity_min;
  const requiredMs = durationMin * 60 * 1000;

  const windows = mats.windows_by_date[dateYmd] || [];
  const busy = mats.busy_by_date[dateYmd] || [];
  const minStartMin = mats.min_start_min_by_date[dateYmd]; // 今日だけ数値、それ以外 null

  const available = [];

  // [startMin,endMin) の window 内で gran ごとに候補を作り、busy と重なるものを除外
  for (const w of windows) {
    const w0 = w[0], w1 = w[1];
    for (let tMin = ceilMinToGran_(w0, granMin); (tMin * 60000 + requiredMs) <= w1 * 60000; tMin += granMin) {
      if (minStartMin !== null && tMin <= minStartMin) continue;
      const endMin = tMin + Math.round(requiredMs / 60000);

      if (isOverlappingMinutes_(tMin, endMin, busy)) continue;

      const dt = new Date(fromStart.getFullYear(), fromStart.getMonth(), fromStart.getDate(), 0, 0, 0, 0);
      dt.setMinutes(tMin);
      available.push(toIsoWithOffset_(dt));
    }
  }

  return {
    available: Array.from(new Set(available)).sort(),
    slot_source_hint: 'materials（windows/busy）から生成'
  };
}

// minutesの重なり判定（busyは [[s,e],...]）
function isOverlappingMinutes_(s, e, busyIntervals){
  for (const b of busyIntervals) {
    const b0 = b[0], b1 = b[1];
    // 半開区間 [s,e) と [b0,b1)
    if (s < b1 && e > b0) return true;
  }
  return false;
}

function ceilMinToGran_(min, gran){
  return Math.ceil(min / gran) * gran;
}

/**
 * BLACKOUTS: 受付不可期間
 * シートが無い場合は制限なし（空配列）
 *
 * 想定ヘッダー（最低限）：from, to, all_day, is_active
 * - all_day TRUE or fromが日付だけ → 当日0:00〜翌日0:00
 */
function listBlackoutsOverlapping_(dayStart, dayEnd) {
  const sh = sh_('BLACKOUTS');
  if (!sh) return []; // ← シート無ければ制限なし

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  // ヘッダーの括弧注釈を許容
  const header = values[0].map(v => String(v).trim().replace(/[（(].*$/, '').trim());
  const idx = indexMap_(header);

  if (idx.from === undefined) throw new Error('BLACKOUTS_MISSING_COLUMN_from');

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];

    const isActive = (idx.is_active === undefined)
      ? true
      : String(row[idx.is_active] ?? '').toUpperCase() !== 'FALSE';
    if (!isActive) continue;

    const fromRaw = row[idx.from];
    if (!fromRaw) continue;

    const toRaw = (idx.to !== undefined) ? row[idx.to] : null;
    const allDayRaw = (idx.all_day !== undefined) ? row[idx.all_day] : null;

    const norm = normalizeBlackout_(fromRaw, toRaw, allDayRaw);
    if (!norm) continue;

    // dayStart-dayEnd と重なるものだけ返す
    if (norm.from < dayEnd && norm.to > dayStart) out.push(norm);
  }

  out.sort((a, b) => a.from.getTime() - b.from.getTime());
  return out;
}

function isInBlackout_(startAt, endAt, blackouts) {
  if (!blackouts || blackouts.length === 0) return false;
  for (const b of blackouts) {
    // overlap: [startAt,endAt) と [b.from,b.to)
    if (startAt < b.to && endAt > b.from) return true;
  }
  return false;
}

function normalizeBlackout_(fromRaw, toRaw, allDayRaw) {
  const fromDate = coerceToDate_(fromRaw);
  if (!fromDate) return null;

  const allDay =
    (allDayRaw !== null && allDayRaw !== undefined && String(allDayRaw).trim() !== '')
      ? (String(allDayRaw).toUpperCase() === 'TRUE')
      : isDateOnlyInput_(fromRaw);

  if (allDay) {
    const d0 = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate(), 0, 0, 0, 0);
    const d1 = new Date(d0.getTime() + 24 * 60 * 60 * 1000);
    return { from: d0, to: d1 };
  }

  const toDate = coerceToDate_(toRaw);
  if (!toDate) throw new Error('BLACKOUTS_TO_REQUIRED_FOR_RANGE');
  if (toDate <= fromDate) throw new Error('BLACKOUTS_INVALID_RANGE');

  return { from: fromDate, to: toDate };
}

function isDateOnlyInput_(v) {
  if (!v && v !== 0) return false;

  if (Object.prototype.toString.call(v) === '[object Date]') {
    return v.getHours() === 0 && v.getMinutes() === 0 && v.getSeconds() === 0;
  }
  if (typeof v === 'string') {
    const s = v.trim();
    return /^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/.test(s);
  }
  return false;
}

function getGranularityMinutes_() {
  // CONFIGシート A列=キー, B列=値 でもOKにしたいなら拡張できるけど、
  // まずは A1: granularity_min / B1: 30 で簡単運用
  const sh = sh_('CONFIG');
  if (!sh) return 30;

  const key = String(sh.getRange('A1').getValue() || '').trim();
  const val = sh.getRange('B1').getValue();

  if (key !== 'granularity_min') return 30;
  const n = Number(val);
  return Number.isFinite(n) && n > 0 ? n : 30;
}

function listOpenWindowsForDate_(dayStart, dayEnd) {
  // 1) SLOTS がある＆当日データがあるならそれを使う（手動上書き用）
  const sh = sh_('SLOTS');
  if (sh) {
    const values = sh.getDataRange().getValues();
    if (values.length >= 2) {
      const headerRaw = values[0].map(v => String(v));
      const header = headerRaw.map(h => normalizeHeader_(h));
      const idx = indexMap_(header);

      const hasCols = idx.slot_start !== undefined && idx.slot_end !== undefined && idx.is_open !== undefined;
      if (hasCols) {
        const out = [];
        for (let r = 1; r < values.length; r++) {
          const row = values[r];
          const isOpen = String(row[idx.is_open] ?? '').toUpperCase() === 'TRUE';
          if (!isOpen) continue;

          const from = coerceToDate_(row[idx.slot_start]);
          const to = coerceToDate_(row[idx.slot_end]);
          if (!from || !to) continue;
          if (to <= from) continue;

          if (from < dayEnd && to > dayStart) out.push({ from, to });
        }
        out.sort((a, b) => a.from.getTime() - b.from.getTime());

        // 当日分のSLOTSが1件でもあればSLOTSを優先
        if (out.length > 0) return out;
      }
    }
  }

  // 2) SLOTSが無い/当日データが無い → CONFIGのデフォルト営業時間を使う
  if (isClosedByWeekday_(dayStart)) return []; // 定休日

  const bh = getBusinessHours_();
  const from = new Date(dayStart.getFullYear(), dayStart.getMonth(), dayStart.getDate(), bh.oh, bh.om, 0, 0);
  const to = new Date(dayStart.getFullYear(), dayStart.getMonth(), dayStart.getDate(), bh.ch, bh.cm, 0, 0);

  if (to <= from) throw new Error('CONFIG_INVALID_business_hours_range');

  // 日境界でクリップ
  const wFrom = new Date(Math.max(from.getTime(), dayStart.getTime()));
  const wTo = new Date(Math.min(to.getTime(), dayEnd.getTime()));
  if (wTo <= wFrom) return [];

  return [{ from: wFrom, to: wTo }];
}

function normalizeHeader_(h) {
  // "is_open (TRUE/FALSE)" → "is_open"
  // "slot_start（日付）" → "slot_start"
  return String(h)
    .trim()
    .replace(/[（(].*$/, '') // 括弧以降削除
    .trim();
}

function ceilToGranFromAnchor_(t, anchor, granMin) {
  // t, anchor: Date
  const step = granMin * 60 * 1000;
  const dt = t.getTime() - anchor.getTime();
  const k = Math.ceil(dt / step);
  return new Date(anchor.getTime() + k * step);
}

/**
 * SLOTSから指定日の open slots を Date配列で返す
 */
function listOpenSlotsForDate_(dayStart, dayEnd) {
  const sh = sh_('SLOTS');
  if (!sh) throw new Error('SLOTS_SHEET_NOT_FOUND');

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const header = values[0].map(String);
  const idx = indexMap_(header);

  if (idx.slot_start === undefined) throw new Error('SLOTS_MISSING_COLUMN_slot_start');
  if (idx.is_open === undefined) throw new Error('SLOTS_MISSING_COLUMN_is_open');

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const isOpen = String(row[idx.is_open] ?? '').toUpperCase() === 'TRUE';
    if (!isOpen) continue;

    const dt = coerceToDate_(row[idx.slot_start]);
    if (!dt) continue;

    if (dt >= dayStart && dt < dayEnd) out.push(dt);
  }

  // 時刻昇順
  out.sort((a, b) => a.getTime() - b.getTime());
  return out;
}

/**
 * 粒度（分）をSLOTSの最小差分から推定
 * 例：10:00,10:30,11:00 => 30
 */
function inferGranularityMinutes_(slots) {
  if (!slots || slots.length < 2) return null;
  let minDiffMs = null;
  for (let i = 1; i < slots.length; i++) {
    const diff = slots[i].getTime() - slots[i - 1].getTime();
    if (diff <= 0) continue;
    if (minDiffMs === null || diff < minDiffMs) minDiffMs = diff;
  }
  if (minDiffMs === null) return null;
  const min = Math.round(minDiffMs / 60000);

  // 現実的な粒度に丸め（1分など誤差のときに暴れないように）
  const allowed = [5, 10, 15, 20, 30, 60];
  let best = allowed[0];
  let bestErr = Math.abs(min - best);
  for (const a of allowed) {
    const err = Math.abs(min - a);
    if (err < bestErr) { best = a; bestErr = err; }
  }
  return best;
}

/**
 * "YYYY-MM-DD" を local 00:00 の Date にする
 */
function parseYmdAsLocalDate_(ymd) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(ymd));
  if (!m) throw new Error('INVALID_DATE_FORMAT');
  const y = Number(m[1]);
  const mo = Number(m[2]) - 1;
  const d = Number(m[3]);
  return new Date(y, mo, d, 0, 0, 0, 0);
}

/**
 * Date/文字列/数値(シリアル)などをDateに寄せる
 */
function coerceToDate_(v) {
  if (!v && v !== 0) return null;

  if (Object.prototype.toString.call(v) === '[object Date]') {
    if (isNaN(v.getTime())) return null;
    return v;
  }

  // スプレッドシートの日時は基本Dateで来るが、文字列の場合も許容
  if (typeof v === 'string') {
    const d = new Date(v);
    if (!isNaN(d.getTime())) return d;
    return null;
  }

  // 万が一の数値（シートのシリアル）: Dateに変換（Google Sheets serial）
  if (typeof v === 'number') {
    // Sheets serial: 1899-12-30起点
    const ms = (v - 25569) * 86400 * 1000;
    const d = new Date(ms);
    return isNaN(d.getTime()) ? null : d;
  }

  return null;
}

/**
 * タイムゾーンオフセット付きISO文字列にする（例: 2026-02-22T10:00:00+09:00）
 */
function toIsoWithOffset_(date) {
  const d = new Date(date.getTime());
  const pad = (n) => String(n).padStart(2, '0');

  const y = d.getFullYear();
  const m = pad(d.getMonth() + 1);
  const da = pad(d.getDate());
  const hh = pad(d.getHours());
  const mi = pad(d.getMinutes());
  const ss = pad(d.getSeconds());

  const offsetMin = -d.getTimezoneOffset(); // JSTなら +540
  const sign = offsetMin >= 0 ? '+' : '-';
  const abs = Math.abs(offsetMin);
  const oh = pad(Math.floor(abs / 60));
  const om = pad(abs % 60);

  return `${y}-${m}-${da}T${hh}:${mi}:${ss}${sign}${oh}:${om}`;
}


/**
 * 週レンジに重なる CONFIRMED 予約を一括取得
 */
function listConfirmedReservationsOverlapping_(rangeStart, rangeEnd) {
  const sh = sh_(SHEET_RESERVATIONS);
  if (!sh) throw new Error(`Sheet not found: ${SHEET_RESERVATIONS}`);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const idx = indexMap_(header);

  // 必須列
  const requiredCols = ['status', 'reserved_start', 'reserved_end'];
  for (const c of requiredCols) {
    if (idx[c] === undefined) throw new Error(`RESERVATIONS_MISSING_COLUMN_${c}`);
  }

  const n = lastRow - 1;

  // 必要列だけ取得
  const colStatus = sh.getRange(2, idx.status + 1, n, 1).getValues();
  const colStart  = sh.getRange(2, idx.reserved_start + 1, n, 1).getValues();
  const colEnd    = sh.getRange(2, idx.reserved_end + 1, n, 1).getValues();

  const out = [];
  for (let i = 0; i < n; i++) {
    if (String(colStatus[i][0] || '').trim() !== 'CONFIRMED') continue;

    const s = coerceToDate_(colStart[i][0]);
    const e = coerceToDate_(colEnd[i][0]);
    if (!s || !e) continue;

    if (s < rangeEnd && e > rangeStart) out.push({ s, e });
  }

  // （任意）少し速くしたいなら startでソート
  out.sort((a, b) => a.s.getTime() - b.s.getTime());
  return out;
}

/**
 * 予約リスト内で重複があるか（高速）
 */
function hasConflictInList_(startAt, endAt, confirmedList) {
  const a0 = startAt.getTime();
  const a1 = endAt.getTime();

  for (const r of confirmedList) {
    const b0 = r.s.getTime();
    const b1 = r.e.getTime();

    // ✅ 半開区間 [a0,a1) と [b0,b1) の重なり判定
    // 例) 9:00-10:00 と 10:00-10:30 はOK（重ならない）
    // 例) 9:30-10:30 と 10:00-10:30 はNG（重なる）
    if (a0 < b1 && a1 > b0) return true;
  }
  return false;
}

/**
 * Date -> "YYYY-MM-DD"
 */
function formatYmd_(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const da = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${da}`;
}


let CONFIG_CACHE_ = null;
function getConfigMap_() {
  if (CONFIG_CACHE_) return CONFIG_CACHE_;
  const sh = sh_('CONFIG');
  if (!sh) return (CONFIG_CACHE_ = {});
  const values = sh.getDataRange().getValues();
  const map = {};
  for (const [k,v] of values) {
    const key = String(k||'').trim();
    if (key) map[key] = String(v ?? '').trim();
  }
  return (CONFIG_CACHE_ = map);
}

function getBusinessHours_() {
  const cfg = getConfigMap_();
  const openStr = (cfg.business_open || '09:00').toString().replaceAll('：', ':').trim();
  const closeStr = (cfg.business_close || '18:00').toString().replaceAll('：', ':').trim();

  const [oh, om] = openStr.split(':').map(Number);
  const [ch, cm] = closeStr.split(':').map(Number);

  if (![oh, om, ch, cm].every(Number.isFinite)) throw new Error('CONFIG_INVALID_business_open_close');

  return { openStr, closeStr, oh, om, ch, cm };
}

function getClosedWeekdays_() {
  const cfg = getConfigMap_();
  const s = (cfg.closed_weekdays || '').trim();
  if (!s) return []; // 指定なしなら定休日なし
  return s.split(',').map(x => Number(String(x).trim())).filter(n => Number.isFinite(n));
}

/**
 * 指定日が定休日か
 * dayStartの曜日で判断（0=日..6=土）
 */
function isClosedByWeekday_(dayStart) {
  const closed = getClosedWeekdays_();
  return closed.includes(dayStart.getDay());
}


function getAvailabilityRangeByDuration_(fromYmd, days, planId, durationMinOverride) {
  let durationMin = durationMinOverride;

  if (!durationMin) {
    if (!planId) throw new Error('MISSING_PARAM_plan_id_or_duration_min');
    const plan = getPlanById_(planId);
    if (!plan || !plan.is_active) throw new Error('PLAN_NOT_FOUND_OR_INACTIVE');
    durationMin = Number(plan.duration_min);
  }
  if (!Number.isFinite(durationMin) || durationMin <= 0) throw new Error('INVALID_duration_min');

  const nDays = Number.isFinite(days) && days > 0 && days <= 14 ? Math.floor(days) : 7;

  const fromStart = parseYmdAsLocalDate_(fromYmd);
  const rangeEnd = new Date(fromStart.getTime() + nDays * 24 * 60 * 60 * 1000);

  const granMin = getGranularityMinutes_();
  const requiredMs = durationMin * 60 * 1000;

  const tz = "Asia/Tokyo";
  const now = new Date();

  // ✅ しきい値：現在日時 + granularity_min
  const threshold = new Date(now.getTime() + granMin * 60 * 1000);

  const confirmed = listConfirmedReservationsOverlapping_(fromStart, rangeEnd);
  const blackouts = listBlackoutsOverlapping_(fromStart, rangeEnd);

  const bh = (typeof getBusinessHours_ === 'function') ? getBusinessHours_() : { openStr: '09:00', closeStr: '18:00' };

  const byDate = {};

  for (let i = 0; i < nDays; i++) {
    const dayStart = new Date(fromStart.getTime() + i * 24 * 60 * 60 * 1000);
    const dayEnd   = new Date(dayStart.getTime() + 24 * 60 * 60 * 1000);
    const dateKey  = formatYmd_(dayStart);

    // ✅ その日の終わり(=翌日0:00)が threshold 以下なら「完全に過去日」扱いで空
    if (dayEnd.getTime() <= threshold.getTime()) {
      byDate[dateKey] = [];
      continue;
    }

    const windows = listOpenWindowsForDate_(dayStart, dayEnd);
    const available = [];

    for (const w of windows) {
      const wStart = new Date(Math.max(w.from.getTime(), dayStart.getTime()));
      const wEnd   = new Date(Math.min(w.to.getTime(), dayEnd.getTime()));

      const bh2 = getBusinessHours_();
      const anchor = new Date(dayStart.getFullYear(), dayStart.getMonth(), dayStart.getDate(), bh2.oh, bh2.om, 0, 0);

      for (let t = ceilToGranFromAnchor_(wStart, anchor, granMin).getTime(); t + requiredMs <= wEnd.getTime(); t += granMin * 60 * 1000) {
        const startAt = new Date(t);
        const endAt   = new Date(t + requiredMs);

        // ✅ しきい値以前は不可（過去扱い）
        if (startAt.getTime() <= threshold.getTime()) continue;

        if (isInBlackout_(startAt, endAt, blackouts)) continue;
        if (hasConflictInList_(startAt, endAt, confirmed)) continue;

        available.push(toIsoWithOffset_(startAt));
      }
    }

    byDate[dateKey] = Array.from(new Set(available)).sort();
  }

  return {
    from: formatYmd_(fromStart),
    days: nDays,
    granularity_min: granMin,
    business_open: bh.openStr,
    business_close: bh.closeStr,
    required_duration_min: durationMin,
    slot_source_hint: `所要時間=${durationMin}分で生成（粒度: ${granMin}分）/ BLACKOUTS適用`,
    by_date: byDate
  };
}

function refreshTodayReservations() {
  const ss = ss_();
  const wsRes = sh_(SHEET_RESERVATIONS);
  if (!wsRes) throw new Error(`Sheet not found: ${SHEET_RESERVATIONS}`);

  let wsToday = sh_(SHEET_TODAY);
  if (!wsToday) {
    ss.insertSheet(SHEET_TODAY);
    wsToday = shFresh_(SHEET_TODAY);
  }

  const tz = "Asia/Tokyo";
  wsToday.clear();

  // A1 に本日の日付（フォント20）
  // A1 に本日の日付（フォント20）
  const todayStr = Utilities.formatDate(new Date(), tz, "yyyy/MM/dd");
  wsToday.getRange("A1").setValue(todayStr).setFontSize(20).setFontWeight("bold");
  // 見栄え：1行目を横に広げて中央寄せ（列数は出力列に合わせて I まで）
  wsToday.getRange("A1:I1").merge().setHorizontalAlignment("left").setVerticalAlignment("middle");

  // ヘッダ（2行目）
  wsToday.getRange(2, 1, 1, 9).setValues([[
    "開始", "終了", "顧客名", "電話番号", "プラン", "ステータス", "要望", "予約ID", "LINE_ID"
  ]]);
  wsToday.setFrozenRows(2);

  const values = wsRes.getDataRange().getValues();
  if (values.length < 2) {
    wsToday.getRange(3, 1).setValue("予約データがありません。");
    return;
  }

  const header = values[0].map(v => String(v).trim());
  const idx = indexMap_(header);

  requiredCols_(idx, ['reserved_start', 'reserved_end', 'status', 'line_user_id', 'reservation_id']);

  const hasPlanNames = idx.plan_names_snapshot !== undefined;
  const hasNameSnap  = idx.name_snapshot !== undefined;
  const hasNote      = idx.note !== undefined;

  const userNameByLineId  = buildUserNameMap_();
  const userPhoneByLineId = buildUserPhoneMap_();

  // 今日（JST）範囲
  const now = new Date();
  const y = Number(Utilities.formatDate(now, tz, "yyyy"));
  const m = Number(Utilities.formatDate(now, tz, "MM")) - 1;
  const d = Number(Utilities.formatDate(now, tz, "dd"));
  const start = new Date(y, m, d, 0, 0, 0, 0);
  const end   = new Date(y, m, d + 1, 0, 0, 0, 0);

  const rows = [];

  for (let r = 1; r < values.length; r++) {
    const row = values[r];

    const status = String(row[idx.status] || '').trim();
    if (status === "CANCELED") continue;

    const startAt = coerceToDate_(row[idx.reserved_start]);
    const endAt   = coerceToDate_(row[idx.reserved_end]);
    if (!startAt || !endAt) continue;
    if (startAt < start || startAt >= end) continue;

    const lineId = String(row[idx.line_user_id] || '').trim();
    const rid    = String(row[idx.reservation_id] || '').trim();

    const customer =
      (hasNameSnap && String(row[idx.name_snapshot] || '').trim())
        ? String(row[idx.name_snapshot]).trim()
        : (userNameByLineId[lineId] || lineId || "（不明）");

    // ★電話番号は必ず文字列化して先頭0を保持
    const rawPhone = userPhoneByLineId[lineId] ?? "";
    const phone = rawPhone === "" ? "" : String(rawPhone).trim();

    const planNames =
      (hasPlanNames && String(row[idx.plan_names_snapshot] || '').trim())
        ? String(row[idx.plan_names_snapshot]).trim()
        : "";

    const note =
      (hasNote && String(row[idx.note] || '').trim())
        ? String(row[idx.note]).trim()
        : "";

    rows.push([
      Utilities.formatDate(startAt, tz, "HH:mm"),
      Utilities.formatDate(endAt,   tz, "HH:mm"),
      customer,
      phone,
      planNames,
      status,
      note,
      rid,
      lineId,
      startAt.getTime()
    ]);
  }

  rows.sort((a, b) => a[9] - b[9]);

  if (rows.length === 0) {
    wsToday.getRange(3, 1).setValue("本日の予約はありません。");
    return;
  }

  wsToday.getRange(3, 1, rows.length, 9).setValues(rows.map(r => r.slice(0, 9)));

  // ★電話番号列(D列)を「プレーンテキスト」にして表示上も0落ちを防ぐ
  wsToday.getRange(3, 4, rows.length, 1).setNumberFormat("@");

  // ===== 列幅（px）を固定 =====
  wsToday.setColumnWidths(1, 2, 60);   // A:開始  B:終了
  wsToday.setColumnWidth(3, 120);      // C:顧客名
  wsToday.setColumnWidth(4, 100);      // D:電話番号
  wsToday.setColumnWidth(5, 200);      // E:プラン
  wsToday.setColumnWidth(6, 100);      // F:ステータス
  wsToday.setColumnWidth(7, 500);      // G:要望
  wsToday.setColumnWidth(8, 170);      // H:予約ID
  wsToday.setColumnWidth(9, 260);      // I:LINE_ID
  // ===========================

  wsToday.getRange("K1").setValue(`更新: ${Utilities.formatDate(new Date(), tz, "yyyy/MM/dd HH:mm:ss")}`);
}

function buildUserPhoneMap_() {
  const sh = sh_(SHEET_USERS);
  if (!sh) return {};

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return {};

  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const idx = indexMap_(header);
  if (idx.line_user_id === undefined || idx.phone === undefined) return {};

  const n = lastRow - 1;

  // 必要列だけ取得
  const colLine  = sh.getRange(2, idx.line_user_id + 1, n, 1).getValues();
  const colPhone = sh.getRange(2, idx.phone + 1, n, 1).getValues();

  const map = {};
  for (let i = 0; i < n; i++) {
    const lineId = String(colLine[i][0] || '').trim();
    const phone  = String(colPhone[i][0] || '').trim();
    if (lineId && phone) map[lineId] = phone;
  }
  return map;
}

function getAdminEmails_() {
  const cfg = getConfigMap_();
  const raw = (cfg.admin_emails || cfg.admin_email || "").trim();
  if (!raw) return [];
  return raw.split(",").map(s => s.trim()).filter(Boolean);
}

function sendAdminMailOnReserve_(reservation, user, planNames, priceStr, note) {
  const emails = getAdminEmails_();
  if (emails.length === 0) return;

  const cfg = getConfigMap_();
  const prefix = (cfg.mail_subject_prefix || "[予約]").trim();

  const tz = "Asia/Tokyo";
  const start = Utilities.formatDate(new Date(reservation.reserved_start), tz, "yyyy/MM/dd HH:mm");
  const end   = Utilities.formatDate(new Date(reservation.reserved_end), tz, "HH:mm");

  const age = calcAgeJst_(user.birthday);

  const subject = `${prefix} 新規予約 ${start}`;
  const body =
    `新規予約が入りました。

    予約ID: ${reservation.reservation_id}
    日時: ${start} - ${end}
    プラン: ${planNames}
    料金: ${priceStr}円
    要望: ${note || ""}

    顧客:
      ニックネーム: ${user.nick_name || ""}
      名前: ${user.name || ""}
      カナ: ${user.kana || ""}
      年齢: ${age ? age + "歳" : ""}
      電話: ${user.phone || ""}
      Email: ${user.email || ""}
      LINE_ID: ${user.line_user_id || ""}

    ステータス: CONFIRMED
    `;

  MailApp.sendEmail({ to: emails.join(","), subject, body });
}

function sendAdminMailOnCancel_(reservation, user, planNames, priceStr) {
  const emails = getAdminEmails_();
  if (emails.length === 0) return;

  const cfg = getConfigMap_();
  const prefix = (cfg.mail_subject_prefix || "[予約]").trim();
  const tz = "Asia/Tokyo";
  const start = Utilities.formatDate(new Date(reservation.reserved_start), tz, "yyyy/MM/dd HH:mm");
  const end   = Utilities.formatDate(new Date(reservation.reserved_end), tz, "HH:mm");
  const age = calcAgeJst_(user.birthday); // user.birthday に生年月日が入ってる前提

  const subject = `${prefix} キャンセル ${start}`;
  const body =
`予約がキャンセルされました。

予約ID: ${reservation.reservation_id}
日時: ${start} - ${end}
プラン: ${planNames}
料金: ${priceStr}円

顧客:
  ニックネーム: ${user.nick_name || ""}
  名前: ${user.name || ""}
  カナ: ${user.kana || ""}
  年齢: ${age ? age + "歳" : ""}
  電話: ${user.phone || ""}
  Email: ${user.email || ""}
  LINE_ID: ${user.line_user_id || ""}

ステータス: CANCELED
`;

  MailApp.sendEmail({ to: emails.join(","), subject, body });
}



function pushLineMessage_(lineUserId, text) {
  if (!lineUserId) return;

  const token = PropertiesService.getScriptProperties().getProperty("LINE_CHANNEL_ACCESS_TOKEN");
  if (!token) throw new Error("LINE_CHANNEL_ACCESS_TOKEN is not set");

  const url = "https://api.line.me/v2/bot/message/push";
  const payload = {
    to: lineUserId,
    messages: [{ type: "text", text }]
  };

  const res = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    console.log("push failed:", code, res.getContentText());
  }
}

function getAdminPhone_() {
  const cfg = getConfigMap_();
  return (cfg.admin_phone || "").trim();
}


/**
 * RESERVATIONS: status が CONFIRMED かつ reserved_end <= 現在時刻 のものを
 * status=COMPLETED に更新し、completed_at に更新時刻を記録する。
 *
 * - completed_at 列が無ければヘッダ行に自動追加します。
 * - reserved_end は Date/文字列/シリアルのどれでも coerceToDate_ で吸収
 * markCompletedReservations() を “列一括更新” にする（I/O削減）
 */
function markCompletedReservations() {
  const tz = "Asia/Tokyo";
  const now = new Date();

  const sh = sh_(SHEET_RESERVATIONS);
  if (!sh) throw new Error(`Sheet not found: ${SHEET_RESERVATIONS}`);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  // header
  let header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  let idx = indexMap_(header);

  requiredCols_(idx, ["status", "reserved_end"]);

  // completed_at 列が無ければ追加
  if (idx.completed_at === undefined) {
    sh.getRange(1, header.length + 1).setValue("completed_at");
    header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    idx = indexMap_(header);
  }

  const n = lastRow - 1;

  // 必要列だけ一括取得
  const statusCol = sh.getRange(2, idx.status + 1, n, 1).getValues();        // [[...]]
  const endCol    = sh.getRange(2, idx.reserved_end + 1, n, 1).getValues();
  const compCol   = sh.getRange(2, idx.completed_at + 1, n, 1).getValues();  // 既存値維持

  let updated = 0;

  for (let i = 0; i < n; i++) {
    const st = String(statusCol[i][0] || "").trim();
    if (st !== "CONFIRMED") continue;

    const endAt = coerceToDate_(endCol[i][0]);
    if (!endAt) continue;

    if (endAt.getTime() <= now.getTime()) {
      statusCol[i][0] = "COMPLETED";
      compCol[i][0] = now;
      updated++;
    }
  }

  if (updated > 0) {
    // 一括書き戻し（2回のsetValuesだけ）
    sh.getRange(2, idx.status + 1, n, 1).setValues(statusCol);
    sh.getRange(2, idx.completed_at + 1, n, 1).setValues(compCol);

    console.log(`[markCompletedReservations] updated: ${updated} rows at ${Utilities.formatDate(now, tz, "yyyy/MM/dd HH:mm:ss")}`);
  }
}



/**
 * 今年＆来年の祝日を BLACKOUTS に追加する（既存があればスキップ）
 * - 既存判定キー：from(yyyy/MM/dd) + all_day(=TRUE) の組み合わせ
 * - to は空でもOK（all_day=TRUEなら normalizeBlackout_ が終日化）
 */
function syncJapaneseHolidaysToBlackouts() {
  const tz = "Asia/Tokyo";
  const ss = ss_();

  let sh = sh_("BLACKOUTS");
  if (!sh) sh = ss.insertSheet("BLACKOUTS");
  sh = shFresh_("BLACKOUTS");   // ✅ここ（insert直後）

  // ---- ヘッダ保証
  let header = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(v => String(v || "").trim());
  const headerLower = header.map(h => h.toLowerCase());

  // 最低限必要な列
  const required = ["from", "to", "all_day", "is_active", "title"];
  // ヘッダ行が空っぽなら作り直し
  if (headerLower.filter(Boolean).length === 0) {
    sh.getRange(1, 1, 1, required.length).setValues([required]);
    header = required;
  } else {
    // 足りない列は末尾に追加
    for (const colName of required) {
      if (!headerLower.includes(colName)) {
        sh.getRange(1, sh.getLastColumn() + 1).setValue(colName);
        header.push(colName);
        headerLower.push(colName);
      }
    }
  }

  // index map（あなたの indexMap_ があればそれを使ってOK）
  const idx = {};
  headerLower.forEach((h, i) => { idx[h] = i; });

  // ---- 既存の「終日ブラックアウト日」セットを作る
  // key = yyyy/MM/dd
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const values = (lastRow >= 2) ? sh.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];

  const existingAllDaySet = new Set();
  for (const row of values) {
    const fromRaw = row[idx.from];
    if (!fromRaw) continue;

    // all_day 判定：TRUE または "TRUE"
    const allDayRaw = row[idx.all_day];
    const allDay = String(allDayRaw ?? "").toUpperCase() === "TRUE";
    if (!allDay) continue;

    // from を yyyy/MM/dd に正規化
    const d = coerceToDate_(fromRaw);
    if (!d) continue;
    const key = Utilities.formatDate(d, tz, "yyyy/MM/dd");
    existingAllDaySet.add(key);
  }

  // ---- 期間：今年1/1〜来年12/31
  const now = new Date();
  const y = Number(Utilities.formatDate(now, tz, "yyyy"));
  const start = new Date(y, 0, 1, 0, 0, 0, 0);
  const end = new Date(y + 1, 11, 31, 23, 59, 59, 999);

  const cal = CalendarApp.getCalendarById(HOLIDAY_CALENDAR_ID);
  if (!cal) throw new Error("Holiday calendar not found / cannot access: " + HOLIDAY_CALENDAR_ID);

  const events = cal.getEvents(start, end);

  // ---- 追加行を作る（既存ならスキップ）
  const rowsToAppend = [];
  for (const ev of events) {
    const d = ev.getStartTime();
    const key = Utilities.formatDate(d, tz, "yyyy/MM/dd");

    if (existingAllDaySet.has(key)) continue; // ★既存ならスキップ

    rowsToAppend.push({
      from: key,
      to: "",
      all_day: true,
      is_active: true,
      title: ev.getTitle(),
    });

    existingAllDaySet.add(key); // 同日が複数出ても二重追加しない
  }

  if (rowsToAppend.length === 0) {
    console.log("[syncJapaneseHolidaysToBlackouts] No new holidays to append.");
    return;
  }

  // ---- appendRow は遅いので setValues で一括追記
  const startRow = sh.getLastRow() + 1;
  const out = rowsToAppend.map(o => {
    // ヘッダ順に並べる
    return headerLower.map(h => (h in o ? o[h] : ""));
  });

  sh.getRange(startRow, 1, out.length, header.length).setValues(out);

  console.log(`[syncJapaneseHolidaysToBlackouts] Appended ${out.length} holidays.`);
}

// 空き枠返却　速度改善版
function getAvailabilityRangeMaterialsByDuration_(fromYmd, days, planId, durationMinOverride) {
  const tz = "Asia/Tokyo";

  // duration
  let durationMin = durationMinOverride;
  if (!durationMin) {
    if (!planId) throw new Error('MISSING_PARAM_plan_id_or_duration_min');
    const plan = getPlanById_(planId);
    if (!plan || !plan.is_active) throw new Error('PLAN_NOT_FOUND_OR_INACTIVE');
    durationMin = Number(plan.duration_min);
  }
  if (!Number.isFinite(durationMin) || durationMin <= 0) throw new Error('INVALID_duration_min');

  // days
  const nDays = Number.isFinite(days) && days > 0 && days <= 14 ? Math.floor(days) : 7;

  // range
  const fromStart = parseYmdAsLocalDate_(fromYmd);
  const rangeEnd = new Date(fromStart.getTime() + nDays * 24 * 60 * 60 * 1000);

  const granMin = getGranularityMinutes_();
  const bh = getBusinessHours_();

  const now = new Date();

  // ✅ しきい値：現在日時 + granularity_min
  const threshold = new Date(now.getTime() + granMin * 60 * 1000);
  const thresholdKey = Utilities.formatDate(threshold, tz, "yyyy-MM-dd");
  const thresholdMin = minutesOfDay_(threshold);

  // ---- 素材を作る
  const windowsByDate = buildOpenWindowsByDate_(fromStart, rangeEnd, bh);
  const busyByDateConfirmed = buildConfirmedBusyByDate_(fromStart, rangeEnd);
  const busyByDateBlackouts = buildBlackoutsBusyByDate_(fromStart, rangeEnd);

  // busy をマージ（CONFIRMED + BLACKOUTS）
  const busyByDate = {};
  for (let i = 0; i < nDays; i++) {
    const d0 = new Date(fromStart.getTime() + i * 86400000);
    const key = formatYmd_(d0);

    const a = busyByDateConfirmed[key] || [];
    const b = busyByDateBlackouts[key] || [];
    busyByDate[key] = mergeIntervals_(a.concat(b));
  }

  // ✅ 過去日（threshold基準）なら windows を空にする
  for (let i = 0; i < nDays; i++) {
    const dayStart = new Date(fromStart.getTime() + i * 86400000);
    const dayEnd   = new Date(dayStart.getTime() + 86400000);
    const key = formatYmd_(dayStart);

    if (dayEnd.getTime() <= threshold.getTime()) {
      windowsByDate[key] = [];
    }
  }

  // ✅ min_start_min_by_date：threshold が属する日だけ設定（跨ぎ対応）
  const minStartMinByDate = {};
  for (let i = 0; i < nDays; i++) {
    const d0 = new Date(fromStart.getTime() + i * 86400000);
    const key = formatYmd_(d0);
    minStartMinByDate[key] = (key === thresholdKey) ? thresholdMin : null;
  }

  return {
    from: formatYmd_(fromStart),
    days: nDays,
    granularity_min: granMin,
    required_duration_min: durationMin,
    business_open: bh.openStr,
    business_close: bh.closeStr,
    windows_by_date: windowsByDate,
    busy_by_date: busyByDate,
    min_start_min_by_date: minStartMinByDate
  };
}

// 週レンジの open windows を「SLOTS一括読み」→日別生成
function buildOpenWindowsByDate_(rangeStart, rangeEnd, bh) {
  const tz = "Asia/Tokyo";
  const nDays = Math.ceil((rangeEnd.getTime() - rangeStart.getTime()) / 86400000);

  // まずは全日をデフォルト営業時間で作る（定休日は空）
  const out = {};
  for (let i = 0; i < nDays; i++) {
    const dayStart = new Date(rangeStart.getTime() + i * 86400000);
    const key = formatYmd_(dayStart);

    if (isClosedByWeekday_(dayStart)) {
      out[key] = [];
      continue;
    }

    const openMin = bh.oh * 60 + bh.om;
    const closeMin = bh.ch * 60 + bh.cm;
    out[key] = (closeMin > openMin) ? [[openMin, closeMin]] : [];
  }

  // SLOTS があれば「該当日が1件でもある日」だけ SLOTS を優先
  const sh = sh_('SLOTS');
  if (!sh) return out;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return out;

  // ヘッダ
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const idx = indexMap_(header);
  if (idx.slot_start === undefined || idx.slot_end === undefined || idx.is_open === undefined) return out;

  // 必要列だけ一括取得
  const n = lastRow - 1;
  const colSlotStart = sh.getRange(2, idx.slot_start + 1, n, 1).getValues();
  const colSlotEnd   = sh.getRange(2, idx.slot_end + 1, n, 1).getValues();
  const colIsOpen    = sh.getRange(2, idx.is_open + 1, n, 1).getValues();

  // 日別に集計
  const slotsByDate = {}; // {key: [[fromMin,toMin], ...]}
  const hasAnyByDate = new Set();

  for (let i = 0; i < n; i++) {
    const isOpen = String(colIsOpen[i][0] ?? '').toUpperCase() === 'TRUE';
    if (!isOpen) continue;

    const from = coerceToDate_(colSlotStart[i][0]);
    const to   = coerceToDate_(colSlotEnd[i][0]);
    if (!from || !to || to <= from) continue;

    // 範囲外は除外
    if (from >= rangeEnd || to <= rangeStart) continue;

    const key = formatYmd_(from);
    hasAnyByDate.add(key);

    const fromMin = minutesOfDay_(from);
    const toMin = minutesOfDay_(to);

    if (!slotsByDate[key]) slotsByDate[key] = [];
    slotsByDate[key].push([fromMin, toMin]);
  }

  // SLOTSがある日だけ上書き（同日が複数windowでもOK）
  for (const key of hasAnyByDate) {
    out[key] = mergeIntervals_(slotsByDate[key] || []);
  }

  return out;
}

// 週レンジの CONFIRMED 予約を一括取得して busy区間にする
function buildConfirmedBusyByDate_(rangeStart, rangeEnd) {
  const sh = sh_(SHEET_RESERVATIONS);
  if (!sh) throw new Error(`Sheet not found: ${SHEET_RESERVATIONS}`);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return {};

  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const idx = indexMap_(header);

  requiredCols_(idx, ['status', 'reserved_start', 'reserved_end']);

  const n = lastRow - 1;

  // 必要列だけ一括取得
  const colStatus = sh.getRange(2, idx.status + 1, n, 1).getValues();
  const colStart  = sh.getRange(2, idx.reserved_start + 1, n, 1).getValues();
  const colEnd    = sh.getRange(2, idx.reserved_end + 1, n, 1).getValues();

  const out = {};

  for (let i = 0; i < n; i++) {
    const st = String(colStatus[i][0] || '').trim();
    if (st !== 'CONFIRMED') continue;

    const s = coerceToDate_(colStart[i][0]);
    const e = coerceToDate_(colEnd[i][0]);
    if (!s || !e || e <= s) continue;

    // 範囲外は除外
    if (s >= rangeEnd || e <= rangeStart) continue;

    // 日跨ぎは日ごとに分割（JS側が楽になる）
    splitIntoDailyIntervals_(s, e, out);
  }

  // merge
  for (const k of Object.keys(out)) out[k] = mergeIntervals_(out[k]);

  return out;
}

// 週レンジの BLACKOUTS を一括取得して busy区間にする
function buildBlackoutsBusyByDate_(rangeStart, rangeEnd) {
  const sh = sh_('BLACKOUTS');
  if (!sh) return {};

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return {};

  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(v => String(v).trim().replace(/[（(].*$/, '').trim());
  const idx = indexMap_(header);

  if (idx.from === undefined) throw new Error('BLACKOUTS_MISSING_COLUMN_from');

  const n = lastRow - 1;

  // 必要列だけ
  const colFrom     = sh.getRange(2, idx.from + 1, n, 1).getValues();
  const colTo       = (idx.to !== undefined) ? sh.getRange(2, idx.to + 1, n, 1).getValues() : null;
  const colAllDay   = (idx.all_day !== undefined) ? sh.getRange(2, idx.all_day + 1, n, 1).getValues() : null;
  const colIsActive = (idx.is_active !== undefined) ? sh.getRange(2, idx.is_active + 1, n, 1).getValues() : null;

  const out = {};

  for (let i = 0; i < n; i++) {
    const isActive = (colIsActive === null)
      ? true
      : String(colIsActive[i][0] ?? '').toUpperCase() !== 'FALSE';
    if (!isActive) continue;

    const fromRaw = colFrom[i][0];
    if (!fromRaw) continue;

    const toRaw = (colTo ? colTo[i][0] : null);
    const allDayRaw = (colAllDay ? colAllDay[i][0] : null);

    const norm = normalizeBlackout_(fromRaw, toRaw, allDayRaw);
    if (!norm) continue;

    if (norm.from >= rangeEnd || norm.to <= rangeStart) continue;

    splitIntoDailyIntervals_(norm.from, norm.to, out);
  }

  for (const k of Object.keys(out)) out[k] = mergeIntervals_(out[k]);

  return out;
}

// 共通ヘルパー（分/日跨ぎ分割/マージ）
function minutesOfDay_(d) {
  return d.getHours() * 60 + d.getMinutes();
}

// [start,end) を日別に分割して out[YYYY-MM-DD] に [startMin,endMin] をpush
function splitIntoDailyIntervals_(startAt, endAt, outByDate) {
  let cur = new Date(startAt.getTime());

  while (cur < endAt) {
    const dayStart = new Date(cur.getFullYear(), cur.getMonth(), cur.getDate(), 0, 0, 0, 0);
    const dayEnd = new Date(dayStart.getTime() + 86400000);

    const segStart = cur;
    const segEnd = new Date(Math.min(dayEnd.getTime(), endAt.getTime()));

    const key = formatYmd_(dayStart);
    const sMin = minutesOfDay_(segStart);
    const eMin = (segEnd.getTime() === dayEnd.getTime()) ? 1440 : minutesOfDay_(segEnd);

    if (!outByDate[key]) outByDate[key] = [];
    outByDate[key].push([sMin, eMin]);

    cur = segEnd;
  }
}

// intervals: [[s,e], ...] をソートしてマージ（半開区間）
function mergeIntervals_(intervals) {
  if (!intervals || intervals.length === 0) return [];
  const a = intervals
    .map(x => [Number(x[0]), Number(x[1])])
    .filter(x => Number.isFinite(x[0]) && Number.isFinite(x[1]) && x[1] > x[0])
    .sort((p, q) => p[0] - q[0]);

  const out = [a[0]];
  for (let i = 1; i < a.length; i++) {
    const [s, e] = a[i];
    const last = out[out.length - 1];
    if (s <= last[1]) {
      last[1] = Math.max(last[1], e);
    } else {
      out.push([s, e]);
    }
  }
  return out;
}

function parseBirthdayToDate_(birthdayRaw) {
  if (!birthdayRaw) return null;

  // Date型
  if (Object.prototype.toString.call(birthdayRaw) === "[object Date]") {
    return isNaN(birthdayRaw.getTime()) ? null : birthdayRaw;
  }

  // 文字列 "YYYY-MM-DD" or "YYYY/MM/DD"
  const s = String(birthdayRaw).trim();
  let m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (!m) m = /^(\d{4})\/(\d{2})\/(\d{2})$/.exec(s);
  if (m) {
    const y = Number(m[1]);
    const mo = Number(m[2]) - 1;
    const d = Number(m[3]);
    const dt = new Date(y, mo, d, 0, 0, 0, 0);
    return isNaN(dt.getTime()) ? null : dt;
  }

  // その他（念のため）
  const dt = new Date(s);
  return isNaN(dt.getTime()) ? null : dt;
}

function calcAgeJst_(birthdayRaw, nowDate) {
  const tz = "Asia/Tokyo";
  const b = parseBirthdayToDate_(birthdayRaw);
  if (!b) return "";

  // JSTの「今日」
  const now = nowDate ? new Date(nowDate) : new Date();
  const y = Number(Utilities.formatDate(now, tz, "yyyy"));
  const m = Number(Utilities.formatDate(now, tz, "MM"));
  const d = Number(Utilities.formatDate(now, tz, "dd"));

  const by = b.getFullYear();
  const bm = b.getMonth() + 1;
  const bd = b.getDate();

  let age = y - by;
  // 今年の誕生日がまだなら -1
  if (m < bm || (m === bm && d < bd)) age--;

  // 異常値ガード（空返し）
  if (!Number.isFinite(age) || age < 0 || age > 130) return "";
  return String(age);
}

// allowed_gender を読むヘルパー
function getAllowedGenders_() {
  const cfg = getConfigMap_();
  const raw = (cfg.allowed_gender || "").trim();
  if (!raw) return []; // 空なら制限なし（全許可）

  return raw
    .split(",")
    .map(s => String(s).trim().toLowerCase())
    .filter(Boolean);
}

function isGenderAllowed_(gender, allowed) {
  if (!allowed || allowed.length === 0) return true; // 制限なし
  const g = String(gender || "").trim().toLowerCase();
  return allowed.includes(g);
}

function getAdminLineUserId_() {
  const cfg = getConfigMap_();
  return String(cfg.admin_line_user_id || "").trim();
}


function renderTodayGanttChart() {
  const ss = ss_();
  const tz = "Asia/Tokyo";

  const resSh = sh_(SHEET_RESERVATIONS);
  if (!resSh) throw new Error(`Sheet not found: ${SHEET_RESERVATIONS}`);

  let ganttSh = sh_(SHEET_TODAY_GANTT);
  if (!ganttSh) {
    ss.insertSheet(SHEET_TODAY_GANTT);
    ganttSh = shFresh_(SHEET_TODAY_GANTT);
  }

  // ===== 今日（JST）範囲 =====
  const now = new Date();
  const y = Number(Utilities.formatDate(now, tz, "yyyy"));
  const m = Number(Utilities.formatDate(now, tz, "MM")) - 1;
  const d = Number(Utilities.formatDate(now, tz, "dd"));
  const dayStart = new Date(y, m, d, 0, 0, 0, 0);
  const dayEnd   = new Date(y, m, d + 1, 0, 0, 0, 0);

  // ===== CONFIG =====
  const granMin = getGranularityMinutes_();  // 既存
  const bh = getBusinessHours_();            // 既存（openStr/closeStr, oh/om, ch/cm）
  const openMin  = bh.oh * 60 + bh.om;
  const closeMin = bh.ch * 60 + bh.cm;

  // ===== 予約取得 =====
  const values = resSh.getDataRange().getValues();
  if (values.length < 2) {
    // シートを空描画
    drawEmptyTodayGantt_(ganttSh, dayStart, openMin, closeMin, granMin, tz);
    return;
  }

  const header = values[0].map(v => String(v).trim());
  const idx = indexMap_(header);
  requiredCols_(idx, ["reserved_start", "reserved_end", "status", "line_user_id"]);

  // 名前解決（既存）
  const userNameByLineId = buildUserNameMap_();

  /** @type {{start:Date,end:Date,label:string}[]} */
  const todays = [];

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const st = String(row[idx.status] || "").trim();
    if (!st || st === "CANCELED") continue;

    const s = coerceToDate_(row[idx.reserved_start]);
    const e = coerceToDate_(row[idx.reserved_end]);
    if (!s || !e || e <= s) continue;

    // 今日と重なるものだけ（開始が今日でも、日跨ぎも一応拾う）
    if (s >= dayEnd || e <= dayStart) continue;

    // 今日の範囲にクリップ
    const cs = new Date(Math.max(s.getTime(), dayStart.getTime()));
    const ce = new Date(Math.min(e.getTime(), dayEnd.getTime()));

    const lineId = String(row[idx.line_user_id] || "").trim();

    const customer =
      (idx.name_snapshot !== undefined && String(row[idx.name_snapshot] || "").trim())
        ? String(row[idx.name_snapshot]).trim()
        : (userNameByLineId[lineId] || lineId || "（不明）");

    const planNames =
      (idx.plan_names_snapshot !== undefined && String(row[idx.plan_names_snapshot] || "").trim())
        ? String(row[idx.plan_names_snapshot]).trim()
        : "";

    const hhmm = Utilities.formatDate(cs, tz, "HH:mm");
    const label = `${hhmm} ${customer}${planNames ? " " + planNames : ""}`.trim();

    todays.push({ start: cs, end: ce, label });
  }

  // 時刻順
  todays.sort((a, b) => a.start.getTime() - b.start.getTime());

  // ===== 時間軸の開始/終了を決定 =====
  // 基本：営業時間（openMin/closeMin）
  // ただし、当日の予約が営業時間外にある場合は、その最小〜最大まで拡張
  let minAxis = openMin;
  let maxAxis = closeMin;

  if (todays.length > 0) {
    const minRes = Math.min(...todays.map(x => minutesOfDay_(x.start)));
    const maxRes = Math.max(...todays.map(x => minutesOfDay_(x.end)));
    minAxis = Math.min(openMin, minRes);
    maxAxis = Math.max(closeMin, maxRes);
  }

  // 粒度に合わせて軸を整形（開始：切り下げ / 終了：切り上げ）
  minAxis = floorToGran_(minAxis, granMin);
  maxAxis = ceilToGran_(maxAxis, granMin);

  // ===== 描画 =====
  ganttSh.clear();

  const dateTitle = Utilities.formatDate(dayStart, tz, "yyyy年M月d日");
  const nCols = Math.round((maxAxis - minAxis) / granMin) + 1; // 終端含む
  const startCol = 2; // B列開始

  // 1行目：日付タイトル（結合）
  ganttSh.getRange(1, startCol, 1, nCols).merge();
  ganttSh.getRange(1, startCol).setValue(dateTitle)
    .setFontSize(18).setFontWeight("bold")
    .setHorizontalAlignment("left").setVerticalAlignment("middle");

  // 2行目：時間ラベル
  const timeRow = [];
  for (let t = minAxis; t <= maxAxis; t += granMin) timeRow.push(minToHHMM_(t));
  ganttSh.getRange(2, startCol, 1, nCols).setValues([timeRow])
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontSize(10);

  ganttSh.setFrozenRows(2);

  // 列幅（時間軸）
  ganttSh.setColumnWidths(startCol, nCols, 42);
  ganttSh.setRowHeight(1, 30);
  ganttSh.setRowHeight(2, 22);

  // 罫線（縦の点線っぽく）
  const gridRows = Math.max(22, 3 + todays.length + 10);
  const gridRange = ganttSh.getRange(2, startCol, gridRows, nCols);
  gridRange.setBorder(null, true, null, true, true, null, "#BDBDBD", SpreadsheetApp.BorderStyle.DOTTED);

  // 予約バー
  const barColor = "#E9A3A3"; // 画像に寄せたピンク
  for (let i = 0; i < todays.length; i++) {
    const row = 3 + i;
    ganttSh.setRowHeight(row, 18);

    const sMin = minutesOfDay_(todays[i].start);
    const eMin = minutesOfDay_(todays[i].end);

    const colStart = startCol + Math.round((sMin - minAxis) / granMin);
    const colEnd   = startCol + Math.round((eMin - minAxis) / granMin);

    const width = Math.max(1, colEnd - colStart); // [start,end) なので end-start
    const rng = ganttSh.getRange(row, colStart, 1, width);
    rng.setBackground(barColor);

    // ラベルはバーの左端に
    ganttSh.getRange(row, colStart).setValue(todays[i].label).setFontSize(10);
  }

  // 見た目調整
  ganttSh.getRange(1, 1, ganttSh.getMaxRows(), 1).setBackground(null); // A列は未使用
  ganttSh.setActiveSelection(ganttSh.getRange("B1"));
}

// 予約0件のときも営業時間ベースで枠だけ描く
function drawEmptyTodayGantt_(ganttSh, dayStart, openMin, closeMin, granMin, tz) {
  ganttSh.clear();

  const minAxis = floorToGran_(openMin, granMin);
  const maxAxis = ceilToGran_(closeMin, granMin);
  const nCols = Math.round((maxAxis - minAxis) / granMin) + 1;
  const startCol = 2;

  const dateTitle = Utilities.formatDate(dayStart, tz, "yyyy年M月d日");
  ganttSh.getRange(1, startCol, 1, nCols).merge();
  ganttSh.getRange(1, startCol).setValue(dateTitle)
    .setFontSize(18).setFontWeight("bold")
    .setHorizontalAlignment("left").setVerticalAlignment("middle");

  const timeRow = [];
  for (let t = minAxis; t <= maxAxis; t += granMin) timeRow.push(minToHHMM_(t));
  ganttSh.getRange(2, startCol, 1, nCols).setValues([timeRow])
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontSize(10);

  ganttSh.setFrozenRows(2);
  ganttSh.setColumnWidths(startCol, nCols, 42);
  ganttSh.setRowHeight(1, 30);
  ganttSh.setRowHeight(2, 22);

  const gridRows = 22;
  const gridRange = ganttSh.getRange(2, startCol, gridRows, nCols);
  gridRange.setBorder(null, true, null, true, true, null, "#BDBDBD", SpreadsheetApp.BorderStyle.DOTTED);
}

// minutes(0-1440) -> "HH:mm"
function minToHHMM_(min) {
  const hh = String(Math.floor(min / 60)).padStart(2, "0");
  const mm = String(min % 60).padStart(2, "0");
  return `${hh}:${mm}`;
}

function floorToGran_(min, gran) {
  return Math.floor(min / gran) * gran;
}
function ceilToGran_(min, gran) {
  return Math.ceil(min / gran) * gran;
}


/** =========================
 *  バッチ処理（定期事項）
 * ========================= */

function reservationDetailsUpdate(){
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {   // 1秒で取れなければスキップ
    console.log('[reservationDetailsUpdate] skipped (lock busy)');
    return;
  }
  try {
    refreshTodayReservations();
    renderTodayGanttChart();
    renderReservationCalendar();
  } finally {
    lock.releaseLock();
  }
}
