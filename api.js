/**
 * ============================
 * 設定（ここだけ埋める）
 * ============================
 */
const GAS_URL = "https://script.google.com/macros/s/AKfycbxpVHAXUQoPB1tgXqXT_Syy0gk2GLFCPZ35dUjUZ_XVi6aap5kud587Vecan4xFavG_8Q/exec";

/**
 * ============================
 * LINE公式アカウント（LIFF）適用時にコメントを外してください
 * ============================
 */
// const LIFF_ID = "YOUR_LIFF_ID";
// ※ index.html に以下も必要
// <script src="https://static.line-scdn.net/liff/edge/2/sdk.js"></script>


/**
 * API：GET/POST
 */
async function apiGet(params) {
  const url = new URL(GAS_URL);
  Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, v));
  const res = await fetch(url.toString(), { method: "GET" });
  return res.json();
}
async function apiPost(payload) {
  const res = await fetch(GAS_URL, {
    method: "POST",
    // ★ headers を付けない（ここが重要）
    body: JSON.stringify(payload),
  });

  // GASはJSONを返すが、念のため text→JSON にする
  const text = await res.text();
  try {
    return JSON.parse(text);
  } catch {
    // エラー調査用にそのまま返す
    return { ok: false, error: "INVALID_JSON_RESPONSE", raw: text };
  }
}

/**
 * UI helpers
 */
const $ = (id) => document.getElementById(id);
function show(el) { el.classList.remove("hidden"); }
function hide(el) { el.classList.add("hidden"); }
function setError(msg) {
  const box = $("errorBox");
  if (!box) return;
  if (!msg) { hide(box); box.textContent = ""; return; }
  box.textContent = msg;
  show(box);
}
function setStatus(msg) {
  const el = $("statusText");
  if (el) el.textContent = msg;
}
function fmtYmd(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth()+1).padStart(2,"0");
  const dd = String(d.getDate()).padStart(2,"0");
  return `${y}-${m}-${dd}`;
}
function fmtTime(dateStrOrDate) {
  const d = new Date(dateStrOrDate);
  const hh = String(d.getHours()).padStart(2,"0");
  const mm = String(d.getMinutes()).padStart(2,"0");
  return `${hh}:${mm}`;
}
function fmtJstDateTime(dateStrOrDate) {
  const d = new Date(dateStrOrDate);
  return `${fmtYmd(d)} ${fmtTime(d)}`;
}

/**
 * state
 */
const state = {
  lineUserId: null,
  user: null,
  plans: [],

  // 複数選択
  selectedPlans: [],
  totalDurationMin: 0,
  totalPrice: 0,

  // 週表示
  weekStart: null,
  granMin: null,
  businessOpen: null,
  businessClose: null,

  selectedStartAt: null,
  lastReservation: null,
};

/**
 * ============================
 * 静的Webページ用：仮ユーザーIDの決定ロジック
 * - URLクエリ ?uid=xxx があればそれを使う
 * - なければ localStorage に保存されたIDを使う
 * - なければ新規生成して保存する
 * ============================
 */
function resolveStaticUserId_() {
  const params = new URLSearchParams(location.search);
  const uidFromQuery = (params.get("uid") || "").trim();
  if (uidFromQuery) {
    localStorage.setItem("STATIC_UID", uidFromQuery);
    return uidFromQuery;
  }

  const saved = (localStorage.getItem("STATIC_UID") || "").trim();
  if (saved) return saved;

  const gen = `WEB-${cryptoRandomId_()}`;
  localStorage.setItem("STATIC_UID", gen);
  return gen;
}

function cryptoRandomId_() {
  // ブラウザのcryptoが使えればそれを使用
  if (typeof crypto !== "undefined" && crypto.getRandomValues) {
    const a = new Uint8Array(16);
    crypto.getRandomValues(a);
    return [...a].map(b => b.toString(16).padStart(2,"0")).join("");
  }
  // フォールバック
  return `${Date.now().toString(16)}-${Math.random().toString(16).slice(2)}`;
}

/**
 * 初期化：静的Webとして起動
 */
document.addEventListener("DOMContentLoaded", () => {
  initStaticWeb();
});

/**
 * ============================
 * 静的Web起動（LIFF無し）
 * ============================
 */
async function initStaticWeb() {
  wireEvents();

  try {
    setError("");
    setStatus("初期化中…");

    // 静的WebではLINE userIdが取れないため、仮IDを使用
    state.lineUserId = resolveStaticUserId_();

    setStatus(`ユーザー確認中…（UID: ${state.lineUserId}）`);
    await loadPlans();
    await checkUser();

    setStatus("準備完了");
  } catch (e) {
    console.error(e);
    setStatus("初期化に失敗");
    setError(String(e?.message || e));
  }
}

/**
 * ============================
 * LINE公式アカウント（LIFF）適用時にコメントを外してください
 * ============================
 */
// async function initLiff() {
//   wireEvents();
//
//   try {
//     setStatus("LIFF初期化中…");
//     setError("");
//
//     await liff.init({ liffId: LIFF_ID });
//
//     // LINE外ブラウザの場合
//     if (!liff.isInClient()) {
//       setStatus("LINEアプリ内で開いてください（LIFF）");
//       return;
//     }
//
//     // 未ログインならログイン
//     if (!liff.isLoggedIn()) {
//       liff.login();
//       return;
//     }
//
//     const profile = await liff.getProfile();
//     state.lineUserId = profile.userId;
//
//     setStatus("ユーザー確認中…");
//     await loadPlans();
//     await checkUser();
//
//   } catch (e) {
//     console.error(e);
//     setStatus("初期化に失敗");
//     setError(String(e?.message || e));
//   }
// }

/**
 * events
 */
function wireEvents() {
  $("btnReload")?.addEventListener("click", () => location.reload());

  $("btnMyPage")?.addEventListener("click", async () => {
    await openMyPage();
  });
  $("btnGoMyPage")?.addEventListener("click", async () => {
    await openMyPage();
  });

  $("btnBackHome")?.addEventListener("click", () => {
    hide($("myPageCard"));
    show($("bookingCard"));
  });

  $("registerForm")?.addEventListener("submit", async (ev) => {
    ev.preventDefault();
    await registerUser();
  });

  $("btnBackToPlan")?.addEventListener("click", () => gotoStep(1));
  $("btnBackToDateTime")?.addEventListener("click", () => gotoStep(2));

  $("btnConfirmReserve")?.addEventListener("click", async () => {
    await reserve();
  });

  $("btnNewBooking")?.addEventListener("click", async () => {
    state.selectedPlans = [];
    state.totalDurationMin = 0;
    state.totalPrice = 0;
    state.selectedStartAt = null;

    // テーブルを初期化
    $("timeTableHead") && ($("timeTableHead").innerHTML = "");
    $("timeTableBody") && ($("timeTableBody").innerHTML = "");

    $("selectedPlanBox") && ($("selectedPlanBox").textContent = "—");

    gotoStep(1);
    renderPlans();
  });

  $("btnPrevWeek")?.addEventListener("click", async () => {
    shiftWeek(-7);
    await loadAvailabilityWeek();
  });
  $("btnNextWeek")?.addEventListener("click", async () => {
    shiftWeek(7);
    await loadAvailabilityWeek();
  });
  $("btnToday")?.addEventListener("click", async () => {
    state.weekStart = startOfWeek(new Date());
    await loadAvailabilityWeek();
  });
}

/**
 * Step control
 */
function gotoStep(n) {
  const steps = [$("step1"), $("step2"), $("step3"), $("step4")].filter(Boolean);
  steps.forEach(s => s.classList.remove("stepper__item--active"));
  steps[n-1]?.classList.add("stepper__item--active");

  const p1 = $("panelPlan");
  const p2 = $("panelDateTime");
  const p3 = $("panelConfirm");
  const p4 = $("panelDone");

  [p1,p2,p3,p4].forEach(p => p && hide(p));
  if (n === 1) p1 && show(p1);
  if (n === 2) p2 && show(p2);
  if (n === 3) p3 && show(p3);
  if (n === 4) p4 && show(p4);
}

/**
 * Data loading
 */
async function loadPlans() {
  const r = await apiGet({ action: "plans" });
  if (!r.ok) throw new Error(r.error || "plans_failed");
  state.plans = r.plans || [];
}

async function checkUser() {
  const r = await apiGet({ action: "me", line_user_id: state.lineUserId });
  if (!r.ok) throw new Error(r.error || "me_failed");

  if (r.exists) {
    state.user = r.user;
    setStatus("予約画面を準備中…");
    hide($("registerCard"));
    show($("bookingCard"));
    gotoStep(1);
    renderPlans();
    setDefaultDate();
  } else {
    state.user = null;
    setStatus("初回登録が必要です");
    hide($("bookingCard"));
    show($("registerCard"));
  }
}

function setDefaultDate() {
  // 週表示の起点を今日の週にする
  state.weekStart = startOfWeek(new Date());
  const from = ymd(state.weekStart);
  const to = ymd(addDays(state.weekStart, 6));
  const label = $("weekLabel");
  if (label) label.textContent = `${from} 〜 ${to}`;
}

/**
 * Registration
 */
async function registerUser() {
  if (!state.lineUserId) {
  state.lineUserId = resolveStaticUserId_(); // 静的Web用の仮UID生成
  }
  setError("");
  setStatus("登録中…");

  const fd = new FormData($("registerForm"));
  const payload = {
    action: "users_upsert",
    line_user_id: state.lineUserId,
    name: String(fd.get("name") || "").trim(),
    kana: String(fd.get("kana") || "").trim(),
    gender: String(fd.get("gender") || "").trim(),
    phone: String(fd.get("phone") || "").replace(/[^0-9]/g, "").trim(),
    email: String(fd.get("email") || "").trim(),
  };

  try {
    const r = await apiPost(payload);
    if (!r.ok) throw new Error(r.error || "register_failed");
    state.user = r.user;
    hide($("registerCard"));
    show($("bookingCard"));
    renderPlans();
    setDefaultDate();
    gotoStep(1);
    setStatus("登録完了。予約できます");
  } catch (e) {
    console.error(e);
    setStatus("登録に失敗");
    setError(String(e?.message || e));
  }
}

/**
 * Booking UI
 */
function renderPlans() {
  const list = $("planList");
  list.innerHTML = "";

  state.selectedPlans = [];
  state.totalDurationMin = 0;
  state.totalPrice = 0;

  const summaryId = "planSummaryBox";
  if (!$(summaryId)) {
    const box = document.createElement("div");
    box.className = "alert";
    box.id = summaryId;
    box.textContent = "プランを選択してください（複数選択可）";
    list.parentElement.insertBefore(box, list);
  } else {
    $(summaryId).textContent = "プランを選択してください（複数選択可）";
  }

  if (!state.plans.length) {
    list.innerHTML = `<div class="alert">利用可能なプランがありません。</div>`;
    return;
  }

  // 決定ボタン（下部）
  const footer = document.createElement("div");
  footer.className = "row row--end";
  footer.innerHTML = `<button class="btn btn--primary" type="button" id="btnPlanDecide" disabled>日時を選ぶ</button>`;

  state.plans.forEach(plan => {
    const el = document.createElement("div");
    el.className = "item";

    const id = `plan_${plan.plan_id}`;
    el.innerHTML = `
      <div class="row row--space">
        <div>
          <div class="item__title">${escapeHtml(plan.plan_name)}</div>
          <div class="item__meta">所要時間: ${Number(plan.duration_min)}分 / 料金: ¥${Number(plan.price).toLocaleString()}</div>
        </div>
        <label class="row" style="gap:8px;">
          <input type="checkbox" id="${id}" data-plan-id="${escapeHtml(plan.plan_id)}" />
          <span class="muted">選択</span>
        </label>
      </div>
    `;

    el.querySelector("input").addEventListener("change", (ev) => {
      const checked = ev.target.checked;

      if (checked) {
        state.selectedPlans.push(plan);
      } else {
        state.selectedPlans = state.selectedPlans.filter(p => p.plan_id !== plan.plan_id);
      }

      // 合計更新
      state.totalDurationMin = state.selectedPlans.reduce((a, p) => a + Number(p.duration_min), 0);
      state.totalPrice = state.selectedPlans.reduce((a, p) => a + Number(p.price), 0);

      // 表示更新
      const names = state.selectedPlans.map(p => p.plan_name).join(" + ");
      $(summaryId).innerHTML = state.selectedPlans.length
        ? `選択中：<b>${escapeHtml(names)}</b><br>合計：<b>${state.totalDurationMin}分</b> / <b>¥${Number(state.totalPrice).toLocaleString()}</b>`
        : "プランを選択してください（複数選択可）";

      const btn = $("btnPlanDecide");
      btn.disabled = state.selectedPlans.length === 0;
      // 右上の選択中プラン表示も更新
      $("selectedPlanBox").textContent = state.selectedPlans.length
        ? `${names}（${state.totalDurationMin}分 / ¥${Number(state.totalPrice).toLocaleString()}）`
        : "—";
    });

    list.appendChild(el);
  });

  list.appendChild(footer);

  $("btnPlanDecide").addEventListener("click", async () => {
    // 週の起点
    state.weekStart = startOfWeek(new Date());
    await loadAvailabilityWeek(); // 週表示取得
    gotoStep(2);
  });
}

async function loadAvailability() {
  setError("");

  if (!state.selectedPlan) return;
  if (!state.selectedDate) return;

  setStatus("空き枠を取得中…");
  $("slotGrid").innerHTML = "";
  $("slotHint").textContent = "";

  try {
    const r = await apiGet({
      action: "availability",
      date: state.selectedDate,
      plan_id: state.selectedPlan.plan_id,
    });
    if (!r.ok) throw new Error(r.error || "availability_failed");

    state.availableSlots = r.available || [];
    $("slotHint").textContent = r.slot_source_hint || "※ 枠情報を参照しています";

    renderSlots();
    setStatus("日時を選択してください");
  } catch (e) {
    console.error(e);
    setStatus("空き枠取得に失敗");
    setError(String(e?.message || e));
  }
}

function renderSlots() {
  const grid = $("slotGrid");
  grid.innerHTML = "";

  const slots = state.availableSlots || [];
  if (!slots.length) {
    grid.innerHTML = `<div class="alert">この日は予約枠がありません（受付不可、または満席/受付終了の可能性）。</div>`;
    return;
  }

  slots.forEach(iso => {
    const btn = document.createElement("div");
    btn.className = "slot";
    btn.textContent = fmtTime(iso);
    btn.addEventListener("click", () => {
      state.selectedStartAt = iso;
      [...grid.children].forEach(c => c.classList.remove("slot--selected"));
      btn.classList.add("slot--selected");
      buildConfirm();
      gotoStep(3);
    });
    grid.appendChild(btn);
  });
}

function buildConfirm() {
  const start = state.selectedStartAt;
  if (!state.selectedPlans.length || !start) return;

  const startDt = new Date(start);
  const endDt = new Date(startDt.getTime() + Number(state.totalDurationMin) * 60 * 1000);
  const names = state.selectedPlans.map(p => p.plan_name).join(" + ");

  $("confirmBox").innerHTML = `
    <div><b>プラン</b>：${escapeHtml(names)}</div>
    <div><b>日時</b>：${escapeHtml(fmtJstDateTime(startDt))} 〜 ${escapeHtml(fmtTime(endDt))}</div>
    <div><b>所要時間（合計）</b>：${Number(state.totalDurationMin)}分</div>
    <div><b>料金（合計）</b>：¥${Number(state.totalPrice).toLocaleString()}</div>
  `;
}

async function reserve() {
  setError("");

  if (!state.selectedPlans || state.selectedPlans.length === 0 || !state.selectedStartAt) {
    setError("プランと日時を選択してください。");
    return;
  }

  setStatus("予約確定中…");

  try {
    const r = await apiPost({
      action: "reserve",
      line_user_id: state.lineUserId,
      plan_ids: state.selectedPlans.map(p => p.plan_id),
      start_at: state.selectedStartAt,
    });

    if (!r.ok) throw new Error(r.error || "reserve_failed");

    state.lastReservation = r.reservation;

    $("doneBox").innerHTML = `
      <div><b>予約番号</b>：${escapeHtml(r.reservation.reservation_id || "")}</div>
      <div><b>日時</b>：${escapeHtml(fmtJstDateTime(r.reservation.reserved_start))} 〜 ${escapeHtml(fmtJstDateTime(r.reservation.reserved_end))}</div>
      <div><b>プラン</b>：${escapeHtml(r.reservation.plan_names || "")}</div>
      <div class="muted small">キャンセルは「予約一覧」から可能です。</div>
    `;

    gotoStep(4);
    setStatus("予約が完了しました");
  } catch (e) {
    console.error(e);
    setStatus("予約に失敗");
    setError(String(e?.message || e));
  }
}

/**
 * My page
 */
async function openMyPage() {
  setError("");
  setStatus("予約一覧を取得中…");

  try {
    const r = await apiGet({
      action: "my_reservations",
      line_user_id: state.lineUserId,
      status: "CONFIRMED",
    });
    if (!r.ok) throw new Error(r.error || "my_reservations_failed");

    renderReservations(r.reservations || []);
    hide($("bookingCard"));
    hide($("registerCard"));
    show($("myPageCard"));
    setStatus("予約一覧");
  } catch (e) {
    console.error(e);
    setStatus("一覧取得に失敗");
    setError(String(e?.message || e));
  }
}

function renderReservations(list) {
  const root = $("reservationList");
  root.innerHTML = "";

  if (!list.length) {
    root.innerHTML = `<div class="alert">予約がありません。</div>`;
    return;
  }

  list.forEach(resv => {
    const el = document.createElement("div");
    el.className = "item";

    const title = `${escapeHtml(resv.plan_name || "プラン")} / ${escapeHtml(fmtJstDateTime(resv.reserved_start))}`;
    const meta = `所要: ${Number(resv.duration_min)}分 / ¥${Number(resv.price).toLocaleString()}`;

    el.innerHTML = `
      <div class="item__title">${title}</div>
      <div class="item__meta">${meta}</div>
      <div class="item__actions">
        <button class="btn btn--danger" type="button">キャンセル</button>
      </div>
    `;

    el.querySelector("button").addEventListener("click", async () => {
      const ok = confirm("この予約をキャンセルしますか？");
      if (!ok) return;
      await cancelReservation(resv.cancel_token);
      await openMyPage();
    });

    root.appendChild(el);
  });
}

async function cancelReservation(cancelToken) {
  setError("");
  setStatus("キャンセル処理中…");

  try {
    const r = await apiPost({ action: "cancel", cancel_token: cancelToken });
    if (!r.ok) throw new Error(r.error || "cancel_failed");
    setStatus("キャンセルしました");
  } catch (e) {
    console.error(e);
    setStatus("キャンセルに失敗");
    setError(String(e?.message || e));
  }
}

/**
 * XSS safe
 */
function escapeHtml(s) {
  return String(s ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function startOfWeek(d){
  // 月曜始まり
  const x = new Date(d);
  const day = (x.getDay() + 6) % 7; // Mon=0 ... Sun=6
  x.setHours(0,0,0,0);
  x.setDate(x.getDate() - day);
  return x;
}
function shiftWeek(days){
  if (!state.weekStart) state.weekStart = startOfWeek(new Date());
  const x = new Date(state.weekStart);
  x.setDate(x.getDate() + days);
  state.weekStart = x;
}
function addDays(d, n){
  const x = new Date(d);
  x.setDate(x.getDate() + n);
  return x;
}
function ymd(d){
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
}
function dowJa(d){
  return ["日","月","火","水","木","金","土"][d.getDay()];
}

async function loadAvailabilityWeek() {
  setError("");

  // ★ 複数選択対応：selectedPlansで判定
  if (!state.selectedPlans || state.selectedPlans.length === 0) {
    setError("プランを選択してください。");
    return;
  }
  if (!state.totalDurationMin || state.totalDurationMin <= 0) {
    setError("所要時間（合計）が不正です。");
    return;
  }

  setStatus("空き枠（週）を取得中…");

  $("slotHint") && ($("slotHint").textContent = "");
  $("timeTableHead") && ($("timeTableHead").innerHTML = "");
  $("timeTableBody") && ($("timeTableBody").innerHTML = "");

  if (!state.weekStart) state.weekStart = startOfWeek(new Date());

  try {
    const r = await apiGet({
      action: "availability_range",
      from: ymd(state.weekStart),
      days: "7",
      // plan_ids は送ってもいいが、GASが使わなくてもOK（duration_minが本命）
      plan_ids: state.selectedPlans.map(p => p.plan_id).join(","),
      duration_min: String(state.totalDurationMin),
    });

    console.log("availability_range response:", r);

    if (!r.ok) throw new Error(r.error || JSON.stringify(r));

    // ★ r.ok 確認後に反映
    state.granMin = r.granularity_min;
    state.businessOpen = r.business_open;
    state.businessClose = r.business_close;

    $("slotHint") && ($("slotHint").textContent = r.slot_source_hint || "");

    const dayResults = [...Array(7)].map((_, i) => {
      const d = addDays(state.weekStart, i);
      const key = ymd(d);
      return { date: d, slots: (r.by_date && r.by_date[key]) ? r.by_date[key] : [] };
    });

    renderWeekTable(dayResults);

    const from = ymd(state.weekStart);
    const to = ymd(addDays(state.weekStart, 6));
    $("weekLabel") && ($("weekLabel").textContent = `${from} 〜 ${to}`);

    setStatus("空き枠を選択してください");
  } catch (e) {
    console.error(e);
    setStatus("空き枠取得に失敗");
    setError(String(e?.message || e));
  }
}

function renderWeekTable(dayResults){
  console.log("renderWeekTable dayResults:", dayResults);
  // dayResults: [{date, slots:[iso...]}, ...7]

  // 表の縦軸（時間）を営業時間と粒度で固定生成（×も表示するため）
  const BUSINESS_OPEN = state.businessOpen || "09:00";
  const BUSINESS_CLOSE = state.businessClose || "18:00";
  const GRAN_MIN = Number(state.granMin) || 30;

  const times = buildTimes_(BUSINESS_OPEN, BUSINESS_CLOSE, GRAN_MIN);

  // ヘッダー
  $("timeTableHead").innerHTML = `
    <tr>
      <th>時間</th>
      ${dayResults.map((dr)=>{
        const d = dr.date;
        const label = `${d.getMonth()+1}/${d.getDate()} (${dowJa(d)})`;
        return `<th>${label}</th>`;
      }).join("")}
    </tr>
  `;

  // body
  const rows = times.map(time => {
    const tds = dayResults.map((dr)=>{
      const iso = findIsoByDateAndTime_(dr.slots, dr.date, time);
      const ok = !!iso;
      const cls = ok ? "timeCell timeCell--ok" : "timeCell timeCell--ng";
      const badge = ok ? `<span class="badge badge--ok">○</span>` : `<span class="badge badge--ng">×</span>`;
      const dataAttr = ok ? `data-iso="${iso}"` : "";
      return `<td><div class="${cls}" ${dataAttr}>${badge}</div></td>`;
    }).join("");

    return `<tr><td>${time}</td>${tds}</tr>`;
  }).join("");

  $("timeTableBody").innerHTML = rows;

  // クリックイベント（delegation）
  const body = $("timeTableBody");
  body.onclick = (ev) => {
    const cell = ev.target.closest(".timeCell--ok");
    if (!cell) return;
    const iso = cell.getAttribute("data-iso");
    if (!iso) return;

    document.querySelectorAll(".timeCell--selected")
      .forEach(x=>x.classList.remove("timeCell--selected"));
    cell.classList.add("timeCell--selected");

    state.selectedStartAt = iso;
    buildConfirm();
    gotoStep(3);
  };
}

function findIsoByDateAndTime_(isos, dateObj, timeStr){
  // isos: ["2026-02-22T10:00:00+09:00", ...]
  const y = dateObj.getFullYear();
  const m = String(dateObj.getMonth()+1).padStart(2,"0");
  const d = String(dateObj.getDate()).padStart(2,"0");
  const prefix = `${y}-${m}-${d}T${timeStr}:`;
  return isos.find(s => String(s).startsWith(prefix)) || null;
}


// ★ 追加：固定時間軸を生成（CONFIGと合わせるなら、この値もサーバから返してもOK）
const BUSINESS_OPEN = "09:00";
const BUSINESS_CLOSE = "18:00";
const GRAN_MIN = 10; // CONFIGの granularity_min と合わせる

function buildTimes_(openStr, closeStr, granMin){
  const [oh, om] = openStr.split(":").map(Number);
  const [ch, cm] = closeStr.split(":").map(Number);
  const start = oh * 60 + om;
  const end = ch * 60 + cm;

  const out = [];
  for (let t = start; t < end; t += granMin) {
    const hh = String(Math.floor(t / 60)).padStart(2,"0");
    const mm = String(t % 60).padStart(2,"0");
    out.push(`${hh}:${mm}`);
  }
  return out;
}


