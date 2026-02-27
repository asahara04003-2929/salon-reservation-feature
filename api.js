/**
 * ============================
 * GAS_URL, LINE公式アカウント（LIFF）
 * ============================
 */
const env = window.__ENV__;
if (!env?.GAS_URL || !env?.LIFF_ID) {
  alert("config.js が読み込まれていません（GAS_URL / LIFF_ID 未設定）");
  throw new Error("Missing __ENV__");
}

const { GAS_URL, LIFF_ID } = env;

/**
 * API：GET/POST
 */
async function apiGet(params) {
  showLoading("読み込み中…");
  try{
    const url = new URL(GAS_URL);
    Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, v));
    const res = await fetch(url.toString(), { method: "GET" });
    return await res.json();
  } finally {
    hideLoading();
  }
}

async function apiPost(payload) {
  showLoading("処理中…");
  try{
    const res = await fetch(GAS_URL, {
      method: "POST",
      body: JSON.stringify(payload),
    });
    const text = await res.text();
    try { return JSON.parse(text); }
    catch { return { ok:false, error:"INVALID_JSON_RESPONSE", raw:text }; }
  } finally {
    hideLoading();
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
  nickName: null,   // ★追加
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
 * 初期化：静的Webとして起動
 */
document.addEventListener("DOMContentLoaded", () => {
  initLiff();
});


/**
 * ============================
 * LINE公式アカウント（LIFF）
 * ============================
 */
async function initLiff() {
  wireEvents();

  try {
    setError("");

    // ✅ ローカル開発：LIFFもログインも全部スキップ
    if (isDevMode_()) {
      setStatus("開発モード（LIFFスキップ）");

      // ダミーのユーザー情報（必要なら localStorage で変えられる）
      state.lineUserId = localStorage.getItem("DEV_LINE_USER_ID") || "DEV_USER_001";
      state.nickName   = localStorage.getItem("DEV_NICKNAME") || "Dev User";

      await loadPlans();
      await checkUser();

      setStatus("準備完了（開発モード）");
      return;
    }

    // ===== 本番（LIFF） =====
    setStatus("LIFF初期化中…");
    await liff.init({ liffId: LIFF_ID });

    // LINEアプリ内必須
    if (!liff.isInClient()) {
      setStatus("LINEアプリ内で開いてください（LIFF）");
      setError("このページはLINEアプリ内（LIFF）でのみ利用できます。");
      return;
    }

    // 未ログインならログインへ
    if (!liff.isLoggedIn()) {
      liff.login();
      return;
    }

    const profile = await liff.getProfile();
    state.lineUserId = profile.userId;
    state.nickName   = profile.displayName || "";

    setStatus("ユーザー確認中…");
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
  initDobPicker(); // ★追加：生年月日ドラムロール
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
  setError("");

  const fd = new FormData($("registerForm"));
  const payload = {
    action: "users_upsert",
    line_user_id: state.lineUserId,
    nick_name: state.nickName || "",
    name: String(fd.get("name") || "").trim(),
    kana: String(fd.get("kana") || "").trim(),
    birthday: String(fd.get("birthday") || "").trim(), // YYYY-MM-DD
    gender: String(fd.get("gender") || "").trim(),     // male/female/other
    phone: String(fd.get("phone") || "").replace(/[^0-9]/g, "").trim(),
    email: String(fd.get("email") || "").trim(),
  };

  // 必須チェック
  if (!payload.birthday) {
    setError("生年月日を選択してください。");
    setStatus("入力を確認");
    return;
  }
  if (!payload.gender) {
    setError("性別を選択してください。");
    setStatus("入力を確認");
    return;
  }

  setStatus("登録中…");

  try {
    const r = await apiPost(payload);
    if (!r.ok) {
      // ★ 性別NGのハンドリング
      if (r.error === "GENDER_NOT_ALLOWED") {
        const map = { male: "男性", female: "女性", other: "その他" };
        const allowed = Array.isArray(r.allowed_genders) ? r.allowed_genders : [];
        const allowedJa = allowed.map(x => map[x] || x).join("、");

        const msg = allowed.length
          ? `このサロンでは登録できる性別が制限されています。\n許可：${allowedJa}\n\nお手数ですが選択を変更してください。`
          : `このサロンでは登録できる性別が制限されています。\nお手数ですが選択を変更してください。`;

        setStatus("登録できません");
        setError(msg);
        await popupError(escHtml(msg).replace(/\n/g, "<br>"), "登録不可");
        return;
      }

      throw new Error(r.error || "register_failed");
    }

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
        <div class="planInfo">
          <div class="item__title">${escapeHtml(plan.plan_name)}</div>
          <div class="item__meta">所要時間: ${Number(plan.duration_min)}分 / 料金: ¥${Number(plan.price).toLocaleString()}</div>
          ${plan.descriptions ? `<div class="planDesc">${escapeHtmlWithBreaks(clamp100_(plan.descriptions))}</div>` : ``}
        </div>

        <label class="row" style="gap:8px; flex-shrink:0;">
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

// async function loadAvailability() {
//   setError("");

//   if (!state.selectedPlan) return;
//   if (!state.selectedDate) return;

//   setStatus("空き枠を取得中…");
//   $("slotGrid").innerHTML = "";
//   $("slotHint").textContent = "";

//   try {
//     const r = await apiGet({
//       action: "availability",
//       date: state.selectedDate,
//       plan_id: state.selectedPlan.plan_id,
//     });
//     if (!r.ok) throw new Error(r.error || "availability_failed");

//     state.availableSlots = r.available || [];
//     $("slotHint").textContent = r.slot_source_hint || "※ 枠情報を参照しています";

//     renderSlots();
//     setStatus("日時を選択してください");
//   } catch (e) {
//     console.error(e);
//     setStatus("空き枠取得に失敗");
//     setError(String(e?.message || e));
//   }
// }

// function renderSlots() {
//   const grid = $("slotGrid");
//   grid.innerHTML = "";

//   const slots = state.availableSlots || [];
//   if (!slots.length) {
//     grid.innerHTML = `<div class="alert">この日は予約枠がありません（受付不可、または満席/受付終了の可能性）。</div>`;
//     return;
//   }

//   slots.forEach(iso => {
//     const btn = document.createElement("div");
//     btn.className = "slot";
//     btn.textContent = fmtTime(iso);
//     btn.addEventListener("click", () => {
//       state.selectedStartAt = iso;
//       [...grid.children].forEach(c => c.classList.remove("slot--selected"));
//       btn.classList.add("slot--selected");
//       buildConfirm();
//       gotoStep(3);
//     });
//     grid.appendChild(btn);
//   });
// }

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
      const ok = await popupConfirm("この予約をキャンセルしますか？");
      if (!ok) return;

      const result = await cancelReservation(resv.cancel_token);
      if (result.refreshed) await openMyPage(); // ←成功時だけ更新
    });

    root.appendChild(el);
  });
}

async function cancelReservation(cancelToken) {
  setError("");
  setStatus("キャンセル処理中…");

  try {
    const r = await apiPost({ action: "cancel", cancel_token: cancelToken });
    const ok = (r.ok === true) || (String(r.ok).toLowerCase() === "true");

    if (!ok) {
      if (r.error === "SAME_DAY_CANCEL_NOT_ALLOWED") {
        const phone = r.admin_phone || "";
        const msg = `当日のキャンセルについては、${phone}へご連絡ください。`;
        const msgHtml = `当日のキャンセルについては、以下電話番号に直接ご連絡ください。<br><br>電話番号：${escHtml(phone)}`;

        setStatus("当日のキャンセルは承っておりません。");
        setError(msg);         // ページ内にも残す（任意）
        await popupSameDayCancel(phone, msgHtml, "当日キャンセル不可"); // ★モーダル
        return { refreshed: false };
      }
      throw new Error(r.error || "cancel_failed");
    }

    setStatus("キャンセルしました");
    return { handled: true, refreshed: true };
  } catch (e) {
    console.error(e);
    setStatus("キャンセルに失敗");
    setError(String(e?.message || e));
    return { handled: false, refreshed: false };
  }
}

/**
 * XSS safe
 */

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

// ★追加：minutes -> "HH:mm"
function minToHHMM_(min){
  const hh = String(Math.floor(min / 60)).padStart(2,"0");
  const mm = String(min % 60).padStart(2,"0");
  return `${hh}:${mm}`;
}

// ★追加：interval差集合（[start,end) minutes）
function subtractIntervals_(windows, busies){
  const w = (windows || []).slice().sort((a,b)=>a[0]-b[0]);
  const b = (busies || []).slice().sort((a,b)=>a[0]-b[0]);
  const out = [];

  for (const [ws,we] of w){
    let cur = ws;
    for (const [bs,be] of b){
      if (be <= cur) continue;
      if (bs >= we) break;
      if (bs > cur) out.push([cur, Math.min(bs, we)]);
      cur = Math.max(cur, be);
      if (cur >= we) break;
    }
    if (cur < we) out.push([cur, we]);
  }
  return out;
}

// ★追加：free intervals -> start minutes list（粒度・所要時間を満たす枠）
function buildStartMins_(freeIntervals, granMin, requiredMin, minStartMin){
  const out = [];
  for (const [fs,fe] of freeIntervals){
    // 粒度で切り上げ
    let t = Math.ceil(fs / granMin) * granMin;

    // 今日制限（minStartMinがある日だけ）
    if (typeof minStartMin === "number") t = Math.max(t, Math.ceil(minStartMin / granMin) * granMin);

    for (; t + requiredMin <= fe; t += granMin) out.push(t);
  }
  return out;
}

async function loadAvailabilityWeek() {
  setError("");

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
      action: "availability_range_materials",   // ★ここが変更点
      from: ymd(state.weekStart),
      days: "7",
      duration_min: String(state.totalDurationMin),
    });

    console.log("availability_range_materials response:", r);
    if (!r.ok) throw new Error(r.error || JSON.stringify(r));

    state.granMin = r.granularity_min;
    state.businessOpen = r.business_open;
    state.businessClose = r.business_close;
    $("slotHint") && ($("slotHint").textContent = "");

    const dayResults = [...Array(7)].map((_, i) => {
      const d = addDays(state.weekStart, i);
      const key = ymd(d);

      const windows = (r.windows_by_date && r.windows_by_date[key]) ? r.windows_by_date[key] : [];
      const busy    = (r.busy_by_date && r.busy_by_date[key]) ? r.busy_by_date[key] : [];
      const minStartMin = (r.min_start_min_by_date) ? r.min_start_min_by_date[key] : null;

      const free = subtractIntervals_(windows, busy);
      const starts = buildStartMins_(free, Number(r.granularity_min), Number(r.required_duration_min), minStartMin);

      // renderWeekTable は iso を探す実装になってるので、
      // いったん "YYYY-MM-DDTHH:mm:00+09:00" 形式の疑似ISO文字列にして渡す
      // （最小修正で済ませるため）
      const minStartIso = new Date(Date.now() + Number(r.granularity_min) * 60000);

      const pseudoIsos = starts
        .map(min => {
          const hhmm = minToHHMM_(min);
          return `${key}T${hhmm}:00+09:00`;
        })
        .filter(iso => new Date(iso).getTime() > minStartIso.getTime()); // ★ここ

      return { date: d, slots: pseudoIsos };
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


/********************************************************************
 * ✅ JS側（予約画面：プラン一覧に説明文を表示、100文字で崩れない）
 * 1) renderPlans() の plan行HTMLを差し替え（下のブロック）
 * 2) 追加CSS（style.css か <style>）を貼る
 ********************************************************************/

/** 100文字までで表示（それ以上は…にする） */
function clamp100_(s) {
  const t = String(s ?? "").trim(); // ← 改行は残す
  return t.length > 100 ? (t.slice(0, 100) + "…") : t;
}

/** 改行も軽く整形して出したい場合（任意） */
function escapeHtml(s) {
  return String(s ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function escapeHtmlWithBreaks(s) {
  // 1) いったん全部エスケープ（<br>も文字として扱われる）
  let t = escapeHtml(s);

  // 2) \n を <br> に変換（Windows改行にも対応）
  t = t.replace(/\r\n|\r|\n/g, "<br>");

  // 3) ユーザーが <br> と書いたものも改行として扱いたい場合：
  //    エスケープ後は &lt;br&gt; / &lt;br/&gt; / &lt;br /&gt; になっているので、それだけ復活
  t = t
    .replace(/&lt;br\s*\/?&gt;/gi, "<br>");

  return t;
}


function showLoading(msg){
  const ov = $("loadingOverlay");
  const tx = $("loadingText");
  if (tx) tx.textContent = msg || "読み込み中…";
  if (ov) ov.classList.remove("hidden");
}
function hideLoading(){
  const ov = $("loadingOverlay");
  if (ov) ov.classList.add("hidden");
}



function isDevMode_() {
  // 例: http://localhost:5500/?dev=1
  const p = new URLSearchParams(location.search);
  return p.get("dev") === "1";
}

// モーダルライブラリ表示（確認）
async function popupConfirm(msg, title="確認") {
  if (!window.Swal) return confirm(msg);
  const r = await Swal.fire({
    icon: "question",
    title,
    text: msg,
    showCancelButton: true,
    confirmButtonText: "はい",
    cancelButtonText: "やめる",
  });
  return r.isConfirmed;
}

// モーダルライブラリ表示（アラート）
function popupError(msgHtml, title = "エラー") {
  // SweetAlert2 読み込み前に呼ばれた場合の保険
  if (!window.Swal) {
    alert(`${title}\n\n${msgHtml}`);
    return;
  }
  return Swal.fire({
    icon: "error",
    title,
    html: msgHtml,
    confirmButtonText: "OK",
  });
}


function normalizeTel_(phone){
  // 数字と+だけ残す（ハイフン等除去）
  return String(phone || "").replace(/[^\d+]/g, "");
}

async function popupSameDayCancel(phone, msgHtml, title = "エラー"){
  const tel = normalizeTel_(phone);
  const r = await Swal.fire({
    icon: "error",
    title,
    html: msgHtml,
    showCancelButton: true,
    confirmButtonText: "電話する",
    cancelButtonText: "閉じる",
  });

  if (r.isConfirmed && tel) {
    // ✅ ユーザー操作（ボタン押下）なので発信がブロックされにくい
    window.location.href = `tel:${tel}`;
  }
}

// HTMLコードのエスケープ
function escHtml(s){
  return String(s ?? "")
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;")
    .replaceAll("'","&#039;");
}

// 生年月日ドラムロールUI
function initDobPicker(){
  const birthDisplay = document.getElementById("birthDisplay");
  const birthValue   = document.getElementById("birthValue");

  const modal  = document.getElementById("dobModal");
  const wYear  = document.getElementById("wheelYear");
  const wMonth = document.getElementById("wheelMonth");
  const wDay   = document.getElementById("wheelDay");

  const btnCancel = document.getElementById("dobCancel");
  const btnDone   = document.getElementById("dobDone");

  if (!birthDisplay || !birthValue || !modal || !wYear || !wMonth || !wDay) return;

  const pad2 = (n) => String(n).padStart(2,"0");
  const daysInMonth = (y,m) => new Date(y, m, 0).getDate(); // m:1-12

  const buildWheel = (el, values, formatter) => {
    el.innerHTML = values.map(v => `<div class="wheelItem" data-value="${v}">${formatter(v)}</div>`).join("");
  };

  const snapToNearest = (el) => {
    const first = el.querySelector(".wheelItem");
    const itemH = first?.offsetHeight || 48;

    // いま中央に最も近い要素を探して、その位置へスナップ
    const center = el.scrollTop + el.clientHeight / 2;
    const items = [...el.querySelectorAll(".wheelItem")];
    if (!items.length) return;

    let best = items[0];
    let bestDist = Infinity;
    for (const it of items) {
      const itCenter = it.offsetTop + (it.offsetHeight / 2);
      const dist = Math.abs(itCenter - center);
      if (dist < bestDist) { bestDist = dist; best = it; }
    }
    const top = best.offsetTop - (el.clientHeight / 2 - itemH / 2);
    el.scrollTo({ top: Math.max(0, top), behavior: "smooth" });
  };

  const getSelectedValue = (el) => {
    const center = el.scrollTop + el.clientHeight / 2;
    const items = [...el.querySelectorAll(".wheelItem")];
    if (!items.length) return null;

    let best = null;
    let bestDist = Infinity;
    for (const it of items) {
      const itCenter = it.offsetTop + (it.offsetHeight / 2);
      const dist = Math.abs(itCenter - center);
      if (dist < bestDist) { bestDist = dist; best = it; }
    }
    return best ? Number(best.dataset.value) : null;
  };

  const setSelectedValue = (el, value) => {
    const items = [...el.querySelectorAll(".wheelItem")];
    const item = items.find(x => Number(x.dataset.value) === Number(value));
    if (!item) return;

    const itemH = item.offsetHeight || 48;
    // itemの中心がホイール中央に来るようにscrollTopを計算
    const top = item.offsetTop - (el.clientHeight / 2 - itemH / 2);
    el.scrollTop = Math.max(0, top);
  };

  const rebuildDays = (y, m, keepDay) => {
    const max = daysInMonth(y, m);
    const days = Array.from({length:max}, (_,i)=>i+1);
    buildWheel(wDay, days, d => `${d}日`);
    const d = Math.min(keepDay ?? 1, max);
    setSelectedValue(wDay, d);
  };

  let timer = null;
  const onWheelScroll = (el, cb) => {
    el.addEventListener("scroll", () => {
      clearTimeout(timer);
      timer = setTimeout(() => {
        snapToNearest(el);
        cb?.();
      }, 80);
    }, { passive:true });
  };

  const openPicker = (initial) => {
    const now = new Date();
    const years = Array.from({length:(now.getFullYear()-1900+1)}, (_,i)=>1900+i);
    buildWheel(wYear, years, y => `${y}年`);
    buildWheel(wMonth, Array.from({length:12},(_,i)=>i+1), m => `${m}月`);

    const init = initial ?? { y:2000, m:1, d:1 };
    rebuildDays(init.y, init.m, init.d);

    // ✅ 先に表示（display:none解除してからスクロール計算）
    modal.classList.remove("hidden");

    // ✅ レイアウト確定後に初期位置をセット（超重要）
    requestAnimationFrame(() => {
      setSelectedValue(wYear, init.y);
      setSelectedValue(wMonth, init.m);
      setSelectedValue(wDay, init.d);
    });
  };

  const closePicker = () => modal.classList.add("hidden");

  birthDisplay.addEventListener("click", () => {
    const v = birthValue.value; // YYYY-MM-DD
    if (v) {
      const [y,m,d] = v.split("-").map(Number);
      openPicker({ y, m, d });
    } else {
      openPicker({ y:2000, m:1, d:1 });
    }
  });

  btnCancel.addEventListener("click", closePicker);

  btnDone.addEventListener("click", () => {
    const y = getSelectedValue(wYear);
    const m = getSelectedValue(wMonth);
    const d = getSelectedValue(wDay);
    if (!y || !m || !d) return;

    birthValue.value = `${y}-${pad2(m)}-${pad2(d)}`;      // 送信用（FormDataで拾う）
    birthDisplay.value = `${y}/${pad2(m)}/${pad2(d)}`;    // 表示用
    closePicker();
  });

  modal.addEventListener("click", (e) => {
    if (e.target === modal) closePicker();
  });
}