// Retirement Planner page: inputs, results, chart and report.
// Calculations live in retire-calc.js (ported unchanged from Jake's calculator).

const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const money = (v) => new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(v);
const pctFixed = (v) => `${v.toFixed(2)}%`;
const wholeMoney = (v) => v.toLocaleString("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 });

// Each report request (name, email, phone, inputs and results) is recorded
// in Bill's own Google Sheet through this Apps Script web app.
const SHEET_URL = "https://script.google.com/macros/s/AKfycbwL1E5LhVh5y6iNWz9yVwe4RROwCvcS-rmwCsVWtwifn9FavGg6welhjE8tyvoTSYWkHQ/exec";
const BOOKING_URL = "https://outlook.office.com/bookwithme/user/233fafd30bb44d579106fe231f8b1cd6@capitalbenefitsinc.com?anonymous&ismsaljsauthenabled&ep=pcard";
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const ACCOUNTS = ["brokerageAccount", "traditionalIraAccount", "rothIraAccount", "account401k", "otherAccounts"];
const PIE = [
  ["Retirement balance", "retirementBalance", "#2a78d6"],
  ["Final year income", "lastYearIncome", "#eb6834"],
  ["Annual expenses", "annualRetirementExpenses", "#1baf7a"],
  ["Social Security", "annualSocialSecurity", "#eda100"],
];

const state = { inputs: structuredClone(RETIREMENT_DEFAULTS), results: null, reportShown: false, formSubmitted: false, charts: {} };

// ---------------------------------------------------------------- form

const FIELDS = {
  "g-personal": [
    { key: "name", type: "text", label: "Full name", required: true, placeholder: "Enter your full name", tip: "Your full legal name for the retirement plan." },
    { key: "email", type: "email", label: "Email address", required: true, placeholder: "Enter your email address", tip: "Your email address for receiving retirement plan updates and reports.", error: "Please enter a valid email address" },
    { key: "phoneNumber", type: "tel", label: "Phone number", placeholder: "Enter your phone number (optional)", tip: "Your phone number for contact purposes. This field is optional." },
  ],
  "g-plan": [
    { key: "currentAge", label: "Current age", min: 14, max: 90, required: true, tip: "Your current age." },
    { key: "retirementAge", label: "Age at retirement", min: 10, max: 90, required: true, tip: "Age at which you plan to retire. This calculator assumes that the year you retire, you do not make any contributions to your retirement savings." },
    { key: "annualIncome", label: "Annual household income", min: 1, max: 1e7, format: "currency", required: true, tip: "Your total household income. If you are married, this should include your spouse's income." },
    { key: "percentToSave", label: "Percent of income to save", min: 0, max: 100, format: "percent", required: true, tip: "The percentage of your annual income you plan to contribute to your retirement savings." },
    { key: "brokerageAccount", label: "Brokerage account", min: 0, max: 1e8, format: "currency", tip: "Amount in your taxable brokerage account." },
    { key: "traditionalIraAccount", label: "Traditional IRA account", min: 0, max: 1e8, format: "currency", tip: "Amount in your Traditional IRA account." },
    { key: "rothIraAccount", label: "Roth IRA account", min: 0, max: 1e8, format: "currency", tip: "Amount in your Roth IRA account." },
    { key: "account401k", label: "401k account", min: 0, max: 1e8, format: "currency", tip: "Amount in your 401k account." },
    { key: "otherAccounts", label: "Other accounts", min: 0, max: 1e8, format: "currency", tip: "Amount in other retirement accounts." },
    { key: "currentSavings", label: "Current retirement savings", min: 0, max: 1e8, format: "currency", required: true, disabled: true, tip: "Total amount that you currently have saved toward your retirement. This is automatically calculated from the accounts above." },
    { key: "expectedIncomeIncrease", label: "Expected income increase", min: 0, max: 20, format: "percent", required: true, tip: "Annual percent increase you expect in your household income." },
    { key: "preRetirementIncomeDesired", label: "Desired retirement income", min: 40, max: 160, format: "percent", required: true, tip: "The percentage of your pre-retirement household income you think you will need in retirement." },
    { key: "yearsOfRetirementIncome", label: "Years of retirement income", min: 1, max: 100, required: true, tip: "Total number of years you expect to use your retirement income." },
  ],
  "g-returns": [
    { key: "returnBeforeRetirement", label: "Rate of return before retirement", min: 0, max: 20, format: "percent", required: true, tip: "This is the annual rate of return you expect from your retirement savings and investments before retirement." },
    { key: "returnDuringRetirement", label: "Rate of return during retirement", min: 0, max: 20, format: "percent", required: true, tip: "This is the annual rate of return you expect from your savings and investments during retirement." },
    { key: "expectedInflation", label: "Expected rate of inflation", min: 0, max: 20, format: "percent", required: true, tip: "This is what you expect for the average long-term inflation rate." },
    { key: "isMarried", type: "checkbox", label: "Married", tip: "Check this box if you are married. Married couples have a higher maximum Social Security benefit." },
    { key: "includeSocialSecurity", type: "checkbox", label: "Include Social Security", tip: "Check this box if you wish to include Social Security benefits in your retirement planning." },
  ],
};

// "1234567.5" -> "1,234,567.5" (digits after the decimal point are left alone).
const withCommas = (str) => {
  const [whole, ...rest] = String(str).split(".");
  const grouped = whole.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return rest.length ? `${grouped}.${rest.join("")}` : grouped;
};
const display = (f, v) => (f && f.format === "currency" ? withCommas(String(v)) : String(v));
const FIELD_BY_KEY = Object.fromEntries(Object.values(FIELDS).flat().map((f) => [f.key, f]));

const rangeHint = (f) =>
  f.format === "currency" ? `$${f.min.toLocaleString()} - $${f.max.toLocaleString()}` : f.format === "percent" ? `${f.min}% - ${f.max}%` : `${f.min} - ${f.max}`;

function fieldHTML(f) {
  const id = `f-${f.key}`;
  const tip = f.tip ? `<button type="button" class="tip-btn" aria-expanded="false" aria-controls="${id}-tip" aria-label="About ${esc(f.label)}">?</button>` : "";
  const tipBox = f.tip ? `<div class="tip" id="${id}-tip" role="tooltip" hidden>${esc(f.tip)}</div>` : "";
  if (f.type === "checkbox") {
    return `<div class="field check" data-key="${f.key}"><input type="checkbox" id="${id}"><label for="${id}">${esc(f.label)}</label>${tip}${tipBox}</div>`;
  }
  const req = f.required ? `<span class="req" aria-hidden="true">*</span>` : "";
  const label = `<div class="row"><label for="${id}">${esc(f.label)}${req}</label>${tip}</div>`;
  if (["text", "email", "tel"].includes(f.type)) {
    return `<div class="field" data-key="${f.key}">${label}<div class="box"><input id="${id}" type="${f.type}" placeholder="${esc(f.placeholder || "")}" autocomplete="${f.type === "email" ? "email" : f.type === "tel" ? "tel" : "name"}"></div><span class="err" hidden></span>${tipBox}</div>`;
  }
  const pre = f.format === "currency" ? `<span>$</span>` : "";
  const suf = f.format === "percent" ? `<span class="suf">%</span>` : "";
  return `<div class="field" data-key="${f.key}">${label}<div class="box${f.disabled ? " disabled" : ""}">${pre}<input id="${id}" type="text" inputmode="decimal" ${f.disabled ? "disabled" : ""}>${suf}</div><span class="hint">Enter: ${rangeHint(f)}</span>${tipBox}</div>`;
}

function buildForm() {
  for (const [group, fields] of Object.entries(FIELDS)) $(group).innerHTML = fields.map(fieldHTML).join("");
  const all = Object.values(FIELDS).flat();
  for (const f of all) {
    const input = $(`f-${f.key}`);
    if (f.type === "checkbox") {
      input.checked = state.inputs[f.key];
      input.addEventListener("change", () => setInput(f.key, input.checked));
    } else if (["text", "email", "tel"].includes(f.type)) {
      input.value = state.inputs.userInfo[f.key];
      if (f.type === "tel") input.addEventListener("input", (ev) => formatPhoneBox(input, ev));
      input.addEventListener("input", () => { state.inputs.userInfo[f.key] = input.value; validateField(f, false); });
      input.addEventListener("blur", () => { input.dataset.touched = "1"; validateField(f, false); });
    } else {
      input.value = display(f, state.inputs[f.key]);
      // Same behavior as the original: keep digits and dots; apply the value
      // as soon as it's in range; snap to the range when leaving the box.
      // Dollar boxes also show thousands commas while typing.
      input.addEventListener("input", () => {
        const caret = input.selectionStart ?? input.value.length;
        const kept = input.value.slice(0, caret).replace(/[^0-9.]/g, "").length;
        const cleaned = input.value.replace(/[^0-9.]/g, "");
        input.value = f.format === "currency" ? withCommas(cleaned) : cleaned;
        if (f.format === "currency" && document.activeElement === input) {
          let pos = 0, seen = 0;
          while (pos < input.value.length && seen < kept) { if (input.value[pos] !== ",") seen++; pos++; }
          input.setSelectionRange(pos, pos);
        }
        const v = parseFloat(cleaned);
        if (!isNaN(v) && v >= f.min && v <= f.max) setInput(f.key, v, false);
      });
      input.addEventListener("blur", () => {
        const v = parseFloat(input.value.replace(/,/g, ""));
        const next = isNaN(v) || v < f.min ? f.min : v > f.max ? f.max : v;
        input.value = display(f, next);
        setInput(f.key, next, false);
      });
    }
  }
  document.querySelectorAll(".tip-btn").forEach((b) => {
    const box = $(b.getAttribute("aria-controls"));
    const show = (on) => { box.hidden = !on; b.setAttribute("aria-expanded", String(on)); };
    b.addEventListener("mouseenter", () => show(true));
    b.addEventListener("mouseleave", () => show(false));
    b.addEventListener("click", () => show(box.hidden));
  });
}

// Phone box: shows (732) 555-1234 as digits are typed. Numbers starting
// with + or 1, or longer than 10 digits, are left exactly as typed.
let lastPhoneDigits = "";
function formatPhoneBox(input, ev) {
  const raw = input.value;
  let digits = raw.replace(/\D/g, "");
  if (/^\s*\+/.test(raw) || digits[0] === "1" || digits.length > 10) { lastPhoneDigits = digits; return; }
  const caret = input.selectionStart ?? raw.length;
  let before = raw.slice(0, caret).replace(/\D/g, "").length;
  // Backspace over ")", " " or "-" removes the digit in front of it.
  if (ev && ev.inputType === "deleteContentBackward" && digits === lastPhoneDigits && before > 0) {
    digits = digits.slice(0, before - 1) + digits.slice(before);
    before -= 1;
  }
  const a = digits.slice(0, 3), b = digits.slice(3, 6), c = digits.slice(6);
  input.value = !digits ? "" : digits.length <= 3 ? `(${a}` : digits.length <= 6 ? `(${a}) ${b}` : `(${a}) ${b}-${c}`;
  lastPhoneDigits = digits;
  if (document.activeElement === input) {
    let pos = 0, seen = 0;
    while (pos < input.value.length && seen < before) { if (/\d/.test(input.value[pos])) seen++; pos++; }
    if (before === 0) pos = digits ? 1 : 0;
    input.setSelectionRange(pos, pos);
  }
}

function validateField(f, force) {
  const input = $(`f-${f.key}`);
  const wrap = input.closest(".field");
  const v = input.value;
  const show = force || state.formSubmitted || input.dataset.touched === "1";
  let message = "";
  if (f.required && !v.trim()) message = `${f.label} is required`;
  else if (f.type === "email" && v && !EMAIL_RE.test(v)) message = f.error || "Please enter a valid value";
  const bad = show && !!message;
  wrap.classList.toggle("has-error", bad);
  const err = wrap.querySelector(".err");
  err.hidden = !bad;
  err.textContent = bad ? message : "";
  return !message;
}

function setInput(key, value, syncBox = true) {
  state.inputs[key] = value;
  if (ACCOUNTS.includes(key)) {
    state.inputs.currentSavings = ACCOUNTS.reduce((a, k) => a + state.inputs[k], 0);
    $("f-currentSavings").value = withCommas(String(state.inputs.currentSavings));
  }
  if (syncBox && $(`f-${key}`) && $(`f-${key}`).type === "text") $(`f-${key}`).value = display(FIELD_BY_KEY[key], value);
  recalc();
}

// ---------------------------------------------------------------- results

function recalc() {
  state.results = calculateRetirement(state.inputs);
  renderResults();
  renderChart();
  if (state.reportShown) renderReport();
}

function renderResults() {
  const e = state.inputs, t = state.results;
  const goalAge = e.retirementAge + e.yearsOfRetirementIncome;
  const lasts = t.savingsRunOutAge >= goalAge;
  const peak = Math.max(...t.yearlyData.map((r) => r.endingBalance));

  $("banner").className = `banner ${t.savingsRunOutAge < goalAge ? "warn" : "ok"}`;
  $("banner").textContent = t.savingsRunOutAge < goalAge ? "You may need to save more." : "Your retirement plan looks healthy.";

  const total = PIE.reduce((a, [, k]) => a + t[k], 0);
  const showPie = peak > 0 && total > 0;
  $("results").innerHTML = `
    <h2>Results summary</h2>
    <p class="verdict ${lasts ? "ok" : "warn"}">${lasts ? `Your retirement savings will last through age ${goalAge}.` : `Retirement savings runs out at age ${t.savingsRunOutAge}.`}</p>
    ${showPie ? `<div class="rp-pie"><div class="pie-box"><canvas id="pie" aria-label="Results breakdown"></canvas></div>
      <ul class="rp-legend">${PIE.map(([label, k, color]) => `<li><span class="sw" style="background:${color}"></span>${label}<span class="v">${money(t[k])}</span></li>`).join("")}</ul></div>`
      : `<p class="note">No data available for chart</p>`}
    <p class="detail">Your plan provides <strong>${money(t.retirementBalance)}</strong> when you retire. This assumes annual retirement expenses of <strong>${money(t.annualRetirementExpenses)}</strong>, which is ${e.preRetirementIncomeDesired}% of your last year's income of <strong>${money(t.lastYearIncome)}</strong>. This includes <strong>${money(t.annualSocialSecurity)}</strong> per year from Social Security.</p>
    <div class="stats">
      <div><b>${e.retirementAge - e.currentAge}</b>Years until retirement</div>
      <div><b>${money(t.lastYearIncome)}</b>Final year income</div>
      <div><b>${money(t.retirementBalance)}</b>Retirement balance</div>
      <div><b>${money(t.annualRetirementExpenses)}</b>Annual expenses</div>
    </div>`;
  if (showPie) drawPie("pie", t, total);
}

function drawPie(id, t, total) {
  if (state.charts[id]) state.charts[id].destroy();
  state.charts[id] = new Chart($(id), {
    type: "pie",
    data: { labels: PIE.map(([l]) => l), datasets: [{ data: PIE.map(([, k]) => t[k]), backgroundColor: PIE.map(([, , c]) => c), borderColor: "#FFFFFF", borderWidth: 2 }] },
    options: {
      responsive: true, maintainAspectRatio: false, animation: false,
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: (c) => `${c.label}: ${money(c.parsed)} (${(c.parsed / total * 100).toFixed(1)}%)` } } },
    },
  });
}

function renderChart() {
  const t = state.results, e = state.inputs;
  const rows = t.yearlyData;
  const retireIdx = rows.findIndex((r) => r.age === e.retirementAge);
  const retireLine = {
    id: "retireLine",
    afterDatasetsDraw(chart) {
      if (retireIdx < 0) return;
      const x = chart.scales.x.getPixelForValue(retireIdx);
      const { ctx, chartArea: a } = chart;
      ctx.save();
      ctx.strokeStyle = "#102A43"; ctx.setLineDash([2, 3]); ctx.lineWidth = 1;
      ctx.beginPath(); ctx.moveTo(x, a.top); ctx.lineTo(x, a.bottom); ctx.stroke();
      ctx.setLineDash([]); ctx.fillStyle = "#102A43"; ctx.font = "500 12px 'Instrument Sans', system-ui, sans-serif";
      ctx.textAlign = x > a.right - 90 ? "right" : "left";
      ctx.fillText(`Retire at ${e.retirementAge}`, x + (x > a.right - 90 ? -6 : 6), a.top + 12);
      ctx.restore();
    },
  };
  if (state.charts.balance) state.charts.balance.destroy();
  state.charts.balance = new Chart($("balanceChart"), {
    type: "line",
    data: {
      labels: rows.map((r) => r.age),
      datasets: [
        // Same line styles as the home page chart: brass solid lead line, slate dashed second line.
        { label: "Savings for retirement", data: rows.map((r) => r.endingBalance), borderColor: "#B7791F", backgroundColor: "#B7791F", borderWidth: 2.5, pointRadius: 0, pointHoverRadius: 4, tension: 0.15 },
        { label: "Withdrawals for retirement", data: rows.map((r) => r.withdrawals), borderColor: "#52606D", backgroundColor: "#52606D", borderWidth: 1.5, borderDash: [5, 4], pointRadius: 0, pointHoverRadius: 4, tension: 0.15 },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false, animation: false,
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: { position: "bottom", align: "start", labels: { boxWidth: 14, boxHeight: 3, font: { family: "Instrument Sans" }, color: "#52606D" } },
        tooltip: { callbacks: { title: (c) => `Age ${c[0].label}`, label: (c) => `${c.dataset.label}: ${wholeMoney(c.parsed.y)}` } },
      },
      scales: {
        x: { grid: { display: false }, ticks: { maxTicksLimit: 8, maxRotation: 0, autoSkipPadding: 12, color: "#52606D", font: { family: "Instrument Sans" }, callback(v) { return `Age ${this.getLabelForValue(v)}`; } } },
        y: { beginAtZero: true, grid: { color: "#E6EBF0" }, ticks: { color: "#52606D", font: { family: "Instrument Sans" }, callback: (v) => wholeMoney(v) } },
      },
    },
    plugins: [retireLine],
  });
}

// ---------------------------------------------------------------- report

const personalFields = () => FIELDS["g-personal"];

function toggleReport() {
  if (state.reportShown) {
    state.reportShown = false;
    $("report").hidden = true;
    $("reportBtn").textContent = "View report";
    $("printBtn").hidden = true;
    $("reportNote").textContent = "";
    return;
  }
  state.formSubmitted = true;
  const ok = personalFields().map((f) => validateField(f, true)).every(Boolean);
  if (!ok) {
    $("reportNote").textContent = "Enter a name and a valid email address above to view the report.";
    $("f-name").closest("form").scrollIntoView({ behavior: matchMedia("(prefers-reduced-motion: reduce)").matches ? "auto" : "smooth" });
    return;
  }
  submitToSheet(state.inputs, state.results);
  state.reportShown = true;
  $("reportBtn").textContent = "Hide report";
  $("printBtn").hidden = false;
  $("reportNote").textContent = `A PDF copy of this report is on its way to ${state.inputs.userInfo.email}.`;
  renderReport();
  $("report").hidden = false;
  $("report").scrollIntoView({ behavior: matchMedia("(prefers-reduced-motion: reduce)").matches ? "auto" : "smooth", block: "start" });
}

function submitToSheet(p, g) {
  const payload = {
    timestamp: new Date().toISOString(), name: p.userInfo.name, email: p.userInfo.email, phone: p.userInfo.phoneNumber,
    currentAge: p.currentAge, retirementAge: p.retirementAge, annualIncome: p.annualIncome, percentToSave: p.percentToSave,
    currentSavings: p.currentSavings, brokerageAccount: p.brokerageAccount, traditionalIraAccount: p.traditionalIraAccount,
    rothIraAccount: p.rothIraAccount, account401k: p.account401k, otherAccounts: p.otherAccounts,
    expectedIncomeIncrease: p.expectedIncomeIncrease, preRetirementIncomeDesired: p.preRetirementIncomeDesired,
    yearsOfRetirementIncome: p.yearsOfRetirementIncome, returnBeforeRetirement: p.returnBeforeRetirement,
    returnDuringRetirement: p.returnDuringRetirement, expectedInflation: p.expectedInflation, isMarried: p.isMarried,
    includeSocialSecurity: p.includeSocialSecurity, retirementBalance: g.retirementBalance, lastYearIncome: g.lastYearIncome,
    annualRetirementExpenses: g.annualRetirementExpenses, annualSocialSecurity: g.annualSocialSecurity,
    savingsRunOutAge: g.savingsRunOutAge, endingBalance: g.endingBalance, site: location.hostname,
  };
  // The sheet's script records the request and emails the PDF report to the
  // client (copy to Bill). Re-opening the same report doesn't send it again.
  const { timestamp, ...same } = payload;
  const key = JSON.stringify(same);
  if (key === state.lastSent) return;
  state.lastSent = key;
  fetch(SHEET_URL, { method: "POST", mode: "no-cors", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) })
    .catch((err) => console.error("Failed to submit to Google Sheets:", err));
}

// Year-by-year table, split in two: working years, then retirement years.
// When printed, the first table starts page 2 and the retirement years start a new page.
function balanceTables(e, t) {
  // Print-only rows: a spacer and the section title repeat at the top of every
  // printed page (and a spacer at the bottom), so the margins don't depend on
  // the browser's print-margin setting.
  const head = (title) => `<thead><tr class="print-gap"><td colspan="7"></td></tr><tr class="print-title"><th colspan="7">${title}</th></tr><tr><th>Age</th><th class="num">Beginning retirement balance</th><th class="num">Investment growth</th><th class="num">Contributions at ${e.percentToSave}% of income</th><th class="num">Annual income desired</th><th class="num">Retirement account withdrawals</th><th class="num">Ending retirement balance</th></tr></thead><tfoot><tr class="print-gap"><td colspan="7"></td></tr></tfoot>`;
  const rowHTML = (r) => `<tr class="${r.age >= e.retirementAge ? "retire-year" : ""}">
          <td>${r.age}</td><td class="num">${money(r.beginningBalance)}</td><td class="num">${money(r.investmentGrowth)}</td><td class="num">${money(r.contributions)}</td>
          <td class="num">${money(r.annualIncomeDesired)}</td><td class="num">${money(r.withdrawals)}</td><td class="num">${money(r.endingBalance)}</td></tr>`;
  const table = (rows, title) => `<div class="table-scroll"><table>${head(title)}<tbody>${rows.map(rowHTML).join("")}</tbody></table></div>`;
  const working = t.yearlyData.filter((r) => r.age < e.retirementAge);
  const retired = t.yearlyData.filter((r) => r.age >= e.retirementAge);
  const parts = [];
  if (working.length) parts.push(`<section class="balances"><h4>Balances by year*: before retirement</h4>${table(working, "Balances by year*: before retirement")}</section>`);
  if (retired.length) parts.push(`<section class="balances"><h4>Balances by year*: retirement, from age ${e.retirementAge}</h4>${table(retired, `Balances by year*: retirement, from age ${e.retirementAge}`)}</section>`);
  return parts.join("");
}

function renderReport() {
  const e = state.inputs, t = state.results;
  const goalAge = e.retirementAge + e.yearsOfRetirementIncome;
  const lasts = t.savingsRunOutAge >= goalAge;
  const annualSave = e.annualIncome * (e.percentToSave / 100);
  const today = new Date().toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" });
  const peak = Math.max(...t.yearlyData.map((r) => r.endingBalance));
  const total = PIE.reduce((a, [, k]) => a + t[k], 0);
  const u = e.userInfo;
  const row = (a, b) => `<tr><td>${a}</td><td>${b}</td></tr>`;
  const sentence = `This assumes annual retirement expenses of ${money(t.annualRetirementExpenses)} which is ${e.preRetirementIncomeDesired}% of your last year's income of ${money(t.lastYearIncome)}.`;
  const contact = `
    Phone: <a href="tel:+17327845789">(732) 784-5789</a><br>
    E-mail: <a href="mailto:bippolito@alexandercapitallp.com">bippolito@alexandercapitallp.com</a><br>
    Website: <a href="https://alexandercapitalvm.com" target="_blank" rel="noopener noreferrer">alexandercapitalvm.com</a><br>
    <a href="${BOOKING_URL}" target="_blank" rel="noopener noreferrer">Book an appointment</a>`;

  $("report").innerHTML = `<article class="report">
    <div class="letterhead">
      <div>
        <div class="co">Alexander Capital L.P.</div>
        <div class="by">Prepared by<br><b>Bill Ippolito</b>, Financial Advisor<br>${contact}</div>
      </div>
      <div class="right">${today}<b>Retirement Planning Analysis</b></div>
    </div>
    <h3>Retirement savings and expenditures</h3>
    ${u.name ? `<p class="prepared"><strong>Prepared for:</strong> ${esc(u.name)}${u.email ? `, ${esc(u.email)}` : ""}${u.phoneNumber ? `, ${esc(u.phoneNumber)}` : ""}</p>` : ""}
    <p class="summary-box ${lasts ? "ok" : "warn"}">${lasts ? `Your retirement savings will last through age ${goalAge}.` : `Retirement savings runs out at age ${t.savingsRunOutAge}.`} ${sentence}</p>
    ${lasts ? "" : `<h4>To help meet your goal, you may wish to do one of the following:</h4>
    <ul>
      <li>Increase contributions to <strong>${pctFixed(t.recommendedSavingsRate)}</strong> of your income (${money(e.annualIncome * t.recommendedSavingsRate / 100)} annual savings, which will increase as your income increases)</li>
      <li>Increase your rate of return before retirement to <strong>${pctFixed(t.recommendedReturnRate)}</strong>.</li>
      <li>Reduce your required income in retirement to <strong>${pctFixed(t.recommendedIncomePercent)}</strong> of your final year's income.</li>
      <li>Delay your retirement until age <strong>${t.recommendedRetirementAge}</strong>.</li>
    </ul>`}
    <div class="report-grid">
      <div>
        <h4>Personal information</h4>
        <table class="info"><tbody>${row("Name", esc(u.name) || "Not provided")}${row("Email", esc(u.email) || "Not provided")}${row("Phone", esc(u.phoneNumber) || "Not provided")}</tbody></table>
        <h4>Retirement plan inputs</h4>
        <table class="info"><tbody>
          ${row("Current age", e.currentAge)}${row("Age at retirement", e.retirementAge)}${row("Household income", money(e.annualIncome))}
          ${row("Current retirement savings", money(e.currentSavings))}
          ${row("Percent of income to save", `${e.percentToSave}% (${money(annualSave)} annually which will increase as your income increases)`)}
          ${row("Expected income increase", `${e.expectedIncomeIncrease}%`)}${row("Years of retirement income", e.yearsOfRetirementIncome)}
          ${row("Desired retirement income", `${e.preRetirementIncomeDesired}%`)}
        </tbody></table>
      </div>
      <div>
        <h4>Investment returns and inflation</h4>
        <table class="info"><tbody>
          ${row("Rate of return before retirement", `${e.returnBeforeRetirement}%`)}${row("Rate of return during retirement", `${e.returnDuringRetirement}%`)}
          ${row("Expected inflation rate", `${e.expectedInflation}%`)}${row("Are you married?", e.isMarried ? "yes" : "no")}
          ${row("Include Social Security?", e.includeSocialSecurity ? "yes" : "no")}
        </tbody></table>
        <h4>Result summary</h4>
        <table class="info"><tbody>
          ${row("Years until retirement", e.retirementAge - e.currentAge)}${row("Your last year's income", money(t.lastYearIncome))}
          ${row("Estimated annual retirement expenditures", money(t.annualRetirementExpenses))}${row("Your ending balance", money(t.endingBalance))}
        </tbody></table>
        ${peak > 0 && total > 0 ? `<div class="rp-pie" style="margin-top:14px"><div class="pie-box"><canvas id="reportPie" aria-label="Results breakdown"></canvas></div>
          <ul class="rp-legend">${PIE.map(([label, k, color]) => `<li><span class="sw" style="background:${color}"></span>${label}<span class="v">${money(t[k])}</span></li>`).join("")}</ul></div>` : ""}
      </div>
    </div>
    ${balanceTables(e, t)}
    <p class="table-note">* This chart reflects fixed annual rates of return, expense, and income growth. It is meant solely as a guide. It is not likely that these and other variables will increase and decrease by a steady annual rate year over year.</p>
    <div class="report-foot">
      <b>Questions about your retirement plan?</b> Contact Bill Ippolito.<br>${contact}
      <p>Schedule a consultation to discuss your personalized retirement strategy and explore opportunities to optimize your financial future.</p>
      <b>Alexander Capital L.P.</b> Building Your Financial Future
    </div>
  </article>`;
  if (peak > 0 && total > 0) drawPie("reportPie", t, total);
}

// ---------------------------------------------------------------- start

buildForm();
recalc();
$("reportBtn").addEventListener("click", toggleReport);
$("printBtn").addEventListener("click", () => window.print());
$("form").addEventListener("submit", (ev) => ev.preventDefault());
