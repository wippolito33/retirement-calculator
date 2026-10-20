// Retirement Calculator: Google Apps Script for Bill Ippolito's leads sheet.
//
// Paste this whole file into the sheet's Extensions > Apps Script editor,
// replacing what's there, then redeploy (Deploy > Manage deployments > pencil >
// Version: New version > Deploy) so the web address stays the same.
//
// Each "View report" click on the website:
//   1. adds a row to the first tab of the sheet, and
//   2. emails the client a PDF of their report, with a copy (bcc) to Bill,
//      sent from Bill's work address through this Google account's Gmail.
//
// The calculation section at the bottom is an exact copy of retire-calc.js
// from the website. If that file changes, paste the new version over it.

// ===================================================================
// Web app: receives "View report" submissions from the website.
// ===================================================================

// Work address: shown as the sender (must be set up in Gmail under
// Settings > Accounts > "Send mail as") and copied on every report.
const ADVISOR_EMAIL = "bippolito@alexandercapitallp.com";
const ADVISOR_NAME = "Bill Ippolito";
const ADVISOR_PHONE = "(732) 784-5789";
const ADVISOR_WEBSITE = "alexandercapitalvm.com";
const BOOKING_URL = "https://outlook.office.com/bookwithme/user/233fafd30bb44d579106fe231f8b1cd6@capitalbenefitsinc.com?anonymous&ismsaljsauthenabled&ep=pcard";
const TIME_ZONE = "America/New_York";

// Safety limits, because the web app is open to anyone on the internet.
// A free Gmail account can email about 100 people a day from a script; each
// report goes to 2 (the client and you), so 40 reports a day stays well inside it.
const MAX_REPORTS_PER_DAY = 40;
const MAX_REPORTS_PER_ADDRESS = 3; // per client email address, per 6 hours

const FIELDS = [
  "timestamp", "name", "email", "phone",
  "currentAge", "retirementAge", "annualIncome", "percentToSave", "currentSavings",
  "brokerageAccount", "traditionalIraAccount", "rothIraAccount", "account401k", "otherAccounts",
  "expectedIncomeIncrease", "preRetirementIncomeDesired", "yearsOfRetirementIncome",
  "returnBeforeRetirement", "returnDuringRetirement", "expectedInflation",
  "isMarried", "includeSocialSecurity",
  "retirementBalance", "lastYearIncome", "annualRetirementExpenses",
  "annualSocialSecurity", "savingsRunOutAge", "endingBalance",
  "site", "emailStatus"
];

// Same limits as the boxes on the website.
const LIMITS = {
  currentAge: [14, 90], retirementAge: [10, 90], annualIncome: [1, 1e7], percentToSave: [0, 100],
  brokerageAccount: [0, 1e8], traditionalIraAccount: [0, 1e8], rothIraAccount: [0, 1e8], account401k: [0, 1e8], otherAccounts: [0, 1e8],
  expectedIncomeIncrease: [0, 20], preRetirementIncomeDesired: [40, 160], yearsOfRetirementIncome: [1, 100],
  returnBeforeRetirement: [0, 20], returnDuringRetirement: [0, 20], expectedInflation: [0, 20],
};
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // one submission at a time so rows never collide
  let sheet, row;
  try {
    const data = JSON.parse(e.postData.contents);
    const inputs = cleanInputs(data);
    // Results are recalculated here from the inputs, so the emailed report
    // always matches the calculator (nothing sent from the page is trusted).
    const results = calculateRetirement(inputs);
    const record = Object.assign({}, inputs, results, {
      timestamp: new Date(),
      name: inputs.userInfo.name, email: inputs.userInfo.email, phone: phone(inputs.userInfo.phoneNumber),
      site: clean(String(data.site || "")).slice(0, 80), emailStatus: "sending",
    });

    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    ensureHeader(sheet);
    sheet.appendRow(FIELDS.map((f) => f === "timestamp" ? record.timestamp : f === "phone" ? record.phone : clean(record[f])));
    row = sheet.getLastRow();

    const status = emailReport(inputs, results, record.phone);
    sheet.getRange(row, FIELDS.indexOf("emailStatus") + 1).setValue(status);
    return ContentService.createTextOutput("ok");
  } catch (err) {
    if (sheet && row) sheet.getRange(row, FIELDS.indexOf("emailStatus") + 1).setValue("error: " + err);
    return ContentService.createTextOutput("error: " + err);
  } finally {
    lock.releaseLock();
  }
}

function ensureHeader(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(FIELDS);
    sheet.setFrozenRows(1);
    return;
  }
  const current = sheet.getRange(1, 1, 1, FIELDS.length).getValues()[0];
  if (current.join("|") !== FIELDS.join("|")) sheet.getRange(1, 1, 1, FIELDS.length).setValues([FIELDS]);
}

function cleanInputs(d) {
  const num = (k) => {
    const v = Number(d[k]);
    const [lo, hi] = LIMITS[k];
    return isFinite(v) ? Math.min(hi, Math.max(lo, v)) : lo;
  };
  const inputs = {};
  Object.keys(LIMITS).forEach((k) => { inputs[k] = num(k); });
  inputs.currentSavings = inputs.brokerageAccount + inputs.traditionalIraAccount + inputs.rothIraAccount + inputs.account401k + inputs.otherAccounts;
  inputs.isMarried = d.isMarried === true;
  inputs.includeSocialSecurity = d.includeSocialSecurity === true;
  inputs.userInfo = {
    name: String(d.name || "").trim().slice(0, 100),
    email: String(d.email || "").trim().slice(0, 200),
    phoneNumber: String(d.phone || "").trim().slice(0, 40),
  };
  return inputs;
}

// ------------------------------------------------------------------ email

function emailReport(inputs, results, phoneText) {
  const u = inputs.userInfo;
  if (!u.name || !EMAIL_RE.test(u.email)) return "not sent: missing name or valid email";

  const props = PropertiesService.getScriptProperties();
  const dayKey = "sent-" + Utilities.formatDate(new Date(), TIME_ZONE, "yyyy-MM-dd");
  const sentToday = Number(props.getProperty(dayKey) || 0);
  if (sentToday >= MAX_REPORTS_PER_DAY) return "not sent: daily limit reached";
  const cache = CacheService.getScriptCache();
  const addrKey = "addr-" + u.email.toLowerCase();
  const sentToAddr = Number(cache.get(addrKey) || 0);
  if (sentToAddr >= MAX_REPORTS_PER_ADDRESS) return "not sent: too many requests for this address";

  const pdf = Utilities.newBlob(reportHtml(inputs, results, phoneText), MimeType.HTML, "report.html")
    .getAs(MimeType.PDF)
    .setName("Retirement Planning Analysis - " + safeFileName(u.name) + ".pdf");

  const first = u.name.split(/\s+/)[0];
  const subject = "Your Retirement Planning Analysis";
  const htmlBody =
    "<p>Hi " + esc(first) + ",</p>" +
    "<p>Thank you for using the Retirement Calculator. Your Retirement Planning Analysis is attached as a PDF.</p>" +
    "<p>If you'd like to talk through your results, reply to this email, call me at " + ADVISOR_PHONE +
    ', or <a href="' + BOOKING_URL + '">book an appointment</a>.</p>' +
    "<p>" + ADVISOR_NAME + "<br>Financial Advisor, Alexander Capital, L.P.<br>" + ADVISOR_PHONE +
    '<br><a href="mailto:' + ADVISOR_EMAIL + '">' + ADVISOR_EMAIL + '</a><br><a href="https://' + ADVISOR_WEBSITE + '">' + ADVISOR_WEBSITE + "</a></p>" +
    '<p style="font-size:11px;color:#52606D">' + DISCLAIMER + "</p>";
  const body =
    "Hi " + first + ",\n\nThank you for using the Retirement Calculator. Your Retirement Planning Analysis is attached as a PDF.\n\n" +
    "If you'd like to talk through your results, reply to this email, call me at " + ADVISOR_PHONE + ", or book an appointment: " + BOOKING_URL + "\n\n" +
    ADVISOR_NAME + "\nFinancial Advisor, Alexander Capital, L.P.\n" + ADVISOR_PHONE + "\n" + ADVISOR_EMAIL + "\n" + ADVISOR_WEBSITE + "\n\n" + DISCLAIMER;

  const options = { name: ADVISOR_NAME, bcc: ADVISOR_EMAIL, replyTo: ADVISOR_EMAIL, htmlBody: htmlBody, attachments: [pdf] };
  let status = "emailed";
  try {
    GmailApp.sendEmail(u.email, subject, body, Object.assign({ from: ADVISOR_EMAIL }, options));
  } catch (err) {
    // The work address isn't set up under "Send mail as" yet: send from the
    // Gmail address instead (replies still go to the work address).
    GmailApp.sendEmail(u.email, subject, body, options);
    status = "emailed (from Gmail address; work address not set up in Send mail as)";
  }
  props.setProperty(dayKey, String(sentToday + 1));
  cache.put(addrKey, String(sentToAddr + 1), 6 * 60 * 60);
  return status;
}

const DISCLAIMER = "Information and interactive calculators are made available to you as self-help tools for your independent use and are not intended to provide investment advice. We cannot and do not guarantee their applicability or accuracy in regards to your individual circumstances. All examples are hypothetical and are for illustrative purposes. We encourage you to seek personalized advice from qualified professionals regarding all personal finance issues. Not investment advice.";

// ------------------------------------------------------------------ PDF report

function money(v) {
  const neg = v < 0;
  const fixed = Math.abs(v).toFixed(2).split(".");
  return (neg ? "-$" : "$") + fixed[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",") + "." + fixed[1];
}
const pct2 = (v) => v.toFixed(2) + "%";

function reportHtml(e, t, phoneText) {
  const u = e.userInfo;
  const goalAge = e.retirementAge + e.yearsOfRetirementIncome;
  const lasts = t.savingsRunOutAge >= goalAge;
  const annualSave = e.annualIncome * (e.percentToSave / 100);
  const today = Utilities.formatDate(new Date(), TIME_ZONE, "MMMM d, yyyy");
  const row = (a, b) => "<tr><td>" + a + '</td><td class="r">' + b + "</td></tr>";
  const contact =
    "Phone: " + ADVISOR_PHONE + "<br>E-mail: " + ADVISOR_EMAIL + "<br>Website: " + ADVISOR_WEBSITE +
    '<br><a href="' + BOOKING_URL + '">Book an appointment</a>';
  const verdict = lasts ? "Your retirement savings will last through age " + goalAge + "." : "Retirement savings runs out at age " + t.savingsRunOutAge + ".";
  const sentence = "This assumes annual retirement expenses of " + money(t.annualRetirementExpenses) + " which is " + e.preRetirementIncomeDesired +
    "% of your last year's income of " + money(t.lastYearIncome) + ".";
  const recs = lasts ? "" :
    "<h4>To help meet your goal, you may wish to do one of the following:</h4><ul>" +
    "<li>Increase contributions to <b>" + pct2(t.recommendedSavingsRate) + "</b> of your income (" + money(e.annualIncome * t.recommendedSavingsRate / 100) + " annual savings, which will increase as your income increases)</li>" +
    "<li>Increase your rate of return before retirement to <b>" + pct2(t.recommendedReturnRate) + "</b>.</li>" +
    "<li>Reduce your required income in retirement to <b>" + pct2(t.recommendedIncomePercent) + "</b> of your final year's income.</li>" +
    "<li>Delay your retirement until age <b>" + t.recommendedRetirementAge + "</b>.</li></ul>";
  const slices = [
    ["Retirement balance", t.retirementBalance, "#2a78d6"], ["Final year income", t.lastYearIncome, "#eb6834"],
    ["Annual expenses", t.annualRetirementExpenses, "#1baf7a"], ["Social Security", t.annualSocialSecurity, "#eda100"],
  ];
  const legend = '<table class="info">' + slices.map((b) => '<tr><td><span style="color:' + b[2] + ';font-size:12pt">&#9632;</span> ' + b[0] +
    '</td><td class="r">' + money(b[1]) + "</td></tr>").join("") + "</table>";
  let pieImg = "";
  try {
    pieImg = '<img src="data:image/png;base64,' + Utilities.base64Encode(piePng(slices.map((b) => [b[1], b[2]]), 300)) + '" width="130" height="130">';
  } catch (err) {
    pieImg = ""; // if the image can't be made, the report still goes out with the legend
  }
  const breakdown = pieImg
    ? '<table style="width:100%;margin-top:10px"><tr><td style="width:140px;vertical-align:middle">' + pieImg + '</td><td style="vertical-align:middle">' + legend + "</td></tr></table>"
    : '<div style="margin-top:10px">' + legend + "</div>";

  const head = "<thead><tr><th>Age</th><th class=\"r\">Beginning retirement balance</th><th class=\"r\">Investment growth</th><th class=\"r\">Contributions at " +
    e.percentToSave + "% of income</th><th class=\"r\">Annual income desired</th><th class=\"r\">Retirement account withdrawals</th><th class=\"r\">Ending retirement balance</th></tr></thead>";
  const yearRow = (r) => "<tr" + (r.age >= e.retirementAge ? ' class="ret"' : "") + "><td>" + r.age + '</td><td class="r">' + money(r.beginningBalance) +
    '</td><td class="r">' + money(r.investmentGrowth) + '</td><td class="r">' + money(r.contributions) + '</td><td class="r">' + money(r.annualIncomeDesired) +
    '</td><td class="r">' + money(r.withdrawals) + '</td><td class="r">' + money(r.endingBalance) + "</td></tr>";
  const working = t.yearlyData.filter((r) => r.age < e.retirementAge);
  const retired = t.yearlyData.filter((r) => r.age >= e.retirementAge);
  const yearTable = (title, rows) => '<div class="pb"><h4>' + title + '</h4><table class="years">' + head + "<tbody>" + rows.map(yearRow).join("") + "</tbody></table></div>";

  return '<html><head><meta charset="utf-8"><style>' +
    "body{font-family:Arial,Helvetica,sans-serif;color:#102A43;font-size:10.5pt;margin:0}" +
    "a{color:#102A43}" +
    ".lh{width:100%;border-bottom:2px solid #102A43;margin-bottom:8px}.lh td{vertical-align:top;padding:0 0 8px}" +
    ".co{font-family:Georgia,serif;font-size:19pt}.by{font-size:9pt;color:#52606D;line-height:1.4}" +
    "h3{font-family:Georgia,serif;font-weight:normal;font-size:15pt;margin:6px 0 4px}h4{font-size:10.5pt;margin:10px 0 4px}" +
    ".box{padding:6px 10px;margin:6px 0}.ok{background:#E6F2EE;color:#0B5D43}.warn{background:#FBE9E7;color:#8C2A21}" +
    "ul{margin:4px 0;padding-left:18px}li{margin:1px 0}" +
    ".grid{width:100%;border-collapse:collapse}.grid>tbody>tr>td{width:50%;vertical-align:top;padding:0 10px 0 0}" +
    ".info{width:100%;border-collapse:collapse;font-size:9.5pt}.info td{border-bottom:1px solid #D9E2EC;padding:3px 0}" +
    ".r{text-align:right}.sw{display:inline-block;width:9px;height:9px;margin-right:6px}" +
    ".pb{page-break-before:always}" +
    ".years{width:100%;border-collapse:collapse;font-size:8.5pt}.years th{font-weight:bold;color:#52606D;border-bottom:1px solid #D9E2EC;padding:3px 5px;vertical-align:bottom}" +
    ".years td{border-bottom:1px solid #D9E2EC;padding:2px 5px}.years tr.ret td{background:#FBF5EA}" +
    ".note{font-size:8pt;color:#52606D}.foot{border-top:1px solid #D9E2EC;margin-top:12px;padding-top:8px;font-size:9pt;color:#52606D}" +
    "</style></head><body>" +
    '<table class="lh"><tr><td><div class="co">Alexander Capital L.P.</div><div class="by">Prepared by<br><b style="color:#102A43">' + ADVISOR_NAME +
    "</b>, Financial Advisor<br>" + contact + '</div></td><td class="r by">' + today + '<br><b style="color:#102A43">Retirement Planning Analysis</b></td></tr></table>' +
    "<h3>Retirement savings and expenditures</h3>" +
    "<p><b>Prepared for:</b> " + esc(u.name) + ", " + esc(u.email) + (phoneText ? ", " + esc(phoneText) : "") + "</p>" +
    '<div class="box ' + (lasts ? "ok" : "warn") + '">' + verdict + " " + sentence + "</div>" + recs +
    '<table class="grid"><tr><td>' +
    '<h4>Personal information</h4><table class="info">' + row("Name", esc(u.name)) + row("Email", esc(u.email)) + row("Phone", esc(phoneText) || "Not provided") + "</table>" +
    '<h4>Retirement plan inputs</h4><table class="info">' + row("Current age", e.currentAge) + row("Age at retirement", e.retirementAge) +
    row("Household income", money(e.annualIncome)) + row("Current retirement savings", money(e.currentSavings)) +
    row("Percent of income to save", e.percentToSave + "% (" + money(annualSave) + " annually which will increase as your income increases)") +
    row("Expected income increase", e.expectedIncomeIncrease + "%") + row("Years of retirement income", e.yearsOfRetirementIncome) +
    row("Desired retirement income", e.preRetirementIncomeDesired + "%") + "</table>" +
    "</td><td>" +
    '<h4>Investment returns and inflation</h4><table class="info">' + row("Rate of return before retirement", e.returnBeforeRetirement + "%") +
    row("Rate of return during retirement", e.returnDuringRetirement + "%") + row("Expected inflation rate", e.expectedInflation + "%") +
    row("Are you married?", e.isMarried ? "yes" : "no") + row("Include Social Security?", e.includeSocialSecurity ? "yes" : "no") + "</table>" +
    '<h4>Result summary</h4><table class="info">' + row("Years until retirement", e.retirementAge - e.currentAge) + row("Your last year's income", money(t.lastYearIncome)) +
    row("Estimated annual retirement expenditures", money(t.annualRetirementExpenses)) + row("Your ending balance", money(t.endingBalance)) + "</table>" +
    breakdown +
    "</td></tr></table>" +
    (working.length ? yearTable("Balances by year*: before retirement", working) : "") +
    (retired.length ? yearTable("Balances by year*: retirement, from age " + e.retirementAge, retired) : "") +
    '<p class="note">* This chart reflects fixed annual rates of return, expense, and income growth. It is meant solely as a guide. It is not likely that these and other variables will increase and decrease by a steady annual rate year over year.</p>' +
    '<div class="foot"><b style="color:#102A43">Questions about your retirement plan?</b> Contact ' + ADVISOR_NAME + ".<br>" + contact +
    "<p>Schedule a consultation to discuss your personalized retirement strategy and explore opportunities to optimize your financial future.</p>" +
    '<b style="color:#102A43">Alexander Capital L.P.</b> Building Your Financial Future<p>' + DISCLAIMER + "</p></div>" +
    "</body></html>";
}

// ------------------------------------------------------------------ pie chart image

// Draws the results pie (same colors and order as the website: clockwise from
// 12 o'clock, thin white lines between slices) and returns PNG bytes.
function piePng(slices, size) {
  const total = slices.reduce((a, s) => a + Math.max(0, s[0]), 0);
  if (!(total > 0)) throw new Error("nothing to chart");
  const rgb = (hex) => [parseInt(hex.slice(1, 3), 16), parseInt(hex.slice(3, 5), 16), parseInt(hex.slice(5, 7), 16)];
  const ends = []; // cumulative end angle (radians from 12 o'clock, clockwise) and color
  let acc = 0;
  slices.forEach((s) => { if (s[0] > 0) { acc += s[0] / total * 2 * Math.PI; ends.push([acc, rgb(s[1])]); } });
  const multi = ends.length > 1;
  const c = size / 2, r = size / 2 - 2, gap = size / 150; // white line half-width in pixels
  const bounds = multi ? ends.map((e, k) => (k === ends.length - 1 ? 0 : e[0])) : [];
  const WHITE = [255, 255, 255];
  // Color at one point: white outside the circle or on a slice boundary.
  const colorAt = (px, py) => {
    const dx = px - c, dy = py - c, d = Math.sqrt(dx * dx + dy * dy);
    if (d > r) return WHITE;
    let a = Math.atan2(dx, -dy); if (a < 0) a += 2 * Math.PI;
    for (let k = 0; k < bounds.length; k++) {
      let diff = Math.abs(a - bounds[k]); if (diff > Math.PI) diff = 2 * Math.PI - diff;
      if (diff < Math.PI / 2 && d * Math.sin(diff) < gap) return WHITE;
    }
    let i = 0; while (i < ends.length - 1 && a > ends[i][0]) i++;
    return ends[i][1];
  };
  // Pixels away from any edge take one sample; edge pixels take 16 for smooth lines.
  const nearEdge = (px, py) => {
    const dx = px - c, dy = py - c, d = Math.sqrt(dx * dx + dy * dy);
    if (Math.abs(d - r) < 1.5) return true;
    if (d > r) return false;
    let a = Math.atan2(dx, -dy); if (a < 0) a += 2 * Math.PI;
    for (let k = 0; k < bounds.length; k++) {
      let diff = Math.abs(a - bounds[k]); if (diff > Math.PI) diff = 2 * Math.PI - diff;
      if (diff < Math.PI / 2 && Math.abs(d * Math.sin(diff) - gap) < 1.5) return true;
    }
    return false;
  };
  const SS = 4;
  const raw = new Array(size * (size * 3 + 1)); // filtered scanlines for PNG: filter byte 0 + RGB
  let o = 0;
  for (let y = 0; y < size; y++) {
    raw[o++] = 0;
    for (let x = 0; x < size; x++) {
      if (!nearEdge(x + 0.5, y + 0.5)) {
        const col = colorAt(x + 0.5, y + 0.5);
        raw[o++] = col[0]; raw[o++] = col[1]; raw[o++] = col[2];
        continue;
      }
      let R = 0, G = 0, B = 0;
      for (let sy = 0; sy < SS; sy++) for (let sx = 0; sx < SS; sx++) {
        const col = colorAt(x + (sx + 0.5) / SS, y + (sy + 0.5) / SS);
        R += col[0]; G += col[1]; B += col[2];
      }
      const n = SS * SS;
      raw[o++] = Math.round(R / n); raw[o++] = Math.round(G / n); raw[o++] = Math.round(B / n);
    }
  }
  return pngBytes(size, size, raw);
}

// Minimal PNG writer (8-bit RGB). Compression uses Apps Script's gzip,
// re-wrapped as the zlib stream PNG expects.
function pngBytes(w, h, raw) {
  const u8 = (arr) => arr.map((b) => b & 255);
  const gz = u8(Utilities.gzip(Utilities.newBlob(raw.map((b) => (b > 127 ? b - 256 : b)))).getBytes());
  let p = 10; const flg = gz[3];
  if (flg & 4) p += 2 + (gz[10] | (gz[11] << 8));
  if (flg & 8) { while (gz[p] !== 0) p++; p++; }
  if (flg & 16) { while (gz[p] !== 0) p++; p++; }
  if (flg & 2) p += 2;
  const deflate = gz.slice(p, gz.length - 8);
  let a = 1, b = 0;
  for (let i = 0; i < raw.length; i++) { a = (a + raw[i]) % 65521; b = (b + a) % 65521; }
  const zlib = [0x78, 0x9c].concat(deflate, [(b >>> 8) & 255, b & 255, (a >>> 8) & 255, a & 255]);
  const crcTable = [];
  for (let n = 0; n < 256; n++) { let k = n; for (let j = 0; j < 8; j++) k = k & 1 ? 0xedb88320 ^ (k >>> 1) : k >>> 1; crcTable[n] = k >>> 0; }
  const crc = (bytes) => { let x = 0xffffffff; for (let i = 0; i < bytes.length; i++) x = crcTable[(x ^ bytes[i]) & 255] ^ (x >>> 8); return (x ^ 0xffffffff) >>> 0; };
  const u32 = (v) => [(v >>> 24) & 255, (v >>> 16) & 255, (v >>> 8) & 255, v & 255];
  const chunk = (type, data) => {
    const td = type.split("").map((ch) => ch.charCodeAt(0)).concat(data);
    return u32(data.length).concat(td, u32(crc(td)));
  };
  const out = [137, 80, 78, 71, 13, 10, 26, 10]
    .concat(chunk("IHDR", u32(w).concat(u32(h), [8, 2, 0, 0, 0])), chunk("IDAT", zlib), chunk("IEND", []));
  return out.map((v) => (v > 127 ? v - 256 : v)); // Apps Script byte arrays are signed
}

// ------------------------------------------------------------------ helpers

function esc(s) {
  return String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function safeFileName(s) {
  return String(s).replace(/[^\w .-]/g, "").trim().slice(0, 60) || "Client";
}

// Text typed by website visitors is stored as plain text, never run as a spreadsheet formula.
function clean(v) {
  if (v === undefined || v === null) return "";
  if (typeof v === "string") {
    v = v.slice(0, 500);
    if (/^[=+\-@]/.test(v)) v = "'" + v;
  }
  return v;
}

// Turns 7325551234, 732-555-1234 or +1 732 555 1234 into (732) 555-1234.
// Anything that isn't a 10-digit US number is kept as typed.
function phone(v) {
  let d = String(v || "").replace(/\D/g, "");
  if (d.length === 11 && d[0] === "1") d = d.slice(1);
  if (d.length !== 10) return clean(v);
  return "(" + d.slice(0, 3) + ") " + d.slice(3, 6) + "-" + d.slice(6);
}

// Run this once from the editor (select testEmail, click Run) to send yourself
// a sample report and check the sender address and the PDF.
function testEmail() {
  const inputs = cleanInputs({ name: "Test Client", email: ADVISOR_EMAIL, phone: "7325551234", currentAge: 45, retirementAge: 65,
    annualIncome: 150000, percentToSave: 10, brokerageAccount: 100000, account401k: 250000, expectedIncomeIncrease: 2,
    preRetirementIncomeDesired: 70, yearsOfRetirementIncome: 25, returnBeforeRetirement: 7, returnDuringRetirement: 5,
    expectedInflation: 2.5, isMarried: true, includeSocialSecurity: true });
  Logger.log(emailReport(inputs, calculateRetirement(inputs), phone(inputs.userInfo.phoneNumber)));
}


// ===================================================================
// Calculations: exact copy of the website's retire-calc.js
// ===================================================================

// Retirement Planner calculations.
//
// Ported line for line from Jake Ippolito's Retirement Planning Calculator
// (github.com/jakeipp/retirement-planner). Only names and formatting were made
// readable; every formula, limit and default is unchanged. If the original
// changes, update this file to match.

// Social Security: sliding scale on final-year income, adjusted for
// retirement age, x1.5 for a married couple, capped at the maximum.
const SS_INCOME_FOR_MAX = 176100;
const SS_MAX_BENEFIT = 48216;
const SS_FULL_RETIREMENT_AGE = 67;

function socialSecurityBenefit(income, isMarried, retirementAge) {
  let benefit = Math.min(income / SS_INCOME_FOR_MAX, 1) * SS_MAX_BENEFIT;
  if (retirementAge < SS_FULL_RETIREMENT_AGE) {
    const yearsEarly = SS_FULL_RETIREMENT_AGE - retirementAge;
    benefit *= 1 - yearsEarly * 0.067;
  } else if (retirementAge > SS_FULL_RETIREMENT_AGE) {
    const yearsLate = Math.min(retirementAge - SS_FULL_RETIREMENT_AGE, 3);
    benefit *= 1 + yearsLate * 0.08;
  }
  if (isMarried) benefit *= 1.5;
  const cap = isMarried ? SS_MAX_BENEFIT * 1.5 : SS_MAX_BENEFIT;
  return Math.min(benefit, cap);
}

// Year-by-year projection. No contribution is made in the year of retirement.
function projectRetirement(e) {
  const {
    currentAge, retirementAge, annualIncome, percentToSave, currentSavings,
    expectedIncomeIncrease, preRetirementIncomeDesired, yearsOfRetirementIncome,
    returnBeforeRetirement, returnDuringRetirement, expectedInflation,
    isMarried, includeSocialSecurity,
  } = e;
  const yearlyData = [];
  let balance = currentSavings;
  let income = annualIncome;
  let lastYearIncome = annualIncome;
  let savingsRunOutAge = retirementAge + yearsOfRetirementIncome;
  const endAge = retirementAge + yearsOfRetirementIncome;

  for (let age = currentAge + 1; age < retirementAge; age++) {
    const beginningBalance = balance;
    const investmentGrowth = beginningBalance * (returnBeforeRetirement / 100);
    income = income * (1 + expectedIncomeIncrease / 100);
    lastYearIncome = income;
    const contributions = income * (percentToSave / 100);
    balance = beginningBalance + investmentGrowth + contributions;
    yearlyData.push({ age, beginningBalance, investmentGrowth, contributions, annualIncomeDesired: 0, withdrawals: 0, endingBalance: balance });
  }
  if (retirementAge > currentAge + 1) {
    income = income * (1 + expectedIncomeIncrease / 100);
    lastYearIncome = income;
  }
  const retirementBalance = balance;
  const annualRetirementExpenses = lastYearIncome * (preRetirementIncomeDesired / 100);
  const annualSocialSecurity = includeSocialSecurity ? socialSecurityBenefit(lastYearIncome, isMarried, retirementAge) : 0;

  let incomeDesired = annualRetirementExpenses;
  for (let age = retirementAge; age < endAge; age++) {
    const beginningBalance = balance;
    const investmentGrowth = beginningBalance * (returnDuringRetirement / 100);
    if (age > retirementAge) incomeDesired = incomeDesired * (1 + expectedInflation / 100);
    const needed = Math.max(0, incomeDesired - annualSocialSecurity);
    const withdrawals = Math.min(needed, beginningBalance + investmentGrowth);
    balance = Math.max(0, beginningBalance + investmentGrowth - withdrawals);
    if (balance === 0 && savingsRunOutAge === endAge) savingsRunOutAge = age;
    yearlyData.push({ age, beginningBalance, investmentGrowth, contributions: 0, annualIncomeDesired: incomeDesired, withdrawals, endingBalance: balance });
  }
  return { retirementBalance, lastYearIncome, annualRetirementExpenses, annualSocialSecurity, savingsRunOutAge, endingBalance: balance, yearlyData };
}

const lastsToGoal = (inputs) => projectRetirement(inputs).savingsRunOutAge >= inputs.retirementAge + inputs.yearsOfRetirementIncome;

// Savings rate needed (binary search 0-100%, rounded up to 0.01).
function recommendedSavingsRate(e) {
  let lo = 0, hi = 100;
  for (let i = 0; i < 50; i++) {
    const mid = (lo + hi) / 2;
    if (lastsToGoal({ ...e, percentToSave: mid })) hi = mid; else lo = mid;
  }
  return Math.ceil(hi * 100) / 100;
}

// Pre-retirement return needed (binary search 0-20%, rounded up to 0.01).
function recommendedReturnRate(e) {
  let lo = 0, hi = 20;
  for (let i = 0; i < 50; i++) {
    const mid = (lo + hi) / 2;
    if (lastsToGoal({ ...e, returnBeforeRetirement: mid })) hi = mid; else lo = mid;
  }
  return Math.ceil(hi * 100) / 100;
}

// Retirement income the savings can support (binary search 40%-current, rounded down).
function recommendedIncomePercent(e) {
  let lo = 40, hi = e.preRetirementIncomeDesired;
  for (let i = 0; i < 50; i++) {
    const mid = (lo + hi) / 2;
    if (lastsToGoal({ ...e, preRetirementIncomeDesired: mid })) lo = mid; else hi = mid;
  }
  return Math.floor(lo);
}

// First retirement age (up to 90) at which the savings last.
function recommendedRetirementAge(e) {
  for (let age = e.retirementAge; age <= 90; age++) {
    if (lastsToGoal({ ...e, retirementAge: age })) return age;
  }
  return 90;
}

function calculateRetirement(e) {
  return {
    ...projectRetirement(e),
    recommendedSavingsRate: recommendedSavingsRate(e),
    recommendedReturnRate: recommendedReturnRate(e),
    recommendedIncomePercent: recommendedIncomePercent(e),
    recommendedRetirementAge: recommendedRetirementAge(e),
  };
}

const RETIREMENT_DEFAULTS = {
  currentAge: 25, retirementAge: 65, annualIncome: 100000, percentToSave: 5, currentSavings: 0,
  brokerageAccount: 0, traditionalIraAccount: 0, rothIraAccount: 0, account401k: 0, otherAccounts: 0,
  expectedIncomeIncrease: 0, preRetirementIncomeDesired: 40, yearsOfRetirementIncome: 25,
  returnBeforeRetirement: 10, returnDuringRetirement: 5, expectedInflation: 0,
  isMarried: false, includeSocialSecurity: false,
  userInfo: { name: "", email: "", phoneNumber: "" },
};
