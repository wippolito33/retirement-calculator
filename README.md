# Retirement Calculator

A one-page retirement calculator for Bill Ippolito, Financial Advisor, Alexander Capital, L.P. It is a standalone copy of the Retirement Planner tab on Portfolio Builder 500.

- `index.html`: the page (layout and styles)
- `retire-calc.js`: the calculations, ported unchanged from Jake Ippolito's calculator
- `retire.js`: inputs, results, chart and report
- `bill-ippolito.jpg`: contact photo
- `apps-script/Code.gs`: the Google Sheet script (records leads, emails the PDF report)

"View report" requires a name and email and sends the request to the Apps Script web app at `SHEET_URL` in `retire.js`. That script (`apps-script/Code.gs`, pasted into the Google Sheet's Apps Script editor) adds a row to the sheet, recalculates the results, and emails the client a PDF of the report from Bill's work address (set up in Gmail under "Send mail as"), with a bcc to Bill. It caps email at 40 reports a day and 3 per client address every 6 hours.

Static site, no build step. Netlify publishes the repo root (see `netlify.toml`) and redeploys on every push to `main`.

If the Retirement Planner on Portfolio Builder 500 changes, copy `retire-calc.js` and `retire.js` over to keep the two in step.
