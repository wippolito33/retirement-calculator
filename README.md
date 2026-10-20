# Retirement Calculator

A one-page retirement calculator for Bill Ippolito, Financial Advisor, Alexander Capital, L.P. It is a standalone copy of the Retirement Planner tab on Portfolio Builder 500.

- `index.html`: the page (layout and styles)
- `retire-calc.js`: the calculations, ported unchanged from Jake Ippolito's calculator
- `retire.js`: inputs, results, chart and report
- `bill-ippolito.jpg`: contact photo

"View report" requires a name and email and records the request (contact details, inputs and results) in Bill's Google Sheet through the Apps Script web app URL (`SHEET_URL`) in `retire.js`.

Static site, no build step. Netlify publishes the repo root (see `netlify.toml`) and redeploys on every push to `main`.

If the Retirement Planner on Portfolio Builder 500 changes, copy `retire-calc.js` and `retire.js` over to keep the two in step.
