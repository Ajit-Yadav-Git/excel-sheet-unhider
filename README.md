# Sheet Unhider (Excel add-in, Python via Pyodide)

All logic is Python ([src/sheets.py](src/sheets.py)). [src/taskpane.js](src/taskpane.js) is a ~20-line bootstrap that loads Pyodide and calls `main()`; Python talks to Excel through Office.js via Pyodide's `js` bridge.

On open, the add-in sets `Office.StartupBehavior.load` (needs the SharedRuntime in the manifest), so the pane loads with every workbook and runs automatically. The first time, open the pane once from Home > Sheet Unhider.

## Run
```
npm install
npm run certs      # one-time: trust the localhost HTTPS cert
npm run serve      # https://localhost:3000   (keep running)
npm run sideload   # second terminal; opens Excel with the add-in
```
Then open `sample/sample-hidden-sheets.xlsx` (Sales, Costs hidden; Secrets very hidden). Expect the pane to report `Unhid 3 sheet(s)`.

## Test
- Sample data: `powershell -File tools/make-sample.ps1` regenerates it (no Python/Excel needed).
- Logic: `python -m pytest tests`.
- Add `src/icon.png` (80x80) for the ribbon button.

## Notes
- Pyodide/Office.js load from CDNs; self-host or pin for production. First load ~10 MB.
- Rollback: `npm run stop` removes the sideloaded add-in.
