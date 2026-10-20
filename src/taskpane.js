/* Bootstrap only: loads Pyodide + sheets.py and runs it. All logic is in sheets.py. */
const log = (m) => { document.getElementById("log").textContent = m; };

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) { log("This add-in runs in Excel only."); return; }
  try {
    // Re-open this pane (and re-run) every time a workbook opens.
    await Office.addin.setStartupBehavior(Office.StartupBehavior.load);
  } catch (e) { console.warn("Startup behavior not set:", e); }
  try {
    log("Loading Python runtime...");
    const py = await loadPyodide();
    py.runPython(await (await fetch("sheets.py")).text());
    const run = async () => log(await py.runPythonAsync("await main()"));
    document.getElementById("run").onclick = run;
    document.getElementById("run").disabled = false;
    await run(); // auto-run on open
  } catch (e) { log("Failed to start Python: " + e.message); }
});
