"""All add-in logic, in Python (runs in Pyodide). taskpane.js only boots the runtime."""
import json


def plan_unhide(sheets):
    """Names of sheets that are not Visible (includes VeryHidden). Pure, testable."""
    return [s["name"] for s in sheets if s["visibility"] != "Visible"]


async def unhide_all():
    """Read sheets, decide in Python, write visibility back. Returns unhidden names."""
    # Imported lazily so plan_unhide can be unit-tested outside Pyodide.
    from js import Excel
    from pyodide.ffi import create_proxy

    result = []

    async def batch(ctx):
        sheets = ctx.workbook.worksheets
        sheets.load("items/name,items/visibility")
        await ctx.sync()
        info = [{"name": s.name, "visibility": s.visibility} for s in sheets.items]
        names = plan_unhide(info)
        for name in names:
            sheets.getItem(name).visibility = "Visible"
        await ctx.sync()
        result.extend(names)

    proxy = create_proxy(batch)
    try:
        await Excel.run(proxy)
    finally:
        proxy.destroy()
    return result


async def main():
    """Entry point called on add-in load; returns a status string for the pane."""
    try:
        names = await unhide_all()
    except Exception as e:  # surface failures in the pane instead of silently dying
        return f"Error: {e}"
    if not names:
        return "No hidden sheets."
    return f"Unhid {len(names)} sheet(s):\n" + "\n".join(names)
