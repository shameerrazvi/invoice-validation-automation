"""Run the VBA macro in LibreOffice and check it matches the Python validator exactly.

Needs LibreOffice installed (its Python UNO bridge must be importable). The macro is
loaded into a copy of the tool workbook with VBA compatibility switched on, run, and
its Results and Exceptions sheets compared row by row with the Python output.

Usage:  python tools/check_vba_parity.py
"""
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from src import config as cfg, validate, workbook  # noqa: E402

import uno  # noqa: E402  (LibreOffice's Python bridge)
from com.sun.star.beans import PropertyValue  # noqa: E402


def prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def basic_source() -> str:
    lines = (ROOT / "excel" / "InvoiceValidation.bas").read_text().splitlines()
    body = [ln for ln in lines if not ln.startswith("Attribute VB_")]
    return "Option VBASupport 1\n" + "\n".join(body)


def run_macro(src: Path, out: Path) -> float:
    port = 2202
    profile = Path(tempfile.mkdtemp())
    office = subprocess.Popen(["soffice", "--headless", "--invisible", "--norestore", "--nologo",
                               f"-env:UserInstallation=file://{profile}",
                               f"--accept=socket,host=localhost,port={port};urp;"],
                              stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        for _ in range(60):
            try:
                ctx = resolver.resolve(f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(0.5)
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(str(src)), "_blank", 0, (prop("Hidden", True),))
        libs = doc.BasicLibraries
        if not libs.hasByName("Standard"):
            libs.createLibrary("Standard")
        lib = libs.getByName("Standard")
        lib.insertByName("InvoiceValidation", basic_source())
        script = doc.getScriptProvider().getScript(
            "vnd.sun.star.script:Standard.InvoiceValidation.RunValidation?language=Basic&location=document")
        seconds, _, _ = script.invoke((), (), ())
        doc.storeToURL(uno.systemPathToFileUrl(str(out)), (prop("FilterName", "Calc MS Excel 2007 XML"),))
        doc.close(True)
        return float(seconds)
    finally:
        office.terminate()
        office.wait(timeout=30)
        shutil.rmtree(profile, ignore_errors=True)


def normalise(df):
    df = df.astype(object).where(df.notna(), "")
    for c in df.columns:
        if "date" in c:
            df[c] = pd.to_datetime(df[c], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
        elif c in ("total_inc_gst",):
            df[c] = df[c].astype(float).round(2)
        elif c == "line_no":
            df[c] = df[c].map(lambda v: "" if v == "" else str(int(float(v))))
        else:
            df[c] = df[c].astype(str)
    return df.reset_index(drop=True)


def main():
    work = Path(tempfile.mkdtemp())
    src, out = work / "tool.xlsx", work / "tool_after_vba.xlsx"
    shutil.copy(cfg.TOOL_WORKBOOK, src)
    seconds = run_macro(src, out)

    data, settings = workbook.read_inputs(cfg.TOOL_WORKBOOK)
    py_res, py_exc = validate.validate(data, settings)
    vba_res = pd.read_excel(out, sheet_name="Results", dtype=str, keep_default_na=False)
    vba_exc = pd.read_excel(out, sheet_name="Exceptions", dtype=str, keep_default_na=False)

    a, b = normalise(py_res), normalise(vba_res)
    c, d = normalise(py_exc), normalise(vba_exc)
    res_ok = a.equals(b)
    exc_ok = c.equals(d)
    print(f"VBA macro ran in LibreOffice in {seconds:.2f}s")
    print(f"Results:    {len(a)} Python rows, {len(b)} VBA rows -> {'IDENTICAL' if res_ok else 'DIFFERENT'}")
    print(f"Exceptions: {len(c)} Python rows, {len(d)} VBA rows -> {'IDENTICAL' if exc_ok else 'DIFFERENT'}")
    if not (res_ok and exc_ok):
        for name, x, y in (("Results", a, b), ("Exceptions", c, d)):
            if len(x) == len(y):
                diff = (x != y).any(axis=1)
                print(name, "first differences:")
                print(pd.concat([x[diff].head(5), y[diff].head(5)], keys=["python", "vba"]).to_string())
        sys.exit(1)
    shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    main()
