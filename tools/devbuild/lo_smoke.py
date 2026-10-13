#!/usr/bin/env python3
"""Proxy smoke run of the development workbook in headless LibreOffice.

Usage: python tools/devbuild/lo_smoke.py [build/IRRBB_Dev.xlsm]

Opens the workbook in LibreOffice with VBA compatibility, runs
TEST_Harness.RunTests, plants a stale row in the result tables, runs it again to
check that previous results are withdrawn, runs it once more with every results
write failing (both errors must be kept) and prints a JSON summary. Needs
LibreOffice and python3-uno. LibreOffice is NOT Excel: this is never Excel
evidence (docs/EXCEL_EVIDENCE.md), only an early warning for runtime defects.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import time
from pathlib import Path

import uno  # type: ignore[import-not-found]
from com.sun.star.beans import PropertyValue  # type: ignore[import-not-found]

PORT = 2007


def pv(name: str, value: object) -> PropertyValue:
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def main() -> int:
    root = Path(__file__).resolve().parents[2]
    path = Path(sys.argv[1]) if len(sys.argv) > 1 else root / "build" / "IRRBB_Dev.xlsm"
    proc = subprocess.Popen(["soffice", "--headless", "--invisible", "--norestore", "--nologo",
                             f"--accept=socket,host=localhost,port={PORT};urp;"],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        ctx = None
        for _ in range(120):
            try:
                ctx = resolver.resolve(f"uno:socket,host=localhost,port={PORT};urp;StarOffice.ComponentContext")
                break
            except Exception:  # noqa: BLE001
                time.sleep(0.5)
        smgr = ctx.ServiceManager
        cp = smgr.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", ctx)
        vba = cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationUpdateAccess",
                                             (pv("nodepath", "/org.openoffice.Office.Calc/Filter/Import/VBA"),))
        vba.setPropertyValue("Load", True)
        vba.setPropertyValue("Executable", True)
        vba.commitChanges()
        desk = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desk.loadComponentFromURL(path.resolve().as_uri(), "_blank", 0,
                                        (pv("Hidden", True), pv("MacroExecutionMode", 4)))
        lib = "VBAProject"
        scripts = doc.getScriptProvider()

        def run(module: str, name: str) -> str | None:
            try:
                scripts.getScript(f"vnd.sun.star.script:{lib}.{module}.{name}?language=Basic&location=document"
                                  ).invoke((), (), ())
                return None
            except Exception as error:  # noqa: BLE001
                return str(error)

        def cells(sheet: str, ref: str):
            return doc.Sheets.getByName(sheet).getCellRangeByName(ref).getDataArray()

        out: dict[str, object] = {"modules": list(doc.BasicLibraries.getByName(lib).getElementNames())}
        out["harness_error"] = run("TEST_Harness", "RunTests")
        out["status"] = [r[1] for r in cells("Checks", "A3:B11")]
        out["cases"] = [r for r in cells("Checks", "A14:F80") if r[0]]
        rows = cells("Checks", "H14:N3000")
        out["assertion_rows"] = sum(1 for r in rows if r[0])
        out["failed_assertions"] = [r for r in rows if r[0] and r[6] != "PASS"]
        checks = doc.Sheets.getByName("Checks")
        checks.getCellRangeByName("A70").setString("STALE")
        checks.getCellRangeByName("H2900").setString("STALE")
        out["rerun_error"] = run("TEST_Harness", "RunTests")
        out["rerun_status"] = cells("Checks", "B3:B3")[0][0]
        out["stale_rows_withdrawn"] = (cells("Checks", "A70:A70")[0][0] in ("", None)
                                       and cells("Checks", "H2900:H2900")[0][0] in ("", None))
        # A run whose every results write fails must end FAILED with both errors kept and
        # no earlier result left in the tables.
        out["fault_error"] = run("TEST_Harness", "TEST_RunTestsWithWriteFault")
        status = [r[1] for r in cells("Checks", "A3:B11")]
        out["fault_status"], out["fault_primary"], out["fault_cleanup"] = status[0], status[7], status[8]
        out["fault_tables_empty"] = (not any(r[0] for r in cells("Checks", "A14:A514"))
                                     and not any(r[0] for r in cells("Checks", "H14:H3000")))
        doc.close(True)
        try:
            desk.terminate()
        except Exception:  # noqa: BLE001
            pass
        print(json.dumps(out, indent=1, default=str))
    finally:
        try:
            proc.wait(timeout=60)
        except subprocess.TimeoutExpired:
            proc.kill()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
