"""Offline tests for functions of tools/Sync-IRRBB-VBA.ps1 that need no Excel.

The functions are loaded from the script's syntax tree and run in PowerShell
7 (pwsh); the tests are skipped where pwsh is not installed. Nothing here
starts Excel or touches a workbook.
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path
from typing import Any

SCRIPT = Path(__file__).resolve().parent / "Sync-IRRBB-VBA.ps1"
PWSH = shutil.which("pwsh")

# Defines every function of the script, and the $NewModulePattern it uses,
# without running the script body.
LOADER = r"""
$ErrorActionPreference = 'Stop'
$ast = [System.Management.Automation.Language.Parser]::ParseFile($env:SYNC_SCRIPT, [ref]$null, [ref]$null)
foreach ($f in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false)) {
    . ([scriptblock]::Create($f.Extent.Text))
}
foreach ($a in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] }, $false)) {
    if ($a.Left.Extent.Text -eq '$NewModulePattern') { . ([scriptblock]::Create($a.Extent.Text)) }
}
"""

HEADER = 'Attribute VB_Name = "shConfig"\r\nAttribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"\r\n'


@unittest.skipUnless(PWSH, "PowerShell 7 (pwsh) is not installed")
class SyncScriptTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)

    def pwsh(self, body: str) -> Any:
        """Run body after the loader; it writes a JSON result to $env:RESULT."""
        result = self.root / "result.json"
        completed = subprocess.run(
            [str(PWSH), "-NoProfile", "-NonInteractive", "-Command", LOADER + body],
            env={**os.environ, "SYNC_SCRIPT": str(SCRIPT), "RESULT": str(result)},
            capture_output=True, text=True, timeout=120, check=False)
        self.assertEqual(completed.returncode, 0, completed.stderr + completed.stdout)
        return json.loads(result.read_text(encoding="utf-8"))

    def code_body(self, raw: bytes) -> str:
        source = self.root / "shConfig.cls"
        source.write_bytes(raw)
        body = self.pwsh(f"$b = Get-VbaCodeBody -SourcePath '{source}'\n"
                         "ConvertTo-Json -InputObject @($b) | Set-Content -LiteralPath $env:RESULT -Encoding utf8")
        text: str = body[0] if body else ""
        return text.replace("\r\n", "\n")

    def test_document_code_keeps_cp1252_text(self) -> None:
        # £, é and € in a comment and a string literal survive.
        code = "Option Explicit\r\n' Prezzo in £, café\r\nPrivate Const S As String = \"€ é £\"\r\n"
        expected = "Option Explicit\n' Prezzo in £, café\nPrivate Const S As String = \"€ é £\""
        self.assertEqual(self.code_body((HEADER + code).encode("cp1252")), expected)
        self.assertEqual(self.code_body((HEADER + code).replace("\r\n", "\n").encode("cp1252")), expected)

    def test_document_code_without_body_or_header(self) -> None:
        self.assertEqual(self.code_body(HEADER.encode("cp1252")), "")
        self.assertEqual(self.code_body(b""), "")
        self.assertEqual(self.code_body("Option Explicit\r\n' é\r\n".encode("cp1252")), "Option Explicit\n' é")

    def write_module(self, folder: Path, file_name: str, header: str) -> None:
        (folder / file_name).write_bytes((header + "Option Explicit\r\n").encode("cp1252"))

    def select(self, folder: Path) -> dict[str, Any]:
        repository = {"core_import": "src/core/CORE_Import.bas", "core_parse": "src/core/CORE_Parse.bas",
                      "core_codes": "src/core/CORE_Codes.bas", "wb_import": "src/workbook/WB_Import.bas",
                      "irrbb_import": "src/modules/IRRBB_Import.bas"}
        entries = "; ".join(f"'{k}' = '{v}'" for k, v in repository.items())
        return dict(self.pwsh(
            f"$repo = @{{ {entries} }}\n"
            "$picked = @(); $failed = ''; $shown = $null\n"
            "try {\n"
            f"  $out = Select-FolderSources -FolderPath '{folder}' -RepositoryNames $repo "
            f"-StagePath '{self.root / 'stage'}' -InformationVariable shown 6>$null\n"
            "  foreach ($o in $out) { if ($o -is [System.IO.FileInfo]) { $picked += $o.Name } }\n"
            "} catch { $failed = $_.Exception.Message }\n"
            "$lines = @(foreach ($r in $shown) { [string]$r })\n"
            "@{ picked = $picked; lines = $lines; failed = $failed } | ConvertTo-Json "
            "| Set-Content -LiteralPath $env:RESULT -Encoding utf8"))

    def test_folder_candidates_need_a_declared_vb_name(self) -> None:
        # The file name never stands in for a missing or invalid VB_Name.
        folder = self.root / "downloads"
        folder.mkdir()
        self.write_module(folder, "CORE_Import.bas", "")
        self.write_module(folder, "CORE_Import (1).bas", 'Attribute VB_Name = "CORE_Import"\r\n')
        self.write_module(folder, "CORE_Parse.bas", "Attribute VB_Name = CORE_Parse\r\n")
        self.write_module(folder, "CORE_Codes.bas",
                          'Attribute VB_Name = "CORE_Codes"\r\nAttribute VB_Name = "CORE_Parse"\r\n')
        self.write_module(folder, "WB_Import.bas", 'Attribute VB_Name = "WB Import"\r\n')
        self.write_module(folder, "CORE_Import (2).bas",
                          'Attribute VB_Name = "CORE_Import"\r\nAttribute VB_Name = "core_import"\r\n')
        self.write_module(folder, "Other.bas", 'Attribute VB_Name = "Other"\r\n')
        self.write_module(folder, "IRRBB_Import.cls", 'Attribute VB_Name = "IRRBB_Import"\r\n')
        result = self.select(folder)
        self.assertEqual(result["failed"], "")
        self.assertEqual(result["picked"], ["CORE_Import.bas"])
        lines = "\n".join(result["lines"])
        for expected in ("CORE_Import.bas (no Attribute VB_Name line; the file name is not used",
                         "CORE_Parse.bas (malformed Attribute VB_Name line)",
                         "CORE_Codes.bas (several Attribute VB_Name lines)",
                         "CORE_Import (2).bas (several Attribute VB_Name lines)",
                         "WB_Import.bas (VB_Name 'WB Import' is not a valid VBA name)",
                         "Other.bas (VB_Name 'Other' is not a repository component)",
                         "IRRBB_Import.cls ('IRRBB_Import' is a .bas component in the repository)"):
            self.assertIn(expected, lines)
        self.assertIn("CORE_Import (1).bas", lines)

    def test_folder_without_a_valid_module_stops_before_excel(self) -> None:
        folder = self.root / "downloads"
        folder.mkdir()
        self.write_module(folder, "CORE_Import.bas", "")
        result = self.select(folder)
        self.assertIn("No repository VBA module found", result["failed"])
        self.assertEqual(result["picked"], [])
        # The reason is shown before the stop, so the user can see why.
        self.assertIn("CORE_Import.bas (no Attribute VB_Name line", "\n".join(result["lines"]))


    def mode(self, answer: str) -> Any:
        return self.pwsh(f"$m = ConvertTo-SyncMode -Answer '{answer}'\n"
                         "ConvertTo-Json -InputObject @{ mode = $m } | Set-Content -LiteralPath $env:RESULT -Encoding utf8")["mode"]

    def test_mode_answer(self) -> None:
        # P/D, short and long words in any case; anything else cancels.
        for answer, expected in (("p", "Production"), (" PROD ", "Production"), ("production", "Production"),
                                 ("D", "Development"), ("dev", "Development"), ("Development", "Development"),
                                 ("", None), ("y", None), ("pd", None), ("production x", None)):
            with self.subTest(answer=answer):
                self.assertEqual(self.mode(answer), expected)

    def manifest_sources(self, mode: str) -> list[str]:
        repo = Path(__file__).resolve().parent.parent
        names = self.pwsh(f"$f = Get-ManifestSources -RepoPath '{repo}' -Mode {mode}\n"
                          "ConvertTo-Json -InputObject @($f | ForEach-Object { $_.Name }) "
                          "| Set-Content -LiteralPath $env:RESULT -Encoding utf8")
        return list(names)

    def test_manifest_sources_per_mode(self) -> None:
        # Production: the production modules and the document modules;
        # development adds every test and example module and nothing else.
        manifest = json.loads((SCRIPT.parent / "vba-projects.json").read_text(encoding="utf-8"))
        documents = {Path(p).name for p in manifest["document_modules"]}
        production = self.manifest_sources("Production")
        development = self.manifest_sources("Development")
        self.assertEqual(set(production), {Path(p).name for p in manifest["configurations"]["production"]["modules"]} | documents)
        self.assertEqual(set(development) - set(production),
                         {Path(p).name for p in manifest["configurations"]["development"]["modules"]
                          if p.startswith(("tests/", "examples/"))})
        self.assertFalse([n for n in production if n.startswith("TEST_")])
        self.assertEqual(len(development), len(set(development)))

    def test_manifest_entry_must_exist(self) -> None:
        repo = self.root / "repo"
        (repo / "tools").mkdir(parents=True)
        (repo / "tools" / "vba-projects.json").write_text(json.dumps(
            {"document_modules": [], "configurations": {"production": {"modules": ["src/core/Missing.bas"]}}}),
            encoding="utf-8")
        result = self.pwsh(f"$failed = ''\ntry {{ $null = Get-ManifestSources -RepoPath '{repo}' -Mode Production }} "
                           "catch { $failed = $_.Exception.Message }\n"
                           "@{ failed = $failed } | ConvertTo-Json | Set-Content -LiteralPath $env:RESULT -Encoding utf8")
        self.assertIn("not found: src/core/Missing.bas", result["failed"])


if __name__ == "__main__":
    unittest.main()
