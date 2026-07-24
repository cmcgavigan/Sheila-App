"""Static checks on install/install.ps1 — parsed, never executed.

Guards the explicit-source contract: migration dry runs only happen for a
caller-supplied -V1Source; the installer must not infer or scan for v1.
"""
from __future__ import annotations

import re
import shutil
import subprocess
from pathlib import Path

import pytest

PS1 = Path(__file__).resolve().parents[1] / "install" / "install.ps1"


def _text() -> str:
    return PS1.read_text(encoding="utf-8")


def test_v1source_is_an_explicit_parameter():
    text = _text()
    assert re.search(r"^param\(", text, re.M), "param() block missing"
    assert "[string]$V1Source" in text
    # param() must precede all executable statements
    assert text.index("param(") < text.index("$ErrorActionPreference")


def test_migration_section_does_not_infer_the_source():
    text = _text()
    section = text[text.index("--- 3b."):text.index("--- 4.")]
    assert "Split-Path" not in section, "migration section must not derive a v1 path"
    assert "$V1Source" in section
    assert "--apply" not in section.split("Write-Host")[0], \
        "the installer itself must never run an apply"


def test_nssm_download_is_pinned_and_service_is_not_system():
    text = _text()
    assert "Get-FileHash" in text and "nssmSha256" in text
    assert "ObjectName 'LocalService'" in text
    assert "profile=private,domain" in text
    assert "[guid]::NewGuid" in text


def test_powershell_syntax_valid():
    exe = shutil.which("powershell") or shutil.which("pwsh")
    if not exe:
        pytest.skip("PowerShell not available on this machine")
    # Language.Parser only parses the file; nothing in it is executed.
    ps = (
        "$t=$null;$e=$null;"
        f"[System.Management.Automation.Language.Parser]::ParseFile('{PS1}',[ref]$t,[ref]$e)|Out-Null;"
        "if($e.Count -eq 0){exit 0}else{$e|ForEach-Object{$_.Message};exit 1}"
    )
    r = subprocess.run([exe, "-NoProfile", "-NonInteractive", "-Command", ps],
                       capture_output=True, text=True, timeout=120)
    assert r.returncode == 0, f"syntax errors:\n{r.stdout}{r.stderr}"
