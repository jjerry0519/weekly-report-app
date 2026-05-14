from __future__ import annotations

import datetime as dt
import json
import os
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import server  # noqa: E402


def check_template() -> dict[str, object]:
    template = server.TEMPLATE_PATH
    required = [
        "xl/workbook.xml",
        "xl/sharedStrings.xml",
        "xl/worksheets/sheet1.xml",
        "xl/worksheets/sheet2.xml",
        "[Content_Types].xml",
        "xl/_rels/workbook.xml.rels",
    ]
    result: dict[str, object] = {"exists": template.exists(), "missing": [], "xml": {}}
    if not template.exists():
        return result
    with zipfile.ZipFile(template) as zf:
        names = set(zf.namelist())
        result["missing"] = [name for name in required if name not in names]
        xml_result: dict[str, str] = {}
        for name in required:
            if name not in names:
                continue
            try:
                ET.fromstring(zf.read(name))
                xml_result[name] = "ok"
            except Exception as exc:  # pragma: no cover - diagnostic output
                xml_result[name] = str(exc)
        result["xml"] = xml_result
    return result


def check_optional_sample() -> dict[str, object]:
    sample = os.environ.get("SAMPLE_SOURCE_XLSX", "").strip()
    if not sample:
        return {"skipped": True, "reason": "SAMPLE_SOURCE_XLSX not set"}
    source = Path(sample)
    if not source.exists():
        return {"skipped": True, "reason": f"sample not found: {source}"}
    start = dt.date.fromisoformat(os.environ.get("SAMPLE_START", "2026-05-08"))
    end = dt.date.fromisoformat(os.environ.get("SAMPLE_END", "2026-05-14"))
    original_report_dir = server.REPORT_DIR
    with tempfile.TemporaryDirectory() as tmp:
        try:
            server.REPORT_DIR = Path(tmp)
            server.ensure_dirs()
            output = server.build_report(source, start, end, "self-check")
            path = server.REPORT_DIR / str(output["file"])
            with zipfile.ZipFile(path) as zf:
                parsed = {}
                for name in ("xl/workbook.xml", "xl/sharedStrings.xml", "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"):
                    ET.fromstring(zf.read(name))
                    parsed[name] = "ok"
        finally:
            server.REPORT_DIR = original_report_dir
    return {"skipped": False, "output": output, "xml": parsed}


def main() -> int:
    checks = {
        "template": check_template(),
        "default_week": [d.isoformat() for d in server.default_week(dt.date(2026, 5, 14))],
        "optional_sample": check_optional_sample(),
    }
    failures: list[str] = []
    template = checks["template"]
    if not template["exists"]:
        failures.append("template file missing")
    if template["missing"]:
        failures.append(f"template missing parts: {template['missing']}")
    bad_xml = {name: status for name, status in template["xml"].items() if status != "ok"}
    if bad_xml:
        failures.append(f"template XML parse failures: {bad_xml}")
    print(json.dumps(checks, ensure_ascii=False, indent=2))
    if failures:
        print("SELF CHECK FAILED:")
        for failure in failures:
            print(f"- {failure}")
        return 1
    print("SELF CHECK OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
