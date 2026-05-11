#!/usr/bin/env python3
"""一鍵重跑帕妃全套輸出，並彙整所有 audit 缺漏到 輸出檔/待補資料.txt。

目的：避免只有某一支程式寫待補資料，導致其他輸出檔缺資料沒有被記錄。
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import shutil
import subprocess
import sys
from collections import defaultdict
from pathlib import Path
from typing import Any

DEFAULT_BASE = Path("/Users/openclaw/創旭/帕妃")
PYTHON = Path("/opt/homebrew/bin/python3")

FLOW_NAME = "帕妃自動產品產生檔"

SCRIPTS = [
    "generate_live_table.py",
    "generate_91_listing.py",
    "generate_website_radio.py",
    "generate_website_color_blocks_file3.py",
    "generate_website_middle_images_file5.py",
    "generate_website_hidden_category_file4.py",
    "generate_website_category_file4.py",
    "generate_website_tags_file7.py",
    "generate_website_recommend_pending.py",
]

SCRIPT_FLOW_NAMES = {
    "generate_live_table.py": "直播表格",
    "generate_91_listing.py": "91 新品上架檔",
    "generate_website_radio.py": "官網檔案2 名稱+大小",
    "generate_website_color_blocks_file3.py": "官網檔案3 色塊",
    "generate_website_middle_images_file5.py": "官網檔案5 中圖合圖",
    "generate_website_hidden_category_file4.py": "官網檔案4 隱藏分類",
    "generate_website_category_file4.py": "官網檔案4 上架分類",
    "generate_website_tags_file7.py": "官網檔案7 洗滌+通用+標籤",
    "generate_website_recommend_pending.py": "官網推薦待填寫",
}


def norm(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def ensure_big_file_alias(base: Path) -> None:
    input_dir = base / "輸入檔"
    alias = input_dir / "商品資料大檔.xlsx"
    real = input_dir / "商品資料(大檔).xlsx"
    if alias.exists() or alias.is_symlink():
        return
    if not real.exists():
        raise FileNotFoundError(f"找不到輸入大檔：{real}")
    alias.symlink_to(real.name)


def run_generators(base: Path, date_code: str, year: str) -> list[dict[str, Any]]:
    code_dir = base / "程式"
    runs: list[dict[str, Any]] = []
    for script in SCRIPTS:
        cmd = [str(PYTHON), script, "--base", str(base), "--date", date_code]
        if script == "generate_91_listing.py" or script in {
            "generate_website_color_blocks_file3.py",
            "generate_website_middle_images_file5.py",
        }:
            cmd += ["--year", year]
        print(f"=== RUN {script}")
        completed = subprocess.run(cmd, cwd=code_dir, text=True, capture_output=True)
        runs.append({
            "script": script,
            "returncode": completed.returncode,
            "stdout": completed.stdout,
            "stderr": completed.stderr,
        })
        if completed.stdout:
            print(completed.stdout, end="")
        if completed.stderr:
            print(completed.stderr, file=sys.stderr, end="")
        if completed.returncode != 0:
            raise RuntimeError(f"{script} 執行失敗，returncode={completed.returncode}")
    return runs


def item_code(item: dict[str, Any]) -> str:
    for key in ("code", "物品編號", "流水號", "sku", "product_code"):
        value = norm(item.get(key))
        if value:
            return value
    # Some website file audits store row-level missing without code; keep row as identity fallback.
    row = norm(item.get("row"))
    return f"未標流水號(row {row})" if row else "未標流水號"


def normalize_missing_item(item: dict[str, Any], audit_path: Path, audit: dict[str, Any]) -> dict[str, str]:
    field = norm(item.get("field") or item.get("missing") or item.get("column") or item.get("欄位") or "未指定欄位")
    reason = norm(item.get("reason") or item.get("原因") or "來源資料缺漏，導致輸出資料不完整")
    source = norm(item.get("source") or item.get("來源"))
    if not source:
        source = infer_source(field, audit)
    output = norm(audit.get("output") or audit_path.with_suffix(""))
    return {
        "code": item_code(item),
        "field": field,
        "reason": reason,
        "source": source,
        "output": output,
        "audit": str(audit_path),
    }


def infer_source(field: str, audit: dict[str, Any]) -> str:
    script = Path(norm(audit.get("script"))).name
    if "成本" in field:
        return "商品資料大檔.xlsx → 商品資料(大檔)：K欄 ray成本需有值"
    if "適合" in field or "推薦" in field or "尺寸" in field:
        return "尺寸表&試穿報告.xlsx：對應流水號的尺寸/適合尺寸欄需有值"
    if "顏色" in field or "色塊" in field:
        return "商品資料大檔.xlsx → 顏色尺寸請用這一個：物品編號、商店顏色、色塊/圖片資料需完整"
    if "價格" in field or "原價" in field or "新品價" in field or "直播" in field:
        return "商品資料大檔.xlsx → 商品資料(大檔)：對應流水號的價格欄位需有值"
    if script:
        return f"請檢查 {script} 對應輸入來源欄位"
    return "請檢查對應輸入來源欄位"


def collect_audit_missing(output_dir: Path) -> list[dict[str, str]]:
    items: list[dict[str, str]] = []
    for audit_path in sorted(output_dir.rglob("*.audit.json")):
        try:
            audit = json.loads(audit_path.read_text(encoding="utf-8"))
        except Exception as exc:
            items.append({
                "code": "未標流水號",
                "field": "audit讀取失敗",
                "reason": f"{audit_path} 無法讀取：{exc}",
                "source": str(audit_path),
                "output": "",
                "audit": str(audit_path),
            })
            continue
        raw_missing = []
        for key in ("missing_or_pending", "missing", "pending"):
            value = audit.get(key)
            if isinstance(value, list):
                raw_missing.extend(value)
        for raw in raw_missing:
            if isinstance(raw, dict):
                items.append(normalize_missing_item(raw, audit_path, audit))
            else:
                items.append({
                    "code": "未標流水號",
                    "field": "未指定欄位",
                    "reason": norm(raw),
                    "source": infer_source("", audit),
                    "output": norm(audit.get("output")),
                    "audit": str(audit_path),
                })
    return dedupe_missing(items)


def dedupe_missing(items: list[dict[str, str]]) -> list[dict[str, str]]:
    seen: set[tuple[str, str, str, str, str]] = set()
    result: list[dict[str, str]] = []
    for item in items:
        key = (item["code"], item["field"], item["reason"], item["source"], item["output"])
        if key in seen:
            continue
        seen.add(key)
        result.append(item)
    return result


def write_missing_report(path: Path, date_code: str, items: list[dict[str, str]], generated_at: str) -> None:
    lines = [
        "待補資料",
        "=" * 40,
        f"產生時間：{generated_at}",
        f"日期批次：{date_code}",
        "說明：這份檔案由 run_all_pafei_outputs.py 統一彙整所有 .audit.json 的缺漏；任何輸出檔缺資料都應記錄在這裡。",
        "重要：商品辨識一律用流水號 / 物品編號；顏色只是要補的欄位，不用顏色去認列流水號。",
        "廠商欄位對應商品資料(大檔) H欄，允許原始空白，不列為錯誤。",
        "",
    ]
    if not items:
        lines += ["目前沒有 audit 記錄到缺資料。", ""]
    else:
        by_code: dict[str, list[dict[str, str]]] = defaultdict(list)
        for item in items:
            by_code[item["code"]].append(item)
        lines += ["一、依流水號整理需補欄位", "-" * 40]
        for i, code in enumerate(sorted(by_code), start=1):
            group = by_code[code]
            fields = "、".join(dict.fromkeys(item["field"] for item in group if item["field"]))
            lines += [f"{i}. 流水號：{code}", f"   需要補：{fields or '未指定欄位'}"]
            for item in group:
                lines += [
                    f"   - {item['field']}：{item['reason']}",
                    f"     補資料位置：{item['source']}",
                    f"     影響輸出：{item['output']}",
                ]
            lines.append("")

        by_source: dict[str, list[dict[str, str]]] = defaultdict(list)
        for item in items:
            by_source[item["source"]].append(item)
        lines += ["二、依原始來源彙總", "-" * 40]
        for i, source in enumerate(sorted(by_source), start=1):
            group = by_source[source]
            fields = "、".join(dict.fromkeys(item["field"] for item in group if item["field"]))
            codes = "、".join(dict.fromkeys(item["code"] for item in group if item["code"]))
            outputs = "、".join(dict.fromkeys(Path(item["output"]).name for item in group if item["output"]))
            lines += [
                f"{i}. 來源：{source}",
                f"   影響欄位：{fields}",
                f"   影響流水號：{codes}",
                f"   影響輸出：{outputs}",
                "",
            ]
    path.write_text("\n".join(lines), encoding="utf-8")


def summarize_outputs(base: Path) -> dict[str, Any]:
    # Keep lightweight summary only; detailed row-count verification can be done externally.
    output_dir = base / "輸出檔"
    return {
        "xlsx_count": len(list(output_dir.rglob("*.xlsx"))),
        "audit_count": len(list(output_dir.rglob("*.audit.json"))),
        "missing_report": str(output_dir / "待補資料.txt"),
        "missing_report_sha256": sha256(output_dir / "待補資料.txt") if (output_dir / "待補資料.txt").exists() else "",
    }


def audit_by_script(output_dir: Path) -> dict[str, dict[str, Any]]:
    audits: dict[str, dict[str, Any]] = {}
    for audit_path in sorted(output_dir.rglob("*.audit.json")):
        try:
            audit = json.loads(audit_path.read_text(encoding="utf-8"))
        except Exception:
            continue
        script_name = Path(norm(audit.get("script"))).name
        if script_name:
            audits[script_name] = audit
    return audits


def write_workflow_code_bundle(base: Path, date_code: str, year: str, generated_at: str) -> Path:
    """輸出單一流程總檔：流程名稱、每個輸出檔、對應程式碼。"""
    code_dir = base / "程式"
    output_dir = base / "輸出檔"
    audits = audit_by_script(output_dir)
    path = output_dir / "帕妃自動產品產生檔-流程與程式碼總檔.md"
    lines: list[str] = [
        f"# {FLOW_NAME}",
        "",
        f"產生時間：{generated_at}",
        f"日期批次：{date_code}",
        f"年份：{year}",
        "",
        "## 檔案輸出總表",
        "",
        "| 順序 | 流程名稱 | 程式碼檔案 | 輸出檔案 | 輸出筆數 | 缺漏項目數 |",
        "|---:|---|---|---|---:|---:|",
    ]
    for index, script in enumerate(SCRIPTS, start=1):
        audit = audits.get(script, {})
        output = norm(audit.get("output"))
        row_count = norm(audit.get("row_count"))
        missing_count = 0
        for key in ("missing_or_pending", "missing", "pending"):
            value = audit.get(key)
            if isinstance(value, list):
                missing_count += len(value)
        lines.append(
            f"| {index} | {SCRIPT_FLOW_NAMES.get(script, script)} | `程式/{script}` | `{output}` | {row_count or 0} | {missing_count} |"
        )
    lines += [
        "",
        "## 統一輸出檔",
        "",
        f"- 待補資料：`{output_dir / '待補資料.txt'}`",
        f"- 總執行紀錄：`{output_dir / 'run_all_pafei_outputs.audit.json'}`",
        f"- 本流程總檔：`{path}`",
        "",
        "## 程式碼內容",
        "",
    ]
    for index, script in enumerate(SCRIPTS, start=1):
        script_path = code_dir / script
        lines += [
            f"### {index}. {SCRIPT_FLOW_NAMES.get(script, script)} — `{script}`",
            "",
        ]
        if not script_path.exists():
            lines += ["```text", f"找不到程式碼檔案：{script_path}", "```", ""]
            continue
        lines += ["```python", script_path.read_text(encoding="utf-8"), "```", ""]
    path.write_text("\n".join(lines), encoding="utf-8")
    return path


def main() -> None:
    parser = argparse.ArgumentParser(description="重跑帕妃全套輸出並彙整待補資料")
    parser.add_argument("--base", default=str(DEFAULT_BASE), help="帕妃資料夾根目錄")
    parser.add_argument("--date", default="0414", help="日期 MMDD，例如 0414")
    parser.add_argument("--year", default="2026", help="年份，例如 2026")
    parser.add_argument("--keep-output", action="store_true", help="不清空輸出檔，直接覆蓋產出")
    args = parser.parse_args()

    base = Path(args.base)
    output_dir = base / "輸出檔"
    generated_at = dt.datetime.now().isoformat(timespec="seconds")

    ensure_big_file_alias(base)
    if not args.keep_output and output_dir.exists():
        shutil.rmtree(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    runs = run_generators(base, args.date, args.year)
    missing_items = collect_audit_missing(output_dir)
    report_path = output_dir / "待補資料.txt"
    write_missing_report(report_path, args.date, missing_items, generated_at)
    workflow_code_bundle_path = write_workflow_code_bundle(base, args.date, args.year, generated_at)

    summary = {
        "generated_at": generated_at,
        "base": str(base),
        "flow_name": FLOW_NAME,
        "workflow_code_bundle": str(workflow_code_bundle_path),
        "date_code": args.date,
        "year": args.year,
        "runs": runs,
        "missing_count": len(missing_items),
        "missing_report": str(report_path),
        **summarize_outputs(base),
    }
    summary_path = output_dir / "run_all_pafei_outputs.audit.json"
    summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"已彙整待補資料：{report_path}")
    print(f"總執行紀錄：{summary_path}")
    print(f"流程與程式碼總檔：{workflow_code_bundle_path}")
    print(f"缺漏項目數：{len(missing_items)}")


if __name__ == "__main__":
    main()
