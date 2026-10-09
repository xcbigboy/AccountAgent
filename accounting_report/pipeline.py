import csv
import hashlib
import json
import os
from datetime import datetime, timezone
from pathlib import Path

from . import __version__
from .config import load_config
from .errors import ReportError
from .extract import ExcelSource
from .numbers import display
from .validate import validate_data
from .word import WordTemplate


def digest(path):
    h = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def write_audit(directory, audit):
    (directory / "audit.json").write_text(json.dumps(audit, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    columns = ["id", "label", "sheet", "label_cell", "cell", "period", "raw", "formula",
               "value", "source_unit", "output_unit", "converted", "rounded", "formatted", "blank_as_zero"]
    with (directory / "audit.csv").open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=columns, extrasaction="ignore")
        writer.writeheader()
        for record in audit["records"]:
            row = dict(record)
            # Keep source formula strings inspectable without executing in Excel.
            for key, value in row.items():
                if isinstance(value, str) and value.lstrip().startswith(("=", "+", "-", "@")):
                    row[key] = "'" + value
            writer.writerow(row)


def run(excel, template, config_path, output_dir, dry_run=False):
    directory = Path(output_dir).resolve()
    if directory.exists():
        raise ReportError(f"输出目录已存在：{directory}；请指定新的目录，避免覆盖历史结果")
    directory.mkdir(parents=True, exist_ok=False)
    audit = {"version": __version__, "started_at_utc": datetime.now(timezone.utc).isoformat(),
             "status": "running", "dry_run": dry_run, "inputs": {}, "records": [],
             "checks": [], "word_changes": [], "warnings": [], "errors": []}
    source = None
    pending = directory / "report.pending.docx"
    try:
        paths = {"excel": Path(excel).resolve(), "template": Path(template).resolve(), "config": Path(config_path).resolve()}
        if paths["excel"].suffix.lower() not in [".xlsx", ".xlsm"]:
            raise ReportError("仅支持 .xlsx/.xlsm，旧 .xls 请另存为 .xlsx")
        if paths["template"].suffix.lower() != ".docx":
            raise ReportError("Word 模板必须是 .docx，旧 .doc 请另存为 .docx")
        for name, path in paths.items():
            audit["inputs"][name] = {"path": str(path), "sha256": digest(path)}
        config = load_config(paths["config"])
        audit["source_unit"] = config["source"]["unit"]
        source = ExcelSource(paths["excel"], config)
        for field in config["fields"]:
            try:
                record = source.read(field)
                fmt = config.get("format", {})
                unit = fmt.get("unit", config["source"]["unit"])
                text, converted, rounded = display(record["value"], config["source"]["unit"], unit,
                                                   fmt.get("decimals", 2), fmt.get("grouping", True), fmt.get("negative", "minus"))
                record.update({"formatted": text, "converted": converted, "rounded": rounded,
                               "source_unit": config["source"]["unit"], "output_unit": unit})
                audit["records"].append(record)
                if record["formula"]:
                    audit["warnings"].append(f"{field['id']} 使用 Excel 缓存的公式值，程序不重算公式；需确认源文件已重算并保存")
                if record["blank_as_zero"]:
                    audit["warnings"].append(f"{field['id']} 按显式 blank_zero 配置将空值填为零")
            except ReportError as exc:
                audit["errors"].append(f"{field['id']}：{exc}")
        audit["checks"], validation_errors = validate_data(config, audit["records"])
        audit["errors"].extend(validation_errors)
        if audit["errors"]:
            raise ReportError("数据验证失败")
        word = WordTemplate(paths["template"])
        audit["word_changes"] = word.fill(config, audit["records"])
        for name, path in paths.items():
            if digest(path) != audit["inputs"][name]["sha256"]:
                raise ReportError(f"运行期间输入文件发生变化：{name}，请重试")
        if not dry_run:
            word.save(pending)
            # Re-open the saved package: no incomplete report gets published.
            if WordTemplate(pending).inventory()["placeholders"]:
                raise ReportError("保存后的 Word 仍有未填占位符")
            audit["report_sha256"] = digest(pending)
        audit["status"] = "validated" if dry_run else "completed"
    except Exception as exc:
        audit["status"] = "failed"
        if str(exc) != "数据验证失败":
            audit["errors"].append(f"{type(exc).__name__}：{exc}")
        pending.unlink(missing_ok=True)
    finally:
        if source is not None:
            source.close()
        audit["finished_at_utc"] = datetime.now(timezone.utc).isoformat()
        try:
            write_audit(directory, audit)
        except Exception:
            pending.unlink(missing_ok=True)
            raise
    if audit["status"] == "completed":
        try:
            os.replace(pending, directory / "report.docx")
        except OSError as exc:
            audit["status"] = "failed"
            audit["errors"].append(f"报告发布失败：{type(exc).__name__}：{exc}")
            audit.pop("report_sha256", None)
            pending.unlink(missing_ok=True)
            write_audit(directory, audit)
    return audit
