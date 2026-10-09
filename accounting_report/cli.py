import argparse
import json
import sys
from pathlib import Path

from .ai import cloud_suggestions, local_suggestions
from .config import load_config
from .errors import ReportError
from .extract import ExcelSource
from .pipeline import run
from .word import WordTemplate


def main(argv=None):
    parser = argparse.ArgumentParser(description="Excel 资产负债表 → Word 会计报告，先验证再填报")
    sub = parser.add_subparsers(dest="command", required=True)
    execute = sub.add_parser("run", help="验证数据和模板并填报")
    execute.add_argument("--excel", required=True)
    execute.add_argument("--template", required=True)
    execute.add_argument("--config", required=True)
    execute.add_argument("--output-dir", required=True, help="新的、不存在的输出目录")
    execute.add_argument("--dry-run", action="store_true", help="完整验证并输出日志，不生成 Word")
    inspect = sub.add_parser("inspect-word", help="列出 Word 占位符及表格单元格原文")
    inspect.add_argument("--template", required=True)
    for name, help_text in [("inspect-excel", "列出配置范围中的科目及坐标"), ("suggest", "生成候选匹配建议，不能自动应用")]:
        p = sub.add_parser(name, help=help_text)
        p.add_argument("--excel", required=True)
        p.add_argument("--config", required=True)
        if name == "suggest":
            p.add_argument("--ai", action="store_true", help="把科目名发送至指定 AI 服务；默认只用本地文字相似度")
            p.add_argument("--endpoint", help="完整 HTTPS Chat Completions 地址")
            p.add_argument("--model", help="服务支持的模型名称")
            p.add_argument("--key-env", default="REPORT_AI_API_KEY")
            p.add_argument("--output", help="保存建议 JSON，路径须不存在")
    args = parser.parse_args(argv)
    try:
        if args.command == "run":
            result = run(args.excel, args.template, args.config, args.output_dir, args.dry_run)
            print(json.dumps({"status": result["status"], "output_dir": str(Path(args.output_dir).resolve()),
                              "records": len(result["records"]), "errors": result["errors"], "warnings": result["warnings"]}, ensure_ascii=False, indent=2))
            return 2 if result["status"] == "failed" else 0
        if args.command == "inspect-word":
            result = WordTemplate(args.template).inventory()
        else:
            config = load_config(args.config)
            source = ExcelSource(args.excel, config)
            try:
                labels = source.labels()
            finally:
                source.close()
            if args.command == "inspect-excel":
                result = labels
            elif args.ai:
                if not args.endpoint or not args.model:
                    raise ReportError("--ai 需要 --endpoint 和 --model")
                result = cloud_suggestions(config, labels, args.endpoint, args.model, args.key_env)
            else:
                result = local_suggestions(config, labels)
        text = json.dumps(result, ensure_ascii=False, indent=2, default=str)
        if getattr(args, "output", None):
            with Path(args.output).open("x", encoding="utf-8") as stream:
                stream.write(text)
        else:
            print(text)
        return 0
    except Exception as exc:
        print(f"失败：{type(exc).__name__}：{exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
