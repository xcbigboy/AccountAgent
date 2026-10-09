from datetime import datetime
from pathlib import Path
from uuid import uuid4
from accounting_report.cli import main

root = Path(__file__).resolve().parent
output = root / "results" / (datetime.now().strftime("%Y%m%d-%H%M%S") + "-" + uuid4().hex[:8])
raise SystemExit(main(["run", "--excel", str(root / "examples/balance_sheet.xlsx"),
                      "--template", str(root / "examples/report_template.docx"),
                      "--config", str(root / "examples/mapping.json"), "--output-dir", str(output)]))
