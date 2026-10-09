"""A small loopback-only web UI for the existing validated pipeline."""
import base64
import binascii
import json
import os
import secrets
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlsplit, parse_qs
from uuid import uuid4
from zipfile import ZipFile, BadZipFile, ZIP_DEFLATED

from .config import load_config
from .errors import ReportError
from .layout import discover, approve
from .pipeline import run
from .word import WordTemplate

MAX_FILE = 10 * 1024 * 1024
MAX_BODY = 30 * 1024 * 1024
WEB = Path(__file__).parent / "web"


def unpack_file(item, suffixes):
    if not isinstance(item, dict) or set(item) != {"name", "data"}:
        raise ReportError("请选择有效文件")
    name = item["name"]
    if not isinstance(name, str) or Path(name).suffix.lower() not in suffixes:
        raise ReportError("文件格式不符，请选择 " + " / ".join(suffixes))
    try:
        if not isinstance(item["data"], str) or len(item["data"]) > MAX_FILE * 4 // 3 + 4:
            raise ValueError()
        data = base64.b64decode(item["data"], validate=True)
    except (ValueError, binascii.Error):
        raise ReportError("文件数据无效或超过 10 MB") from None
    if not data or len(data) > MAX_FILE:
        raise ReportError("文件为空或超过 10 MB")
    return data


def check_zip(path):
    try:
        with ZipFile(path) as archive:
            infos = archive.infolist()
            if len(infos) > 3000 or sum(i.file_size for i in infos) > 100 * 1024 * 1024:
                raise ReportError("文件解压内容过大，请简化文件后重试")
    except BadZipFile:
        raise ReportError("文件不是有效的 Excel / Word 文档") from None


class App:
    def __init__(self, project, data_dir):
        self.project = Path(project).resolve()
        self.data = Path(data_dir).resolve()
        self.data.mkdir(parents=True, exist_ok=True)
        self.templates_dir = self.data / "templates"
        self.templates_dir.mkdir(exist_ok=True)
        self.token = secrets.token_urlsafe(32)
        self.lock = threading.RLock()
        self.jobs = {}
        self.templates = {"example": {"id": "example", "name": "资产负债表会计报告 · 示例",
            "description": "配套示例表：两期金额 · 32 个字段 · 金额显示为万元",
            "word": self.project / "examples/report_template.docx",
            "config": self.project / "examples/mapping.json", "example": True}}
        if not all(self.templates["example"][key].is_file() for key in ("word", "config")):
            raise ReportError("请在完整项目目录中启动界面，未找到 examples 示例模板")
        for directory in self.templates_dir.iterdir():
            if not directory.is_dir() or len(directory.name) != 32 or any(c not in "0123456789abcdef" for c in directory.name):
                continue
            try:
                meta = json.loads((directory / "template.json").read_text(encoding="utf-8"))
                load_config(directory / "mapping.json")
                WordTemplate(directory / "report.docx")
                self.templates[directory.name] = {"id": directory.name, "name": str(meta["name"]),
                    "description": "自定义模板 · 已保存到本机", "word": directory / "report.docx",
                    "config": directory / "mapping.json", "example": False}
            except (OSError, ValueError, KeyError, ReportError, BadZipFile):
                continue

    def catalogue(self):
        with self.lock:
            return {"templates": [{k: v for k, v in item.items() if k not in ("word", "config")}
                                  for item in self.templates.values()],
                    "ai_ready": bool(os.environ.get("DEEPSEEK_API_KEY")), "max_file_mb": 10}

    def add_template(self, payload):
        name = payload.get("name", "")
        if not isinstance(name, str):
            raise ReportError("模板名称格式无效")
        name = name.strip()
        if not name or len(name) > 60:
            raise ReportError("请填写 1–60 字的模板名称")
        word = unpack_file(payload.get("word"), (".docx",))
        config = unpack_file(payload.get("config"), (".json",))
        tid = uuid4().hex
        directory = self.templates_dir / tid
        directory.mkdir()
        (directory / "report.docx").write_bytes(word)
        (directory / "mapping.json").write_bytes(config)
        check_zip(directory / "report.docx")
        load_config(directory / "mapping.json")
        WordTemplate(directory / "report.docx")
        (directory / "template.json").write_text(json.dumps({"name": name}, ensure_ascii=False), encoding="utf-8")
        with self.lock:
            self.templates[tid] = {"id": tid, "name": name, "description": "自定义模板 · 已保存到本机",
                "word": directory / "report.docx", "config": directory / "mapping.json", "example": False}
        return {"id": tid, "name": name}

    def generate(self, payload):
        tid = payload.get("template")
        with self.lock:
            template = self.templates.get(tid) if isinstance(tid, str) else None
        if not template:
            raise ReportError("请选择可用的报告模板")
        if type(payload.get("demo", False)) is not bool or type(payload.get("use_ai", False)) is not bool:
            raise ReportError("运行选项格式无效")
        data = (self.project / "examples/balance_sheet.xlsx").read_bytes() if payload.get("demo") else unpack_file(payload.get("excel"), (".xlsx", ".xlsm"))
        if payload.get("use_ai") and not os.environ.get("DEEPSEEK_API_KEY"):
            raise ReportError("本机尚未设置 DEEPSEEK_API_KEY；可关闭 AI，按模板直接生成")
        jid = uuid4().hex
        directory = self.data / "runs" / jid
        directory.mkdir(parents=True)
        excel = directory / "source.xlsx"
        excel.write_bytes(data)
        check_zip(excel)
        # Snapshot the selected template: later registrations cannot change this job.
        word = directory / "template.docx"
        config = directory / "mapping.json"
        word.write_bytes(template["word"].read_bytes())
        config.write_bytes(template["config"].read_bytes())
        job = {"id": jid, "directory": directory, "lock": threading.Lock(), "status": "running"}
        with self.lock:
            self.jobs[jid] = job
        if payload.get("use_ai"):
            proposal = discover(excel, config)
            (directory / "proposal.json").write_text(json.dumps(proposal, ensure_ascii=False, indent=2), encoding="utf-8")
            job["status"] = "needs_review"
            labels = {f["id"]: f["source"]["labels"][0] for f in proposal["candidate_config"]["fields"]}
            periods = {f["id"]: f["source"].get("period", "") for f in load_config(config)["fields"]}
            return {"id": jid, "status": "needs_review", "matches": [{**item, "label": labels[item["field"]], "period": periods[item["field"]]} for item in proposal["answer"]["matches"]],
                    "sheet": proposal["answer"]["sheet"]}
        return self.finish(job, config)

    def finish(self, job, config):
        directory = job["directory"]
        result = run(directory / "source.xlsx", directory / "template.docx", config, directory / "output")
        job["status"] = result["status"]
        files = ["audit.json", "audit.csv"]
        if result["status"] == "completed":
            files.insert(0, "report.docx")
            with ZipFile(directory / "output/report-and-audit.zip", "w", ZIP_DEFLATED) as archive:
                for name in files:
                    archive.write(directory / "output" / name, name)
            files.append("report-and-audit.zip")
        return {"id": job["id"], "status": result["status"], "errors": result["errors"], "warnings": result["warnings"],
                "records": len(result["records"]), "checks": len(result["checks"]), "files": files}

    def confirm(self, payload):
        if payload.get("accept") is not True:
            raise ReportError("请先核对并确认科目映射")
        with self.lock:
            job = self.jobs.get(payload.get("id")) if isinstance(payload.get("id"), str) else None
        if not job:
            raise ReportError("识别结果已失效，请重新上传")
        if not job["lock"].acquire(blocking=False):
            raise ReportError("正在生成，请勿重复提交")
        try:
            if job["status"] != "needs_review":
                raise ReportError("该识别结果已处理，请重新上传")
            directory = job["directory"]
            approved = directory / "approved-mapping.json"
            job["status"] = "running"
            try:
                approve(directory / "source.xlsx", directory / "mapping.json", directory / "proposal.json",
                        directory / "template.docx", approved, directory / "preflight")
            except ReportError as exc:
                job["status"] = "failed"
                audit_path = directory / "preflight/audit.json"
                errors = json.loads(audit_path.read_text(encoding="utf-8")).get("errors", []) if audit_path.is_file() else []
                return {"id": job["id"], "status": "failed", "errors": [str(exc), *errors],
                        "warnings": [], "files": ["audit.json", "audit.csv"] if audit_path.is_file() else [], "records": 0, "checks": 0}
            return self.finish(job, approved)
        finally:
            job["lock"].release()


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass  # Do not log upload data, tokens, private paths or download URLs.

    @property
    def app(self):
        return self.server.app

    def send(self, status, body, content_type="application/json; charset=utf-8"):
        data = json.dumps(body, ensure_ascii=False).encode("utf-8") if isinstance(body, dict) else body
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("X-Frame-Options", "DENY")
        self.send_header("Content-Security-Policy", "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; connect-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'self'")
        self.end_headers()
        self.wfile.write(data)

    def allowed_host(self):
        port = self.server.server_port
        return self.headers.get("Host") in (f"127.0.0.1:{port}", f"localhost:{port}")

    def do_GET(self):
        if not self.allowed_host():
            return self.send(403, {"error": "无效的本机访问地址"})
        route = urlsplit(self.path)
        if route.path == "/":
            html = (WEB / "index.html").read_text(encoding="utf-8").replace("__TOKEN__", self.app.token)
            return self.send(200, html.encode("utf-8"), "text/html; charset=utf-8")
        if route.path in ("/app.js", "/style.css"):
            kind = "text/javascript" if route.path.endswith(".js") else "text/css"
            return self.send(200, (WEB / route.path[1:]).read_bytes(), kind + "; charset=utf-8")
        if route.path == "/api/templates":
            return self.send(200, self.app.catalogue())
        if route.path.startswith("/download/"):
            parts = route.path.split("/")
            if len(parts) != 4 or not secrets.compare_digest(parse_qs(route.query).get("token", [""])[0].encode(), self.app.token.encode()):
                return self.send(403, {"error": "下载链接已失效"})
            with self.app.lock:
                job = self.app.jobs.get(parts[2])
            if not job or parts[3] not in ("report.docx", "audit.json", "audit.csv", "report-and-audit.zip"):
                return self.send(404, {"error": "文件不存在"})
            path = job["directory"] / "output" / parts[3]
            if not path.is_file() and parts[3] in ("audit.json", "audit.csv"):
                path = job["directory"] / "preflight" / parts[3]
            if not path.is_file():
                return self.send(404, {"error": "文件未生成"})
            data = path.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "application/octet-stream")
            self.send_header("Content-Disposition", f'attachment; filename="{parts[3]}"')
            self.send_header("Content-Length", str(len(data)))
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-Content-Type-Options", "nosniff")
            self.end_headers()
            self.wfile.write(data)
            return
        self.send(404, {"error": "页面不存在"})

    def do_POST(self):
        port = self.server.server_port
        origin = self.headers.get("Origin")
        if (not self.allowed_host() or not secrets.compare_digest(self.headers.get("X-App-Token", "").encode(), self.app.token.encode())
                or (origin is not None and origin not in (f"http://127.0.0.1:{port}", f"http://localhost:{port}"))):
            return self.send(403, {"error": "页面已失效，请刷新后重试"})
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if not 0 < length <= MAX_BODY:
                raise ReportError("上传内容为空或过大，每个文件最多 10 MB")
            if self.headers.get_content_type() != "application/json":
                raise ReportError("请求格式无效")
            payload = json.loads(self.rfile.read(length))
            if not isinstance(payload, dict):
                raise ReportError("请求格式无效")
            actions = {"/api/template": self.app.add_template, "/api/run": self.app.generate, "/api/approve": self.app.confirm}
            action = actions.get(self.path)
            if not action:
                return self.send(404, {"error": "操作不存在"})
            self.send(200, action(payload))
        except (ReportError, ValueError, TypeError, UnicodeError) as exc:
            self.send(400, {"error": str(exc) if isinstance(exc, ReportError) else "文件或请求格式无效"})
        except Exception:
            self.send(500, {"error": "处理失败，请检查文件是否可读，并确认使用完整的项目文件"})


def create_server(project, data_dir, port=8765):
    app = App(project, data_dir)
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    server.app = app
    server.daemon_threads = True
    return server


def serve(project=None, data_dir=None, port=8765, open_browser=True):
    project = Path(project or Path.cwd()).resolve()
    server = create_server(project, data_dir or project / "results/ui", port)
    url = f"http://127.0.0.1:{server.server_port}"
    print(f"会计报告界面已启动：{url}\n文件保存在 {server.app.data}\n按 Ctrl+C 关闭。", flush=True)
    if open_browser:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
