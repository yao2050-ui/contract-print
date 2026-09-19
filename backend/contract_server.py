# -*- coding: utf-8 -*-
"""
合同生成后端服务（云端部署版）
POST /generate  body: {"record_id": "recXXX"}  -> 生成合同并回填附件
GET  /health    -> 健康检查
"""
import json
import mimetypes
import os
import re
import tempfile
import threading
import time
import traceback
from datetime import datetime
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import unquote

import requests

import generate_contracts as gc

# 内存锁：防止同一记录被并发重复生成
_generating_lock = threading.Lock()
_generating_records = set()

BASE_TOKEN = os.environ.get("FEISHU_BASE_TOKEN", "bascnOa5an1A3Oo6XBrj9JwVHqc")
TABLE_ID = os.environ.get("FEISHU_TABLE_ID", "tblRRSw6jaLD2IWi")
API_BASE = "https://base-api.feishu.cn/open-apis"
TEMPLATE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates")
STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")
PERSONAL_TOKEN = os.environ.get("FEISHU_PERSONAL_TOKEN", "")
HEADERS = {"Authorization": "Bearer " + PERSONAL_TOKEN}

FIELD_ID = {
    "企业名称": "fldOGL0k2F", "地址": "fldLoBMTRR", "房间号": "fldf3fYeNq",
    "使用面积": "fldkfx4Rpd", "地址费": "fldGePXm5s", "租赁起始日期": "fldBt9QBNz",
    "租赁截止日期": "fldbWBy240", "收款抬头": "fldIOKndZZ", "收款信息": "fldmryYlkb",
    "合同类型": "fldwtNRNge", "大写金额": "fldha0wB1Z", "租赁合同附件": "fld8RQfApb",
    "合同生成状态": "fldnyA7mTY",
}
ID2FIELD = {v: k for k, v in FIELD_ID.items()}


def fetch_record(record_id):
    url = "%s/bitable/v1/apps/%s/tables/%s/records/%s" % (API_BASE, BASE_TOKEN, TABLE_ID, record_id)
    r = requests.get(url, headers=HEADERS, timeout=30)
    r.raise_for_status()
    data = r.json()
    if data.get("code") != 0:
        raise RuntimeError("读取记录失败: %s" % data.get("msg"))
    fields_raw = data["data"]["record"]["fields"]
    rec = {"record_id": record_id}
    for fid, val in fields_raw.items():
        name = ID2FIELD.get(fid, fid)
        rec[name] = val
    return rec


def upload_attachment(record_id, file_path):
    file_name = os.path.basename(file_path)
    file_size = os.path.getsize(file_path)
    url = "%s/drive/v1/medias/upload_all" % API_BASE
    with open(file_path, "rb") as f:
        files = {"file": (file_name, f, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")}
        data = {
            "file_name": file_name,
            "parent_type": "bitable_file",
            "parent_node": BASE_TOKEN,
            "size": str(file_size),
        }
        r = requests.post(url, headers=HEADERS, data=data, files=files, timeout=60)
    r.raise_for_status()
    res = r.json()
    if res.get("code") != 0:
        raise RuntimeError("上传附件失败: %s" % res.get("msg"))
    return res["data"]["file_token"]


def set_attachment(record_id, file_token):
    url = "%s/bitable/v1/apps/%s/tables/%s/records/%s" % (API_BASE, BASE_TOKEN, TABLE_ID, record_id)
    body = {"fields": {"租赁合同附件": [{"file_token": file_token}]}}
    r = requests.put(url, headers=HEADERS, json=body, timeout=30)
    r.raise_for_status()
    res = r.json()
    if res.get("code") != 0:
        raise RuntimeError("回填附件失败: %s" % res.get("msg"))


def set_status(record_id, option):
    url = "%s/bitable/v1/apps/%s/tables/%s/records/%s" % (API_BASE, BASE_TOKEN, TABLE_ID, record_id)
    body = {"fields": {"合同生成状态": option}}
    r = requests.put(url, headers=HEADERS, json=body, timeout=30)
    r.raise_for_status()
    res = r.json()
    if res.get("code") != 0:
        raise RuntimeError("更新状态失败: %s" % res.get("msg"))


def generate_and_backfill(record_id):
    # 第一道防线：内存锁，彻底防止同一记录并发生成
    with _generating_lock:
        if record_id in _generating_records:
            return {
                "ok": True, "record_id": record_id,
                "skipped": True, "message": "正在生成中，跳过重复请求"
            }
        _generating_records.add(record_id)

    try:
        rec = fetch_record(record_id)
        ctype = gc.to_text(rec.get("合同类型"))
        if not ctype:
            raise RuntimeError("记录缺少合同类型")

        # 第二道防线：检查状态
        current_status = gc.to_text(rec.get("合同生成状态"))
        if current_status in ("已生成", "生成中"):
            return {
                "ok": True, "record_id": record_id,
                "企业名称": gc.to_text(rec.get("企业名称")),
                "合同类型": ctype,
                "skipped": True, "message": "状态为%s，跳过" % current_status,
            }

        # 第三道防线：检查附件字段是否已有内容
        existing_attach = rec.get("租赁合同附件")
        if existing_attach and isinstance(existing_attach, list) and len(existing_attach) > 0:
            try:
                set_status(record_id, "已生成")
            except Exception:
                pass
            return {
                "ok": True, "record_id": record_id,
                "企业名称": gc.to_text(rec.get("企业名称")),
                "合同类型": ctype,
                "skipped": True, "message": "已有附件，跳过",
            }

        tmpdir = tempfile.mkdtemp(prefix="contract_")
        base_name = re.sub(r'[\\/:*?"<>|]', "_", "%s-%s.docx" % (ctype, gc.to_text(rec.get("企业名称"))))
        out_path = os.path.join(tmpdir, base_name)

        # 标记为生成中
        set_status(record_id, "生成中")

        try:
            result = gc.generate_one(rec, TEMPLATE_DIR, out_path)
            if not result.get("ok"):
                set_status(record_id, "生成失败")
                raise RuntimeError("生成失败: %s" % result.get("原因"))

            file_token = upload_attachment(record_id, out_path)
            set_attachment(record_id, file_token)
            set_status(record_id, "已生成")
        except Exception as e:
            set_status(record_id, "生成失败")
            raise

        return {
            "ok": True,
            "record_id": record_id,
            "企业名称": result["企业名称"],
            "合同类型": ctype,
            "file_token": file_token,
            "替换处数": result["替换处数"],
            "未替换占位符": result["未替换占位符"],
        }
    finally:
        with _generating_lock:
            _generating_records.discard(record_id)


class Handler(BaseHTTPRequestHandler):
    def _send_json(self, code, obj):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self._send_json(200, {})

    def do_GET(self):
        if self.path == "/health":
            self._send_json(200, {"ok": True, "time": datetime.now().isoformat()})
            return
        # 静态文件服务
        path = unquote(self.path.split("?")[0])
        if path == "/":
            path = "/index.html"
        file_path = os.path.join(STATIC_DIR, path.lstrip("/"))
        if os.path.exists(file_path) and os.path.isfile(file_path):
            mime, _ = mimetypes.guess_type(file_path)
            with open(file_path, "rb") as f:
                data = f.read()
            self.send_response(200)
            self.send_header("Content-Type", mime or "application/octet-stream")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        self._send_json(404, {"error": "not found", "path": path})

    def do_POST(self):
        if self.path != "/generate":
            self._send_json(404, {"error": "not found"})
            return
        try:
            length = int(self.headers.get("Content-Length", 0))
            raw = self.rfile.read(length) if length else b"{}"
            payload = json.loads(raw.decode("utf-8"))
            record_id = payload.get("record_id") or payload.get("recordId")
            if not record_id:
                self._send_json(400, {"ok": False, "error": "缺少 record_id"})
                return
            result = generate_and_backfill(record_id)
            self._send_json(200, result)
        except Exception as e:
            traceback.print_exc()
            self._send_json(500, {"ok": False, "error": str(e)})

    def log_message(self, fmt, *args):
        print("[%s] %s" % (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), fmt % args))


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8765))
    server = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    print("合同生成服务启动: 0.0.0.0:%d" % port)
    server.serve_forever()
