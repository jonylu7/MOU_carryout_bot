"""
carry_out_automation/hitl_server.py

Human-in-the-Loop 審核伺服器（輕量 Flask）。
協調人打開瀏覽器連結即可預覽並核准 / 退回攜出申請單。

依賴：
    pip install flask
"""

from __future__ import annotations

import logging
import os
import threading
import time
import uuid
from pathlib import Path

from flask import Flask, abort, redirect, render_template_string, request, url_for

import config

logger = logging.getLogger(__name__)

app = Flask(__name__)

# ── 審核任務狀態儲存（記憶體，單機使用足夠）──
_pending: dict[str, dict] = {}
# { token: {"docx_path": ..., "info": {...}, "result": None | "approved" | "rejected", "comment": ""} }

# ─────────────────────────────────────────────────
# HTML 範本
# ─────────────────────────────────────────────────
REVIEW_PAGE = """
<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<title>攜出申請單審核</title>
<style>
  body { font-family: "Noto Sans TC", sans-serif; max-width: 800px; margin: 40px auto; padding: 0 20px; color: #333; }
  h1 { color: #1a5276; border-bottom: 2px solid #1a5276; padding-bottom: 8px; }
  table { width: 100%; border-collapse: collapse; margin: 20px 0; }
  th { background: #1a5276; color: white; padding: 10px; text-align: left; }
  td { padding: 10px; border: 1px solid #ddd; }
  tr:nth-child(even) { background: #f8f9fa; }
  .btn { padding: 12px 28px; border: none; border-radius: 6px; font-size: 16px; cursor: pointer; margin: 8px; }
  .approve { background: #27ae60; color: white; }
  .reject  { background: #e74c3c; color: white; }
  textarea { width: 100%; padding: 10px; border-radius: 6px; border: 1px solid #ccc; font-size: 14px; }
  .done { text-align: center; padding: 60px; font-size: 20px; }
  .files { background: #f0f3f4; padding: 12px; border-radius: 6px; white-space: pre-wrap; font-size: 13px; }
</style>
</head>
<body>
<h1>📋 攜出申請單審核</h1>
<table>
  <tr><th>欄位</th><th>內容</th></tr>
  <tr><td>申請日期</td><td>{{ info.apply_date }}</td></tr>
  <tr><td>預定攜出日期</td><td>{{ info.carry_date }}</td></tr>
  <tr><td>攜出時間</td><td>{{ info.carry_time }}</td></tr>
  <tr><td>議題範圍</td><td>{{ info.topic }}</td></tr>
  <tr><td>填寫人數</td><td>{{ info.count }} 人</td></tr>
  <tr><td>填寫人</td><td>{{ info.filers }}</td></tr>
</table>

<h2>攜出檔案清單</h2>
<div class="files">{{ info.file_list }}</div>

<h2>🔎 預覽 Word 檔</h2>
<p>申請單已存於：<code>{{ info.docx_path }}</code></p>
<p style="color:#666;">請自行開啟確認內容後，再按下方按鈕。</p>

<form method="POST" action="{{ action_url }}">
  <p><label>退回意見（退回時請填寫）：</label></p>
  <textarea name="comment" rows="3" placeholder="例：請補充檔案描述..."></textarea>
  <br>
  <button class="btn approve" name="decision" value="approved">✅ 核准送出</button>
  <button class="btn reject"  name="decision" value="rejected">❌ 退回修改</button>
</form>
</body>
</html>
"""

DONE_PAGE = """
<!DOCTYPE html><html lang="zh-TW"><head><meta charset="UTF-8"><title>完成</title>
<style>body{font-family:sans-serif;text-align:center;padding:80px;}</style></head>
<body><h2>{{ msg }}</h2><p>本頁面可以關閉。</p></body></html>
"""


# ─────────────────────────────────────────────────
# Flask 路由
# ─────────────────────────────────────────────────

@app.route("/review/<token>", methods=["GET"])
def review_page(token: str):
    if token not in _pending:
        abort(404)
    task = _pending[token]
    if task["result"] is not None:
        return render_template_string(DONE_PAGE, msg="此申請已處理完畢。")

    action_url = url_for("review_action", token=token)
    return render_template_string(REVIEW_PAGE, info=task["info"], action_url=action_url)


@app.route("/review/<token>/action", methods=["POST"])
def review_action(token: str):
    if token not in _pending:
        abort(404)
    decision = request.form.get("decision")
    comment  = request.form.get("comment", "")
    _pending[token]["result"]  = decision
    _pending[token]["comment"] = comment
    msg = "✅ 已核准，系統將自動完成後續流程。" if decision == "approved" else "❌ 已退回，請通知填寫人修改。"
    return render_template_string(DONE_PAGE, msg=msg)


@app.route("/health")
def health():
    return "ok"


# ─────────────────────────────────────────────────
# 公開介面
# ─────────────────────────────────────────────────

def create_review_task(docx_path: Path, info: dict) -> str:
    """
    建立一個審核任務，回傳審核 URL（協調人在瀏覽器開啟）。
    info 至少包含：apply_date, carry_date, carry_time, topic, count, filers, file_list, docx_path
    """
    token = uuid.uuid4().hex
    info["docx_path"] = str(docx_path)
    _pending[token] = {"docx_path": str(docx_path), "info": info, "result": None, "comment": ""}
    url = f"http://{config.HITL_HOST}:{config.HITL_PORT}/review/{token}"
    logger.info(f"審核連結：{url}")
    return url


def wait_for_decision(token: str, timeout: int = config.HITL_TIMEOUT_SECS) -> tuple[str, str]:
    """
    阻塞等待協調人做出決定，回傳 (decision, comment)。
    decision = "approved" | "rejected" | "timeout"
    """
    start = time.time()
    while time.time() - start < timeout:
        result = _pending[token]["result"]
        if result is not None:
            comment = _pending[token]["comment"]
            del _pending[token]
            return result, comment
        time.sleep(3)
    del _pending[token]
    return "timeout", ""


def run_server_background():
    """在背景執行緒啟動 Flask（只啟動一次）。"""
    t = threading.Thread(
        target=lambda: app.run(
            host=config.HITL_HOST,
            port=config.HITL_PORT,
            debug=False,
            use_reloader=False,
        ),
        daemon=True,
    )
    t.start()
    time.sleep(1)  # 等待伺服器啟動
    logger.info(f"HITL 審核伺服器已啟動：http://{config.HITL_HOST}:{config.HITL_PORT}")
