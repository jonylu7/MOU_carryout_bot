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
  body { font-family: "Noto Sans TC", sans-serif; max-width: 820px; margin: 40px auto; padding: 0 20px; color: #333; }
  h1 { color: #1a5276; border-bottom: 2px solid #1a5276; padding-bottom: 8px; }
  h2 { color: #1a5276; margin-top: 28px; }
  .label { font-size: 13px; color: #888; margin-bottom: 4px; }
  .email-preview {
    background: #f7f9fb; border: 1px solid #d0d7de; border-radius: 6px;
    padding: 18px 22px; white-space: pre-wrap; font-size: 14px; line-height: 1.8;
  }
  .meta { font-size: 13px; color: #555; margin-bottom: 10px; }
  .docx-path { background: #f0f3f4; padding: 8px 12px; border-radius: 4px; font-family: monospace; font-size: 13px; }
  .btn { padding: 12px 28px; border: none; border-radius: 6px; font-size: 16px; cursor: pointer; margin: 8px; }
  .approve { background: #27ae60; color: white; }
  .reject  { background: #e74c3c; color: white; }
  textarea { width: 100%; padding: 10px; border-radius: 6px; border: 1px solid #ccc; font-size: 14px; }
  hr { border: none; border-top: 1px solid #e0e0e0; margin: 24px 0; }
</style>
</head>
<body>
<h1>📋 攜出申請單審核</h1>

<h2>📧 核准後將寄出以下信件</h2>
{% if info.mof_to_name %}
<div class="meta">
  <b>To：</b>{{ info.mof_to_name }}　{{ info.mof_to_email }}
  {% if info.mof_cc %}<br><b>CC：</b>{{ info.mof_cc }}{% endif %}
</div>
{% endif %}
<div class="email-preview">{{ info.mof_body }}</div>

<hr>
<h2>📄 Word 申請單</h2>
<p class="label">請先開啟確認範例截圖是否完整，修改後存檔，再按下方按鈕。</p>
<div class="docx-path">{{ info.docx_path }}</div>

<hr>
<form method="POST" action="{{ action_url }}">
  <p><label>退回意見（退回時請填寫）：</label></p>
  <textarea name="comment" rows="3" placeholder="例：請補充檔案描述..."></textarea>
  <br>
  <button class="btn approve" name="decision" value="approved">✅ 確認送出（轉 PDF 並寄信）</button>
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
