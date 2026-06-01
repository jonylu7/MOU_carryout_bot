"""
carry_out_automation/email_sender.py

發送攜出通知 Email（SMTP / Gmail）。
依賴：Python 標準庫 smtplib（無需額外安裝）
"""

from __future__ import annotations

import logging
import smtplib
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Optional

import config
from sheets_reader import CarryOutRequest

logger = logging.getLogger(__name__)


def _build_message(
    subject: str,
    body_html: str,
    to_addresses: list[str],
    attachment_path: Optional[Path] = None,
) -> MIMEMultipart:
    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["From"]    = config.EMAIL_SENDER
    msg["To"]      = ", ".join(to_addresses)

    msg.attach(MIMEText(body_html, "html", "utf-8"))

    if attachment_path and attachment_path.exists():
        with open(attachment_path, "rb") as f:
            part = MIMEApplication(f.read(), Name=attachment_path.name)
        part["Content-Disposition"] = f'attachment; filename="{attachment_path.name}"'
        msg.attach(part)

    return msg


def _send(msg: MIMEMultipart, to_addresses: list[str]) -> None:
    with smtplib.SMTP(config.EMAIL_SMTP_HOST, config.EMAIL_SMTP_PORT) as server:
        server.ehlo()
        server.starttls()
        server.login(config.EMAIL_SENDER, config.EMAIL_PASSWORD)
        server.sendmail(config.EMAIL_SENDER, to_addresses, msg.as_string())
    logger.info(f"Email 已發送至：{to_addresses}")


# ─────────────────────────────────────────────────
# 審核通知（給協調人，含審核連結）
# ─────────────────────────────────────────────────

def send_review_request(
    review_url: str,
    requests: list[CarryOutRequest],
    docx_path: Path,
    mof_coord: Optional[dict] = None,
) -> None:
    filers = "、".join(set(r.填寫人 for r in requests))
    carry_date = requests[0].預定攜出日期 if requests else "—"
    topic      = requests[0].議題範圍    if requests else "—"

    file_list_html = "<ul>" + "".join(
        f"<li>{r.填寫人}：{r.檔案名稱}</li>" for r in requests
    ) + "</ul>"

    # 財資中心信件預覽
    if mof_coord:
        preview_text = _build_mof_body(requests, mof_coord)
        to_name  = mof_coord["name"]
        to_email = mof_coord["email"]
        cc_str   = "、".join(config.EMAIL_MOF_CC) if config.EMAIL_MOF_CC else "（無）"
        mof_preview_html = f"""
<hr>
<h3>📧 核准後將自動寄出以下信件</h3>
<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse; font-size:13px;">
  <tr><td><b>To</b></td><td>{to_name}　{to_email}</td></tr>
  <tr><td><b>CC</b></td><td>{cc_str}</td></tr>
</table>
<pre style="background:#f4f4f4; padding:12px; border-radius:4px; font-size:13px; white-space:pre-wrap;">{preview_text}</pre>
"""
    else:
        mof_preview_html = ""

    body = f"""
<html><body style="font-family:sans-serif; color:#333;">
<h2>【攜出申請單已產生，請審閱後送出】</h2>
<table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;">
  <tr><td><b>預定攜出日期</b></td><td>{carry_date}</td></tr>
  <tr><td><b>議題範圍</b></td><td>{topic}</td></tr>
  <tr><td><b>填寫人</b></td><td>{filers}</td></tr>
</table>
<h3>攜出檔案：</h3>
{file_list_html}
<hr>
<p><b>📄 Word 檔位置（可直接開啟編輯）：</b><br>
<code>{docx_path}</code></p>
<p style="color:#c0392b;">⚠️ 請先開啟 Word 檔確認範例截圖是否完整，必要時手動補充，修改後儲存。</p>
{mof_preview_html}
<hr>
<p>確認無誤後，點擊下方按鈕將自動轉為 PDF 並寄送給財資中心及填寫人：</p>
<p>
  <a href="{review_url}" style="
    background:#1a5276; color:white; padding:12px 24px;
    text-decoration:none; border-radius:6px; font-size:15px;
  ">✅ 確認送出（轉 PDF 並寄信）</a>
</p>
<p style="color:#888; font-size:12px;">此連結有效期限 {config.HITL_TIMEOUT_SECS // 3600} 小時。</p>
</body></html>
"""
    subject = f"【待審核】攜出申請 {carry_date}（{topic}）{filers}"
    msg = _build_message(
        subject=subject,
        body_html=body,
        to_addresses=[config.EMAIL_COORDINATOR],
        attachment_path=docx_path,
    )
    _send(msg, [config.EMAIL_COORDINATOR])


# ─────────────────────────────────────────────────
# 核准通知（給所有填寫人）
# ─────────────────────────────────────────────────

def send_approval_notice(
    requests: list[CarryOutRequest],
    docx_path: Path,
    carryout_folder: Path,
) -> None:
    carry_date = requests[0].預定攜出日期 if requests else "—"
    carry_time = requests[0].預定攜出時間 if requests else "—"
    topic      = requests[0].議題範圍    if requests else "—"

    filer_emails = list({r.填寫人信箱 for r in requests if r.填寫人信箱})
    all_to = filer_emails + [config.EMAIL_COORDINATOR]

    items_html = "<ul>" + "".join(
        f"<li><b>{r.填寫人}</b>：{r.檔案名稱}</li>" for r in requests
    ) + "</ul>"

    body = f"""
<html><body style="font-family:sans-serif; color:#333;">
<h2>【攜出申請已核准】</h2>
<p>您的攜出申請已通過審核，請依下列時程攜帶。</p>
<table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;">
  <tr><td><b>預定攜出日期</b></td><td>{carry_date}</td></tr>
  <tr><td><b>攜出時間</b></td><td>{carry_time}</td></tr>
  <tr><td><b>議題範圍</b></td><td>{topic}</td></tr>
</table>
<h3>核准攜出檔案：</h3>
{items_html}
<p>攜出申請單（Word）已附件，請攜帶列印版本至財資中心。</p>
<p>攜出資料請放至：<code>{carryout_folder}</code></p>
<hr>
<p style="color:#888; font-size:12px;">此為系統自動發送，如有疑問請聯繫 {config.EMAIL_COORDINATOR}</p>
</body></html>
"""
    subject = f"【已核准】攜出申請 {carry_date}（{topic}）"
    msg = _build_message(
        subject=subject,
        body_html=body,
        to_addresses=all_to,
        attachment_path=docx_path,
    )
    _send(msg, all_to)


# ─────────────────────────────────────────────────
# 財資中心攜出申請通知（核准後寄給承辦人）
# ─────────────────────────────────────────────────

def _build_mof_body(requests: list[CarryOutRequest], mof_coord: dict) -> str:
    """產生寄給財資中心承辦人的純文字信件內容。"""
    from datetime import datetime as _dt

    carry_date = requests[0].預定攜出日期 if requests else ""
    carry_time = requests[0].預定攜出時間 if requests else ""
    topic      = requests[0].議題範圍    if requests else ""

    date_obj = None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y/%m/%d", "%Y-%m-%d"):
        try:
            date_obj = _dt.strptime(str(carry_date).strip(), fmt)
            break
        except ValueError:
            continue
    weekday_map = ["一", "二", "三", "四", "五", "六", "日"]
    date_str = date_obj.strftime("%m/%d") if date_obj else str(carry_date)
    weekday  = weekday_map[date_obj.weekday()] if date_obj else ""

    time_obj = None
    for tfmt in ("%H:%M:%S", "%H:%M"):
        try:
            time_obj = _dt.strptime(str(carry_time).strip(), tfmt)
            break
        except ValueError:
            continue
    if time_obj:
        ampm     = "上午" if time_obj.hour < 12 else "下午"
        hour     = time_obj.hour if time_obj.hour <= 12 else time_obj.hour - 12
        time_str = f"{ampm}{hour}點"
    else:
        time_str = str(carry_time)

    topic_desc_map = {
        "子題一":      "調降營所稅及促產條例落日之政策效果",
        "子題一子議題": "稅制改革引起之地下經濟",
        "子題二":      "台灣租稅與年金改革及世代間不平等",
        "稅收估計":    "使用數據分析進行各稅別稅收估計",
    }
    research_topic = next((v for k, v in topic_desc_map.items() if k in topic), topic)

    file_lines = [
        f.strip()
        for r in requests
        for f in str(r.檔案名稱).split("\n")
        if f.strip()
    ]
    file_list_str = "\n".join(file_lines)

    return (
        f"{mof_coord['name']}你好：\n\n"
        f"我是盧威任，政大團隊本週預計申請{research_topic}研究案的資料攜出，"
        f"申請單如附件，請您查收。\n\n"
        f"本次申請攜出檔案（共{len(file_lines)}件）如下：\n"
        f"{file_list_str}\n\n"
        f"預計{date_str}（{weekday}）{time_str}辦理攜出，請問是否方便呢？\n"
        f"非常感謝您的幫忙！\n\n"
        f"盧威任 敬上"
    )


def send_mof_request(
    requests: list[CarryOutRequest],
    docx_path: Path,
    mof_coord: dict,   # {"name": ..., "email": ...}
) -> None:
    from datetime import datetime as _dt

    carry_date = requests[0].預定攜出日期 if requests else ""
    topic      = requests[0].議題範圍    if requests else ""

    date_obj = None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y/%m/%d", "%Y-%m-%d"):
        try:
            date_obj = _dt.strptime(str(carry_date).strip(), fmt)
            break
        except ValueError:
            continue
    date_str = date_obj.strftime("%m/%d") if date_obj else str(carry_date)

    body_text = _build_mof_body(requests, mof_coord)
    subject   = f"財政部財資中心資料攜出申請（{date_str}）"
    to_addr   = config.EMAIL_MOF_TO
    cc_addrs  = config.EMAIL_MOF_CC

    # 轉成 HTML，在「盧威任 敬上」前插入簽名圖片
    has_sig = config.SIGNATURE_IMAGE_PATH.exists()
    sig_tag = '<img src="cid:sig" style="height:45px; vertical-align:middle;"><br>' if has_sig else ""
    html_body = (
        "<html><body style='font-family:sans-serif; line-height:1.8; color:#333;'>"
        + body_text.replace(
            "盧威任 敬上",
            f"{sig_tag}盧威任 敬上",
        ).replace("\n", "<br>")
        + "</body></html>"
    )

    msg_outer   = MIMEMultipart("mixed")
    msg_outer["Subject"] = subject
    msg_outer["From"]    = config.EMAIL_SENDER
    msg_outer["To"]      = to_addr
    if cc_addrs:
        msg_outer["Cc"] = ", ".join(cc_addrs)

    msg_related = MIMEMultipart("related")
    msg_related.attach(MIMEText(html_body, "html", "utf-8"))
    if has_sig:
        with open(config.SIGNATURE_IMAGE_PATH, "rb") as f:
            sig_img = MIMEImage(f.read())
        sig_img.add_header("Content-ID", "<sig>")
        sig_img.add_header("Content-Disposition", "inline")
        msg_related.attach(sig_img)
    msg_outer.attach(msg_related)

    if docx_path and docx_path.exists():
        with open(docx_path, "rb") as f:
            part = MIMEApplication(f.read(), Name=docx_path.name)
        part["Content-Disposition"] = f'attachment; filename="{docx_path.name}"'
        msg_outer.attach(part)

    _send(msg_outer, [to_addr] + cc_addrs)


# ─────────────────────────────────────────────────
# 退回通知
# ─────────────────────────────────────────────────

def send_rejection_notice(
    requests: list[CarryOutRequest],
    comment: str,
) -> None:
    carry_date = requests[0].預定攜出日期 if requests else "—"
    topic      = requests[0].議題範圍    if requests else "—"
    filer_emails = list({r.填寫人信箱 for r in requests if r.填寫人信箱})

    body = f"""
<html><body style="font-family:sans-serif; color:#333;">
<h2>【攜出申請需修改】</h2>
<p>您的攜出申請（{carry_date}，{topic}）已被退回，請依以下意見修改後重新送出：</p>
<blockquote style="background:#fef9e7; padding:12px; border-left:4px solid #f39c12;">
  {comment or "（無附加說明）"}
</blockquote>
<p>修改後請重新填寫 Google Drive 申請表，並通知協調人。</p>
<hr>
<p style="color:#888; font-size:12px;">如有疑問請聯繫 {config.EMAIL_COORDINATOR}</p>
</body></html>
"""
    subject = f"【請修改】攜出申請 {carry_date}（{topic}）"
    all_to = filer_emails + [config.EMAIL_COORDINATOR]
    msg = _build_message(subject=subject, body_html=body, to_addresses=all_to)
    _send(msg, all_to)
