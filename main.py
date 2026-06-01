"""
carry_out_automation/main.py

攜出申請自動化主程式。

執行方式：
    python main.py              # 立即執行一次
    python main.py --dry-run    # 只讀不寫（測試用）

排程（每週一早上 9:00）：
    使用 cron：  0 9 * * 1  cd /path/to/carry_out_automation && python main.py
    或呼叫 schedule_task.py 設定 macOS launchd / Windows Task Scheduler
"""

from __future__ import annotations

import argparse
import logging
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path

import config
import email_sender
import form_filler
import hitl_server
import sheets_reader

# ─────────────────────────────────────────────────
# 日誌設定
# ─────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(Path(__file__).parent / "automation.log", encoding="utf-8"),
    ],
)
logger = logging.getLogger("main")


# ─────────────────────────────────────────────────
# 主流程
# ─────────────────────────────────────────────────

def run(dry_run: bool = False):
    logger.info("=" * 60)
    logger.info(f"攜出自動化流程啟動  {'[DRY RUN]' if dry_run else ''}")
    logger.info("=" * 60)

    # ── 步驟 1：啟動 HITL 審核伺服器 ──────────────────
    if not dry_run:
        hitl_server.run_server_background()

    # ── 步驟 2：從 Google Drive xlsx 讀取待處理申請 ───
    logger.info("讀取 Google Drive 申請表…")
    try:
        pending, df, raw_bytes = sheets_reader.fetch_pending_requests()
    except Exception as e:
        logger.error(f"讀取 Google Drive 失敗：{e}")
        return

    if not pending:
        logger.info("目前無待處理申請，流程結束。")
        return

    logger.info(f"共找到 {len(pending)} 筆待處理申請")

    # ── 步驟 3：依（攜出日期 + 議題）分組 ─────────────
    groups = sheets_reader.group_by_date_and_topic(pending)
    logger.info(f"分為 {len(groups)} 個攜出批次")

    approved_rows: list[int] = []

    for (carry_date, topic), requests in groups.items():
        logger.info(f"\n─── 批次：{carry_date} / {topic}（{len(requests)} 人）───")

        carry_time = requests[0].預定攜出時間 or "10:00"

        # ── 步驟 4：填寫攜出申請單 .docx ──────────────
        # 將攜出日期正規化為 YYYYMMDD（去除時間戳記）
        _dt = form_filler._parse_date(carry_date)
        date_tag = _dt.strftime("%Y%m%d") if _dt else re.sub(r"[^0-9]", "", carry_date)[:8]
        ym_tag   = _dt.strftime("%Y%m")   if _dt else date_tag[:6]

        fname     = f"MOU攜出{topic}資料申請{date_tag}.docx"
        output_dir = config.CARRYOUT_FORM_DIR / ym_tag / date_tag
        docx_path = output_dir / fname

        if dry_run:
            logger.info(f"[DRY RUN] 將產生：{docx_path}")
            continue

        try:
            form_filler.fill_form(
                requests=requests,
                carry_date=carry_date,
                carry_time=carry_time,
                topic=topic,
                output_path=docx_path,
            )
        except Exception as e:
            logger.error(f"表單填寫失敗：{e}")
            continue

        # ── 步驟 5：HITL 審核 ────────────────────────
        filers     = "、".join(r.填寫人 for r in requests)
        file_lines = "\n".join(
            f"{r.填寫人}：{r.檔案名稱}（{r.格式內容說明[:30]}…）"
            for r in requests
        )
        review_info = {
            "apply_date": datetime.today().strftime("%Y/%m/%d"),
            "carry_date": carry_date,
            "carry_time": carry_time,
            "topic":      topic,
            "count":      len(requests),
            "filers":     filers,
            "file_list":  file_lines,
        }

        review_url = hitl_server.create_review_task(docx_path, review_info)

        mof_coord_preview = next(
            (v for k, v in config.MOU_COORDINATORS.items() if k in topic), None
        )

        # 發送審核 email（含連結、Word 附件、財資中心信件預覽）
        try:
            email_sender.send_review_request(
                review_url, requests, docx_path, mof_coord=mof_coord_preview
            )
            logger.info(f"審核通知 Email 已發送，等待協調人決定…")
            logger.info(f"（也可直接開啟：{review_url}）")
        except Exception as e:
            logger.warning(f"Email 發送失敗：{e}  →  請手動開啟審核連結：{review_url}")

        # 等待決定
        token = review_url.split("/")[-1]
        decision, comment = hitl_server.wait_for_decision(token)
        logger.info(f"審核結果：{decision}  意見：{comment or '（無）'}")

        # ── 步驟 6：依決定處理 ────────────────────────
        if decision == "approved":
            # ── Step 2：docx → PDF ────────────────────
            try:
                pdf_path = form_filler.docx_to_pdf(docx_path)
            except Exception as e:
                logger.error(f"PDF 轉換失敗：{e}，改以 docx 附件寄出")
                pdf_path = docx_path

            # 建立攜出資料夾，複製 PDF（或 docx）
            carryout_folder = form_filler.create_carryout_folder(carry_date, topic)
            shutil.copy2(pdf_path, carryout_folder / pdf_path.name)


            # 寄送正式攜出申請通知給財資中心承辦人（附 PDF）
            mof_coord = next(
                (v for k, v in config.MOU_COORDINATORS.items() if k in topic),
                None,
            )
            if mof_coord:
                try:
                    email_sender.send_mof_request(requests, pdf_path, mof_coord)
                    logger.info(f"已寄送攜出通知給財資中心：{mof_coord['name']}（{mof_coord['email']}）")
                except Exception as e:
                    logger.warning(f"財資中心通知信寄送失敗：{e}")
            else:
                logger.warning(f"找不到子題「{topic}」對應的財資中心承辦人")

            # 記錄待更新的列號
            approved_rows.extend(r.row_index for r in requests)
            logger.info(f"✅ 批次核准完成：{carry_date} / {topic}")

        elif decision == "rejected":
            try:
                email_sender.send_rejection_notice(requests, comment)
            except Exception as e:
                logger.warning(f"退回通知 Email 失敗：{e}")
            logger.info(f"❌ 批次退回：{carry_date} / {topic}")

        else:  # timeout
            logger.warning(f"⏰ 審核逾時（{config.HITL_TIMEOUT_SECS}s），請手動處理")

    # ── 步驟 7：更新 Google Drive xlsx（送出申請 = TRUE）
    if approved_rows and not dry_run:
        try:
            sheets_reader.mark_as_submitted(approved_rows, df)
            logger.info(f"已將 {len(approved_rows)} 筆標記為已送出")
        except Exception as e:
            logger.error(f"更新 xlsx 失敗：{e}")

    logger.info("\n流程結束。")


# ─────────────────────────────────────────────────
# 入口
# ─────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="攜出申請自動化")
    parser.add_argument("--dry-run", action="store_true", help="只讀不寫（測試模式）")
    args = parser.parse_args()
    run(dry_run=args.dry_run)
