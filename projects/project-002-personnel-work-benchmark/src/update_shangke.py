#!/usr/bin/env python3
"""从邮箱下载全渠道做商客战报，按人员维度提取沙太商客发展数据。"""
import os
import re
import sys
import json
import imaplib
import email
import shutil
import openpyxl
from email.header import decode_header
from datetime import datetime

CONFIG = "/Users/mr.g/Documents/Codex/Workspace/projects/project-005-broadband-distribution-system/src/email_config.json"
DATA_DIR = "/Users/mr.g/Documents/Codex/Workspace/projects/project-002-personnel-work-benchmark/data"
REPORT_FILE = "全渠道做商客战报.xlsx"
OUT_JSON = "商客发展情况.json"
SHEET_NAME = "2_人员维度"
TARGET_YINGFU = "天河沙太城中村营销服务中心"


def ds(s):
    if not s:
        return ""
    parts = decode_header(s)
    return "".join(
        p.decode(c or "utf-8", errors="replace") if isinstance(p, bytes) else str(p)
        for p, c in parts
    )


def parse_num(v):
    if v is None:
        return 0
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip().replace(",", "")
    if not s:
        return 0
    try:
        return float(s)
    except ValueError:
        return 0


def report_date_from_filename(fn):
    m = re.search(r"(\d{8})", fn)
    if m:
        d = m.group(1)
        return f"{d[:4]}-{d[4:6]}-{d[6:]}"
    m = re.search(r"[【（](\d{4})[】）]", fn)
    if m:
        mmdd = m.group(1)
        return f"{datetime.now().year}-{mmdd[:2]}-{mmdd[2:]}"
    m = re.search(r"(\d{4})", fn)
    if m:
        mmdd = m.group(1)
        return f"{datetime.now().year}-{mmdd[:2]}-{mmdd[2:]}"
    return ""


def download_latest_report():
    """从邮箱下载最新全渠道做商客战报附件，保存为 全渠道做商客战报.xlsx + 月度归档。"""
    cfg = json.load(open(CONFIG))
    conn = imaplib.IMAP4_SSL(cfg["imap_server"], cfg["imap_port"])
    conn.login(cfg["email"], cfg["password"])
    conn.select("INBOX")
    status, mids = conn.search(None, "ALL")
    if status != "OK":
        conn.logout()
        return None, None
    all_ids = mids[0].split()
    found = None
    for mid in reversed(all_ids[-150:]):
        status, data = conn.fetch(mid, "(RFC822)")
        if status != "OK":
            continue
        msg = email.message_from_bytes(data[0][1])
        subj = ds(msg.get("Subject", ""))
        if not msg.is_multipart():
            continue
        for part in msg.walk():
            fn = ds(part.get_filename())
            if not fn:
                continue
            if "全渠道做商客战报" not in fn or not fn.lower().endswith(".xlsx"):
                continue
            if "全渠道做商客战报" in subj:
                payload = part.get_payload(decode=True)
                if payload:
                    found = (fn, payload)
                    break
        if found:
            break
    conn.logout()
    if not found:
        return None, None
    fn, payload = found
    os.makedirs(DATA_DIR, exist_ok=True)
    path = os.path.join(DATA_DIR, REPORT_FILE)
    with open(path, "wb") as f:
        f.write(payload)
    month = datetime.now().strftime("%Y%m")
    archive = os.path.join(DATA_DIR, f"全渠道做商客战报_{month}.xlsx")
    if not os.path.exists(archive):
        shutil.copy2(path, archive)
    return fn, path


def extract(report_path, report_file=None):
    """读取人员维度表，按现有 personnel.json 编码匹配，输出商客发展情况 JSON。"""
    pd_json = json.load(open(os.path.join(DATA_DIR, "personnel.json"), encoding="utf-8"))
    by_code = {str(p.get("code", "")).strip(): p["name"] for p in pd_json["personnel"]}
    by_name = {p["name"]: str(p.get("code", "")).strip() for p in pd_json["personnel"]}

    wb = openpyxl.load_workbook(report_path, data_only=True, read_only=True)
    ws = wb[SHEET_NAME]
    rows = ws.iter_rows(values_only=True)
    header = [str(v or "").strip() for v in next(rows)]
    col = {name: i for i, name in enumerate(header) if name}
    idx_yingfu = col.get("营服")
    idx_name = col.get("成员")
    idx_code = col.get("揽装人编码")
    idx_shangqi = col.get("增存商企")
    idx_weixiao = col.get("小微ICT")
    idx_dev299 = col.get("业务量汇总")
    if any(i is None for i in (idx_yingfu, idx_name, idx_code, idx_shangqi, idx_weixiao, idx_dev299)):
        wb.close()
        raise RuntimeError(f"人员维度表缺少必要列: {header}")

    people = {}
    skipped = []
    for row in rows:
        if row[idx_yingfu] != TARGET_YINGFU:
            continue
        name = str(row[idx_name] or "").strip()
        code = str(row[idx_code] or "").strip()
        if not name and not code:
            continue
        if code in by_code:
            key = code
        elif name in by_name:
            key = by_name[name]
        else:
            skipped.append(name or code)
            continue
        people[key] = {
            "name": by_code.get(key, name),
            "shangqi": int(parse_num(row[idx_shangqi])),
            "weixiao": round(parse_num(row[idx_weixiao]), 2),
            "dev299": int(parse_num(row[idx_dev299])),
        }
    wb.close()

    report_file = report_file or os.path.basename(report_path)
    out = {
        "report_file": report_file,
        "report_date": report_date_from_filename(report_file),
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "people": people,
    }
    out_path = os.path.join(DATA_DIR, OUT_JSON)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    return out, skipped


def main(download=True):
    fn = None
    if download:
        fn, path = download_latest_report()
        if not fn or not path:
            path = os.path.join(DATA_DIR, REPORT_FILE)
            if not os.path.exists(path):
                print("  ⚠️ 未找到全渠道做商客战报邮件，且本地无历史文件")
                return 1
            print(f"  ⚠️ 未找到新邮件，使用本地文件: {REPORT_FILE}")
    else:
        path = os.path.join(DATA_DIR, REPORT_FILE)
    try:
        out, skipped = extract(path, report_file=fn)
    except Exception as e:
        print(f"  ❌ 商客战报解析失败: {e}")
        return 1
    print(f"  ✅ 商客战报: {out['report_date']} 匹配 {len(out['people'])} 人，跳过 {len(skipped)} 人")
    if skipped:
        print(f"     未匹配(忽略): {', '.join(skipped)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
