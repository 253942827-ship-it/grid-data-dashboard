#!/usr/bin/env python3
"""人员数据看板自动更新"""
import os, sys, json, imaplib, email, re, subprocess, datetime, shutil, openpyxl
from email.header import decode_header
from datetime import datetime as dt, date

CONFIG = "/Users/mr.g/Documents/Codex/Workspace/projects/project-005-broadband-distribution-system/src/email_config.json"
DATA_DIR = "/Users/mr.g/Documents/Codex/Workspace/projects/project-002-personnel-work-benchmark/data"
PROJ_DIR = "/Users/mr.g/Documents/Codex/Workspace/projects/project-002-personnel-work-benchmark"
WS_DIR = "/Users/mr.g/Documents/Codex/Workspace"
BACKUP_DIR = os.path.join(PROJ_DIR, "data_archive")

def ds(s):
    if not s: return ''
    parts = decode_header(s)
    return ''.join([p.decode(c or 'utf-8', errors='replace') if isinstance(p, bytes) else str(p) for p, c in parts])

def match_target(orig):
    name = re.sub(r'[（(]\d+[）)]', '', orig)
    name = re.sub(r'^\d+', '', name).strip()
    for kw, t in {'新装高套':'新装高套竣工清单.xlsx','存量高套':'存量高套竣工清单.xlsx',
        '关键一单':'关键一单清单.xlsx','杠保':'杠保清单.xlsx','质态':'质态相关清单.xlsx',
        '宽带离网':'宽带离网清单.xlsx'}.items():
        if kw in name: return t
    return None

def get_data_month(fp):
    """读取文件的最新数据月份"""
    try:
        wb = openpyxl.load_workbook(fp, data_only=True)
        ws = wb.active
        # 找日期列
        for col in range(1, min(ws.max_column+1, 20)):
            h = str(ws.cell(1, col).value or '').strip()
            if '日期' in h:
                for r in range(2, min(ws.max_row+1, 10)):
                    dv = ws.cell(r, col).value
                    if isinstance(dv, (dt, date)):
                        wb.close()
                        return dv.month if isinstance(dv, date) else dv.date().month
                    elif isinstance(dv, int) and dv > 20260000:
                        wb.close()
                        return int(str(dv)[4:6])
        wb.close()
    except: pass
    return None

def main():
    today = dt.now()
    cur_month = today.month
    print(f"=== 自动更新 {today.strftime('%Y-%m-%d %H:%M')} ===")
    
    # 先保存当前数据为备份（防止下载到旧数据）
    tmp_backup = {}
    for f in ['新装高套竣工清单.xlsx','存量高套竣工清单.xlsx','关键一单清单.xlsx',
              '杠保清单.xlsx','质态相关清单.xlsx','宽带离网清单.xlsx']:
        fp = os.path.join(DATA_DIR, f)
        if os.path.exists(fp):
            tmp_backup[f] = open(fp, 'rb').read()
    
    # 连接邮箱下载
    cfg = json.load(open(CONFIG))
    try:
        conn = imaplib.IMAP4_SSL(cfg['imap_server'], cfg['imap_port'])
        conn.login(cfg['email'], cfg['password'])
        conn.select('INBOX')
        status, mids = conn.search(None, 'ALL')
        if status != 'OK': print("❌ 搜索失败"); return 1
        all_ids = mids[0].split()
        downloaded = set()
        for mid in reversed(all_ids):
            if len(downloaded) >= 6: break
            status, data = conn.fetch(mid, '(RFC822)')
            if status != 'OK': continue
            msg = email.message_from_bytes(data[0][1])
            subj = ds(msg['Subject'])
            if not any(kw in subj for kw in ['清单','高套','竣工']): continue
            if not msg.is_multipart(): continue
            for part in msg.walk():
                if part.get_content_maintype() == 'multipart': continue
                fn = ds(part.get_filename())
                if not fn: continue
                target = match_target(fn)
                if target and target not in downloaded:
                    payload = part.get_payload(decode=True)
                    if payload:
                        with open(os.path.join(DATA_DIR, target), 'wb') as f:
                            f.write(payload)
                        downloaded.add(target)
                        print(f'  ✅ {target} ({len(payload)/1024:.0f}KB)')
        conn.logout()
    except Exception as e:
        print(f"❌ 邮箱下载失败: {e}")
        # 恢复备份
        for fname, content in tmp_backup.items():
            with open(os.path.join(DATA_DIR, fname), 'wb') as f:
                f.write(content)
        print("  已恢复原始数据")
    
    # ★ 月份检查：如果下载的文件是上个月的数据，恢复备份
    test_file = os.path.join(DATA_DIR, '新装高套竣工清单.xlsx')
    data_month = get_data_month(test_file)
    if data_month and data_month < cur_month:
        print(f"⚠️ 检测到下载了 {data_month}月数据（当前月份={cur_month}月），恢复为备份数据")
        for fname, content in tmp_backup.items():
            with open(os.path.join(DATA_DIR, fname), 'wb') as f:
                f.write(content)
        # 如果备份也是旧数据，尝试从存档恢复
        if get_data_month(test_file) and get_data_month(test_file) < cur_month:
            archive_dir = os.path.join(BACKUP_DIR, f"2026-{cur_month:02d}")
            if os.path.exists(archive_dir):
                for fn in os.listdir(archive_dir):
                    if fn.endswith('.xlsx'):
                        shutil.copy2(os.path.join(archive_dir, fn), os.path.join(DATA_DIR, fn))
                print(f"✅ 从存档 {archive_dir} 恢复数据")
    
    # 生成看板
    for gen in ['generate_dashboard.py', 'generate_tangxia.py']:
        fp = os.path.join(PROJ_DIR, 'src', gen)
        print(f"🔄 {gen}...")
        r = subprocess.run(['python3', fp], capture_output=True, text=True, timeout=120)
        print(r.stdout[-80:] if r.returncode == 0 else f"❌ {r.stderr[-80:]}")
    
    # 上传GitHub
    print("📤 上传到 GitHub Pages...")
    cmds = f"""
cd {WS_DIR}
cp {PROJ_DIR}/docs/dashboard.html docs/personnel-dashboard.html
cp {PROJ_DIR}/docs/tangxia_dashboard.html docs/tangxia_dashboard.html
git add docs/personnel-dashboard.html docs/tangxia_dashboard.html
git commit -m "自动更新 $(date +%Y-%m-%d)" 2>/dev/null || true
git -c http.version=HTTP/1.1 push origin main --force 2>&1 | tail -3
"""
    r = subprocess.run(['/bin/zsh', '-lc', cmds], capture_output=True, text=True, timeout=60)
    print(r.stdout[-100:] if r.returncode == 0 else f"❌ {r.stderr[-100:]}")
    return 0

if __name__ == '__main__':
    sys.exit(main())
