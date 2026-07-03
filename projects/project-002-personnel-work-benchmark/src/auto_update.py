#!/usr/bin/env python3
"""人员数据看板自动更新 - 下载邮箱附件→归档旧文件→更新数据→生成看板→上传GitHub"""

import os, sys, json, imaplib, email, re, subprocess, datetime, shutil
from email.header import decode_header

CONFIG = "/Users/mr.g/Documents/Codex/Workspace/projects/project-005-broadband-distribution-system/src/email_config.json"
DATA_DIR = "/Users/mr.g/Documents/Codex/Workspace/projects/project-002-personnel-work-benchmark/data"
PROJ_DIR = "/Users/mr.g/Documents/Codex/Workspace/projects/project-002-personnel-work-benchmark"
WS_DIR = "/Users/mr.g/Documents/Codex/Workspace"

def main():
    today = datetime.date.today()
    print(f"=== 自动更新 {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')} ===")
    cfg = json.load(open(CONFIG))
    
    conn = imaplib.IMAP4_SSL(cfg['imap_server'], cfg['imap_port'])
    conn.login(cfg['email'], cfg['password'])
    conn.select('INBOX')
    
    status, mids = conn.search(None, 'ALL')
    if status != 'OK': print("❌ 搜索失败"); return 1
    msg_ids = mids[0].split()[-50:]
    print(f"  扫描最近 {len(msg_ids)} 封邮件")
    
    downloaded = {}
    for mid in msg_ids:
        status, data = conn.fetch(mid, '(RFC822)')
        if status != 'OK': continue
        msg = email.message_from_bytes(data[0][1])
        subj = decode_str(msg['Subject'])
        if not any(kw in subj for kw in ['清单','高套','竣工','新装','存量','关键一单','杠保','质态','宽带离网']):
            continue
        if not msg.is_multipart(): continue
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart': continue
            fn = decode_str(part.get_filename())
            if not fn or not fn.lower().endswith('.xlsx'): continue
            target = match_filename(fn)
            if target:
                payload = part.get_payload(decode=True)
                if payload:
                    old_size = downloaded.get(target, (None, 0))[1]
                    if len(payload) > old_size:
                        save_path = os.path.join(DATA_DIR, target)
                        # ★ 归档旧文件：保存前把旧文件重命名加上月份标识
                        if os.path.exists(save_path):
                            old_mtime = datetime.datetime.fromtimestamp(os.path.getmtime(save_path))
                            month_tag = old_mtime.strftime('%Y%m')
                            archived_name = target.replace('.xlsx', f'_{month_tag}.xlsx')
                            archived_path = os.path.join(DATA_DIR, archived_name)
                            if not os.path.exists(archived_path):
                                shutil.copy2(save_path, archived_path)
                                print(f"  📦 归档: {target} → {archived_name}")
                        with open(save_path, 'wb') as f:
                            f.write(payload)
                        downloaded[target] = (save_path, len(payload))
                        print(f"  ✅ {target} ({len(payload)/1024:.0f}KB)")
    conn.logout()
    
    if not downloaded:
        print("\n⚠️ 未找到匹配附件")
        return 0
    print(f"\n✅ 下载/更新了 {len(downloaded)} 个文件")
    
    for gen in ['generate_dashboard.py', 'generate_tangxia.py']:
        fp = os.path.join(PROJ_DIR, 'src', gen)
        print(f"🔄 {gen}...")
        r = subprocess.run(['python3', fp], capture_output=True, text=True, timeout=120)
        print(r.stdout[-80:] if r.returncode == 0 else f"❌ {r.stderr[-80:]}")
    
    print("📤 上传到 GitHub Pages...")
    cmds = f"""
    cd {WS_DIR}
    cp {PROJ_DIR}/docs/dashboard.html docs/personnel-dashboard.html
    cp {PROJ_DIR}/docs/tangxia_dashboard.html docs/tangxia_dashboard.html
    git add docs/personnel-dashboard.html docs/tangxia_dashboard.html
    git commit -m "自动更新 $(date +%Y-%m-%d)"
    git -c http.version=HTTP/1.1 push origin main --force
    """
    r = subprocess.run(['/bin/zsh', '-lc', cmds], capture_output=True, text=True, timeout=60)
    print(r.stdout[-80:] if r.returncode == 0 else f"❌ {r.stderr[-80:]}")
    return 0

def decode_str(s):
    if not s: return ''
    parts = decode_header(s)
    return ''.join([p.decode(c or 'utf-8', errors='replace') if isinstance(p, bytes) else str(p) for p, c in parts])

def match_filename(orig):
    name = re.sub(r'[（(]\d+[）)]', '', orig)
    name = re.sub(r'^\d+', '', name).strip()
    for kw, target in {
        '新装高套': '新装高套竣工清单.xlsx',
        '存量高套': '存量高套竣工清单.xlsx',
        '关键一单': '关键一单清单.xlsx',
        '杠保': '杠保清单.xlsx',
        '质态': '质态相关清单.xlsx',
        '宽带离网': '宽带离网清单.xlsx',
        '上月新装': '上月新装高套清单.xlsx',
        '上月存量': '上月存量高套清单.xlsx',
    }.items():
        if kw in name: return target
    return None

if __name__ == '__main__':
    sys.exit(main())
