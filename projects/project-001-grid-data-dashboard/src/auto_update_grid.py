#!/usr/bin/env python3
"""网格数据看板自动更新：每天7点检查邮箱→下载最新网格数据→重新生成看板→上传GitHub"""
import os, sys, json, imaplib, email, re, subprocess, datetime, shutil
from email.header import decode_header

CONFIG = "/Users/mr.g/Documents/Codex/Workspace/projects/project-005-broadband-distribution-system/src/email_config.json"
DATA_DIR = "/Users/mr.g/Documents/Codex/Workspace/projects/project-001-grid-data-dashboard/data"
PROJ_DIR = "/Users/mr.g/Documents/Codex/Workspace/projects/project-001-grid-data-dashboard"
GEN_SCRIPT = os.path.join(PROJ_DIR, "src", "generate_dashboard.py")
REPO = "253942827-ship-it/grid-data-dashboard"

def ds(s):
    if not s: return ''
    parts = decode_header(s)
    return ''.join([p.decode(c or 'utf-8', errors='replace') if isinstance(p, bytes) else str(p) for p, c in parts])

def main():
    today = datetime.date.today()
    print(f"=== 网格看板自动更新 {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')} ===")
    
    cfg = json.load(open(CONFIG))
    TOKEN = os.environ.get('GITHUB_PAT') or cfg.get('github_token', '')
    if not TOKEN:
        # 兜底：从 git remote URL 提取 PAT（个人访问令牌）
        try:
            remote = subprocess.run(
                ['git', 'remote', 'get-url', 'origin'],
                capture_output=True, text=True, timeout=10,
                cwd=PROJ_DIR
            ).stdout.strip()
            m = re.search(r'https://([^@]+)@github\.com', remote)
            if m:
                TOKEN = m.group(1)
        except Exception:
            pass
    if not TOKEN:
        print("❌ 未找到 GitHub 令牌（GITHUB_PAT / github_token / git remote 中均无）")
        return 1
    try:
        conn = imaplib.IMAP4_SSL(cfg['imap_server'], cfg['imap_port'])
        conn.login(cfg['email'], cfg['password'])
        conn.select('INBOX')
        status, mids = conn.search(None, 'ALL')
        if status != 'OK': print("❌ 搜索失败"); return 1
        
        downloaded = None
        for mid in reversed(mids[0].split()):
            status, data = conn.fetch(mid, '(RFC822)')
            if status != 'OK': continue
            msg = email.message_from_bytes(data[0][1])
            subj = ds(msg['Subject'])
            if '网格单元运营数据' not in subj: continue
            if not msg.is_multipart(): continue
            for part in msg.walk():
                if part.get_content_maintype() == 'multipart': continue
                fn = ds(part.get_filename())
                if fn and '网格单元运营数据' in fn and fn.endswith('.xlsx'):
                    payload = part.get_payload(decode=True)
                    if payload:
                        # 备份旧文件
                        save_path = os.path.join(DATA_DIR, '网格单元运营数据最新.xlsx')
                        # 先存到临时路径
                        tmp_path = os.path.join(DATA_DIR, f'_new_{fn}')
                        with open(tmp_path, 'wb') as f:
                            f.write(payload)
                        # 用临时文件覆盖正式文件（避免部分写入）
                        if os.path.exists(save_path):
                            # 存档旧版本
                            archive_name = f'网格单元运营数据_{datetime.date.today().strftime("%Y%m")}.xlsx'
                            shutil.copy2(save_path, os.path.join(DATA_DIR, archive_name))
                        shutil.move(tmp_path, save_path)
                        # 也覆盖正式文件名
                        shutil.copy2(save_path, os.path.join(DATA_DIR, fn))
                        downloaded = fn
                        print(f'✅ 下载: {fn} ({len(payload)/1024:.0f}KB)')
                    break
            if downloaded: break
        conn.logout()
    except Exception as e:
        print(f"❌ 邮箱访问失败: {e}")
        return 1
    
    if not downloaded:
        print("📭 邮箱中无网格数据文件，跳过更新")
        return 0
    
    # 重新生成看板
    print("🔄 生成网格看板...")
    r = subprocess.run(['python3', GEN_SCRIPT], capture_output=True, text=True, timeout=120)
    print(r.stdout[-100:] if r.returncode == 0 else f"❌ {r.stderr[-100:]}")
    
    # 上传到GitHub Pages
    print("📤 上传到GitHub Pages...")
    import urllib.request, base64, json as j
    headers = {'Authorization': f'Bearer {TOKEN}', 'Accept': 'application/vnd.github+json'}
    html_path = os.path.join(PROJ_DIR, 'docs', 'dashboard.html')
    url = f'https://api.github.com/repos/{REPO}/contents/docs/dashboard.html'
    try:
        req = urllib.request.Request(url, headers=headers, method='GET')
        sha = j.loads(urllib.request.urlopen(req).read()).get('sha')
    except:
        sha = None
    with open(html_path, 'rb') as f:
        content = base64.b64encode(f.read()).decode('utf-8')
    data = j.dumps({'message': f'网格自动更新 {today}', 'content': content, 'sha': sha} if sha else {'message': f'网格自动更新 {today}', 'content': content}).encode()
    req = urllib.request.Request(url, data=data, headers=headers, method='PUT')
    try:
        r2 = urllib.request.urlopen(req)
        print(f'✅ 上传成功 ({r2.status})')
    except Exception as e:
        print(f'❌ 上传失败: {e}')
    return 0

if __name__ == '__main__':
    sys.exit(main())
