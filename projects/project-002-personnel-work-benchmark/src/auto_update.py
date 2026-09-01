#!/usr/bin/env python3
"""人员数据看板自动更新"""
import os, sys, json, imaplib, email, re, subprocess, datetime, shutil, openpyxl, base64, urllib.error, urllib.request
from email.header import decode_header
from datetime import datetime as dt, date
import update_shangke

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

def github_api_request(method, path, token, payload=None):
    req = urllib.request.Request(
        'https://api.github.com' + path,
        method=method,
        data=json.dumps(payload).encode('utf-8') if payload is not None else None,
        headers={
            'Authorization': 'Bearer ' + token,
            'Accept': 'application/vnd.github+json',
            'User-Agent': 'personnel-dashboard-automation',
            'Content-Type': 'application/json',
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode('utf-8')), None
    except urllib.error.HTTPError as e:
        return None, f"HTTP {e.code}: {e.read().decode('utf-8', 'replace')}"
    except Exception as e:
        return None, str(e)

def publish_via_github_api(files):
    remote_url = subprocess.check_output(
        ['git', '-C', WS_DIR, 'config', '--get', 'remote.origin.url'],
        text=True,
    ).strip()
    token = remote_url.split('https://', 1)[1].split('@', 1)[0]
    repo = remote_url.split('github.com/', 1)[1].removesuffix('.git')
    repo_path = '/repos/' + repo

    head, err = github_api_request('GET', repo_path + '/commits/main', token)
    if err:
        raise RuntimeError(f"获取远端 main 失败: {err}")

    entries = []
    for src, dst in files:
        with open(os.path.join(WS_DIR, src), 'rb') as f:
            content = f.read()
        blob, err = github_api_request(
            'POST', repo_path + '/git/blobs', token,
            {'content': base64.b64encode(content).decode('ascii'), 'encoding': 'base64'},
        )
        if err:
            raise RuntimeError(f"上传 {dst} blob 失败: {err}")
        entries.append({'path': dst, 'mode': '100644', 'type': 'blob', 'sha': blob['sha']})

    tree, err = github_api_request(
        'POST', repo_path + '/git/trees', token,
        {'base_tree': head['commit']['tree']['sha'], 'tree': entries},
    )
    if err:
        raise RuntimeError(f"创建 tree 失败: {err}")

    commit, err = github_api_request(
        'POST', repo_path + '/git/commits', token,
        {
            'message': f"自动更新 {dt.now().strftime('%Y-%m-%d')}",
            'tree': tree['sha'],
            'parents': [head['sha']],
        },
    )
    if err:
        raise RuntimeError(f"创建 commit 失败: {err}")

    ref, err = github_api_request(
        'PATCH', repo_path + '/git/refs/heads/main', token,
        {'sha': commit['sha'], 'force': True},
    )
    if err:
        raise RuntimeError(f"更新 main 失败: {err}")
    print(f"  ✅ GitHub API 上传成功: {ref['object']['sha']}")

def main():
    today = dt.now()
    print(f"=== 自动更新 {today.strftime('%Y-%m-%d %H:%M')} ===")
    
    # 先保存当前数据为备份（防止下载到旧数据）
    tmp_backup = {}
    for f in ['新装高套竣工清单.xlsx','存量高套竣工清单.xlsx','关键一单清单.xlsx',
              '杠保清单.xlsx','质态相关清单.xlsx','宽带离网清单.xlsx']:
        fp = os.path.join(DATA_DIR, f)
        if os.path.exists(fp):
            tmp_backup[f] = open(fp, 'rb').read()
    backup_month = get_data_month(os.path.join(DATA_DIR, '新装高套竣工清单.xlsx'))
    
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
    
    # ★ 月份检查：只回退比原备份更旧的数据，月初允许使用上月清单
    test_file = os.path.join(DATA_DIR, '新装高套竣工清单.xlsx')
    data_month = get_data_month(test_file)
    if data_month and backup_month and data_month < backup_month:
        print(f"⚠️ 检测到下载了 {data_month}月数据（原备份为{backup_month}月），恢复为备份数据")
        for fname, content in tmp_backup.items():
            with open(os.path.join(DATA_DIR, fname), 'wb') as f:
                f.write(content)
        # 如果备份也是旧数据，尝试从存档恢复
        if get_data_month(test_file) and backup_month and get_data_month(test_file) < backup_month:
            archive_dir = os.path.join(BACKUP_DIR, f"2026-{backup_month:02d}")
            if os.path.exists(archive_dir):
                for fn in os.listdir(archive_dir):
                    if fn.endswith('.xlsx'):
                        shutil.copy2(os.path.join(archive_dir, fn), os.path.join(DATA_DIR, fn))
                print(f"✅ 从存档 {archive_dir} 恢复数据")

    # 商客发展情况：下载最新全渠道做商客战报并提取
    print("🔄 商客发展情况...")
    try:
        sk_rc = update_shangke.main(download=True)
        if sk_rc != 0:
            print("  ⚠️ 商客战报无更新，继续使用现有数据")
    except Exception as e:
        print(f"  ⚠️ 商客战报处理失败: {e}")
    
    # 生成看板
    gen_failed = False
    for gen in ['generate_dashboard.py', 'generate_tangxia.py']:
        fp = os.path.join(PROJ_DIR, 'src', gen)
        print(f"🔄 {gen}...")
        try:
            r = subprocess.run(['python3', fp], capture_output=True, text=True, timeout=120)
        except subprocess.TimeoutExpired:
            print(f"❌ {gen} 生成超时")
            gen_failed = True
            continue
        if r.returncode != 0:
            gen_failed = True
            print(f"❌ {gen}: {r.stderr[-200:]}")
        else:
            print(r.stdout[-80:])
    if gen_failed:
        print("❌ 看板生成失败，未上传 GitHub Pages")
        return 1
    
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
    try:
        r = subprocess.run(['/bin/zsh', '-lc', cmds], capture_output=True, text=True, timeout=60)
        push_error = None if r.returncode == 0 else r.stderr[-200:]
    except subprocess.TimeoutExpired:
        r = None
        push_error = 'git push 超时'

    if push_error:
        print(f"⚠️ {push_error}")
        print("📡 改用 GitHub API 上传...")
        try:
            publish_via_github_api([
                ('docs/personnel-dashboard.html', 'docs/personnel-dashboard.html'),
                ('docs/tangxia_dashboard.html', 'docs/tangxia_dashboard.html'),
            ])
        except Exception as e:
            print(f"❌ GitHub API 上传失败: {e}")
            return 1
    else:
        print(r.stdout[-100:] if r.returncode == 0 else f"❌ {r.stderr[-100:]}")
    return 0

if __name__ == '__main__':
    sys.exit(main())
