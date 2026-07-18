#!/usr/bin/env python3
# 宽带分销系统 - Flask Web 应用
import os, sqlite3, json
from datetime import datetime
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, session, jsonify, g

PROJ_DIR = "/Users/mr.g/Documents/Codex/Workspace/projects/project-005-broadband-distribution-system"
DB_PATH = os.path.join(PROJ_DIR, "data", "broadband.db")

app = Flask(__name__, template_folder=os.path.join(PROJ_DIR, "src", "templates"),
            static_folder=os.path.join(PROJ_DIR, "src", "static"))
app.secret_key = "broadband-dist-2026-secret"

def get_db():
    if 'db' not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA journal_mode=WAL")
    return g.db

@app.teardown_appcontext
def close_db(exception):
    db = g.pop('db', None)
    if db: db.close()

def init_db():
    db = get_db()
    db.executescript("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        role TEXT NOT NULL CHECK(role IN ('admin','order_clerk','finance','commission')),
        name TEXT NOT NULL,
        active INTEGER DEFAULT 1,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS channels (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        contact TEXT,
        phone TEXT,
        commission_rate REAL DEFAULT 0,
        settle_method TEXT DEFAULT 'monthly',
        status TEXT DEFAULT 'active',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
        CREATE TABLE IF NOT EXISTS orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_date DATE,
            order_channel TEXT,
            order_method TEXT,
            customer_name TEXT,
            package_value REAL DEFAULT 0,
            broadband_no TEXT,
            main_card TEXT,
            sub_card_1 TEXT,
            sub_card_2 TEXT,
            id_number TEXT,
            order_code TEXT,
            install_address TEXT,
            order_status TEXT DEFAULT 'pending',
            express_no TEXT,
            contact_phone TEXT,
            channel_id INTEGER REFERENCES channels(id),
            entry_user_id INTEGER REFERENCES users(id),
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

    CREATE TABLE IF NOT EXISTS settlements (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        order_id INTEGER REFERENCES orders(id),
        channel_id INTEGER REFERENCES channels(id),
        settle_period TEXT,
        settle_amount REAL DEFAULT 0,
        commission_rate REAL DEFAULT 0,
        status TEXT DEFAULT 'pending',
        operator_id INTEGER REFERENCES users(id),
        settled_at TIMESTAMP,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS commissions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        order_id INTEGER REFERENCES orders(id),
        channel_id INTEGER REFERENCES channels(id),
        cashback_amount REAL DEFAULT 0,
        retro_amount REAL DEFAULT 0,
        retro_rule TEXT,
        status TEXT DEFAULT 'pending',
        operator_id INTEGER REFERENCES users(id),
        processed_at TIMESTAMP,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS monthly_monitor (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        order_id INTEGER REFERENCES orders(id),
        monitor_month TEXT,
        is_alive INTEGER DEFAULT 1,
        is_active INTEGER DEFAULT 1,
        traffic_mb REAL DEFAULT 0,
        call_minutes REAL DEFAULT 0,
        revenue REAL DEFAULT 0,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(order_id, monitor_month)
    );
    """)
    db.commit()

def seed_data():
    db = get_db()
    if db.execute("SELECT COUNT(*) FROM users").fetchone()[0] > 0: return
    from werkzeug.security import generate_password_hash
    pw = lambda p: generate_password_hash(p, method="pbkdf2:sha256")
    for u in [('admin',pw('admin123'),'admin','系统管理员'),
              ('clerk',pw('clerk123'),'order_clerk','订单录入员'),
              ('finance',pw('finance123'),'finance','财务结算员'),
              ('commission',pw('comm123'),'commission','佣金管理员')]:
        db.execute("INSERT INTO users (username,password,role,name) VALUES (?,?,?,?)", u)
    db.execute("INSERT INTO channels (name,contact,phone,commission_rate) VALUES (?,?,?,?)",
               ('测试渠道A','张三','13800138001',0.15))
    db.execute("INSERT INTO channels (name,contact,phone,commission_rate) VALUES (?,?,?,?)",
               ('测试渠道B','李四','13800138002',0.12))
    db.commit()

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session: return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

def role_required(*roles):
    def decorator(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            if session.get('role') not in roles: return "权限不足", 403
            return f(*args, **kwargs)
        return decorated
    return decorator

@app.route('/login', methods=['GET','POST'])
def login():
    if request.method == 'POST':
        from werkzeug.security import check_password_hash
        db = get_db()
        user = db.execute("SELECT * FROM users WHERE username=? AND active=1",
                         (request.form['username'],)).fetchone()
        if user and check_password_hash(user['password'], request.form['password']):
            session.update(user_id=user['id'], username=user['username'],
                          role=user['role'], name=user['name'])
            return redirect(url_for('dashboard'))
        return render_template('login.html', error='用户名或密码错误')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/')
@login_required
def dashboard():
    db = get_db()
    return render_template('dashboard.html',
        order_count=db.execute("SELECT COUNT(*) FROM orders").fetchone()[0],
        channel_count=db.execute("SELECT COUNT(*) FROM channels WHERE status='active'").fetchone()[0],
        settle_count=db.execute("SELECT COUNT(*) FROM settlements WHERE status='pending'").fetchone()[0],
        recent_orders=db.execute("""
            SELECT o.*, c.name as channel_name FROM orders o
            LEFT JOIN channels c ON o.channel_id=c.id
            ORDER BY o.created_at DESC LIMIT 10""").fetchall())

@app.route('/orders')
@login_required
@role_required('admin','order_clerk')
def orders():
    db = get_db()
    return render_template('orders.html',
        orders=db.execute("""
            SELECT o.*, c.name as channel_name, u.name as entry_name FROM orders o
            LEFT JOIN channels c ON o.channel_id=c.id
            LEFT JOIN users u ON o.entry_user_id=u.id
            ORDER BY o.created_at DESC LIMIT 100""").fetchall(),
        channels=db.execute("SELECT * FROM channels WHERE status='active'").fetchall())

@app.route('/orders/add', methods=['POST'])
@login_required
@role_required('admin','order_clerk')
def add_order():
    db = get_db()
    biz = request.form['order_code'].strip()
    if not biz: return jsonify({'error':'业务号不能为空'}),400
    if db.execute("SELECT id FROM orders WHERE order_code=?",(biz,)).fetchone():
        return jsonify({'error':'业务号已存在，请检查是否重复'}),400
    db.execute("INSERT INTO orders (order_date,order_channel,order_method,customer_name,"
               "package_value,broadband_no,main_card,sub_card_1,sub_card_2,id_number,"
               "order_code,install_address,express_no,contact_phone,channel_id,entry_user_id) "
               "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
        (request.form.get('order_date',''), request.form.get('order_channel',''),
         request.form.get('order_method',''), request.form.get('customer_name',''),
         float(request.form.get('package_value',0)), request.form.get('broadband_no',''),
         request.form.get('main_card',''), request.form.get('sub_card_1',''),
         request.form.get('sub_card_2',''), request.form.get('id_number',''),
         code, request.form.get('install_address',''),
         request.form.get('express_no',''), request.form.get('contact_phone',''),
         request.form['channel_id'], session['user_id']))
    db.commit()
    return jsonify({'ok':True})

@app.route('/orders/<int:oid>/delete', methods=['POST'])
@login_required
@role_required('admin')
def delete_order(oid):
    db = get_db()
    db.execute("UPDATE orders SET order_status='deleted' WHERE id=?",(oid,))
    db.commit()
    return jsonify({'ok':True})

@app.route('/orders/batch-import', methods=['POST'])
@login_required
@role_required('admin','order_clerk')
def batch_import():
    import openpyxl, io
    file = request.files.get('file')
    if not file: return jsonify({'error':'请选择文件'}),400
    try:
        ws = openpyxl.load_workbook(io.BytesIO(file.read()),data_only=True).active
        db = get_db()
        imported = skipped = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]: continue
            biz = str(row[0]).strip()
            if not biz: continue
            if db.execute("SELECT id FROM orders WHERE order_code=?",(biz,)).fetchone():
                skipped += 1; continue
            cid = row[1] if len(row)>1 else None
            db.execute("INSERT INTO orders (order_date,order_channel,order_method,customer_name,"
                       "package_value,broadband_no,main_card,sub_card_1,sub_card_2,id_number,"
                       "order_code,install_address,express_no,contact_phone,entry_user_id) "
                       "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (str(row[0]) if len(row)>0 and row[0] else '',
                 str(row[1]) if len(row)>1 and row[1] else '',
                 str(row[2]) if len(row)>2 and row[2] else '',
                 str(row[3]) if len(row)>3 and row[3] else '',
                 float(row[4]) if len(row)>4 and row[4] else 0,
                 str(row[5]) if len(row)>5 and row[5] else '',
                 str(row[6]) if len(row)>6 and row[6] else '',
                 str(row[7]) if len(row)>7 and row[7] else '',
                 str(row[8]) if len(row)>8 and row[8] else '',
                 str(row[9]) if len(row)>9 and row[9] else '',
                 biz,
                 str(row[11]) if len(row)>11 and row[11] else '',
                 str(row[12]) if len(row)>12 and row[12] else '',
                 str(row[13]) if len(row)>13 and row[13] else '',
                 session['user_id']))
            imported += 1
        db.commit()
        return jsonify({'ok':True, 'imported':imported, 'skipped':skipped})
    except Exception as e:
        return jsonify({'error':f'导入失败: {str(e)}'}),400

@app.route('/settlements')
@login_required
@role_required('admin','finance')
def settlements():
    db = get_db()
    return render_template('settlements.html',
        settlements=db.execute("""
            SELECT s.*, o.order_code, c.name as channel_name FROM settlements s
            JOIN orders o ON s.order_id=o.id
            LEFT JOIN channels c ON s.channel_id=c.id
            ORDER BY s.created_at DESC LIMIT 100""").fetchall(),
        channels=db.execute("SELECT * FROM channels WHERE status='active'").fetchall())

@app.route('/settlements/generate', methods=['POST'])
@login_required
@role_required('admin','finance')
def generate_settlement():
    db = get_db()
    period = request.form.get('period', datetime.now().strftime('%Y-%m'))
    cid = request.form.get('channel_id')
    rows = db.execute("""
        SELECT o.id, o.channel_id, o.package_value, c.commission_rate FROM orders o
        JOIN channels c ON o.channel_id=c.id
        WHERE o.status='active' AND o.id NOT IN (
            SELECT order_id FROM settlements WHERE settle_period=?)""" +
        (" AND o.channel_id=?" if cid else ""), [period] + ([cid] if cid else [])).fetchall()
    count = 0
    for r in rows:
        amt = r['package_value'] * (r['commission_rate'] or 0)
        db.execute("INSERT INTO settlements (order_id,channel_id,settle_period,"
                   "settle_amount,commission_rate,status,operator_id) VALUES (?,?,?,?,?,'pending',?)",
            (r['id'], r['channel_id'], period, amt, r['commission_rate'], session['user_id']))
        count += 1
    db.commit()
    return jsonify({'ok':True, 'count':count})

@app.route('/settlements/<int:sid>/confirm', methods=['POST'])
@login_required
@role_required('admin')
def confirm_settlement(sid):
    db = get_db()
    db.execute("UPDATE settlements SET status='settled', settled_at=CURRENT_TIMESTAMP WHERE id=?",(sid,))
    db.commit()
    return jsonify({'ok':True})

@app.route('/commissions')
@login_required
@role_required('admin','commission')
def commissions():
    db = get_db()
    return render_template('commissions.html',
        commissions=db.execute("""
            SELECT cm.*, o.order_code, c.name as channel_name FROM commissions cm
            JOIN orders o ON cm.order_id=o.id
            LEFT JOIN channels c ON cm.channel_id=c.id
            ORDER BY cm.created_at DESC LIMIT 100""").fetchall(),
        orders=db.execute("""
            SELECT o.id, o.order_code, c.name as channel_name FROM orders o
            LEFT JOIN channels c ON o.channel_id=c.id
            WHERE o.order_status='active' ORDER BY o.created_at DESC""").fetchall())

@app.route('/commissions/retro', methods=['POST'])
@login_required
@role_required('admin','commission')
def retro_commission():
    db = get_db()
    oid = request.form['order_id']
    order = db.execute("SELECT * FROM orders WHERE id=?",(oid,)).fetchone()
    if not order: return jsonify({'error':'订单不存在'}),400
    db.execute("INSERT INTO commissions (order_id,channel_id,cashback_amount,"
               "retro_amount,retro_rule,status,operator_id) VALUES (?,?,?,?,?,'processed',?)",
        (oid, order['channel_id'], float(request.form.get('cashback',0)),
         float(request.form.get('retro_amount',0)), request.form.get('rule',''),
         session['user_id']))
    db.commit()
    return jsonify({'ok':True})

@app.route('/channels')
@login_required
@role_required('admin')
def channels():
    db = get_db()
    return render_template('channels.html',
        channels=db.execute("""
            SELECT c.*, (SELECT COUNT(*) FROM orders WHERE channel_id=c.id) as order_count
            FROM channels c ORDER BY c.created_at DESC""").fetchall())

@app.route('/channels/add', methods=['POST'])
@login_required
@role_required('admin')
def add_channel():
    db = get_db()
    db.execute("INSERT INTO channels (name,contact,phone,commission_rate,settle_method) "
               "VALUES (?,?,?,?,?)",
        (request.form['name'], request.form.get('contact',''),
         request.form.get('phone',''), float(request.form.get('commission_rate',0)),
         request.form.get('settle_method','monthly')))
    db.commit()
    return jsonify({'ok':True})

@app.route('/monitor')
@login_required
@role_required('admin','finance','commission')
def monitor():
    db = get_db()
    month = request.args.get('month', datetime.now().strftime('%Y-%m'))
    return render_template('monitor.html',
        stats=db.execute("""
            SELECT c.name as channel_name,
                COUNT(DISTINCT o.id) as total_orders,
                COUNT(DISTINCT CASE WHEN mm.is_alive=1 THEN o.id END) as alive_count,
                COUNT(DISTINCT CASE WHEN mm.is_active=1 THEN o.id END) as active_count,
                COALESCE(SUM(mm.revenue),0) as total_revenue
            FROM channels c LEFT JOIN orders o ON o.channel_id=c.id AND o.order_status='active'
            LEFT JOIN monthly_monitor mm ON mm.order_id=o.id AND mm.monitor_month=?
            GROUP BY c.id""", (month,)).fetchall(),
        month=month)

@app.route('/users')
@login_required
@role_required('admin')
def users():
    return render_template('users.html',
        users=get_db().execute(
            "SELECT id,username,role,name,active,created_at FROM users").fetchall())

@app.route('/users/add', methods=['POST'])
@login_required
@role_required('admin')
def add_user():
    from werkzeug.security import generate_password_hash
    db = get_db()
    try:
        db.execute("INSERT INTO users (username,password,role,name) VALUES (?,?,?,?)",
            (request.form['username'], generate_password_hash(request.form['password']),
             request.form['role'], request.form['name']))
        db.commit()
        return jsonify({'ok':True})
    except sqlite3.IntegrityError:
        return jsonify({'error':'用户名已存在'}),400

if __name__ == '__main__':
    with app.app_context():
        init_db()
        seed_data()
    print(f"数据库: {DB_PATH}")
    app.run(host='127.0.0.1', port=8766, debug=True)
