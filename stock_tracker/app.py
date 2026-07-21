from flask import Flask, render_template, request, jsonify, session, Response
import sqlite3
import os
import csv
import io
from datetime import datetime

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'ast_stock_2026_secret')

EDITOR_PIN  = os.environ.get('EDITOR_PIN',  'ast2026')
VIEWER_PIN  = os.environ.get('VIEWER_PIN',  '')
DB_PATH     = os.path.join(os.path.dirname(__file__), 'stock.db')


# ── DB helpers ────────────────────────────────────────────────────────────────

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    conn.executescript('''
        CREATE TABLE IF NOT EXISTS items (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            name        TEXT NOT NULL UNIQUE,
            category    TEXT NOT NULL DEFAULT '',
            unit        TEXT NOT NULL DEFAULT 'pcs',
            description TEXT NOT NULL DEFAULT '',
            created_at  TEXT NOT NULL DEFAULT (datetime('now','localtime'))
        );
        CREATE TABLE IF NOT EXISTS transactions (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            item_id    INTEGER NOT NULL REFERENCES items(id),
            type       TEXT NOT NULL CHECK(type IN ('IN','OUT')),
            quantity   REAL NOT NULL CHECK(quantity > 0),
            notes      TEXT NOT NULL DEFAULT '',
            reference  TEXT NOT NULL DEFAULT '',
            logged_by  TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
        );
        CREATE TABLE IF NOT EXISTS transaction_subitems (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            transaction_id INTEGER NOT NULL REFERENCES transactions(id),
            name           TEXT NOT NULL,
            notes          TEXT NOT NULL DEFAULT ''
        );
    ''')
    conn.commit()
    conn.close()


def resolve_item(conn, name):
    """Return item id for name, creating the item if it does not exist."""
    name = (name or '').strip()
    if not name:
        return None
    row = conn.execute(
        'SELECT id FROM items WHERE LOWER(name)=LOWER(?)', (name,)
    ).fetchone()
    if row:
        return row['id']
    conn.execute('INSERT INTO items (name) VALUES (?)', (name,))
    return conn.execute('SELECT last_insert_rowid()').fetchone()[0]


# ── Auth ──────────────────────────────────────────────────────────────────────

@app.route('/api/auth', methods=['POST'])
def auth():
    pin = (request.json or {}).get('pin', '')
    if pin == EDITOR_PIN:
        session['role'] = 'editor'
        return jsonify({'ok': True, 'role': 'editor'})
    if VIEWER_PIN and pin == VIEWER_PIN:
        session['role'] = 'viewer'
        return jsonify({'ok': True, 'role': 'viewer'})
    return jsonify({'ok': False, 'msg': 'Incorrect PIN'}), 401


@app.route('/api/auth/check')
def auth_check():
    return jsonify({'role': session.get('role', 'viewer')})


@app.route('/api/auth/logout', methods=['POST'])
def logout():
    session.clear()
    return jsonify({'ok': True})


# ── Items ─────────────────────────────────────────────────────────────────────

@app.route('/api/items', methods=['GET'])
def get_items():
    conn = get_db()
    rows = conn.execute('''
        SELECT
            i.id, i.name, i.category, i.unit, i.description, i.created_at,
            COALESCE(SUM(CASE WHEN t.type='IN' THEN t.quantity
                              WHEN t.type='OUT' THEN -t.quantity END), 0) AS current_stock,
            MAX(t.created_at)  AS last_transaction,
            COUNT(t.id)        AS tx_count
        FROM items i
        LEFT JOIN transactions t ON t.item_id = i.id
        GROUP BY i.id
        ORDER BY i.category, i.name
    ''').fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route('/api/items', methods=['POST'])
def add_item():
    if session.get('role') != 'editor':
        return jsonify({'error': 'Unauthorized'}), 403
    data = request.json or {}
    name = (data.get('name') or '').strip()
    if not name:
        return jsonify({'error': 'Name is required'}), 400
    conn = get_db()
    try:
        conn.execute(
            'INSERT INTO items (name, category, unit, description) VALUES (?,?,?,?)',
            (name, data.get('category','').strip(),
             data.get('unit','pcs').strip(),
             data.get('description','').strip())
        )
        conn.commit()
    except sqlite3.IntegrityError:
        conn.close()
        return jsonify({'error': f'Item "{name}" already exists'}), 400
    conn.close()
    return jsonify({'ok': True})


@app.route('/api/items/<int:item_id>', methods=['DELETE'])
def delete_item(item_id):
    if session.get('role') != 'editor':
        return jsonify({'error': 'Unauthorized'}), 403
    conn = get_db()
    # Cascade: remove subitems → transactions → item
    tx_ids = [r['id'] for r in conn.execute(
        'SELECT id FROM transactions WHERE item_id=?', (item_id,)
    ).fetchall()]
    if tx_ids:
        ph = ','.join('?' * len(tx_ids))
        conn.execute(f'DELETE FROM transaction_subitems WHERE transaction_id IN ({ph})', tx_ids)
    conn.execute('DELETE FROM transactions WHERE item_id=?', (item_id,))
    conn.execute('DELETE FROM items WHERE id=?', (item_id,))
    conn.commit()
    conn.close()
    return jsonify({'ok': True})


# ── Transactions ──────────────────────────────────────────────────────────────

@app.route('/api/transactions', methods=['GET'])
def get_transactions():
    item_id = request.args.get('item_id')
    tx_type = request.args.get('type')
    limit   = min(int(request.args.get('limit', 500)), 2000)

    query  = '''
        SELECT t.id, t.item_id, i.name AS item_name, i.unit,
               i.category, t.type, t.quantity, t.notes,
               t.reference, t.logged_by, t.created_at
        FROM transactions t
        JOIN items i ON i.id = t.item_id
    '''
    params, conds = [], []
    if item_id:
        conds.append('t.item_id = ?'); params.append(item_id)
    if tx_type:
        conds.append('t.type = ?'); params.append(tx_type)
    if conds:
        query += ' WHERE ' + ' AND '.join(conds)
    query += ' ORDER BY t.created_at DESC LIMIT ?'
    params.append(limit)

    conn = get_db()
    rows = conn.execute(query, params).fetchall()
    tx_list = [dict(r) for r in rows]

    # Attach subitems in one additional query (avoids N+1)
    if tx_list:
        tx_ids = [t['id'] for t in tx_list]
        ph = ','.join('?' * len(tx_ids))
        subs = conn.execute(
            f'SELECT * FROM transaction_subitems WHERE transaction_id IN ({ph}) ORDER BY id',
            tx_ids
        ).fetchall()
        sub_map = {}
        for s in subs:
            sub_map.setdefault(s['transaction_id'], []).append(
                {'id': s['id'], 'name': s['name'], 'notes': s['notes']}
            )
        for t in tx_list:
            t['subitems'] = sub_map.get(t['id'], [])

    conn.close()
    return jsonify(tx_list)


@app.route('/api/transactions', methods=['POST'])
def add_transaction():
    if session.get('role') != 'editor':
        return jsonify({'error': 'Unauthorized'}), 403
    data = request.json or {}

    conn = get_db()

    # Resolve item — accept either item_id or item_name (auto-creates if new)
    item_id = data.get('item_id')
    if not item_id:
        item_id = resolve_item(conn, data.get('item_name', ''))
    if not item_id:
        conn.close()
        return jsonify({'error': 'Item name is required'}), 400

    try:
        qty = float(data['quantity'])
        assert qty > 0
    except (KeyError, ValueError, AssertionError):
        conn.close()
        return jsonify({'error': 'Quantity must be a positive number'}), 400

    if data.get('type') not in ('IN', 'OUT'):
        conn.close()
        return jsonify({'error': 'Type must be IN or OUT'}), 400

    if data['type'] == 'OUT':
        row = conn.execute('''
            SELECT COALESCE(SUM(CASE WHEN type='IN' THEN quantity ELSE -quantity END),0) AS stock
            FROM transactions WHERE item_id=?
        ''', (item_id,)).fetchone()
        if row and row['stock'] < qty:
            conn.close()
            return jsonify({'error': f'Insufficient stock (available: {row["stock"]})'}), 400

    conn.execute(
        'INSERT INTO transactions (item_id,type,quantity,notes,reference,logged_by) VALUES (?,?,?,?,?,?)',
        (item_id, data['type'], qty,
         data.get('notes','').strip(),
         data.get('reference','').strip(),
         data.get('logged_by','').strip())
    )
    tx_id = conn.execute('SELECT last_insert_rowid()').fetchone()[0]

    for sub in (data.get('subitems') or []):
        name = (sub.get('name') or '').strip()
        if name:
            conn.execute(
                'INSERT INTO transaction_subitems (transaction_id, name, notes) VALUES (?,?,?)',
                (tx_id, name, (sub.get('notes') or '').strip())
            )

    conn.commit()
    conn.close()
    return jsonify({'ok': True})


@app.route('/api/transactions/<int:tx_id>', methods=['DELETE'])
def delete_transaction(tx_id):
    if session.get('role') != 'editor':
        return jsonify({'error': 'Unauthorized'}), 403
    conn = get_db()
    conn.execute('DELETE FROM transaction_subitems WHERE transaction_id=?', (tx_id,))
    conn.execute('DELETE FROM transactions WHERE id=?', (tx_id,))
    conn.commit()
    conn.close()
    return jsonify({'ok': True})


# ── Export ────────────────────────────────────────────────────────────────────

@app.route('/api/export/csv')
def export_csv():
    conn = get_db()
    rows = conn.execute('''
        SELECT t.created_at AS "Date/Time", i.name AS "Item",
               i.category AS "Category", t.type AS "Type",
               t.quantity AS "Quantity", i.unit AS "Unit",
               t.reference AS "Reference", t.notes AS "Notes",
               t.logged_by AS "Logged By"
        FROM transactions t
        JOIN items i ON i.id = t.item_id
        ORDER BY t.created_at DESC
    ''').fetchall()
    conn.close()
    buf = io.StringIO()
    writer = csv.writer(buf)
    if rows:
        writer.writerow([d[0] for d in rows[0].description])
    else:
        writer.writerow(['Date/Time','Item','Category','Type','Quantity','Unit','Reference','Notes','Logged By'])
    for r in rows:
        writer.writerow(list(r))
    filename = f'stock_export_{datetime.now().strftime("%Y%m%d_%H%M")}.csv'
    return Response(buf.getvalue(), mimetype='text/csv',
                    headers={'Content-Disposition': f'attachment; filename={filename}'})


# ── Main ──────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


if __name__ == '__main__':
    init_db()
    import socket
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
    except Exception:
        ip = '127.0.0.1'
    print(f'\n  Stock Tracker running at:')
    print(f'  Local:   http://127.0.0.1:5000')
    print(f'  Network: http://{ip}:5000  <-- share this link\n')
    app.run(host='0.0.0.0', port=5000, debug=False)
