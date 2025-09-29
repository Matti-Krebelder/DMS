from flask import Flask, render_template, request, redirect, url_for, session, jsonify, send_file
from flask_cors import CORS
from markupsafe import Markup
import sqlite3
import os
import random
import string
from datetime import datetime
import qrcode
import io
import base64
import shutil
import csv
from docx import Document
from docx.shared import Inches, Pt, Cm
from io import BytesIO
import json
import re # Import re for regex operations
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}}, supports_credentials=True)
app.secret_key = 'your-secret-key'

def init_user_db():
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS users
                 (id TEXT PRIMARY KEY, name TEXT NOT NULL)''')
    c.execute('''CREATE TABLE IF NOT EXISTS lager
                 (id TEXT PRIMARY KEY, name TEXT NOT NULL, created_by TEXT, 
                   access_users TEXT, system_type TEXT DEFAULT 'personal')''')
    c.execute("INSERT OR IGNORE INTO users VALUES ('CKS-123-432-132', 'Matti')")
    c.execute("INSERT OR IGNORE INTO users VALUES ('CKS-456-789-012', 'Hubert')")
    c.execute("INSERT OR IGNORE INTO users VALUES ('CKS-789-012-345', 'Admin')")
    c.execute("INSERT OR IGNORE INTO users VALUES ('CKS-725-283-382', 'Christoffer Rentsch')")
    c.execute("INSERT OR IGNORE INTO users VALUES ('CKS-123-456-789', 'Steffen Mascher')")
    conn.commit()
    conn.close()

def create_warehouse_db(lager_id):
    conn = sqlite3.connect(f'{lager_id}.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE geraete
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                   name TEXT NOT NULL,
                   barcode TEXT UNIQUE NOT NULL,
                   lagerplatz TEXT NOT NULL,
                   status TEXT DEFAULT 'verfügbar',
                   beschreibung TEXT,
                   seriennummer TEXT,
                   modell TEXT,
                   instrumentenart TEXT,
                   inventarnummer TEXT,
                   kaufdatum TEXT,
                   preis REAL)''')
    c.execute('''CREATE TABLE ausleihen
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                   ausleih_id TEXT NOT NULL,
                   mitarbeiter_id TEXT NOT NULL,
                   mitarbeiter_name TEXT NOT NULL,
                   zielort TEXT NOT NULL,
                   datum TEXT NOT NULL,
                   rueckgabe_qr TEXT NOT NULL,
                   status TEXT DEFAULT 'ausgeliehen',
                   email TEXT,
                   klasse TEXT)''')
    c.execute('''CREATE TABLE ausleih_details
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                   ausleih_id TEXT NOT NULL,
                   geraet_id INTEGER NOT NULL,
                   geraet_barcode TEXT NOT NULL,
                   FOREIGN KEY(ausleih_id) REFERENCES ausleihen(ausleih_id),
                   FOREIGN KEY(geraet_id) REFERENCES geraete(id))''')
    c.execute('''CREATE TABLE label_layouts
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                   name TEXT NOT NULL,
                   layout_data TEXT NOT NULL,
                   is_default INTEGER DEFAULT 0,
                   created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                   updated_at TEXT DEFAULT CURRENT_TIMESTAMP)''')
    conn.commit()
    conn.close()

def generate_random_id(length=6):
    return ''.join(random.choices(string.digits, k=length))

def get_db_connection(lager_id):
    return sqlite3.connect(f'{lager_id}.db')

def generate_qr_code(data):
    qr = qrcode.QRCode(version=1, box_size=10, border=5)
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buffer = io.BytesIO()
    img.save(buffer, format='PNG')
    buffer.seek(0)
    img_str = base64.b64encode(buffer.getvalue()).decode()
    return f"data:image/png;base64,{img_str}"

def get_lager_system_type(lager_id):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute("SELECT system_type FROM lager WHERE id = ?", (lager_id,))
    result = c.fetchone()
    conn.close()
    return result[0] if result else 'personal'

def backup_db(lager_id, operation):
    os.makedirs('backups', exist_ok=True)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    shutil.copy(f'{lager_id}.db', f'backups/{timestamp}_{operation}_{lager_id}.db')

@app.route('/')
def login():
    if 'user_id' in session:
        return redirect(url_for('dashboard'))
    return render_template('login.html', title="Login")

@app.route('/login', methods=['POST'])
def do_login():
    user_id = request.form['user_id']
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute("SELECT name FROM users WHERE id = ?", (user_id,))
    user = c.fetchone()
    conn.close()
    if user:
        session['user_id'] = user_id
        session['user_name'] = user[0]
        return redirect(url_for('dashboard'))
    else:
        return redirect(url_for('login'))

@app.route('/dashboard')
def dashboard():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute("SELECT id, name FROM lager WHERE created_by = ? OR access_users LIKE ?", 
              (session['user_id'], f"%{session['user_id']}%"))
    lagers = c.fetchall()
    conn.close()
    return render_template('dashboard.html', title="Dashboard", lagers=lagers)

@app.route('/create_lager', methods=['GET', 'POST'])
def create_lager():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    if request.method == 'POST':
        name = request.form['name']
        access_users = request.form.getlist('access_users')
        system_type = request.form.get('system_type', 'personal')
        lager_id = generate_random_id(8)
        while os.path.exists(f'{lager_id}.db'):
            lager_id = generate_random_id(8)
        conn = sqlite3.connect('users.db')
        c = conn.cursor()
        c.execute("INSERT INTO lager VALUES (?, ?, ?, ?, ?)", 
                  (lager_id, name, session['user_id'], ','.join(access_users), system_type))
        conn.commit()
        conn.close()
        create_warehouse_db(lager_id)
        return redirect(url_for('warehouse', lager_id=lager_id))
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute("SELECT id, name FROM users WHERE id != ?", (session['user_id'],))
    users = c.fetchall()
    conn.close()
    return render_template('create_lager.html', title="Neues Lager", users=users)

@app.route('/lager/<lager_id>')
def warehouse(lager_id):
    if 'user_id' not in session:
        return redirect(url_for('login'))
    if not os.path.exists(f'{lager_id}.db'):
        return redirect(url_for('dashboard'))
    session['current_lager'] = lager_id
    return render_template('warehouse.html', title="Lager", lager_id=lager_id)

@app.route('/devices')
def devices():
    if 'current_lager' not in session:
        return redirect(url_for('dashboard'))
    
    search = request.args.get('search', '')
    status_filters = [f for f in request.args.getlist('status') if f]
    art_filters = [f for f in request.args.getlist('art') if f]
    klasse_filters = [f for f in request.args.getlist('klasse') if f]
    sort_by = request.args.get('sort_by', 'name')
    group_by = request.args.get('group_by', 'model')  # Added grouping option
    
    conn = get_db_connection(session['current_lager'])
    c = conn.cursor()
    
    # Build query with filters
    query = """SELECT g.* FROM geraete g
               LEFT JOIN ausleih_details ad ON g.id = ad.geraet_id
               LEFT JOIN ausleihen a ON ad.ausleih_id = a.ausleih_id AND a.status = 'ausgeliehen'
               WHERE 1=1"""
    params = []
    
    if search:
        query += " AND (g.name LIKE ? OR g.barcode LIKE ? OR g.lagerplatz LIKE ? OR g.seriennummer LIKE ? OR g.modell LIKE ? OR g.instrumentenart LIKE ? OR a.mitarbeiter_name LIKE ? OR a.klasse LIKE ?)"
        params.extend([f'%{search}%'] * 8)
    
    if status_filters:
        status_conditions = []
        if "verfügbar" in status_filters:
            status_conditions.append("g.status = 'verfügbar'")
        if "ausgeliehen" in status_filters:
            status_conditions.append("g.status LIKE 'ausgeliehen%'")
        if status_conditions:
            query += f" AND ({' OR '.join(status_conditions)})"
    
    if art_filters:
        placeholders = ','.join('?' for _ in art_filters)
        query += f" AND g.instrumentenart IN ({placeholders})"
        params.extend(art_filters)
    
    if klasse_filters:
        placeholders = ','.join('?' for _ in klasse_filters)
        query += f" AND a.klasse IN ({placeholders})"
        params.extend(klasse_filters)
    
    # Handle sorting
    if sort_by == 'instrumentenart':
        query += " ORDER BY g.instrumentenart, g.name"
    elif sort_by == 'lagerplatz':
        query += " ORDER BY g.lagerplatz, g.name"
    elif sort_by == 'status':
        query += " ORDER BY g.status, g.name"
    elif sort_by == 'model':  # Added model sorting
        query += " ORDER BY g.modell, g.name"
    else:
        query += " ORDER BY g.name"
    
    c.execute(query, params)
    devices_list = c.fetchall()
    
    grouped_devices = {}
    if group_by == 'model':
        for device in devices_list:
            model = device[7] or 'Unbekanntes Modell'  # modell is at index 7
            if model not in grouped_devices:
                grouped_devices[model] = []
            grouped_devices[model].append(device)
    elif group_by == 'instrument':
        for device in devices_list:
            instrument = device[8] or 'Unbekanntes Instrument'  # instrumentenart is at index 8
            if instrument not in grouped_devices:
                grouped_devices[instrument] = []
            grouped_devices[instrument].append(device)
    elif group_by == 'status':
        for device in devices_list:
            status = 'Verfügbar' if device[4] == 'verfügbar' else 'Ausgeliehen'
            if status not in grouped_devices:
                grouped_devices[status] = []
            grouped_devices[status].append(device)
    else:
        # No grouping - put all devices in one group
        grouped_devices['Alle Geräte'] = devices_list
    
    # Get filter options
    c.execute("SELECT DISTINCT instrumentenart FROM geraete WHERE instrumentenart IS NOT NULL")
    instrumentenarten = [row[0] for row in c.fetchall()]
    c.execute("SELECT DISTINCT klasse FROM ausleihen WHERE klasse IS NOT NULL")
    klassen = [row[0] for row in c.fetchall()]
    conn.close()
    
    # Handle AJAX requests for live search
    if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
        return render_template('devices_list.html', devices=devices_list, grouped_devices=grouped_devices, group_by=group_by)
    
    return render_template('devices.html', title="Geräte", devices=devices_list, grouped_devices=grouped_devices,
                         search=search, status_filters=status_filters, 
                         art_filters=art_filters, klasse_filters=klasse_filters,
                         sort_by=sort_by, group_by=group_by, instrumentenarten=instrumentenarten, klassen=klassen)

@app.route('/add_device', methods=['GET', 'POST'])
def add_device():
    if 'current_lager' not in session:
        return redirect(url_for('dashboard'))
    
    if request.method == 'POST':
        name = request.form['name']
        lagerplatz = request.form['lagerplatz']
        beschreibung = request.form['beschreibung']
        seriennummer = request.form['seriennummer']
        modell = request.form['modell']
        instrumentenart = request.form['instrumentenart']
        inventarnummer = request.form['inventarnummer']
        kaufdatum = request.form['kaufdatum']
        preis = request.form.get('preis', 0)
        
        conn = get_db_connection(session['current_lager'])
        c = conn.cursor()
        
        # Generate unique barcode
        while True:
            barcode = generate_random_id(6)
            c.execute("SELECT id FROM geraete WHERE barcode = ?", (barcode,))
            if not c.fetchone():
                break
        
        backup_db(session['current_lager'], 'before_add_device')
        c.execute("INSERT INTO geraete (name, barcode, lagerplatz, beschreibung, seriennummer, modell, instrumentenart, inventarnummer, kaufdatum, preis) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                  (name, barcode, lagerplatz, beschreibung, seriennummer, modell, instrumentenart, inventarnummer, kaufdatum, preis))
        conn.commit()
        backup_db(session['current_lager'], 'after_add_device')
        conn.close()
        return redirect(url_for('devices'))
    
    # Get existing instrument types for datalist
    conn = get_db_connection(session['current_lager'])
    c = conn.cursor()
    c.execute("SELECT DISTINCT instrumentenart FROM geraete WHERE instrumentenart IS NOT NULL")
    instrumentenarten = [row[0] for row in c.fetchall()]
    conn.close()
    
    return render_template('add_device.html', title="Gerät hinzufügen", instrumentenarten=instrumentenarten)

@app.route('/edit_device/<int:device_id>', methods=['GET', 'POST'])
def edit_device(device_id):
    if 'current_lager' not in session:
        return redirect(url_for('dashboard'))
    
    conn = get_db_connection(session['current_lager'])
    c = conn.cursor()
    
    if request.method == 'POST':
        name = request.form['name']
        barcode = request.form['barcode']
        lagerplatz = request.form['lagerplatz']
        beschreibung = request.form['beschreibung']
        seriennummer = request.form['seriennummer']
        modell = request.form['modell']
        instrumentenart = request.form['instrumentenart']
        inventarnummer = request.form['inventarnummer']
        kaufdatum = request.form['kaufdatum']
        preis = request.form.get('preis', 0)
        
        # Check if barcode is unique (excluding current device)
        c.execute("SELECT id FROM geraete WHERE barcode = ? AND id != ?", (barcode, device_id))
        if c.fetchone():
            conn.close()
            return redirect(url_for('edit_device', device_id=device_id))
        
        backup_db(session['current_lager'], 'before_edit_device')
        c.execute("UPDATE geraete SET name = ?, barcode = ?, lagerplatz = ?, beschreibung = ?, seriennummer = ?, modell = ?, instrumentenart = ?, inventarnummer = ?, kaufdatum = ?, preis = ? WHERE id = ?",
                  (name, barcode, lagerplatz, beschreibung, seriennummer, modell, instrumentenart, inventarnummer, kaufdatum, preis, device_id))
        conn.commit()
        backup_db(session['current_lager'], 'after_edit_device')
        conn.close()
        return redirect(url_for('devices'))
    
    # Get device data and instrument types
    c.execute("SELECT * FROM geraete WHERE id = ?", (device_id,))
    device = c.fetchone()
    c.execute("SELECT DISTINCT instrumentenart FROM geraete WHERE instrumentenart IS NOT NULL")
    instrumentenarten = [row[0] for row in c.fetchall()]
    conn.close()
    
    return render_template('edit_device.html', title="Gerät bearbeiten", device=device, instrumentenarten=instrumentenarten)

@app.route('/delete_device/<int:device_id>')
def delete_device(device_id):
    if 'current_lager' not in session:
        return redirect(url_for('dashboard'))
    
    conn = get_db_connection(session['current_lager'])
    c = conn.cursor()
    backup_db(session['current_lager'], 'before_delete_device')
    c.execute("DELETE FROM geraete WHERE id = ?", (device_id,))
    conn.commit()
    backup_db(session['current_lager'], 'after_delete_device')
    conn.close()
    return redirect(url_for('devices'))

@app.route('/borrow', methods=['GET', 'POST'])
def borrow():
    if 'current_lager' not in session:
        return redirect(url_for('dashboard'))
    
    system_type = get_lager_system_type(session['current_lager'])
    
    if request.method == 'POST':
        if 'add_device' in request.form:
            barcode = request.form['barcode']
            if 'borrow_list' not in session:
                session['borrow_list'] = []
            
            conn = get_db_connection(session['current_lager'])
            c = conn.cursor()
            c.execute("SELECT id, name, barcode FROM geraete WHERE barcode = ? AND status = 'verfügbar'", (barcode,))
            device = c.fetchone()
            conn.close()
            
            if device and device[0] not in [d['id'] for d in session['borrow_list']]:
                session['borrow_list'].append({
                    'id': device[0], 'name': device[1], 'barcode': device[2]
                })
                session.modified = True
                
        elif 'complete_borrow' in request.form:
            if system_type == 'personal':
                borrower_name = session['user_name']
                borrower_id = session['user_id']
                email = klasse = None
            else:
                borrower_name = request.form['borrower_name']
                borrower_id = request.form.get('borrower_id', 'N/A')
                email = request.form.get('email')
                klasse = request.form.get('klasse')
            
            if session.get('borrow_list'):
                ausleih_id = generate_random_id(4)
                conn = get_db_connection(session['current_lager'])
                c = conn.cursor()
                backup_db(session['current_lager'], 'before_borrow')
                
                c.execute("INSERT INTO ausleihen (ausleih_id, mitarbeiter_id, mitarbeiter_name, zielort, datum, rueckgabe_qr, email, klasse) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                          (ausleih_id, borrower_id, borrower_name, 'N/A',
                           datetime.now().strftime('%Y-%m-%d %H:%M:%S'), ausleih_id, email, klasse))
                
                for device in session['borrow_list']:
                    c.execute("INSERT INTO ausleih_details (ausleih_id, geraet_id, geraet_barcode) VALUES (?, ?, ?)",
                              (ausleih_id, device['id'], device['barcode']))
                    c.execute("UPDATE geraete SET status = ? WHERE id = ?",
                              (f"ausgeliehen an {borrower_name}", device['id']))
                
                conn.commit()
                backup_db(session['current_lager'], 'after_borrow')
                conn.close()
                session['borrow_list'] = []
                session.modified = True
                return redirect(url_for('borrow_success', ausleih_id=ausleih_id))
    
    borrow_list = session.get('borrow_list', [])
    return render_template('borrow.html', title="Ausleihen", borrow_list=borrow_list, system_type=system_type)

@app.route('/borrow_success/<ausleih_id>')
def borrow_success(ausleih_id):
    if 'current_lager' not in session:
        return redirect(url_for('dashboard'))
    
    system_type = get_lager_system_type(session['current_lager'])
    conn = get_db_connection(session['current_lager'])
    c = conn.cursor()
    
    c.execute("""SELECT a.*, GROUP_CONCAT(g.name, ', ') as devices, GROUP_CONCAT(g.barcode, ', ') as barcodes
                 FROM ausleihen a
                 JOIN ausleih_details ad ON a.ausleih_id = ad.ausleih_id
                 JOIN geraete g ON ad.geraet_id = g.id
                 WHERE a.ausleih_id = ?""", (ausleih_id,))
    borrow = c.fetchone()
    
    c.execute("""SELECT g.name, g.barcode
                 FROM ausleih_details ad
                 JOIN geraete g ON ad.geraet_id = g.id
                 WHERE ad.ausleih_id = ?""", (ausleih_id,))
    devices = c.fetchall()
    conn.close()
    
    return render_template('borrow_success.html', title="Ausleihe erfolgreich", 
                         borrow=borrow, devices=devices, system_type=system_type, 
                         generate_qr_code=generate_qr_code)

@app.route('/return', methods=['GET', 'POST'])
def return_devices():
    if 'current_lager' not in session:
        return redirect(url_for('dashboard'))
    
    system_type = get_lager_system_type(session['current_lager'])
    
    if request.method == 'POST':
        if 'scan_qr' in request.form:
            qr_code = request.form['qr_code']
            return redirect(url_for('return_devices', qr=qr_code))
        elif 'complete_return' in request.form:
            ausleih_id = request.form['ausleih_id']
            device_ids = request.form.getlist('return_devices')
            
            conn = get_db_connection(session['current_lager'])
            c = conn.cursor()
            backup_db(session['current_lager'], 'before_return')
            
            for device_id in device_ids:
                c.execute("UPDATE geraete SET status = 'verfügbar' WHERE id = ?", (device_id,))
                c.execute("DELETE FROM ausleih_details WHERE ausleih_id = ? AND geraet_id = ?", (ausleih_id, device_id))
            
            c.execute("SELECT COUNT(*) FROM ausleih_details WHERE ausleih_id = ?", (ausleih_id,))
            remaining = c.fetchone()[0]
            if remaining == 0:
                c.execute("UPDATE ausleihen SET status = 'zurückgegeben' WHERE ausleih_id = ?", (ausleih_id,))
            
            conn.commit()
            backup_db(session['current_lager'], 'after_return')
            conn.close()
            return redirect(url_for('return_devices'))
    
    qr_code = request.args.get('qr')
    devices_to_return = []
    
    if qr_code:
        conn = get_db_connection(session['current_lager'])
        c = conn.cursor()
        
        if system_type == 'personal':
            c.execute("""SELECT g.id, g.name, g.barcode, ad.ausleih_id
                         FROM geraete g
                         JOIN ausleih_details ad ON g.id = ad.geraet_id
                         JOIN ausleihen a ON ad.ausleih_id = a.ausleih_id
                         WHERE a.rueckgabe_qr = ? AND a.status = 'ausgeliehen'""", (qr_code,))
        else:
            c.execute("""SELECT g.id, g.name, g.barcode, ad.ausleih_id
                         FROM geraete g
                         JOIN ausleih_details ad ON g.id = ad.geraet_id
                         JOIN ausleihen a ON ad.ausleih_id = a.ausleih_id
                         WHERE g.barcode = ? AND a.status = 'ausgeliehen'""", (qr_code,))
        
        devices_to_return = c.fetchall()
        conn.close()
    
    return render_template('return.html', title="Zurückgeben", 
                         devices_to_return=devices_to_return, system_type=system_type)

@app.route('/inventory')
def inventory():
    if 'current_lager' not in session:
        return redirect(url_for('dashboard'))
    
    search = request.args.get('search', '')
    status_filters = [f for f in request.args.getlist('status') if f]
    art_filters = [f for f in request.args.getlist('art') if f]
    klasse_filters = [f for f in request.args.getlist('klasse') if f]
    sort_by = request.args.get('sort_by', 'name')
    group_by = request.args.get('group_by', 'model')  # Added grouping to inventory
    
    conn = get_db_connection(session['current_lager'])
    c = conn.cursor()
    
    # Build query similar to devices but with borrow info
    query = """SELECT g.*, a.mitarbeiter_name, a.zielort, a.datum, a.email, a.klasse
               FROM geraete g
               LEFT JOIN ausleih_details ad ON g.id = ad.geraet_id
               LEFT JOIN ausleihen a ON ad.ausleih_id = a.ausleih_id AND a.status = 'ausgeliehen'
               WHERE 1=1"""
    params = []
    
    if search:
        query += " AND (g.name LIKE ? OR g.barcode LIKE ? OR g.lagerplatz LIKE ? OR g.seriennummer LIKE ? OR g.modell LIKE ? OR g.instrumentenart LIKE ? OR a.mitarbeiter_name LIKE ? OR a.klasse LIKE ?)"
        params.extend([f'%{search}%'] * 8)
    
    if status_filters:
        status_conditions = []
        if "verfügbar" in status_filters:
            status_conditions.append("g.status = 'verfügbar'")
        if "ausgeliehen" in status_filters:
            status_conditions.append("g.status LIKE 'ausgeliehen%'")
        if status_conditions:
            query += f" AND ({' OR '.join(status_conditions)})"
    
    if art_filters:
        placeholders = ','.join('?' for _ in art_filters)
        query += f" AND g.instrumentenart IN ({placeholders})"
        params.extend(art_filters)
    
    if klasse_filters:
        placeholders = ','.join('?' for _ in klasse_filters)
        query += f" AND a.klasse IN ({placeholders})"
        params.extend(klasse_filters)
    
    # Handle sorting
    if sort_by == 'instrumentenart':
        query += " ORDER BY g.instrumentenart, g.name"
    elif sort_by == 'lagerplatz':
        query += " ORDER BY g.lagerplatz, g.name"
    elif sort_by == 'status':
        query += " ORDER BY g.status, g.name"
    elif sort_by == 'model':  # Added model sorting
        query += " ORDER BY g.modell, g.name"
    else:
        query += " ORDER BY g.name"
    
    c.execute(query, params)
    devices_list = c.fetchall()
    
    grouped_devices = {}
    if group_by == 'model':
        for device in devices_list:
            model = device[7] or 'Unbekanntes Modell'
            if model not in grouped_devices:
                grouped_devices[model] = []
            grouped_devices[model].append(device)
    elif group_by == 'instrument':
        for device in devices_list:
            instrument = device[8] or 'Unbekanntes Instrument'
            if instrument not in grouped_devices:
                grouped_devices[instrument] = []
            grouped_devices[instrument].append(device)
    elif group_by == 'status':
        for device in devices_list:
            status = 'Verfügbar' if device[4] == 'verfügbar' else 'Ausgeliehen'
            if status not in grouped_devices:
                grouped_devices[status] = []
            grouped_devices[status].append(device)
    else:
        grouped_devices['Alle Geräte'] = devices_list
    
    # Get filter options
    c.execute("SELECT DISTINCT instrumentenart FROM geraete WHERE instrumentenart IS NOT NULL")
    instrumentenarten = [row[0] for row in c.fetchall()]
    c.execute("SELECT DISTINCT klasse FROM ausleihen WHERE klasse IS NOT NULL")
    klassen = [row[0] for row in c.fetchall()]
    conn.close()
    
    return render_template('inventory.html', title="Inventar", devices=devices_list, grouped_devices=grouped_devices,
                         search=search, status_filters=status_filters,
                         art_filters=art_filters, klasse_filters=klasse_filters,
                         sort_by=sort_by, group_by=group_by, instrumentenarten=instrumentenarten, klassen=klassen)

@app.route('/remove_from_borrow/<int:device_id>')
def remove_from_borrow(device_id):
    if 'borrow_list' in session:
        session['borrow_list'] = [d for d in session['borrow_list'] if d['id'] != device_id]
        session.modified = True
    return redirect(url_for('borrow'))

@app.route('/manage_lager')
def manage_lager():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute("SELECT id, name, created_by, access_users, system_type FROM lager WHERE created_by = ?", (session['user_id'],))
    lagers = c.fetchall()
    conn.close()
    
    return render_template('manage_lager.html', title="Lager verwalten", lagers=lagers)

@app.route('/edit_lager/<lager_id>', methods=['GET', 'POST'])
def edit_lager(lager_id):
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    
    if request.method == 'POST':
        name = request.form['name']
        access_users = request.form.getlist('access_users')
        system_type = request.form.get('system_type', 'personal')
        c.execute("UPDATE lager SET name = ?, access_users = ?, system_type = ? WHERE id = ? AND created_by = ?",
                  (name, ','.join(access_users), system_type, lager_id, session['user_id']))
        conn.commit()
        conn.close()
        return redirect(url_for('manage_lager'))
    
    c.execute("SELECT * FROM lager WHERE id = ? AND created_by = ?", (lager_id, session['user_id']))
    lager = c.fetchone()
    if not lager:
        conn.close()
        return redirect(url_for('manage_lager'))
    
    c.execute("SELECT id, name FROM users WHERE id != ?", (session['user_id'],))
    users = c.fetchall()
    conn.close()
    
    return render_template('edit_lager.html', title="Lager bearbeiten", lager=lager, users=users)

@app.route('/delete_lager/<lager_id>')
def delete_lager(lager_id):
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute("DELETE FROM lager WHERE id = ? AND created_by = ?", (lager_id, session['user_id']))
    conn.commit()
    conn.close()
    
    if os.path.exists(f'{lager_id}.db'):
        os.remove(f'{lager_id}.db')
    
    return redirect(url_for('manage_lager'))

@app.route('/export')
def export():
    if 'current_lager' not in session:
        return redirect(url_for('dashboard'))
    
    format_type = request.args.get('format', 'csv')
    search = request.args.get('search', '')
    status_filters = [f for f in request.args.getlist('status') if f]
    art_filters = [f for f in request.args.getlist('art') if f]
    klasse_filters = [f for f in request.args.getlist('klasse') if f]
    
    conn = get_db_connection(session['current_lager'])
    c = conn.cursor()
    
    # Build query with filters (similar to devices/inventory)
    query = """SELECT g.*, a.mitarbeiter_name, a.zielort, a.datum, a.email, a.klasse
               FROM geraete g
               LEFT JOIN ausleih_details ad ON g.id = ad.geraet_id
               LEFT JOIN ausleihen a ON ad.ausleih_id = a.ausleih_id AND a.status = 'ausgeliehen'
               WHERE 1=1"""
    params = []
    
    if search:
        query += " AND (g.name LIKE ? OR g.barcode LIKE ? OR g.lagerplatz LIKE ? OR g.seriennummer LIKE ? OR g.modell LIKE ? OR g.instrumentenart LIKE ? OR a.mitarbeiter_name LIKE ? OR a.klasse LIKE ?)"
        params.extend([f'%{search}%'] * 8)
    
    if status_filters:
        status_conditions = []
        if "verfügbar" in status_filters:
            status_conditions.append("g.status = 'verfügbar'")
        if "ausgeliehen" in status_filters:
            status_conditions.append("g.status LIKE 'ausgeliehen%'")
        if status_conditions:
            query += f" AND ({' OR '.join(status_conditions)})"
    
    if art_filters:
        placeholders = ','.join('?' for _ in art_filters)
        query += f" AND g.instrumentenart IN ({placeholders})"
        params.extend(art_filters)
    
    if klasse_filters:
        placeholders = ','.join('?' for _ in klasse_filters)
        query += f" AND a.klasse IN ({placeholders})"
        params.extend(klasse_filters)
    
    query += " ORDER BY g.instrumentenart, g.name"
    c.execute(query, params)
    devices_list = c.fetchall()
    conn.close()
    
    if format_type == 'csv':
        output = io.StringIO()
        writer = csv.writer(output)
        writer.writerow(['Name', 'Barcode', 'Lagerplatz', 'Status', 'Beschreibung', 'Seriennummer', 'Modell', 'Instrumentenart', 'Inventar-Nummer', 'Kaufdatum', 'Preis', 'Ausgeliehen an', 'Email', 'Klasse'])
        
        for device in devices_list:
            writer.writerow([device[1], device[2], device[3], device[4], device[5], device[6], device[7], device[8], device[9] or '', device[10] or '', device[11] or '', device[12] or '', device[15] or '', device[16] or ''])
        
        output.seek(0)
        return send_file(io.BytesIO(output.getvalue().encode('utf-8')), 
                        mimetype='text/csv', as_attachment=True, download_name='export.csv')
    
    elif format_type == 'word':
        doc = Document()
        doc.add_heading('Geräte Export', 0)
        
        table = doc.add_table(rows=1, cols=14)
        hdr_cells = table.rows[0].cells
        headers = ['Name', 'Barcode', 'Lagerplatz', 'Status', 'Beschreibung', 'Seriennummer', 'Modell', 'Instrumentenart', 'Inventar-Nummer', 'Kaufdatum', 'Preis', 'Ausgeliehen an', 'Email', 'Klasse']
        
        for i, header in enumerate(headers):
            hdr_cells[i].text = header
        
        for device in devices_list:
            row_cells = table.add_row().cells
            row_cells[0].text = device[1]
            row_cells[1].text = device[2]
            row_cells[2].text = device[3]
            row_cells[3].text = device[4]
            row_cells[4].text = device[5] or ''
            row_cells[5].text = device[6] or ''
            row_cells[6].text = device[7] or ''
            row_cells[7].text = device[8] or ''
            row_cells[8].text = device[9] or ''
            row_cells[9].text = device[10] or ''
            row_cells[10].text = str(device[11] or '') + ' €'
            row_cells[11].text = device[12] or ''
            row_cells[12].text = device[15] or ''
            row_cells[13].text = device[16] or ''
        
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return send_file(buffer, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document', 
                        as_attachment=True, download_name='export.docx')
    
    elif format_type == 'word_labels':
        
        # Get the default label layout for this warehouse
        conn_layout = get_db_connection(session['current_lager'])
        c_layout = conn_layout.cursor()
        c_layout.execute("SELECT layout_data FROM label_layouts WHERE is_default = 1 LIMIT 1")
        layout_result = c_layout.fetchone()
        
        if not layout_result:
            # If no default layout, get the first available layout
            c_layout.execute("SELECT layout_data FROM label_layouts ORDER BY created_at DESC LIMIT 1")
            layout_result = c_layout.fetchone()
        
        conn_layout.close()
        
        if not layout_result:
            # If no layouts exist, return error
            return send_file(io.BytesIO(b'No label layouts found. Please create a label layout first.'), 
                           mimetype='text/plain', as_attachment=True, download_name='error.txt')
        
        layout_data = json.loads(layout_result[0])
        
        doc = Document()
        
        # Set minimal margins
        sections = doc.sections
        for section in sections:
            section.top_margin = Cm(0.5)
            section.bottom_margin = Cm(0.5)
            section.left_margin = Cm(0.5)
            section.right_margin = Cm(0.5)
        
        # Get label dimensions from layout and convert pixels to cm (96 DPI)
        label_width_px = layout_data.get('width', 480)
        label_height_px = layout_data.get('height', 288)
        
        # Convert pixels to cm: 1 inch = 2.54 cm, 96 pixels = 1 inch
        label_width_cm = (label_width_px / 96) * 2.54
        label_height_cm = (label_height_px / 96) * 2.54
        
        # A4 page dimensions in cm minus margins
        page_width_cm = 21 - 1  # A4 width minus margins
        page_height_cm = 29.7 - 1  # A4 height minus margins
        
        # Calculate labels per page
        labels_per_row = max(1, int(page_width_cm / label_width_cm))
        labels_per_col = max(1, int(page_height_cm / label_height_cm))
        labels_per_page = labels_per_row * labels_per_col
        
        def get_field_value(device, field_type, borrower_info=None):
            """Get the actual value for a field type from device data"""
            field_mapping = {
                'device_name': device[1],
                'barcode': device[2],
                'location': device[3],
                'status': device[4],
                'description': device[5] or '',
                'serial_number': device[6] or '',
                'modell': device[7] or '',
                'instrument_type': device[8] or '',
                'inventory_number': device[9] or '',
                'purchase_date': device[10] or '',
                'price': f"{device[11] or 0} €",
                'borrower_name': borrower_info[0] if borrower_info else '',
                'borrower_id': borrower_info[1] if borrower_info else '',
                'email': borrower_info[2] if borrower_info else '',
                'borrow_date': borrower_info[3] if borrower_info else '',
                'destination': borrower_info[4] if borrower_info else '',
                'klasse': borrower_info[5] if borrower_info else ''
            }
            return field_mapping.get(field_type, field_type)
        
        def create_qr_code_image(content, size_cm=2):
            """Generate QR code image and return as BytesIO"""
            qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_L,
                box_size=10,
                border=1,
            )
            qr.add_data(content)
            qr.make(fit=True)
            
            img = qr.make_image(fill_color="black", back_color="white")
            buffer = BytesIO()
            img.save(buffer, format='PNG')
            buffer.seek(0)
            return buffer
        
        # Create labels table
        table = doc.add_table(rows=labels_per_col, cols=labels_per_row)
        table.alignment = WD_TABLE_ALIGNMENT.LEFT
        
        # Set exact cell dimensions and remove borders
        for row in table.rows:
            row.height = Cm(label_height_cm)
            for cell in row.cells:
                cell.width = Cm(label_width_cm)
                # Remove cell margins using correct API
                tc = cell._element
                tcPr = tc.get_or_add_tcPr()
                tcMar = parse_xml(r'<w:tcMar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                                 r'<w:top w:w="0" w:type="dxa"/>'
                                 r'<w:left w:w="0" w:type="dxa"/>'
                                 r'<w:bottom w:w="0" w:type="dxa"/>'
                                 r'<w:right w:w="0" w:type="dxa"/></w:tcMar>')
                tcPr.append(tcMar)
        
        # Remove table borders
        tbl = table._element
        tblPr = tbl.tblPr
        if tblPr is None:
            tblPr = parse_xml(r'<w:tblPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
            tbl.insert(0, tblPr)
        
        tblBorders = parse_xml(r'<w:tblBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                              r'<w:top w:val="none"/><w:left w:val="none"/>'
                              r'<w:bottom w:val="none"/><w:right w:val="none"/>'
                              r'<w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders>')
        tblPr.append(tblBorders)
        
        label_index = 0
        for device in devices_list:
            if label_index >= labels_per_page:
                # Add new page
                doc.add_page_break()
                table = doc.add_table(rows=labels_per_col, cols=labels_per_row)
                table.alignment = WD_TABLE_ALIGNMENT.LEFT
                
                # Set dimensions and formatting for new table
                for row in table.rows:
                    row.height = Cm(label_height_cm)
                    for cell in row.cells:
                        cell.width = Cm(label_width_cm)
                        tc = cell._element
                        tcPr = tc.get_or_add_tcPr()
                        tcMar = parse_xml(r'<w:tcMar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                                         r'<w:top w:w="0" w:type="dxa"/>'
                                         r'<w:left w:w="0" w:type="dxa"/>'
                                         r'<w:bottom w:w="0" w:type="dxa"/>'
                                         r'<w:right w:w="0" w:type="dxa"/></w:tcMar>')
                        tcPr.append(tcMar)
                
                # Remove borders from new table
                tbl = table._element
                tblPr = tbl.tblPr
                if tblPr is None:
                    tblPr = parse_xml(r'<w:tblPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
                    tbl.insert(0, tblPr)
                
                tblBorders = parse_xml(r'<w:tblBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                                      r'<w:top w:val="none"/><w:left w:val="none"/>'
                                      r'<w:bottom w:val="none"/><w:right w:val="none"/>'
                                      r'<w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders>')
                tblPr.append(tblBorders)
                
                label_index = 0
            
            row_idx = label_index // labels_per_row
            col_idx = label_index % labels_per_row
            cell = table.rows[row_idx].cells[col_idx]
            
            # Clear cell content
            cell._element.clear_content()
            
            # Get borrower info if device is borrowed
            borrower_info = None
            if device[12]:  # mitarbeiter_name exists
                borrower_info = (device[12], '', device[15] or '', device[14] or '', device[13] or '', device[16] or '')
            
            # Process fields from layout, sorted by Y position for proper layering
            fields = layout_data.get('fields', [])
            sorted_fields = sorted(fields, key=lambda f: f.get('y', 0))
            
            for field in sorted_fields:
                field_type = field.get('type', 'text')
                field_x = field.get('x', 0)
                field_y = field.get('y', 0)
                field_width = field.get('width', 100)
                field_height = field.get('height', 20)
                
                if field_type == 'qr_code':
                    # Generate QR code with warehouse ID + device barcode
                    qr_content = f"{session['current_lager']}-{device[2]}"
                    qr_size_cm = min((field_width / 96) * 2.54, 3)  # Convert px to cm, max 3cm
                    
                    qr_buffer = create_qr_code_image(qr_content, qr_size_cm)
                    
                    # Add QR code to document
                    qr_para = cell.add_paragraph()
                    qr_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    
                    # Add spacing based on Y position
                    if field_y > 0:
                        qr_para.space_before = Pt(field_y * 0.75)  # Convert pixels to points
                    
                    qr_run = qr_para.add_run()
                    qr_run.add_picture(qr_buffer, width=Cm(qr_size_cm))
                    
                else:
                    # Handle text fields
                    field_value = get_field_value(device, field_type, borrower_info)
                    if field_value:
                        text_para = cell.add_paragraph()
                        text_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        
                        # Add spacing based on Y position
                        if field_y > 0:
                            text_para.space_before = Pt(field_y * 0.75)  # Convert pixels to points
                        
                        text_run = text_para.add_run(str(field_value))
                        
                        # Parse font size properly
                        font_size_raw = field.get('fontSize', '8px')
                        font_size = 8  # default
                        
                        # Extract numeric value from font size
                        if isinstance(font_size_raw, str):
                            match = re.search(r'(\d+)', font_size_raw)
                            if match:
                                font_size = int(match.group(1))
                        elif isinstance(font_size_raw, (int, float)):
                            font_size = int(font_size_raw)
                        
                        # Apply font size (clamp between 6-24pt for labels)
                        text_run.font.size = Pt(max(6, min(font_size, 24)))
                        
                        # Apply bold if specified
                        if field.get('fontWeight') == 'bold':
                            text_run.bold = True
            
            label_index += 1
        
        # Save document
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return send_file(buffer, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document', 
                        as_attachment=True, download_name='custom_labels.docx')
    
    return send_file(io.BytesIO(b'Not implemented'), mimetype='text/plain')

@app.route('/label-layout')
def label_layout():
    if 'current_lager' not in session:
        return redirect(url_for('dashboard'))
    
    conn = get_db_connection(session['current_lager'])
    c = conn.cursor()
    c.execute("SELECT id, name, is_default, created_at FROM label_layouts ORDER BY is_default DESC, name")
    labels = c.fetchall()
    conn.close()
    
    return render_template('label_selection.html', title="Label Editor", labels=labels)

@app.route('/label-layout/edit/<int:label_id>')
def edit_label_layout(label_id):
    if 'current_lager' not in session:
        return redirect(url_for('dashboard'))
    
    conn = get_db_connection(session['current_lager'])
    c = conn.cursor()
    c.execute("SELECT id, name, layout_data FROM label_layouts WHERE id = ?", (label_id,))
    label = c.fetchone()
    conn.close()
    
    if not label:
        return redirect(url_for('label_layout'))
    
    return render_template('label_layout.html', title=f"Label Editor - {label[1]}", 
                         label_id=label[0], label_name=label[1], layout_data=label[2])

@app.route('/label-layout/new')
def new_label_layout():
    if 'current_lager' not in session:
        return redirect(url_for('dashboard'))
    
    return render_template('label_layout.html', title="Neues Label erstellen", 
                         label_id=None, label_name="", layout_data="{}")

@app.route('/save-layout', methods=['POST'])
def save_layout():
    if 'current_lager' not in session:
        return jsonify({'error': 'No warehouse selected'}), 400
    
    data = request.json
    label_id = data.get('label_id')
    label_name = data.get('name', 'Unbenannt')
    layout_data = json.dumps(data.get('layout', {}))
    
    conn = get_db_connection(session['current_lager'])
    c = conn.cursor()
    
    if label_id:
        c.execute("UPDATE label_layouts SET name = ?, layout_data = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
                  (label_name, layout_data, label_id))
    else:
        c.execute("INSERT INTO label_layouts (name, layout_data) VALUES (?, ?)",
                  (label_name, layout_data))
        label_id = c.lastrowid
    
    conn.commit()
    conn.close()
    
    return jsonify({'success': True, 'message': 'Layout gespeichert', 'label_id': label_id})

@app.route('/set-default-label/<int:label_id>', methods=['POST'])
def set_default_label(label_id):
    if 'current_lager' not in session:
        return jsonify({'error': 'No warehouse selected'}), 400
    
    conn = get_db_connection(session['current_lager'])
    c = conn.cursor()
    
    c.execute("UPDATE label_layouts SET is_default = 0")
    c.execute("UPDATE label_layouts SET is_default = 1 WHERE id = ?", (label_id,))
    
    conn.commit()
    conn.close()
    
    return jsonify({'success': True, 'message': 'Standard-Label gesetzt'})

@app.route('/delete-label/<int:label_id>', methods=['POST'])
def delete_label(label_id):
    if 'current_lager' not in session:
        return jsonify({'error': 'No warehouse selected'}), 400
    
    conn = get_db_connection(session['current_lager'])
    c = conn.cursor()
    c.execute("DELETE FROM label_layouts WHERE id = ?", (label_id,))
    conn.commit()
    conn.close()
    
    return jsonify({'success': True, 'message': 'Label gelöscht'})

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

if __name__ == '__main__':
    init_user_db()
    app.run(debug=True, host='0.0.0.0', port=5000)
