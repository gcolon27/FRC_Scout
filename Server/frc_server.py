"""
FRC 2026 MASTER SERVER
1. Ensure 'html5-qrcode.min.js' is in this folder.
2. Install dependencies: pip install flask pandas odfpy
3. Run this script (works on Mac & Linux).
4. Access http://localhost:8000
"""
import sys
import os

# ==========================================================
# PRE-FLIGHT CHECK 1: Python Dependencies
# ==========================================================
missing_modules = []
try:
    import flask
except ImportError:
    missing_modules.append('flask')

try:
    import pandas
except ImportError:
    missing_modules.append('pandas')

try:
    import odf
except ImportError:
    missing_modules.append('odfpy')

if missing_modules:
    print("\n" + "!"*55)
    print("🛑 SERVER STARTUP FAILED: Missing Python Packages")
    print("!"*55)
    print("Please install the required dependencies by running this command in your terminal:\n")
    print(f"👉  pip install {' '.join(missing_modules)}\n")
    sys.exit(1)

import sqlite3
import json
import logging
import pandas as pd
from io import BytesIO
from datetime import datetime, timezone
from flask import Flask, render_template_string, request, jsonify, Response, send_file
import flask.cli

app = Flask(__name__)
DB_FILE = 'frc_scouting.db'

# --- SILENCE FLASK LOGGING ---
log = logging.getLogger('werkzeug')
log.setLevel(logging.ERROR)
flask.cli.show_server_banner = lambda *args: None

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS matches (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            match_number INTEGER,
            team_number INTEGER,
            scout_name TEXT,
            device_id TEXT,
            data_json TEXT,
            scan_time TEXT,
            timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(match_number, team_number)
        )
    ''')
    
    try:
        c.execute("ALTER TABLE matches ADD COLUMN device_id TEXT")
    except sqlite3.OperationalError:
        pass 
        
    conn.commit()
    conn.close()

def parse_match_data(row):
    db_id, raw_json, scan_time, device_id = row
    try:
        d = json.loads(raw_json)
    except:
        d = {}

    return {
        'id': db_id,
        'scan_time_utc': scan_time if scan_time else d.get('scan_time_utc', 'N/A'),
        'deviceId': device_id if device_id else d.get('dev', d.get('deviceId', 'Legacy')),
        'matchNumber': d.get('m', d.get('matchNumber', 0)),
        'teamNumber': d.get('t', d.get('teamNumber', 0)),
        'scoutName': d.get('s', d.get('scoutName', 'Unknown')),
        'autoBalls': d.get('ab', d.get('autoBalls', 0)),
        'autoClimb': d.get('ac', d.get('autoClimb', 'None')),
        'teleBalls': d.get('tb', d.get('teleBalls', 0)),
        'endClimb': d.get('ec', d.get('endClimb', 'None')),
        'outcome': d.get('o', d.get('outcome', 'Tie')),
        'defense': 'Yes' if d.get('df', d.get('defense')) in [1, True, 'Yes'] else 'No',
        'broken': 'Yes' if d.get('br', d.get('broken')) in [1, True, 'Yes'] else 'No',
        'notes': d.get('n', d.get('notes', ''))
    }

# --- SCANNER INTERFACE (CATPPUCCIN MOCHA) ---
SCANNER_HTML = """
<!DOCTYPE html>
<html>
<head>
    <title>FRC Scanner</title>
    <script src="/html5-qrcode.min.js" type="text/javascript"></script>
    <style>
        body { background: #1e1e2e; color: #cdd6f4; font-family: sans-serif; text-align: center; padding: 10px; }
        
        .nav { margin-bottom: 30px; display: flex; justify-content: center; gap: 20px; }
        .nav a { color: #89b4fa; text-decoration: none; font-size: 1.2rem; border-bottom: 2px solid #89b4fa; padding-bottom: 3px; transition: 0.3s; }
        .nav a:hover { color: #f9e2af; border-color: #f9e2af; }
        
        #reader { width: 100%; max-width: 500px; margin: 0 auto; border: 2px solid #89b4fa; background: #181825; border-radius: 10px; overflow: hidden; padding-bottom: 15px;}
        
        /* Styled the native 'Scan an Image File' link so it looks good in dark mode */
        #reader a { color: #89b4fa; text-decoration: none; font-weight: bold; font-size: 1.1rem; }
        #reader a:hover { color: #f9e2af; }
        
        .success-box { background: #181825; border: 2px solid #a6e3a1; padding: 20px; border-radius: 10px; margin-top: 20px; display: none; }
        .btn { padding: 15px 30px; font-size: 1.2rem; border: 2px solid #89b4fa; border-radius: 6px; cursor: pointer; margin-top: 10px; width: 100%; max-width: 300px; font-weight: bold; transition: 0.3s; }
        .btn-primary { background: #313244; color: #89b4fa; }
        .btn-primary:hover { background: #89b4fa; color: #1e1e2e; }
        
        /* Mode Selector Styles */
        .mode-btn { background: #181825; color: #cdd6f4; border: 2px solid #313244; padding: 10px 15px; font-size: 1rem; cursor: pointer; border-radius: 6px; margin: 5px; transition: 0.3s; }
        .mode-btn.active { background: #89b4fa; color: #1e1e2e; border-color: #89b4fa; font-weight: bold; }
        .mode-btn:hover { border-color: #89b4fa; }
        
        #usb-input { width: 100%; max-width: 500px; padding: 15px; font-size: 1.2rem; background: #313244; color: #cdd6f4; border: 2px solid #89b4fa; border-radius: 8px; box-sizing: border-box; outline: none; }
        #usb-input:focus { box-shadow: 0 0 8px #89b4fa; }
    </style>
</head>
<body>
    <div class="nav">
        <a href="/">📷 Scanner</a>
        <a href="/view">📊 View Data</a>
    </div>

    <div id="mode-selector" style="margin-bottom: 20px;">
        <button id="btn-camera" class="mode-btn active" onclick="setMode('camera')">📷 Camera / File</button>
        <button id="btn-usb" class="mode-btn" onclick="setMode('usb')">⌨️ USB Scanner</button>
    </div>

    <div id="reader"></div>

    <div id="usb-container" style="display: none; margin-bottom: 20px;">
        <p style="color: #89b4fa; font-weight: bold; margin-bottom: 10px;">Click the box below and scan your code</p>
        <input type="text" id="usb-input" placeholder="Awaiting scan..." autocomplete="off">
    </div>

    <div id="error-msg" style="color: #f38ba8; display:none;">
        <h3>⚠️ MISSING FILE</h3>
        <p>Please ensure <b>html5-qrcode.min.js</b> is in the server folder.</p>
    </div>

    <div id="successDisplay" class="success-box">
        <h2 style="margin:0 0 10px 0; color: #a6e3a1;">✅ Scan Successful!</h2>
        <div id="scanDetails" style="font-size: 1.2rem; margin-bottom: 15px; color: #cdd6f4;"></div>
        <button class="btn btn-primary" onclick="submitToDb()">💾 Save to Database</button>
    </div>

    <script>
        let currentData = null;
        let scanner = null;
        let currentMode = localStorage.getItem('scannerMode') || 'camera';

        function setMode(mode) {
            currentMode = mode;
            localStorage.setItem('scannerMode', mode);
            
            document.getElementById('btn-camera').classList.remove('active');
            document.getElementById('btn-usb').classList.remove('active');
            
            document.getElementById('reader').style.display = 'none';
            document.getElementById('usb-container').style.display = 'none';
            
            if (mode === 'camera') {
                document.getElementById('reader').style.display = 'block';
                document.getElementById('btn-camera').classList.add('active');
            } else if (mode === 'usb') {
                document.getElementById('usb-container').style.display = 'block';
                document.getElementById('btn-usb').classList.add('active');
                document.getElementById('usb-input').value = '';
                document.getElementById('usb-input').focus();
            }
        }

        function handleValidData(raw) {
            currentData = {
                deviceId: raw.dev || 'Unknown',
                matchNumber: raw.m,
                teamNumber: raw.t,
                scoutName: raw.s || 'Unknown',
                autoBalls: raw.ab || 0,
                autoClimb: raw.ac || 'None',
                teleBalls: raw.tb || 0,
                endClimb: raw.ec || 'None',
                outcome: raw.o || 'Tie',
                defense: raw.df === 1 ? 'Yes' : 'No',
                broken: raw.br === 1 ? 'Yes' : 'No',
                notes: raw.n || ''
            };

            document.getElementById('scanDetails').innerHTML =
                `Match <strong>${currentData.matchNumber}</strong> - Team <strong>${currentData.teamNumber}</strong><br>` +
                `Scout: ${currentData.scoutName}`;

            document.getElementById('successDisplay').style.display = 'block';
            document.getElementById('reader').style.display = 'none';
            document.getElementById('usb-container').style.display = 'none';
            document.getElementById('mode-selector').style.display = 'none';
            
            try {
                if (scanner && scanner.getState() === 2) {
                    scanner.pause();
                }
            } catch(e) { }
        }

        function processCameraScan(decodedText) {
            try {
                let raw = JSON.parse(decodedText);
                handleValidData(raw);
            } catch(e) { 
                // Silently ignore invalid parses from camera.
            }
        }

        document.getElementById('usb-input').addEventListener('keydown', function(e) {
            if (e.key === 'Enter') {
                e.preventDefault();
                let val = this.value.trim();
                if(val !== '') {
                    let raw = null;
                    try {
                        raw = JSON.parse(val);
                    } catch(err) {
                        alert("❌ Invalid QR Data received from USB Scanner.");
                        this.value = ''; 
                        return;
                    }
                    this.value = ''; 
                    handleValidData(raw); 
                }
            }
        });

        // Initialize Native Scanner (This re-enables the built-in file uploader!)
        try {
            scanner = new Html5QrcodeScanner("reader", { 
                fps: 10, 
                qrbox: 250
            });
            scanner.render(processCameraScan);
        } catch(e) {
            document.getElementById('error-msg').style.display = 'block';
            document.getElementById('reader').style.display = 'none';
        }

        async function submitToDb() {
            if(!currentData) return;
            try {
                const res = await fetch('/api/submit', {
                    method: 'POST', headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify(currentData)
                });
                if(res.ok) {
                    location.reload(); 
                } else {
                    alert("❌ Error saving");
                }
            } catch(e) { alert("❌ Network Error"); }
        }

        // Apply saved mode on load
        setTimeout(() => setMode(currentMode), 100);
    </script>
</body>
</html>
"""

# --- DASHBOARD HTML (CATPPUCCIN MOCHA) ---
DASHBOARD_HTML = """
<!DOCTYPE html>
<html>
<head>
    <title>FRC Data</title>
    <style>
        :root {
            --base: #1e1e2e; --mantle: #181825; --text: #cdd6f4;
            --surface0: #313244; --blue: #89b4fa; --yellow: #f9e2af;
            --red: #f38ba8; --green: #a6e3a1; --subtext0: #a6adc8;
        }
        body { background: var(--base); color: var(--text); font-family: sans-serif; padding: 20px; }
        
        .nav { margin-bottom: 30px; display: flex; justify-content: center; gap: 20px; }
        .nav a { color: var(--blue); text-decoration: none; font-size: 1.2rem; border-bottom: 2px solid var(--blue); padding-bottom: 3px; transition: 0.3s; }
        .nav a:hover { color: var(--yellow); border-color: var(--yellow); }
        
        table { width: 100%; border-collapse: collapse; background: var(--mantle); border: 1px solid var(--surface0); }
        th, td { padding: 10px; text-align: center; border-bottom: 1px solid var(--surface0); }
        th { background: var(--surface0); color: var(--blue); border-bottom: 2px solid var(--blue); }
        
        .btn-ods { background: var(--surface0); color: var(--blue); border: 2px solid var(--blue); padding: 10px 15px; text-decoration: none; border-radius: 5px; display:inline-block; margin-bottom:20px; font-weight: bold; transition: 0.3s; }
        .btn-ods:hover { background: var(--blue); color: var(--base); }
        
        .btn-util { background: transparent; color: var(--text); border: 1px solid var(--text); padding: 5px 10px; cursor: pointer; border-radius: 4px; font-weight: bold; transition: 0.3s;}
        .btn-util:hover { background: var(--text); color: var(--base); }
        .btn-edit { background: transparent; color: var(--yellow); border: 1px solid var(--yellow); padding: 5px 10px; cursor: pointer; border-radius: 4px; font-weight: bold; transition: 0.3s;}
        .btn-edit:hover { background: var(--yellow); color: var(--base); }
        .btn-del { background: transparent; color: var(--red); border: 1px solid var(--red); padding: 5px 10px; cursor: pointer; border-radius: 4px; font-weight: bold; transition: 0.3s;}
        .btn-del:hover { background: var(--red); color: var(--base);}
        
        /* Modals */
        .modal { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background-color: rgba(17,17,27,0.85); justify-content: center; align-items: center; z-index: 1000; }
        .modal-content { background-color: var(--mantle); border: 3px solid var(--blue); border-radius: 12px; padding: 25px; width: 90%; max-width: 500px; max-height: 90vh; overflow-y: auto;}
        
        .modal input, .modal select, .modal textarea { width: 100%; background: var(--surface0); color: var(--text); border: 1px solid var(--subtext0); padding: 8px; border-radius: 6px; box-sizing: border-box; margin-bottom: 15px;}
        .modal label { font-weight: bold; color: var(--blue); display: block; margin-bottom: 4px; text-align: left;}
        
        .flex-row { display: flex; gap: 10px; }
        .flex-row > div { flex: 1; }
    </style>
</head>
<body>
    <div class="nav">
        <a href="/">📷 Scanner</a>
        <a href="/view">📊 View Data</a>
    </div>

    <a href="/api/export_ods" class="btn-ods">📥 Download ODS (Spreadsheet)</a>

    <table>
        <thead>
            <tr>
                <th>Scan Time (UTC)</th>
                <th>Device</th>
                <th>Match</th>
                <th>Team</th>
                <th>Scout</th>
                <th>Auto</th>
                <th>A-Climb</th>
                <th>Tele</th>
                <th>E-Climb</th>
                <th>Defense</th>
                <th>Broken</th>
                <th>Notes</th>
                <th>Action</th>
            </tr>
        </thead>
        <tbody>
            {% for r in matches %}
            <tr>
                <td style="font-size: 0.9em; color: var(--subtext0);">{{ r.scan_time_utc }}</td>
                <td style="color: var(--subtext0);">{{ r.deviceId }}</td>
                <td style="font-weight: bold; color: var(--yellow);">{{ r.matchNumber }}</td>
                <td style="font-weight: bold; color: var(--blue);">{{ r.teamNumber }}</td>
                <td>{{ r.scoutName }}</td>
                <td>{{ r.autoBalls }}</td>
                <td>{{ r.autoClimb }}</td>
                <td>{{ r.teleBalls }}</td>
                <td>{{ r.endClimb }}</td>
                <td style="color: {% if r.defense == 'Yes' %}var(--green){% else %}var(--subtext0){% endif %};">{{ r.defense }}</td>
                <td style="color: {% if r.broken == 'Yes' %}var(--red){% else %}var(--subtext0){% endif %};">{{ r.broken }}</td>
                <td>
                    <button class="btn-util" onclick="openNotes({{ r.id }})" {% if not r.notes %}style="opacity:0.3; pointer-events:none;"{% endif %}>📄</button>
                </td>
                <td style="white-space: nowrap;">
                    <button class="btn-edit" onclick="openEdit({{ r.id }})">✏️</button>
                    <button class="btn-del" onclick="del({{ r.id }})">X</button>
                </td>
            </tr>
            {% endfor %}
        </tbody>
    </table>

    <script>
        const matchData = {{ matches | tojson }};
    </script>

    <div id="notesModal" class="modal" onclick="closeModals()">
        <div class="modal-content" onclick="event.stopPropagation()">
            <h2 style="color: var(--yellow); margin-top:0;">Scout Notes</h2>
            <p id="displayNotes" style="font-size: 1.1rem; line-height: 1.5; background: var(--surface0); padding: 15px; border-radius: 8px;"></p>
            <button class="btn-util" style="width: 100%; margin-top: 15px; padding: 10px;" onclick="closeModals()">Close</button>
        </div>
    </div>

    <div id="editModal" class="modal" onclick="closeModals()">
        <div class="modal-content" onclick="event.stopPropagation()">
            <h2 style="color: var(--blue); margin-top:0; text-align: center;">Edit Match Data</h2>
            <input type="hidden" id="e_id">
            
            <div class="flex-row">
                <div><label>Match</label><input type="number" id="e_match"></div>
                <div><label>Team</label><input type="number" id="e_team"></div>
            </div>
            <label>Scout Name</label><input type="text" id="e_scout">
            
            <div class="flex-row">
                <div><label>Auto Balls</label><input type="number" id="e_auto"></div>
                <div>
                    <label>Auto Climb</label>
                    <select id="e_aClimb">
                        <option>None</option><option>L1</option><option>L2</option><option>L3</option>
                    </select>
                </div>
            </div>

            <div class="flex-row">
                <div><label>Tele Balls</label><input type="number" id="e_tele"></div>
                <div>
                    <label>End Climb</label>
                    <select id="e_eClimb">
                        <option>None</option><option>L1</option><option>L2</option><option>L3</option>
                    </select>
                </div>
            </div>

            <div class="flex-row">
                <div>
                    <label>Outcome</label>
                    <select id="e_outcome">
                        <option>Tie</option><option>Win</option><option>Loss</option>
                    </select>
                </div>
                <div>
                    <label>Defense</label>
                    <select id="e_def"><option>No</option><option>Yes</option></select>
                </div>
                <div>
                    <label>Broken</label>
                    <select id="e_broke"><option>No</option><option>Yes</option></select>
                </div>
            </div>

            <label>Notes</label><textarea id="e_notes" rows="3"></textarea>

            <div class="flex-row" style="margin-top: 10px;">
                <button class="btn-edit" style="width: 100%; padding: 10px;" onclick="saveEdit()">💾 Save Changes</button>
                <button class="btn-util" style="width: 100%; padding: 10px;" onclick="closeModals()">Cancel</button>
            </div>
        </div>
    </div>

    <script>
        function closeModals() {
            document.getElementById('notesModal').style.display = 'none';
            document.getElementById('editModal').style.display = 'none';
        }

        function openNotes(id) {
            const m = matchData.find(x => x.id === id);
            if(m) {
                document.getElementById('displayNotes').innerText = m.notes || "No notes provided.";
                document.getElementById('notesModal').style.display = 'flex';
            }
        }

        function openEdit(id) {
            const m = matchData.find(x => x.id === id);
            if(m) {
                document.getElementById('e_id').value = m.id;
                document.getElementById('e_match').value = m.matchNumber;
                document.getElementById('e_team').value = m.teamNumber;
                document.getElementById('e_scout').value = m.scoutName;
                document.getElementById('e_auto').value = m.autoBalls;
                document.getElementById('e_aClimb').value = m.autoClimb;
                document.getElementById('e_tele').value = m.teleBalls;
                document.getElementById('e_eClimb').value = m.endClimb;
                document.getElementById('e_outcome').value = m.outcome;
                document.getElementById('e_def').value = m.defense;
                document.getElementById('e_broke').value = m.broken;
                document.getElementById('e_notes').value = m.notes;
                
                document.getElementById('editModal').style.display = 'flex';
            }
        }

        async function saveEdit() {
            const id = document.getElementById('e_id').value;
            const payload = {
                matchNumber: document.getElementById('e_match').value,
                teamNumber: document.getElementById('e_team').value,
                scoutName: document.getElementById('e_scout').value,
                autoBalls: document.getElementById('e_auto').value,
                autoClimb: document.getElementById('e_aClimb').value,
                teleBalls: document.getElementById('e_tele').value,
                endClimb: document.getElementById('e_eClimb').value,
                outcome: document.getElementById('e_outcome').value,
                defense: document.getElementById('e_def').value,
                broken: document.getElementById('e_broke').value,
                notes: document.getElementById('e_notes').value
            };

            try {
                const res = await fetch('/api/edit/' + id, {
                    method: 'POST', 
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify(payload)
                });
                
                if(res.ok) {
                    location.reload();
                } else {
                    const errorMsg = await res.json();
                    alert("❌ Error saving: " + errorMsg.error);
                }
            } catch(e) { alert("❌ Network Error"); }
        }

        async function del(id) {
            if(confirm("Delete match data permanently?")) {
                await fetch('/api/delete/' + id, { method: 'POST' });
                location.reload();
            }
        }
    </script>
</body>
</html>
"""

# --- ROUTES ---
@app.route('/')
def home():
    return render_template_string(SCANNER_HTML)

@app.route('/html5-qrcode.min.js')
def serve_js():
    try:
        return send_file('html5-qrcode.min.js', mimetype='application/javascript')
    except FileNotFoundError:
        return "Error: html5-qrcode.min.js not found in server folder!", 404

@app.route('/view')
def view_data():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT id, data_json, scan_time, device_id FROM matches ORDER BY match_number DESC")
    rows = c.fetchall()
    conn.close()

    clean_data = [parse_match_data(r) for r in rows]
    return render_template_string(DASHBOARD_HTML, matches=clean_data)

@app.route('/api/submit', methods=['POST'])
def submit():
    try:
        data = request.json
        utc_now = datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S')
        data['scan_time_utc'] = utc_now

        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute('''
            INSERT OR REPLACE INTO matches
            (match_number, team_number, scout_name, device_id, data_json, scan_time)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (
            data['matchNumber'], 
            data['teamNumber'], 
            data['scoutName'],
            data.get('deviceId', 'Unknown'),
            json.dumps(data),
            utc_now
        ))
        conn.commit()
        conn.close()
        
        scout_name = data.get('scoutName', 'Unknown')
        device_id = data.get('deviceId', 'Unknown')
        print(f"✅ {scout_name} submitted a scan from {device_id} at {utc_now}")
        
        return jsonify({'success': True, 'timestamp': utc_now})
    except Exception as e:
        print(f"❌ SERVER ERROR: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/edit/<int:match_id>', methods=['POST'])
def edit_match(match_id):
    try:
        data = request.json
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        
        c.execute("SELECT data_json, scan_time, device_id FROM matches WHERE id = ?", (match_id,))
        row = c.fetchone()
        if not row:
            return jsonify({'error': 'Match not found'}), 404
            
        raw_json, scan_time, device_id = row
        existing = json.loads(raw_json) if raw_json else {}
        
        existing['m'] = int(data['matchNumber'])
        existing['t'] = int(data['teamNumber'])
        existing['s'] = data['scoutName']
        existing['ab'] = int(data['autoBalls'])
        existing['ac'] = data['autoClimb']
        existing['tb'] = int(data['teleBalls'])
        existing['ec'] = data['endClimb']
        existing['o'] = data['outcome']
        existing['df'] = 1 if data['defense'] == 'Yes' else 0
        existing['br'] = 1 if data['broken'] == 'Yes' else 0
        existing['n'] = data['notes']
        
        if 'scan_time_utc' not in existing: existing['scan_time_utc'] = scan_time
        if 'dev' not in existing: existing['dev'] = device_id
        
        c.execute('''
            UPDATE matches
            SET match_number = ?, team_number = ?, scout_name = ?, data_json = ?
            WHERE id = ?
        ''', (existing['m'], existing['t'], existing['s'], json.dumps(existing), match_id))
        
        conn.commit()
        conn.close()
        print(f"✏️  Match {existing['m']} (Team {existing['t']}) was manually edited.")
        return jsonify({'success': True})
    except sqlite3.IntegrityError:
         return jsonify({'error': 'Match and Team combination already exists.'}), 400
    except Exception as e:
         return jsonify({'error': str(e)}), 500

@app.route('/api/delete/<int:match_id>', methods=['POST'])
def delete_match(match_id):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("DELETE FROM matches WHERE id = ?", (match_id,))
    conn.commit()
    conn.close()
    return jsonify({'success': True})

@app.route('/api/export_ods')
def export_ods():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT id, data_json, scan_time, device_id FROM matches ORDER BY match_number")
    rows = c.fetchall()
    conn.close()

    parsed_data = []
    for r in rows:
        d = parse_match_data(r)
        parsed_data.append({
            'Scan Time (UTC)': d['scan_time_utc'],
            'Device ID': d['deviceId'],
            'Match': d['matchNumber'],
            'Team': d['teamNumber'],
            'Scout': d['scoutName'],
            'Auto Balls': d['autoBalls'],
            'Auto Climb': d['autoClimb'],
            'Teleop Balls': d['teleBalls'],
            'Endgame Climb': d['endClimb'],
            'Outcome': d['outcome'],
            'Defense': d['defense'],
            'Robot Broke': d['broken'],
            'Notes': d['notes']
        })

    df = pd.DataFrame(parsed_data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='odf') as writer:
        df.to_excel(writer, index=False)
    
    output.seek(0)
    return Response(
        output.getvalue(),
        mimetype="application/vnd.oasis.opendocument.spreadsheet",
        headers={"Content-disposition": "attachment; filename=frc_data.ods"}
    )

if __name__ == '__main__':
    if not os.path.exists('html5-qrcode.min.js'):
        print("\n" + "!"*55)
        print("🛑 SERVER STARTUP FAILED: Missing Scanner File")
        print("!"*55)
        print("The scanner requires 'html5-qrcode.min.js' to function.")
        print("👉 Please download it and place it in the SAME folder as this script.\n")
        sys.exit(1)

    init_db()
    
    print("\n" + "="*50)
    print("🔗 Scanner Interface: http://localhost:8000")
    print("🛑 Press CTRL+C to quit")
    print("="*50 + "\n")
    
    app.run(host='0.0.0.0', port=8000, debug=False)
