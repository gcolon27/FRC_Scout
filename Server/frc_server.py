"""
FRC 2026 MASTER SERVER
1. Ensure 'html5-qrcode.min.js' is in this folder.
2. Install dependencies: pip install flask pandas odfpy
3. Run this script.
4. Access http://localhost:5000
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
    import odf # This is the internal name for odfpy
except ImportError:
    missing_modules.append('odfpy')

if missing_modules:
    print("\n" + "!"*55)
    print("🛑 SERVER STARTUP FAILED: Missing Python Packages")
    print("!"*55)
    print("Please install the required dependencies by running this command in your terminal:\n")
    print(f"👉  pip install {' '.join(missing_modules)}\n")
    sys.exit(1)

# --- Safe to import everything else now ---
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
            data_json TEXT,
            scan_time TEXT,
            timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(match_number, team_number)
        )
    ''')
    conn.commit()
    conn.close()

# --- SCANNER INTERFACE (OFFLINE) ---
SCANNER_HTML = """
<!DOCTYPE html>
<html>
<head>
    <title>FRC Scanner</title>
    <script src="/html5-qrcode.min.js" type="text/javascript"></script>
    <style>
        body { background: #1a1a1a; color: white; font-family: sans-serif; text-align: center; padding: 10px; }
        .nav { margin-bottom: 20px; }
        .nav a { color: #00ccff; margin: 0 10px; text-decoration: none; font-size: 1.2rem; border-bottom: 1px solid #00ccff; }
        #reader { width: 100%; max-width: 500px; margin: 0 auto; border: 4px solid #333; border-radius: 10px; }
        .success-box { background: #004d26; border: 2px solid #00d968; padding: 20px; border-radius: 10px; margin-top: 20px; display: none; }
        .btn { padding: 15px 30px; font-size: 1.2rem; border: none; border-radius: 6px; cursor: pointer; margin-top: 10px; width: 100%; max-width: 300px; }
        .btn-primary { background: #0066cc; color: white; }
    </style>
</head>
<body>
    <div class="nav">
        <a href="/">📷 Scanner</a>
        <a href="/view">📊 View Data</a>
    </div>

    <div id="reader"></div>
    <div id="error-msg" style="color: red; display:none;">
        <h3>⚠️ MISSING FILE</h3>
        <p>Please ensure <b>html5-qrcode.min.js</b> is in the server folder.</p>
    </div>

    <div id="successDisplay" class="success-box">
        <h2 style="margin:0 0 10px 0;">✅ Scan Successful!</h2>
        <div id="scanDetails" style="font-size: 1.2rem; margin-bottom: 15px; color: #ccffdd;"></div>
        <button class="btn btn-primary" onclick="submitToDb()">💾 Save to Database</button>
    </div>

    <script>
        let currentData = null;
        try {
            let scanner = new Html5QrcodeScanner("reader", { fps: 10, qrbox: 250 });

            function onScan(decodedText) {
                try {
                    let raw = JSON.parse(decodedText);
                    currentData = {
                        matchNumber: raw.m,
                        teamNumber: raw.t,
                        scoutName: raw.s || 'Unknown',
                        autoBalls: raw.ab || 0,
                        autoClimb: raw.ac || 'None',
                        teleBalls: raw.tb || 0,
                        endClimb: raw.ec || 'None',
                        outcome: raw.o || 'None',
                        defense: raw.df === 1 ? 'Yes' : 'No',
                        broken: raw.br === 1 ? 'Yes' : 'No',
                        notes: raw.n || ''
                    };

                    document.getElementById('scanDetails').innerHTML =
                        `Match <strong>${currentData.matchNumber}</strong> - Team <strong>${currentData.teamNumber}</strong><br>` +
                        `Scout: ${currentData.scoutName}`;

                    document.getElementById('successDisplay').style.display = 'block';
                    document.getElementById('reader').style.display = 'none';
                    scanner.pause();

                } catch(e) { console.error("Parse error", e); }
            }
            scanner.render(onScan);
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
                    alert("✅ Saved!");
                    location.reload();
                } else {
                    alert("❌ Error saving");
                }
            } catch(e) { alert("❌ Network Error"); }
        }
    </script>
</body>
</html>
"""

# --- DASHBOARD HTML ---
DASHBOARD_HTML = """
<!DOCTYPE html>
<html>
<head>
    <title>FRC Data</title>
    <style>
        body { background: #1a1a1a; color: white; font-family: sans-serif; padding: 20px; }
        .nav { margin-bottom: 30px; text-align: center; }
        .nav a { color: #00ccff; margin: 0 15px; text-decoration: none; font-size: 1.2rem; border-bottom: 1px solid #00ccff; }
        table { width: 100%; border-collapse: collapse; background: #2a2a2a; }
        th, td { padding: 10px; text-align: left; border-bottom: 1px solid #444; }
        th { background: #333; color: #00ccff; }
        .btn-ods { background: #00a854; color: white; padding: 10px; text-decoration: none; border-radius: 5px; display:inline-block; margin-bottom:20px;}
        .btn-del { background: #ff3333; color: white; border: none; padding: 5px; cursor: pointer; }
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
                <th>Match</th>
                <th>Team</th>
                <th>Scout</th>
                <th>Auto</th>
                <th>A-Climb</th>
                <th>Tele</th>
                <th>E-Climb</th>
                <th>Notes</th>
                <th>Action</th>
            </tr>
        </thead>
        <tbody>
            {% for r in matches %}
            <tr>
                <td>{{ r.scan_time_utc }}</td>
                <td>{{ r.matchNumber }}</td>
                <td>{{ r.teamNumber }}</td>
                <td>{{ r.scoutName }}</td>
                <td>{{ r.autoBalls }}</td>
                <td>{{ r.autoClimb }}</td>
                <td>{{ r.teleBalls }}</td>
                <td>{{ r.endClimb }}</td>
                <td>{{ r.notes }}</td>
                <td><button class="btn-del" onclick="del({{ r.id }})">X</button></td>
            </tr>
            {% endfor %}
        </tbody>
    </table>
    <script>
        async function del(id) {
            if(confirm("Delete match?")) {
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
    c.execute("SELECT id, data_json FROM matches ORDER BY match_number DESC")
    rows = c.fetchall()
    conn.close()

    clean_data = []
    for r in rows:
        try:
            d = json.loads(r[1])
            d['id'] = r[0]
            if 'scan_time_utc' not in d:
                d['scan_time_utc'] = 'N/A'
            clean_data.append(d)
        except: pass

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
            (match_number, team_number, scout_name, data_json, scan_time)
            VALUES (?, ?, ?, ?, ?)
        ''', (
            data['matchNumber'], 
            data['teamNumber'], 
            data['scoutName'], 
            json.dumps(data),
            utc_now
        ))
        conn.commit()
        conn.close()
        
        # --- CUSTOM TERMINAL LOG ---
        scout_name = data.get('scoutName', 'Unknown')
        print(f"✅ {scout_name} submitted a scan at {utc_now}")
        
        return jsonify({'success': True, 'timestamp': utc_now})
    except Exception as e:
        print(f"❌ SERVER ERROR: {str(e)}")
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
    c.execute("SELECT data_json FROM matches ORDER BY match_number")
    rows = c.fetchall()
    conn.close()

    parsed_data = []
    for r in rows:
        d = json.loads(r[0])
        parsed_data.append({
            'Scan Time (UTC)': d.get('scan_time_utc', 'N/A'),
            'Match': d.get('matchNumber'),
            'Team': d.get('teamNumber'),
            'Scout': d.get('scoutName'),
            'Auto Balls': d.get('autoBalls'),
            'Auto Climb': d.get('autoClimb'),
            'Teleop Balls': d.get('teleBalls'),
            'Endgame Climb': d.get('endClimb'),
            'Outcome': d.get('outcome'),
            'Defense': d.get('defense'),
            'Robot Broke': d.get('broken'),
            'Notes': d.get('notes')
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
    # ==========================================================
    # PRE-FLIGHT CHECK 2: Scanner Javascript File
    # ==========================================================
    if not os.path.exists('html5-qrcode.min.js'):
        print("\n" + "!"*55)
        print("🛑 SERVER STARTUP FAILED: Missing Scanner File")
        print("!"*55)
        print("The scanner requires 'html5-qrcode.min.js' to function.")
        print("👉 Please download it and place it in the SAME folder as this script.\n")
        sys.exit(1)

    # Initialize the database and boot the server
    init_db()
    
    print("\n" + "="*50)
    print("🔗 Scanner Interface: http://localhost:5000")
    print("🛑 Press CTRL+C to quit")
    print("="*50 + "\n")
    
    app.run(host='0.0.0.0', port=5000, debug=False)