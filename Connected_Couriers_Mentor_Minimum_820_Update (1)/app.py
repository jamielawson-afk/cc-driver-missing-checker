"""Connected Couriers: compare Amazon route drivers with a supplied roster."""
import csv
import io
import os
import re
import sqlite3
import secrets
import time
from werkzeug.security import check_password_hash
from pathlib import Path
from flask import redirect, url_for, flash
import unicodedata
from flask import Flask, render_template, request, session, abort
from openpyxl import load_workbook

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 8 * 1024 * 1024
app.secret_key = os.environ.get('SECRET_KEY', 'development-only-change-this')
app.config.update(SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SECURE=True, SESSION_COOKIE_SAMESITE='Lax', PERMANENT_SESSION_LIFETIME=3600)
_attempts = {}

@app.route('/login', methods=['GET','POST'])
def office_login():
    if session.get('office'):
        return redirect('/')
    error = None
    if request.method == 'POST':
        ip = request.remote_addr or 'unknown'
        now = time.monotonic()
        recent = [t for t in _attempts.get(ip, []) if now-t < 900]
        if len(recent) >= 10:
            error = 'Too many attempts. Try again later.'
        else:
            password_hash = os.environ.get('OFFICE_PASSWORD_HASH', '')
            if password_hash and check_password_hash(password_hash, request.form.get('password','')):
                _attempts.pop(ip, None)
                session.clear()
                session.permanent = True
                session['office'] = True
                session['csrf'] = secrets.token_urlsafe(32)
                return redirect('/')
            recent.append(now)
            _attempts[ip] = recent
            error = 'Incorrect office password.'
    return render_template('login.html', error=error)

@app.before_request
def protect_office():
    if request.endpoint in ('health','office_login','static'):
        return
    if not session.get('office'):
        return redirect('/login')
    if request.method == 'POST' and request.form.get('csrf_token','') != session.get('csrf'):
        abort(400)

@app.post('/logout')
def office_logout():
    session.clear()
    return redirect('/login')

@app.context_processor
def csrf_context():
    return {'csrf_token': session.get('csrf','')}

DB_PATH = Path(os.environ.get('MASTER_DB_PATH', '/data/mentor.sqlite3' if Path('/data').is_dir() else 'mentor.sqlite3'))

def db():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    conn.execute('CREATE TABLE IF NOT EXISTS master (id INTEGER PRIMARY KEY, name TEXT NOT NULL, name_key TEXT NOT NULL UNIQUE, code1 TEXT NOT NULL, code2 TEXT NOT NULL, UNIQUE(code1, code2))')
    # Seed once on first database creation. Do not restore deleted entries later.
    if not conn.execute('SELECT 1 FROM master LIMIT 1').fetchone() and not conn.execute('PRAGMA user_version').fetchone()[0]:
        import json
        seed = json.loads((Path(__file__).parent / 'master_seed.json').read_text())
        conn.executemany('INSERT INTO master (name, name_key, code1, code2) VALUES (?, ?, ?, ?)', [(r['name'], normalise(r['name']), r['code1'], r['code2']) for r in seed])
        conn.execute('PRAGMA user_version = 1')
    conn.commit()
    return conn

def master_list():
    with db() as conn:
        return [dict(row) for row in conn.execute('SELECT id, name, code1, code2 FROM master ORDER BY name COLLATE NOCASE')]


# Known, reviewed spelling variants, not general fuzzy matching.
ALIASES = {
    'scott keiran': 'scott kiernan',
    'rhys robins': 'rhys robbins',
    'drew johson': 'drew johnson',
    'kirstry perham': 'kirsty perham',
    'macy rees davies': 'macy davies',
    'joemark bulotano': 'joemark bultano',
}


def normalise(value):
    s = unicodedata.normalize('NFKC', str(value or '')).casefold().strip()
    s = ' '.join(re.sub(r'[^a-z\s-]', ' ', s).replace('-', ' ').split())
    return ALIASES.get(s, s)



# Reviewed common variants; avoid general fuzzy matching that confuses drivers.
FIRST_NAME_ALIASES = {
    'scott keiran': 'scott kiernan',
    'frank': 'franklyn', 'franklin': 'franklyn', 'frankie': 'franklyn',
    'chris': 'christopher', 'bonny': 'bonnykimberley',
    'jim': 'james', 'jimmy': 'james', 'mike': 'michael',
    'dan': 'daniel', 'danny': 'daniel', 'matt': 'matthew',
    'tom': 'thomas', 'tommy': 'thomas', 'will': 'william',
    'bill': 'william', 'billy': 'william', 'rob': 'robert',
    'robbie': 'robert', 'bob': 'robert', 'alex': 'alexander',
    'andy': 'andrew', 'dave': 'david', 'tim': 'timothy',
    'jon': 'jonathan', 'john': 'jonathan',
}

def match_keys(name):
    parts = normalise(name).split()
    if len(parts) < 2:
        return ()
    first = FIRST_NAME_ALIASES.get(parts[0], parts[0])
    surname = parts[-1]
    # The full name and first/surname pair both work with middle names.
    return (normalise(name), first + ' ' + surname, first[0] + ' ' + surname)

def compare_route_drivers(routes, drivers):
    # One-to-one matching; initials only if unique on BOTH sides.
    remaining_routes = set(range(len(routes)))
    remaining_drivers = set(range(len(drivers)))
    keys = {i: match_keys(name) for i, name in enumerate(routes)}
    roster_keys = {i: match_keys(name) for i, name in enumerate(drivers)}
    for level in (0, 1, 2):
        route_groups, driver_groups = {}, {}
        for i in sorted(remaining_routes):
            if len(keys[i]) > level:
                route_groups.setdefault(keys[i][level], []).append(i)
        for j in sorted(remaining_drivers):
            if len(roster_keys[j]) > level:
                driver_groups.setdefault(roster_keys[j][level], []).append(j)
        for key, route_ids in route_groups.items():
            candidates = driver_groups.get(key, [])
            if len(route_ids) == len(candidates) == 1:
                remaining_routes.remove(route_ids[0])
                remaining_drivers.remove(candidates[0])
    return [routes[i] for i in sorted(remaining_routes)]


def read_rows(file):
    if not file or not file.filename:
        return []
    extension = file.filename.rsplit('.', 1)[-1].casefold()
    raw = file.read()
    if extension == 'xlsx':
        workbook = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        return [list(row) for row in workbook.active.iter_rows(values_only=True)]
    if extension == 'csv':
        return list(csv.reader(io.StringIO(raw.decode('utf-8-sig'))))
    raise ValueError('Only .xlsx or .csv files are supported.')


def valid_name(text):
    name = normalise(text)
    words = name.split()
    return 2 <= len(words) <= 5 and all(len(word) >= 2 for word in words)


def route_names(rows):
    if not rows:
        return []
    header = [normalise(v) for v in rows[0]]
    columns = ['driver name', 'driver', 'driver full name']
    idx = next((header.index(h) for h in columns if h in header), None)
    if idx is None:
        raise ValueError('Amazon Routes file must contain a Driver name column.')
    found = {}
    for row in rows[1:]:
        cell = str(row[idx] or '') if len(row) > idx else ''
        # Amazon may put several drivers on a single route, separated by pipes.
        for item in cell.split('|'):
            name = item.strip()
            if valid_name(name):
                found.setdefault(normalise(name), name)
    return list(found.values())


def driver_names(rows):
    if not rows:
        return []
    header = [normalise(v) for v in rows[0]]
    first_keys = ('first name', 'firstname', 'forename')
    last_keys = ('last name', 'lastname', 'surname', 'family name')
    first = next((header.index(x) for x in first_keys if x in header), None)
    last = next((header.index(x) for x in last_keys if x in header), None)
    full = next((header.index(x) for x in ('driver name', 'full name', 'name') if x in header), None)
    is_header = (first is not None and last is not None) or full is not None
    found = {}
    for row in rows[1:] if is_header else rows:
        cells = [str(v or '').strip() for v in row]
        if first is not None and last is not None:
            name = ' '.join(cells[x] for x in (first, last) if x < len(cells))
        elif full is not None:
            name = cells[full] if full < len(cells) else ''
        elif len(cells) >= 2 and len(cells[0].split()) == len(cells[1].split()) == 1:
            name = f'{cells[0]} {cells[1]}'
        else:
            name = cells[0] if cells else ''
        for item in name.split('|'):
            item = item.strip()
            if valid_name(item):
                found.setdefault(normalise(item), item)
    return list(found.values())


def pasted_rows(text):
    if not text.strip():
        return []
    return list(csv.reader(io.StringIO(text), delimiter='\t')) if '\t' in text else [
        [part.strip() for part in line.split(',')] if ',' in line else [line.strip()]
        for line in text.splitlines()
    ]


@app.after_request
def headers(response):
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['Referrer-Policy'] = 'no-referrer'
    response.headers['Cache-Control'] = 'no-store'
    return response


@app.route('/', methods=['GET', 'POST'])
def index():
    result = None
    error = None
    if request.method == 'POST':
        try:
            routes = route_names(read_rows(request.files.get('routes')))
            roster = pasted_rows(request.form.get('driver_text', ''))
            drivers = driver_names(roster)
            if not routes:
                raise ValueError('No route driver names found in the Amazon file.')
            if not drivers:
                raise ValueError('No names found in the second driver list.')
            missing = compare_route_drivers(routes, drivers)
            result = {'route_count': len(routes), 'list_count': len(drivers),
                      'matched_count': len(routes) - len(missing), 'missing': missing}
        except (ValueError, UnicodeError, OSError, KeyError, TypeError) as exc:
            error = str(exc)
        except Exception:
            app.logger.exception('Unable to compare driver files')
            error = 'Could not read the files. Check that both files are valid and try again.'
    master = master_list()
    return render_template('index.html', results=result, error=error, master=master)

@app.post('/master/add')
def add_master():
    name = ' '.join(request.form.get('name', '').split())
    code1 = request.form.get('code1', '').strip()
    code2 = request.form.get('code2', '').strip()
    if not valid_name(name) or not code1 or not code2 or len(code1) > 100 or len(code2) > 100:
        flash('Enter a full name and both codes.')
    else:
        try:
            with db() as conn:
                conn.execute('INSERT INTO master (name, name_key, code1, code2) VALUES (?, ?, ?, ?)', (name, normalise(name), code1, code2))
            flash('Driver added to Master List.')
        except sqlite3.IntegrityError:
            flash('Driver name or combined code pair already exists.')
    return redirect('/#master')

@app.post('/master/edit/<int:driver_id>')
def edit_master(driver_id):
    name = ' '.join(request.form.get('name', '').split())
    code1 = request.form.get('code1', '').strip()
    code2 = request.form.get('code2', '').strip()
    if not valid_name(name) or not code1 or not code2 or len(code1) > 100 or len(code2) > 100:
        flash('Enter a valid name and both codes.')
    else:
        try:
            with db() as conn:
                conn.execute('UPDATE master SET name=?, name_key=?, code1=?, code2=? WHERE id=?', (name, normalise(name), code1, code2, driver_id))
            flash('Driver updated.')
        except sqlite3.IntegrityError:
            flash('Driver name or combined code pair already exists.')
    return redirect('/#master')

@app.post('/master/delete/<int:driver_id>')
def delete_master(driver_id):
    with db() as conn:
        conn.execute('DELETE FROM master WHERE id=?', (driver_id,))
    flash('Driver deleted.')
    return redirect('/#master')





def parse_driver_training_report(file):
    rows = read_rows(file)
    if not rows:
        raise ValueError('Please upload a Driver Report (.xlsx or .csv).')
    headers = [str(x or '').strip() for x in rows[0]]
    required = ('First Name', 'Last Name', 'FICO® Safe Driving Score')
    if any(key not in headers for key in required):
        raise ValueError('This is not a recognised Mentor Driver Report. Expected first/last name codes and FICO Safe Driving Score columns.')
    columns = {key: i for i, key in enumerate(headers) if key}
    with db() as conn:
        drivers = [dict(r) for r in conn.execute('SELECT name, code1, code2 FROM master')]
    by_pair = {}
    for d in drivers:
        key = (d['code1'].strip().casefold(), d['code2'].strip().casefold())
        by_pair.setdefault(key, []).append(d['name'])
    output = []
    for values in rows[1:]:
        def get(key):
            i = columns.get(key)
            return values[i] if i is not None and i < len(values) else None
        first = str(get('First Name') or '').strip()
        last = str(get('Last Name') or '').strip()
        score_value = get('FICO® Safe Driving Score')
        if not first and not last and score_value is None:
            continue
        try:
            score = float(score_value)
        except (ValueError, TypeError):
            continue
        if not (0 <= score <= 850):
            continue
        names = by_pair.get((first.casefold(), last.casefold()), [])
        matched = len(names) == 1
        name = names[0] if matched else 'Unmatched driver' if not names else 'Ambiguous driver'
        findings = []
        if score < 820:
            for count_col, rating_col, title, suggestion in [
                ('Speeding', 'Speeding Rating', 'Speeding', 'Review speed limits, use cruise control where appropriate, and avoid rushing to meet delivery times.'),
                ('Hard Braking', 'Braking Rating', 'Harsh braking', 'Increase following distance, scan further ahead, and brake progressively.'),
                ('Hard Cornering', 'Cornering Rating', 'Harsh cornering', 'Reduce speed before turns and roundabouts; steer smoothly.'),
                ('Hard Acceleration', 'Acceleration Rating', 'Harsh acceleration', 'Pull away gently and increase speed smoothly.'),
                ('Distraction', 'Distraction Rating', 'Driver distraction', 'Keep attention on the road; pull over safely before interacting with devices.'),
                ('Phone Calls', None, 'Phone calls', 'Do not make, accept, or handle calls while driving; stop in a safe place first.'),
                ('Phone Manipulations', None, 'Phone handling', 'Do not touch or operate a phone while driving; set navigation before departure.'),
            ]:
                value = get(count_col)
                rating = str(get(rating_col) or '').strip().casefold() if rating_col else ''
                try: count = float(value or 0)
                except (ValueError, TypeError): count = 0
                if count > 0 or rating in ('medium risk', 'high risk'):
                    findings.append({'issue': title, 'detail': suggestion})
            if not findings:
                findings.append({'issue': 'General driving review', 'detail': 'Review the detailed Mentor feedback with the driver, provide a coached drive, and monitor the next report. The available event fields do not identify a specific cause of the low score.'})
        output.append({'name': name, 'matched': matched, 'code1': first, 'code2': last, 'score': int(score) if score.is_integer() else score,
                       'below': score < 820, 'gap': max(0, round(820-score,1)), 'findings': findings,
                       'trips': get('Total Driver Trips')})
    output.sort(key=lambda d: (d['score'], d['name']))
    return {'drivers': output, 'below': sum(x['below'] for x in output), 'matched': sum(x['matched'] for x in output),
            'unmatched': sum(not x['matched'] for x in output)}

@app.route('/training', methods=['GET', 'POST'])
def training():
    report = None
    error = None
    if request.method == 'POST':
        try:
            report = parse_driver_training_report(request.files.get('driver_report'))
        except (ValueError, UnicodeError, OSError, KeyError, TypeError) as exc:
            error = str(exc)
        except Exception:
            app.logger.exception('Unable to process driver training report')
            error = 'Could not read this Driver Report. Please check the file and try again.'
    return render_template('training.html', report=report, error=error)

@app.route('/health')
def health():
    return {'ok': True}


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', '8080')))
