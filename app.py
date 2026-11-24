"""Connected Couriers: compare Amazon route drivers with a supplied roster."""
import csv
import io
import os
import re
import unicodedata
from flask import Flask, render_template, request
from openpyxl import load_workbook

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 8 * 1024 * 1024

# Known, reviewed spelling variants, not general fuzzy matching.
ALIASES = {
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
        name = str(row[idx] or '').strip() if len(row) > idx else ''
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
        if valid_name(name):
            found.setdefault(normalise(name), name)
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
            roster = read_rows(request.files.get('drivers'))
            if not roster:
                roster = pasted_rows(request.form.get('driver_text', ''))
            drivers = driver_names(roster)
            if not routes:
                raise ValueError('No route driver names found in the Amazon file.')
            if not drivers:
                raise ValueError('No names found in the second driver list.')
            registered = {normalise(n) for n in drivers}
            missing = [n for n in routes if normalise(n) not in registered]
            result = {'route_count': len(routes), 'list_count': len(drivers),
                      'matched_count': len(routes) - len(missing), 'missing': missing}
        except (ValueError, UnicodeError, OSError, KeyError, TypeError) as exc:
            error = str(exc)
        except Exception:
            app.logger.exception('Unable to compare driver files')
            error = 'Could not read the files. Check that both files are valid and try again.'
    return render_template('index.html', results=result, error=error)


@app.route('/health')
def health():
    return {'ok': True}


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', '8080')))
