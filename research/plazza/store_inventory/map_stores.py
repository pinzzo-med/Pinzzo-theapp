"""Map Plazza's store network: which store serves each point of a Bengaluru grid,
with ETA, fees, distance and the pincode Plazza resolves. Resumable (stores.sqlite)."""
import json, re, sqlite3, sys, time, urllib.error, urllib.parse, urllib.request
from datetime import datetime, timezone
from pathlib import Path

HERE = Path(__file__).parent
db = sqlite3.connect(HERE / 'stores.sqlite')  # separate file: the catalogue crawl writes plazza.sqlite
db.execute('''CREATE TABLE IF NOT EXISTS locations(lat REAL, lng REAL, serviceable INT, store_id TEXT, pincode TEXT,
    locality TEXT, eta_min INT, eta_text TEXT, deferred_text TEXT, secondary_eta TEXT, distance_m REAL,
    platform_fee REAL, packaging REAL, convenience REAL, delivery REAL, split INT, raw TEXT, at TEXT,
    PRIMARY KEY(lat, lng))''')
HEADERS = {'Accept': 'application/json', 'X-App-Type': 'web_app', 'User-Agent': 'Mozilla/5.0',
           'Content-Type': 'application/json'}
URL = 'https://auth-service.plazza.in/api/v1/'
STEP = 0.0135  # about 1.5 km
LAT0, LAT1, LNG0, LNG1 = 12.80, 13.16, 77.44, 77.80


def call(path, params=None, data=None):
    url = URL + path + ('?' + urllib.parse.urlencode(params) if params else '')
    req = urllib.request.Request(url, headers=HEADERS, data=json.dumps(data).encode() if data is not None else None,
                                 method='POST' if data is not None else 'GET')
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.load(r)


row = db.execute("SELECT value FROM meta WHERE key='map_contact'").fetchone() if db.execute(
    "SELECT name FROM sqlite_master WHERE name='meta'").fetchone() else None
db.execute('CREATE TABLE IF NOT EXISTS meta(key TEXT PRIMARY KEY, value TEXT)')
if row:
    contact = row[0]
else:
    contact = call('auth/create-guest', data={})['contact_id']
    db.execute("INSERT INTO meta VALUES ('map_contact', ?)", (contact,))
    db.commit()

points = []
lat = LAT0
while lat <= LAT1:
    lng = LNG0
    while lng <= LNG1:
        points.append((round(lat, 4), round(lng, 4)))
        lng += STEP
    lat += STEP
done = {(a, b) for a, b in db.execute('SELECT lat, lng FROM locations')}
todo = [p for p in points if p not in done]
print(f'{len(points)} grid points, {len(todo)} to do', flush=True)

strikes = 0
for i, (la, ln) in enumerate(todo):
    time.sleep(1.5)
    try:
        d = call('contact-addresses/service-details', {'latitude': la, 'longitude': ln, 'contact_id': contact})
        strikes = 0
    except urllib.error.HTTPError as e:
        if e.code in (403, 429):
            strikes += 1
            if strikes > 1:
                print('throttled or blocked twice: stopping', flush=True)
                sys.exit(2)
            time.sleep(120)
            continue
        d = {'error': e.code}
    except Exception as e:
        d = {'error': repr(e)}
    addr = (d.get('address') or {}).get('formatted_address') or ''
    pin = re.search(r'\b(5\d{5})\b', addr)
    db.execute('INSERT OR REPLACE INTO locations VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)', (
        la, ln, int(bool(d.get('serviceable'))), d.get('store_id'), pin.group(1) if pin else None, d.get('locality'),
        d.get('deliveryTime'), d.get('instant_delivery_time'), d.get('deferred_delivery_time'),
        d.get('secondary_delivery_time'), d.get('distance_meters'), d.get('platform_fee'), d.get('packaging_charges'),
        d.get('convenience_fee'), d.get('delivery_charges'), int(bool(d.get('show_split_delivery'))),
        json.dumps(d, ensure_ascii=False), datetime.now(timezone.utc).isoformat(timespec='seconds')))
    db.commit()
    if i % 50 == 0:
        s = db.execute('SELECT count(*), sum(serviceable), count(DISTINCT store_id) FROM locations').fetchone()
        print(f'{i}/{len(todo)} points; serviceable {s[1]}; stores {s[2]}', flush=True)

s = db.execute('SELECT count(*), sum(serviceable), count(DISTINCT store_id), count(DISTINCT pincode) FROM locations').fetchone()
print(f'done: {s[0]} points, {s[1]} serviceable, {s[2]} stores, {s[3]} pincodes', flush=True)
