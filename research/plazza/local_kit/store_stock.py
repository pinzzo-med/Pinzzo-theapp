"""Plazza per-store stock collector. Plain Python 3.9+, no extra packages.

For each store in stores.json: pin a guest session to a location that store serves,
replay every search in queries.txt, and save stock, price and delivery type per product
into stock.sqlite. Safe to stop and re-run: it skips work already saved.

  python3 store_stock.py                          # all stores
  python3 store_stock.py --stores aa49,b666       # only stores whose id starts with these
  python3 store_stock.py --budget 3000            # stop after 3000 requests

About 1 request per second. Stops by itself if Plazza throttles (HTTP 429) or blocks (403).
"""
import argparse, json, sqlite3, sys, time, urllib.error, urllib.parse, urllib.request
from datetime import datetime, timezone
from pathlib import Path

HERE = Path(__file__).parent
ap = argparse.ArgumentParser()
ap.add_argument('--stores', default='', help='comma-separated store id prefixes (default: all)')
ap.add_argument('--budget', type=int, default=10**9, help='max requests this run')
args = ap.parse_args()

H = {'Accept': 'application/json', 'X-App-Type': 'web_app', 'User-Agent': 'Mozilla/5.0', 'Content-Type': 'application/json'}
AUTH = 'https://auth-service.plazza.in/api/v1/'
SEARCH = 'https://cma.plazza.in/api/mobile/catalogue/search'

db = sqlite3.connect(HERE / 'stock.sqlite')
db.executescript('''
CREATE TABLE IF NOT EXISTS stores(store_id TEXT PRIMARY KEY, contact_id TEXT, lat REAL, lng REAL, pincodes TEXT, localities TEXT);
CREATE TABLE IF NOT EXISTS done(store_id TEXT, q TEXT, PRIMARY KEY(store_id, q));
CREATE TABLE IF NOT EXISTS stock(store_id TEXT, product_id TEXT, store_inventory REAL, delivery_type TEXT, fast_delivery INT,
  selling_price REAL, mrp REAL, discount REAL, offer_price REAL, raw TEXT, at TEXT, PRIMARY KEY(store_id, product_id));
CREATE TABLE IF NOT EXISTS errors(at TEXT, store_id TEXT, q TEXT, err TEXT);
''')
for s in json.loads((HERE / 'stores.json').read_text(encoding='utf-8')):
    db.execute('INSERT OR IGNORE INTO stores VALUES (?,?,?,?,?,?)',
               (s['store_id'], None, s['lat'], s['lng'], s['pincodes'], s['localities']))
db.commit()
queries = [q.strip() for q in (HERE / 'queries.txt').read_text(encoding='utf-8').splitlines() if q.strip()]
wanted = [p.strip() for p in args.stores.split(',') if p.strip()]

now = lambda: datetime.now(timezone.utc).isoformat(timespec='seconds')
made = 0


def say(msg):
    print(f'{now()} {msg}', flush=True)


def http(url, params=None, data=None):
    global made
    for attempt in range(4):
        time.sleep(1.0)
        made += 1
        try:
            req = urllib.request.Request(url + ('?' + urllib.parse.urlencode(params) if params else ''), headers=H,
                                         data=json.dumps(data).encode() if data is not None else None,
                                         method='POST' if data is not None else 'GET')
            with urllib.request.urlopen(req, timeout=30) as r:
                return json.load(r)
        except urllib.error.HTTPError as e:
            if e.code in (403, 429):
                if attempt == 0:
                    say(f'HTTP {e.code}: pausing 2 minutes')
                    time.sleep(120)
                    continue
                say(f'HTTP {e.code} again: Plazza is throttling or blocking. Stopping.')
                sys.exit(2)
            if e.code == 422:
                return None
            err = f'HTTP {e.code}'
        except Exception as e:
            err = repr(e)
        db.execute('INSERT INTO errors VALUES (?,?,?,?)', (now(), url[-40:], str(params)[:80], err))
        db.commit()
        time.sleep(2 ** (attempt + 1))
    return None


def pin(sid, contact, lat, lng):
    d = http(AUTH + 'contact-addresses/service-details', {'latitude': lat, 'longitude': lng, 'contact_id': contact})
    if not d or d.get('store_id') != sid:
        say(f'warning: location for store {sid[:8]} resolved to {(d or {}).get("store_id")}; skipping this store')
        return False
    return True


say(f'{len(queries)} searches per store; stores: {", ".join(wanted) or "all"}')
for sid, contact, lat, lng in db.execute('SELECT store_id, contact_id, lat, lng FROM stores ORDER BY store_id').fetchall():
    if wanted and not any(sid.startswith(w) for w in wanted):
        continue
    done = {q for (q,) in db.execute('SELECT q FROM done WHERE store_id=?', (sid,))}
    todo = [q for q in queries if q not in done]
    if not todo:
        say(f'store {sid[:8]}: already complete')
        continue
    if not contact:
        contact = http(AUTH + 'auth/create-guest', data={})['contact_id']
        db.execute('UPDATE stores SET contact_id=? WHERE store_id=?', (contact, sid))
        db.commit()
    if not pin(sid, contact, lat, lng):
        continue
    pinned_at = time.time()
    say(f'store {sid[:8]}: {len(todo)} searches left')
    for i, q in enumerate(todo):
        if made >= args.budget:
            say('request budget reached; run again to continue')
            sys.exit(0)
        if time.time() - pinned_at > 540:  # the app refreshes the location every 10 minutes
            pin(sid, contact, lat, lng)
            pinned_at = time.time()
        page = 1
        while True:
            d = http(SEARCH, {'query': q, 'page': page, 'page_size': 50, 'contact_id': contact})
            if not d or not d.get('success'):
                break
            x = d['data']
            for p in (x.get('results') or []) + ((x.get('out_of_stock') or {}).get('products') or []):
                di = p.get('discount_info') or {}
                db.execute('INSERT OR REPLACE INTO stock VALUES (?,?,?,?,?,?,?,?,?,?,?)', (
                    sid, p['product_id'], p.get('store_inventory', p.get('inventory_quantity')), p.get('delivery_type'),
                    int(bool(p.get('fast_delivery'))), p.get('plazza_selling_price_incl_gst'), p.get('mrp'),
                    p.get('discount'), di.get('price'), json.dumps(p, ensure_ascii=False), now()))
            if not (x.get('pagination') or {}).get('has_next'):
                break
            page += 1
        db.execute('INSERT OR IGNORE INTO done VALUES (?,?)', (sid, q))
        db.commit()
        if i % 200 == 0:
            n = db.execute('SELECT count(*) FROM stock WHERE store_id=?', (sid,)).fetchone()[0]
            say(f'store {sid[:8]}: {i}/{len(todo)} searches, {n} products, {made} requests')

say(f'finished; {db.execute("SELECT count(*) FROM stock").fetchone()[0]} store x product rows saved in stock.sqlite')
