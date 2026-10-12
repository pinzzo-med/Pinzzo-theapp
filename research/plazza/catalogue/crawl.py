"""Plazza catalogue collector: search-driven, rate-limited, resumable.

Pass 1: two-letter prefixes + digits. Pass 2: snowball on brands, name words and
salts found so far, until no new products or the request budget runs out.
Stops on HTTP 403/429 (after one long pause). State lives in plazza.sqlite.
"""
import itertools, json, re, sqlite3, string, sys, time, urllib.error, urllib.parse, urllib.request
from datetime import datetime, timezone
from pathlib import Path

HERE = Path(__file__).parent
DB = HERE / 'plazza.sqlite'
LOG = HERE / 'crawl.log'
API = 'https://cma.plazza.in/api/mobile/catalogue/search'
CONTACT = '30836b00-b5a8-4650-a7a8-39675f6776a5'  # one guest session, no address set
DELAY = 1.0
BUDGET = int(sys.argv[1]) if len(sys.argv) > 1 else 8000
HEADERS = {'Accept': 'application/json', 'X-App-Type': 'web_app', 'User-Agent': 'Mozilla/5.0'}

db = sqlite3.connect(DB)
db.executescript('''
CREATE TABLE IF NOT EXISTS queries(q TEXT PRIMARY KEY, pass INT, status TEXT, total INT, pages INT, new_products INT, done_at TEXT);
CREATE TABLE IF NOT EXISTS products(product_id TEXT PRIMARY KEY, raw TEXT, listed_as TEXT, first_seen TEXT, last_seen TEXT, hits INT);
CREATE TABLE IF NOT EXISTS errors(at TEXT, q TEXT, page INT, error TEXT);
''')


def now():
    return datetime.now(timezone.utc).isoformat(timespec='seconds')


def log(msg):
    line = f'{now()} {msg}'
    print(line, flush=True)
    with LOG.open('a') as f:
        f.write(line + '\n')


requests_made = 0


def get(q, page):
    global requests_made
    url = API + '?' + urllib.parse.urlencode({'query': q, 'page': page, 'page_size': 50, 'contact_id': CONTACT})
    for attempt in range(4):
        time.sleep(DELAY)
        requests_made += 1
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=HEADERS), timeout=30) as r:
                return json.load(r)
        except urllib.error.HTTPError as e:
            if e.code in (403, 429):
                if attempt == 0:
                    log(f'HTTP {e.code} on {q!r}; pausing 120 s once')
                    time.sleep(120)
                    continue
                log(f'HTTP {e.code} again: Plazza is throttling or blocking. Stopping.')
                sys.exit(2)
            if e.code == 422:
                raise
            err = f'HTTP {e.code}'
        except Exception as e:  # network hiccup
            err = repr(e)
        db.execute('INSERT INTO errors VALUES (?,?,?,?)', (now(), q, page, err))
        db.commit()
        time.sleep(2 ** (attempt + 1))
    return None


def save(p, listed_as):
    pid = p.get('product_id')
    if not pid:
        return 0
    t = now()
    cur = db.execute('SELECT 1 FROM products WHERE product_id=?', (pid,)).fetchone()
    if cur:
        db.execute('UPDATE products SET raw=?, listed_as=?, last_seen=?, hits=hits+1 WHERE product_id=?',
                   (json.dumps(p, ensure_ascii=False), listed_as, t, pid))
        return 0
    db.execute('INSERT INTO products VALUES (?,?,?,?,?,1)', (pid, json.dumps(p, ensure_ascii=False), listed_as, t, t))
    return 1


def run_query(q, pas):
    if db.execute('SELECT 1 FROM queries WHERE q=?', (q,)).fetchone():
        return
    new = total = pages = 0
    page = 1
    status = 'ok'
    while True:
        try:
            d = get(q, page)
        except urllib.error.HTTPError as e:
            status = f'http_{e.code}'
            break
        if not d or not d.get('success'):
            status = 'failed'
            break
        x = d['data']
        for p in x.get('results') or []:
            new += save(p, 'results')
        for p in (x.get('out_of_stock') or {}).get('products') or []:
            new += save(p, 'out_of_stock_block')
        pg = x.get('pagination') or {}
        total, pages = pg.get('total_results', 0), pg.get('total_pages', 0)
        if not pg.get('has_next'):
            break
        page += 1
    db.execute('INSERT OR REPLACE INTO queries VALUES (?,?,?,?,?,?,?)', (q, pas, status, total, pages, new, now()))
    db.commit()


def snowball_terms():
    seen = {r[0] for r in db.execute('SELECT q FROM queries')}
    terms = set()
    for (raw,) in db.execute('SELECT raw FROM products'):
        p = json.loads(raw)
        if p.get('brand'):
            terms.add(p['brand'].lower().strip())
        w = re.findall(r'[A-Za-z]{3,}', p.get('name') or '')
        if w:
            terms.add(w[0].lower())
        for salt in re.split(r'[,+]| and ', p.get('salt_composition') or ''):
            s = re.sub(r'\(.*?\)', '', salt).strip().lower()
            if 3 <= len(s) <= 40:
                terms.add(s)
    return sorted(t for t in terms if t not in seen)


def count():
    return db.execute('SELECT count(*) FROM products').fetchone()[0]


log(f'start: {count()} products already stored, budget {BUDGET} requests')
seeds = [''.join(t) for t in itertools.product(string.ascii_lowercase, repeat=2)] + list(string.digits)
for i, q in enumerate(seeds):
    if requests_made >= BUDGET:
        break
    run_query(q, 1)
    if i % 50 == 0:
        log(f'pass 1: {i}/{len(seeds)} queries, {count()} products, {requests_made} requests')

rounds = 0
while requests_made < BUDGET:
    terms = snowball_terms()
    if not terms:
        break
    rounds += 1
    before = count()
    log(f'pass 2 round {rounds}: {len(terms)} new terms, {before} products')
    for i, q in enumerate(terms):
        if requests_made >= BUDGET:
            break
        run_query(q, 2)
        if i % 100 == 0:
            log(f'  round {rounds}: {i}/{len(terms)}, {count()} products, {requests_made} requests')
    if count() == before:
        break

log(f'finished: {count()} products, {requests_made} requests this run, '
    f'{db.execute("SELECT count(*) FROM errors").fetchone()[0]} errors logged')
