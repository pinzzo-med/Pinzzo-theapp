"""Turn stock.sqlite into CSV files. Plain Python, no extra packages.

  python3 export_stock.py
Creates: store_stock.csv (one row per store x product), store_summary.csv (one row per store)
"""
import csv, json, sqlite3
from pathlib import Path

HERE = Path(__file__).parent
db = sqlite3.connect(HERE / 'stock.sqlite')

cols = ['store_id', 'store_pincodes', 'product_id', 'name', 'brand', 'manufacturers', 'salt_composition',
        'packaging_detail', 'tag', 'prescription_required', 'mrp', 'selling_price', 'discount', 'offer_price_999',
        'store_inventory', 'delivery_type', 'fast_delivery', 'collected_at']
stores = {s: p for s, p in db.execute('SELECT store_id, pincodes FROM stores')}
with (HERE / 'store_stock.csv').open('w', newline='', encoding='utf-8-sig') as f:
    w = csv.writer(f)
    w.writerow(cols)
    for sid, pid, inv, dt, fast, sp, mrp, disc, offer, raw, at in db.execute(
            'SELECT store_id, product_id, store_inventory, delivery_type, fast_delivery, selling_price, mrp, discount, '
            'offer_price, raw, at FROM stock ORDER BY store_id, product_id'):
        p = json.loads(raw)
        w.writerow([sid, stores.get(sid), pid, p.get('name'), p.get('brand'), p.get('manufacturers'),
                    p.get('salt_composition'), p.get('packaging_detail'), p.get('tag'), p.get('prescription_required'),
                    mrp, sp, disc, offer, inv, dt, fast, at])

with (HERE / 'store_summary.csv').open('w', newline='', encoding='utf-8-sig') as f:
    w = csv.writer(f)
    w.writerow(['store_id', 'pincodes', 'localities', 'searches_done', 'products_seen', 'in_stock', 'in_stock_pct',
                'units_in_stock', 'deferred_or_zero'])
    total = len([q for q in (HERE / 'queries.txt').read_text(encoding='utf-8').splitlines() if q.strip()])
    for sid, pins, locs in db.execute('SELECT store_id, pincodes, localities FROM stores ORDER BY store_id'):
        done = db.execute('SELECT count(*) FROM done WHERE store_id=?', (sid,)).fetchone()[0]
        seen, ins, units = db.execute('SELECT count(*), sum(store_inventory>0), sum(CASE WHEN store_inventory>0 '
                                      'THEN store_inventory ELSE 0 END) FROM stock WHERE store_id=?', (sid,)).fetchone()
        w.writerow([sid, pins, locs, f'{done}/{total}', seen, ins or 0,
                    round((ins or 0) / seen * 100, 1) if seen else '', units or 0, seen - (ins or 0)])
print('wrote store_stock.csv and store_summary.csv')
