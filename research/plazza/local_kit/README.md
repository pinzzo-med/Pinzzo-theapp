# Plazza per-store stock: run it on your laptop

Works on Ubuntu or Windows with Python 3.9 or newer. No extra packages needed.

## Your share of the work
The cloud session is collecting stores `7e4b6aef`, `7e91be9f` and `a258ba87` (`4436157a` is already done).
Your laptop collects the other 5:

| Store id starts with | Area |
|---|---|
| `aa49e0e4` | Kalyan Nagar, Hennur, Thanisandra |
| `b666cfc6` | Bellandur, Rayasandra, Kodathi |
| `bb28573b` | HSR Layout, BTM, Koramangala |
| `f7070977` | Basavanagudi, Vijayanagar, Chickpet |
| `fe4c5211` | Anjanapura, Gottigere, Talaghattapura |

## Run it

**Ubuntu** (open Terminal in this folder):
```bash
python3 --version          # must be 3.9 or newer
python3 store_stock.py --stores aa49e0e4,b666cfc6,bb28573b,f7070977,fe4c5211
```

**Windows** (open Command Prompt in this folder):
```
python --version
python store_stock.py --stores aa49e0e4,b666cfc6,bb28573b,f7070977,fe4c5211
```

- It takes about 1 hour per store, so roughly 6 hours for all 5. Start it in the evening and leave the laptop plugged in with sleep turned off.
- It prints progress every 200 searches.
- **Stopped or the laptop slept?** Run the same command again. It carries on where it stopped.
- If it prints "Plazza is throttling or blocking. Stopping.", **don't** force it. Wait a few hours and run again.
- Don't speed it up or run several copies at once. One request per second keeps it from being blocked.

## When it finishes
```bash
python3 export_stock.py        # Windows: python export_stock.py
```
This creates `store_stock.csv` and `store_summary.csv`. Send me **`stock.sqlite`** (or the two CSVs) and I'll merge them with the cloud stores into one workbook.

## Files
| File | What it is |
|---|---|
| `store_stock.py` | The collector |
| `stores.json` | Plazza's 9 stores, with a location each one serves and its pincodes |
| `queries.txt` | The searches to replay at each store |
| `export_stock.py` | Turns the results into CSV files |
| `stock.sqlite` | Created on first run; all results go here |
