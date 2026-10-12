# Plazza competitive intelligence: Phase 1 report and observation plan

Observed 12 October 2026. Target: **plazza.in** (PLAZZA HEALTH PRIVATE LIMITED). Benchmark: Pinzzo.

Every statement is marked **observed** (seen directly), **claimed** (Plazza's own statement, not verified) or **pending** (to be collected).

## 1. Phase 1 findings

| Topic | Finding | Kind |
|---|---|---|
| Company | PLAZZA HEALTH PRIVATE LIMITED, founded 2024 | claimed |
| Address | 5th Floor, Tower A, No. 51, Innov8 Mantri Commercio, Bellandur, Bengaluru 560103 | claimed |
| Market | **Bengaluru only.** No Gurugram presence is claimed | claimed |
| Promise | "Medicines in 15 minutes", 24/7, handling time 5–15 min | claimed |
| Assortment | 20,000+, 35,000+ and 100,000+ are all claimed in different places on the same page | claimed, inconsistent |
| Minimum order | None | claimed |
| Delivery charge | ₹0 in the page's offer data; real checkout fees not yet seen | claimed |
| Payments | COD, cards, UPI, net banking, wallets | claimed |
| Prescriptions | Upload in the app; a "saved prescriptions" page exists | claimed / observed |
| Channels | Website, Android app (`com.plazza.app`), iOS app, email (support@, founders@) | observed |
| Phone | Their structured data shows a placeholder number, "+91-XXXXX-XXXXX" | observed |
| App rating | 4.1 from 31 ratings, from their own page markup (not checked on the stores) | claimed |
| Website technology | Next.js app. Products, prices, stock and even policy text load in the browser from `/api/` | observed |
| Crawling rules | `robots.txt` allows pages but **disallows `/api/`** for all bots | observed |

## 2. What this means for the extraction plan

Your brief says to respect robots.txt and to switch to compliant manual observation where automated collection is not permitted (sections 3 and 22). That applies here:

- **Automated SKU, price, availability and policy collection: not performed.** All of that data is served only from `/api/`, which Plazza tells every bot not to use. Running a browser bot over thousands of pages would still hit `/api/` thousands of times.
- **Allowed and done:** domain verification, robots.txt, sitemap, and the public structured data in page HTML. Evidence is saved in `evidence/`.
- **Everything else is collected by a person** browsing as an ordinary customer at one fixed Bengaluru address, recorded in the workbook. This is how most retail pricing teams benchmark competitors.

**Market note:** Plazza serves Bengaluru and Pinzzo serves Gurugram, so they don't compete for the same customers today. This study benchmarks Plazza's **model**: speed promise, fees, assortment strategy, Rx flow and membership. Its prices are not a direct reference for Gurugram. For local price benchmarking, run the same workbook against apps that deliver in Gurugram.

## 3. Observer protocol (about 6–8 hours spread over one week)

**Set-up**
1. Use the Plazza app or plazza.in in a normal browser. Log in with a company phone number, never a personal or patient one.
2. Choose **one fixed Bengaluru address**, for example a public landmark near Bellandur (560103), and use it for every observation. Record it in the *Serviceability* tab.
3. Take a screenshot of every screen you record. Name it after the row ID (for example `SKU-0001.png` or `B05.png`) and save it in `evidence/screenshots/`.

**Day 1 (2–3 h): categories and a product sample**
- *Category Tree:* every category and subcategory in the navigation, in order.
- *Master SKU Catalogue:* about **20 products per main category, 200 or more in total**. Include Pinzzo's top sellers when you have that list. Copy MRP, selling price, pack and stock status exactly as shown. Leave fields blank when they aren't shown.

**Day 2 (1–2 h): search**
- Run the 20 queries pre-filled in *Search Tests* and record the count, exact-match position and top 5 results.

**Day 3 (1–2 h): baskets and fees, stopping before payment**
- Build the 12 baskets (₹99 to ₹2,999) in *Basket Experiments*. Copy every fee line and the total. A *Total check* of 0 means the figures add up.
- Fill in the *Checkout Fee Matrix*, especially the free-delivery threshold, small-cart fee and late-night fee.

**Days 1–7: delivery promises at four times of day**
- Note the ETA text on the home page, cart and checkout at 8–10, 12–2, 6–8 and 9–11 IST. That's 28 or more rows in *Delivery ETA*.

**Throughout the week:** *Promotions*, *Prescription Workflows*, *Membership and Loyalty*, *Customer Experience*, *Policies and Compliance* (copy the terms and refund pages), and *WhatsApp and Ordering Channels*.

**Week 2 and later (optional):** re-check 50 products from the sample in *Change Detection* to spot price and stock changes.

**Never:** place or pay for an order, upload a real prescription or patient data, message their WhatsApp in bulk, or use scripts or bots on their site or app.

## 4. Analysis I will produce once the workbook is filled
- Catalogue depth and category mix (sample-based, with sample sizes stated)
- Observed availability % by category, using only SKUs with a known stock status as the denominator
- Discount distribution (median and average by category) and how often selling price exceeds MRP
- Fee curve by basket value, free-delivery threshold, and fees as a % of the basket
- How the ETA promise varies by time of day (promised, not actual, delivery times)
- Search quality: exact-match success rate, no-result rate, typo handling
- Rx, membership and channel comparison, the *Pinzzo Comparison* table, and the five most defensible opportunities for Pinzzo

## 5. Files
| File | What it is |
|---|---|
| `Plazza_Competitive_Intelligence.xlsx` | All 20 tabs. Green cells are Phase 1 facts; formulas calculate the metrics; dropdowns keep entries consistent |
| `Plazza_Competitive_Analysis.md` | This report |
| `build_workbook.py` | Rebuilds the blank workbook |
| `evidence/` | robots.txt, sitemap.xml and home page HTML, captured 12 Oct 2026 |
