"""Generate a fictional, deliberately messy sales export (data/sales_raw.csv).

Real exports from shops and CRMs are rarely clean, so the file includes the usual
problems on purpose: duplicate rows, inconsistent region labels, prices typed as
"$1,299.00", mixed date formats, rows with missing quantities and blank lines.
All names and figures are made up.
"""

import csv
import random
from datetime import date, timedelta
from pathlib import Path

random.seed(42)
OUT = Path(__file__).parent / "data" / "sales_raw.csv"

REPS = {
    "North": ["Emma Clarke", "Liam Turner"],
    "South": ["Sofia Ramos", "Noah Bennett"],
    "East": ["Olivia Chen", "Mason Wright"],
    "West": ["Ava Patel", "Ethan Brooks"],
}
REGION_WEIGHTS = {"North": 0.30, "South": 0.20, "East": 0.28, "West": 0.22}

# product, category, list price, typical max quantity, popularity weight
PRODUCTS = [
    ("Laptop Pro 14", "Tech", 1299.00, 4, 6),
    ("27-inch Monitor", "Tech", 329.00, 6, 8),
    ("USB-C Dock", "Tech", 149.00, 8, 7),
    ("Noise-Cancelling Headset", "Tech", 189.00, 6, 7),
    ("Wireless Mouse", "Tech", 29.90, 20, 10),
    ("Standing Desk", "Furniture", 549.00, 4, 5),
    ("Ergonomic Chair", "Furniture", 289.00, 6, 7),
    ("Bookshelf", "Furniture", 159.00, 4, 4),
    ("Filing Cabinet", "Furniture", 199.00, 3, 4),
    ("Copy Paper Box", "Office", 38.50, 30, 12),
    ("Printer Toner", "Office", 72.00, 10, 9),
    ("Notebook 5-Pack", "Office", 12.90, 40, 10),
]

# Nov 2025 .. Oct 2026 with a holiday peak and a summer dip
MONTHS = [(2025, 11), (2025, 12)] + [(2026, m) for m in range(1, 11)]
SEASON = {11: 1.15, 12: 1.35, 1: 0.85, 2: 0.85, 3: 0.95, 4: 1.0, 5: 1.0,
          6: 0.9, 7: 0.85, 8: 0.9, 9: 1.05, 10: 1.1}


def month_days(year: int, month: int) -> int:
    nxt = date(year + (month == 12), month % 12 + 1, 1)
    return (nxt - timedelta(days=1)).day


def main() -> None:
    rows = []
    order_no = 10001
    for i, (year, month) in enumerate(MONTHS):
        n_orders = round(190 * SEASON[month] * (1 + 0.012 * i))
        for _ in range(n_orders):
            region = random.choices(list(REGION_WEIGHTS), weights=REGION_WEIGHTS.values())[0]
            product, category, price, max_qty, _ = random.choices(
                PRODUCTS, weights=[p[4] for p in PRODUCTS])[0]
            discount = random.choice([0, 0, 0, 0.05, 0.10])
            rows.append({
                "order_id": f"SO-{order_no}",
                "order_date": date(year, month, random.randint(1, month_days(year, month))),
                "region": region,
                "sales_rep": random.choice(REPS[region]),
                "product": product,
                "category": category,
                "quantity": random.randint(1, max_qty),
                "unit_price": round(price * (1 - discount), 2),
            })
            order_no += 1
    rows.sort(key=lambda r: (r["order_date"], r["order_id"]))

    # --- make it messy -------------------------------------------------------
    out = []
    for r in rows:
        r = dict(r)
        if random.random() < 0.10:
            r["order_date"] = r["order_date"].strftime("%m/%d/%Y")
        else:
            r["order_date"] = r["order_date"].isoformat()
        if random.random() < 0.08:
            r["region"] = random.choice([r["region"].lower(), r["region"].upper(), f" {r['region']} "])
        if random.random() < 0.05:
            r["unit_price"] = f"${r['unit_price']:,.2f}"
        out.append(r)

    for r in random.sample(out, 6):
        r["quantity"] = ""
    for r in random.sample(out, 25):
        out.insert(out.index(r) + 1, dict(r))
    blank = {k: "" for k in out[0]}
    for pos in random.sample(range(len(out)), 4):
        out.insert(pos, dict(blank))

    OUT.parent.mkdir(exist_ok=True)
    with OUT.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(out[0]))
        writer.writeheader()
        writer.writerows(out)
    print(f"Wrote {len(out)} rows to {OUT}")


if __name__ == "__main__":
    main()
