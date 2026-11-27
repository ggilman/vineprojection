"""Write tests/fixtures/parity.json: random orders plus the Python app's projected
totals for them, so the extension's JavaScript math can be checked against it.

Run from the repo root: python extension/tests/make_parity_fixture.py
"""

import json
import random
import sys
from pathlib import Path

import pandas as pd

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO))
import app  # noqa: E402

random.seed(7)
base = pd.Timestamp("2026-11-26")
orders = [
    {
        "date": (base - pd.Timedelta(days=random.randint(0, 45))).strftime("%Y-%m-%d"),
        "cents": random.randint(500, 60000),
    }
    for _ in range(120)
]
df = pd.DataFrame(
    {
        "Order Date": [pd.Timestamp(o["date"]) for o in orders],
        "Shipped Date": [None] * len(orders),
        "Estimated Tax Value": [o["cents"] / 100 for o in orders],
        "Order Type": ["ORDER"] * len(orders),
        "Cancelled Date": [None] * len(orders),
    }
)
df = app.apply_date_basis(df, app.DATE_BASIS_ORDER)
proj, _ = app.generate_decay_projection(df, base, 30, 6500.0, 35)
out = {
    "base": "2026-11-26",
    "window": 30,
    "horizon": 35,
    "orders": orders,
    "totals": [round(v * 100) for v in proj["Projected Total"]],
}
dest = Path(__file__).parent / "fixtures" / "parity.json"
dest.parent.mkdir(exist_ok=True)
dest.write_text(json.dumps(out), encoding="utf-8")
print(f"wrote {dest}")
