"""Time the Python validator on batches from 300 to 10,000 invoices.

Each batch is generated with the same mix of planted problems, so the benchmark
also checks that every planted issue is still caught at scale.
"""
import statistics
import time

import pandas as pd

from src import config as cfg
from src.generate_batch import BatchGenerator
from src.validate import validate

SIZES = [300, 3_000, 10_000]
REPEATS = 3


def main():
    rows = []
    for n in SIZES:
        data = BatchGenerator(n_invoices=n, seed=cfg.RANDOM_SEED + n).run()
        timings = []
        for _ in range(REPEATS):
            start = time.perf_counter()
            results, exceptions = validate(data, cfg.SETTINGS)
            timings.append(time.perf_counter() - start)
        found = set(zip(exceptions["invoice_key"], exceptions["code"]))
        seeded = data["seeded"]
        caught = sum((k, c) in found for k, c in zip(seeded["invoice_key"], seeded["code"]))
        rows.append({
            "invoices": len(data["invoices"]),
            "invoice_lines": len(data["invoice_lines"]),
            "posted_history": len(data["posted"]),
            "seconds": round(statistics.median(timings), 3),
            "invoices_per_second": int(len(data["invoices"]) / statistics.median(timings)),
            "planted_issues": len(seeded),
            "planted_caught": caught,
            "held": int((results["status"] == "HOLD").sum()),
        })
        print(rows[-1])
    out = pd.DataFrame(rows)
    cfg.OUTPUT_DIR.mkdir(exist_ok=True)
    out.to_csv(cfg.OUTPUT_DIR / "benchmark.csv", index=False)
    return out


if __name__ == "__main__":
    main()
