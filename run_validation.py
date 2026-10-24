"""Validate an invoice batch held in the Excel tool workbook.

Usage:
    python run_validation.py                                   # sample batch
    python run_validation.py --workbook my_batch.xlsx --output results.xlsx
"""
import argparse
import time
from pathlib import Path

from src import validate, workbook
from src import config as cfg


def main():
    parser = argparse.ArgumentParser(description="Pre-posting invoice validation")
    parser.add_argument("--workbook", type=Path, default=cfg.TOOL_WORKBOOK)
    parser.add_argument("--output", type=Path, default=cfg.OUTPUT_DIR / "Validation_Results.xlsx")
    args = parser.parse_args()

    data, settings = workbook.read_inputs(args.workbook)
    start = time.perf_counter()
    results, exceptions = validate.validate(data, settings)
    runtime = time.perf_counter() - start
    summary = validate.summarise(results, exceptions)
    workbook.write_results_workbook(results, exceptions, summary, settings, args.output, runtime)
    results.to_csv(args.output.with_name("results.csv"), index=False)
    exceptions.to_csv(args.output.with_name("exceptions.csv"), index=False)

    counts = results["status"].value_counts()
    print(f"{len(results)} invoices validated in {runtime:.2f}s")
    for status in ["PASS", "REVIEW", "HOLD"]:
        print(f"  {status:<7}{counts.get(status, 0):>5}")
    print(f"Report: {args.output}")


if __name__ == "__main__":
    main()
