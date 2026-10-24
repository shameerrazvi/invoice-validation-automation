"""Redraw the README charts from the sample batch and the latest benchmark."""
import pandas as pd

from src import charts, validate, workbook
from src import config as cfg


def main():
    data, settings = workbook.read_inputs(cfg.TOOL_WORKBOOK)
    results, exceptions = validate.validate(data, settings)
    cfg.IMAGES_DIR.mkdir(exist_ok=True)
    charts.batch_outcome(results, cfg.IMAGES_DIR / "batch_outcome.png")
    charts.exceptions_by_check(exceptions, results, cfg.IMAGES_DIR / "exceptions_by_check.png")
    charts.results_preview(results, cfg.IMAGES_DIR / "results_preview.png")
    bench = cfg.OUTPUT_DIR / "benchmark.csv"
    if bench.exists():
        charts.benchmark_chart(pd.read_csv(bench), cfg.IMAGES_DIR / "benchmark.png")


if __name__ == "__main__":
    main()
