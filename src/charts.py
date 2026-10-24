"""README charts: batch outcome, exceptions by check, and the scale benchmark."""
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import pandas as pd

from . import config as cfg

NAVY, DARK, GREY = "#1F4E79", "#2E3A46", "#8A949E"
STATUS = {"PASS": "#4E8F3A", "REVIEW": "#D68910", "HOLD": "#B03A2E"}

plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 10, "axes.titlesize": 13, "axes.titleweight": "bold",
    "axes.titlelocation": "left", "axes.spines.top": False, "axes.spines.right": False,
    "axes.edgecolor": GREY, "axes.labelcolor": DARK, "xtick.color": DARK, "ytick.color": DARK,
    "figure.dpi": 150, "savefig.bbox": "tight",
})


def batch_outcome(results: pd.DataFrame, path):
    order = ["PASS", "REVIEW", "HOLD"]
    counts = results["status"].value_counts().reindex(order, fill_value=0)
    values = results.groupby("status")["total_inc_gst"].sum().reindex(order, fill_value=0)
    fig, ax = plt.subplots(figsize=(10, 2.6))
    left = 0
    total = counts.sum()
    for status in order:
        n = counts[status]
        ax.barh([0], [n], left=left, color=STATUS[status], height=0.55)
        ax.text(left + n / 2, 0, f"{status}\n{n} invoices\n${values[status] / 1000:,.0f}k", ha="center", va="center",
                color="white", fontsize=10, fontweight="bold")
        left += n
    ax.set_xlim(0, total)
    ax.set_yticks([])
    ax.spines["left"].set_visible(False)
    ax.set_xlabel("Invoices in the batch")
    ax.set_title(f"Sample batch of {total} invoices: {counts['HOLD']} held before posting")
    fig.savefig(path)
    plt.close(fig)


def exceptions_by_check(exceptions: pd.DataFrame, results: pd.DataFrame, path):
    per_code = exceptions.groupby("code")["invoice_key"].nunique().reindex(list(cfg.CODES), fill_value=0)
    df = per_code[per_code > 0].iloc[::-1]
    labels = [f"{c}  {cfg.CODES[c]}" for c in df.index]
    colours = [STATUS["REVIEW"] if c == "A01" else STATUS["HOLD"] for c in df.index]
    fig, ax = plt.subplots(figsize=(10, 4.8))
    bars = ax.barh(labels, df.values, color=colours)
    for bar, n in zip(bars, df.values):
        ax.text(bar.get_width() + 0.4, bar.get_y() + bar.get_height() / 2, str(n), va="center", fontsize=9, color=DARK)
    ax.set_xlim(0, df.max() * 1.12)
    ax.set_xlabel("Invoices flagged")
    ax.set_title("Invoices flagged by each check")
    handles = [plt.Rectangle((0, 0), 1, 1, color=STATUS[s]) for s in ("HOLD", "REVIEW")]
    ax.legend(handles, ["Blocks posting (HOLD)", "Needs a check (REVIEW)"], frameon=False, loc="upper right")
    fig.savefig(path)
    plt.close(fig)


def benchmark_chart(bench: pd.DataFrame, path):
    fig, ax = plt.subplots(figsize=(10, 4))
    ax.plot(bench["invoices"], bench["seconds"], marker="o", color=NAVY, lw=2)
    for _, row in bench.iterrows():
        ax.annotate(f"{row['seconds']:.2f} s", (row["invoices"], row["seconds"]), textcoords="offset points",
                    xytext=(0, 9), ha="center", fontsize=9, color=DARK)
    ax.set_xticks(bench["invoices"])
    ax.xaxis.set_major_formatter(mtick.FuncFormatter(lambda x, _: f"{int(x):,}"))
    ax.set_xlim(0, bench["invoices"].max() * 1.08)
    ax.set_ylim(0, bench["seconds"].max() * 1.3)
    ax.set_xlabel("Invoices in batch")
    ax.set_ylabel("Seconds")
    ax.set_title("Python validation time by batch size")
    fig.savefig(path)
    plt.close(fig)


def results_preview(results: pd.DataFrame, path, n_per_code: int = 1):
    """A picture of the Results sheet: one PASS, one REVIEW and one HOLD invoice per check."""
    picks = [results[results["status"] == "PASS"].head(2), results[results["exception_codes"] == "A01"].head(1)]
    for code in [c for c in cfg.CODES if c != "A01"]:
        picks.append(results[results["exception_codes"].str.startswith(code)].head(n_per_code))
    df = pd.concat(picks).drop_duplicates("invoice_key")
    rows = []
    for r in df.itertuples():
        detail = r.details if len(r.details) <= 80 else r.details[:77] + "..."
        rows.append([r.invoice_key, r.vendor_name.replace(" Pty Ltd", ""), f"{r.total_inc_gst:,.2f}", r.status,
                     r.exception_codes, detail])
    headers = ["Invoice", "Vendor", "Amount", "Status", "Codes", "Details"]
    fig, ax = plt.subplots(figsize=(13, 0.3 * (len(rows) + 1)))
    ax.axis("off")
    table = ax.table(cellText=rows, colLabels=headers, loc="center", cellLoc="left",
                     colWidths=[0.095, 0.17, 0.08, 0.065, 0.07, 0.52])
    table.auto_set_font_size(False)
    table.set_fontsize(8.5)
    table.scale(1, 1.35)
    fills = {"PASS": "#E2EFDA", "REVIEW": "#FFF2CC", "HOLD": "#F8CBAD"}
    for (row, col), cell in table.get_celld().items():
        cell.set_edgecolor("#D0D5DA")
        cell.PAD = 0.02 if col == 5 else 0.06
        if row == 0:
            cell.set_facecolor(NAVY)
            cell.set_text_props(color="white", fontweight="bold")
        elif col == 3:
            cell.set_facecolor(fills[rows[row - 1][3]])
        if col == 2 and row > 0:
            cell.set_text_props(ha="right")
            cell._loc = "right"
    ax.set_title("Results sheet (extract)", loc="left", fontsize=12, fontweight="bold", pad=6)
    fig.savefig(path)
    plt.close(fig)
