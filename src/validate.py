"""Pre-posting validation of a supplier invoice batch.

Each check returns rows of (invoice_key, line_no, code, action, detail). An invoice
with any HOLD row must not be posted; REVIEW rows need a person to look but don't
block posting on their own. The VBA macro in excel/InvoiceValidation.bas applies the
same checks with the same wording, so both tools produce identical results.
"""
import numpy as np
import pandas as pd

from . import config as cfg
from .abn import clean_abn, is_valid_abn

EXC_COLUMNS = ["invoice_key", "line_no", "code", "action", "detail"]
CODE_ORDER = {c: i for i, c in enumerate(cfg.CODES)}


def money(x) -> str:
    return f"{x:,.2f}"


def dmy(d) -> str:
    return pd.Timestamp(d).strftime("%d/%m/%Y")


def normalise_invoice_number(s: str) -> str:
    s = str(s).strip().upper()
    for token in ("INV", "-", " ", "/"):
        s = s.replace(token, "")
    return s.lstrip("0")


def _rows(df, code, action, detail, line_col=None):
    if df.empty:
        return pd.DataFrame(columns=EXC_COLUMNS)
    return pd.DataFrame({
        "invoice_key": df["invoice_key"].values,
        "line_no": df[line_col].values if line_col else pd.NA,
        "code": code,
        "action": action if isinstance(action, str) else np.asarray(action),
        "detail": list(detail),
    })


def validate(data: dict, settings: dict) -> tuple[pd.DataFrame, pd.DataFrame]:
    s = settings
    inv = data["invoices"].copy()
    lines = data["invoice_lines"].copy()
    vendors = data["vendors"].set_index("vendor_id")
    po_lines = data["po_lines"]
    receipts = data["receipts"]
    posted = data["posted"]
    batch_date = pd.Timestamp(s["batch_date"])

    inv["inv_order"] = np.arange(len(inv))
    inv["invoice_date"] = pd.to_datetime(inv["invoice_date"])
    inv = inv.join(vendors[["abn", "status", "gst_registered", "payment_terms_days", "bsb", "account", "vendor_name"]],
                   on="vendor_id")
    found = inv["status"].notna()
    out = []

    # V01 vendor master
    missing = inv[~found]
    out.append(_rows(missing, "V01", "HOLD", [f"Vendor {v} is not in the vendor master" for v in missing["vendor_id"]]))
    inactive = inv[found & (inv["status"] != "Active")]
    out.append(_rows(inactive, "V01", "HOLD", [f"Vendor {v} is inactive" for v in inactive["vendor_id"]]))

    # V02 ABN
    abn_digits = inv["abn_on_invoice"].map(clean_abn)
    master_digits = inv["abn"].map(clean_abn)
    no_abn = abn_digits == ""
    bad_abn = ~no_abn & ~inv["abn_on_invoice"].map(is_valid_abn)
    other_abn = ~no_abn & ~bad_abn & found & (abn_digits != master_digits)
    out.append(_rows(inv[no_abn], "V02", "HOLD", ["No ABN on the invoice"] * int(no_abn.sum())))
    out.append(_rows(inv[bad_abn], "V02", "HOLD", [f"ABN {a} fails the ATO checksum" for a in inv.loc[bad_abn, "abn_on_invoice"]]))
    out.append(_rows(inv[other_abn], "V02", "HOLD",
                     [f"ABN {a} does not match vendor master ABN {m}"
                      for a, m in zip(inv.loc[other_abn, "abn_on_invoice"], inv.loc[other_abn, "abn"])]))

    # V03 bank details printed on the invoice
    printed = inv["bsb_on_invoice"].astype(str).str.strip() != ""
    differs = found & printed & ((inv["bsb_on_invoice"] != inv["bsb"]) | (inv["account_on_invoice"] != inv["account"]))
    d = inv[differs]
    out.append(_rows(d, "V03", "HOLD",
                     [f"Bank account on invoice {b} {a} differs from vendor master {mb} {ma}; confirm with the vendor "
                      f"on a known phone number" for b, a, mb, ma in
                      zip(d["bsb_on_invoice"], d["account_on_invoice"], d["bsb"], d["account"])]))

    # V04 duplicates: against posted history, then within the batch
    inv["inv_norm"] = inv["invoice_number"].map(normalise_invoice_number)
    hist = posted.assign(inv_norm=posted["invoice_number"].map(normalise_invoice_number))
    hist = hist.sort_values("posted_date", kind="mergesort").drop_duplicates(["vendor_id", "inv_norm"])
    m = inv.merge(hist[["vendor_id", "inv_norm", "invoice_number", "posted_date"]], on=["vendor_id", "inv_norm"],
                  suffixes=("", "_posted")).sort_values("inv_order")
    out.append(_rows(m, "V04", "HOLD", [f"Already posted on {dmy(p)} as invoice {n}"
                                        for p, n in zip(m["posted_date"], m["invoice_number_posted"])]))
    first_key = inv.groupby(["vendor_id", "inv_norm"])["invoice_key"].transform("first")
    again = inv[inv["invoice_key"] != first_key]
    out.append(_rows(again, "V04", "HOLD", [f"Same invoice as {k} in this batch" for k in first_key[again.index]]))

    # V05 purchase order (header)
    po_head = po_lines.groupby("po_number").agg(po_vendor=("vendor_id", "first"), po_status=("po_status", "first"))
    inv = inv.join(po_head, on="po_number")
    no_po = inv["po_number"].astype(str).str.strip() == ""
    not_found = ~no_po & inv["po_vendor"].isna()
    other = ~no_po & ~not_found & (inv["po_vendor"] != inv["vendor_id"])
    closed = ~no_po & ~not_found & ~other & (inv["po_status"] == "Closed")
    out.append(_rows(inv[no_po], "V05", "HOLD", ["No PO number on the invoice"] * int(no_po.sum())))
    out.append(_rows(inv[not_found], "V05", "HOLD", [f"PO {p} not found" for p in inv.loc[not_found, "po_number"]]))
    out.append(_rows(inv[other], "V05", "HOLD", [f"PO {p} belongs to vendor {pv}" for p, pv in
                                                 zip(inv.loc[other, "po_number"], inv.loc[other, "po_vendor"])]))
    out.append(_rows(inv[closed], "V05", "HOLD", [f"PO {p} is closed" for p in inv.loc[closed, "po_number"]]))
    inv["po_ok"] = ~(no_po | not_found | other | closed)

    # line checks
    lines["line_order"] = np.arange(len(lines))
    lines = lines.merge(inv[["invoice_key", "po_number", "po_ok"]], on="invoice_key", how="left")
    pol = po_lines.merge(receipts, on=["po_number", "po_line"], how="left")
    lines = lines.merge(pol[["po_number", "po_line", "unit_price_ex_gst", "qty_received", "gst_code"]],
                        on=["po_number", "po_line"], how="left", suffixes=("", "_po"))
    lines = lines.sort_values("line_order")
    checked = lines["po_ok"].fillna(False).astype(bool)
    on_po = checked & lines["unit_price_ex_gst_po"].notna()

    bad_line = lines[checked & ~on_po]
    out.append(_rows(bad_line, "V05", "HOLD", [f"PO line {pl} is not on PO {p}"
                                               for pl, p in zip(bad_line["po_line"], bad_line["po_number"])], "line_no"))

    tol = s["price_tolerance"]
    over = lines[on_po & (lines["unit_price_ex_gst"] > lines["unit_price_ex_gst_po"] * (1 + tol) + 1e-9)]
    out.append(_rows(over, "V06", "HOLD",
                     [f"Unit price {money(u)} is {(u / p - 1) * 100:.1f}% above PO price {money(p)}"
                      for u, p in zip(over["unit_price_ex_gst"], over["unit_price_ex_gst_po"])], "line_no"))

    qty_over = lines[on_po & (lines["qty"] > lines["qty_received"].fillna(0))]
    out.append(_rows(qty_over, "V07", "HOLD",
                     [f"Qty {int(q)} invoiced but only {int(r)} received"
                      for q, r in zip(qty_over["qty"], qty_over["qty_received"].fillna(0))], "line_no"))

    calc = (lines["qty"] * lines["unit_price_ex_gst"]).round(2)
    arith = lines[(calc - lines["line_ex_gst"]).abs() > s["line_tolerance"] + 1e-9]
    out.append(_rows(arith, "V08", "HOLD", [f"Line total {money(t)} but qty x price = {money(c)}"
                                            for t, c in zip(arith["line_ex_gst"], calc[arith.index])], "line_no"))

    expected_gst = (lines["line_ex_gst"] * s["gst_rate"]).round(2)
    gst_wrong = lines[on_po & (lines["gst_code"] == "GST")
                      & ((lines["line_gst"] - lines["line_ex_gst"] * s["gst_rate"]).abs() > s["gst_tolerance"] + 1e-9)]
    gst_free = lines[on_po & (lines["gst_code"] == "FRE") & (lines["line_gst"] > 0.005)]
    gst_rows = pd.concat([
        _rows(gst_wrong, "V09", "HOLD", [f"Line GST {money(g)} should be {money(e)}"
                                         for g, e in zip(gst_wrong["line_gst"], expected_gst[gst_wrong.index])], "line_no")
        .assign(_o=gst_wrong["line_order"].values),
        _rows(gst_free, "V09", "HOLD", [f"GST {money(g)} charged on a GST-free item" for g in gst_free["line_gst"]],
              "line_no").assign(_o=gst_free["line_order"].values),
    ]).sort_values("_o").drop(columns="_o")
    out.append(gst_rows)

    # header totals
    sums = lines.groupby("invoice_key")[["line_ex_gst", "line_gst"]].sum()
    inv = inv.join(sums, on="invoice_key")
    tt = s["total_tolerance"] + 1e-9
    e1 = inv[(inv["line_ex_gst"].fillna(0) - inv["total_ex_gst"]).abs() > tt]
    e2 = inv[(inv["line_gst"].fillna(0) - inv["gst_total"]).abs() > tt]
    e3 = inv[(inv["total_ex_gst"] + inv["gst_total"] - inv["total_inc_gst"]).abs() > tt]
    totals = pd.concat([
        _rows(e1, "V08", "HOLD", [f"Invoice ex-GST total {money(t)} but lines add to {money(x)}"
                                  for t, x in zip(e1["total_ex_gst"], e1["line_ex_gst"].fillna(0))]).assign(_o=1),
        _rows(e2, "V08", "HOLD", [f"Invoice GST {money(t)} but lines add to {money(x)}"
                                  for t, x in zip(e2["gst_total"], e2["line_gst"].fillna(0))]).assign(_o=2),
        _rows(e3, "V08", "HOLD", [f"Invoice total {money(t)} does not equal ex-GST plus GST ({money(x + g)})"
                                  for t, x, g in zip(e3["total_inc_gst"], e3["total_ex_gst"], e3["gst_total"])]).assign(_o=3),
    ])
    out.append(totals.drop(columns="_o") if not totals.empty else totals)

    unreg = inv[found & (inv["gst_registered"] == "N") & (inv["gst_total"] > 0.005)]
    out.append(_rows(unreg, "V09", "HOLD", [f"Vendor is not registered for GST but charged {money(g)} GST"
                                            for g in unreg["gst_total"]]))

    # V10 dates
    future = inv[inv["invoice_date"] > batch_date]
    out.append(_rows(future, "V10", "HOLD", [f"Invoice dated {dmy(d)}, after the batch date" for d in future["invoice_date"]]))
    stale = inv[(batch_date - inv["invoice_date"]).dt.days > s["stale_days"]]
    out.append(_rows(stale, "V10", "REVIEW", [f"Invoice dated {dmy(d)}, more than {int(s['stale_days'])} days old"
                                              for d in stale["invoice_date"]]))

    # A01 approval
    big = inv[inv["total_inc_gst"] >= s["approval_limit"]]
    out.append(_rows(big, "A01", "REVIEW", [f"Total {money(t)} needs a second approver (limit {money(s['approval_limit'])})"
                                            for t in big["total_inc_gst"]]))

    found_rows = [o for o in out if not o.empty]
    exc = pd.concat(found_rows, ignore_index=True) if found_rows else pd.DataFrame(columns=EXC_COLUMNS)
    exc["_seq"] = np.arange(len(exc))
    order = inv.set_index("invoice_key")["inv_order"]
    exc = exc.assign(_inv=exc["invoice_key"].map(order), _code=exc["code"].map(CODE_ORDER),
                     _header=exc["line_no"].isna().astype(int))
    exc = exc.sort_values(["_inv", "_code", "_header", "_seq"], kind="mergesort")
    exc = exc[EXC_COLUMNS].reset_index(drop=True)
    exc["line_no"] = exc["line_no"].astype("Int64")

    results = build_results(inv, exc)
    return results, exc


def build_results(inv: pd.DataFrame, exc: pd.DataFrame) -> pd.DataFrame:
    grouped = exc.groupby("invoice_key", sort=False)
    codes = grouped["code"].agg(lambda c: ", ".join(sorted(set(c), key=CODE_ORDER.get)))
    details = grouped["detail"].agg("; ".join)
    hold = grouped["action"].agg(lambda a: (a == "HOLD").any())
    res = inv[["invoice_key", "vendor_id", "vendor_name", "invoice_number", "invoice_date", "total_inc_gst",
               "payment_terms_days"]].copy()
    res["due_date"] = res["invoice_date"] + pd.to_timedelta(res["payment_terms_days"], unit="D")
    held = set(hold.index[hold.astype(bool)])
    res["status"] = np.where(res["invoice_key"].isin(held), "HOLD",
                             np.where(res["invoice_key"].isin(codes.index), "REVIEW", "PASS"))
    codes, details = codes.to_dict(), details.to_dict()
    res["exception_codes"] = [codes.get(k, "") for k in res["invoice_key"]]
    res["details"] = [details.get(k, "") for k in res["invoice_key"]]
    res["vendor_name"] = res["vendor_name"].astype(object).where(res["vendor_name"].notna(), "")
    return res.drop(columns="payment_terms_days")[
        ["invoice_key", "vendor_id", "vendor_name", "invoice_number", "invoice_date", "total_inc_gst", "due_date",
         "status", "exception_codes", "details"]]


def summarise(results: pd.DataFrame, exc: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for status in ["PASS", "REVIEW", "HOLD"]:
        sel = results[results["status"] == status]
        rows.append({"item": f"Invoices {status}", "count": len(sel), "value_inc_gst": round(sel["total_inc_gst"].sum(), 2)})
    rows.append({"item": "Invoices in batch", "count": len(results), "value_inc_gst": round(results["total_inc_gst"].sum(), 2)})
    for code, desc in cfg.CODES.items():
        keys = exc.loc[exc["code"] == code, "invoice_key"].unique()
        val = results.loc[results["invoice_key"].isin(keys), "total_inc_gst"].sum()
        rows.append({"item": f"{code} {desc}", "count": len(keys), "value_inc_gst": round(val, 2)})
    return pd.DataFrame(rows)
