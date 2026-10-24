"""Generate a synthetic daily invoice batch and its reference data.

What an AP officer has in front of them before posting a day's invoices:
the vendor master, open purchase order lines, goods receipts, the posted-invoice
history and the batch itself (headers + lines). No real company data is used.

Known problems are planted in about 1 in 5 invoices and recorded in
seeded_issues.csv so the validator can be checked against a ground truth.
"""
from datetime import timedelta

import numpy as np
import pandas as pd

from . import config as cfg
from .abn import random_invalid_abn, random_valid_abn

PLACES = ["Ironbark", "Redgum", "Wattle", "Banksia", "Kestrel", "Tasman", "Murray", "Granite",
          "Harbourside", "Summit", "Southbank", "Eastgate", "Westbrook", "Highfield", "Riverbend",
          "Stonebridge", "Clearwater", "Oakridge", "Goldfield", "Sandstone", "Bayview", "Northpoint",
          "Lakeside", "Kingfisher", "Lyrebird", "Acacia", "Mallee", "Jarrah", "Derwent", "Otway",
          "Flinders", "Corella", "Bellbird", "Ashgrove", "Glenrock"]

CATEGORIES = {
    # name: (vendor words, (item, typical unit price ex GST), share GST-free)
    "Packaging": (["Packaging", "Cartons", "Pallets"],
                  [("Shipping carton 600x400", 2.85), ("Stretch wrap roll", 24.50), ("Timber pallet", 18.00),
                   ("Packing tape 48mm", 3.10)], 0.0),
    "Food Ingredients": (["Foods", "Ingredients", "Grain Co"],
                         [("Rolled oats 25kg", 38.00), ("Raw sugar 25kg", 31.50), ("Canola oil 20L", 64.00),
                          ("Flour 12.5kg", 17.40)], 1.0),
    "Freight": (["Freight", "Transport", "Logistics"],
                [("Pallet freight Melb-Syd", 95.00), ("Courier satchel", 12.50), ("Linehaul per tonne", 140.00)], 0.0),
    "Maintenance": (["Maintenance", "Electrical", "Facility Services"],
                    [("Labour hour", 98.00), ("Call-out fee", 120.00), ("LED panel", 85.00)], 0.0),
    "Chemicals": (["Chemicals", "Cleaning Supplies"],
                  [("Floor cleaner 20L", 72.00), ("Sanitiser 5L", 29.00), ("Degreaser 20L", 88.00)], 0.0),
    "IT": (["IT Solutions", "Technology"],
           [("Laptop", 1450.00), ("Monitor 27in", 389.00), ("Support hour", 145.00), ("Licence per user", 32.00)], 0.0),
}

TERMS = [14, 30, 30, 30, 45, 60]


class BatchGenerator:
    def __init__(self, n_invoices: int = 300, seed: int = cfg.RANDOM_SEED, plant_issues: bool = True):
        self.rng = np.random.default_rng(seed)
        self.n_invoices = n_invoices
        self.plant = plant_issues
        self.batch_date = cfg.SETTINGS["batch_date"]
        self.seeded = []

    def pick(self, items):
        return items[int(self.rng.integers(len(items)))]

    def bank(self):
        return (f"{self.rng.integers(10, 99)}{self.rng.integers(0, 10)}-{self.rng.integers(0, 1000):03d}",
                f"{self.rng.integers(10_000_000, 99_999_999)}")

    # ---------- reference data ----------
    def build_vendors(self, n_vendors):
        rows, names = [], set()
        cats = list(CATEGORIES)
        for i in range(n_vendors):
            cat = cats[i % len(cats)]
            while True:
                name = f"{self.pick(PLACES)} {self.pick(CATEGORIES[cat][0])} Pty Ltd"
                if name not in names:
                    names.add(name)
                    break
            bsb, acct = self.bank()
            rows.append({"vendor_id": f"V{20001 + i}", "vendor_name": name, "abn": random_valid_abn(self.rng),
                         "category": cat, "gst_registered": "Y", "status": "Active",
                         "payment_terms_days": int(self.pick(TERMS)), "bsb": bsb, "account": acct})
        v = pd.DataFrame(rows)
        # a few small suppliers below the GST registration threshold
        small = self.rng.choice(v.index[v["category"].isin(["Maintenance", "Freight"])], size=max(1, n_vendors // 15),
                                replace=False)
        v.loc[small, "gst_registered"] = "N"
        self.vendors = v

    def build_pos(self, n_pos):
        v = self.vendors
        rows = []
        for i in range(n_pos):
            vend = v.iloc[int(self.rng.integers(len(v)))]
            items = CATEGORIES[vend["category"]][1]
            gst_free_share = CATEGORIES[vend["category"]][2]
            n_lines = int(self.rng.choice([1, 1, 2, 2, 3, 4]))
            po_number = f"{4500100000 + i}"
            status = "Closed" if self.rng.random() < 0.04 else "Open"
            for ln in range(1, n_lines + 1):
                item, base = items[int(self.rng.integers(len(items)))]
                price = round(base * float(np.exp(self.rng.normal(0, 0.08))), 2)
                qty = int(self.rng.choice([1, 2, 5, 10, 20, 40, 80, 120, 200])) if base < 100 else \
                    int(self.rng.choice([1, 2, 3, 4, 6, 8, 12]))
                no_gst = vend["gst_registered"] == "N" or self.rng.random() < gst_free_share
                rows.append({"po_number": po_number, "po_line": ln * 10, "vendor_id": vend["vendor_id"],
                             "item": item, "qty_ordered": qty, "unit_price_ex_gst": price,
                             "gst_code": "FRE" if no_gst else "GST", "po_status": status})
        self.po_lines = pd.DataFrame(rows)

    def build_receipts(self):
        r = self.po_lines[["po_number", "po_line", "qty_ordered"]].copy()
        u = self.rng.random(len(r))
        partial = np.maximum(1, np.floor(r["qty_ordered"] * self.rng.uniform(0.4, 0.9, len(r)))).astype(int)
        r["qty_received"] = np.where(u < 0.82, r["qty_ordered"], np.where(u < 0.96, partial, 0))
        self.receipts = r.drop(columns="qty_ordered")

    # ---------- the batch ----------
    def invoice_number(self, vendor_id):
        n = int(self.rng.integers(10_000, 99_999))
        style = int(vendor_id[-2:]) % 4
        return [f"INV-{n}", f"{n}", f"INV{n:07d}", f"TI{n}"][style]

    def build_batch(self):
        v = self.vendors.set_index("vendor_id")
        po = self.po_lines.merge(self.receipts, on=["po_number", "po_line"])
        open_pos = po[(po["po_status"] == "Open") & (po["qty_received"] > 0)]
        usable = open_pos.groupby("po_number").filter(lambda g: v.loc[g["vendor_id"].iloc[0], "status"] == "Active")
        po_numbers = usable["po_number"].unique()
        chosen = self.rng.choice(po_numbers, size=self.n_invoices, replace=len(po_numbers) < self.n_invoices)

        by_po = {po_no: grp for po_no, grp in usable.groupby("po_number")}
        headers, lines = [], []
        for k, po_number in enumerate(chosen, start=1):
            pol = by_po[po_number]
            vend_id = pol["vendor_id"].iloc[0]
            vend = v.loc[vend_id]
            inv_date = self.batch_date - timedelta(days=int(self.rng.integers(1, 21)))
            key = f"B{k:05d}"
            n_take = int(self.rng.integers(1, len(pol) + 1))
            take = pol.iloc[np.sort(self.rng.choice(len(pol), size=n_take, replace=False))]
            for ln, (_, row) in enumerate(take.iterrows(), start=1):
                qty = int(row["qty_received"]) if self.rng.random() < 0.85 else max(1, int(row["qty_received"]) // 2)
                price = row["unit_price_ex_gst"]
                line_ex = round(qty * price, 2)
                gst = round(line_ex * cfg.SETTINGS["gst_rate"], 2) if row["gst_code"] == "GST" else 0.0
                lines.append({"invoice_key": key, "line_no": ln, "po_line": int(row["po_line"]), "item": row["item"],
                              "qty": qty, "unit_price_ex_gst": price, "line_ex_gst": line_ex, "line_gst": gst})
            printed_bank = self.rng.random() < 0.6
            abn = vend["abn"] if self.rng.random() < 0.7 else vend["abn"].replace(" ", "")
            headers.append({"invoice_key": key, "vendor_id": vend_id, "invoice_number": self.invoice_number(vend_id),
                            "invoice_date": inv_date,
                            "received_date": min(self.batch_date, inv_date + timedelta(days=int(self.rng.integers(0, 4)))),
                            "abn_on_invoice": abn, "po_number": po_number,
                            "bsb_on_invoice": vend["bsb"] if printed_bank else "",
                            "account_on_invoice": vend["account"] if printed_bank else ""})
        self.headers = pd.DataFrame(headers)
        self.lines = pd.DataFrame(lines)
        self.recalc_totals()

    def recalc_totals(self, keys=None):
        sums = self.lines.groupby("invoice_key")[["line_ex_gst", "line_gst"]].sum().round(2)
        h = self.headers.set_index("invoice_key")
        idx = h.index if keys is None else pd.Index(keys)
        h.loc[idx, "total_ex_gst"] = sums.loc[idx, "line_ex_gst"]
        h.loc[idx, "gst_total"] = sums.loc[idx, "line_gst"]
        h.loc[idx, "total_inc_gst"] = (h.loc[idx, "total_ex_gst"] + h.loc[idx, "gst_total"]).round(2)
        self.headers = h.reset_index()

    def build_history(self, n_posted):
        v = self.vendors
        rows = []
        for _ in range(n_posted):
            vend_id = v["vendor_id"].iloc[int(self.rng.integers(len(v)))]
            rows.append({"vendor_id": vend_id, "invoice_number": self.invoice_number(vend_id),
                         "total_inc_gst": round(float(np.exp(self.rng.normal(7.3, 1.1))), 2),
                         "posted_date": self.batch_date - timedelta(days=int(self.rng.integers(2, 365)))})
        posted = pd.DataFrame(rows)
        # a genuine invoice number in the batch must not collide with history by chance
        batch_keys = set(zip(self.headers["vendor_id"], self.headers["invoice_number"]))
        posted = posted[[k not in batch_keys for k in zip(posted["vendor_id"], posted["invoice_number"])]]
        self.posted = posted.reset_index(drop=True)

    # ---------- planted problems ----------
    def seed(self, key, code, note):
        self.seeded.append({"invoice_key": key, "code": code, "note": note})

    def plant_issues(self, rate: float = 1.0):
        rng = self.rng
        h = self.headers.set_index("invoice_key")
        v = self.vendors.set_index("vendor_id")
        po = self.po_lines.merge(self.receipts, on=["po_number", "po_line"])
        lines = self.lines
        keys = list(h.index)
        rng.shuffle(keys)
        pool = iter(keys)

        def n(base):
            return max(1, int(round(base * rate)))

        used = set()

        def take(cond=lambda k: True):
            while True:
                k = next(pool)
                if k not in used and cond(k):
                    used.add(k)
                    return k

        positions = lines.groupby("invoice_key").indices
        line_of = lambda k: list(positions[k])

        # V01 inactive vendor: deactivate the two vendors with the fewest invoices in the batch;
        # every one of their invoices should then be held
        counts = h["vendor_id"].value_counts()
        for vend_id in counts.index[-n(2):]:
            v.at[vend_id, "status"] = "Inactive"
            for k in h.index[h["vendor_id"] == vend_id]:
                used.add(k)
                self.seed(k, "V01", "vendor deactivated in vendor master")

        # V02 ABN problems
        for style in ["missing", "missing", "checksum", "checksum", "different", "different"][: n(6)]:
            k = take()
            h.at[k, "abn_on_invoice"] = {"missing": "", "checksum": random_invalid_abn(rng),
                                         "different": random_valid_abn(rng)}[style]
            self.seed(k, "V02", f"ABN {style}")

        # V03 bank details changed on the invoice
        for _ in range(n(5)):
            k = take()
            h.at[k, "bsb_on_invoice"], h.at[k, "account_on_invoice"] = self.bank()
            self.seed(k, "V03", "invoice shows different bank account")

        # V04 duplicates: already posted, and twice in the same batch
        already_posted = []
        for _ in range(n(6)):
            k = take()
            already_posted.append({
                "vendor_id": h.at[k, "vendor_id"], "invoice_number": h.at[k, "invoice_number"],
                "total_inc_gst": h.at[k, "total_inc_gst"],
                "posted_date": self.batch_date - timedelta(days=int(rng.integers(5, 40)))})
            self.seed(k, "V04", "already posted")
        self.posted = pd.concat([self.posted, pd.DataFrame(already_posted)], ignore_index=True)
        batch_dups = []
        for _ in range(n(3)):
            k = take()
            batch_dups.append(k)
            self.seed(f"{k}-DUP", "V04", "same invoice twice in the batch")

        # V05 PO problems
        other_vendor_po = lambda vend: po.loc[po["vendor_id"] != vend, "po_number"].iloc[int(rng.integers(100))]
        closed_pos = po.loc[po["po_status"] == "Closed", "po_number"].unique()
        for style in ["none", "none", "not_found", "not_found", "other_vendor", "closed"][: n(6)]:
            k = take()
            if style == "none":
                h.at[k, "po_number"] = ""
            elif style == "not_found":
                h.at[k, "po_number"] = f"{int(h.at[k, 'po_number']) + 7_000_000}"
            elif style == "other_vendor":
                h.at[k, "po_number"] = other_vendor_po(h.at[k, "vendor_id"])
            else:
                h.at[k, "po_number"] = closed_pos[int(rng.integers(len(closed_pos)))]
            self.seed(k, "V05", f"PO {style.replace('_', ' ')}")

        # V06 price above PO
        for _ in range(n(10)):
            k = take()
            i = line_of(k)[0]
            lines.at[i, "unit_price_ex_gst"] = round(lines.at[i, "unit_price_ex_gst"] * rng.uniform(1.03, 1.15), 2)
            self.reline(i)
            self.seed(k, "V06", "unit price above PO")

        # V07 quantity above received (only where something is still to be received)
        rec = dict(zip(zip(po["po_number"], po["po_line"]), po["qty_received"]))
        ordered = dict(zip(zip(po["po_number"], po["po_line"]), po["qty_ordered"]))

        def partial_line(k):
            for i in line_of(k):
                pk = (h.at[k, "po_number"], lines.at[i, "po_line"])
                if pk in rec and rec[pk] < ordered[pk]:
                    return i
            return None
        for _ in range(n(10)):
            k = take(lambda k: partial_line(k) is not None)
            i = partial_line(k)
            pk = (h.at[k, "po_number"], lines.at[i, "po_line"])
            lines.at[i, "qty"] = int(rec[pk] + max(1, (ordered[pk] - rec[pk]) // 2))
            self.reline(i)
            self.seed(k, "V07", "billed for goods not yet received")

        # V08 arithmetic
        for _ in range(n(3)):
            k = take()
            i = line_of(k)[0]
            lines.at[i, "line_ex_gst"] = round(lines.at[i, "line_ex_gst"] + float(rng.choice([10, 50, 100])), 2)
            if lines.at[i, "line_gst"] > 0:
                lines.at[i, "line_gst"] = round(lines.at[i, "line_ex_gst"] * 0.1, 2)
            self.seed(k, "V08", "line total does not equal qty x price")
        header_errors = []
        for _ in range(n(3)):
            k = take()
            header_errors.append(k)
            self.seed(k, "V08", "invoice total does not equal sum of lines")

        # V09 GST
        # a GST error on a tiny line can fall inside the 5 cent tolerance, so use lines of $50+
        gst_line = lambda k: next((i for i in line_of(k) if lines.at[i, "line_gst"] > 0
                                   and lines.at[i, "line_ex_gst"] >= 50), None)
        free_line = lambda k: next((i for i in line_of(k) if lines.at[i, "line_gst"] == 0), None)
        for _ in range(n(3)):
            k = take(lambda k: gst_line(k) is not None)
            i = gst_line(k)
            lines.at[i, "line_gst"] = round(lines.at[i, "line_ex_gst"] * float(rng.choice([0.0, 0.11, 0.15])), 2)
            self.seed(k, "V09", "line GST not 10%")
        for _ in range(n(2)):
            k = take(lambda k: free_line(k) is not None and v.at[h.at[k, "vendor_id"], "gst_registered"] == "Y")
            i = free_line(k)
            lines.at[i, "line_gst"] = round(lines.at[i, "line_ex_gst"] * 0.1, 2)
            self.seed(k, "V09", "GST charged on GST-free item")
        unregistered = set(v.index[v["gst_registered"] == "N"])
        k = take(lambda k: h.at[k, "vendor_id"] in unregistered)
        for i in line_of(k):
            lines.at[i, "line_gst"] = round(lines.at[i, "line_ex_gst"] * 0.1, 2)
        self.seed(k, "V09", "vendor not registered for GST but charged GST")

        # V10 dates
        for style in ["future", "future", "stale", "stale"][: n(4)]:
            k = take()
            h.at[k, "invoice_date"] = self.batch_date + timedelta(days=int(rng.integers(3, 30))) if style == "future" \
                else self.batch_date - timedelta(days=int(rng.integers(380, 500)))
            h.at[k, "received_date"] = self.batch_date
            self.seed(k, "V10", f"invoice dated {'in the future' if style == 'future' else 'over 12 months ago'}")

        self.headers = h.reset_index()
        self.vendors = v.reset_index()
        self.lines = lines
        self.recalc_totals()

        # totals that don't add up are applied after the recalculation
        h = self.headers.set_index("invoice_key")
        for k in header_errors:
            h.at[k, "total_inc_gst"] = round(h.at[k, "total_inc_gst"] + float(rng.choice([100, 250, 1000])), 2)
        self.headers = h.reset_index()

        # in-batch duplicates: same invoice keyed again with the number reformatted
        dup_headers = self.headers[self.headers["invoice_key"].isin(batch_dups)].copy()
        dup_headers["invoice_number"] = [n.replace("INV-", "INV") if "INV-" in n else "0" + n
                                         for n in dup_headers["invoice_number"]]
        dup_headers["invoice_key"] = dup_headers["invoice_key"] + "-DUP"
        dup_lines = self.lines[self.lines["invoice_key"].isin(batch_dups)].copy()
        dup_lines["invoice_key"] = dup_lines["invoice_key"] + "-DUP"
        self.headers = pd.concat([self.headers, dup_headers], ignore_index=True)
        self.lines = pd.concat([self.lines, dup_lines], ignore_index=True)

    def reline(self, i):
        ln = self.lines
        ln.at[i, "line_ex_gst"] = round(ln.at[i, "qty"] * ln.at[i, "unit_price_ex_gst"], 2)
        if ln.at[i, "line_gst"] > 0:
            ln.at[i, "line_gst"] = round(ln.at[i, "line_ex_gst"] * 0.1, 2)

    def run(self):
        scale = self.n_invoices / 300
        self.build_vendors(max(60, int(60 * scale ** 0.5)))
        self.build_pos(max(500, int(500 * scale)))
        self.build_receipts()
        self.build_batch()
        self.build_history(int(2500 * scale))
        if self.plant:
            self.plant_issues(rate=scale)
        h = self.headers.sort_values("invoice_key").reset_index(drop=True)
        cols = ["invoice_key", "vendor_id", "invoice_number", "invoice_date", "received_date", "abn_on_invoice",
                "po_number", "total_ex_gst", "gst_total", "total_inc_gst", "bsb_on_invoice", "account_on_invoice"]
        self.headers = h[cols]
        self.lines = self.lines.sort_values(["invoice_key", "line_no"]).reset_index(drop=True)
        return {
            "vendors": self.vendors,
            "po_lines": self.po_lines,
            "receipts": self.receipts,
            "invoices": self.headers,
            "invoice_lines": self.lines,
            "posted": self.posted.sort_values("posted_date").reset_index(drop=True),
            "seeded": pd.DataFrame(self.seeded).sort_values("invoice_key").reset_index(drop=True),
        }


def main():
    data = BatchGenerator().run()
    cfg.SAMPLE_DIR.mkdir(parents=True, exist_ok=True)
    for name, df in data.items():
        df.to_csv(cfg.SAMPLE_DIR / f"{'seeded_issues' if name == 'seeded' else name}.csv", index=False)
    for name, df in data.items():
        print(f"  {name:<14} {len(df):>6,}")
    return data


if __name__ == "__main__":
    main()
