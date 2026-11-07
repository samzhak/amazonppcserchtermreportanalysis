#!/usr/bin/env python3
"""
Amazon Sponsored Products Search Term Report Analyzer
-----------------------------------------------------
Takes the Sponsored Products "Search term" report (xlsx or csv, downloaded from
the Amazon Ads console) and produces an Excel workbook with:

  Summary          - totals, settings used, and the size of the opportunity
  Negatives        - campaign + search term pairs to negate, in priority tiers
  Word Waste       - words/phrases that keep burning spend (negative phrase ideas)
  Harvest          - converting terms worth adding as exact match keywords
  ASIN Targets     - ASIN search terms to add as product targets or negate
  Campaigns        - performance and wasted spend by campaign
  Targeting Types  - Auto vs Phrase vs Broad vs Exact vs Product targeting

Usage:
  python search_term_analyzer.py report.xlsx
  python search_term_analyzer.py report.xlsx --target-acos 0.35 --out my_analysis.xlsx

The core idea: a search term has "earned" a negative once it has spent about as
much as you can afford to pay for one order (average order value x target ACOS)
without producing an order. Fixed rules like "10 clicks, no orders" are too
aggressive when your conversion rate is ~2%.
"""
import argparse
import math
import re
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

FONT = "Arial"
AUTO_TARGETS = {"close-match", "loose-match", "substitutes", "complements"}
ASIN_RE = re.compile(r"^b0[a-z0-9]{8}$")
STOPWORDS = {
    "for", "the", "and", "with", "to", "of", "in", "on", "a", "an", "my", "your",
    "best", "new", "that", "this", "from", "at", "by", "or", "is", "are", "me",
}


# ----------------------------------------------------------------- loading ---
def load_report(path: str) -> pd.DataFrame:
    p = Path(path)
    if p.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
        df = pd.read_excel(p)
    else:
        df = pd.read_csv(p)
    df.columns = [str(c).strip() for c in df.columns]  # Amazon adds trailing spaces
    return df


def find_col(df, label, *prefixes, required=True):
    for c in df.columns:
        if any(c.lower().startswith(p.lower()) for p in prefixes):
            return c
    if required:
        sys.exit(f"Could not find a '{label}' column. Columns found: {list(df.columns)}")
    return None


def standardize(df: pd.DataFrame) -> pd.DataFrame:
    m = {
        "Campaign": find_col(df, "Campaign Name", "Campaign Name"),
        "Ad Group": find_col(df, "Ad Group Name", "Ad Group Name", required=False),
        "Targeting": find_col(df, "Targeting", "Targeting"),
        "Match Type": find_col(df, "Match Type", "Match Type"),
        "Search Term": find_col(df, "Customer Search Term", "Customer Search Term"),
        "Impressions": find_col(df, "Impressions", "Impressions"),
        "Clicks": find_col(df, "Clicks", "Clicks"),
        "Spend": find_col(df, "Spend", "Spend"),
        "Sales": find_col(df, "Sales", "7 Day Total Sales", "14 Day Total Sales"),
        "Orders": find_col(df, "Orders", "7 Day Total Orders", "14 Day Total Orders"),
    }
    out = pd.DataFrame({k: df[v] for k, v in m.items() if v})
    if "Ad Group" not in out:
        out["Ad Group"] = ""
    for c in ["Impressions", "Clicks", "Spend", "Sales", "Orders"]:
        out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0)
    out["Search Term"] = out["Search Term"].astype(str).str.strip().str.lower()
    out["Targeting"] = out["Targeting"].astype(str).str.strip()
    out["Type"] = [classify(t, mt) for t, mt in zip(out["Targeting"], out["Match Type"])]
    out["Is ASIN"] = out["Search Term"].str.match(ASIN_RE)
    return out


def classify(targeting: str, match_type) -> str:
    t = targeting.lower()
    if t in AUTO_TARGETS:
        return "Auto"
    if t.startswith(("asin=", "asin-expanded=", "category=", "keyword-group=")):
        return "Product/Category targeting"
    mt = str(match_type).strip().upper()
    return mt.title() if mt in ("EXACT", "PHRASE", "BROAD") else "Other"


# ------------------------------------------------------------- aggregation ---
def add_ratios(d: pd.DataFrame) -> pd.DataFrame:
    d = d.copy()
    d["ACOS"] = (d["Spend"] / d["Sales"]).where(d["Sales"] > 0)
    d["CVR"] = (d["Orders"] / d["Clicks"]).where(d["Clicks"] > 0)
    d["CPC"] = (d["Spend"] / d["Clicks"]).where(d["Clicks"] > 0)
    return d


def agg(df, keys):
    g = (
        df.groupby(keys, dropna=False)
        .agg(Impressions=("Impressions", "sum"), Clicks=("Clicks", "sum"),
             Spend=("Spend", "sum"), Sales=("Sales", "sum"), Orders=("Orders", "sum"))
        .reset_index()
    )
    return add_ratios(g)


def joined(series, limit=3):
    vals = sorted(set(series.astype(str)))
    s = ", ".join(vals[:limit])
    return s + (f" (+{len(vals) - limit} more)" if len(vals) > limit else "")


# ---------------------------------------------------------------- analysis ---
def analyze(df, target_acos, aov_override, watch_clicks, harvest_orders, own_asins=None):
    total = df[["Impressions", "Clicks", "Spend", "Sales", "Orders"]].sum()
    aov = aov_override or (total["Sales"] / total["Orders"] if total["Orders"] else 0)
    breakeven = aov * target_acos  # spend you can afford per order

    kw = df[~df["Is ASIN"]]

    # --- negatives: campaign + term pairs with no orders -------------------
    pairs = agg(kw, ["Campaign", "Search Term"])
    term_orders = kw.groupby("Search Term")["Orders"].sum()
    trig = kw.groupby(["Campaign", "Search Term"])["Targeting"].agg(joined).rename("Triggered By")
    kind = kw.groupby(["Campaign", "Search Term"])["Type"].agg(joined).rename("Campaign Type")
    pairs = pairs.merge(trig.reset_index()).merge(kind.reset_index())
    zero = pairs[pairs["Orders"] == 0].copy()
    zero_all = agg(df, ["Campaign", "Search Term"])
    zero_all = zero_all[zero_all["Orders"] == 0]
    zero_total_spend = zero_all["Spend"].sum()

    def tier(r):
        if r["Spend"] >= breakeven:
            return "1 - Negate now"
        if r["Spend"] >= 0.5 * breakeven:
            return "2 - Likely negate"
        if r["Clicks"] >= watch_clicks or r["Spend"] >= 0.25 * breakeven:
            return "3 - Watch"
        return None

    zero["Priority"] = zero.apply(tier, axis=1)
    neg = zero[zero["Priority"].notna()].copy()
    neg["Converts Elsewhere"] = neg["Search Term"].map(
        lambda t: "Yes - negate in this campaign only" if term_orders.get(t, 0) > 0 else "No")
    is_kw = (kw["Targeting"].str.lower() == kw["Search Term"]).groupby(
        [kw["Campaign"], kw["Search Term"]]).any().rename("is_kw").reset_index()
    neg = neg.merge(is_kw, how="left")
    neg["Suggested Action"] = neg["is_kw"].map(
        lambda x: "Targeted keyword itself: pause or cut bid" if x else "Add as negative exact")
    neg = neg.sort_values(["Priority", "Spend"], ascending=[True, False])
    neg = neg[["Priority", "Campaign", "Search Term", "Triggered By", "Campaign Type",
               "Impressions", "Clicks", "Spend", "CPC", "Converts Elsewhere", "Suggested Action"]]

    # --- word waste: n-grams that burn spend --------------------------------
    terms = agg(kw, ["Search Term"])
    rows = []
    for _, r in terms.iterrows():
        toks = [t for t in re.findall(r"[a-z0-9']+", r["Search Term"])
                if t not in STOPWORDS and len(t) > 1]
        grams = set(toks) | {f"{a} {b}" for a, b in zip(toks, toks[1:])}
        for g in grams:
            rows.append((g, r["Search Term"], r["Clicks"], r["Spend"], r["Sales"], r["Orders"]))
    gdf = pd.DataFrame(rows, columns=["Word/Phrase", "term", "Clicks", "Spend", "Sales", "Orders"])
    gdf["converting"] = (gdf["Orders"] > 0).astype(int)
    words = (gdf.groupby("Word/Phrase")
             .agg(Terms=("term", "nunique"), Clicks=("Clicks", "sum"), Spend=("Spend", "sum"),
                  Sales=("Sales", "sum"), Orders=("Orders", "sum"),
                  **{"Converting Terms": ("converting", "sum")})
             .reset_index())
    words = add_ratios(words)
    words = words[(words["Terms"] >= 3) & (words["Spend"] >= 0.3 * breakeven)
                  & ((words["Orders"] == 0) | (words["ACOS"] >= 2 * target_acos))]
    words["Verdict"] = words["Orders"].map(lambda o: "No orders" if o == 0 else "ACOS 2x+ target")
    words = words.sort_values("Spend", ascending=False)[
        ["Word/Phrase", "Verdict", "Terms", "Converting Terms", "Clicks", "Spend", "Sales", "Orders", "ACOS"]]

    # --- harvest: converting keyword terms ----------------------------------
    exact_targets = set(df[df["Type"] == "Exact"]["Targeting"].str.lower())
    src = kw[kw["Type"] != "Exact"]
    h = agg(src, ["Search Term"])
    h["Sources"] = h["Search Term"].map(src.groupby("Search Term")["Type"].agg(joined))
    h["Already An Exact Target?"] = h["Search Term"].map(lambda t: "Yes" if t in exact_targets else "No")
    ok = h[(h["Orders"] >= 1) & (h["ACOS"].notna()) & (h["ACOS"] <= target_acos)].copy()
    ok["Priority"] = ok["Orders"].map(lambda o: "1 - Harvest" if o >= harvest_orders else "2 - Test")
    ok["Suggested Bid (target ACOS)"] = (ok["Sales"] / ok["Clicks"] * target_acos).round(2)
    harvest = ok.sort_values(["Priority", "Sales"], ascending=[True, False])[
        ["Priority", "Search Term", "Sources", "Already An Exact Target?", "Clicks", "Spend", "Sales",
         "Orders", "ACOS", "CVR", "Suggested Bid (target ACOS)"]]

    # --- ASIN terms ----------------------------------------------------------
    a = agg(df[df["Is ASIN"]], ["Search Term"]).rename(columns={"Search Term": "ASIN"})
    a["ASIN"] = a["ASIN"].str.upper()
    own = {x.upper() for x in re.findall(r"(?i)b0[a-z0-9]{8}", " ".join(df["Campaign"].unique()))}
    own |= {x.upper() for x in (own_asins or [])}

    def asin_action(r):
        if r["ASIN"] in own:
            return "Your own ASIN (defend, not a new target)"
        if r["Orders"] == 0 and r["Spend"] >= 0.5 * breakeven:
            return "Negate as product target"
        if r["Orders"] >= harvest_orders and pd.notna(r["ACOS"]) and r["ACOS"] <= target_acos:
            return "Add as product target"
        if r["Orders"] >= 1 and pd.notna(r["ACOS"]) and r["ACOS"] <= target_acos:
            return "Test as product target"
        if r["Orders"] == 0 and r["Spend"] >= 0.25 * breakeven:
            return "Watch"
        return None

    a["Action"] = a.apply(asin_action, axis=1)
    asin = a[a["Action"].notna()].sort_values(["Action", "Spend"], ascending=[True, False])[
        ["Action", "ASIN", "Impressions", "Clicks", "Spend", "Sales", "Orders", "ACOS", "CVR"]]

    # --- campaigns & targeting types ----------------------------------------
    camp = agg(df, ["Campaign"])
    camp["Share of Spend"] = camp["Spend"] / total["Spend"]
    camp["Spend On Zero-Order Terms"] = camp["Campaign"].map(zero_all.groupby("Campaign")["Spend"].sum()).fillna(0)
    camp = camp.sort_values("Spend", ascending=False)[
        ["Campaign", "Spend", "Share of Spend", "Sales", "Orders", "Clicks", "ACOS", "CVR", "CPC",
         "Spend On Zero-Order Terms"]]

    types = agg(df, ["Type"]).sort_values("Spend", ascending=False)
    types["Share of Spend"] = types["Spend"] / total["Spend"]
    types = types[["Type", "Spend", "Share of Spend", "Sales", "Orders", "Clicks", "ACOS", "CVR", "CPC"]]
    auto = agg(df[df["Type"] == "Auto"], ["Targeting"]).sort_values("Spend", ascending=False)
    auto = auto[["Targeting", "Spend", "Sales", "Orders", "Clicks", "ACOS", "CVR", "CPC"]]

    return dict(total=total, aov=aov, breakeven=breakeven, zero_total_spend=zero_total_spend,
                neg=neg, words=words, harvest=harvest, asin=asin, camp=camp, types=types, auto=auto,
                aov_overridden=bool(aov_override))


# ------------------------------------------------------------------ output ---
HEAD_FILL = PatternFill("solid", start_color="1F3864")
MONEY, PCT, INT = '$#,##0.00', '0.0%', '#,##0'
MONEY_COLS = {"Spend", "Sales", "CPC", "Spend On Zero-Order Terms", "Suggested Bid (target ACOS)"}
PCT_COLS = {"ACOS", "CVR", "Share of Spend"}
INT_COLS = {"Impressions", "Clicks", "Orders", "Terms", "Converting Terms"}
WIDE = {"Campaign": 58, "Action": 38, "Search Term": 38, "Triggered By": 38, "Sources": 34, "Word/Phrase": 30,
        "Converts Elsewhere": 34, "Suggested Action": 38, "Targeting": 26, "Type": 28}


def write_table(wb, name, df, note=None):
    ws = wb.create_sheet(name)
    start = 1
    if note:
        ws.cell(1, 1, note).font = Font(name=FONT, italic=True, size=9, color="595959")
        start = 3
    cols = list(df.columns)
    for j, c in enumerate(cols, 1):
        cell = ws.cell(start, j, c)
        cell.font = Font(name=FONT, bold=True, color="FFFFFF", size=10)
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(j)].width = WIDE.get(c, 14)
    for i, row in enumerate(df.itertuples(index=False), start + 1):
        for j, v in enumerate(row, 1):
            if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
                v = None
            cell = ws.cell(i, j, v)
            cell.font = Font(name=FONT, size=10)
            c = cols[j - 1]
            if c in MONEY_COLS:
                cell.number_format = MONEY
            elif c in PCT_COLS:
                cell.number_format = PCT
            elif c in INT_COLS:
                cell.number_format = INT
    ws.freeze_panes = ws.cell(start + 1, 1)
    if len(df):
        ws.auto_filter.ref = f"A{start}:{get_column_letter(len(cols))}{start + len(df)}"
    return ws, {c: get_column_letter(j) for j, c in enumerate(cols, 1)}, start


def build_workbook(res, args, out_path):
    wb = Workbook()
    wb.remove(wb.active)
    summ = wb.create_sheet("Summary")

    neg_ws, neg_c, neg_s = write_table(
        wb, "Negatives", res["neg"],
        "Zero-order search terms that have spent a meaningful share of one order's break-even spend. "
        "Add as negative exact in the named campaign.")
    write_table(wb, "Word Waste", res["words"],
                "Words/phrases (3+ distinct terms) with no orders or very high ACOS. Check 'Converting Terms' before "
                "adding a negative phrase.")
    har_ws, har_c, har_s = write_table(
        wb, "Harvest", res["harvest"],
        "Search terms from auto/phrase/broad/product targeting that converted at or under target ACOS. "
        "'Already An Exact Target?' only sees exact keywords that appear in this report.")
    write_table(wb, "ASIN Targets", res["asin"],
                "ASINs that appeared as search terms. Add winners as product targets; negate wasteful ones.")
    write_table(wb, "Campaigns", res["camp"])
    ty_ws, _, _ = write_table(wb, "Targeting Types", res["types"])
    # auto breakdown under the types table
    r0 = len(res["types"]) + 5
    ty_ws.cell(r0 - 1, 1, "Auto campaigns by targeting group").font = Font(name=FONT, bold=True, size=10)
    for j, c in enumerate(res["auto"].columns, 1):
        cell = ty_ws.cell(r0, j, c)
        cell.font = Font(name=FONT, bold=True, color="FFFFFF", size=10)
        cell.fill = HEAD_FILL
    for i, row in enumerate(res["auto"].itertuples(index=False), r0 + 1):
        for j, v in enumerate(row, 1):
            if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
                v = None
            cell = ty_ws.cell(i, j, v)
            cell.font = Font(name=FONT, size=10)
            c = res["auto"].columns[j - 1]
            cell.number_format = MONEY if c in MONEY_COLS else PCT if c in PCT_COLS else INT if c in INT_COLS else "General"

    # ---- Summary sheet ------------------------------------------------------
    def put(r, label, val, fmt=None, bold=False, blue=False):
        a = summ.cell(r, 1, label)
        a.font = Font(name=FONT, size=10, bold=bold)
        b = summ.cell(r, 2, val)
        b.font = Font(name=FONT, size=10, bold=bold, color="0000FF" if blue else "000000")
        if fmt:
            b.number_format = fmt
        return b

    def head(r, text):
        c = summ.cell(r, 1, text)
        c.font = Font(name=FONT, bold=True, color="FFFFFF", size=10)
        c.fill = HEAD_FILL
        summ.cell(r, 2).fill = HEAD_FILL

    summ.column_dimensions["A"].width = 52
    summ.column_dimensions["B"].width = 18
    summ.column_dimensions["C"].width = 90
    summ.cell(1, 1, "Search Term Report Analysis").font = Font(name=FONT, bold=True, size=14)
    summ.cell(2, 1, f"Source: {Path(args.report).name}").font = Font(name=FONT, size=9, color="595959")

    t = res["total"]
    head(4, "Account totals")
    put(5, "Spend", float(t["Spend"]), MONEY)
    put(6, "Sales (attributed)", float(t["Sales"]), MONEY)
    put(7, "Orders", int(t["Orders"]), INT)
    put(8, "Clicks", int(t["Clicks"]), INT)
    put(9, "Impressions", int(t["Impressions"]), INT)
    put(10, "ACOS", "=IF(B6>0,B5/B6,0)", PCT, bold=True)
    put(11, "Conversion rate (orders / clicks)", "=IF(B8>0,B7/B8,0)", PCT)
    put(12, "Average CPC", "=IF(B8>0,B5/B8,0)", MONEY)

    head(14, "Settings used (blue = input to the script; re-run the script to change the lists)")
    put(15, "Target ACOS", args.target_acos, PCT, blue=True)
    if res["aov_overridden"]:
        put(16, "Average order value", res["aov"], MONEY, blue=True)
    else:
        put(16, "Average order value (sales / orders)", "=IF(B7>0,B6/B7,0)", MONEY)
    put(17, "Break-even spend per order (AOV x target ACOS)", "=B16*B15", MONEY, bold=True)
    summ.cell(17, 3, "Spend you can afford per order. A zero-order term that has spent this much has had a fair chance.").font = Font(name=FONT, size=9, color="595959")
    put(18, "'Watch' tier: clicks without an order", args.watch_clicks, INT, blue=True)
    put(19, "'Harvest' tier: minimum orders", args.harvest_orders, INT, blue=True)

    head(21, "Opportunity")
    nl, nc = neg_c, neg_s
    rng = lambda col: f"Negatives!${neg_c[col]}${neg_s + 1}:${neg_c[col]}${neg_s + max(len(res['neg']), 1)}"
    put(22, "Spend on all zero-order campaign/term pairs (incl. ASINs)", float(res["zero_total_spend"]), MONEY)
    put(23, "  as % of total spend", "=IF(B5>0,B22/B5,0)", PCT)
    summ.cell(22, 3, "Mostly a long tail of terms with 1-3 clicks. Not all of it is recoverable; use the tiers below.").font = Font(name=FONT, size=9, color="595959")
    for r, pr in [(24, "1 - Negate now"), (25, "2 - Likely negate"), (26, "3 - Watch")]:
        put(r, f"Spend in tier: {pr}", f'=SUMIFS({rng("Spend")},{rng("Priority")},"{pr}")', MONEY, bold=(r == 24))
        put(r + 3, f"  # terms in tier: {pr}", f'=COUNTIF({rng("Priority")},"{pr}")', INT)
    hrng = lambda col: f"Harvest!${har_c[col]}${har_s + 1}:${har_c[col]}${har_s + max(len(res['harvest']), 1)}"
    put(30, "Harvest candidates (1 - Harvest tier)", f'=COUNTIF({hrng("Priority")},"1 - Harvest")', INT)
    put(31, "  sales from those terms so far", f'=SUMIFS({hrng("Sales")},{hrng("Priority")},"1 - Harvest")', MONEY)

    head(33, "Read this before acting")
    notes = [
        "Amazon's 7-day attribution window means sales from the last ~7 days of the report period are incomplete; recent zero-order terms may still convert.",
        "Sales include 'other SKU' (halo) sales, which flatters ACOS slightly.",
        "Break-even spend uses account-wide average order value. If your ASINs have very different prices or margins, set --aov or review per ASIN.",
        "Negatives are suggested at campaign level as negative exact. Check 'Converts Elsewhere' first.",
        "Thresholds are rules of thumb, not statistics. At low conversion rates a few clicks tell you very little.",
    ]
    for i, n in enumerate(notes, 34):
        c = summ.cell(i, 1, f"{i - 33}. {n}")
        c.font = Font(name=FONT, size=10)
    wb.move_sheet("Summary", offset=-wb.index(wb["Summary"]))
    wb.save(out_path)


def main():
    ap = argparse.ArgumentParser(description="Analyze an Amazon Sponsored Products search term report.")
    ap.add_argument("report", help="Search term report (.xlsx or .csv)")
    ap.add_argument("--target-acos", type=float, default=0.30, help="Target ACOS as a fraction (default 0.30)")
    ap.add_argument("--aov", type=float, default=None, help="Average order value; default = sales / orders")
    ap.add_argument("--watch-clicks", type=int, default=15, help="Clicks with no order to put a term on the watch list")
    ap.add_argument("--harvest-orders", type=int, default=2, help="Orders needed for the 'Harvest' tier")
    ap.add_argument("--own-asins", default="", help="Comma-separated ASINs you advertise (also auto-detected from campaign names)")
    ap.add_argument("--out", default=None, help="Output .xlsx path")
    args = ap.parse_args()

    df = standardize(load_report(args.report))
    res = analyze(df, args.target_acos, args.aov, args.watch_clicks, args.harvest_orders,
                  [x for x in args.own_asins.split(",") if x])
    out = args.out or f"{Path(args.report).stem}_analysis.xlsx"
    build_workbook(res, args, out)
    print(f"Wrote {out}")
    print(f"Spend ${res['total']['Spend']:,.2f} | Sales ${res['total']['Sales']:,.2f} | "
          f"Break-even spend per order ${res['breakeven']:,.2f}")
    print(res["neg"]["Priority"].value_counts().sort_index().to_string())


if __name__ == "__main__":
    main()
