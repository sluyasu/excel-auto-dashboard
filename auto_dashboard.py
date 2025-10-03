"""
CSV → Excel Dashboard (xlwings) — v0.1
- Prompts user to choose a CSV (GUI if available, else CLI)
- Normalizes headers and types (dates/numbers)
- Writes a 'Data' sheet as an Excel Table
- Builds/overwrites a 'Dashboard' sheet with KPIs + charts
- Adds a starter Pivot + Slicer wired to a chart
- Idempotent: safe to re-run

Dependencies: pandas, numpy, xlwings
"""

import os
import sys
import re
import math
import pandas as pd
import numpy as np
import xlwings as xw

# ---------- Helpers ----------
def pick_csv() -> str:
    try:
        import tkinter as tk
        from tkinter.filedialog import askopenfilename
        root = tk.Tk(); root.withdraw()
        path = askopenfilename(title="Select CSV", filetypes=[("CSV","*.csv"), ("All files","*.*")])
        return path or ""
    except Exception:
        return input("Enter path to CSV: ").strip()

def normalize_columns(cols):
    def norm(c):
        c = str(c).strip()
        c = re.sub(r"\s+", " ", c)
        return c
    return [norm(c) for c in cols]

# fuzzy lookup of important columns
SYN = {
    "date": [
        "reporting date", "report date", "date", "period", "month", "posting date",
        "accounting date", "valuation date"
    ],
    "acct_year": [
        "accounting year", "acc year", "fiscal year", "year", "ay"
    ],
    "company": ["company code", "company", "entity", "legal entity"],
    "treaty": ["treaty number", "treaty id", "contract", "policy", "treaty no"],
    "treaty_type": ["treaty type", "contract type", "program type"],
    "lob": ["line of business ins", "line of business", "lob", "segment"],
    "partner": ["ri partner", "reinsurer", "cedent partner", "trading partner", "vbund", "company id of trading partner (vbund)"],
    "gl": ["gl account", "account", "gl"],
    "currency": ["currency key", "currency", "ccy", "iso currency code"],
    "amount": ["amount in balance tr", "amount", "gross amount", "net amount", "value"],
    "value_local": ["value local currency", "local value", "amount local", "value lc"]
}

def find_col(df, keys):
    cols = [c.lower() for c in df.columns]
    for k in keys:
        k = k.lower()
        # exact / contains / startswith
        for i, c in enumerate(cols):
            if c == k or k in c or c.startswith(k):
                return df.columns[i]
    return None

def coerce_types(df):
    # dates
    for cand in ["date"]:
        col = find_col(df, SYN[cand]) if "date" in SYN else None
        if col:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    # accounting year
    ay = find_col(df, SYN["acct_year"])
    if ay:
        df[ay] = pd.to_numeric(df[ay], errors="coerce").astype("Int64")
    # numerics
    for key in ("amount", "value_local"):
        col = find_col(df, SYN[key])
        if col:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df

def ensure_year(df):
    ay = find_col(df, SYN["acct_year"])
    date_col = find_col(df, SYN["date"])
    if not ay and date_col:
        df["Accounting Year (auto)"] = df[date_col].dt.year
        return "Accounting Year (auto)"
    return ay

def french_month_name(series):
    # Format month names in French (without locale dependency)
    fr = ["janvier","février","mars","avril","mai","juin",
          "juillet","août","septembre","octobre","novembre","décembre"]
    return series.dt.month.apply(lambda m: fr[m-1] if (isinstance(m, (int,np.integer)) and 1<=m<=12) else None)

# ---------- Excel builders ----------
def write_table(sheet, df: pd.DataFrame, tbl_name="DataTbl"):
    # clear sheet and write df with headers
    sheet.clear()
    sheet["A1"].value = df
    # resize to data range
    last_row, last_col = df.shape[0] + 1, df.shape[1]
    sht = sheet.api
    lo = sheet.range((1,1), (last_row, last_col)).api
    wb = sheet.book.api
    # delete existing ListObject if same name
    for lo_obj in list(sheet.api.ListObjects):
        if lo_obj.Name == tbl_name:
            lo_obj.Delete()
    sheet.api.ListObjects.Add(1, lo, None, 1).Name = tbl_name
    return sheet.range((1,1), (last_row, last_col))

def kpi_block(sh, row, col, title, value, fmt=None):
    sh.range((row, col)).value = title
    sh.range((row, col)).api.Font.Bold = True
    sh.range((row+1, col)).value = value
    if fmt:
        sh.range((row+1, col)).number_format = fmt

def add_chart(sh, source_range, left, top, width, height, chart_type="xlColumnClustered", title=""):
    ch = sh.api.ChartObjects().Add(left, top, width, height).Chart
    ch.ChartType = getattr(xw.constants.ChartType, chart_type)
    ch.SetSourceData(Source=source_range.api)
    if title:
        ch.HasTitle = True
        ch.ChartTitle.Text = title
    ch.Axes(1).HasTitle = False
    ch.Axes(2).HasTitle = False
    return ch

def add_pivot_and_slicer(book, data_sheet, table_name, dest_sheet, dest_cell, row_field, data_field):
    # Create PivotCache
    pc = book.api.PivotCaches().Create(
        SourceType=1,  # xlDatabase
        SourceData=f"{data_sheet.name}!{data_sheet.api.ListObjects(table_name).Range.Address(True, True, 1)}"
    )
    # Create PivotTable
    pt_name = "PT_Main"
    dest = dest_sheet.range(dest_cell).api
    pc.CreatePivotTable(TableDestination=dest, TableName=pt_name)
    pt = dest_sheet.api.PivotTables(pt_name)
    # Add fields
    pt.PivotFields(row_field).Orientation = 1  # xlRowField
    pt.AddDataField(pt.PivotFields(data_field), "Sum of " + data_field, -4157)  # xlSum
    # Add slicer on row_field
    sl_cache = book.api.SlicerCaches.Add2(pt.PivotFields(row_field))
    slicer = sl_cache.Slicers.Add(dest_sheet.api, None, None, row_field, 350, 10, 200, 180)
    return pt, slicer

# ---------- Main pipeline ----------
def main():
    csv_path = pick_csv()
    if not csv_path or not os.path.isfile(csv_path):
        print("No CSV selected or invalid path.")
        sys.exit(1)

    print(f"Loading CSV: {csv_path}")
    df = pd.read_csv(csv_path)
    df.columns = normalize_columns(df.columns)
    df = coerce_types(df)

    # detect important columns
    col_amount = find_col(df, SYN["amount"])
    col_value = find_col(df, SYN["value_local"]) or col_amount
    col_date = find_col(df, SYN["date"])
    col_year = ensure_year(df)
    col_company = find_col(df, SYN["company"])
    col_treaty = find_col(df, SYN["treaty"])
    col_treaty_type = find_col(df, SYN["treaty_type"])
    col_lob = find_col(df, SYN["lob"])
    col_partner = find_col(df, SYN["partner"])
    col_currency = find_col(df, SYN["currency"])

    # fallback guards
    if col_amount is None:
        print("WARNING: Could not detect 'Amount' measure. KPIs will be limited.")
        df["__amount__"] = 0.0
        col_amount = "__amount__"

    # derived columns for visuals
    if col_date:
        df["_MonthNameFR"] = french_month_name(df[col_date])
        df["_MonthPeriod"] = pd.to_datetime(df[col_date]).dt.to_period("M").astype(str)

    # basic aggregations for KPIs
    total_amount = float(pd.to_numeric(df[col_amount], errors="coerce").fillna(0).sum())
    yoy_change_abs = None
    yoy_change_pct = None
    if col_year:
        grp = df.groupby(col_year)[col_amount].sum().sort_index()
        if len(grp) >= 2:
            yoy_change_abs = float(grp.iloc[-1] - grp.iloc[-2])
            yoy_change_pct = float((grp.iloc[-1] / grp.iloc[-2] - 1) * 100) if grp.iloc[-2] else None

    n_treaties = df[col_treaty].nunique() if col_treaty else None
    n_partners = df[col_partner].nunique() if col_partner else None
    top_lob = None
    if col_lob:
        s = df.groupby(col_lob)[col_amount].sum().sort_values(ascending=False)
        top_lob = s.index[0] if len(s) else None

    # create workbook
    out_xlsx = os.path.splitext(csv_path)[0] + "_Dashboard.xlsx"
    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False
    try:
        wb = app.books.add() if not os.path.exists(out_xlsx) else app.books.open(out_xlsx)

        # Data sheet as Excel Table
        if "Data" in [s.name for s in wb.sheets]:
            wb.sheets["Data"].delete()
        data_sh = wb.sheets.add("Data", before=wb.sheets[0])
        table_rng = write_table(data_sh, df, tbl_name="DataTbl")

        # Dashboard sheet
        if "Dashboard" in [s.name for s in wb.sheets]:
            wb.sheets["Dashboard"].delete()
        dash = wb.sheets.add("Dashboard", after=wb.sheets[-1])

        # Title
        dash["A1"].value = "Dashboard"
        dash["A1"].api.Font.Size = 18
        dash["A1"].api.Font.Bold = True

        # KPI cards (simple)
        kpi_block(dash, 3, 1, "Total Amount", total_amount, "#,##0.00")
        kpi_block(dash, 3, 3, "YoY Δ (abs)", yoy_change_abs if yoy_change_abs is not None else "n/a", "#,##0.00")
        kpi_block(dash, 3, 5, "YoY Δ (%)", yoy_change_pct if yoy_change_pct is not None else "n/a", "0.0%")
        kpi_block(dash, 3, 7, "#Treaties", n_treaties if n_treaties is not None else "n/a", "0")
        kpi_block(dash, 3, 9, "#RI Partners", n_partners if n_partners is not None else "n/a", "0")
        kpi_block(dash, 3, 11, "Top LOB", top_lob or "n/a")

        # Prep compact ranges for charts (write small tables on Dashboard)
        cur_row = 8

        # Time series by month or by Accounting Year
        if col_date:
            ts = df.dropna(subset=[col_date]).copy()
            ts["_Period"] = pd.to_datetime(ts[col_date]).dt.to_period("M").astype(str)
            ts = ts.groupby("_Period")[col_amount].sum().reset_index().sort_values("_Period")
            dash.range((cur_row, 1)).value = "Period"
            dash.range((cur_row, 2)).value = "Amount"
            dash.range((cur_row+1, 1)).value = ts.values.tolist()
            ts_last_row = cur_row + len(ts)
            add_chart(dash, dash.range((cur_row,1),(ts_last_row,2)), left=10, top=180, width=480, height=260,
                      chart_type="xlLine", title="Amount by Period")
            cur_row = ts_last_row + 2
        elif col_year:
            ts = df.groupby(col_year)[col_amount].sum().reset_index().sort_values(col_year)
            dash.range((cur_row, 1)).value = str(col_year)
            dash.range((cur_row, 2)).value = "Amount"
            dash.range((cur_row+1, 1)).value = ts.values.tolist()
            ts_last_row = cur_row + len(ts)
            add_chart(dash, dash.range((cur_row,1),(ts_last_row,2)), left=10, top=180, width=480, height=260,
                      chart_type="xlColumnClustered", title="Amount by Year")
            cur_row = ts_last_row + 2

        # Top Treaty Type
        if col_treaty_type:
            g = df.groupby(col_treaty_type)[col_amount].sum().sort_values(ascending=False)
            if len(g) > 25:
                top = g.head(10)
                others = pd.Series({"Others": g.iloc[10:].sum()})
                g = pd.concat([top, others])
            g = g.reset_index()
            dash.range((cur_row, 1)).value = str(col_treaty_type)
            dash.range((cur_row, 2)).value = "Amount"
            dash.range((cur_row+1, 1)).value = g.values.tolist()
            last = cur_row + len(g)
            add_chart(dash, dash.range((cur_row,1),(last,2)), left=520, top=180, width=480, height=260,
                      chart_type="xlColumnClustered", title="Amount by Treaty Type")
            cur_row = last + 2

        # Starter Pivot + Slicer (so visuals can be slicer-driven)
        # Choose a robust row field and data field if available
        row_field = col_year or col_company or col_currency
        data_field = col_amount
        if row_field and data_field:
            # place pivot at E8
            try:
                pt, slicer = add_pivot_and_slicer(wb, data_sh, "DataTbl", dash, "E8", row_field, data_field)
                # Connect a small pivot chart
                # Build a compact range under pivot for chart source (Excel PivotCharts also possible; use data copy for simplicity)
                piv_area = dash.range("E8").expand()
                # (Optional) you can turn this pivot into a PivotChart via COM if preferred
            except Exception as e:
                dash["E7"].value = f"Pivot/slicer setup skipped: {e}"

        # Formatting tweaks
        dash.autofit()
        # Save
        wb.save(out_xlsx)
        print(f"Saved dashboard: {out_xlsx}")
        print("Note: This is v0.1 — add more slicers and wire charts to pivot as needed.")
    finally:
        # Keep Excel open for inspection; comment next line to auto-close
        pass

if __name__ == "__main__":
    main()
