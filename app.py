"""DTDC Invoice Reconciliation App for GutBasket Foods.

Upload monthly DTDC invoice + Shopify orders + SKU dimensions + zone rates,
get back overcharged shipments and a dispute-ready XLSX.
"""

from __future__ import annotations

import io
import math
from dataclasses import dataclass

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="DTDC Reconciliation", layout="wide")


# ---------------------------------------------------------------------------
# File loading helpers
# ---------------------------------------------------------------------------

def read_tabular(uploaded_file) -> pd.DataFrame:
    """Read an uploaded file as a DataFrame.

    Handles XLSX files mis-named as .csv (DTDC commonly does this) and
    falls back through common encodings.
    """
    if uploaded_file is None:
        return pd.DataFrame()

    raw = uploaded_file.getvalue()
    name = uploaded_file.name.lower()

    # XLSX magic bytes: PK\x03\x04 (zip). XLS magic: \xd0\xcf\x11\xe0.
    if raw[:4] == b"PK\x03\x04" or raw[:4] == b"\xd0\xcf\x11\xe0" or name.endswith((".xlsx", ".xls")):
        return pd.read_excel(io.BytesIO(raw))

    for enc in ("utf-8", "utf-8-sig", "latin1", "cp1252"):
        try:
            return pd.read_csv(io.BytesIO(raw), encoding=enc, dtype=str, keep_default_na=False)
        except UnicodeDecodeError:
            continue
        except Exception:
            try:
                return pd.read_excel(io.BytesIO(raw))
            except Exception:
                continue

    return pd.read_csv(io.BytesIO(raw))


def strip_order_id(value: str) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    if s.startswith("#"):
        s = s[1:]
    return s.strip()


# ---------------------------------------------------------------------------
# Core logic
# ---------------------------------------------------------------------------

def normalize_sku_combo(s) -> str:
    """Sort items alphabetically, strip spaces, join with ', '."""
    if s is None:
        return ""
    text = str(s).strip()
    if not text or text.lower() == "nan":
        return ""
    items = [x.strip() for x in text.replace(", ", ",").split(",") if x.strip()]
    items = sorted(items)
    return ", ".join(items)


def build_order_sku_map(df_orders: pd.DataFrame) -> dict[str, list[str]]:
    """Group Shopify rows by order Name -> list of SKUs (with *N qty suffix)."""
    required = {"Name", "Lineitem sku", "Lineitem quantity"}
    missing = required - set(df_orders.columns)
    if missing:
        raise ValueError(f"Orders file missing columns: {sorted(missing)}")

    order_sku_map: dict[str, list[str]] = {}
    for raw_name, group in df_orders.groupby("Name", dropna=False):
        order_id = strip_order_id(raw_name)
        if not order_id:
            continue
        items: list[str] = []
        for _, row in group.iterrows():
            sku = str(row["Lineitem sku"]).strip()
            if not sku or sku.lower() == "nan":
                continue
            try:
                qty = int(float(str(row["Lineitem quantity"]).strip() or 0))
            except (ValueError, TypeError):
                qty = 1
            if qty <= 0:
                continue
            items.append(f"{sku}*{qty}" if qty > 1 else sku)
        if items:
            order_sku_map[order_id] = items
    return order_sku_map


def build_sku_lookup(df_sku: pd.DataFrame) -> dict[str, dict]:
    """Build normalized SKU combo -> dimensions/weight lookup."""
    df = df_sku.copy()
    # First column is the combo key (often unnamed)
    key_col = df.columns[0]

    # Resolve weight column: prefer 'Weight (kg)' if non-empty, else 'Weight'
    weight_kg_col = "Weight (kg)" if "Weight (kg)" in df.columns else None
    weight_col = "Weight" if "Weight" in df.columns else None
    if weight_kg_col is None and weight_col is None:
        raise ValueError("SKU dimensions file needs a 'Weight' or 'Weight (kg)' column")

    for c in ("L (cm)", "B (cm)", "H (cm)"):
        if c not in df.columns:
            raise ValueError(f"SKU dimensions file missing column: {c}")

    lookup: dict[str, dict] = {}
    for _, row in df.iterrows():
        key = normalize_sku_combo(row[key_col])
        if not key or key in lookup:
            continue

        # Pick weight: 'Weight (kg)' overrides if present and non-empty
        w = None
        if weight_kg_col:
            try:
                v = float(row[weight_kg_col])
                if v > 0:
                    w = v
            except (ValueError, TypeError):
                pass
        if w is None and weight_col:
            try:
                v = float(row[weight_col])
                if v > 0:
                    w = v
            except (ValueError, TypeError):
                pass
        if w is None:
            continue

        try:
            L = float(row["L (cm)"])
            B = float(row["B (cm)"])
            H = float(row["H (cm)"])
        except (ValueError, TypeError):
            continue

        lookup[key] = {"weight": w, "L": L, "B": B, "H": H}
    return lookup


def build_booked_map(df_booked: pd.DataFrame) -> dict[str, str]:
    if df_booked is None or df_booked.empty:
        return {}
    if "DD_CNNO" not in df_booked.columns or "DD_REFNO" not in df_booked.columns:
        return {}
    out: dict[str, str] = {}
    for _, row in df_booked.iterrows():
        cn = str(row["DD_CNNO"]).strip()
        ref = strip_order_id(row["DD_REFNO"])
        if cn and ref:
            out[cn] = ref
    return out


@dataclass
class ReconciliationResult:
    overcharged: pd.DataFrame
    unmatched: pd.DataFrame
    summary: dict
    all_rows: pd.DataFrame  # debugging / full view


def reconcile(
    df_invoice: pd.DataFrame,
    order_sku_map: dict[str, list[str]],
    sku_lookup: dict[str, dict],
    booked_map: dict[str, str] | None = None,
) -> ReconciliationResult:
    booked_map = booked_map or {}

    # Normalize invoice
    inv = df_invoice.copy()
    if "CONSIGNMENT_NO" not in inv.columns:
        raise ValueError("Invoice missing CONSIGNMENT_NO column")

    # Resolve order ID column
    if "ZCUST_REF" in inv.columns:
        inv["_order_id"] = inv["ZCUST_REF"].apply(strip_order_id)
    else:
        inv["_order_id"] = inv["CONSIGNMENT_NO"].astype(str).map(
            lambda cn: booked_map.get(str(cn).strip(), "")
        )

    rows = []
    unmatched_rows = []

    for _, r in inv.iterrows():
        try:
            charged_weight = float(str(r.get("CHARGED_WEIGHT", "")).strip() or 0)
        except (ValueError, TypeError):
            charged_weight = 0.0
        try:
            basic_freight = float(str(r.get("BASIC_FREIGHT", "")).strip() or 0)
        except (ValueError, TypeError):
            basic_freight = 0.0

        cn = str(r.get("CONSIGNMENT_NO", "")).strip()
        order_id = str(r.get("_order_id", "")).strip()

        skus = order_sku_map.get(order_id, [])
        sku_combo_key = normalize_sku_combo(", ".join(skus)) if skus else ""
        sku_display = ", ".join(skus) if skus else ""

        info = sku_lookup.get(sku_combo_key)

        if charged_weight <= 0:
            unmatched_rows.append({
                "CONSIGNMENT_NO": cn,
                "ZCUST_REF": order_id,
                "SKU": sku_display,
                "Reason": "Zero/invalid CHARGED_WEIGHT",
            })
            continue

        if not order_id:
            unmatched_rows.append({
                "CONSIGNMENT_NO": cn,
                "ZCUST_REF": "",
                "SKU": "",
                "Reason": "No order ID (missing ZCUST_REF / Booked.csv mapping)",
            })
            continue

        if not skus:
            unmatched_rows.append({
                "CONSIGNMENT_NO": cn,
                "ZCUST_REF": order_id,
                "SKU": "",
                "Reason": "Order ID not found in Shopify export",
            })
            continue

        if info is None:
            unmatched_rows.append({
                "CONSIGNMENT_NO": cn,
                "ZCUST_REF": order_id,
                "SKU": sku_display,
                "Reason": "SKU combo not in dimensions master",
            })
            continue

        actual_weight = info["weight"]
        vol_weight = (info["L"] * info["B"] * info["H"]) / 4000.0
        expected_weight = actual_weight if actual_weight <= 3 else vol_weight

        our_slab = math.ceil(expected_weight)
        dtdc_slab = math.ceil(charged_weight)
        if dtdc_slab <= 0:
            unmatched_rows.append({
                "CONSIGNMENT_NO": cn,
                "ZCUST_REF": order_id,
                "SKU": sku_display,
                "Reason": "DTDC slab is zero",
            })
            continue

        rate_per_kg = basic_freight / dtdc_slab
        actual_charge = our_slab * rate_per_kg
        difference = basic_freight - actual_charge

        rows.append({
            "CONSIGNMENT_NO": cn,
            "ZCUST_REF": order_id,
            "SKU": sku_display,
            "CHARGED_WEIGHT": charged_weight,
            "DTDC_SLAB": dtdc_slab,
            "BASIC_FREIGHT": basic_freight,
            "Rate/kg": rate_per_kg,
            "Our Weight": our_slab,
            "Actual Charge": actual_charge,
            "Difference": difference,
        })

    all_df = pd.DataFrame(rows)
    overcharged = (
        all_df[all_df["Difference"] > 0].sort_values("Difference", ascending=False).reset_index(drop=True)
        if not all_df.empty
        else pd.DataFrame(columns=[
            "CONSIGNMENT_NO", "ZCUST_REF", "SKU", "CHARGED_WEIGHT", "DTDC_SLAB",
            "BASIC_FREIGHT", "Rate/kg", "Our Weight", "Actual Charge", "Difference",
        ])
    )
    unmatched_df = pd.DataFrame(unmatched_rows)

    matched_count = len(all_df)
    total_invoice_rows = len(inv)
    summary = {
        "total_shipments": total_invoice_rows,
        "matched": matched_count,
        "unmatched": len(unmatched_df),
        "overcharged_count": len(overcharged),
        "total_overcharge": float(overcharged["Difference"].sum()) if not overcharged.empty else 0.0,
        "total_basic_freight": float(all_df["BASIC_FREIGHT"].sum()) if not all_df.empty else 0.0,
    }

    return ReconciliationResult(
        overcharged=overcharged,
        unmatched=unmatched_df,
        summary=summary,
        all_rows=all_df,
    )


# ---------------------------------------------------------------------------
# XLSX export — matches GutBasket's established template
# ---------------------------------------------------------------------------

def build_xlsx(overcharged: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Weight Disputes"

    headers = [
        "CONSIGNMENT_NO", "ZCUST_REF", "SKU", "CHARGED_WEIGHT", "DTDC_SLAB",
        "BASIC_FREIGHT", "Rate/kg", "Our Weight", "Actual Charge", "Difference",
    ]
    header_fill = PatternFill("solid", fgColor="305496")
    header_font = Font(bold=True, color="FFFFFF")

    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    fmt_map = {
        "CHARGED_WEIGHT": "0.000",
        "DTDC_SLAB": "0",
        "BASIC_FREIGHT": '"₹"#,##0.00',
        "Rate/kg": "#,##0.00",
        "Our Weight": "0",
        "Actual Charge": "#,##0.00",
        "Difference": "#,##0.00",
    }

    row_idx = 2
    for _, r in overcharged.iterrows():
        for col, h in enumerate(headers, start=1):
            cell = ws.cell(row=row_idx, column=col, value=r[h])
            if h in fmt_map:
                cell.number_format = fmt_map[h]
        row_idx += 1

    # Totals row
    if not overcharged.empty:
        total_label = ws.cell(row=row_idx, column=1, value="TOTAL")
        total_label.font = Font(bold=True)
        bf_cell = ws.cell(row=row_idx, column=6, value=float(overcharged["BASIC_FREIGHT"].sum()))
        bf_cell.number_format = fmt_map["BASIC_FREIGHT"]
        bf_cell.font = Font(bold=True)
        diff_cell = ws.cell(row=row_idx, column=10, value=float(overcharged["Difference"].sum()))
        diff_cell.number_format = fmt_map["Difference"]
        diff_cell.font = Font(bold=True)

    widths = [18, 14, 40, 14, 11, 14, 11, 11, 14, 14]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_missing_skus_xlsx(missing_combos: list[str]) -> bytes:
    """XLSX template for missing SKU combos — fill in L/B/H/Weight and paste back into master."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Missing SKUs"

    headers = ["SKU_Combo", "L (cm)", "B (cm)", "H (cm)", "Weight (kg)"]
    header_fill = PatternFill("solid", fgColor="C00000")
    header_font = Font(bold=True, color="FFFFFF")

    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for i, combo in enumerate(sorted(missing_combos), start=2):
        ws.cell(row=i, column=1, value=combo)

    widths = [50, 12, 12, 12, 14]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

def column_check(df: pd.DataFrame, required: list[str]) -> tuple[bool, list[str]]:
    if df is None or df.empty:
        return False, required
    missing = [c for c in required if c not in df.columns]
    return (len(missing) == 0), missing


st.title("DTDC Invoice Reconciliation")
st.caption("Upload monthly files and generate a weight-dispute report for GutBasket Foods.")

with st.sidebar:
    st.header("Files")
    invoice_file = st.file_uploader("DTDC Invoice (CSV/XLSX)", type=["csv", "xlsx", "xls"], key="inv")
    orders_file = st.file_uploader("Shopify Orders Export (CSV)", type=["csv", "xlsx"], key="ord")
    sku_files = st.file_uploader(
        "SKU Dimensions Master (CSV) — multiple allowed",
        type=["csv", "xlsx"],
        accept_multiple_files=True,
        key="sku",
    )
    zones_file = st.file_uploader("Zone Rates (CSV) — optional fallback", type=["csv", "xlsx"], key="zone")
    booked_file = st.file_uploader(
        "Booked.csv (optional, only if invoice has no ZCUST_REF)",
        type=["csv", "xlsx"],
        key="bkd",
    )
    run = st.button("Run Reconciliation", type="primary", use_container_width=True)

# Validation panel
status_cols = st.columns(4)
file_specs = [
    ("Invoice", invoice_file, ["CONSIGNMENT_NO", "CHARGED_WEIGHT", "BASIC_FREIGHT"]),
    ("Orders", orders_file, ["Name", "Lineitem sku", "Lineitem quantity"]),
    ("SKU Master", sku_files[0] if sku_files else None, ["L (cm)", "B (cm)", "H (cm)"]),
    ("Zone Rates", zones_file, ["Zone", "Rate"]),
]

for col, (label, f, required) in zip(status_cols, file_specs):
    with col:
        if f is None:
            st.info(f"**{label}**\nNot uploaded")
        else:
            try:
                df = read_tabular(f)
                ok, missing = column_check(df, required)
                if ok:
                    st.success(f"**{label}** ✓\n{len(df)} rows")
                else:
                    st.warning(f"**{label}** missing: {missing}")
            except Exception as e:
                st.error(f"**{label}** failed: {e}")

st.divider()

if run:
    if invoice_file is None or orders_file is None or not sku_files:
        st.error("Invoice, Orders, and SKU Dimensions are required.")
        st.stop()

    try:
        with st.spinner("Reading files..."):
            df_invoice = read_tabular(invoice_file)
            df_orders = read_tabular(orders_file)
            df_sku = pd.concat([read_tabular(f) for f in sku_files], ignore_index=True)
            df_booked = read_tabular(booked_file) if booked_file else pd.DataFrame()

        with st.spinner("Reconciling..."):
            order_map = build_order_sku_map(df_orders)
            sku_lookup = build_sku_lookup(df_sku)
            booked_map = build_booked_map(df_booked)
            result = reconcile(df_invoice, order_map, sku_lookup, booked_map)
    except Exception as e:
        st.error(f"Failed: {e}")
        st.stop()

    s = result.summary
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Total Shipments", s["total_shipments"])
    m2.metric("Matched", s["matched"])
    m3.metric("Unmatched", s["unmatched"])
    m4.metric("Overcharged", s["overcharged_count"])
    m5.metric("Total Overcharge", f"₹{s['total_overcharge']:,.2f}")

    st.subheader("Overcharged Shipments (dispute candidates)")
    if result.overcharged.empty:
        st.info("No overcharged shipments found.")
    else:
        display = result.overcharged.copy()
        st.dataframe(
            display,
            use_container_width=True,
            column_config={
                "CHARGED_WEIGHT": st.column_config.NumberColumn(format="%.3f"),
                "BASIC_FREIGHT": st.column_config.NumberColumn(format="₹%.2f"),
                "Rate/kg": st.column_config.NumberColumn(format="%.2f"),
                "Actual Charge": st.column_config.NumberColumn(format="₹%.2f"),
                "Difference": st.column_config.NumberColumn(format="₹%.2f"),
            },
        )

        xlsx_bytes = build_xlsx(result.overcharged)
        st.download_button(
            "Download Dispute XLSX",
            data=xlsx_bytes,
            file_name="DTDC_Weight_Disputes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

    if not result.unmatched.empty:
        st.subheader(f"Unmatched Shipments ({len(result.unmatched)})")
        st.caption("These could not be priced — usually means a new SKU combo to add to the dimensions master.")
        st.dataframe(result.unmatched, use_container_width=True)
        # Distinct unmatched SKU combos
        new_combos = (
            result.unmatched[result.unmatched["Reason"] == "SKU combo not in dimensions master"]["SKU"]
            .dropna()
            .unique()
            .tolist()
        )
        if new_combos:
            st.markdown(f"**SKU combos to add to master ({len(new_combos)}):**")
            missing_xlsx = build_missing_skus_xlsx(new_combos)
            st.download_button(
                "Download Missing SKUs XLSX",
                data=missing_xlsx,
                file_name="Missing_SKU_Dimensions.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            for c in sorted(new_combos):
                st.code(c, language=None)
else:
    st.info("Upload the required files in the sidebar and click **Run Reconciliation**.")
