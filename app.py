"""DTDC Invoice Reconciliation App for GutBasket Foods.

Upload monthly DTDC invoice + Shopify orders + SKU dimensions + zone rates,
get back overcharged shipments and a dispute-ready XLSX.
"""

from __future__ import annotations

import io
import math
import os
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="DTDC Reconciliation", layout="wide")


# ---------------------------------------------------------------------------
# Persistent storage
# ---------------------------------------------------------------------------
# DATA_DIR points to a Railway Volume mount in production (e.g. /data) and
# defaults to ./data for local dev. Files persist across reconciliation runs.

DATA_DIR = Path(os.environ.get("DATA_DIR", "data"))
SKU_MASTER_PATH = DATA_DIR / "sku_master.csv"
ZONE_RATES_PATH = DATA_DIR / "zone_rates.csv"


def ensure_data_dir() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)


def load_saved_csv(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    try:
        return pd.read_csv(path)
    except Exception:
        return pd.DataFrame()


def save_csv(df: pd.DataFrame, path: Path) -> None:
    ensure_data_dir()
    df.to_csv(path, index=False)


def file_status(path: Path) -> str:
    if not path.exists():
        return "no saved version"
    mtime = datetime.fromtimestamp(path.stat().st_mtime)
    try:
        rows = len(pd.read_csv(path))
    except Exception:
        rows = "?"
    return f"{rows} rows · saved {mtime:%Y-%m-%d %H:%M}"



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


def merge_sku_master(existing: pd.DataFrame, incoming: pd.DataFrame) -> pd.DataFrame:
    """Merge an incoming SKU dimensions file into the existing master.

    Dedup by normalized SKU combo key (the first column). Latest upload wins —
    this lets users correct dimensions for an existing combo by re-uploading.
    Returns a new DataFrame; caller persists it.
    """
    if incoming is None or incoming.empty:
        return existing.copy() if not existing.empty else pd.DataFrame()
    if existing is None or existing.empty:
        return incoming.copy()

    # Align columns: union, preserving existing order then any new ones
    cols = list(existing.columns)
    for c in incoming.columns:
        if c not in cols:
            cols.append(c)
    existing = existing.reindex(columns=cols)
    incoming = incoming.reindex(columns=cols)

    key_col = cols[0]
    incoming_keys = {
        normalize_sku_combo(v) for v in incoming[key_col].astype(str).tolist() if normalize_sku_combo(v)
    }

    def _key(v) -> str:
        return normalize_sku_combo(v)

    keep_existing = existing[~existing[key_col].apply(_key).isin(incoming_keys)]
    merged = pd.concat([keep_existing, incoming], ignore_index=True)
    return merged


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
    st.header("Required this month")
    invoice_file = st.file_uploader("DTDC Invoice (CSV/XLSX)", type=["csv", "xlsx", "xls"], key="inv")
    orders_file = st.file_uploader("Shopify Orders Export (CSV)", type=["csv", "xlsx"], key="ord")

    st.header("Saved master files")
    st.caption("These persist between runs. Upload only when you have updates.")
    st.markdown(f"**SKU Dimensions** — {file_status(SKU_MASTER_PATH)}")
    sku_files = st.file_uploader(
        "Upload to merge into SKU master (multiple allowed)",
        type=["csv", "xlsx"],
        accept_multiple_files=True,
        key="sku",
    )
    st.markdown(f"**Zone Rates** — {file_status(ZONE_RATES_PATH)}")
    zones_file = st.file_uploader(
        "Upload to replace Zone Rates",
        type=["csv", "xlsx"],
        key="zone",
    )

    st.header("Optional")
    booked_file = st.file_uploader(
        "Booked.csv (only if invoice has no ZCUST_REF)",
        type=["csv", "xlsx"],
        key="bkd",
    )
    run = st.button("Run Reconciliation", type="primary", use_container_width=True)

# Validation panel
status_cols = st.columns(4)


def _check_saved(label: str, path: Path, required: list[str]) -> None:
    if not path.exists():
        st.error(f"**{label}** ✗\nNo upload, no saved version")
        return
    df = load_saved_csv(path)
    ok, missing = column_check(df, required)
    if ok:
        st.success(f"**{label}** ✓ (saved)\n{len(df)} rows")
    else:
        st.warning(f"**{label}** saved but missing: {missing}")


# Invoice
with status_cols[0]:
    if invoice_file is None:
        st.info("**Invoice**\nNot uploaded")
    else:
        try:
            df = read_tabular(invoice_file)
            ok, missing = column_check(df, ["CONSIGNMENT_NO", "CHARGED_WEIGHT", "BASIC_FREIGHT"])
            if ok:
                st.success(f"**Invoice** ✓\n{len(df)} rows")
            else:
                st.warning(f"**Invoice** missing: {missing}")
        except Exception as e:
            st.error(f"**Invoice** failed: {e}")

# Orders
with status_cols[1]:
    if orders_file is None:
        st.info("**Orders**\nNot uploaded")
    else:
        try:
            df = read_tabular(orders_file)
            ok, missing = column_check(df, ["Name", "Lineitem sku", "Lineitem quantity"])
            if ok:
                st.success(f"**Orders** ✓\n{len(df)} rows")
            else:
                st.warning(f"**Orders** missing: {missing}")
        except Exception as e:
            st.error(f"**Orders** failed: {e}")

# SKU Master — uploaded preview, or saved
with status_cols[2]:
    if sku_files:
        try:
            preview = pd.concat([read_tabular(f) for f in sku_files], ignore_index=True)
            ok, missing = column_check(preview, ["L (cm)", "B (cm)", "H (cm)"])
            if ok:
                existing = load_saved_csv(SKU_MASTER_PATH)
                merged_preview = merge_sku_master(existing, preview)
                st.success(
                    f"**SKU Master** ✓ (will merge)\n"
                    f"+{len(preview)} new → {len(merged_preview)} total"
                )
            else:
                st.warning(f"**SKU Master** missing: {missing}")
        except Exception as e:
            st.error(f"**SKU Master** failed: {e}")
    else:
        _check_saved("SKU Master", SKU_MASTER_PATH, ["L (cm)", "B (cm)", "H (cm)"])

# Zone Rates
with status_cols[3]:
    if zones_file is not None:
        try:
            df = read_tabular(zones_file)
            ok, missing = column_check(df, ["Zone", "Rate"])
            if ok:
                st.success(f"**Zone Rates** ✓ (will replace)\n{len(df)} rows")
            else:
                st.warning(f"**Zone Rates** missing: {missing}")
        except Exception as e:
            st.error(f"**Zone Rates** failed: {e}")
    else:
        _check_saved("Zone Rates", ZONE_RATES_PATH, ["Zone", "Rate"])

st.divider()

if run:
    if invoice_file is None or orders_file is None:
        st.error("Invoice and Orders are required.")
        st.stop()
    if not sku_files and not SKU_MASTER_PATH.exists():
        st.error("No SKU Dimensions master saved yet. Upload one to start.")
        st.stop()

    try:
        with st.spinner("Reading files..."):
            df_invoice = read_tabular(invoice_file)
            df_orders = read_tabular(orders_file)
            df_booked = read_tabular(booked_file) if booked_file else pd.DataFrame()

            # SKU master: merge incoming uploads into saved master, persist
            saved_sku = load_saved_csv(SKU_MASTER_PATH)
            if sku_files:
                incoming_sku = pd.concat([read_tabular(f) for f in sku_files], ignore_index=True)
                df_sku = merge_sku_master(saved_sku, incoming_sku)
                save_csv(df_sku, SKU_MASTER_PATH)
                st.info(
                    f"SKU master updated: {len(saved_sku)} → {len(df_sku)} rows "
                    f"(+{len(df_sku) - len(saved_sku)} new/updated)"
                )
            else:
                df_sku = saved_sku

            # Zone rates: replace saved if uploaded
            if zones_file is not None:
                df_zones = read_tabular(zones_file)
                save_csv(df_zones, ZONE_RATES_PATH)
                st.info(f"Zone rates saved: {len(df_zones)} rows")

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
