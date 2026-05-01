# DTDC Invoice Reconciliation App — Build Spec for Claude Code

## Context

GutBasket Foods Pvt. Ltd. is a D2C gut health brand based in Gandhinagar, Gujarat. They ship via DTDC courier and need to reconcile monthly invoices to catch weight/rate overcharges. Currently this is done manually with spreadsheets. Build a web app so any team member can upload files and get the reconciliation output.

---

## What the App Does

Users upload 4–5 files, the app processes them, and outputs a downloadable XLSX showing only overcharged shipments with the dispute amount.

---

## Required File Uploads (UI should have labeled upload slots)

### 1. DTDC Invoice (Required)
- **Format**: CSV or XLSX (DTDC often sends XLSX files with .csv extension — auto-detect)
- **Key columns used**:
  - `CONSIGNMENT_NO` — DTDC tracking number
  - `ZCUST_REF` — Customer reference / order ID (e.g., GB720085)
  - `DESTINATION_CATEGORY` — Zone code (e.g., 01_LOCAL, 04_ROI_TO_METRO)
  - `CHARGED_WEIGHT` — Weight in kg that DTDC charged for
  - `BASIC_FREIGHT` — Freight amount charged (₹)
  - `COD_FOD_CHARGE` — COD charge
  - `SUB_TOTAL` — Total charge excl. GST
  - `BOOKING_DATE` — Date of booking
  - `DELIVERY_CITY` — Destination city

### 2. Shopify Orders Export (Required)
- **Format**: CSV
- **Key columns used**:
  - `Name` — Order ID (e.g., GB720085, #GB720085). Strip any `#` prefix.
  - `Lineitem sku` — SKU code
  - `Lineitem quantity` — Quantity of that SKU in the order
  - `Lineitem name` — Product name (for display)
- **Important**: One order can have multiple rows (one per line item). These must be grouped.

### 3. SKU Dimensions Master (Required)
- **Format**: CSV
- **Structure**: First column is the SKU combo key (may be named `Unnamed: 0` or blank header). Remaining columns:
  - `L (cm)` — Length in cm
  - `B (cm)` — Breadth in cm
  - `H (cm)` — Height in cm
  - `Weight` or `Weight (kg)` — Actual weight in kg
- **SKU combo key format**: Comma-separated SKU codes. For multi-quantity items, suffix with `*N` (e.g., `GBSproutsJar-1*2`). Single items have no suffix.
- **Examples**:
  - `Kanji401` (single SKU order)
  - `GBSproutsJar-1, Kanji201` (multi-item order)
  - `GBSproutsJar-1*2` (same item × 2)

### 4. Zone Rates (Required)
- **Format**: CSV with two columns:
  - `Zone` — Zone name (e.g., `01_LOCAL`, `02_WITHIN_STATE`, `03_WITHIN_ZONE`, `04_ROI_TO_METRO`, `05_ROIA`, `06_ROIB`, `07_SPL_DEST`)
  - `Rate` — Rate per kg in ₹

### 5. Booked.csv (Optional)
- Only needed if DTDC invoice doesn't have `ZCUST_REF` column (some PDF-extracted invoices)
- Maps `DD_CNNO` (tracking number) → `DD_REFNO` (order ID)

---

## Core Business Logic

### Step 1: Build Order → SKU Combo Mapping

From Shopify orders export:

```python
order_sku_map = {}
for order_name, group in df_orders.groupby('Name'):
    items = []
    for _, row in group.iterrows():
        sku = str(row['Lineitem sku']).strip()
        qty = int(row['Lineitem quantity'])
        if not sku or sku == 'nan':
            continue
        if qty > 1:
            items.append(f"{sku}*{qty}")
        else:
            items.append(sku)
    order_sku_map[order_name] = items
```

### Step 2: Normalize SKU Combo Keys

Both the SKU dimensions master keys AND the order-derived combos must be normalized identically for matching:

```python
def normalize_sku_combo(s):
    """Sort items alphabetically, strip spaces, join with ', '"""
    if not s:
        return ''
    items = [x.strip() for x in str(s).replace(', ', ',').split(',') if x.strip()]
    items = sorted(items)
    return ', '.join(items)
```

Build lookup from dimensions master:
```python
sku_lookup = {}
for _, row in df_sku.iterrows():
    key = normalize_sku_combo(row['SKU_Combo'])  # first column
    if key and key not in sku_lookup:  # first occurrence wins
        sku_lookup[key] = {
            'weight': float(row['Weight']),
            'L': float(row['L (cm)']),
            'B': float(row['B (cm)']),
            'H': float(row['H (cm)']),
        }
```

### Step 3: Weight Calculation Rules

```python
import math

# For each invoice row, after matching to SKU:
actual_weight = sku_info['weight']  # from dimensions master
volumetric_weight = (sku_info['L'] * sku_info['B'] * sku_info['H']) / 4000

# RULE: If actual weight ≤ 3 kg → use actual weight
#        If actual weight > 3 kg → use volumetric weight
if actual_weight <= 3:
    expected_weight = actual_weight
else:
    expected_weight = volumetric_weight

# SLAB: 1 kg ceiling (round up to next whole kg)
our_slab = math.ceil(expected_weight)        # e.g., 1.3 → 2, 2.0 → 2
dtdc_slab = math.ceil(charged_weight)        # DTDC's charged weight, same ceiling
```

### Step 4: Freight Calculation (v2 — Derived Rate Method)

This is the preferred method. Instead of using zone rate card, derive the rate from DTDC's own charges:

```python
# Derive rate per kg from DTDC's own invoice
rate_per_kg = basic_freight / dtdc_slab  # what DTDC charged per kg

# Apply that same rate to our correct slab
actual_charge = our_slab * rate_per_kg

# Difference = what they charged minus what they should have
difference = basic_freight - actual_charge
```

**Why this method**: Uses DTDC's own rate against them. They can't dispute the rate — only the weight slab. Makes the dispute much stronger.

### Step 5: Filter and Output

- **Only include rows where `difference > 0`** (overcharged shipments)
- **Sort by `difference` descending** (largest overcharges first)
- Add totals row at the bottom

---

## Output XLSX Format

Single sheet with these columns (this exact format matches GutBasket's established template):

| Column | Header | Format | Description |
|--------|--------|--------|-------------|
| A | CONSIGNMENT_NO | Text | DTDC tracking number |
| B | ZCUST_REF | Text | GutBasket order ID |
| C | SKU | Text | Comma-separated SKU combo |
| D | CHARGED_WEIGHT | Number (0.000) | DTDC's raw charged weight in kg |
| E | DTDC_SLAB | Number (0) | ceil(CHARGED_WEIGHT) |
| F | BASIC_FREIGHT | Currency (₹#,##0.00) | What DTDC charged |
| G | Rate/kg | Currency (#,##0.00) | BASIC_FREIGHT / DTDC_SLAB |
| H | Our Weight | Number (0) | Our slab = ceil(expected_weight) |
| I | Actual Charge | Currency (#,##0.00) | Our Weight × Rate/kg |
| J | Difference | Currency (#,##0.00) | BASIC_FREIGHT − Actual Charge |

**Last row**: Totals for columns F (BASIC_FREIGHT) and J (Difference)

---

## UI Requirements

### Upload Page
- 4 labeled file upload slots (Invoice, Orders Export, SKU Dimensions, Zone Rates)
- Optional 5th slot for Booked.csv
- "Run Reconciliation" button
- Show file validation status (correct columns detected or not)

### Results Page
- **Summary card at top**:
  - Total shipments in invoice
  - SKU matched count / total
  - Overcharged shipment count
  - Total overcharge amount (₹)
  - Unmatched SKU count
- **Overcharged shipments table** (sortable, filterable)
- **Unmatched SKUs list** (so user knows which combos to add to the master)
- **Download XLSX button** — generates the file in the format above

### Nice to Have
- Save zone rates and SKU dimensions so they don't need re-upload every month
- Show month-over-month trend if historical data exists
- Drag-and-drop file upload

---

## Edge Cases to Handle

1. **File type detection**: DTDC CSVs are often XLSX. Check magic bytes or try both parsers.
2. **Encoding**: Try UTF-8 first, fall back to latin1/cp1252.
3. **Order ID format**: Shopify exports may have `#` prefix (e.g., `#GB720085`). Strip it. DTDC invoice uses `GB720085`.
4. **Missing ZCUST_REF**: If invoice doesn't have this column, use Booked.csv to map CONSIGNMENT_NO → order ID.
5. **SKU combo not in master**: Mark as unmatched, show in separate section. Don't include in overcharge calculations.
6. **Zero charged weight**: Skip row (division by zero in rate calculation).
7. **Multiple SKU dimension files**: Allow uploading more than one and merge them.
8. **Duplicate SKU combos in master**: Use first occurrence.

---

## Tech Stack Suggestion

- **Frontend**: React or Next.js with Tailwind
- **Backend**: Python (FastAPI or Flask) — pandas for data processing, openpyxl for XLSX generation
- **Hosting**: Vercel (frontend) + Railway/Render (Python backend), or single Streamlit app for simplicity
- **No database needed** for MVP — all processing is file-in, file-out

For the simplest MVP, a **Streamlit app** hosted on Streamlit Cloud (free) would work. Single Python file, file uploaders built in, table display built in, download button built in.

---

## Sample Zone Rates (current)

```csv
Zone,Rate
01_LOCAL,36
02_WITHIN_STATE,36
03_WITHIN_ZONE,42
04_ROI_TO_METRO,50
05_ROIA,67
06_ROIB,67
07_SPL_DEST,67
```

---

## Validation Checklist

After building, test with these checks:
- [ ] XLSX disguised as CSV loads correctly
- [ ] Multi-item orders (2+ SKUs) match correctly
- [ ] Quantity suffix (*2, *3) handled in SKU matching
- [ ] Weight rule: 2.5kg product uses actual weight (slab = 3)
- [ ] Weight rule: 3.5kg product uses volumetric weight
- [ ] Exactly 3.0kg product uses actual weight (≤3 rule)
- [ ] ceil(1.0) = 1, ceil(1.001) = 2, ceil(2.5) = 3
- [ ] Only overcharged rows appear in output
- [ ] Totals row sums correctly
- [ ] Unmatched SKUs shown separately, not included in overcharge total
