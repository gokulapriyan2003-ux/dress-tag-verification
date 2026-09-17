# Technosport Tag Verification Engine - Domain Rules & Permanent Memory

This document defines the permanent project rules, domain-specific taxonomy, and edge-case solutions learned across engine iterations (v2.7 through v3.1).

---

## 1. Technosport SKU Taxonomy & Gender Mapping
Technosport SKUs follow strict prefix conventions. The verifier must ALWAYS maintain gender consistency:

| Prefix | Meaning | Gender Category | Tag Category |
| :--- | :--- | :--- | :--- |
| `BT` | Boys Topwear (e.g., `BTB208...`) | `BOYS` / `KIDS` | `Kid's` / `Kids Clothing` |
| `BP` | Boys Bottomwear / Pants | `BOYS` / `KIDS` | `Kid's` / `Kids Clothing` |
| `BS` | Boys Shorts | `BOYS` / `KIDS` | `Kid's` / `Kids Clothing` |
| `GP` | Girls Apparel | `GIRLS` / `KIDS` | `Kid's` / `Girls Clothing` |
| `KD` | Kids Apparel | `KIDS` | `Kid's` / `Kids Clothing` |
| `MT`, `MS`, `MP`, `MV`, `MI`, `MJ`, `MCS`, `M` | Men's Apparel | `MENS` | `Men's Clothing` |
| `WT`, `WS`, `WP`, `WV`, `WI`, `WJ`, `WCS`, `W` | Women's Apparel | `WOMENS` | `Women's Clothing` |

### Critical Rule: Gender Conflict Guard
- When falling back to barcode lookup (`excel_idx_barcode`), the engine MUST verify `gender_matches()`.
- A `BT...` / `B...` / `K...` (Boys/Kids) tag must **NEVER** be matched to an Excel row with `Category: "Men's Clothing"` or `Gender: "MENS"`.
- If an Excel row has an outdated or mismatched Category, it must be overridden using the verified Google Sheet master for that style (e.g. style `B208` -> `GENDER: KIDS` -> `Kid's`).

---

## 2. Tag Verification Modes
1. **D2C Dress Tag File**:
   - Fields: `Description`, `Fit` (conditional), `Category`, `MRP`, `SKU`, `EAN`, `Size`, `Color`, `Qty`, `Lot No (Google Sheet)`, `Lot No (GS1 Master)`.
2. **B2B Box Sticker Tag File**:
   - Verifies `Description`, `Lot No`, `Qty`, `Total MRP`, `MRP`, `SKU`, `EAN`, `Size`.
3. **B2B Bundle Sticker Tag File**:
   - Verifies `Color` instead of `Description`, plus `Lot No`, `Qty`, `Total MRP`, `SKU`, `EAN`, `Size`.

---

## 3. Key Learned Solutions & Normalizations
- **Fit Field Presence Check**:
  If the physical tag does not contain a `Fit:` line, skip the `Fit` check completely instead of showing a false mismatch.
  Normalize common reference typos: `RUGULAR FIT` -> `Regular Fit`.
- **Size-Tiered Pricing**:
  Styles like `OR51`, `OR40`, etc., have differential pricing across sizes (e.g., S/M/L vs XL/2XL). The engine must match the exact size code against the Google Sheet row to fetch the correct tiered MRP.
- **Unlisted / New Colorway Simulation**:
  For new colorways not yet registered in GS1 (e.g., `BTB208OFW`, `RST`, `BTE`), the engine creates simulated master entries referencing the base style (`B208`) from Google Sheets and dynamic color mappings.
- **Assorted / Pack Codes**:
  Suffixes like `ASC08Y001`, `008`, `2PK`, `3PK` represent multi-packs or batch variants. The engine strips batch codes when checking base styles while retaining them for pack quantity and price calculations.
- **Category Synonyms**:
  `Kid's`, `Kids Clothing`, and `Boy's` are equivalent apparel categories and must evaluate to `Match`.
- **Garment Category Prefix Typo Tolerance (v3.2)**:
  In Master sheets, pants/bottomwear are occasionally entered with `WT` or `MT` instead of `WP` or `MP` (e.g. `WR134/5 WOMENS BASIC TRACKPANT` with `WTW134CHAMED` vs tag `WPW134CHAMED`). The engine recognizes they share the same gender and core SKU (`Style + Color + Size`), automatically linking and evaluating them as `✅ Match`.
