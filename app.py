import io
import openpyxl
import pandas as pd
import streamlit as st
import builtins

st.set_page_config(page_title="Lease Overview App", layout="wide")
st.title("🏢 Multi-Tab Lease Overview App")

st.markdown("""
Upload an Excel workbook containing the necessary sheets:
* **`Square_Meters`** / `Square_Metres`
* **`Prices`**
* **`Add-on`**
* **`Add-on participation`**
""")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])


def format_cz_number(val):
    """Formats numeric floats into Czech number format (e.g. 1 234,56), hiding zero values."""
    if pd.isnull(val):
        return ""
    try:
        val_float = float(val)
        if abs(val_float) < 0.001:
            return ""
        return (
            f"{val_float:,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", " ")
        )
    except (ValueError, TypeError):
        return str(val)


def get_column_configs(df, numeric_cols):
    """Generates column configuration to right-align numeric columns."""
    configs = {}
    for col in df.columns:
        if col in numeric_cols:
            configs[col] = st.column_config.TextColumn(col, alignment="right")
        else:
            configs[col] = st.column_config.TextColumn(col, alignment="left")
    return configs


def process_original_lease_data(df_raw):
    """Processes raw lease DataFrame, coercing numbers and suppressing empty rows."""
    df = df_raw.copy().dropna(how="all", axis=1).dropna(how="all", axis=0)

    numeric_cols = list(df.columns[1:])
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    # Filter out rows where all numeric values are 0
    df = df[~(df[numeric_cols] == 0).all(axis=1)].reset_index(drop=True)

    df_display = df.copy()
    for col in numeric_cols:
        df_display[col] = df_display[col].apply(format_cz_number)

    configs = get_column_configs(df_display, numeric_cols)
    return df, df_display, configs, numeric_cols


def process_price_list_data(df_raw):
    """Formats numeric columns in tabular data to Czech formatting and builds configs."""
    df_prices_display = df_raw.copy()
    numeric_cols = []
    
    for col in df_prices_display.columns:
        if pd.api.types.is_numeric_dtype(df_prices_display[col]) or any(
            isinstance(val, (int, float)) for val in df_prices_display[col].dropna()
        ):
            numeric_cols.append(col)
            df_prices_display[col] = df_prices_display[col].apply(format_cz_number)
            
    configs = get_column_configs(df_prices_display, numeric_cols)
    return df_prices_display, configs


if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names

        sqm_sheet = next((s for s in sheet_names if "square" in s.lower()), sheet_names[0])
        prices_sheet = next((s for s in sheet_names if "price" in s.lower()), None)
        addon_sheet = next((s for s in sheet_names if "add-on" in s.lower() and "participation" not in s.lower()), None)
        part_sheet = next((s for s in sheet_names if "participation" in s.lower()), None)

        df_sqm_raw = pd.read_excel(uploaded_file, sheet_name=sqm_sheet)

        tab1, tab2, tab3, tab4, tab5 = st.tabs(
            [
                "📋 Panel 1: Main Lease List",
                "🏷️ Panel 2: Price List",
                "➕ Panel 3: Add-on Rates",
                "🤝 Panel 4: Participation Rules",
                "🧮 Panel 5: Calculated Lease Summary",
            ]
        )

        # ------------------------------------------------------------------
        # PANEL 1: Main Lease List (Raw Quantities)
        # ------------------------------------------------------------------
        with tab1:
            st.subheader(f"Main Lease Overview (Sheet: `{sqm_sheet}`)")
            df_sqm_clean, df_sqm_display, configs1, num_cols1 = process_original_lease_data(df_sqm_raw)
            st.dataframe(
                df_sqm_display, 
                width="stretch", 
                height="content",
                column_config=configs1
            )

        # ------------------------------------------------------------------
        # PANEL 2: Master Price List
        # ------------------------------------------------------------------
        with tab2:
            st.subheader(f"Master Prices per Unit " + (f"(Sheet: `{prices_sheet}`)" if prices_sheet else ""))
            if prices_sheet:
                df_prices_raw = pd.read_excel(uploaded_file, sheet_name=prices_sheet)
                df_prices_raw.columns = df_prices_raw.columns.str.strip()
                df_prices_display, configs2 = process_price_list_data(df_prices_raw)
                st.dataframe(
                    df_prices_display, 
                    width="stretch", 
                    height="content",
                    column_config=configs2
                )
            else:
                st.warning("No sheet found for Prices.")

        # ------------------------------------------------------------------
        # PANEL 3: Add-on Rates
        # ------------------------------------------------------------------
        addon_rates = {}
        with tab3:
            st.subheader(f"Add-on Percentage Rates " + (f"(Sheet: `{addon_sheet}`)" if addon_sheet else ""))
            if addon_sheet:
                df_addon_raw = pd.read_excel(uploaded_file, sheet_name=addon_sheet)
                df_addon_raw.columns = df_addon_raw.columns.str.strip()
                
                # Build Add-on rates lookup
                if "Typ" in df_addon_raw.columns and "Add-on" in df_addon_raw.columns:
                    for _, row in df_addon_raw.iterrows():
                        key = str(row["Typ"]).strip()
                        rate = pd.to_numeric(row["Add-on"], errors="coerce")
                        if pd.notnull(rate):
                            addon_rates[key] = float(rate)

                df_addon_display, configs_addon = process_price_list_data(df_addon_raw)
                st.dataframe(
                    df_addon_display,
                    width="stretch",
                    height="content",
                    column_config=configs_addon
                )
            else:
                st.warning("No sheet found for Add-on rates.")

        # ------------------------------------------------------------------
        # PANEL 4: Add-on Participation Table
        # ------------------------------------------------------------------
        df_part_raw = None
        with tab4:
            st.subheader(f"Add-on Participation Matrix " + (f"(Sheet: `{part_sheet}`)" if part_sheet else ""))
            if part_sheet:
                df_part_raw = pd.read_excel(uploaded_file, sheet_name=part_sheet)
                df_part_raw.columns = df_part_raw.columns.str.strip()
                st.dataframe(
                    df_part_raw,
                    width="stretch",
                    height="content"
                )
            else:
                st.warning("No sheet found for Add-on participation.")

        # ------------------------------------------------------------------
        # PANEL 5: Calculated Lease Summary with Participation Logic
        # ------------------------------------------------------------------
        with tab5:
            st.subheader("Calculated Lease Totals (CZK) - Conditional Participation Included")

            if prices_sheet and "Kód" in df_prices_raw.columns and "Cena za jednotku (Bez DPH)" in df_prices_raw.columns:
                # Build lookup mapping: Kód -> Unit Price
                df_prices_base = df_prices_raw.drop_duplicates(subset=["Kód"], keep="first")
                price_map = dict(zip(df_prices_base["Kód"], df_prices_base["Cena za jednotku (Bez DPH)"]))

                # Helper function to get participation status ("YES" -> True)
                def check_participation(company, typ_str, category):
                    if df_part_raw is None or df_part_raw.empty:
                        return True
                    match = df_part_raw[
                        (df_part_raw["Společnost"].astype(str).str.strip() == str(company).strip()) &
                        (df_part_raw["Typ"].astype(str).str.strip().str.lower() == str(typ_str).strip().lower())
                    ]
                    if not match.empty and category in match.columns:
                        val = str(match[category].values[0]).strip().upper()
                        return val in ["YES", "ANO", "TRUE", "1"]
                    return False

                label_col = df_sqm_clean.columns[0]
                calc_dict = {label_col: df_sqm_clean[label_col]}
                calc_cols = []

                for col in num_cols1:
                    if col in price_map:
                        unit_price = price_map[col]

                        # Calculate calculated amount per company for column 'col'
                        calculated_series = []
                        for idx, row in df_sqm_clean.iterrows():
                            company = row[label_col]
                            sqm_val = row[col]

                            addon_pct = 0.0

                            # 1) Office space columns (starting with K - Typ)
                            if col.startswith("K - Typ"):
                                typ_name = col.replace("K - ", "").strip()  # e.g., "Typ A", "Typ D-Archiv"
                                
                                # Mandatory office base add-on
                                addon_pct += addon_rates.get("Výtah, schodiště, lobby", 0.0)

                                # Conditional add-ons based on participation table
                                if check_participation(company, typ_name, "JM interní"):
                                    addon_pct += addon_rates.get("JM interní", 0.0)
                                if check_participation(company, typ_name, "JM reprezentativní"):
                                    addon_pct += addon_rates.get("JM reprezentativní", 0.0)
                                if check_participation(company, typ_name, "Chodby, kuchyň, copy"):
                                    addon_pct += addon_rates.get("Chodby, kuchyň, copy", 0.0)

                            # 2) Terasy column
                            elif col == "Terasy":
                                if check_participation(company, "Terasy", "Terasy") or check_participation(company, "Typ A", "Terasy"):
                                    addon_pct += addon_rates.get("Terasy", 0.0)

                            # 3) Sklad column
                            elif col == "Sklad":
                                addon_pct += addon_rates.get("Sklad", 0.0)

                            # Calculate final amount for this item
                            item_total = sqm_val * (1.0 + addon_pct) * unit_price
                            calculated_series.append(item_total)

                        calc_dict[col] = calculated_series
                        calc_cols.append(col)

                df_calc = pd.DataFrame(calc_dict)

                # Total cost column per company across all lease codes
                if calc_cols:
                    df_calc["Celková Cena (CZK)"] = df_calc[calc_cols].sum(axis=1)

                # Filter out rows where total calculated cost is 0
                if "Celková Cena (CZK)" in df_calc.columns:
                    df_calc = df_calc[df_calc["Celková Cena (CZK)"] > 0].reset_index(drop=True)

                # Format table for UI
                df_calc_display, configs5 = process_price_list_data(df_calc)

                st.dataframe(
                    df_calc_display, 
                    width="stretch", 
                    height="content",
                    column_config=configs5
                )
            else:
                st.info("Panel 5 will display calculated totals once all required sheets and columns are available.")

    except Exception as e:
        st.error(f"Error reading workbook: {e}")
