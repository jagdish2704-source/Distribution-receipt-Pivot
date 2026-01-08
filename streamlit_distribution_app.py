import streamlit as st
import pandas as pd
import numpy as np
import io

# --- CORE LOGIC (Reused) ---

def generate_product_column(df):
    """
    Applies the Product column generation logic.
    Logic:
    0. Clean headers: Replace "=-" with "" (blank).
    1. Insert 'Product' at Column S (Index 18).
    2. F Col (Index 5) contains 'GIMA WHITE GOLD' -> Set 'GIMA WHITE GOLD'.
    3. F Col (Index 5) ends with 'OIL' -> Copy F.
    4. R Col (Index 17) contains 'YARN' (case-insensitive) -> Set 'YARN'.
    5. R Col (Index 17) contains 'STORE' (case-insensitive) -> Set 'STORE'.
    6. Fallback -> Copy G Col (Index 6).
    """
    if df is None: return None
    
    # 0. Clean headers: Remove "=-"
    df.columns = df.columns.str.replace("=-", "", regex=False).str.strip()

    # Validation of shape
    if len(df.columns) < 18:
        st.error(f"Input file has only {len(df.columns)} columns. Requires at least 18 (Col R).")
        return None

    # Indices
    IDX_F = 5
    IDX_G = 6
    IDX_R = 17
    IDX_S = 18

    # Extract Series
    col_f_str = df.iloc[:, IDX_F].fillna("").astype(str).str.strip()
    col_r_str = df.iloc[:, IDX_R].fillna("").astype(str)
    col_g_vals = df.iloc[:, IDX_G]

    # Logic Conditions
    cond_gima = col_f_str.str.contains("GIMA WHITE GOLD", case=False, regex=False)
    cond_oil = col_f_str.str.endswith("OIL")
    cond_yarn = col_r_str.str.contains("YARN", case=False, na=False)
    cond_store = col_r_str.str.contains("STORE", case=False, na=False)

    # Use numpy select for precedence
    new_s_values = np.select(
        [cond_gima, cond_oil, cond_yarn, cond_store],
        ["GIMA WHITE GOLD", df.iloc[:, IDX_F], "YARN", "STORE"],
        default=col_g_vals
    )

    # Insert 'Product' logic
    # If Product already exists, we drop it to avoid duplicates or errors on insert
    if "Product" in df.columns:
        df = df.drop(columns=["Product"])
        
    df.insert(IDX_S, "Product", new_s_values)
    return df

def generate_pivot_table(df):
    """
    Generates the pivot table with:
    - Row: Product, SITE_CODE
    - Column: SHIP_SITE, Grand Total
    - Values: QUANTITY, AMOUNT
    - Subtotals: Product-wise
    """
    required_cols = ["SITE_CODE", "SHIP_SITE", "QUANTITY", "AMOUNT", "Product"]
    missing = [c for c in required_cols if c not in df.columns]
    
    if missing:
        st.error(f"Missing columns for Pivot: {missing}")
        return None

    # Ensure numeric types - Handle commas
    for col in ["QUANTITY", "AMOUNT"]:
        df[col] = df[col].astype(str).str.replace(",", "", regex=False)
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # Pivot Main (Detailed)
    pivot_main = pd.pivot_table(
        df,
        index=["Product", "SITE_CODE"],
        columns=["SHIP_SITE"],
        values=["QUANTITY", "AMOUNT"],
        aggfunc="sum",
        fill_value=0
    )
    
    if pivot_main.empty:
        st.error("Pivot table is empty.")
        return None

    # Pivot Subtotals (Product-wise)
    pivot_sub = pd.pivot_table(
        df,
        index=["Product"],
        columns=["SHIP_SITE"],
        values=["QUANTITY", "AMOUNT"],
        aggfunc="sum",
        fill_value=0
    )
    
    # Add "Total" level
    pivot_sub.index = pd.MultiIndex.from_arrays(
        [pivot_sub.index, ["Total"] * len(pivot_sub)], 
        names=["Product", "SITE_CODE"]
    )
    
    # Combine and Sort
    final_pivot = pd.concat([pivot_main, pivot_sub])
    final_pivot.sort_index(inplace=True)
    
    # Swap levels to SHIP_SITE on top
    final_pivot.columns = final_pivot.columns.swaplevel(0, 1)

    # Calculate Horizontal Totals
    total_qty = final_pivot.xs('QUANTITY', axis=1, level=1).sum(axis=1)
    total_amt = final_pivot.xs('AMOUNT', axis=1, level=1).sum(axis=1)
    
    # Add Grand Total
    final_pivot[('Grand Total', 'QUANTITY')] = total_qty
    final_pivot[('Grand Total', 'AMOUNT')] = total_amt
    
    # Sort columns
    final_pivot.sort_index(axis=1, level=0, inplace=True)
    
    return final_pivot

def to_excel(df_processed, df_pivot):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Sheet 1: Processed Data
        df_processed.to_excel(writer, sheet_name="Processed Data", index=False)
        
        # Sheet 2: Pivot Table
        if df_pivot is not None:
            # Sort Columns Logic
            sites = [s for s in df_pivot.columns.get_level_values(0).unique() if s != "Grand Total"]
            sites.sort()
            if "Grand Total" in df_pivot.columns.get_level_values(0):
                sites.append("Grand Total")
            
            sorted_cols = []
            for site in sites:
                sorted_cols.append((site, "QUANTITY"))
                sorted_cols.append((site, "AMOUNT"))
            
            existing_cols = [c for c in sorted_cols if c in df_pivot.columns]
            pivot_sorted = df_pivot[existing_cols]

            styler = pivot_sorted.style.format("{:,.2f}")
            
            # Styling
            styler.set_table_styles([
                {'selector': 'th', 'props': [('font-size', '10pt'), ('text-align', 'center'), ('font-weight', 'bold'), ('background-color', '#f8f9fa'), ('border', '1px solid #ddd')]}
            ])
            
            def highlight_cols(series):
                site, metric = series.name
                if site == "Grand Total":
                    return ['background-color: #dce5f0; font-weight: bold; border-left: 2px solid #aaa;']*len(series)
                if metric == "QUANTITY":
                    return ['background-color: #e3f2fd']*len(series)
                elif metric == "AMOUNT":
                    return ['background-color: #e8f5e9']*len(series)
                return ['']*len(series)
            
            def highlight_totals(series):
                if series.name[1] == "Total":
                    return ['font-weight: bold; background-color: #cfcfcf; border-top: 2px solid #555;']*len(series)
                return ['']*len(series)

            styler.apply(highlight_cols, axis=0)
            styler.apply(highlight_totals, axis=1)
            
            styler.to_excel(writer, sheet_name="Distribution Pivot")
            
    return output.getvalue()

# --- STREAMLIT APP ---

st.set_page_config(page_title="Distribution Details Generator", layout="wide")

st.title("📊 Distribution Details Generator")
st.markdown("Upload your CSV file to generate the Product column and Distribution Pivot Table with analytics.")

uploaded_file = st.file_uploader("Upload CSV File", type=["csv", "txt"])

if uploaded_file is not None:
    try:
        df = pd.read_csv(uploaded_file, encoding='ISO-8859-1') # Fallback encoding
        st.success(f"Loaded {uploaded_file.name} with {len(df)} rows.")

        if st.button("Generate Product Column"):
            with st.spinner("Processing Logic..."):
                processed_df = generate_product_column(df.copy())
                if processed_df is not None:
                    st.session_state['processed_df'] = processed_df
                    st.success("Product Column Generated!")
                    st.dataframe(processed_df.head())

        if 'processed_df' in st.session_state:
            if st.button("Generate Distribution Pivot"):
                with st.spinner("Creating Pivot Table..."):
                    pivot_df = generate_pivot_table(st.session_state['processed_df'])
                    if pivot_df is not None:
                        st.session_state['pivot_df'] = pivot_df
                        st.success("Pivot Table Generated!")
                        st.dataframe(pivot_df)

            if 'pivot_df' in st.session_state:
                excel_data = to_excel(st.session_state['processed_df'], st.session_state['pivot_df'])
                st.download_button(
                    label="📥 Download Excel Report",
                    data=excel_data,
                    file_name="Distribution_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Error loading file: {e}")
