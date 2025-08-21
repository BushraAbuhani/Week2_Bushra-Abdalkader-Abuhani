# app.py
# -*- coding: utf-8 -*-
# Excel Manager — Baby Violet Enhanced Edition
# Bushra Abu Hani
# Features: Create/Delete Excel files, Add/Edit/Delete rows & columns, Visual effects, Fully editable table

import streamlit as st
import pandas as pd
from pathlib import Path
import os

# ========= Page Setup =========
st.set_page_config(page_title="Excel Manager — Bushra Abu Hani", layout="wide", page_icon="📊")

# ========= CSS for Baby Violet Theme & Animations =========
st.markdown("""
<style>
  body {
      background: linear-gradient(160deg, #E6E6FA 0%, #D8BFD8 100%);
  }
  .shimmer {
      background: linear-gradient(90deg, rgba(255,255,255,0) 0%, rgba(255,255,255,.6) 50%, rgba(255,255,255,0) 100%);
      background-size: 200% 100%;
      animation: shimmer 2.2s infinite;
      -webkit-background-clip: text;
      background-clip: text;
      color: transparent;
      font-size: 2.4rem;
  }
  @keyframes shimmer { 0%{background-position: -200% 0} 100%{background-position: 200% 0} }
  .fadein { animation: fadein 0.8s ease-in-out; }
  @keyframes fadein { from {opacity: 0; transform: translateY(-6px)} to {opacity: 1; transform: translateY(0)} }
  .pulse { animation: pulse 2.5s infinite; }
  @keyframes pulse { 0%{box-shadow: 0 0 0 0 rgba(59,130,246,.4)} 70%{box-shadow: 0 0 0 12px rgba(59,130,246,0)} 100%{box-shadow: 0 0 0 0 rgba(59,130,246,0)} }
  .card {
      border-radius: 20px;
      border: 1px solid rgba(0,0,0,.08);
      padding: 1rem;
      background: linear-gradient(180deg, rgba(230,230,250,.9), rgba(220,220,250,.8));
      box-shadow: 0 8px 20px rgba(0,0,0,0.1);
  }
  button {
      background-color: #D8BFD8 !important;  
      color: #fff !important;
      font-weight: bold;
  }
  button:hover {
      background-color: #C8A2C8 !important;
      color: #fff !important;
  }
</style>
""", unsafe_allow_html=True)

# ========= Header =========
st.markdown('<h1 class="shimmer">📊 Excel Data Manager — Bushra Abu Hani</h1>', unsafe_allow_html=True)

# ========= Data Folder Setup =========
DATA_FOLDER = Path("All Data - Bushra Abu Hani")
DATA_FOLDER.mkdir(parents=True, exist_ok=True)

# ========= Helper Functions =========
def list_excel_files(folder: Path):
    return sorted([f for f in os.listdir(folder) if f.lower().endswith(".xlsx")])

def safe_read_excel(path: Path) -> pd.DataFrame:
    if not path.exists() or path.stat().st_size == 0:
        return pd.DataFrame()
    try:
        return pd.read_excel(path)
    except Exception as e:
        st.warning(f"Failed to read file: {e}")
        return pd.DataFrame()

def safe_write_excel(path: Path, df: pd.DataFrame):
    try:
        df.to_excel(path, index=False)
        st.toast("Saved successfully ✅", icon="✅")
    except Exception as e:
        st.error(f"Failed to save: {e}")

def create_excel(path: Path):
    if path.exists():
        st.info("File already exists.")
    else:
        safe_write_excel(path, pd.DataFrame())
        st.balloons()

# ========= Sidebar =========
with st.sidebar:
    st.header("📂 Files")
    files = list_excel_files(DATA_FOLDER)
    current_file = st.selectbox("Select Excel file", options=["— None —"] + files, index=0)

    st.markdown("---")
    st.subheader("➕ Create New File")
    new_name = st.text_input("File name (e.g., data.xlsx)", value="")
    if st.button("Create File", use_container_width=True):
        if not new_name.lower().endswith(".xlsx"):
            st.warning("Add the .xlsx extension.")
        else:
            create_excel(DATA_FOLDER / new_name)
            st.rerun()

    st.markdown("---")
    st.subheader("🗑️ Delete File")
    if current_file != "— None —":
        confirm_del = st.checkbox("Confirm deletion")
        if st.button("Delete Selected File", disabled=not confirm_del, use_container_width=True):
            try:
                os.remove(DATA_FOLDER / current_file)
                st.toast("File deleted 🗑️", icon="🗑️")
                st.rerun()
            except Exception as e:
                st.error(f"Failed to delete file: {e}")

# ========= Main =========
if current_file != "— None —":
    path = DATA_FOLDER / current_file
    st.markdown(f"#### 🗃️ Current File: `{current_file}`")
    df = safe_read_excel(path)

    # Remove old Name/Row Name columns if exist
    df = df[[c for c in df.columns if c.lower() not in ["name", "row name"]]]

    # --- Column & Row Management ---
    with st.container():
        st.markdown('<div class="card fadein">', unsafe_allow_html=True)
        st.subheader("Columns & Rows Management")
        c1, c2, c3 = st.columns([1,1,1])

        # Add Columns
        with c1:
            st.write("➕ Add Columns")
            ncols = st.number_input("Number of Columns", min_value=1, max_value=20, value=1, step=1)
            col_names = []
            for i in range(ncols):
                col_name = st.text_input(f"Column {i+1} Name", value=f"Column_{i+1}", key=f"col_name_{i}")
                # Prevent user from naming a column "Name"
                if col_name.strip().lower() == "name":
                    col_name = f"Column_{len(df.columns)+1}"
                col_names.append(col_name)
            if st.button("Add Columns", use_container_width=True):
                for name in col_names:
                    original_name = name
                    k = 2
                    while name in df.columns:
                        name = f"{original_name}_{k}"
                        k += 1
                    df[name] = None
                safe_write_excel(path, df)
                st.balloons()
                st.rerun()

        # Delete Columns
        with c2:
            st.write("➖ Delete Columns")
            cols_to_drop = st.multiselect("Select columns to delete", options=list(df.columns))
            if st.button("Delete Selected Columns", use_container_width=True, disabled=len(cols_to_drop)==0):
                df = df.drop(columns=cols_to_drop, errors="ignore")
                safe_write_excel(path, df)
                st.toast("Columns deleted 🗑️", icon="🗑️")
                st.rerun()

        # Add Rows
        with c3:
            st.write("➕ Add Rows")
            nrows = st.number_input("Number of Rows", min_value=1, max_value=50, value=1, step=1)
            if st.button("Add Rows", use_container_width=True):
                blanks = pd.DataFrame([{c: None for c in df.columns} for _ in range(nrows)])
                df = pd.concat([df, blanks], ignore_index=True)
                safe_write_excel(path, df)
                st.balloons()
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    # --- Editable Table ---
    st.markdown("### ✏️ Editable Table")
    original_dtypes = df.dtypes.to_dict()
    df_editable = df.astype(str)

    edited_df = st.data_editor(
        df_editable,
        hide_index=True,
        use_container_width=True,
        num_rows="dynamic",
        key="editor",
        column_config={col: st.column_config.TextColumn(label=col) for col in df_editable.columns}
    )

    def restore_dtypes(original_dtypes, edited_df):
        for col, dtype in original_dtypes.items():
            if pd.api.types.is_numeric_dtype(dtype):
                edited_df[col] = pd.to_numeric(edited_df[col], errors='coerce')
        return edited_df

    # --- Action Buttons ---
    st.markdown('<div class="card fadein">', unsafe_allow_html=True)
    s1, s2, s3 = st.columns([1,1,1])

    # Save table
    with s1:
        if st.button("💾 Save Table", use_container_width=True):
            edited_df_to_save = restore_dtypes(original_dtypes, edited_df)
            safe_write_excel(path, edited_df_to_save)
            st.balloons()

    # Delete rows
    with s2:
        if not edited_df.empty:
            rows_to_delete = st.multiselect(
                "Select Rows to Delete",
                options=list(range(len(edited_df))),
                format_func=lambda x: f"Row {x+1}"
            )
            if st.button("🗑️ Delete Selected Rows", disabled=len(rows_to_delete)==0, use_container_width=True):
                new_df = edited_df.drop(index=rows_to_delete).reset_index(drop=True)
                safe_write_excel(path, new_df)
                st.toast("Rows deleted 🗑️", icon="🗑️")
                st.balloons()
                st.rerun()

    # Reload
    with s3:
        if st.button("⟲ Reload Table", use_container_width=True):
            st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

else:
    st.info("Select a file from the sidebar or create a new one.")

# ========= Special File: Warehouse Number One =========
with st.expander("🏭 Create 'Warehouse Number One' File"):
    w_name = "Warehouse Number One.xlsx"
    cols = ["Product", "Quantity", "Amount", "Weight", "Product Serial Number", "Product Supplier"]
    if st.button("Create Template File", type="primary"):
        path = DATA_FOLDER / w_name
        if path.exists():
            st.warning("File already exists.")
        else:
            df = pd.DataFrame(columns=cols)
            df.to_excel(path, index=False)
            st.success("File created ✅")
            st.balloons()

# ========= Footer =========
st.markdown("""
<div style="margin-top: 1rem; padding: .8rem 1rem; border-radius: 20px;
            background: linear-gradient(90deg, #D8BFD8 0%, #E6E6FA 100%);
            color: #fff; font-weight: bold; text-align: center;">
✨ Enjoy!
</div>
""", unsafe_allow_html=True)
