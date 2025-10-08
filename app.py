import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime, timedelta
from openpyxl import load_workbook

st.set_page_config(page_title="PJUM Perdin Processor", layout="wide")

# --------------------------------------------
# Helper functions
# --------------------------------------------
def parse_date(value):
    if pd.isna(value):
        return pd.NaT
    if isinstance(value, (datetime, pd.Timestamp)):
        return pd.to_datetime(value).normalize()
    text = str(value).strip()
    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y", "%d-%m-%y", "%Y-%m-%d", "%m/%d/%Y"):
        try:
            return pd.to_datetime(datetime.strptime(text, fmt)).normalize()
        except Exception:
            continue
    return pd.to_datetime(text, dayfirst=True, errors="coerce").normalize()

def to_excel_bytes(df):
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="PJUM_Period")
    bio.seek(0)
    return bio

def add_business_days(start_date, add_days, holidays_set):
    if pd.isna(start_date):
        return pd.NaT
    current = pd.to_datetime(start_date).date()
    added = 0
    while added < add_days:
        current += timedelta(days=1)
        if current.weekday() >= 5 or current in holidays_set:
            continue
        added += 1
    return pd.to_datetime(current)

def standardize_input_df(df):
    df = df.rename(columns=lambda x: str(x).strip())
    for col in ["Posting Date", "End Date", "Expired Date", "Posting Date PJUM"]:
        if col in df.columns:
            df[col] = df[col].apply(parse_date)
    if "Nilai" in df.columns:
        df["Nilai"] = pd.to_numeric(df["Nilai"], errors="coerce")
    return df

def compute_sla_and_status(df, holidays_df, ref_date):
    holidays_df = holidays_df.copy()
    if "Tanggal" in holidays_df.columns:
        holidays_df["Tanggal"] = holidays_df["Tanggal"].apply(parse_date)
    holidays = set(d.date() for d in holidays_df["Tanggal"].dropna().unique())

    df = df.copy()
    sla_dates, sla_days = [], []

    for _, row in df.iterrows():
        header = str(row.get("Header Text", "")).strip()
        base_date = row.get("End Date", pd.NaT) or row.get("Posting Date", pd.NaT)
        days_to_add = 20 if header.startswith("S-") else 10
        sla_dates.append(add_business_days(base_date, days_to_add, holidays) if not pd.isna(base_date) else pd.NaT)
        sla_days.append(days_to_add)

    df["SLA PJUM + 10/20HK"] = sla_dates
    status_list = []

    for _, row in df.iterrows():
        posting_pjum = row.get("Posting Date PJUM", pd.NaT)
        sla_date = row.get("SLA PJUM + 10/20HK", pd.NaT)
        end_date = row.get("End Date", pd.NaT)

        if not pd.isna(posting_pjum):
            status_list.append("Telat" if posting_pjum > sla_date else "Tidak Telat")
        elif not pd.isna(end_date) and end_date > pd.to_datetime(ref_date):
            status_list.append("Kegiatan Belum Berakhir")
        elif ref_date > sla_date:
            status_list.append("Lewat Jatuh Tempo")
        else:
            status_list.append("Belum Jatuh Tempo")

    df["Status"] = status_list
    df["Nomor SAP PJUM Pertama Kali"] = ""
    df["Tanggal Diterima GPBN"] = pd.NaT
    df["Status No SAP PJUM"] = ""
    df["Status Final"] = ""
    return df

def compute_status_after_manual(df):
    df = df.copy()
    for idx, row in df.iterrows():
        manual_no = str(row.get("Nomor SAP PJUM Pertama Kali", "")).strip()
        doc_pjum = str(row.get("Doc SAP PJUM", "")).strip()
        if manual_no:
            df.at[idx, "Status No SAP PJUM"] = (
                "No Doc SAP Sama" if manual_no == doc_pjum else "No Doc SAP Berbeda"
            )
    for idx, row in df.iterrows():
        tgl_diterima, sla_date, orig_status = (
            row.get("Tanggal Diterima GPBN", pd.NaT),
            row.get("SLA PJUM + 10/20HK", pd.NaT),
            row.get("Status", ""),
        )
        if not pd.isna(tgl_diterima) and not pd.isna(sla_date):
            df.at[idx, "Status Final"] = (
                "Tidak Telat" if tgl_diterima <= sla_date else orig_status
            )
        else:
            df.at[idx, "Status Final"] = orig_status
    return df

# --------------------------------------------
# UI
# --------------------------------------------
st.title("PJUM Perdin — Streamlit Processor")
st.markdown("Upload a single Excel file with **two sheets:** `Report` and `Holidays`. The app will compute SLA & status automatically.")

uploaded = st.file_uploader("Upload Combined Excel (Report + Holidays)", type=["xlsx"], help="Sheet1: Report, Sheet2: Holidays")

ref_date = st.date_input("Reference date (default = today)", value=datetime.today().date())

if uploaded is None:
    st.info("Please upload the combined Excel file to start.")
    st.stop()

# Read both sheets safely
try:
    xls = pd.ExcelFile(uploaded, engine="openpyxl")
    report_df = pd.read_excel(xls, sheet_name=0, dtype=str)
    holidays_df = pd.read_excel(xls, sheet_name=1, dtype=str)
except Exception as e:
    st.error(f"Failed to read Excel: {e}")
    st.stop()

report_df = standardize_input_df(report_df)
processed = compute_sla_and_status(report_df, holidays_df, pd.to_datetime(ref_date))

st.markdown("### Preview (first 50 rows)")
st.dataframe(processed.head(50), use_container_width=True)

# Manual workflow selection
st.markdown("---")
manual_mode = st.radio("Manual workflow", ["Only compute SLA/status & download (no manual input)", "Manual input for Telat rows now (recommended)"], index=1)

df_for_edit = processed.copy()

if manual_mode.startswith("Manual input"):
    telat_rows = df_for_edit[df_for_edit["Status"] == "Telat"].copy()
    if telat_rows.empty:
        st.info("No rows with Status = 'Telat' found.")
    else:
        st.markdown("**Rows with Telat — fill Nomor SAP PJUM Pertama Kali and Tanggal Diterima GPBN if available.**")
        edited = st.data_editor(
            telat_rows[
                [
                    "Assignment", "Doc SAP", "Doc SAP PJUM", "Header Text",
                    "End Date", "SLA PJUM + 10/20HK", "Status",
                    "Nomor SAP PJUM Pertama Kali", "Tanggal Diterima GPBN"
                ]
            ],
            use_container_width=True,
        )
        if st.button("Apply manual changes"):
            idxs = df_for_edit[df_for_edit["Status"] == "Telat"].index
            for i, idx in enumerate(idxs):
                for col in ["Nomor SAP PJUM Pertama Kali", "Tanggal Diterima GPBN"]:
                    df_for_edit.at[idx, col] = edited.iloc[i][col]
            df_for_edit["Tanggal Diterima GPBN"] = df_for_edit["Tanggal Diterima GPBN"].apply(parse_date)
            df_for_edit = compute_status_after_manual(df_for_edit)
            st.success("Manual changes applied.")
            st.dataframe(df_for_edit.head(50), use_container_width=True)

else:
    st.info("Only compute SLA/status selected. Displaying results directly.")
    st.dataframe(processed.head(100), use_container_width=True)

# Prepare final output
final_df = df_for_edit if manual_mode.startswith("Manual") else processed
date_cols = [c for c in final_df.columns if "Date" in c or "Tanggal" in c]
for c in date_cols:
    final_df[c] = pd.to_datetime(final_df[c], errors="coerce").dt.strftime("%d/%m/%y")

st.markdown("---")
st.subheader("Final preview (first 100 rows)")
st.dataframe(final_df.head(100), use_container_width=True)

# Downloads
st.markdown("### Download Data")
col1, col2 = st.columns(2)
with col1:
    st.download_button(
        "⬇️ Download Sebelum Manual Input (.xlsx)",
        data=to_excel_bytes(processed),
        file_name=f"PJUM_before_{datetime.today().strftime('%Y%m%d')}.xlsx",
    )
with col2:
    st.download_button(
        "⬇️ Download Setelah Manual Input (.xlsx)",
        data=to_excel_bytes(final_df),
        file_name=f"PJUM_after_{datetime.today().strftime('%Y%m%d')}.xlsx",
    )
