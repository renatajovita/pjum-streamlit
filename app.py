# app.py
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime, timedelta
import calendar
from openpyxl import load_workbook
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

st.set_page_config(page_title="automatics reporting", layout="wide")

# -------------------------
# Helper functions
# -------------------------
def parse_date(value):
    if pd.isna(value):
        return pd.NaT
    if isinstance(value, (datetime, pd.Timestamp)):
        return pd.to_datetime(value).normalize()
    text = str(value).strip()
    # Explicitly handle both 2-digit and 4-digit year formats, always day-first
    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y", "%d-%m-%y", "%Y-%m-%d", "%m/%d/%Y"):
        try:
            parsed = datetime.strptime(text, fmt)
            # normalize to midnight
            return pd.to_datetime(parsed).normalize()
        except Exception:
            continue
    # fallback: let pandas infer but enforce dayfirst
    try:
        return pd.to_datetime(text, dayfirst=True, errors="coerce").normalize()
    except Exception:
        return pd.NaT

def to_excel_bytes(df):
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="PJUM_Period")
    bio.seek(0)
    return bio

def add_business_days(start_date, add_days, holidays_set):
    # start_date is a Timestamp or datetime.date; add_days is int (>=0)
    if pd.isna(start_date):
        return pd.NaT
    current = pd.to_datetime(start_date).date()
    added = 0
    while added < add_days:
        current += timedelta(days=1)
        # skip weekends
        if current.weekday() >= 5:
            continue
        # skip holidays
        if current in holidays_set:
            continue
        added += 1
    return pd.to_datetime(current)

def standardize_input_df(df):
    # Ensure expected columns exist and standardize types
    # Convert column names to trimmed strings
    df = df.rename(columns=lambda x: str(x).strip())
    # parse date-like columns heuristically:
    date_candidates = ["Posting Date", "End Date", "Expired Date", "Posting Date PJUM"]
    for col in date_candidates:
        if col in df.columns:
            df[col] = df[col].apply(parse_date)
    # numeric candidate
    if "Nilai" in df.columns:
        df["Nilai"] = pd.to_numeric(df["Nilai"], errors="coerce")
    return df

def compute_sla_and_status(df, holidays_df, ref_date):
    # build holiday set of date objects
    holidays_df = holidays_df.copy()
    if "Tanggal" in holidays_df.columns:
        holidays_df["Tanggal"] = holidays_df["Tanggal"].apply(parse_date)
    holidays = set()
    for d in holidays_df["Tanggal"].dropna().unique():
        holidays.add(pd.to_datetime(d).date())

    # Ensure columns exist
    if "Header Text" not in df.columns:
        st.error("Header Text column not found in uploaded report.")
        return df

    df = df.copy()

    # compute SLA days per row
    sla_days_list = []
    sla_dates = []
    for _, row in df.iterrows():
        header = str(row.get("Header Text", "")).strip()
        # find base date to add from: use 'End Date' or 'Posting Date' if End Date missing?
        # We'll use 'End Date' as in your formula (O)
        base_date = row.get("End Date", pd.NaT)
        if pd.isna(base_date):
            # fallback: Posting Date (if exists)
            base_date = row.get("Posting Date", pd.NaT)
        # determine days to add
        days_to_add = None
        if header.startswith("S-"):
            days_to_add = 20
        elif header.startswith("ST-"):
            days_to_add = 10
        else:
            # if unclear, default to 10 (or you can set 0)
            days_to_add = 10

        if pd.isna(base_date):
            sla_dates.append(pd.NaT)
            sla_days_list.append(days_to_add)
        else:
            sla_date = add_business_days(base_date, days_to_add, holidays)
            sla_dates.append(sla_date)
            sla_days_list.append(days_to_add)

    df["SLA PJUM + 10/20HK"] = sla_dates

    # Compute Status: implement logic you gave, with ref_date variable:
    # Excel logic you gave interpreted as:
    # IF(Posting Date PJUM <> "";
    #     IF(Posting Date PJUM - SLA > 0; "Telat"; "Tidak Telat");
    #     IF(End Date > ref_date; "Kegiatan Belum Berakhir";
    #         IF(ref_date - SLA > 0; "Lewat Jatuh Tempo"; "Belum Jatuh Tempo")))
    pd.options.mode.chained_assignment = None
    status_list = []
    for _, row in df.iterrows():
        posting_pjum = row.get("Posting Date PJUM", pd.NaT)
        sla_date = row.get("SLA PJUM + 10/20HK", pd.NaT)
        end_date = row.get("End Date", pd.NaT)
        if not pd.isna(posting_pjum):
            # compare posting_pjum and sla_date
            if pd.isna(sla_date):
                status_list.append("Check SLA")
            else:
                # if posting_pjum > sla_date => Telat
                if pd.to_datetime(posting_pjum) > pd.to_datetime(sla_date):
                    status_list.append("Telat")
                else:
                    status_list.append("Tidak Telat")
        else:
            if not pd.isna(end_date) and pd.to_datetime(end_date) > pd.to_datetime(ref_date):
                status_list.append("Kegiatan Belum Berakhir")
            else:
                if pd.isna(sla_date):
                    status_list.append("Check SLA")
                else:
                    # if ref_date - sla_date > 0 => Lewat Jatuh Tempo
                    if pd.to_datetime(ref_date) > pd.to_datetime(sla_date):
                        status_list.append("Lewat Jatuh Tempo")
                    else:
                        status_list.append("Belum Jatuh Tempo")
    df["Status"] = status_list

    # add empty manual columns per spec
    df["Nomor SAP PJUM Pertama Kali"] = df.get("Nomor SAP PJUM Pertama Kali", "")
    df["Tanggal Diterima GPBN"] = df["Tanggal Diterima GPBN"].apply(parse_date) if "Tanggal Diterima GPBN" in df.columns else pd.NaT
    # compute Status No SAP PJUM (blank if no manual input)
    df["Status No SAP PJUM"] = ""
    # Status Final will be created later after manual input
    df["Status Final"] = ""
    return df

def compute_status_after_manual(df):
    """Given df that may contain manual inputs, compute Status No SAP PJUM and Final status according to rules."""
    df = df.copy()
    # Status No SAP PJUM
    for idx, row in df.iterrows():
        manual_no = str(row.get("Nomor SAP PJUM Pertama Kali", "")).strip()
        doc_pjum = str(row.get("Doc SAP PJUM", "")).strip()
        if manual_no == "" or manual_no.lower() == "nan":
            df.at[idx, "Status No SAP PJUM"] = ""
        else:
            if manual_no == doc_pjum:
                df.at[idx, "Status No SAP PJUM"] = "No Doc SAP Sama"
            else:
                df.at[idx, "Status No SAP PJUM"] = "No Doc SAP Berbeda"

    # Final status: if Tanggal Diterima GPBN exists and <= SLA => becomes Tidak Telat
    for idx, row in df.iterrows():
        tgl_diterima = row.get("Tanggal Diterima GPBN", pd.NaT)
        sla_date = row.get("SLA PJUM + 10/20HK", pd.NaT)
        orig_status = row.get("Status", "")
        if not pd.isna(tgl_diterima) and not pd.isna(sla_date):
            try:
                if pd.to_datetime(tgl_diterima) <= pd.to_datetime(sla_date):
                    df.at[idx, "Status Final"] = "Tidak Telat"
                else:
                    # if manual date > SLA, keep Telat
                    df.at[idx, "Status Final"] = orig_status if orig_status else ""
            except Exception:
                df.at[idx, "Status Final"] = orig_status if orig_status else ""
        else:
            df.at[idx, "Status Final"] = orig_status if orig_status else ""
    # If any Status Final set, we will drop the original "Status" column and keep "Status Final" as final.
    return df

# -------------------------
# UI
# -------------------------
st.title("PJUM Perdin — Streamlit Processor")
st.markdown("Upload BPM report (xlsx) and Holidays (xlsx). This app will add SLA, Status, manual-input columns and allow final download.")

# Sidebar menu
menu = st.sidebar.selectbox("Menu", ["PJUM Perdin", "PJUM Kegiatan (placeholder)", "Status Pembayaran (placeholder)", "Dashboard (placeholder)"])
st.sidebar.markdown("---")
st.sidebar.info("For now only *PJUM Perdin* is active. Other modules are placeholders.")

if menu != "PJUM Perdin":
    st.header(menu)
    st.info("This section is still empty (placeholder). Come back soon.")
    st.stop()

# PJUM Perdin UI
st.header("PJUM Perdin (Report PJUM Perdin)")
col1, col2 = st.columns([1, 1])
with col1:
    uploaded = st.file_uploader("Upload PJUM Perdin report (.xlsx)", type=["xlsx"], help="Make sure the sheet starts with headers like Assignment, Doc SAP, Header Text, Posting Date, End Date, Posting Date PJUM, Doc SAP PJUM, etc.")
with col2:
    holidays_file = st.file_uploader("Upload Holidays xlsx (Tanggal, Hari, Keterangan)", type=["xlsx"], help="Holidays file must have a column 'Tanggal' with dates.")

# reference date (instead of hard-coded DATE(2025;9;10) — user can control)
ref_date = st.date_input("Reference date for status checks (default = today)", value=datetime.today().date())

if uploaded is None:
    st.info("Upload the BPM report xlsx to start processing.")
    st.stop()

# read uploaded excel into df safely even if partially corrupted
try:
    # Try normal openpyxl first
    report_df = pd.read_excel(uploaded, engine="openpyxl", dtype=str)
except Exception as e:
    st.warning("⚠️ openpyxl failed — trying raw recovery mode...")

    import zipfile, tempfile, re
    from xml.etree import ElementTree as ET

    try:
        # Extract sheet XML manually (Excel files are ZIP archives)
        with tempfile.TemporaryDirectory() as tmpdir:
            with zipfile.ZipFile(uploaded, "r") as z:
                z.extractall(tmpdir)
            
            # Cari semua sheet
            sheet_files = [f for f in z.namelist() if f.startswith("xl/worksheets/sheet")]
            if not sheet_files:
                st.error("No valid sheets found in the uploaded Excel file.")
                st.stop()

            # Ambil sheet pertama yang masih valid XML
            sheet_path = None
            for f in sheet_files:
                try:
                    ET.parse(f"{tmpdir}/{f}")
                    sheet_path = f"{tmpdir}/{f}"
                    break
                except ET.ParseError:
                    continue

            if not sheet_path:
                st.error("All sheet XMLs are corrupted beyond recovery.")
                st.stop()

            # Parse XML sheet jadi DataFrame sederhana
            tree = ET.parse(sheet_path)
            root = tree.getroot()

            # ambil semua cell
            rows = []
            for row in root.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row"):
                values = []
                for cell in row:
                    val = cell.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v")
                    values.append(val.text if val is not None else "")
                rows.append(values)

            if not rows:
                st.error("No readable data found in the Excel sheet.")
                st.stop()

            headers = [str(x).strip() for x in rows[0]]
            report_df = pd.DataFrame(rows[1:], columns=headers)
            report_df = report_df.astype(str)

    except Exception as e2:
        st.error(f"Failed to recover Excel file: {e}\n\n{e2}")
        st.stop()

# Holidays handling
if holidays_file is not None:
    try:
        holidays_df = pd.read_excel(holidays_file, engine="openpyxl", dtype=str)
    except Exception as e:
        st.error(f"Failed to read holidays Excel: {e}")
        st.stop()
else:
    holidays_df = pd.DataFrame(columns=["Tanggal", "Hari", "Keterangan"])
    st.warning("No holidays file uploaded. Only weekends (Sat/Sun) will be considered non-business days.")

# Standardize and compute
report_df = standardize_input_df(report_df)
processed = compute_sla_and_status(report_df, holidays_df, pd.to_datetime(ref_date))

st.markdown("### Preview")

st.caption("You can filter or sort directly from the table below.")
st.data_editor(
    processed.head(50),
    use_container_width=True,
    num_rows="dynamic",
    column_config=None,
    hide_index=True
)

# Option: show only Telat rows for manual input
st.markdown("---")
st.subheader("Manual Input for 'Telat' Rows")

st.markdown("### Manual input for Telat rows (recommended)")

df_for_edit = processed.copy()

# Filter rows yang Telat untuk manual input
telat_filter = df_for_edit["Status"] == "Telat"
telat_rows = df_for_edit[telat_filter].copy()

if len(telat_rows) == 0:
    st.info("No rows with Status = 'Telat' found.")
else:
    st.markdown("**Rows with `Telat` — fill `Nomor SAP PJUM Pertama Kali` and `Tanggal Diterima GPBN` if available.**")
    st.caption("Edit cells inline, then click 'Apply manual changes'. Date cells accept dd/mm/yyyy or yyyy-mm-dd.")

    if "Nomor SAP PJUM Pertama Kali" not in telat_rows.columns:
        telat_rows["Nomor SAP PJUM Pertama Kali"] = ""
    if "Tanggal Diterima GPBN" not in telat_rows.columns:
        telat_rows["Tanggal Diterima GPBN"] = pd.NaT

    show_cols = [c for c in telat_rows.columns if c in [
        "Assignment", "Doc SAP", "Doc SAP PJUM", "Header Text",
        "End Date", "SLA PJUM + 10/20HK", "Status",
        "Nomor SAP PJUM Pertama Kali", "Tanggal Diterima GPBN",
        "Status No SAP PJUM", "Status Final"
    ]]

    edited = st.data_editor(telat_rows[show_cols], num_rows="dynamic", use_container_width=True)

    if st.button("Apply manual changes"):
        if "Assignment" in edited.columns and "Assignment" in df_for_edit.columns:
            edited_map = edited.set_index("Assignment")
            df_for_edit = df_for_edit.set_index("Assignment")
            for aid in edited_map.index:
                for col in ["Nomor SAP PJUM Pertama Kali", "Tanggal Diterima GPBN"]:
                    if col in edited_map.columns:
                        df_for_edit.at[aid, col] = edited_map.at[aid, col]
            df_for_edit = df_for_edit.reset_index()
        else:
            idxs = df_for_edit[df_for_edit["Status"] == "Telat"].index
            for i, idx in enumerate(idxs):
                for col in ["Nomor SAP PJUM Pertama Kali", "Tanggal Diterima GPBN"]:
                    if col in edited.columns:
                        df_for_edit.at[idx, col] = edited.iloc[i][col]

        if "Tanggal Diterima GPBN" in df_for_edit.columns:
            df_for_edit["Tanggal Diterima GPBN"] = df_for_edit["Tanggal Diterima GPBN"].apply(parse_date)

        df_for_edit = compute_status_after_manual(df_for_edit)

        st.success("Manual inputs applied and statuses recomputed.")
        st.dataframe(df_for_edit[df_for_edit["Status"] == "Telat"].head(50), use_container_width=True)

        # Sekalian tampilkan summary setelah manual input
        st.markdown("### Summary Setelah Manual Input")
        status_counts = df_for_edit["Status Final"].replace("", np.nan).fillna(df_for_edit["Status"]).value_counts()
        total_rows = len(df_for_edit)
        st.write(f"**Total Rows: {total_rows}**")
        st.write(status_counts.to_frame("Jumlah").reset_index().rename(columns={"index": "Status"}))

# After possible manual changes recompute final outputs
# If Status Final has non-empty values, drop original Status column and keep final as requested
final_df = df_for_edit.copy()

# If any 'Status Final' non-empty, replace
if final_df["Status Final"].replace("", np.nan).notna().any():
    # drop old Status
    if "Status" in final_df.columns:
        final_df = final_df.drop(columns=["Status"])
    # rename Status Final -> Status Final (keep as is)
else:
    # remove Status Final column if all empty to prevent clutter
    if "Status Final" in final_df.columns:
        if final_df["Status Final"].replace("", np.nan).isna().all():
            final_df = final_df.drop(columns=["Status Final"])

# ensure column order: original columns + requested appended columns in order
# Build base columns as those in original upload, then append requested
orig_cols = [c for c in report_df.columns]
append_order = ["SLA PJUM + 10/20HK","Status","Nomor SAP PJUM Pertama Kali","Tanggal Diterima GPBN","Status No SAP PJUM","Status Final"]
cols_to_keep = orig_cols + [c for c in append_order if c in final_df.columns]
final_df = final_df.reindex(columns=cols_to_keep)

# Ensure displayed date format
date_cols = [c for c in final_df.columns if 'Date' in c or 'Tanggal' in c]
for c in date_cols:
    final_df[c] = final_df[c].dt.strftime("%d/%m/%y")

st.markdown("---")
st.subheader("Final preview (first 100 rows)")
st.dataframe(final_df.head(100), use_container_width=True)
# -----------------------------
# DOWNLOAD SECTION
# -----------------------------
st.markdown("### Download Data")

col_dl1, col_dl2 = st.columns(2)

with col_dl1:
    st.markdown("**⬇️ Download Sebelum Manual Input**")
    bio_before = to_excel_bytes(processed)
    st.download_button(
        label="Download Sebelum Manual Input (.xlsx)",
        data=bio_before,
        file_name=f"PJUM_Period_before_manual_{datetime.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

with col_dl2:
    st.markdown("**⬇️ Download Setelah Manual Input / Apply Changes**")
    bio_after = to_excel_bytes(final_df)
    st.download_button(
        label="Download Setelah Manual Input (.xlsx)",
        data=bio_after,
        file_name=f"PJUM_Period_after_manual_{datetime.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

st.markdown("### Notes & tips")
st.markdown("""
- Make sure your uploaded BPM report has the expected column names starting in cell A1 (Assignment, Doc SAP, Header Text, Posting Date, End Date, Posting Date PJUM, Doc SAP PJUM, ...).
- Holidays file must have column 'Tanggal' with dates; the app will skip those dates when computing SLA.
- Reference date defaults to today; change it if you want to re-evaluate statuses for a different date.
- Manual edits are applied only after clicking **Apply manual changes**.
""")
