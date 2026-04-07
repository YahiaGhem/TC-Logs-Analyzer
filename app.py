import io
import re
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Complaint Workbook Generator", layout="wide")

st.title("Complaint Workbook Generator")
st.write(
    "Upload the TC Logs Excel file and the Complaints CSV file. "
    "The app will combine them, remove exact duplicate complaints, "
    "build the summary sheets, and return a formatted Excel workbook."
)

# ---------------------------
# Helpers
# ---------------------------
REQUIRED_MAIN_COLS = [
    "VIN", "VMAKE", "VMODEL", "MODEL_YR",
    "DC_PREFIX", "DC_NAME", "DC_FAULT", "COMMENT"
]

def normalize_text_columns(df: pd.DataFrame) -> pd.DataFrame:
    for col in ["VIN", "VMAKE", "VMODEL", "DC_PREFIX", "DC_NAME", "DC_FAULT", "COMMENT", "SOURCE"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df.loc[df[col].isin(["nan", "None", ""]), col] = pd.NA
    return df

def standardize_csv_columns(df_csv: pd.DataFrame) -> pd.DataFrame:
    if "DATE_REP" in df_csv.columns and "DATE_REPORT" not in df_csv.columns:
        df_csv = df_csv.rename(columns={"DATE_REP": "DATE_REPORT"})
    if "ODOMETER" in df_csv.columns and "ODOMETER (KM)" not in df_csv.columns:
        df_csv = df_csv.rename(columns={"ODOMETER": "ODOMETER (KM)"})
    if "RN" in df_csv.columns and "Column1" not in df_csv.columns:
        df_csv = df_csv.rename(columns={"RN": "Column1"})
    return df_csv

def validate_columns(df: pd.DataFrame, label: str) -> None:
    missing = [col for col in REQUIRED_MAIN_COLS if col not in df.columns]
    if missing:
        raise ValueError(f"{label} is missing required columns: {', '.join(missing)}")

def style_worksheet(ws):
    header_fill = PatternFill("solid", start_color="1F3864", end_color="1F3864")
    alt_fill = PatternFill("solid", start_color="DCE6F1", end_color="DCE6F1")

    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    body_font = Font(name="Arial", size=10)

    center_align = Alignment(horizontal="center", vertical="center")
    wrap_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    thin = Side(style="thin", color="B0B0B0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = border

    ws.row_dimensions[1].height = 22

    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        fill = alt_fill if row_idx % 2 == 0 else PatternFill()

        for cell in row:
            cell.font = body_font
            cell.fill = fill
            cell.border = border
            cell.alignment = wrap_align

        ws.row_dimensions[row_idx].height = 24

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max((len(str(c.value or "")) for c in col), default=0)
        ws.column_dimensions[col_letter].width = min(max_len + 4, 55)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

def build_output_workbook(xlsx_file, csv_file) -> tuple[bytes, dict]:
    # Read uploads
    df_xlsx = pd.read_excel(xlsx_file)
    df_csv = pd.read_csv(csv_file)

    # Standardize CSV naming differences
    df_csv = standardize_csv_columns(df_csv)

    # Validate
    validate_columns(df_xlsx, "TC Logs Excel file")
    validate_columns(df_csv, "Complaints CSV file")

    # Add source labels
    df_xlsx["SOURCE"] = "TC Logs"
    df_csv["SOURCE"] = "Complaints_CSV"

    # Align columns, then combine
    all_columns = sorted(set(df_xlsx.columns).union(set(df_csv.columns)))
    df_xlsx = df_xlsx.reindex(columns=all_columns)
    df_csv = df_csv.reindex(columns=all_columns)
    df = pd.concat([df_xlsx, df_csv], ignore_index=True)

    # Clean text
    df = normalize_text_columns(df)

    # Remove exact duplicate complaints across both files
    df = df.drop_duplicates(subset=[
        "VIN",
        "VMAKE",
        "VMODEL",
        "MODEL_YR",
        "DC_PREFIX",
        "DC_NAME",
        "DC_FAULT",
        "COMMENT"
    ]).reset_index(drop=True)

    # ---------------------------
    # Sheet 1: Filtered Logs
    # ---------------------------
    df_sheet1 = df[[
        "SOURCE",
        "VIN",
        "VMAKE",
        "VMODEL",
        "MODEL_YR",
        "DC_PREFIX",
        "DC_NAME",
        "DC_FAULT",
        "COMMENT"
    ]].copy()

    df_sheet1 = df_sheet1.rename(columns={
        "VMAKE": "MAKE",
        "VMODEL": "MODEL",
        "MODEL_YR": "MODEL YEAR",
        "DC_PREFIX": "DC PREFIX",
        "DC_NAME": "DC NAME",
        "DC_FAULT": "DC FAULT",
        "COMMENT": "COMPLAINT"
    })

    model_counts = df_sheet1["MODEL"].value_counts(dropna=False)
    df_sheet1["MODEL_COUNT"] = df_sheet1["MODEL"].map(model_counts)

    df_sheet1 = df_sheet1.sort_values(
        by=["MODEL_COUNT", "MODEL", "MODEL YEAR", "DC PREFIX", "DC FAULT"],
        ascending=[False, True, False, True, True]
    ).drop(columns=["MODEL_COUNT"]).reset_index(drop=True)

    # ---------------------------
    # Sheet 2: Vehicle Year Counts
    # ---------------------------
    df_sheet2_base = df[[
        "VIN",
        "VMAKE",
        "VMODEL",
        "MODEL_YR"
    ]].copy()

    df_sheet2_base = df_sheet2_base.dropna(subset=["VIN", "VMAKE", "VMODEL", "MODEL_YR"])

    df_sheet2_base = df_sheet2_base.drop_duplicates(
        subset=["VIN", "VMAKE", "VMODEL", "MODEL_YR"]
    )

    df_sheet2 = (
        df_sheet2_base.groupby(["VMAKE", "VMODEL", "MODEL_YR"])
        .size()
        .reset_index(name="TOTAL")
    )

    df_sheet2 = df_sheet2.rename(columns={
        "VMAKE": "MAKE",
        "VMODEL": "MODEL",
        "MODEL_YR": "MODEL YEAR"
    })

    df_sheet2 = df_sheet2.sort_values(
        by=["TOTAL", "MAKE", "MODEL", "MODEL YEAR"],
        ascending=[False, True, True, True]
    ).reset_index(drop=True)

    # ---------------------------
    # Sheet 3: Vehicle Year Fault Counts
    # ---------------------------
    df_sheet3_base = df[[
        "VMAKE",
        "VMODEL",
        "MODEL_YR",
        "DC_FAULT"
    ]].copy()

    df_sheet3_base = df_sheet3_base.dropna(subset=["VMAKE", "VMODEL", "MODEL_YR", "DC_FAULT"])

    df_sheet3 = (
        df_sheet3_base.groupby(["VMAKE", "VMODEL", "MODEL_YR", "DC_FAULT"])
        .size()
        .reset_index(name="TOTAL")
    )

    df_sheet3 = df_sheet3.rename(columns={
        "VMAKE": "MAKE",
        "VMODEL": "MODEL",
        "MODEL_YR": "MODEL YEAR",
        "DC_FAULT": "DC FAULT"
    })

    df_sheet3 = df_sheet3.sort_values(
        by=["TOTAL", "MAKE", "MODEL", "MODEL YEAR", "DC FAULT"],
        ascending=[False, True, True, True, True]
    ).reset_index(drop=True)

    # ---------------------------
    # Sheet 4: Recall Counts
    # ---------------------------
    recall_pattern = re.compile(r"\b\d{2}[A-Za-z]\b")
    comments = df["COMMENT"].dropna().astype(str)

    recall_matches = []
    for comment in comments:
        matches = recall_pattern.findall(comment.upper())
        for match in matches:
            recall_matches.append(match)

    df_sheet4 = pd.DataFrame({"RECALL": recall_matches})

    if not df_sheet4.empty:
        df_sheet4["RECALL"] = "Recall " + df_sheet4["RECALL"]
        df_sheet4 = (
            df_sheet4.groupby("RECALL")
            .size()
            .reset_index(name="TOTAL")
            .sort_values(by=["TOTAL", "RECALL"], ascending=[False, True])
            .reset_index(drop=True)
        )
    else:
        df_sheet4 = pd.DataFrame(columns=["RECALL", "TOTAL"])

    # ---------------------------
    # Write workbook to memory
    # ---------------------------
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_sheet1.to_excel(writer, sheet_name="Filtered Logs", index=False)
        df_sheet2.to_excel(writer, sheet_name="Vehicle Year Counts", index=False)
        df_sheet3.to_excel(writer, sheet_name="Vehicle Year Fault Counts", index=False)
        df_sheet4.to_excel(writer, sheet_name="Recall Counts", index=False)

    output.seek(0)

    wb = load_workbook(output)
    for sheet_name in wb.sheetnames:
        style_worksheet(wb[sheet_name])

    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    stats = {
        "combined_rows_after_dedup": len(df),
        "sheet1_rows": len(df_sheet1),
        "sheet2_rows": len(df_sheet2),
        "sheet3_rows": len(df_sheet3),
        "sheet4_rows": len(df_sheet4),
    }
    return final_output.getvalue(), stats

# ---------------------------
# UI
# ---------------------------
col1, col2 = st.columns(2)

with col1:
    xlsx_upload = st.file_uploader(
        "Upload TC Logs Excel file",
        type=["xlsx"],
        accept_multiple_files=False
    )

with col2:
    csv_upload = st.file_uploader(
        "Upload Complaints CSV file",
        type=["csv"],
        accept_multiple_files=False
    )

generate = st.button("Generate workbook", type="primary")

if generate:
    if xlsx_upload is None or csv_upload is None:
        st.error("Please upload both files first.")
    else:
        try:
            with st.spinner("Processing files..."):
                file_bytes, stats = build_output_workbook(xlsx_upload, csv_upload)

            st.success("Workbook created successfully.")

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Combined rows", stats["combined_rows_after_dedup"])
            c2.metric("Vehicle/Year rows", stats["sheet2_rows"])
            c3.metric("Vehicle/Year/Fault rows", stats["sheet3_rows"])
            c4.metric("Recall rows", stats["sheet4_rows"])

            st.download_button(
                label="Download processed workbook",
                data=file_bytes,
                file_name="TC_Logs_Filtered.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error: {e}")