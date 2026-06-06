import datetime
import io
import os
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
import streamlit as st

# -------- CONFIGURATION --------
# તમારી સાચી ગૂગલ શીટની લિંક
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1MbGVYp71KS9RkK07V4Lqmuq6QSN9j7fAsbxJrYCUnXY/edit?usp=sharing"

st.set_page_config(page_title="Spool Detail Report Generator", layout="wide")
st.title("📊 Spool Detail Report Generator")


# -------- GOOGLE SHEET FETCHER --------
def get_web_dataframe(url, sheet_name="Sheet2"):
    try:
        # લિંક માંથી સાચી File ID અલગ કરવી
        if "spreadsheets/d/" in url:
            file_id = url.split("spreadsheets/d/")[1].split("/")[0]
        elif "id=" in url:
            file_id = url.split("id=")[1].split("&")[0]
        else:
            st.error("❌ ગૂગલ શીટની લિંકનું ફોર્મેટ ખોટું છે.")
            return None

        # એક્સેલ ફોર્મેટમાં એક્સપોર્ટ કરવાની ડાયરેક્ટ લિંક
      # નવી સુધારેલી લાઈન:
d_url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx&sheet={sheet_name}"
        # પંડાસ દ્વારા સીધો લાઈવ ડેટા રીડ કરવો
        df = pd.read_excel(d_url, sheet_name=sheet_name)
        df.columns = df.columns.str.strip()
        return df

    except Exception as e:
        st.error(f"❌ ગૂગલ શીટ ડેટા લોડ કરવામાં ભૂલ: {e}")
        return None


# -------- HELPERS --------
def clean_val(x):
    if pd.isna(x) or str(x).strip().lower() in ["na", "nan", "none", ""]:
        return ""
    if isinstance(x, (pd.Timestamp, datetime.date)):
        return x.strftime("%d-%m-%Y")
    try:
        if isinstance(x, (float, int)):
            val = float(x)
            return str(int(val)) if val.is_integer() else str(val)
    except:
        pass
    return str(x).strip()


def get_ndt_info(
    full_df, lot_no, lot_col, rep_col, date_col, test_type, xr_col=None
):
    if not lot_no:
        return None, ""

    target_lot = clean_val(lot_no)
    matches = full_df[full_df[lot_col].apply(clean_val) == target_lot]

    for _, r in matches.iterrows():
        rep_no = clean_val(r.get(rep_col, ""))
        rep_date = r.get(date_col, "")
        xr_no = clean_val(r.get(xr_col, "")) if xr_col else ""

        if rep_no and rep_date:
            fmt_date = clean_val(rep_date)

            iso_no = clean_val(r.get("ISO No/Drawing No/Line No", ""))
            spool_no = clean_val(r.get("Spool Unique No.", ""))
            joint_no = clean_val(r.get("Joint No.", ""))

            extra_parts = []
            if iso_no:
                extra_parts.append(iso_no)
            if spool_no:
                extra_parts.append(spool_no)
            if joint_no:
                extra_parts.append(joint_no)

            extra_text = f" ({' | '.join(extra_parts)})" if extra_parts else ""

            remark = f"For {test_type} LOT No {target_lot} SUPPORTING Report No is {rep_no}"
            if xr_no:
                remark += f", XR No is {xr_no}"
            remark += extra_text
            remark += f" and Date is {fmt_date}"

            return remark, "fully cleared"

    return None, "offered"


# -------- PDF BUILDER FUNCTION --------
def generate_pdf_bytes(usr_no, usr_df, full_df, existing_columns):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=15,
        leftMargin=15,
        topMargin=15,
        bottomMargin=15,
    )

    elements = []

    title_style = ParagraphStyle(
        "Title",
        fontSize=14,
        textColor=HexColor("#2C3E50"),
        alignment=1,
        spaceAfter=12,
    )
    body_style = ParagraphStyle("Body", fontSize=6.5, leading=9, alignment=1)
    header_style = ParagraphStyle(
        "Header",
        fontSize=7,
        leading=8,
        alignment=1,
        textColor=colors.whitesmoke,
        fontName="Helvetica-Bold",
    )

    elements.append(Paragraph(f"SPOOL DETAIL REPORT - {usr_no}", title_style))

    table_data = [[Paragraph(col, header_style) for col in existing_columns]]
    final_remarks = set()

    for _, row in usr_df.iterrows():
        joint_type = clean_val(row.get("Type of Joint", "")).upper()
        current_ndt_status = ""
        lot_override = None

        if joint_type == "EB":
            lot_val = clean_val(row.get("Induction bend  DPT Lot no", ""))
            rmk, status = get_ndt_info(
                full_df,
                lot_val,
                "Induction bend  DPT Lot no",
                "Induction bend  DPT Report No",
                "Induction bend  DPT Date",
                "Induction DPT",
            )
            if lot_val:
                current_ndt_status = status
                if rmk:
                    final_remarks.add(rmk)

        elif joint_type in ["LET", "SOF", "SOB"]:
            raw_percent = row.get("DPT %", "")
            try:
                if pd.isna(raw_percent) or raw_percent == "":
                    percent_value = 0
                else:
                    percent_value = float(raw_percent)
                    if percent_value > 1:
                        percent_value = percent_value / 100
            except:
                percent_value = 0

            if percent_value == 1:
                lot_override = "100%"
                rep_no = clean_val(row.get("DPT REPORT NO", ""))
                rep_date = clean_val(row.get("DPT DATE", ""))

                if rep_no and rep_date:
                    current_ndt_status = "fully cleared"
                    iso_no = clean_val(row.get("ISO No/Drawing No/Line No", ""))
                    spool_no = clean_val(row.get("Spool Unique No.", ""))
                    joint_no = clean_val(row.get("Joint No.", ""))

                    extra_parts = []
                    if iso_no:
                        extra_parts.append(iso_no)
                    if spool_no:
                        extra_parts.append(spool_no)
                    if joint_no:
                        extra_parts.append(joint_no)

                    extra_text = (
                        f" ({' | '.join(extra_parts)})" if extra_parts else ""
                    )
                    remark = f"For DPT 100% SUPPORTING Report No is {rep_no}{extra_text} and Date is {rep_date}"
                    final_remarks.add(remark)
                else:
                    current_ndt_status = "offered"
            else:
                lot_val = clean_val(row.get("DPT LOT NO", ""))
                rmk, status = get_ndt_info(
                    full_df,
                    lot_val,
                    "DPT LOT NO",
                    "DPT REPORT NO",
                    "DPT DATE",
                    "DPT",
                )
                if lot_val:
                    current_ndt_status = status
                    if rmk:
                        final_remarks.add(rmk)

        elif joint_type == "BW":
            lot_val = clean_val(row.get("RT LOT NO", ""))
            rmk, status = get_ndt_info(
                full_df,
                lot_val,
                "RT LOT NO",
                "RT REPORT NO",
                "RT DATE",
                "RT",
                xr_col="XR NO",
            )
            if lot_val:
                current_ndt_status = status
                if rmk:
                    final_remarks.add(rmk)

        formatted_row = []
        for col_name in existing_columns:
            if col_name == "NDT Status":
                val = current_ndt_status
            elif col_name == "DPT LOT NO" and lot_override:
                val = lot_override
            else:
                val = clean_val(row.get(col_name, ""))
            formatted_row.append(Paragraph(val if val else "-", body_style))

        table_data.append(formatted_row)

    column_widths = []
    for col in existing_columns:
        if "ISO No" in col:
            column_widths.append(130)
        elif "Date" in col or "DATE" in col:
            column_widths.append(60)
        elif "NDT Status" in col:
            column_widths.append(65)
        else:
            column_widths.append(42)

    table = Table(
        table_data,
        colWidths=column_widths,
        rowHeights=[40] * len(table_data),
        repeatRows=1,
    )
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), HexColor("#2C3E50")),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ]
        )
    )
    elements.append(table)

    if final_remarks:
        elements.append(Spacer(1, 20))
        for rmk in sorted(final_remarks):
            elements.append(
                Paragraph(
                    rmk,
                    ParagraphStyle(
                        "RemarkStyle",
                        fontSize=9,
                        fontName="Helvetica-Bold",
                        leading=12,
                    ),
                )
            )

    first_row = usr_df.iloc[0]
    ndt_clearance_date = clean_val(first_row.get("NDT CLEARANCE", ""))
    fd_date = clean_val(first_row.get("FD Date", ""))
    dc_no = clean_val(
        first_row.get("Shift to Laydown / Painting Yard DC No", "")
    )

    elements.append(Spacer(1, 15))
    extra_style = ParagraphStyle(
        "ExtraStyle", fontSize=9, fontName="Helvetica-Bold", leading=12
    )
    elements.append(
        Paragraph(f"NDT CLEARANCE DATE:- {ndt_clearance_date}", extra_style)
    )
    elements.append(Paragraph(f"FD DATE:- {fd_date}", extra_style))
    elements.append(Paragraph(f"DC No:- {dc_no}", extra_style))

    doc.build(elements)
    buffer.seek(0)
    return buffer


# -------- MAIN WEB APP LOGIC --------
# લાઈવ ગૂગલ શીટ માંથી ડેટા લોડ કરો
with st.spinner("গૂગલ શીટમાંથી લાઈવ ડેટા લોડ થઈ રહ્યો છે..."):
    df = get_web_dataframe(GOOGLE_SHEET_URL, sheet_name="Sheet2")

if df is not None:
    full_df = df.copy()

    # સાઇડબારમાં ઇનપુટ સેટિંગ્સ
    st.sidebar.header("🔍 સર્ચ પેનલ")
    usr_no = st.sidebar.text_input(
        "Spool Unique No. લખો:", placeholder="e.g., A-41101"
    ).strip()

    if usr_no:
        usr_df = df[df["Spool Unique No."].astype(str) == str(usr_no)].copy()

        if usr_df.empty:
            st.warning(f"⚠️ {usr_no} માટે કોઈ રેકોર્ડ મળ્યો નથી.")
        else:
            st.success(f"✅ {len(usr_df)} રેકોર્ડ્સ મળ્યા!")

            required_columns = [
                "ISO No/Drawing No/Line No",
                "Joint No.",
                "Type of Joint",
                "WELD NPD",
                "Spool Unique No.",
                "Induction bend  DPT Lot no",
                "FIT UP Date",
                "Welder No",
                "WELD VISUAL REPORT NO",
                "VISUAL Date",
                "DPT LOT NO",
                "RT REPORT NO",
                "XR NO",
                "RT LOT NO",
                "NDT Status",
            ]
            existing_columns = [
                col
                for col in required_columns
                if col in df.columns or col == "NDT Status"
            ]

            # લાઈવ ડેટા સ્ક્રીન પર બતાવવા માટે
            st.subheader("📋 લાઈવ ડેટા પ્રીવ્યૂ")
            st.dataframe(usr_df[existing_columns], use_container_width=True)

            # PDF જનરેટ અને ડાઉનલોડ બટન
            st.sidebar.markdown("---")
            st.sidebar.subheader("📥 રીપોર્ટ ડાઉનલોડ")

            pdf_data = generate_pdf_bytes(
                usr_no, usr_df, full_df, existing_columns
            )

            st.sidebar.download_button(
                label="📥 Download PDF Report",
                data=pdf_data,
                file_name=f"Spool_Report_{usr_no}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
