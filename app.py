import streamlit as st
import pandas as pd
from io import BytesIO

# ===== CONFIG =====
# Colors per shelf.
# A, B, C, F, H, Others are from your QDE5-style template.
# D and I are explicitly set so they are not white.
SHELF_COLORS = {
    "A": "#00B050",
    "B": "#0070C0",
    "C": "#FFFF00",
    "D": "#9B30FF",   # purple (custom, change if needed)
    "E": None,
    "F": "#CC0000",
    "G": None,
    "H": "#FFC000",
    "I": "#996600",   # brown (from original palette)
    "J": None,
    "K": None,
    "L": None,
    "M": None,
    "N": None,
    "O": None,
    "Others": "#000000",
}

SHELF_ORDER = list("ABCDEFGHIJKLMNO") + ["Others"]


def detect_shelf(label: str) -> str:
    """
    Shelf is always digit number 9 in the label ID.
    Example: HAZ-A101I123 -> 9th char = 'I' -> shelf I.
    If 9th char is not A–O, send it to Others.
    """
    if label is None:
        return "Others"

    text = str(label).strip()
    if len(text) < 9:
        return "Others"

    ch = text[8]  # 0-based index → 9th character
    ch_up = ch.upper()

    if ch_up in SHELF_ORDER[:-1]:  # A–O
        return ch_up

    return "Others"


def build_layout(labels):
    """
    Take a flat list of labels, group by shelf, and arrange into
    columns A–O + Others, each column filled top-down.
    Returns a pandas DataFrame ready to write to Excel.
    """
    groups = {shelf: [] for shelf in SHELF_ORDER}

    for label in labels:
        if label is None or str(label).strip() == "":
            continue
        shelf = detect_shelf(label)
        if shelf not in groups:
            shelf = "Others"
        groups[shelf].append(str(label).strip())

    max_len = max((len(vals) for vals in groups.values()), default=0)

    data = {}
    for shelf in SHELF_ORDER:
        col_vals = groups[shelf]
        col_vals = col_vals + [""] * (max_len - len(col_vals))
        data[shelf] = col_vals

    df_out = pd.DataFrame(data)
    return df_out


def to_excel_with_colors(df_out: pd.DataFrame) -> bytes:
    """
    Write DataFrame to Excel with:
    - Row 1: color hex code (#RRGGBB) for each shelf column, with background color (if defined)
    - Row 2: shelf letter (A, B, C, ... Others), with same background color (if defined)
    - Data starting from row 3
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # Write data WITHOUT headers, starting at row 2 (Excel row 3)
        df_out.to_excel(
            writer,
            sheet_name="Sheet1",
            index=False,
            header=False,
            startrow=2
        )

        workbook = writer.book
        worksheet = writer.sheets["Sheet1"]

        for col_idx, shelf in enumerate(SHELF_ORDER):
            color_hex = SHELF_COLORS.get(shelf)

            # Format for row 1 (color code)
            row1_format = workbook.add_format(
                {
                    "align": "center",
                    "valign": "vcenter",
                    "border": 1,
                }
            )
            # Format for row 2 (shelf header)
            row2_format = workbook.add_format(
                {
                    "bold": True,
                    "align": "center",
                    "valign": "vcenter",
                    "border": 1,
                }
            )

            if color_hex:
                row1_format.set_bg_color(color_hex)
                row2_format.set_bg_color(color_hex)
                # Row 1: hex code text
                worksheet.write(0, col_idx, color_hex, row1_format)
            else:
                # No color defined: keep row 1 blank, just border
                worksheet.write(0, col_idx, "", row1_format)

            # Row 2: shelf letter
            worksheet.write(1, col_idx, shelf, row2_format)

        # Set column width
        worksheet.set_column(0, len(SHELF_ORDER) - 1, 22)

    output.seek(0)
    return output.getvalue()


# ===== STREAMLIT APP =====

st.title("Shelf-based Bin Label Layout (Excel Upload Only)")

st.write(
    """
Upload an **Excel file (.xlsx / .xls)** with bin label IDs.

This app will:
- Detect the shelf from **digit 9** of each label ID.
- Place labels into columns **A–O** or **Others**.
- Generate an Excel file where:
  - **Row 1** = color hex code (background colored).
  - **Row 2** = shelf letter (background colored).
  - **Row 3+** = label IDs.
"""
)

uploaded_file = st.file_uploader(
    "Upload Excel file with label IDs",
    type=["xlsx", "xls"],
)

labels = []

if uploaded_file is not None:
    try:
        # Read all sheets, let user pick one
        xls = pd.ExcelFile(uploaded_file)
        sheet_name = st.selectbox("Select sheet to use", xls.sheet_names)

        df_in = pd.read_excel(xls, sheet_name=sheet_name)

        if df_in.empty:
            st.error("Selected sheet is empty.")
        else:
            label_column_name = st.selectbox(
                "Select the column that contains label IDs",
                df_in.columns,
            )

            labels = df_in[label_column_name].tolist()
            st.write(f"Detected **{len(labels)}** labels from sheet **{sheet_name}**.")

    except Exception as e:
        st.error(f"Failed to read Excel file: {e}")

generate = st.button("Generate shelf layout Excel")

if generate:
    if not labels:
        st.error("No labels found. Please upload a file and select the correct sheet/column.")
    else:
        df_out = build_layout(labels)
        excel_bytes = to_excel_with_colors(df_out)

        st.success("Excel generated successfully.")
        st.dataframe(df_out.head(20))  # Preview

        st.download_button(
            label="Download shelf layout Excel",
            data=excel_bytes,
            file_name="shelf_labels_layout.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
