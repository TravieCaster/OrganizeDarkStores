import streamlit as st
import pandas as pd
from io import BytesIO

# ===== CONFIG: COLORS FROM "Copy of QDE5 (1).xlsx" HEADER =====
# These hex values are taken directly from the sheet header fills.
SHELF_COLORS = {
    "A": "#00B050",
    "B": "#0070C0",
    "C": "#FFFF00",
    "D": "#6FC5E6",
    "E": "#000000",
    "F": "#CC0000",
    "G": "#000000",
    "H": "#FFC000",
    "I": "#000000",
    "J": "#000000",
    "K": "#000000",
    "L": "#000000",
    "M": "#000000",
    "N": "#000000",
    "O": "#000000",
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
        if label is None:
            continue
        text = str(label).strip()
        if text == "":
            continue

        shelf = detect_shelf(text)
        if shelf not in groups:
            shelf = "Others"
        groups[shelf].append(text)

    max_len = max((len(vals) for vals in groups.values()), default=0)

    data = {}
    for shelf in SHELF_ORDER:
        col_vals = groups[shelf]
        col_vals = col_vals + [""] * (max_len - len(col_vals))
        data[shelf] = col_vals

    df_out = pd.DataFrame(data)
    return df_out


def write_output_workbook(sheets_labels: dict) -> bytes:
    """
    sheets_labels: dict { sheet_name -> list_of_labels }

    For each sheet in the input, create a corresponding sheet in the
    output workbook with:
    - Row 1: color hex code (#RRGGBB) with background color
    - Row 2: shelf letter (A..O, Others) with same background
    - Row 3+: label IDs
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        for sheet_name, labels in sheets_labels.items():
            df_out = build_layout(labels)

            # Write data WITHOUT headers, starting at row 2 (Excel row 3)
            df_out.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False,
                header=False,
                startrow=2,
            )

            worksheet = writer.sheets[sheet_name]

            for col_idx, shelf in enumerate(SHELF_ORDER):
                color_hex = SHELF_COLORS.get(shelf)

                # Row 1 (color code)
                row1_format = workbook.add_format(
                    {
                        "align": "center",
                        "valign": "vcenter",
                        "border": 1,
                    }
                )
                # Row 2 (shelf header)
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
                    # Put hex text in Row 1
                    worksheet.write(0, col_idx, color_hex, row1_format)
                else:
                    worksheet.write(0, col_idx, "", row1_format)

                # Shelf letter in Row 2
                worksheet.write(1, col_idx, shelf, row2_format)

            # Column widths
            worksheet.set_column(0, len(SHELF_ORDER) - 1, 22)

    output.seek(0)
    return output.getvalue()


# ===== STREAMLIT APP =====

st.title("Shelf Layout Generator (QDE5 Header Colours)")

st.write(
    """
Upload an **Excel file (.xlsx)** with bin label IDs.

For **each sheet** in the file:
- All non-empty cells are treated as label IDs.
- Shelf is taken from **digit 9** of the label.
- Labels are placed into columns **A–O** and **Others**.
- Headers use the **exact colours** from your QDE5 template:
  - **Row 1**: hex colour code with background.
  - **Row 2**: shelf letter with same background.
  - **Row 3+**: label IDs.
"""
)

uploaded_file = st.file_uploader(
    "Upload Excel file with label IDs (all sheets will be processed)",
    type=["xlsx"],
)

generate = st.button("Generate shelf layout Excel")

if generate:
    if uploaded_file is None:
        st.error("Please upload an Excel (.xlsx) file first.")
    else:
        try:
            # Read all sheets, no assumptions about columns.
            # header=None so we don't treat any row as header; we just take everything.
            xls_dict = pd.read_excel(uploaded_file, sheet_name=None, header=None)
        except Exception as e:
            st.error(f"Failed to read Excel file: {e}")
        else:
            sheets_labels = {}
            for sheet_name, df in xls_dict.items():
                # Flatten all non-empty cells as labels
                values = df.to_numpy().ravel()
                labels = [v for v in values if pd.notna(v) and str(v).strip() != ""]
                sheets_labels[sheet_name] = labels

            if not sheets_labels:
                st.error("No data found in the uploaded workbook.")
            else:
                # Build output workbook
                excel_bytes = write_output_workbook(sheets_labels)

                # Quick preview from first sheet
                first_sheet = next(iter(sheets_labels.keys()))
                preview_df = build_layout(sheets_labels[first_sheet])
                st.write(f"Preview from sheet: **{first_sheet}**")
                st.dataframe(preview_df.head(20))

                st.success("Shelf layout Excel generated successfully.")
                st.download_button(
                    label="Download shelf layout Excel",
                    data=excel_bytes,
                    file_name="shelf_labels_layout_QDE5_colours.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
