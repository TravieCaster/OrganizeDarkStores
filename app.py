import streamlit as st
import pandas as pd
from io import BytesIO

# ===== CONFIG FROM YOUR "Copy of QDE5 (1).xlsx" HEADER =====
# Only shelves that actually had a specific color in the file
SHELF_COLORS = {
    "A": "#00B050",
    "B": "#0070C0",
    "C": "#FFFF00",
    "F": "#CC0000",
    "H": "#FFC000",
    "Others": "#000000",
    # Shelves D, E, G, I, J, K, L, M, N, O had theme/default fills in your file,
    # so we leave them without explicit color (they’ll show as no fill).
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

        # Set a decent column width
        worksheet.set_column(0, len(SHELF_ORDER) - 1, 22)

    output.seek(0)
    return output.getvalue()


# ===== STREAMLIT APP =====

st.title("Shelf-based Bin Label Layout (QDE5 Style)")

st.write(
    """
This app:
- Detects the shelf from **digit 9** of each label ID.
- Places labels into columns **A–O** or **Others**.
- Colors the headers using the same scheme as your QDE5 template:
  - **Row 1**: color hex code (e.g. `#00B050`) with that background.
  - **Row 2**: shelf letter with same background.
  - **Row 3+**: labels.
"""
)

input_mode = st.radio(
    "How do you want to provide labels?",
    ("Paste list", "Upload Excel/CSV"),
    horizontal=True,
)

labels = []

if input_mode == "Paste list":
    text = st.text_area(
        "Paste labels here (one per line or separated by commas):",
        height=200,
        placeholder="Example:\nHAZ-A101A110\nHAZ-A101B110\nHAZ-A101I123\nSTOWAGE_1_A_001",
    )
    if text:
        raw = []
        for line in text.splitlines():
            for part in line.split(","):
                val = part.strip()
                if val:
                    raw.append(val)
        labels = raw

else:  # Upload file
    uploaded_file = st.file_uploader(
        "Upload an Excel or CSV file that contains label IDs",
        type=["xlsx", "xls", "csv"],
    )
    label_column_name = st.text_input(
        "Column name that contains labels (if blank, app will use the first column)",
        value="",
    )

    if uploaded_file is not None:
        if uploaded_file.name.lower().endswith(".csv"):
            df_in = pd.read_csv(uploaded_file)
        else:
            df_in = pd.read_excel(uploaded_file)

        if label_column_name and label_column_name in df_in.columns:
            labels = df_in[label_column_name].tolist()
        else:
            first_col = df_in.columns[0]
            labels = df_in[first_col].tolist()

        st.write("Detected labels:", len(labels))


generate = st.button("Generate Excel")

if generate:
    if not labels:
        st.error("No labels found. Please paste or upload labels first.")
    else:
        df_out = build_layout(labels)
        excel_bytes = to_excel_with_colors(df_out)

        st.success("Excel generated successfully.")
        st.dataframe(df_out.head(20))  # preview

        st.download_button(
            label="Download color-coded Excel",
            data=excel_bytes,
            file_name="shelf_labels_QDE5_style.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
