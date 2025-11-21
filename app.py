import streamlit as st
import pandas as pd
from io import BytesIO

# ===== CONFIG =====

# Shelf header colors copied from your Furjan.xlsx
SHELF_COLORS = {
    "A": "#FFFFFF",
    "B": "#339900",
    "C": "#9B30FF",
    "D": "#FFFF00",
    "E": "#00FFFF",
    "F": "#CC0000",
    "G": "#F88017",
    "H": "#FF00FF",
    "I": "#996600",
    "J": "#00FF00",
    "K": "#FF6565",
    "L": "#9999FE",
    "M": "#C7721C",
    "N": "#F7B0BB",
    "O": "#C6F700",
    "Others": None,  # no fill
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

    if ch_up in SHELF_COLORS and ch_up != "Others":
        return ch_up

    return "Others"


def build_layout(labels):
    """
    Take a flat list of labels, group by shelf, and arrange into
    columns A–O + Others, each column filled top-down.
    Returns a pandas DataFrame ready to write to Excel.
    """
    # Initialize groups
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
        # pad with empty strings to same length
        col_vals = col_vals + [""] * (max_len - len(col_vals))
        data[shelf] = col_vals

    df_out = pd.DataFrame(data)
    return df_out


def to_excel_with_colors(df_out: pd.DataFrame) -> bytes:
    """
    Write DataFrame to Excel with header colors matching your template.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # Start data from row 1, we’ll write headers manually on row 0
        df_out.to_excel(writer, sheet_name="Sheet1", index=False, startrow=1)

        workbook = writer.book
        worksheet = writer.sheets["Sheet1"]

        # Write headers with colors
        for col_idx, col_name in enumerate(SHELF_ORDER):
            header_format = workbook.add_format(
                {
                    "bold": True,
                    "align": "center",
                    "valign": "vcenter",
                    "border": 1,
                }
            )

            color = SHELF_COLORS.get(col_name)
            if color:
                header_format.set_bg_color(color)

            worksheet.write(0, col_idx, col_name, header_format)

        # Set a decent column width
        worksheet.set_column(0, len(SHELF_ORDER) - 1, 22)

    output.seek(0)
    return output.getvalue()


# ===== STREAMLIT APP =====

st.title("Shelf Label Color Coder (Furjan Style)")

st.write(
    """
Paste a list of label IDs or upload a file, and this app will:
- Detect the shelf from **digit 9** of the label ID.
- Put labels with shelves A–O into their shelf columns.
- Put all others into the **Others** column.
- Generate an Excel file with **the same header colors** as your Furjan sheet.
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
        # split by newline or comma
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
            # take first column by default
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
            file_name="shelf_labels_colored.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
