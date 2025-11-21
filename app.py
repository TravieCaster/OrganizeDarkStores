import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook

# ===== COLOR EXTRACTION HELPERS =====

def get_cell_color_hex(cell):
    """
    Get the cell's background color as a #RRGGBB hex string.
    If no usable RGB color, return None.
    """
    fill = cell.fill
    if fill is None:
        return None

    fg = fill.fgColor
    if fg is None:
        return None

    # Most manual fills in Excel come through as rgb = "FFRRGGBB"
    if fg.type == "rgb" and fg.rgb:
        rgb = fg.rgb
        # Handle ARGB (e.g. "FFRRGGBB") → take last 6 chars (RRGGBB)
        if len(rgb) == 8:
            rgb = rgb[2:]
        if len(rgb) == 6:
            return "#" + rgb.upper()

    # If indexed/theme/etc, we skip exact mapping for simplicity
    return None


def extract_color_groups(workbook):
    """
    For each sheet, build a dict:
    sheet_name -> { color_hex -> [labels] }

    - Scans all cells.
    - Uses non-empty cell values as labels.
    - Groups by background color.
    """
    sheets_dict = {}

    for ws in workbook.worksheets:
        color_groups = {}

        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if value is None:
                    continue

                text = str(value).strip()
                if not text:
                    continue

                color_hex = get_cell_color_hex(cell)
                # If no color, treat as white (or generic) group so nothing is lost
                if color_hex is None:
                    color_hex = "#FFFFFF"

                if color_hex not in color_groups:
                    color_groups[color_hex] = []

                color_groups[color_hex].append(text)

        sheets_dict[ws.title] = color_groups

    return sheets_dict


def build_df_from_color_groups(color_groups):
    """
    From { color_hex -> [labels] } create:
    - DataFrame with each column = color_hex, rows = labels.
    - Ordered list of colors (column order).
    """
    if not color_groups:
        return pd.DataFrame(), []

    # Sort colors to have deterministic column order
    colors = sorted(color_groups.keys())

    max_len = max(len(vals) for vals in color_groups.values())
    data = {}

    for color in colors:
        vals = color_groups[color]
        vals = vals + [""] * (max_len - len(vals))
        data[color] = vals

    df_out = pd.DataFrame(data)
    return df_out, colors


def write_grouped_workbook(sheets_dict):
    """
    Given:
      sheets_dict: { sheet_name -> { color_hex -> [labels] } }

    Create an Excel file in memory:
      - One sheet per input sheet.
      - Row 1: color hex text + background
      - Row 2: blank header row with same background (visual band)
      - Row 3+: labels in each color group column
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        for sheet_name, color_groups in sheets_dict.items():
            df_out, colors = build_df_from_color_groups(color_groups)

            # If sheet had no labels, just create an empty sheet
            if df_out.empty:
                df_out.to_excel(writer, sheet_name=sheet_name, index=False)
                continue

            # Write data WITHOUT headers, starting at row 2 (Excel row 3)
            df_out.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False,
                header=False,
                startrow=2,
            )

            worksheet = writer.sheets[sheet_name]

            # Row 0: color hex code text with colored background
            # Row 1: blank row, same background (visual header band)
            for col_idx, color_hex in enumerate(colors):
                # Format for row 0 (color code)
                header1_format = workbook.add_format(
                    {
                        "align": "center",
                        "valign": "vcenter",
                        "border": 1,
                    }
                )
                # Format for row 1 (blank, but colored)
                header2_format = workbook.add_format(
                    {
                        "bold": True,
                        "align": "center",
                        "valign": "vcenter",
                        "border": 1,
                    }
                )

                # Only set bg if color looks like #RRGGBB
                if isinstance(color_hex, str) and color_hex.startswith("#") and len(color_hex) == 7:
                    header1_format.set_bg_color(color_hex)
                    header2_format.set_bg_color(color_hex)

                # Row 0: write color hex
                worksheet.write(0, col_idx, color_hex, header1_format)
                # Row 1: blank but colored
                worksheet.write(1, col_idx, "", header2_format)

            # Set nice column width
            worksheet.set_column(0, len(colors) - 1, 22)

    output.seek(0)
    return output.getvalue()


# ===== STREAMLIT APP =====

st.title("Bin Label Color Grouper (by cell color)")

st.write(
    """
Upload an **Excel file (.xlsx)** where each sheet contains bin labels
with different **cell background colors**.

This app will:
- Read each sheet in the file.
- Group labels by **cell background color**.
- Create a new Excel where:
  - Each sheet mirrors the original sheet name.
  - Each **column = one color** from that sheet.
  - **Row 1** shows the color hex code (e.g. `#339900`) and is filled with that color.
  - **Row 2** is a blank header row, also filled with that color.
  - From **Row 3 down** are the labels that had that color.
"""
)

uploaded_file = st.file_uploader(
    "Upload Excel file (.xlsx) with colored bin labels",
    type=["xlsx"],
)

if uploaded_file is not None:
    st.info("File uploaded. Click **Generate grouped Excel** to process.")

generate = st.button("Generate grouped Excel")

if generate:
    if uploaded_file is None:
        st.error("Please upload an Excel (.xlsx) file first.")
    else:
        try:
            # Load workbook from uploaded file
            wb = load_workbook(uploaded_file, data_only=True)
        except Exception as e:
            st.error(f"Failed to read Excel file: {e}")
        else:
            # Extract color groups per sheet
            sheets_dict = extract_color_groups(wb)

            # Build output workbook
            excel_bytes = write_grouped_workbook(sheets_dict)

            # Simple preview: show first processed sheet if any
            first_sheet_name = next(iter(sheets_dict.keys()), None)
            if first_sheet_name:
                df_preview, colors_preview = build_df_from_color_groups(
                    sheets_dict[first_sheet_name]
                )
                if not df_preview.empty:
                    st.write(f"Preview from sheet: **{first_sheet_name}**")
                    st.dataframe(df_preview.head(20))

            st.success("Grouped Excel generated successfully.")

            st.download_button(
                label="Download grouped Excel",
                data=excel_bytes,
                file_name="bin_labels_grouped_by_color.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
