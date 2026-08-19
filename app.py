import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Breakout List Merger", page_icon="📊", layout="wide")

st.title("📊 Breakout List — Merge All Sheets Into One")
st.write(
    "Upload all your Breakout List Excel files (Sheet-1 to Sheet-7, or any number of files). "
    "The app will combine every sheet from every file into a single merged tab, "
    "remove duplicate rows, and let you download the final combined file."
)

uploaded_files = st.file_uploader(
    "Upload Excel files (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True,
    help="You can select and drop multiple files at once.",
)

# Optional: let user decide whether to drop exact duplicate rows
drop_duplicates = st.checkbox("Remove duplicate rows (exact matches)", value=True)


def load_and_combine(files):
    """Read every sheet of every uploaded file and stack them into one DataFrame."""
    all_frames = []
    file_summary = []

    for f in files:
        try:
            xls = pd.ExcelFile(f)
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                if df.empty:
                    continue
                df["Source File"] = f.name
                df["Source Sheet"] = sheet_name
                all_frames.append(df)
                file_summary.append(
                    {"File": f.name, "Sheet": sheet_name, "Rows": len(df)}
                )
        except Exception as e:
            st.error(f"Could not read {f.name}: {e}")

    if not all_frames:
        return None, pd.DataFrame()

    combined = pd.concat(all_frames, ignore_index=True, sort=False)

    # Drop columns that shouldn't appear in the final combined output
    columns_to_drop = ["Source File", "Source Sheet", "Date"]
    combined = combined.drop(columns=[c for c in columns_to_drop if c in combined.columns])

    # Rename columns for the final output
    combined = combined.rename(columns={"Stock": "Symbol"})

    return combined, pd.DataFrame(file_summary)


# Each entry: (base_url, suffix, display_prefix, uses_symbol_in_formula)
# uses_symbol_in_formula = True  -> build a live =HYPERLINK(...) formula that
#                                    references that row's own Symbol cell
#                                    (behaves like a normal dragged-down formula)
# uses_symbol_in_formula = False -> static display only (no live formula)
HYPERLINK_SPECS = {
    "NSE Chart": (
        "https://www.nseindia.com/get-quotes/equity?symbol=",
        "",
        "n",
        True,
    ),  # Static display
    "Trading View": (
        "https://www.tradingview.com/symbols/",
        "",
        "Tre ",
        True,
    ),
    "History Data": (
        "https://www.equitypandit.com/historical-data/",
        "",
        "his ",
        True,
    ),  # Replace with actual base URL
    "Screener": (
        "https://www.screener.in/company/",
        "",
        "Scr ",
        True,
    ),  # Replace with actual base URL structure
    "Zerodha": (
        "https://zerodha.com/markets/stocks/NSE/",
        "",
        "Z ",
        True,
    ),  # Replace with actual base URL
    "Chartlink": (
        "https://chartink.com/stocks/",
        ".html",
        "CL ",
        True,
    ),  # Replace with actual base URL
    "Marketsmith": (
        "https://marketsmithindia.com/mstool/eval/",
        "/evaluation.jsp",
        "ms ",
        True,
    ),  # Replace with actual base URL
}


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Convert a DataFrame to an in-memory Excel file (single 'Final List' tab),
    with live HYPERLINK formulas added per row for each configured site."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Final List")

        worksheet = writer.sheets["Final List"]
        headers = list(df.columns)

        # Auto-fit original column widths for readability
        for i, col in enumerate(headers, start=1):
            max_len = max(
                df[col].astype(str).map(len).max() if len(df) else 0, len(str(col))
            )
            worksheet.column_dimensions[
                worksheet.cell(row=1, column=i).column_letter
            ].width = min(max_len + 2, 40)

        # Add hyperlink columns (NSE Chart, Trading View, History Data, Screener,
        # Zerodha, Chartlink, Marketsmith), each written as a real Excel formula
        # per row, referencing that row's own Symbol cell.
        if "Symbol" in headers and len(df) > 0:
            symbol_col_idx = headers.index("Symbol") + 1  # 1-based
            symbol_col_letter = worksheet.cell(row=1, column=symbol_col_idx).column_letter

            start_col = len(headers) + 1
            for offset, (label, (base_url, suffix, prefix, use_formula)) in enumerate(
                HYPERLINK_SPECS.items()
            ):
                col_idx = start_col + offset
                col_letter = worksheet.cell(row=1, column=col_idx).column_letter

                # Header
                header_cell = worksheet.cell(row=1, column=col_idx, value=label)
                header_cell.font = header_cell.font.copy(bold=True)

                for row_num in range(2, len(df) + 2):
                    sym_ref = f"{symbol_col_letter}{row_num}"
                    cell = worksheet.cell(row=row_num, column=col_idx)
                    if use_formula:
                        # e.g. =HYPERLINK("https://.../"&A2&".html","CL "&A2)
                        formula = (
                            f'=HYPERLINK("{base_url}"&{sym_ref}&"{suffix}",'
                            f'"{prefix}"&{sym_ref})'
                        )
                        cell.value = formula
                    else:
                        # Static display only (no live formula), e.g. NSE Chart
                        cell.value = prefix

                worksheet.column_dimensions[col_letter].width = max(len(label) + 2, 14)

    return output.getvalue()


if uploaded_files:
    combined_df, summary_df = load_and_combine(uploaded_files)

    if combined_df is not None:
        original_count = len(combined_df)

        if drop_duplicates:
            # Drop duplicates based on actual data columns only (ignore source tracking columns)
            data_cols = [
                c for c in combined_df.columns if c not in ("Source File", "Source Sheet")
            ]
            combined_df = combined_df.drop_duplicates(subset=data_cols, keep="first")

        st.success(
            f"Merged {len(uploaded_files)} file(s) → {len(summary_df)} sheet(s) → "
            f"{original_count} rows read, {len(combined_df)} rows in final combined tab."
        )

        with st.expander("📄 File / Sheet summary"):
            st.dataframe(summary_df, use_container_width=True)

        st.subheader("Combined Data Preview")
        st.dataframe(combined_df, use_container_width=True)

        excel_bytes = to_excel_bytes(combined_df)
        today_str = datetime.now().strftime("%Y-%m-%d")

        st.download_button(
            label="⬇️ Download Combined Excel File",
            data=excel_bytes,
            file_name=f"Breakout_List_Combined_{today_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("👆 Upload one or more Excel files to get started.")
