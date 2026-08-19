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


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Convert a DataFrame to an in-memory Excel file (single combined tab)."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Final List")

        # Auto-fit column widths for readability
        worksheet = writer.sheets["Final List"]
        for i, col in enumerate(df.columns, start=1):
            max_len = max(
                df[col].astype(str).map(len).max() if len(df) else 0, len(str(col))
            )
            worksheet.column_dimensions[
                worksheet.cell(row=1, column=i).column_letter
            ].width = min(max_len + 2, 40)
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
