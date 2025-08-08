import streamlit as st

from utils import load_excel_file, compare_excels, export_diff_to_excel

__version__ = "0.1.0-beta"

ALLOWED_EXTENSIONS = ["xlsx"]
MAX_FILE_SIZE = 100 * 1024 * 1024  # 100MB

st.set_page_config(page_title="Simple Excel Diff Tool v {__version__}", layout="wide")


def init_session_state():
    defaults = {
        "comparison_result": None,
        "prev_file1": None,
        "prev_file2": None
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def validate_file(file):
    if file is None:
        return False
    if file.size > MAX_FILE_SIZE:
        st.error(f"File size exceeds {MAX_FILE_SIZE // 1024 // 1024}MB limit")
        return False
    if not file.name.lower().endswith(tuple(ALLOWED_EXTENSIONS)):
        st.error(f"Invalid file type. Please upload {', '.join(ALLOWED_EXTENSIONS)} files")
        return False
    return True


def filter_dataframe(df, key_prefix):
    try:
        filter_col = st.selectbox("Select column to filter", df.columns, key=f'{key_prefix}FilterCol')
        filter_values = st.multiselect(f"Select values for '{filter_col}'", df[filter_col].unique(), key=f'{key_prefix}Filter')
        return df[df[filter_col].isin(filter_values)] if filter_values else df
    except Exception as e:
        st.error(f"Error loading file: {str(e)}")
        return None


def process_comparison(file1, file2):
    try:
        with st.expander("🔍 Preview of Uploaded Files", expanded=False):

            df1 = load_excel_file(file1)
            df2 = load_excel_file(file2)
            if df1 is None or df2 is None:
                st.error("Error loading one or both files")
                return

            st.write("**Old File:**")
            filtered_df1 = filter_dataframe(df1, "OldFile")
            if filtered_df1 is not None:
                st.dataframe(filtered_df1, use_container_width=True)

            st.write("**New File:**")
            filtered_df2 = filter_dataframe(df2, "NewFile")
            if filtered_df2 is not None:
                st.dataframe(filtered_df2, use_container_width=True)

            if filtered_df1 is None or filtered_df2 is None:
                return



        compare_clicked = st.button("🧙🏻‍♀️ Compare Files")
        if compare_clicked:
            st.session_state.comparison_result = None

        if st.session_state.comparison_result is not None or compare_clicked:
            with st.spinner("Comparing files..."):
                if st.session_state.comparison_result is None:
                    try:
                        result_df, changes = compare_excels(filtered_df1, filtered_df2)
                        st.session_state.comparison_result = result_df
                    except Exception as e:
                        st.error(f"Comparison failed: {str(e)}")
                        return
                result_df = st.session_state.comparison_result

            st.success("Comparison complete!")
            st.subheader("⚖️ Differences")
            st.write("**DIFF File:**")

            def highlight_special(val):
                return 'background-color: gold' if '→' in str(val) else ''

 

            filtered_result_df = filter_dataframe(result_df, "Result")
            if filtered_result_df is not None:
                styled_df = filtered_result_df.style.applymap(highlight_special)
                st.dataframe(styled_df, use_container_width=True)

            st.download_button(
                label="⬇️ Download Diff as Excel",
                data=export_diff_to_excel(result_df),
                file_name="excel_diff_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"An unexpected error occurred: {str(e)}")

def main():
    init_session_state()

    # UI Components
    st.title("📊 Simple Excel Diff Tool ")
    
    with st.expander("🚫 Limitations", expanded=False):
        st.markdown("""
- **File Format**: Only `.xlsx` files are supported (`.xls` format is not supported)
- **Headers**: First row considered as heder row
- **Column Handling**: 
  - Comparison is based on columns from the OLD file
  - New columns in the NEW file will be ignored
  - Excel's internal data filters should be disabled before upload
- **Memory Usage**: Large files may require increased memory allocation
- **Sheet Support**: Only first worksheet is compared

                """)

   

    footer_html = f"""
<style>
.footer {{
    position: fixed;
    left: 0;
    bottom: 0;
    width: 100%;
    background-color: rgba(240, 242, 246, 0.9);
    color: #666666;
    text-align: center;
    padding: 8px;
    font-size: 12px;
    border-top: 1px solid #cccccc;
    z-index: 999;
}}
.main .block-container {{
    padding-bottom: 5rem;
}}
.footer a {{
    color: #666666;
    text-decoration: none;
}}
.footer a:hover {{
    text-decoration: underline;
}}
</style>
<div class="footer">
    <p style="margin: 0;">Simple Excel Diff Tool v {__version__} by <a href="mailto:mitien@gmail.com">Mitien</a> | Built with Streamlit</p> 
</div>
"""
    st.markdown(footer_html, unsafe_allow_html=True)

    # File upload with validation
    col1, col2 = st.columns(2)
    with col1:
        file1 = st.file_uploader("📂 Upload OLD Excel file", type=ALLOWED_EXTENSIONS, key="old")
        if file1 and validate_file(file1):
            if st.session_state.get('prev_file1', None) != file1.name:
                st.session_state.comparison_result = None
                st.session_state.prev_file1 = file1.name

    with col2:
        file2 = st.file_uploader("📂 Upload NEW Excel file", type=ALLOWED_EXTENSIONS, key="new")
        if file2 and validate_file(file2):
            if st.session_state.get('prev_file2', None) != file2.name:
                st.session_state.comparison_result = None
                st.session_state.prev_file2 = file2.name

    if file1 and file2:
        try:
            process_comparison(file1, file2)
        except Exception as e:
            st.error(f"Error during comparison: {str(e)}")


if __name__ == "__main__":
    main()
