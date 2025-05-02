# app.py

import streamlit as st
import pandas as pd
from excel_processor_app_en import process_excel  # Import updated function

# Set page config to wide mode but with left-aligned content
st.set_page_config(page_title="Excel Processor", layout="wide")

# Apply custom CSS to left-align content
st.markdown("""
<style>
    .block-container {
        padding-left: 2rem;
        padding-right: 2rem;
        max-width: 100%;
    }
    h1, h2, h3, p, div.stMarkdown {
        text-align: left !important;
    }
    div.stButton > button {
        float: left;
        margin-left: 0;
    }
    section[data-testid="stSidebar"] {
        width: 18rem !important;
    }
</style>
""", unsafe_allow_html=True)

st.title("🏗 Excel Property Filter")

st.write("""
### Instructions:
1. Upload an Excel file with property data
2. Click the "Process Data" button
3. The system will automatically:
   - Filter records by OVACLS (keeping only 201-212, 234, 295)
   - Calculate SCORE for all records
   - Group data by OVACLS and NEIGHBORHOOD
   - Keep only records with SCORE above average in their group
   - Sort results by OVACLS and NEIGHBORHOOD
4. You can download the filtered records in Excel format
""")

# Upload area with full width
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

# Process button in full width
if uploaded_file:
    if st.button("Process Data", use_container_width=True):
        try:
            with st.spinner('Processing data...'):
                result_df = process_excel(uploaded_file)
            
            if len(result_df) > 0:
                # Group statistics
                groups_stats = result_df.groupby('OVACLS').agg(
                    Records=('SCORE', 'count')
                ).reset_index()
                
                st.success(f"Found {len(result_df)} records with SCORE above average in their groups")
                
                # Create tabs for different types of data
                tab1, tab2, tab3 = st.tabs(["Results", "Statistics", "Detailed Log"])
                
                with tab1:
                    # Show results
                    st.subheader("Filtered Records")
                    st.dataframe(result_df, use_container_width=True)
                    
                    # Download button
                    @st.cache_data
                    def convert_df_to_excel(df):
                        from io import BytesIO
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df.to_excel(writer, index=False)
                        return output.getvalue()

                    excel_bytes = convert_df_to_excel(result_df)

                    st.download_button(
                        label="📥 Download Results as Excel",
                        data=excel_bytes,
                        file_name="filtered_results.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                with tab2:
                    # Show group statistics
                    st.subheader("OVACLS Group Statistics")
                    st.dataframe(groups_stats, use_container_width=True)
                    
                    # Show more detailed statistics
                    st.subheader("Detailed Group Statistics")
                    detailed_stats = result_df.groupby(['OVACLS', 'NEIGHBORHOOD']).agg(
                        Count=('SCORE', 'count'),
                        Average_SCORE=('SCORE', 'mean'),
                        Min_SCORE=('SCORE', 'min'),
                        Max_SCORE=('SCORE', 'max')
                    ).reset_index()
                    st.dataframe(detailed_stats, use_container_width=True)
                
                with tab3:
                    # Show processing log
                    st.subheader("Detailed Processing Log")
                    if 'process_log' in st.session_state:
                        st.text_area("Processing log:", st.session_state['process_log'], height=600)
                    else:
                        st.info("Processing log not available")
            else:
                st.warning("No records found matching the filtering criteria")
                
                # Show log anyway to understand why
                if 'process_log' in st.session_state:
                    st.subheader("Detailed Processing Log")
                    st.text_area("Processing log:", st.session_state['process_log'], height=400)
                
        except Exception as e:
            st.error(f"Error processing file: {e}")
            
            # Show log in case of error for diagnostics
            if 'process_log' in st.session_state:
                st.subheader("Processing Log (contains errors)")
                st.text_area("Processing log:", st.session_state['process_log'], height=400)

# If file is uploaded, show original data at the bottom
if uploaded_file and 'df' not in st.session_state:
    try:
        # Just read data for display, without processing
        df = pd.read_excel(uploaded_file)
        st.session_state['df'] = df
    except Exception as e:
        st.error(f"Error reading file: {e}")

if 'df' in st.session_state:
    with st.expander("Original Data"):
        st.dataframe(st.session_state['df'].head(1000), use_container_width=True)  # Limit for performance