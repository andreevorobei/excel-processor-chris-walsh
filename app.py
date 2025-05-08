# app.py

import streamlit as st
import pandas as pd
from excel_processor_app_en import process_excel  # Import updated function

# Set page config to wide mode but with left-aligned content
st.set_page_config(page_title="Excel Processor", layout="wide")

# Apply custom CSS to left-align content and improve tab styling
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
    
    /* Improved Tab Styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0px;
        border-bottom: 1px solid #333;
        margin-bottom: 20px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 40px;
        white-space: pre-wrap;
        background-color: transparent;
        border-radius: 4px 4px 0 0;
        padding: 10px 16px;
        margin-right: 4px;
        font-size: 14px;
        font-weight: 500;
        color: #aaaaaa;
        border: none;
    }
    .stTabs [aria-selected="true"] {
        background-color: transparent;
        border-bottom: 2px solid #ff4b4b;
        color: white;
        font-weight: 600;
    }
    .stTabs [data-baseweb="tab"]:hover {
        background-color: rgba(255, 255, 255, 0.05);
        color: white;
    }
    
    /* Custom styles for main tabs (Results, Statistics, Log) */
    div[data-testid="stHorizontalBlock"] {
        gap: 0;
        border-bottom: none;
    }
    button[role="tab"] {
        background: transparent;
        color: #aaaaaa;
        border: none;
        border-radius: 0;
        padding: 0.75rem 1.5rem;
        font-size: 1rem;
        font-weight: 500;
        position: relative;
    }
    button[role="tab"][aria-selected="true"] {
        color: white;
        border-bottom: 2px solid #ff4b4b;
        font-weight: 600;
    }
    button[role="tab"]:hover {
        color: white;
        background-color: rgba(255, 255, 255, 0.05);
    }
    
    /* Custom styles for table headers */
    .dataframe th {
        font-weight: bold;
        background-color: #262730;
    }
    /* Style for dark mode tables */
    .dataframe {
        border: 1px solid #4e4e4e;
    }
    .dataframe td, .dataframe th {
        border: 1px solid #4e4e4e;
        padding: 8px;
    }
    
    /* Custom styling for tab navigation area at top */
    div[data-testid="block-container"] > div:nth-child(1) > div > div > div {
        background-color: #1e1e24;
        padding: 0;
        border-bottom: 1px solid #333;
    }
    
    /* Custom styling for success message */
    div[data-baseweb="notification"] {
        margin-bottom: 1rem;
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
                st.success(f"Found {len(result_df)} records with SCORE above average in their groups")
                
                # Create tabs for different types of data with custom styling
                tabs = st.tabs(["Results", "Statistics", "Detailed Log"])
                
                with tabs[0]:  # Results tab
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
                
                with tabs[1]:  # Statistics tab
                    # Get combined statistics
                    combined_ovacls_stats = st.session_state['combined_ovacls_stats']
                    combined_detailed_stats = st.session_state['combined_detailed_stats']
                    
                    # OVACLS Group Statistics
                    st.subheader("OVACLS Group Statistics")
                    
                    # Add index column for display
                    indexed_stats = combined_ovacls_stats.copy()
                    indexed_stats.insert(0, 'Index', range(len(indexed_stats)))
                    
                    # Reorder columns for better readability
                    indexed_stats = indexed_stats[['Index', 'OVACLS', 'Records_Before', 'Records_After']]
                    
                    st.dataframe(indexed_stats, use_container_width=True)
                    
                    # Detailed Group Statistics
                    st.subheader("Detailed Group Statistics")
                    
                    # Format and display detailed statistics
                    formatted_detailed_stats = combined_detailed_stats.copy()
                    
                    # Add index for display
                    formatted_detailed_stats.insert(0, 'Index', range(len(formatted_detailed_stats)))
                    
                    # Format numeric columns to 2 decimal places
                    for col in ['Average_SCORE', 'Min_SCORE', 'Max_SCORE']:
                        if col in formatted_detailed_stats.columns:
                            formatted_detailed_stats[col] = formatted_detailed_stats[col].apply(lambda x: f"{x:.2f}" if x > 0 else "0.00")
                    
                    st.dataframe(formatted_detailed_stats, use_container_width=True)
                
                with tabs[2]:  # Detailed Log tab
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