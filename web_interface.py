try:
    import streamlit as st
    import pandas as pd
    from datetime import datetime
    import os
    from transfer_system import TransferOptimizer
except ImportError as e:
    print(f"Import error: {e}")
    print("Please run: pip install -r requirements.txt")
    exit(1)

# Page configuration
st.set_page_config(
    page_title="Smart Transfer Optimization System",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS styles
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .kpi-card {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1f77b4;
        margin-bottom: 1rem;
    }
    .success-check {
        color: #28a745;
        font-weight: bold;
    }
    .warning-check {
        color: #ffc107;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

def main():
    st.markdown('<h1 class="main-header">📦 Smart Transfer Optimization System</h1>', unsafe_allow_html=True)
    
    # Sidebar - File upload and parameter settings
    with st.sidebar:
        st.header("📁 File Settings")
        
        uploaded_file = st.file_uploader(
            "Upload Excel File", 
            type=['xlsx'],
            help="Supports .xlsx format Excel files"
        )
        
        st.header("⚙️ Processing Options")
        process_btn = st.button("🚀 Start Processing", type="primary")
        
        if uploaded_file and process_btn:
            st.info("File is ready, click Start Processing to run transfer analysis")
    
    # Main content area
    if uploaded_file is not None:
        if process_btn:
            with st.spinner("Processing file, please wait..."):
                try:
                    # Save uploaded file
                    file_path = f"temp_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    with open(file_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    
                    # Initialize transfer optimizer
                    optimizer = TransferOptimizer()
                    
                    # Process file
                    output_file, suggestions = optimizer.process_file(file_path)
                    
                    # Display processing results
                    st.success("✅ File processing completed!")
                    
                    # Display transfer recommendations
                    if suggestions:
                        st.subheader("📋 Transfer Recommendations Details")
                        suggestions_df = pd.DataFrame(suggestions)
                        st.dataframe(suggestions_df)
                        
                        # Display statistical information
                        st.subheader("📊 Statistical Summary")
                        
                        col1, col2, col3, col4 = st.columns(4)
                        
                        with col1:
                            st.metric("Total Recommendations", len(suggestions))
                        
                        with col2:
                            total_qty = sum(t['Transfer Qty'] for t in suggestions)
                            st.metric("Total Transfer Qty", f"{total_qty:,.0f}")
                        
                        with col3:
                            nd_count = len([t for t in suggestions if t['Transfer Type'] == 'ND'])
                            st.metric("ND Type Transfers", nd_count)
                        
                        with col4:
                            emergency_count = len([t for t in suggestions if t['Receive Priority'] == 'Emergency'])
                            st.metric("Emergency Transfers", emergency_count)
                        
                        # Download buttons
                        st.subheader("💾 Export Results")
                        
                        col_dl1, col_dl2 = st.columns(2)
                        
                        with col_dl1:
                            # CSV Download
                            csv_data = suggestions_df.to_csv(index=False).encode('utf-8')
                            st.download_button(
                                label="📥 Download CSV",
                                data=csv_data,
                                file_name=f"transfer_suggestions_{datetime.now().strftime('%Y%m%d')}.csv",
                                mime="text/csv"
                                
                            )
                        
                        with col_dl2:
                            # Excel Download
                            if os.path.exists(output_file):
                                with open(output_file, "rb") as f:
                                    excel_data = f.read()
                                st.download_button(
                                    label="📥 Download Excel",
                                    data=excel_data,
                                    file_name=output_file,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    
                                )
                    
                    else:
                        st.info("ℹ️ No transfer suggestions needed for current data")
                    
                    # Clean up temporary files
                    try:
                        os.remove(file_path)
                    except:
                        pass
                        
                except Exception as e:
                    st.error(f"❌ Error processing file: {str(e)}")
        
        else:
            # Display file preview
            st.info("📄 File uploaded. Click 'Start Processing' to run transfer analysis")
            
            try:
                # Read file preview
                df = pd.read_excel(uploaded_file)
                
                st.subheader("File Preview")
                st.dataframe(df.head(10))
                
                st.subheader("File Information")
                col_info1, col_info2 = st.columns(2)
                
                with col_info1:
                    st.metric("Total Rows", df.shape[0])
                    st.metric("Total Columns", df.shape[1])
                
                with col_info2:
                    st.metric("Article Field", "Exists" if 'Article' in df.columns else "Missing")
                    st.metric("RP Type Field", "Exists" if 'RP Type' in df.columns else "Missing")
                
            except Exception as e:
                st.warning(f"⚠️ Error previewing file: {str(e)}")
    
    else:
        # Display welcome message and user guide
        st.info("👋 Welcome to Smart Transfer Optimization System")
        
        with st.expander("📖 User Guide", expanded=True):
            st.markdown("""
            ### System Features
            - **Smart Transfer Analysis**: Automatically generates transfer suggestions based on inventory, sales, and safety stock data
            - **Multi-dimensional Statistics**: Provides detailed statistical analysis and visual reports
            - **Quality Assurance**: Automatically performs quality checks to ensure data accuracy
            
            ### File Requirements
            **Required Columns**:
            - `Article`: Product code (12 digits)
            - `RP Type`: Type identifier (ND/RF)
            - `Site`: Store location
            - `OM`: Organizational unit
            - `SaSa Net Stock`: Current inventory
            - `Safety Stock`: Safety stock
            
            **Optional Columns**:
            - `Last Month Sold Qty`: Last month sales quantity
            - `MTD Sold Qty`: Month-to-date sales quantity
            - `Pending Received`: Pending received inventory
            
            ### Processing Workflow
            1. Upload Excel file (.xlsx format)
            2. Click 'Start Processing' button
            3. View transfer suggestions and statistical analysis
            4. Download result files (CSV/Excel format)
            """)
        
        # Display sample data structure
        with st.expander("📋 Sample Data Structure"):
            sample_data = {
                'Article': ['123456789012', '123456789012', '123456789013'],
                'RP Type': ['ND', 'RF', 'RF'],
                'Site': ['Warehouse A', 'Warehouse B', 'Warehouse C'],
                'OM': ['1001', '1001', '1002'],
                'SaSa Net Stock': [150, 80, 200],
                'Safety Stock': [60, 60, 80],
                'Last Month Sold Qty': [50, 60, 40],
                'MTD Sold Qty': [45, 55, 35],
                'Pending Received': [0, 10, 5]
            }
            st.dataframe(pd.DataFrame(sample_data))

if __name__ == "__main__":
    main()