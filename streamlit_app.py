import streamlit as st
import pandas as pd
import os
import tempfile
from datetime import datetime
from invoice_verifier import ComprehensiveInvoiceVerifier

# Page configuration
st.set_page_config(
    page_title="Invoice Verification System",
    page_icon="✅",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        padding: 1rem 0;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #666;
        text-align: center;
        margin-bottom: 2rem;
    }
    .success-box {
        padding: 1rem;
        background-color: #d4edda;
        border: 2px solid #28a745;
        border-radius: 5px;
        margin: 1rem 0;
    }
    .error-box {
        padding: 1rem;
        background-color: #f8d7da;
        border: 2px solid #dc3545;
        border-radius: 5px;
        margin: 1rem 0;
    }
    .info-box {
        padding: 1rem;
        background-color: #d1ecf1;
        border: 2px solid #17a2b8;
        border-radius: 5px;
        margin: 1rem 0;
    }
    .field-check {
        display: inline-block;
        padding: 0.3rem 0.8rem;
        margin: 0.2rem;
        border-radius: 15px;
        font-weight: bold;
    }
    .field-pass {
        background-color: #d4edda;
        color: #155724;
    }
    .field-fail {
        background-color: #f8d7da;
        color: #721c24;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown('<div class="main-header">📋 Invoice Verification System</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Comprehensive 7-Field Verification: Quantity • Weight • Description • HS Code • Country • Year • Unit Value • Total Value</div>', unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.image("https://via.placeholder.com/300x100/1f77b4/ffffff?text=Invoice+Verifier", use_column_width=True)
    
    st.markdown("### 📖 How It Works")
    st.info("""
    **1. Upload Files**
    - Your Excel invoice (.xlsx or .xlsb)
    - FedEx PDF invoice
    
    **2. Automatic Verification**
    - Checks all 7 critical fields
    - Compares line by line
    - Detects discrepancies
    
    **3. Download Report**
    - Detailed verification report
    - Field-by-field comparison
    - Clear PASS/FAIL results
    """)
    
    st.markdown("### ✅ Fields Checked")
    fields = [
        "Quantity",
        "Net Weight (kg)",
        "Description",
        "HS Code (970531)",
        "Country of Origin",
        "Year",
        "Unit Value",
        "Total Value"
    ]
    for field in fields:
        st.markdown(f"✓ {field}")
    
    st.markdown("---")
    st.markdown("### 💡 Tips")
    st.success("""
    • Excel file should have standard format
    • PDF should be FedEx commercial invoice
    • Results appear instantly
    • Report can be downloaded
    """)

# Main content area
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📤 Upload Your Excel Invoice")
    excel_file = st.file_uploader(
        "Choose your Excel invoice file",
        type=['xlsx', 'xlsb'],
        help="Upload your invoice in Excel format (.xlsx or .xlsb)"
    )
    if excel_file:
        st.success(f"✓ File uploaded: {excel_file.name}")
        st.info(f"📊 File size: {len(excel_file.getvalue()) / 1024:.1f} KB")

with col2:
    st.markdown("### 📤 Upload FedEx PDF Invoice")
    pdf_file = st.file_uploader(
        "Choose FedEx PDF invoice file",
        type=['pdf'],
        help="Upload the FedEx commercial invoice PDF"
    )
    if pdf_file:
        st.success(f"✓ File uploaded: {pdf_file.name}")
        st.info(f"📊 File size: {len(pdf_file.getvalue()) / 1024:.1f} KB")

# Verification button
st.markdown("---")
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    verify_button = st.button(
        "🔍 Start Verification",
        type="primary",
        disabled=(excel_file is None or pdf_file is None),
        use_container_width=True
    )

# Process verification
if verify_button and excel_file and pdf_file:
    with st.spinner("🔄 Processing invoices... This may take a few seconds..."):
        try:
            # Create temporary files
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx' if excel_file.name.endswith('.xlsx') else '.xlsb') as tmp_excel:
                tmp_excel.write(excel_file.getvalue())
                excel_path = tmp_excel.name
            
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_pdf:
                tmp_pdf.write(pdf_file.getvalue())
                pdf_path = tmp_pdf.name
            
            # Convert .xlsb to .xlsx if needed
            if excel_file.name.endswith('.xlsb'):
                st.info("📦 Converting .xlsb to .xlsx format...")
                from pyxlsb import open_workbook
                
                xlsx_path = excel_path.replace('.xlsb', '.xlsx')
                with open_workbook(excel_path) as wb:
                    for sheet_name in wb.sheets:
                        with wb.get_sheet(sheet_name) as sheet:
                            data = []
                            for row in sheet.rows():
                                data.append([cell.v for cell in row])
                            
                            df = pd.DataFrame(data)
                            df.to_excel(xlsx_path, index=False, header=False)
                            break
                
                excel_path = xlsx_path
            
            # Run verification
            st.info("🔍 Verifying all 7 fields...")
            verifier = ComprehensiveInvoiceVerifier(excel_path, pdf_path)
            results = verifier.verify()
            
            # Generate report
            report_path = tempfile.mktemp(suffix='.txt')
            verifier.generate_report(report_path)
            
            # Display results
            st.markdown("---")
            st.markdown("## 📊 Verification Results")
            
            # Get status
            contact_ok = results['contact_name']['status'] == '✓ PASS'
            purpose_ok = results['purpose_of_shipment']['status'] == '✓ PASS'
            summary = results['line_items']['summary']
            items_ok = (summary['perfect_matches'] == summary['total_excel'] and
                       len(summary['mismatches']) == 0 and
                       len(summary['unmatched_excel']) == 0 and
                       len(summary['unmatched_pdf']) == 0)
            
            overall_pass = contact_ok and purpose_ok and items_ok
            
            # Overall status
            if overall_pass:
                st.markdown("""
                <div class="success-box">
                    <h2 style="color: #28a745; margin: 0;">✅ ALL VERIFICATION CHECKS PASSED!</h2>
                    <p style="margin: 0.5rem 0 0 0;">Your invoice matches the FedEx invoice perfectly on all 7 fields!</p>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown("""
                <div class="error-box">
                    <h2 style="color: #dc3545; margin: 0;">⚠️ DISCREPANCIES DETECTED</h2>
                    <p style="margin: 0.5rem 0 0 0;">Please review the detailed comparison below.</p>
                </div>
                """, unsafe_allow_html=True)
            
            # Detailed results in columns
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.markdown("### 📌 Header Verification")
                status_icon = "✅" if contact_ok else "❌"
                st.markdown(f"{status_icon} **Contact Name:** {results['contact_name']['status']}")
                
                status_icon = "✅" if purpose_ok else "❌"
                st.markdown(f"{status_icon} **Purpose:** {results['purpose_of_shipment']['status']}")
            
            with col2:
                st.markdown("### 📦 Line Items Summary")
                st.metric("Total Items (Excel)", summary['total_excel'])
                st.metric("Total Items (FedEx)", summary['total_pdf'])
                st.metric("Perfect Matches", summary['perfect_matches'], 
                         delta=None if summary['perfect_matches'] == summary['total_excel'] else f"{summary['total_excel'] - summary['perfect_matches']} issues")
            
            with col3:
                st.markdown("### 🔍 Discrepancies")
                st.metric("Partial Matches", len(summary['mismatches']))
                st.metric("Excel Only", len(summary['unmatched_excel']))
                st.metric("FedEx Only", len(summary['unmatched_pdf']))
            
            # Show discrepancies if any
            if not overall_pass:
                st.markdown("---")
                st.markdown("## 🔎 Detailed Discrepancies")
                
                if summary['mismatches']:
                    st.markdown("### ⚠️ Items with Field Mismatches")
                    for idx, mismatch in enumerate(summary['mismatches'], 1):
                        with st.expander(f"📋 Discrepancy #{idx} - Match Score: {mismatch['match_score']}", expanded=(idx <= 3)):
                            excel_item = mismatch['excel_item']
                            pdf_item = mismatch['pdf_item']
                            diffs = mismatch['differences']
                            
                            col1, col2 = st.columns(2)
                            
                            with col1:
                                st.markdown("**YOUR EXCEL:**")
                                st.write(f"Quantity: {excel_item.get('quantity')}")
                                st.write(f"Weight: {excel_item.get('net_weight')} kg")
                                st.write(f"HS Code: {excel_item.get('hs_code')}")
                                st.write(f"Country: {excel_item.get('country_code')}")
                                st.write(f"Year: {excel_item.get('year')}")
                                st.write(f"Unit Value: ${excel_item.get('unit_value')}")
                                st.write(f"Total: ${excel_item.get('total_value')}")
                            
                            with col2:
                                st.markdown("**FEDEX PDF:**")
                                st.write(f"Quantity: {pdf_item.get('quantity')}")
                                st.write(f"Weight: {pdf_item.get('net_weight')} kg")
                                st.write(f"HS Code: {pdf_item.get('hs_code')}")
                                st.write(f"Country: {pdf_item.get('country_code')}")
                                st.write(f"Year: {pdf_item.get('year')}")
                                st.write(f"Unit Value: ${pdf_item.get('unit_value')}")
                                st.write(f"Total: ${pdf_item.get('total_value')}")
                            
                            st.markdown("**Field Status:**")
                            fields_status = {
                                'Quantity': diffs['quantity'],
                                'Weight': diffs['net_weight'],
                                'HS Code': diffs['hs_code'],
                                'Country': diffs['country_code'],
                                'Year': diffs['year'],
                                'Unit Value': diffs['unit_value'],
                                'Total': diffs['total_value']
                            }
                            
                            status_html = ""
                            for field, status in fields_status.items():
                                css_class = "field-pass" if status else "field-fail"
                                icon = "✓" if status else "✗"
                                status_html += f'<span class="field-check {css_class}">{icon} {field}</span> '
                            
                            st.markdown(status_html, unsafe_allow_html=True)
                
                if summary['unmatched_excel']:
                    st.markdown("### 📤 Items Only in Your Excel")
                    for idx, item in enumerate(summary['unmatched_excel'], 1):
                        with st.expander(f"Excel Item #{idx}"):
                            st.write(f"**Quantity:** {item.get('quantity')}")
                            st.write(f"**Weight:** {item.get('net_weight')} kg")
                            st.write(f"**Description:** {item.get('description')[:100]}...")
                            st.write(f"**Country:** {item.get('country_code')} | **Year:** {item.get('year')}")
                
                if summary['unmatched_pdf']:
                    st.markdown("### 📥 Items Only in FedEx PDF")
                    for idx, item in enumerate(summary['unmatched_pdf'], 1):
                        with st.expander(f"FedEx Item #{idx}"):
                            st.write(f"**Quantity:** {item.get('quantity')}")
                            st.write(f"**Weight:** {item.get('net_weight')} kg")
                            st.write(f"**Description:** {item.get('description')[:100]}...")
                            st.write(f"**Country:** {item.get('country_code')} | **Year:** {item.get('year')}")
            
            # Download report button
            st.markdown("---")
            st.markdown("## 📥 Download Full Report")
            
            with open(report_path, 'r', encoding='utf-8') as f:
                report_content = f.read()
            
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                st.download_button(
                    label="📄 Download Verification Report",
                    data=report_content,
                    file_name=f"verification_report_{timestamp}.txt",
                    mime="text/plain",
                    type="primary",
                    use_container_width=True
                )
            
            # Cleanup
            try:
                os.unlink(excel_path)
                os.unlink(pdf_path)
                os.unlink(report_path)
            except:
                pass
                
        except Exception as e:
            st.error(f"❌ Error during verification: {str(e)}")
            st.exception(e)

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666; padding: 2rem 0;">
    <p>Invoice Verification System v1.0 | Checks All 7 Critical Fields</p>
    <p>Quantity • Net Weight • Description • HS Code • Country • Year • Unit Value • Total Value</p>
</div>
""", unsafe_allow_html=True)

# Instructions if no files uploaded
if excel_file is None or pdf_file is None:
    st.markdown("---")
    st.info("""
    ### 👆 Get Started
    
    1. **Upload your Excel invoice** (left box)
    2. **Upload FedEx PDF invoice** (right box)
    3. **Click "Start Verification"** button
    4. **Review results** and download detailed report
    
    **The system will automatically:**
    - Extract all 7 fields from both invoices
    - Compare every field line by line
    - Show exact discrepancies (if any)
    - Generate a downloadable report
    """)
