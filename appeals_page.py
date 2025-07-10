import streamlit as st
import pandas as pd
import io
from datetime import datetime
import base64
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
from dotenv import load_dotenv

load_dotenv('secrets.env')

def compile_appeals_files(uploaded_files):
    """
    Compile multiple appeals files into a single Excel file.
    Extracts PAYMENT SUMMARY sheet from each file and combines them.
    
    Args:
        uploaded_files: List of uploaded Excel files
        
    Returns:
        bytes: Compiled Excel file as bytes
    """
    compiled_data = []
    file_summary = []
    
    for uploaded_file in uploaded_files:
        try:
            # Read the Excel file
            excel_file = pd.ExcelFile(uploaded_file)
            
            # Check if PAYMENT SUMMARY sheet exists
            if 'PAYMENT SUMMARY' in excel_file.sheet_names:
                # Read the PAYMENT SUMMARY sheet with headers in row 2 (index 1)
                df = pd.read_excel(uploaded_file, sheet_name='PAYMENT SUMMARY', header=1)
                
                # Remove any rows that appear to be totals or empty rows
                if not df.empty:
                    # First, remove completely empty rows
                    df = df.dropna(how='all')
                    
                    # Find actual data rows by looking for numeric values in the first column (S/N)
                    # or meaningful data patterns
                    data_rows = []
                    for i in range(len(df)):
                        row = df.iloc[i]
                        first_col = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
                        
                        # Skip rows that are clearly totals or summaries
                        if any(word in first_col.upper() for word in ['TOTAL', 'SUM', 'GRAND', 'SUBTOTAL', 'SUMMARY']):
                            continue
                            
                        # Skip completely empty rows
                        if first_col == "" and row.isna().all():
                            continue
                            
                        # Keep rows that have meaningful data (numeric S/N or other data)
                        try:
                            # If first column is numeric (S/N), it's likely a data row
                            float(first_col)
                            data_rows.append(i)
                        except:
                            # If not numeric but has other meaningful data, keep it
                            non_empty_cols = sum(1 for val in row if pd.notna(val) and str(val).strip() != "")
                            if non_empty_cols >= 3:  # At least 3 non-empty columns suggests real data
                                data_rows.append(i)
                    
                    # Keep only the identified data rows
                    if data_rows:
                        df_clean = df.iloc[data_rows].copy()
                    else:
                        df_clean = df.copy()  # Fallback to keep all data if filtering fails
                    
                    # Create the standardized structure with your specified columns
                    template_columns = [
                        'S_N', 'CLAIM_TYPE', 'BATCH_NUMBER', 'HOSPITAL', 'NUMBER_OF_CLAIMS',
                        'ENCOUNTER_MONTH', 'DATE_OF_RECEIPT', 'APPROVED_PA_VALUE_N',
                        'AMOUNT_RECOMMENDED_FOR_PAYMENT_N', 'VARIANCE', 'VARIANCE1',
                        'NARRATION', 'Source_File', 'PROVIDER_CODE', 'Paiddate',
                        'SCH_NO', 'APPEAL_NO', 'SCH_NUM'
                    ]
                    
                    # Columns that should be left blank (for later manual completion)
                    blank_columns = ['PROVIDER_CODE', 'Paiddate', 'SCH_NO', 'APPEAL_NO', 'SCH_NUM']
                    
                    # Create standardized dataframe with same number of rows as original
                    data_dict = {}
                    for col in template_columns:
                        if col in blank_columns:
                            data_dict[col] = [''] * len(df_clean)  # Leave these blank
                        else:
                            data_dict[col] = [''] * len(df_clean)  # Initialize, will be filled
                    
                    standardized_df = pd.DataFrame(data_dict)
                    
                    # Comprehensive mapping from original columns to standardized structure
                    column_mapping = {
                        'S/N': 'S_N',
                        'AMOUNT RECOMMENDED FOR PAYMENT (N)': 'AMOUNT_RECOMMENDED_FOR_PAYMENT_N',
                        'ENROLLEE NAME': 'HOSPITAL',
                        'PROVIDER NAME': 'HOSPITAL',
                        'HOSPITAL NAME': 'HOSPITAL',
                        'CLAIM TYPE': 'CLAIM_TYPE',
                        'BATCH NUMBER': 'BATCH_NUMBER',
                        'BATCH NO': 'BATCH_NUMBER',
                        'NUMBER OF CLAIMS': 'NUMBER_OF_CLAIMS',
                        'NO OF CLAIMS': 'NUMBER_OF_CLAIMS',
                        'ENCOUNTER MONTH': 'ENCOUNTER_MONTH',
                        'DATE OF RECEIPT': 'DATE_OF_RECEIPT',
                        'APPROVED PA VALUE (N)': 'APPROVED_PA_VALUE_N',
                        'VARIANCE': 'VARIANCE',
                        'VARIANCE1': 'VARIANCE1',
                        'NARRATION': 'NARRATION',
                        'NARRATIVE': 'NARRATION'
                    }
                    
                    # Copy data from original columns where they exist (except blank columns)
                    for original_col, new_col in column_mapping.items():
                        if original_col in df_clean.columns and new_col not in blank_columns:
                            # Handle numeric columns properly
                            if new_col in ['S_N']:
                                standardized_df[new_col] = df_clean[original_col].apply(
                                    lambda x: str(int(float(x))) if pd.notna(x) and str(x).replace('.', '').isdigit() else str(x) if pd.notna(x) else ''
                                )
                            else:
                                standardized_df[new_col] = df_clean[original_col].astype(str)
                    
                    # Add source file information (not blank)
                    standardized_df['Source_File'] = uploaded_file.name
                    
                    compiled_data.append(standardized_df)
                    file_summary.append({
                        'File': uploaded_file.name,
                        'Rows': len(df_clean),
                        'Status': 'Success'
                    })
            else:
                file_summary.append({
                    'File': uploaded_file.name,
                    'Rows': 0,
                    'Status': 'No PAYMENT SUMMARY sheet found'
                })
                
        except Exception as e:
            file_summary.append({
                'File': uploaded_file.name,
                'Rows': 0,
                'Status': f'Error: {str(e)}'
            })
    
    return compiled_data, file_summary

def create_compiled_excel(compiled_data):
    """
    Create an Excel file from compiled data.
    
    Args:
        compiled_data: List of DataFrames to combine
        
    Returns:
        bytes: Excel file as bytes
    """
    if not compiled_data:
        return None
    
    # Combine all dataframes
    combined_df = pd.concat(compiled_data, ignore_index=True)
    
    # Create Excel file
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        combined_df.to_excel(writer, sheet_name='Compiled Appeals', index=False)
    
    output.seek(0)
    return output.getvalue()

def get_download_link(excel_bytes, filename="compiled_appeals.xlsx"):
    """
    Generate a download link for the Excel file.
    
    Args:
        excel_bytes (bytes): Excel file as bytes
        filename (str): Name of the file to download
        
    Returns:
        str: HTML link for downloading the Excel file
    """
    b64 = base64.b64encode(excel_bytes).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">Download Compiled Appeals Excel File</a>'
    return href

def extract_schedule_from_filename(filename):
    """
    Extract schedule number from filename.
    Example: "Payment Schedule 9497_APPEAL (BUPA).xlsx" -> "9497"
    """
    import re
    # Look for pattern like "Schedule XXXX" or "SCH XXXX" 
    pattern = r'(?:Schedule|SCH)\s*(\d+)'
    match = re.search(pattern, filename, re.IGNORECASE)
    return match.group(1) if match else None

def compare_with_finance(compiled_data, finance_file):
    """
    Compare compiled appeals data with finance file.
    
    Args:
        compiled_data: List of DataFrames from appeals
        finance_file: Uploaded finance file
        
    Returns:
        DataFrame: Comparison results
    """
    if not compiled_data or not finance_file:
        return None
    
    try:
        # Read finance file - look for 'CLAIMS RECEIVED WEEKLY REPORT' sheet
        finance_xls = pd.ExcelFile(finance_file)
        finance_sheet = None
        
        # Find the correct sheet name
        for sheet in finance_xls.sheet_names:
            if 'CLAIMS RECEIVED' in sheet.upper() or 'WEEKLY REPORT' in sheet.upper():
                finance_sheet = sheet
                break
        
        if not finance_sheet:
            return None
            
        finance_df = pd.read_excel(finance_file, sheet_name=finance_sheet)
        
        # Compile appeals data
        appeals_combined = pd.concat(compiled_data, ignore_index=True)
        
        # Group appeals by schedule number and sum amounts
        appeals_summary = []
        
        for _, row in appeals_combined.iterrows():
            filename = row.get('Source_File', '')
            schedule_num = extract_schedule_from_filename(filename)
            amount_col = 'AMOUNT_RECOMMENDED_FOR_PAYMENT_N'
            
            if schedule_num and amount_col in row:
                try:
                    amount = float(str(row[amount_col]).replace(',', ''))
                    appeals_summary.append({
                        'Schedule_Number': schedule_num,
                        'Appeals_Amount': amount,
                        'Source_File': filename
                    })
                except:
                    continue
        
        appeals_summary_df = pd.DataFrame(appeals_summary)
        if appeals_summary_df.empty:
            return None
            
        # Group by schedule number and sum
        appeals_grouped = appeals_summary_df.groupby('Schedule_Number').agg({
            'Appeals_Amount': 'sum',
            'Source_File': lambda x: ', '.join(x.unique())
        }).reset_index()
        
        # Find matching data in finance file
        comparison_results = []
        
        for _, appeal_row in appeals_grouped.iterrows():
            schedule_num = appeal_row['Schedule_Number']
            appeals_amount = appeal_row['Appeals_Amount']
            
            # Filter finance data for this schedule
            finance_filtered = finance_df[
                finance_df['Claim Batch No/Sch No'].astype(str).str.contains(str(schedule_num), na=False)
            ]
            
            if not finance_filtered.empty:
                # Sum the Claims_Advised_Amount for this schedule
                finance_amount = finance_filtered['Claims_Advised_Amount'].sum()
                variance = appeals_amount - finance_amount
                
                comparison_results.append({
                    'Schedule_Number': schedule_num,
                    'Appeals_Amount': appeals_amount,
                    'Finance_Amount': finance_amount,
                    'Variance': variance,
                    'Source_Files': appeal_row['Source_File']
                })
            else:
                comparison_results.append({
                    'Schedule_Number': schedule_num,
                    'Appeals_Amount': appeals_amount,
                    'Finance_Amount': 0,
                    'Variance': appeals_amount,
                    'Source_Files': appeal_row['Source_File']
                })
        
        comparison_df = pd.DataFrame(comparison_results)
        
        return comparison_df
        
    except Exception as e:
        st.error(f"Error comparing with finance data: {str(e)}")
        return None

def send_appeals_notification_email(missing_schedules, amount_mismatches):
    """
    Send email notification for appeals comparison issues.
    
    Args:
        missing_schedules (DataFrame): Schedules in appeals but not in finance
        amount_mismatches (DataFrame): Schedules with amount differences
        
    Returns:
        bool: True if email sent successfully, False otherwise
    """
    sender_email = os.getenv("OFFICE_SENDER_EMAIL")
    recipient_email = "ifeoluwa.adeniyi@avonhealthcare.com"
    cc_email = ["ifeoluwa.adeniyi@avonhealthcare.com",
                "adedamola.ayeni@avonhealthcare.com",
                "adebola.adesoyin@avonhealthcare.com",
                "financedepartment@avonhealthcare.com",
                "claims_officers@avonhealthcare.com",
                "bi_dataanalytics@avonhealthcare.com"
                ]
    password = os.getenv("OUTLOOK_APP_PASSWORD")
    
    if not sender_email or not password:
        st.error("Gmail credentials not configured. Please check OFFICE_SENDER_EMAIL and OUTLOOK_APP_PASSWORD secrets.")
        return False
    
    # Create message
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = recipient_email
    message['Cc'] = ", ".join(cc_email)
    message["Subject"] = "Appeals Finance Comparison Alert - Discrepancies Found"
    
    # Create email body
    body = f"""
    <html>
    <body>
        <h2>Appeals Finance Comparison Alert</h2>
        <p>This is an automated notification from the Appeals Compilation system.</p>
        <p><strong>Comparison completed at:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        
        <h3>Summary:</h3>
        <ul>
            <li><strong>Schedules missing in Finance:</strong> {len(missing_schedules)}</li>
            <li><strong>Amount mismatches:</strong> {len(amount_mismatches)}</li>
        </ul>
    """
    
    # Add missing schedules details
    if not missing_schedules.empty:
        body += """
        <h3>🚨 Schedules in Appeals but NOT in Finance:</h3>
        <table border="1" style="border-collapse: collapse; width: 100%;">
            <tr style="background-color: #f2f2f2;">
                <th>Schedule Number</th>
                <th>Appeals Amount</th>
                <th>Source Files</th>
            </tr>
        """
        for _, row in missing_schedules.iterrows():
            body += f"""
            <tr>
                <td>{row['Schedule_Number']}</td>
                <td>{row['Appeals_Amount']:,.2f}</td>
                <td>{row['Source_Files']}</td>
            </tr>
            """
        body += "</table><br>"
        
        total_missing_amount = missing_schedules['Appeals_Amount'].sum()
        body += f"<p><strong>Total amount missing in Finance:</strong> {total_missing_amount:,.2f}</p>"
    
    # Add amount mismatches details
    if not amount_mismatches.empty:
        body += """
        <h3>⚠️ Amount Mismatches (Both in Appeals and Finance):</h3>
        <table border="1" style="border-collapse: collapse; width: 100%;">
            <tr style="background-color: #f2f2f2;">
                <th>Schedule Number</th>
                <th>Appeals Amount</th>
                <th>Finance Amount</th>
                <th>Variance</th>
                <th>Source Files</th>
            </tr>
        """
        for _, row in amount_mismatches.iterrows():
            variance_color = "red" if row['Variance'] > 0 else "blue"
            body += f"""
            <tr>
                <td>{row['Schedule_Number']}</td>
                <td>{row['Appeals_Amount']:,.2f}</td>
                <td>{row['Finance_Amount']:,.2f}</td>
                <td style="color: {variance_color};">{row['Variance']:,.2f}</td>
                <td>{row['Source_Files']}</td>
            </tr>
            """
        body += "</table><br>"
        
        total_variance = amount_mismatches['Variance'].sum()
        body += f"<p><strong>Total variance:</strong> {total_variance:,.2f}</p>"
    
    body += """
        <p><strong>Action Required:</strong></p>
        <ul>
            <li>Review the missing schedules in the finance system</li>
            <li>Investigate amount discrepancies for variance resolution</li>
            <li>Check the Appeals Compilation system for detailed reports</li>
        </ul>
        
        <p>This is an automated message from the Appeals Compilation system.</p>
        <p><em>Please do not reply to this email.</em></p>
    </body>
    </html>
    """
    
    message.attach(MIMEText(body, "html"))
    
    try:
        # Create SMTP session using Gmail
        server = smtplib.SMTP("smtp.office365.com", 587)
        server.starttls()
        server.login(sender_email, password)
        allrecipients = [recipient_email] + cc_email

        # Send email
        server.sendmail(sender_email, allrecipients, message.as_string())
        server.quit()
        st.success("Email notification sent successfully for discrepancies found!")
        return True
    except Exception as e:
        st.error(f"Could not send email notification: {str(e)}")
        return False

def show_appeals_page():
    """
    Display the Appeals page interface.
    """
    st.header("Appeals Compilation & Finance Comparison")
    st.markdown("""
    Upload multiple appeals Excel files to compile their PAYMENT SUMMARY sheets and compare with finance data.
    The system will extract schedule numbers from filenames and compare amounts with the finance report.
    """)
    
    # Create two columns for uploads
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Appeals Files")
        # File uploader for multiple appeals files
        uploaded_files = st.file_uploader(
            "Upload Appeals Excel Files",
            type=['xlsx', 'xls'],
            accept_multiple_files=True,
            help="Select multiple Excel files containing appeals data with PAYMENT SUMMARY sheets",
            key="appeals_uploader"
        )
    
    with col2:
        st.subheader("Finance File")
        # File uploader for finance file
        finance_file = st.file_uploader(
            "Upload Finance Excel File",
            type=['xlsx', 'xls'],
            help="Upload the finance file containing 'CLAIMS RECEIVED WEEKLY REPORT' sheet",
            key="finance_uploader"
        )
    
    if uploaded_files:
        st.subheader(f"Appeals Files Selected: {len(uploaded_files)}")
        
        # Show list of uploaded files with extracted schedule numbers
        for i, file in enumerate(uploaded_files, 1):
            schedule_num = extract_schedule_from_filename(file.name)
            if schedule_num:
                st.write(f"{i}. {file.name} → Schedule {schedule_num}")
            else:
                st.write(f"{i}. {file.name} → ⚠️ Could not extract schedule number")
        
        # Processing options
        process_option = st.radio(
            "Choose processing option:",
            ["Compile Appeals Only", "Compile Appeals + Compare with Finance"],
            help="Select whether to just compile appeals or also compare with finance data"
        )
        
        if st.button("Process Appeals Files", type="primary"):
            with st.spinner("Processing appeals files..."):
                # Compile the files
                compiled_data, file_summary = compile_appeals_files(uploaded_files)
                
                # Show processing summary
                st.subheader("Processing Summary")
                summary_df = pd.DataFrame(file_summary)
                st.dataframe(summary_df, use_container_width=True)
                
                # Create and offer download if we have data
                if compiled_data:
                    excel_bytes = create_compiled_excel(compiled_data)
                    
                    if excel_bytes:
                        # Show compilation statistics
                        total_rows = sum(len(df) for df in compiled_data)
                        successful_files = len([s for s in file_summary if s['Status'] == 'Success'])
                        
                        st.success(f"Successfully compiled {successful_files} files with {total_rows} total rows")
                        
                        # Store compiled data in session state for later use
                        st.session_state['compiled_data'] = compiled_data
                        
                        # Generate download link for compiled appeals
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        filename = f"compiled_appeals_{timestamp}.xlsx"
                        
                        download_link = get_download_link(excel_bytes, filename)
                        st.markdown(f"**Download Compiled Appeals:** {download_link}", unsafe_allow_html=True)
                        
                        # If finance comparison is requested and finance file is uploaded
                        if process_option == "Compile Appeals + Compare with Finance" and finance_file:
                            st.subheader("Finance Comparison")
                            with st.spinner("Comparing with finance data..."):
                                comparison_df = compare_with_finance(compiled_data, finance_file)
                                
                                if comparison_df is not None and not comparison_df.empty:
                                    st.success("Finance comparison completed successfully!")
                                    
                                    # Store comparison results in session state
                                    st.session_state['comparison_df'] = comparison_df
                                    
                                    # Display comparison results
                                    st.dataframe(comparison_df, use_container_width=True)
                                    
                                    # Show summary statistics
                                    total_variance = comparison_df['Variance'].sum()
                                    matches = len(comparison_df[comparison_df['Variance'] == 0])
                                    mismatches = len(comparison_df[comparison_df['Variance'] != 0])
                                    missing_in_finance = len(comparison_df[comparison_df['Finance_Amount'] == 0])
                                    
                                    col1, col2, col3, col4 = st.columns(4)
                                    col1.metric("Total Variance", f"{total_variance:,.2f}")
                                    col2.metric("Exact Matches", matches)
                                    col3.metric("Amount Variances", mismatches - missing_in_finance)
                                    col4.metric("Missing in Finance", missing_in_finance)
                                    
                                    # Highlight critical issues
                                    if missing_in_finance > 0:
                                        st.error(f"🚨 {missing_in_finance} schedule(s) found in appeals but missing in finance!")
                                    
                                    if mismatches - missing_in_finance > 0:
                                        st.warning(f"⚠️ {mismatches - missing_in_finance} schedule(s) have amount discrepancies!")
                                    
                                    if matches == len(comparison_df):
                                        st.success("✅ All schedules match perfectly between appeals and finance!")
                                    
                                    # **NEW: Email notification section**
                                    if missing_in_finance > 0 or (mismatches - missing_in_finance) > 0:
                                        st.subheader("📧 Email Notification")
                                        st.info("Discrepancies found! You can send an email notification to the relevant teams.")
                                        
                                        # Show what will be included in the email
                                        with st.expander("📋 Email Preview"):
                                            st.write("**Email will include:**")
                                            if missing_in_finance > 0:
                                                st.write(f"- {missing_in_finance} schedules missing in finance")
                                            if (mismatches - missing_in_finance) > 0:
                                                st.write(f"- {mismatches - missing_in_finance} amount discrepancies")
                                            st.write("- Detailed tables with schedule numbers and amounts")
                                            st.write("- Action items for the finance and claims teams")
                                        
                                        # Send Email Button
                                        if st.button("📧 Send Email Notification", 
                                                   type="secondary",
                                                   help="Send email notification about the discrepancies found"):
                                            with st.spinner("Sending email notification..."):
                                                missing_schedules = comparison_df[comparison_df['Finance_Amount'] == 0]
                                                amount_mismatches = comparison_df[(comparison_df['Finance_Amount'] != 0) & (comparison_df['Variance'] != 0)]
                                                
                                                email_sent = send_appeals_notification_email(missing_schedules, amount_mismatches)
                                                
                                                if email_sent:
                                                    st.balloons()
                                                    st.success("✅ Email notification sent successfully!")
                                                else:
                                                    st.error("❌ Failed to send email notification. Please check your email configuration.")
                                    else:
                                        st.success("✅ No discrepancies found - no email notification needed!")
                                    
                                    # Create comparison Excel file
                                    comparison_output = io.BytesIO()
                                    with pd.ExcelWriter(comparison_output, engine='openpyxl') as writer:
                                        comparison_df.to_excel(writer, sheet_name='Finance Comparison', index=False)
                                        pd.concat(compiled_data, ignore_index=True).to_excel(writer, sheet_name='Compiled Appeals', index=False)
                                    
                                    comparison_output.seek(0)
                                    comparison_bytes = comparison_output.getvalue()
                                    
                                    comparison_filename = f"appeals_finance_comparison_{timestamp}.xlsx"
                                    comparison_link = get_download_link(comparison_bytes, comparison_filename)
                                    st.markdown(f"**Download Comparison Report:** {comparison_link}", unsafe_allow_html=True)
                                    
                                else:
                                    st.warning("Could not perform finance comparison. Please check that the finance file contains 'CLAIMS RECEIVED WEEKLY REPORT' sheet.")
                        
                        elif process_option == "Compile Appeals + Compare with Finance" and not finance_file:
                            st.warning("Please upload a finance file to perform comparison.")
                        
                        # Show preview of compiled data
                        if st.checkbox("Show Preview of Compiled Data"):
                            combined_preview = pd.concat(compiled_data, ignore_index=True)
                            st.subheader("Data Preview")
                            st.dataframe(combined_preview.head(20), use_container_width=True)
                    else:
                        st.error("Failed to create compiled Excel file")
                else:
                    st.error("No valid data found in the uploaded files. Please check that your files contain PAYMENT SUMMARY sheets.")
    
    # **NEW: Manual email section for previously processed data**
    if 'comparison_df' in st.session_state:
        st.markdown("---")
        st.subheader("📧 Manual Email Notification")
        st.info("You have previously processed comparison data. You can send email notifications for discrepancies.")
        
        comparison_df = st.session_state['comparison_df']
        missing_in_finance = len(comparison_df[comparison_df['Finance_Amount'] == 0])
        amount_mismatches = len(comparison_df[(comparison_df['Finance_Amount'] != 0) & (comparison_df['Variance'] != 0)])
        
        if missing_in_finance > 0 or amount_mismatches > 0:
            col1, col2 = st.columns([2, 1])
            
            with col1:
                st.write(f"**Current Issues:**")
                if missing_in_finance > 0:
                    st.write(f"- {missing_in_finance} schedules missing in finance")
                if amount_mismatches > 0:
                    st.write(f"- {amount_mismatches} amount discrepancies")
            
            with col2:
                if st.button("📧 Send Email for Current Issues", 
                           type="secondary",
                           help="Send email notification for the current comparison results"):
                    with st.spinner("Sending email notification..."):
                        missing_schedules = comparison_df[comparison_df['Finance_Amount'] == 0]
                        amount_mismatches_df = comparison_df[(comparison_df['Finance_Amount'] != 0) & (comparison_df['Variance'] != 0)]
                        
                        email_sent = send_appeals_notification_email(missing_schedules, amount_mismatches_df)
                        
                        if email_sent:
                            st.balloons()
                            st.success("✅ Email notification sent successfully!")
                        else:
                            st.error("❌ Failed to send email notification. Please check your email configuration.")
        else:
            st.success("✅ No discrepancies found in current data - no email notification needed!")
    
    else:
        st.info("Please upload one or more appeals Excel files to begin compilation.")