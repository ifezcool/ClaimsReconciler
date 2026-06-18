import streamlit as st
import pandas as pd
import io
from datetime import datetime
import re
import html
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
from dotenv import load_dotenv
from config import get_cc_list, get_to_email, logger
from utils import validate_email_list

load_dotenv('secrets.env')

COMPILATION_CONFIGS = {
    "telemedicine": {
        "label": "Telemedicine",
        "label_lower": "telemedicine",
        "sheet_name": "Compiled Telemedicine",
        "amount_label": "Telemedicine_Amount",
        "session_compiled": "telemedicine_compiled_data",
        "session_comparison": "telemedicine_comparison_df",
        "session_uploader": "telemedicine_uploader",
        "session_finance_uploader": "finance_uploader_telemedicine",
        "email_subject": "Telemedicine Finance Comparison Alert",
        "page_header": "Telemedicine Compilation & Finance Comparison",
        "finance_comparison_filename": "telemedicine_finance_comparison",
        "compiled_filename_prefix": "compiled_telemedicine",
    },
    "ambulance": {
        "label": "Ambulance",
        "label_lower": "ambulance",
        "sheet_name": "Compiled Ambulance",
        "amount_label": "Ambulance_Amount",
        "session_compiled": "ambulance_compiled_data",
        "session_comparison": "ambulance_comparison_df",
        "session_uploader": "ambulance_uploader",
        "session_finance_uploader": "finance_uploader_ambulance",
        "email_subject": "Ambulance Finance Comparison Alert",
        "page_header": "Ambulance Compilation & Finance Comparison",
        "finance_comparison_filename": "ambulance_finance_comparison",
        "compiled_filename_prefix": "compiled_ambulance",
    },
    "appeals": {
        "label": "Appeals",
        "label_lower": "appeals",
        "sheet_name": "Compiled Appeals",
        "amount_label": "Appeals_Amount",
        "session_compiled": "compiled_data",
        "session_comparison": "comparison_df",
        "session_uploader": "appeals_uploader",
        "session_finance_uploader": "finance_uploader",
        "email_subject": "Appeals Finance Comparison Alert",
        "page_header": "Appeals Compilation & Finance Comparison",
        "finance_comparison_filename": "appeals_finance_comparison",
        "compiled_filename_prefix": "compiled_appeals",
    },
}

TEMPLATE_COLUMNS = [
    'S_N', 'CLAIM_TYPE', 'BATCH_NUMBER', 'HOSPITAL', 'NUMBER_OF_CLAIMS',
    'ENCOUNTER_MONTH', 'DATE_OF_RECEIPT', 'APPROVED_PA_VALUE_N',
    'AMOUNT_RECOMMENDED_FOR_PAYMENT_N', 'VARIANCE', 'VARIANCE1',
    'NARRATION', 'Source_File', 'PROVIDER_CODE', 'Paiddate',
    'SCH_NO', 'APPEAL_NO', 'SCH_NUM'
]
BLANK_COLUMNS = ['Paiddate', 'SCH_NO', 'APPEAL_NO', 'SCH_NUM']

COLUMN_MAPPING = {
    'S/N': 'S_N',
    'AMOUNT RECOMMENDED FOR PAYMENT (N)': 'AMOUNT_RECOMMENDED_FOR_PAYMENT_N',
    'ENROLLEE NAME': 'HOSPITAL',
    'PROVIDER NAME': 'HOSPITAL',
    'HOSPITAL NAME': 'HOSPITAL',
    'HOSPITAL': 'HOSPITAL',
    'PROVIDER CODE': 'PROVIDER_CODE',
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
    'NARRATIVE': 'NARRATION',
}

def compile_files(uploaded_files):
    compiled_data = []
    file_summary = []

    for uploaded_file in uploaded_files:
        try:
            excel_file = pd.ExcelFile(uploaded_file)
            if 'PAYMENT SUMMARY' not in excel_file.sheet_names:
                file_summary.append({'File': uploaded_file.name, 'Rows': 0, 'Status': 'No PAYMENT SUMMARY sheet found'})
                continue

            df = pd.read_excel(uploaded_file, sheet_name='PAYMENT SUMMARY', header=1)
            if df.empty:
                file_summary.append({'File': uploaded_file.name, 'Rows': 0, 'Status': 'Empty sheet'})
                continue

            df = df.dropna(how='all')
            data_rows = []
            for i in range(len(df)):
                row = df.iloc[i]
                first_col = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
                if any(w in first_col.upper() for w in ['TOTAL', 'SUM', 'GRAND', 'SUBTOTAL', 'SUMMARY']):
                    continue
                if first_col == "" and row.isna().all():
                    continue
                try:
                    float(first_col)
                    data_rows.append(i)
                except ValueError:
                    non_empty = sum(1 for val in row if pd.notna(val) and str(val).strip() != "")
                    if non_empty >= 3:
                        data_rows.append(i)

            if not data_rows:
                file_summary.append({'File': uploaded_file.name, 'Rows': 0, 'Status': 'No valid data rows found'})
                continue
            df_clean = df.iloc[data_rows].copy()

            data_dict = {}
            for col in TEMPLATE_COLUMNS:
                data_dict[col] = [''] * len(df_clean)
            standardized_df = pd.DataFrame(data_dict)

            for orig_col, new_col in COLUMN_MAPPING.items():
                if orig_col in df_clean.columns and new_col not in BLANK_COLUMNS:
                    standardized_df[new_col] = df_clean[orig_col].apply(
                        lambda x: str(int(float(x))) if pd.notna(x) and str(x).replace('.', '').isdigit()
                        else str(x) if pd.notna(x) else ''
                    )

            standardized_df['Source_File'] = uploaded_file.name
            compiled_data.append(standardized_df)
            file_summary.append({'File': uploaded_file.name, 'Rows': len(df_clean), 'Status': 'Success'})

        except Exception as e:
            file_summary.append({'File': uploaded_file.name, 'Rows': 0, 'Status': f'Error: {str(e)}'})
            logger.warning(f"Compile failed for {uploaded_file.name}: {e}")

    return compiled_data, file_summary

def create_compiled_excel(compiled_data, sheet_name="Compiled Data"):
    if not compiled_data:
        return None
    combined_df = pd.concat(compiled_data, ignore_index=True)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)
    return output.getvalue()

def extract_schedule_from_filename(filename):
    pattern = r'(?:Schedule|SCH)\s*(\d+)'
    match = re.search(pattern, filename, re.IGNORECASE)
    return match.group(1) if match else None

def compare_with_finance(compiled_data, finance_file, config):
    if not compiled_data or not finance_file:
        return None

    label_lower = config["label_lower"]
    amount_label = config["amount_label"]

    try:
        finance_xls = pd.ExcelFile(finance_file)
        finance_sheet = None
        for sheet in finance_xls.sheet_names:
            if 'CLAIMS RECEIVED' in sheet.upper() or 'WEEKLY REPORT' in sheet.upper():
                finance_sheet = sheet
                break
        if not finance_sheet:
            return None

        finance_df = pd.read_excel(finance_file, sheet_name=finance_sheet)
        combined = pd.concat(compiled_data, ignore_index=True)

        summary = []
        for _, row in combined.iterrows():
            fname = row.get('Source_File', '')
            schedule_num = extract_schedule_from_filename(fname)
            amt_col = 'AMOUNT_RECOMMENDED_FOR_PAYMENT_N'
            if schedule_num and amt_col in row:
                try:
                    amount = float(str(row[amt_col]).replace(',', ''))
                    summary.append({'Schedule_Number': schedule_num, amount_label: amount, 'Source_File': fname})
                except (ValueError, TypeError):
                    continue

        if not summary:
            return None

        summary_df = pd.DataFrame(summary)
        grouped = summary_df.groupby('Schedule_Number').agg({
            amount_label: 'sum',
            'Source_File': lambda x: ', '.join(x.unique())
        }).reset_index()

        results = []
        for _, crow in grouped.iterrows():
            sch = crow['Schedule_Number']
            cat_amt = crow[amount_label]
            finance_df['_sch_num'] = pd.to_numeric(finance_df['Claim Batch No/Sch No'], errors='coerce')
            ffiltered = finance_df[finance_df['_sch_num'] == float(sch)]
            if not ffiltered.empty:
                fin_amt = pd.to_numeric(ffiltered['Claims_Advised_Amount'], errors='coerce').sum()
                variance = cat_amt - fin_amt
                results.append({
                    'Schedule_Number': sch,
                    amount_label: cat_amt,
                    'Finance_Amount': fin_amt,
                    'Variance': variance,
                    'Source_Files': crow['Source_File']
                })
            else:
                results.append({
                    'Schedule_Number': sch,
                    amount_label: cat_amt,
                    'Finance_Amount': 0,
                    'Variance': cat_amt,
                    'Source_Files': crow['Source_File']
                })

        return pd.DataFrame(results)

    except Exception as e:
        st.error(f"Error comparing with finance data: {str(e)}")
        logger.error(f"compare_with_finance error for {label_lower}: {e}", exc_info=True)
        return None

def send_notification_email(missing_schedules, amount_mismatches, config):
    sender_email = os.getenv("OFFICE_SENDER_EMAIL")
    password = os.getenv("OUTLOOK_APP_PASSWORD")
    recipient_email = get_to_email()
    cc_list = get_cc_list("default")

    if not sender_email or not password:
        st.error("Office 365 credentials not configured.")
        return False

    try:
        validate_email_list([recipient_email], context=f"send_notification_email/{config['label_lower']}/to")
        validate_email_list(cc_list, context=f"send_notification_email/{config['label_lower']}/cc")
    except ValueError as e:
        logger.error(f"Invalid email addresses: {e}")
        st.error(f"Invalid email addresses. Check config: {e}")
        return False

    label = config["label"]
    amount_label = config["amount_label"]
    subject = config["email_subject"]

    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = recipient_email
    message['Cc'] = ", ".join(cc_list)
    message["Subject"] = f"{subject} - Discrepancies Found"

    body = f"""<html><body>
        <h2>{label} Finance Comparison Alert</h2>
        <p>This is an automated notification from the {label} Compilation system.</p>
        <p><strong>Comparison completed at:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        <h3>Summary:</h3>
        <ul>
            <li><strong>Schedules missing in Finance:</strong> {len(missing_schedules)}</li>
            <li><strong>Amount mismatches:</strong> {len(amount_mismatches)}</li>
        </ul>
    """

    if not missing_schedules.empty:
        body += """<h3>🚨 Schedules NOT in Finance:</h3>
        <table border="1" style="border-collapse: collapse; width: 100%;">
            <tr style="background-color: #f2f2f2;"><th>Schedule Number</th><th>{label} Amount</th><th>Source Files</th></tr>
        """.replace("{label}", label)
        for _, row in missing_schedules.iterrows():
            sch_esc = html.escape(str(row['Schedule_Number']))
            src_esc = html.escape(str(row['Source_Files']))
            body += f"<tr><td>{sch_esc}</td><td>{row[amount_label]:,.2f}</td><td>{src_esc}</td></tr>"
        body += "</table><br>"
        body += f"<p><strong>Total amount missing in Finance:</strong> {missing_schedules[amount_label].sum():,.2f}</p>"

    if not amount_mismatches.empty:
        body += """<h3>⚠️ Amount Mismatches:</h3>
        <table border="1" style="border-collapse: collapse; width: 100%;">
            <tr style="background-color: #f2f2f2;"><th>Schedule Number</th><th>{label} Amount</th><th>Finance Amount</th><th>Variance</th><th>Source Files</th></tr>
        """.replace("{label}", label)
        for _, row in amount_mismatches.iterrows():
            vcolor = "red" if row['Variance'] > 0 else "blue"
            sch_esc = html.escape(str(row['Schedule_Number']))
            src_esc = html.escape(str(row['Source_Files']))
            body += f"<tr><td>{sch_esc}</td><td>{row[amount_label]:,.2f}</td><td>{row['Finance_Amount']:,.2f}</td><td style='color: {vcolor};'>{row['Variance']:,.2f}</td><td>{src_esc}</td></tr>"
        body += "</table><br>"
        body += f"<p><strong>Total variance:</strong> {amount_mismatches['Variance'].sum():,.2f}</p>"

    body += """
        <p><strong>Action Required:</strong></p>
        <ul>
            <li>Review the missing schedules in the finance system</li>
            <li>Investigate amount discrepancies for variance resolution</li>
            <li>Check the Compilation system for detailed reports</li>
        </ul>
        <p>This is an automated message. Please do not reply.</p>
    </body></html>
    """

    message.attach(MIMEText(body, "html"))

    try:
        server = smtplib.SMTP("smtp.office365.com", 587)
        server.starttls()
        server.login(sender_email, password)
        all_recipients = [recipient_email] + cc_list
        server.sendmail(sender_email, all_recipients, message.as_string())
        server.quit()
        st.success("Email notification sent successfully!")
        return True
    except Exception as e:
        st.error(f"Could not send email: {str(e)}")
        logger.error(f"send_notification_email failed for {label}: {e}", exc_info=True)
        return False

def show_compilation_page(config):
    label = config["label"]
    label_lower = config["label_lower"]
    sheet_name = config["sheet_name"]
    amount_label = config["amount_label"]
    session_compiled = config["session_compiled"]
    session_comparison = config["session_comparison"]
    session_uploader = config["session_uploader"]
    session_finance_uploader = config["session_finance_uploader"]

    st.header(config["page_header"])
    st.markdown(f"""
    Upload multiple {label} Excel files to compile their PAYMENT SUMMARY sheets and compare with finance data.
    """)

    col1, col2 = st.columns(2)

    with col1:
        st.subheader(f"{label} Files")
        uploaded_files = st.file_uploader(
            f"Upload {label} Excel Files",
            type=['xlsx', 'xls'],
            accept_multiple_files=True,
            help=f"Select multiple Excel files containing {label} data with PAYMENT SUMMARY sheets",
            key=session_uploader
        )

    with col2:
        st.subheader("Finance File")
        finance_file = st.file_uploader(
            "Upload Finance Excel File",
            type=['xlsx', 'xls'],
            help="Upload the finance file containing 'CLAIMS RECEIVED WEEKLY REPORT' sheet",
            key=session_finance_uploader
        )

    if uploaded_files:
        st.subheader(f"{label} Files Selected: {len(uploaded_files)}")
        for i, f in enumerate(uploaded_files, 1):
            sch = extract_schedule_from_filename(f.name)
            note = f"→ Schedule {sch}" if sch else "→ ⚠️ Could not extract schedule number"
            st.write(f"{i}. {f.name} {note}")

        process_option = st.radio(
            "Choose processing option:",
            [f"Compile {label} Only", f"Compile {label} + Compare with Finance"],
        )

        if st.button(f"Process {label} Files", type="primary"):
            with st.spinner(f"Processing {label} files..."):
                compiled_data, file_summary = compile_files(uploaded_files)

                st.subheader("Processing Summary")
                st.dataframe(pd.DataFrame(file_summary), use_container_width=True)

                if compiled_data:
                    excel_bytes = create_compiled_excel(compiled_data, sheet_name)
                    if excel_bytes:
                        total_rows = sum(len(d) for d in compiled_data)
                        successful = sum(1 for s in file_summary if s['Status'] == 'Success')
                        st.success(f"Successfully compiled {successful} files with {total_rows} total rows")

                        st.session_state[session_compiled] = compiled_data

                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        compiled_filename = f"{config['compiled_filename_prefix']}_{timestamp}.xlsx"
                        st.download_button(
                            label=f"📥 Download Compiled {label}",
                            data=excel_bytes,
                            file_name=compiled_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )

                        if process_option == f"Compile {label} + Compare with Finance" and finance_file:
                            st.subheader("Finance Comparison")
                            with st.spinner("Comparing with finance data..."):
                                comparison_df = compare_with_finance(compiled_data, finance_file, config)

                                if comparison_df is not None and not comparison_df.empty:
                                    st.success("Finance comparison completed!")
                                    st.session_state[session_comparison] = comparison_df
                                    st.dataframe(comparison_df, use_container_width=True)

                                    total_variance = comparison_df['Variance'].sum()
                                    matches = len(comparison_df[comparison_df['Variance'] == 0])
                                    mismatches = len(comparison_df[comparison_df['Variance'] != 0])
                                    missing_in_finance = len(comparison_df[comparison_df['Finance_Amount'] == 0])

                                    c1, c2, c3, c4 = st.columns(4)
                                    c1.metric("Total Variance", f"{total_variance:,.2f}")
                                    c2.metric("Exact Matches", matches)
                                    c3.metric("Amount Variances", mismatches - missing_in_finance)
                                    c4.metric("Missing in Finance", missing_in_finance)

                                    if missing_in_finance > 0:
                                        st.error(f"🚨 {missing_in_finance} schedule(s) missing in finance!")
                                    if mismatches - missing_in_finance > 0:
                                        st.warning(f"⚠️ {mismatches - missing_in_finance} amount discrepancies!")
                                    if matches == len(comparison_df):
                                        st.success("✅ All schedules match perfectly!")

                                    if missing_in_finance > 0 or (mismatches - missing_in_finance) > 0:
                                        st.subheader("📧 Email Notification")
                                        st.info("Discrepancies found! You can send an email notification.")

                                        if st.button("📧 Send Email Notification", type="secondary", key=f"send_email_{label_lower}"):
                                            with st.spinner("Sending email..."):
                                                ms = comparison_df[comparison_df['Finance_Amount'] == 0]
                                                am = comparison_df[(comparison_df['Finance_Amount'] != 0) & (comparison_df['Variance'] != 0)]
                                                send_notification_email(ms, am, config)
                                    else:
                                        st.success("No discrepancies found!")

                                    comparison_output = io.BytesIO()
                                    with pd.ExcelWriter(comparison_output, engine='openpyxl') as writer:
                                        comparison_df.to_excel(writer, sheet_name='Finance Comparison', index=False)
                                        pd.concat(compiled_data, ignore_index=True).to_excel(writer, sheet_name=sheet_name, index=False)
                                    comparison_output.seek(0)

                                    st.download_button(
                                        label="📥 Download Comparison Report",
                                        data=comparison_output.getvalue(),
                                        file_name=f"{config['finance_comparison_filename']}_{timestamp}.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    )
                                else:
                                    st.warning("Could not perform finance comparison. Check that the finance file contains 'CLAIMS RECEIVED WEEKLY REPORT' sheet.")

                        elif process_option == f"Compile {label} + Compare with Finance" and not finance_file:
                            st.warning("Please upload a finance file to perform comparison.")

                        if st.checkbox("Show Preview of Compiled Data"):
                            combined = pd.concat(compiled_data, ignore_index=True)
                            st.dataframe(combined.head(20), use_container_width=True)

    if session_comparison in st.session_state:
        st.markdown("---")
        st.subheader("📧 Manual Email Notification")
        comparison_df = st.session_state[session_comparison]
        missing_in_finance = len(comparison_df[comparison_df['Finance_Amount'] == 0])
        amount_mismatches = len(comparison_df[(comparison_df['Finance_Amount'] != 0) & (comparison_df['Variance'] != 0)])

        if missing_in_finance > 0 or amount_mismatches > 0:
            c1, c2 = st.columns([2, 1])
            with c1:
                st.write(f"**Issues:** {missing_in_finance} missing, {amount_mismatches} amount mismatches")
            with c2:
                if st.button("📧 Send Email for Current Issues", type="secondary", key=f"manual_email_{label_lower}"):
                    with st.spinner("Sending email..."):
                        ms = comparison_df[comparison_df['Finance_Amount'] == 0]
                        am = comparison_df[(comparison_df['Finance_Amount'] != 0) & (comparison_df['Variance'] != 0)]
                        send_notification_email(ms, am, config)
        else:
            st.success("No discrepancies found!")
    elif not uploaded_files:
        st.info(f"Please upload one or more {label} Excel files to begin compilation.")
