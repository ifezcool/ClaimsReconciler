import streamlit as st
import pyodbc
from dotenv import load_dotenv
import os
import pandas as pd
import numpy as np
from datetime import datetime

load_dotenv('secrets.env')

# #Authentication logic
# def login_screen():
#     st.header("This app is private")
#     st.subheader("Please log in")
#     st.button("Log in with Google", on_click=st.login)

# if not st.user.is_logged_in:
#     login_screen()
# else:
#     if st.user.email != os.getenv("ALLOWED_EMAIL1"):
#         st.error("You are not authorized to access this app")
#         st.button("Log out", on_click=st.logout)

#     else:
#         st.header(f"Welcome, {st.user.name}!")
#         st.button("Log out", on_click=st.logout)

# Upload file
def render_dbpage():
    uploaded_appeals_file = st.file_uploader("Upload Claims file(Please truncate claimstbl before performing any upload)", type=["xlsx"])
    if uploaded_appeals_file is not None:
        # Read with explicit handling of mixed data types
        from openpyxl import load_workbook
        import io

        # Load workbook with data_only=True to get formula results
        excel_bytes = uploaded_appeals_file.read()
        workbook = load_workbook(io.BytesIO(excel_bytes), data_only=True)
        sheet = workbook.active

        data = list(sheet.values)
        headers = [str(h).strip() if h else '' for h in data[0]]
        rows = data[1:]
        df = pd.DataFrame(rows, columns=headers)
        df.columns = df.columns.str.strip()  # Normalize headers
        st.write("DataFrame shape:", df.shape)
        st.write("Column names in Excel file:")
        st.write(list(df.columns))
        st.write(df.head())

        server = os.getenv("server")
        database = os.getenv("database")
        username = os.getenv("dbusername")
        password = os.getenv("password")

        try:
            conn = pyodbc.connect(
                "DRIVER={ODBC Driver 17 for SQL Server};SERVER="
                + server
                + ';DATABASE='
                + database
                + ';UID='
                + username
                + ';PWD='
                + password
            )
            cursor = conn.cursor()

            # Define columns mapping - Excel header to DB column
            column_mapping = {
                'S/N': 'S_N',
                'DCO NAME': 'DCO_NAME',
                'PROVIDER NAME': 'PROVIDER_NAME',
                'PROVIDER CODE': 'PROVIDER_CODE',
                'FIRST NAME': 'FIRST_NAME',
                'MIDDLE NAME': 'MIDDLE_NAME',
                'SURNAME': 'SURNAME',
                'ENROLLEE NAME': 'ENROLLEE_NAME',
                'AVON OLD ENROLEEID': 'AVON_OLD_ENROLEEID',
                'MEMBER NO': 'MEMBER_NO',
                'PLAN NAME': 'PLAN_NAME',
                'SEX': 'SEX',
                'ICD Codes': 'ICD_Codes',
                'DIAGNOSIS': 'DIAGNOSIS',
                'CPT CODES': 'CPT_CODES',
                'SERVICE DESCRIPTION': 'SERVICE_DESCRIPTION',
                'PA AMOUNT': 'PA_AMOUNT',
                'ENCOUNTER DATE (DD/MM/YYYY)': 'ENCOUNTER_DATE_DD_MM_YYYY',
                'NO. OF UNITS': 'NO_OF_UNITS',
                'AMOUNT CLAIMED': 'AMOUNT_CLAIMED',
                'DATE CLAIM RECEIVED ': 'DATE_CLAIM_RECEIVED',  # Note the space
                'AVONPACODE': 'AVONPACODE',
                'CLAIMS INTERN': 'CLAIMS_INTERN',
                'IS PA ATTACHED': 'IS_PA_ATTACHED',
                'FIRST NAME MATCH': 'FIRST_NAME_MATCH',
                'MIDDLE NAME MATCH': 'MIDDLE_NAME_MATCH',
                'SURNAME MATCH': 'SURNAME_MATCH',
                'ANY 2 OF THE 3 NAMES MATCH': 'ANY_2_OF_THE_3_NAMES_MATCH',
                'SEX MATCH': 'SEX_MATCH',
                'DIAGNOSIS MATCH': 'DIAGNOSIS_MATCH',
                'SERVICE MATCH': 'SERVICE_MATCH',
                'ENC DATE MATCH': 'ENC_DATE_MATCH',
                'UNIT MATCH': 'UNIT_MATCH',
                'AMOUNT MATCH': 'AMOUNT_MATCH',
                'CLAIM CATEGORY': 'CLAIM_CATEGORY',
                'PROVIDER RISK RATING': 'PROVIDER_RISK_RATING',
                'IS ENROLLEE REGISTD. WITH THIS PROVIDER ': 'IS_ENROLLEE_REGISTD_WITH_THIS_PROVIDER',
                'IS ENROLLEE ACTIVE': 'IS_ENROLLEE_ACTIVE',
                'IS ENROLLEE CAPITATED': 'IS_ENROLLEE_CAPITATED',
                'CLAIMS VETTER': 'CLAIMS_VETTER',
                'PROVIDER ALLOCATION': 'PROVIDER_ALLOCATION',
                'IS PA REQUIRED': 'IS_PA_REQUIRED',
                'REASON FOR PA REQUIREMENT': 'REASON_FOR_PA_REQUIREMENT',
                'IS SERVCE COVERED BY PLAN': 'IS_SERVCE_COVERED_BY_PLAN',
                'WAITING PERIOD START DATE': 'WAITING_PERIOD_START_DATE',
                'WAITING PERIOD END DATE': 'WAITING_PERIOD_END_DATE',
                'WAS WAITING PERIOD OBSERVED': 'WAS_WAITING_PERIOD_OBSERVED',
                'APPLICABLE LIMIT (UNITS)': 'APPLICABLE_LIMIT_UNITS',
                'CUMMUL. UNIT USED PTD': 'CUMMUL_UNIT_USED_PTD',
                'UNITS IN THIS CLAIM': 'UNITS_IN_THIS_CLAIM',
                'BAL. UNIT LEFT AFTER THIS CLAIM ': 'BAL_UNIT_LEFT_AFTER_THIS_CLAIM',
                'APPLICABLE LIMITS (NAIRA)': 'APPLICABLE_LIMITS_NAIRA',
                'CUMMUL. NAIRA VALUE USED PTD': 'CUMMUL_NAIRA_VALUE_USED_PTD',
                'NAIRA VALUE OF THIS REQUEST': 'NAIRA_VALUE_OF_THIS_REQUEST',
                'BAL. NAIRA LEFT AFTER THIS CLAIM ': 'BAL_NAIRA_LEFT_AFTER_THIS_CLAIM',
                'IS PROVIDER ACCREDITED TO PROVIDE SERVICE?': 'IS_PROVIDER_ACCREDITED_TO_PROVIDE_SERVICE',
                'WAS ACCURATE TARIFF USED': 'WAS_ACCURATE_TARIFF_USED',
                'CLAIMS OFFICER RECOMMD. AMT': 'CLAIMS_OFFICER_RECOMMD_AMT',
                'DIFF BTW CO RECOMMEND. &  CLAIMED': 'DIFF_BTW_CO_RECOMMEND_CLAIMED',
                'REASON FOR DIFF BTW AMT CLAIMED & AMT PAID': 'REASON_FOR_DIFF_BTW_AMT_CLAIMED_AMT_PAID',
                'COMMENT/REFERENCE': 'COMMENT_REFERENCE',
                'Additional services related to the PA originally issued': 'Additional_services_related_to_the_PA_originally_issued',
                'Agreed tariff applied (Higher)': 'Agreed_tariff_applied_Higher',
                'Agreed tariff applied (Lower)': 'Agreed_tariff_applied_Lower',
                'No tariff applied': 'No_tariff_applied',
                'Primary care Services (PA not Required)': 'Primary_care_Services_PA_not_Required',
                'PA PREVIOUSLY PAID': 'PA_PREVIOUSLY_PAID',
                'TARIFF CHECK': 'TARIFF_CHECK',
                'IS SERVCE COVERED BY PLAN.1': 'IS_SERVCE_COVERED_BY_PLAN2',
                'MGR RECOMMD. AMT': 'MGR_RECOMMD_AMT',
                'DIFF BTW MGR RECOMMD. &  CLAIMED': 'DIFF_BTW_MGR_RECOMMD_CLAIMED',
                'REASON FOR DIFF BTW AMT CLAIMED & AMT PAID.1': 'REASON_FOR_DIFF_BTW_AMT_CLAIMED_AMT_PAID3',
                'COMMENT/REFERENCE.1': 'COMMENT_REFERENCE4',
                'DIAGNOSIS CONSISTENT WITH ENROLLEE\'S AGE': 'DIAGNOSIS_CONSISTENT_WITH_ENROLLEE_S_AGE',
                'DIAGNOSIS CONSISTENT WITH ENROLLEE\'S GENDER': 'DIAGNOSIS_CONSISTENT_WITH_ENROLLEE_S_GENDER',
                'DIAGNOSIS CONSISTENT WITH SERVICE STATED': 'DIAGNOSIS_CONSISTENT_WITH_SERVICE_STATED',
                'SERVICE CONSISTENT WITH DIAGNOSIS STATED': 'SERVICE_CONSISTENT_WITH_DIAGNOSIS_STATED',
                'SERVICE CONSISTENT WITH ENROLLEE\'S AGE': 'SERVICE_CONSISTENT_WITH_ENROLLEE_S_AGE',
                'SERVICE CONSISTENT WITH ENROLLEE\'S GENDER': 'SERVICE_CONSISTENT_WITH_ENROLLEE_S_GENDER',
                'SERVICE CONSISTENT WITH TREATMENT PROTOCOL': 'SERVICE_CONSISTENT_WITH_TREATMENT_PROTOCOL',
                'HOD RECOMMD. AMOUNT': 'HOD_RECOMMD_AMOUNT',
                'DIFF BTW HOD RECOMMD. &  CLAIMED': 'DIFF_BTW_HOD_RECOMMD_CLAIMED',
                'REASON FOR DIFF BTW AMT CLAIMED & AMT RECOMMENDED': 'REASON_FOR_DIFF_BTW_AMT_CLAIMED_AMT_RECOMMENDED',
                'COMMENT/REFERENCE.2': 'COMMENT_REFERENCE5',
                'DIFF BTW PA AMOUNT & AMOUNT RECOMMENDED': 'DIFF_BTW_PA_AMOUNT_AMOUNT_RECOMMENDED',
                'REASON FOR VARIANCE BTW PA AMOUNT & AMOUNT RECOMMENDED': 'REASON_FOR_VARIANCE_BTW_PA_AMOUNT_AMOUNT_RECOMMENDED',
                'COMMENT/REFERENCE.3': 'COMMENT_REFERENCE6',
                'SCH NO': 'SCH_NO',
                'Name & Date': 'Name_Date',
                'Enrollee & Date': 'Enrollee_Date',
                'ReviewedDate': 'ReviewedDate',
                'PostedDate': 'PostedDate',
                'PaidDate': 'PaidDate',
                'ClaimBatch': 'ClaimBatch',
                'ClaimNoFnx': 'ClaimNoFnx',
                'ClaimNo': 'ClaimNo',
                'Correct_ClaimNo': 'Correct_ClaimNo',
                'Benefits': 'Benefits',
                'ProviderClass': 'ProviderClass',
                'OpdIpd': 'OpdIpd'
            }

            # Get the actual database column names in order
            db_columns = list(column_mapping.values())
            
            # Date columns that need special handling
            date_columns = ['ENCOUNTER_DATE_DD_MM_YYYY', 'DATE_CLAIM_RECEIVED', 'ReviewedDate', 'PostedDate', 'PaidDate']

            # Create table with proper data types
            column_definitions = []
            for col in db_columns:
                if col in date_columns:
                    column_definitions.append(f'{col} datetime2')
                else:
                    column_definitions.append(f'{col} VARCHAR(MAX)')

            create_table_query = f"""
            IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='claimstbl' AND xtype='U')
            CREATE TABLE claimstbl (
                {', '.join(column_definitions)}
            )
            """
            cursor.execute(create_table_query)

            # Function to safely convert date strings
            def convert_date(date_str):
                if pd.isna(date_str) or date_str == '' or date_str == 'NIL':
                    return None
                try:
                    # Try parsing as datetime if it's already a datetime object
                    if isinstance(date_str, datetime):
                        return date_str
                    # Try different date formats
                    for fmt in ['%m/%d/%Y %H:%M', '%d/%m/%Y %H:%M', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d']:
                        try:
                            return datetime.strptime(str(date_str), fmt)
                        except ValueError:
                            continue
                    return None
                except:
                    return None

            # Function to clean and prepare value
            def clean_value(val, col_name):
                if pd.isna(val) or val == '' or str(val).strip() == '':
                                    return None
                
                # Handle date columns
                if col_name in date_columns:
                    return convert_date(val)
                
                # Handle numeric columns that might be stored as strings
                if col_name in ['PA_AMOUNT', 'AMOUNT_CLAIMED', 'NO_OF_UNITS']:
                    try:
                        # Clean numeric strings
                        clean_val = str(val).replace(',', '').strip()
                        return float(clean_val) if clean_val and clean_val != 'NIL' else None
                    except:
                        return str(val).strip()
                
                # Regular string columns
                return str(val).strip()

            # Prepare the INSERT statement
            insert_query = f"""
            INSERT INTO claimstbl ({', '.join(db_columns)}) 
            VALUES ({', '.join(['?'] * len(db_columns))})
            """

            # Debug: Show column mapping mismatches
            st.write("Checking column mapping:")
            excel_cols = set(df.columns)
            mapped_cols = set(column_mapping.keys())
            missing_in_excel = mapped_cols - excel_cols
            missing_in_mapping = excel_cols - mapped_cols
            
            if missing_in_excel:
                st.warning(f"Columns in mapping but not in Excel: {missing_in_excel}")
            if missing_in_mapping:
                st.warning(f"Columns in Excel but not in mapping: {missing_in_mapping}")

            # Insert data with proper error handling
            progress_bar = st.progress(0)
            total_rows = len(df)
            successful_inserts = 0
            
            for index, row in df.iterrows():
                try:
                    values = []
                    for db_col in db_columns:
                        # Find the Excel column name for this DB column
                        excel_col = None
                        for excel_name, db_name in column_mapping.items():
                            if db_name == db_col:
                                excel_col = excel_name
                                break
                        
                        if excel_col and excel_col in df.columns:
                            val = clean_value(row[excel_col], db_col)
                            values.append(val)
                        else:
                            values.append(None)
                    
                    cursor.execute(insert_query, values)
                    successful_inserts += 1
                    progress_bar.progress((index + 1) / total_rows)
                    
                except Exception as e:
                    st.error(f"Error inserting row {index + 1}: {str(e)}")
                    st.error(f"Row data: {dict(row)}")
                    # Show the problematic values
                    st.error(f"Values being inserted: {values[:10]}...")  # Show first 10 values
                    break
            
            conn.commit()
            st.success(f"Successfully uploaded {successful_inserts} out of {total_rows} records to the server")
            
            # Show some statistics
            st.write("Upload Summary:")
            st.write(f"Total rows processed: {total_rows}")
            st.write(f"Successfully inserted: {successful_inserts}")
            st.write(f"Failed inserts: {total_rows - successful_inserts}")
            
        except Exception as e:
            st.error(f"Database connection error: {str(e)}")
            
        finally:
            if 'conn' in locals():
                conn.close()
                st.write("Upload is complete, please copy the following command and run it on SSMS")
                st.write("INSERT INTO [Claims Schedules Consolidated Mastersheet] SELECT * FROM claimstbl")