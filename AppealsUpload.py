import streamlit as st
import pyodbc
from dotenv import load_dotenv
import os
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
import io

load_dotenv('secrets.env')

def render_appeals_upload():
    uploaded_file = st.file_uploader("Upload Appeals file (Please truncate appealstbl before uploading)", type=["xlsx"])
    if uploaded_file is not None:
        # Read Excel with formulas resolved
        excel_bytes = uploaded_file.read()
        workbook = load_workbook(io.BytesIO(excel_bytes), data_only=True)
        sheet = workbook.active

        data = list(sheet.values)
        headers = [str(h).strip() if h else '' for h in data[0]]
        rows = data[1:]
        df = pd.DataFrame(rows, columns=headers)
        df.columns = df.columns.str.strip()

        st.write("**Excel Column Names Found:**")
        st.write(df.columns.tolist())
        
        st.write("**Preview of uploaded data:**")
        st.write(df.head())

        # DB credentials
        server = os.getenv("server")
        database = os.getenv("database")
        username = os.getenv("dbusername")
        password = os.getenv("password")

        # Column mapping: Excel column -> Database column
        column_mapping = {
            'S_N': 'S_N',
            'CLAIM_TYPE': 'CLAIM_TYPE',
            'BATCH_NUMBER': 'BATCH_NUMBER',
            'HOSPITAL': 'HOSPITAL',
            'NUMBER_OF_CLAIMS': 'NUMBER_OF_CLAIMS',
            'ENCOUNTER_MONTH': 'ENCOUNTER_MONTH',
            'DATE_OF_RECEIPT': 'DATE_OF_RECEIPT',
            'APPROVED_PA_VALUE_N': 'APPROVED_PA_VALUE_N',
            'AMOUNT_RECOMMENDED_FOR_PAYMENT_N': 'AMOUNT_RECOMMENDED_FOR_PAYMENT_N',
            'VARIANCE': 'VARIANCE',
            'VARIANCE1': 'VARIANCE1',
            'NARRATION': 'NARRATION',
            'Source_File': 'Source_File',
            'PROVIDER_CODE': 'PROVIDER_CODE',
            'Paiddate': 'Paiddate',
            'SCH_NO': 'SCH_NO',
            'APPEAL_NO': 'APPEAL_NO',
            'SCH_NUM': 'SCH_NUM'
        }

        # Validate columns exist
        missing_cols = [col for col in column_mapping.keys() if col not in df.columns]
        if missing_cols:
            st.error(f"**Missing columns in Excel file:** {missing_cols}")
            st.warning("**Available columns:** " + ", ".join(df.columns.tolist()))
            st.stop()

        try:
            conn = pyodbc.connect(
                "DRIVER={ODBC Driver 17 for SQL Server};"
                f"SERVER={server};DATABASE={database};UID={username};PWD={password}"
            )
            cursor = conn.cursor()

            db_columns = list(column_mapping.values())
            date_columns = ['DATE_OF_RECEIPT', 'Paiddate']

            # Create table
            column_definitions = []
            for col in db_columns:
                if col in date_columns:
                    column_definitions.append(f'[{col}] datetime2')
                else:
                    column_definitions.append(f'[{col}] VARCHAR(MAX)')

            create_table_query = f"""
            IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='appealstbl' AND xtype='U')
            CREATE TABLE appealstbl (
                {', '.join(column_definitions)}
            )
            """
            cursor.execute(create_table_query)
            conn.commit()

            # Date converter
            def convert_date(date_val):
                if pd.isna(date_val) or date_val in ['', 'NIL', None]:
                    return None
                if isinstance(date_val, datetime):
                    return date_val
                for fmt in ['%d/%m/%Y', '%Y-%m-%d', '%m/%d/%Y', '%d-%m-%Y']:
                    try:
                        return datetime.strptime(str(date_val), fmt)
                    except (ValueError, AttributeError):
                        continue
                return None

            # Cleaner
            def clean_value(val, col_name):
                if pd.isna(val):
                    return None
                val_str = str(val).strip()
                if val_str == '' or val_str.upper() == 'NIL':
                    return None
                if col_name in date_columns:
                    return convert_date(val)
                return val_str

            # Insert query
            insert_query = f"""
            INSERT INTO appealstbl ([{'], ['.join(db_columns)}])
            VALUES ({', '.join(['?'] * len(db_columns))})
            """

            progress_bar = st.progress(0)
            total_rows = len(df)
            success_count = 0
            error_rows = []

            for i, row in df.iterrows():
                try:
                    values = []
                    for excel_col, db_col in column_mapping.items():
                        cell_value = row.get(excel_col)
                        cleaned = clean_value(cell_value, db_col)
                        values.append(cleaned)
                    
                    cursor.execute(insert_query, values)
                    success_count += 1
                except Exception as e:
                    error_rows.append((i+1, str(e)))
                    if len(error_rows) <= 5:  # Show first 5 errors
                        st.warning(f"Row {i+1} failed: {e}")
                
                progress_bar.progress((i + 1) / total_rows)

            conn.commit()
            
            if error_rows:
                st.warning(f"⚠️ {len(error_rows)} rows failed out of {total_rows}")
            
            st.success(f"✅ Successfully uploaded {success_count}/{total_rows} rows")

        except Exception as e:
            st.error(f"❌ Database error: {e}")
            import traceback
            st.code(traceback.format_exc())
        finally:
            if 'conn' in locals():
                conn.close()
            
            if success_count > 0:
                st.info("**Next Step:** Run this in SSMS to consolidate:")
                st.code("INSERT INTO Compiled_Appeals SELECT * FROM appealstbl", language="sql")