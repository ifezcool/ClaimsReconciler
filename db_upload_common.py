import streamlit as st
import pyodbc
from dotenv import load_dotenv
import os
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
import io
from config import logger

load_dotenv('secrets.env')

def _get_connection():
    return pyodbc.connect(
        "DRIVER={ODBC Driver 17 for SQL Server};"
        f"SERVER={os.getenv('server')};"
        f"DATABASE={os.getenv('database')};"
        f"UID={os.getenv('dbusername')};"
        f"PWD={os.getenv('password')}"
    )

def _convert_date(date_val):
    if pd.isna(date_val) or date_val in ['', 'NIL']:
        return None
    if isinstance(date_val, datetime):
        return date_val
    for fmt in ['%d/%m/%Y', '%Y-%m-%d', '%m/%d/%Y', '%m/%d/%Y %H:%M', '%d/%m/%Y %H:%M', '%Y-%m-%d %H:%M:%S']:
        try:
            return datetime.strptime(str(date_val), fmt)
        except ValueError:
            continue
    return None

def _clean_value(val, col_name, date_columns, numeric_columns):
    if pd.isna(val) or str(val).strip() == '':
        return None
    if col_name in date_columns:
        return _convert_date(val)
    if col_name in numeric_columns:
        try:
            clean = str(val).replace(',', '').strip()
            if clean and clean != 'NIL':
                col_type = numeric_columns[col_name]
                if col_type == 'INT':
                    return int(float(clean))
                return float(clean)
            return None
        except (ValueError, TypeError):
            return str(val).strip()
    if col_name in ['BATCH_NUMBER', 'PROVIDER_CODE']:
        try:
            return str(int(float(val)))
        except (ValueError, TypeError):
            pass
    return str(val).strip()

def _build_column_definitions(db_columns, date_columns, numeric_columns):
    defs = []
    for col in db_columns:
        if col in date_columns:
            defs.append(f'{col} datetime2')
        elif col in numeric_columns:
            defs.append(f'{col} {numeric_columns[col]}')
        else:
            defs.append(f'{col} VARCHAR(MAX)')
    return defs

def render_generic_upload(
    table_name,
    column_mapping,
    date_columns=None,
    numeric_columns=None,
    consolidate_target=None,
    file_label="file",
    uploader_help="",
):
    if date_columns is None:
        date_columns = []
    if numeric_columns is None:
        numeric_columns = {}

    db_columns = list(column_mapping.values())

    uploaded_file = st.file_uploader(
        f"Upload {file_label} file",
        type=["xlsx"],
        help=uploader_help,
    )
    if uploaded_file is None:
        return

    excel_bytes = uploaded_file.read()
    workbook = load_workbook(io.BytesIO(excel_bytes), data_only=True)
    sheet = workbook.active
    data = list(sheet.values)
    headers = [str(h).strip() if h else '' for h in data[0]]
    rows = data[1:]
    df = pd.DataFrame(rows, columns=headers)
    df.columns = df.columns.str.strip()

    st.write("Preview of uploaded data:")
    st.write(df.head())

    excel_cols = set(df.columns)
    mapped_cols = set(column_mapping.keys())
    missing_in_excel = mapped_cols - excel_cols
    missing_in_mapping = excel_cols - mapped_cols

    if missing_in_excel:
        st.warning(f"Columns in mapping but not in Excel: {missing_in_excel}")
    if missing_in_mapping:
        st.info(f"Columns in Excel but not in mapping: {missing_in_mapping}")

    required = set(column_mapping.keys()) - {'VARIANCE', 'VARIANCE1', 'NARRATION', 'NARRATIVE',
        'Paiddate', 'SCH_NO', 'APPEAL_NO', 'SCH_NUM', 'Source_File'}
    missing_required = required - excel_cols
    if missing_required:
        st.error(f"Required columns missing from Excel: {missing_required}. Refusing to proceed.")
        return

    truncate_ok = st.checkbox(f"I have truncated table '{table_name}' before uploading", value=False)
    if not truncate_ok:
        st.warning(f"Please confirm that '{table_name}' has been truncated before uploading.")
        proceed = st.checkbox("Proceed without truncation confirmation", value=False)
        if not proceed:
            st.info("Check the box above to confirm truncation or proceed anyway.")
            return

    try:
        conn = _get_connection()
        cursor = conn.cursor()
        conn.autocommit = False

        col_defs = _build_column_definitions(db_columns, date_columns, numeric_columns)
        create_query = f"""
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='{table_name}' AND xtype='U')
        CREATE TABLE {table_name} (
            {', '.join(col_defs)}
        )
        """
        cursor.execute(create_query)

        insert_query = f"""
        INSERT INTO {table_name} ({', '.join(db_columns)})
        VALUES ({', '.join(['?'] * len(db_columns))})
        """

        progress_bar = st.progress(0)
        total_rows = len(df)
        success_count = 0
        failed_rows = []

        for i, row in df.iterrows():
            try:
                values = []
                for db_col in db_columns:
                    excel_col = next(k for k, v in column_mapping.items() if v == db_col)
                    if excel_col in df.columns:
                        values.append(_clean_value(row[excel_col], db_col, date_columns, numeric_columns))
                    else:
                        values.append(None)
                cursor.execute(insert_query, values)
                success_count += 1
            except Exception as e:
                failed_rows.append((i + 1, str(e)))
                logger.warning(f"Row {i+1} failed in {table_name}: {e}")
            progress_bar.progress((i + 1) / total_rows)

        if failed_rows:
            st.warning(f"Failed rows ({len(failed_rows)}):")
            for idx, err in failed_rows:
                st.write(f"  Row {idx}: {err}")

        if success_count > 0:
            conn.commit()
            st.success(f"Uploaded {success_count}/{total_rows} rows successfully to '{table_name}'.")

            if consolidate_target:
                st.info(f"Consolidate to '{consolidate_target}'?")
                if st.checkbox(f"INSERT INTO {consolidate_target} SELECT * FROM {table_name}"):
                    try:
                        cursor.execute(f"INSERT INTO {consolidate_target} SELECT * FROM {table_name}")
                        conn.commit()
                        st.success(f"Data consolidated into '{consolidate_target}' successfully!")
                    except Exception as e:
                        st.error(f"Consolidation failed: {e}")
                        conn.rollback()
        else:
            conn.rollback()
            st.error("No rows were successfully inserted. Transaction rolled back.")

    except Exception as e:
        st.error(f"Database error: {e}")
        logger.error(f"Database error in {table_name} upload: {e}", exc_info=True)
        if 'conn' in locals():
            try:
                conn.rollback()
            except Exception:
                pass
    finally:
        if 'conn' in locals():
            conn.close()
