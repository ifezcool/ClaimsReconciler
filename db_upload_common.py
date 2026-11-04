import streamlit as st
import pyodbc
from dotenv import load_dotenv
import math
import os
import hashlib
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
    raise ValueError(f"Unrecognized date format: '{date_val}' — expected DD/MM/YYYY, YYYY-MM-DD, or MM/DD/YYYY")

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
    result = str(val).strip()
    if isinstance(val, float) and math.isfinite(val) and val == int(val):
        result = str(int(val))
    return result

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

def _normalize_header_text(text: str) -> str:
    """Normalize a header name: strip, collapse whitespace, replace NBSP, casefold."""
    if not isinstance(text, str):
        text = str(text)
    # Replace non-breaking spaces (U+00A0) with regular spaces
    text = text.replace('\u00a0', ' ')
    # Strip leading/trailing whitespace
    text = text.strip()
    # Collapse inner whitespace (multiple spaces/tabs/newlines -> single space)
    text = ' '.join(text.split())
    # Case-insensitive comparison
    return text.casefold()


def _apply_column_aliases(df, column_aliases):
    """Apply column alias mapping to DataFrame columns.

    Returns: (df_renamed, info_messages: list[str], error_msg: str | None)
    - If column_aliases is falsy: return (df, [], None).
    - For EACH (canonical, alternates) pair independently:
        canon_cols = columns whose normalized name equals normalized canonical
        alt_cols   = columns whose normalized name matches any normalized alternate
      * canon_cols and alt_cols both non-empty -> return (df, [], error naming both original headers
        and saying to remove one and re-upload). Do not guess.
      * more than one alt column matched -> return (df, [], error listing them).
      * alt_cols non-empty -> rename that column to the canonical string; append info message
        "Column '<original header stripped>' was read as '<canonical>'."
      * only canon_cols present but its original text differs from the canonical string (case/space/nbsp
        variant) -> rename it to the canonical string, no message needed.
    - Apply all renames together with df.rename(columns=...) after the loop.
    - Pairs must never affect each other.
    - Never mutate df.Columns in place in a loop.
    """
    if not column_aliases:
        return df, [], None

    # Build norm_cols: normalized header -> list of original column names
    norm_cols = {}
    for col in df.columns:
        norm = _normalize_header_text(col)
        norm_cols.setdefault(norm, []).append(col)

    info_messages = []
    error_msg = None
    rename_map = {}  # original_header_string -> canonical_string

    # Process each pair independently
    for canonical_key, alternates in column_aliases.items():
        norm_canonical = _normalize_header_text(canonical_key)

        # Columns matching the canonical (by normalized name)
        canon_cols = [c for c in df.columns if _normalize_header_text(c) == norm_canonical]

        # Columns matching any alternate (by normalized name)
        alt_cols = []
        for alt in alternates:
            norm_alt = _normalize_header_text(alt)
            for c in df.columns:
                if _normalize_header_text(c) == norm_alt and c not in alt_cols:
                    alt_cols.append(c)

        # Both canonical and alternate present in this pair -> error
        if canon_cols and alt_cols:
            # Use original header texts for the error message
            canon_orig = canon_cols[0]
            alt_orig = alt_cols[0]
            error_msg = (
                f"Column '{canon_orig}' and '{alt_orig}' both present. "
                "Please remove one and re-upload."
            )
            return df, [], error_msg

        # More than one alternate column matched -> error
        if len(alt_cols) > 1:
            alt_origs = ", ".join(alt_cols)
            error_msg = (
                f"Multiple alternate columns matched: {alt_origs}. "
                "Please remove all but one and re-upload."
            )
            return df, [], error_msg

        # Alternate column(s) present but no canonical -> rename first alt to canonical
        if alt_cols and not canon_cols:
            alt_orig = alt_cols[0]
            rename_map[alt_orig] = canonical_key
            info_messages.append(f"Column '{alt_orig}' was read as '{canonical_key}'.")

        # Only canonical columns present (possibly with case/space/nbsp variant);
        # if the original text differs from canonical string, rename to canonical (no message)
        # We handle this below by checking if any column's normalized name matches canonical
        # but its original text differs. Since we already checked both-present case above,
        # if we reach here with canon_cols and not alt_cols, we may need to rename if the
        # existing column name is a variant. However, the spec says: "only canon_cols present but
        # its original text differs from the canonical string ... -> rename it to the canonical string,
        # no message needed." We detect this by checking if the stored original column name(s)
        # are not exactly the canonical string (after normalization they match, but text differs).
        # We'll rename any canon_col whose stripped/normalized text != canonical_key text.
        if canon_cols and not alt_cols:
            for c in canon_cols:
                # If the original column name (as stored in the DataFrame) differs from the
                # canonical string (after our normalization), rename it.
                # We compare the original text to the canonical key.
                if c != canonical_key:
                    rename_map[c] = canonical_key
                    # No info message needed per spec for this case

    # Apply all renames together
    if rename_map:
        df = df.rename(columns=rename_map)

    return df, info_messages, error_msg


def render_generic_upload(
    table_name,
    column_mapping,
    date_columns=None,
    numeric_columns=None,
    consolidate_target=None,
    file_label="file",
    uploader_help="",
    column_aliases=None,
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

    # FIX 2: Use SHA-256 of file bytes as file identity everywhere
    file_hash = hashlib.sha256(uploaded_file.getvalue()).hexdigest()

    # Cache keys based on table name and file hash (FIX 2)
    cache_key = f"_parsed_df_{table_name}"
    file_hash_key = f"_parsed_file_hash_{table_name}"

    # FIX 2: Parse + preview + column validation (no DB access on rerun)
    # Only re-parse if file hash changed
    if st.session_state.get(file_hash_key) != file_hash:
        excel_bytes = uploaded_file.read()
        workbook = load_workbook(io.BytesIO(excel_bytes), data_only=True)
        sheet = workbook.active
        data = list(sheet.values)
        headers = [str(h).strip() if h else '' for h in data[0]]
        rows = data[1:]
        df = pd.DataFrame(rows, columns=headers)
        df.columns = df.columns.str.strip()

        # FIX 3: Apply column alias normalization BEFORE missing-columns/required-columns checks
        df, alias_infos, alias_error = _apply_column_aliases(df, column_aliases)

        # Store alias results in session_state next to the cached DataFrame
        # so they display on EVERY rerun, not only on the run that parsed the file
        st.session_state[cache_key] = df
        st.session_state[file_hash_key] = file_hash
        st.session_state[f"_alias_infos_{table_name}"] = alias_infos
        st.session_state[f"_alias_error_{table_name}"] = alias_error
    else:
        df = st.session_state[cache_key]

    # After the if/else, ALWAYS (cached or not): display error or info messages
    if st.session_state.get(f"_alias_error_{table_name}"):
        st.error(st.session_state[f"_alias_error_{table_name}"])
        return

    # Display stored info messages (one per pair that had an alias applied)
    for info in st.session_state.get(f"_alias_infos_{table_name}", []):
        st.info(info)

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

    # Truncate confirmation (existing logic, gates the load button)
    truncate_ok = st.checkbox(f"I have truncated table '{table_name}' before uploading", value=False)
    if not truncate_ok:
        st.warning(f"Please confirm that '{table_name}' has been truncated before uploading.")
        proceed = st.checkbox("Proceed without truncation confirmation", value=False)
        if not proceed:
            st.info("Check the box above to confirm truncation or proceed anyway.")
            return

    # FIX 2: Check if this file is already staged in the current session
    staged_key = f"_staged_{table_name}"
    staged_state = st.session_state.get(staged_key)
    file_hash_stored = staged_state.get("file_hash") if staged_state else None
    already_loaded = (file_hash_stored == file_hash) if file_hash_stored else False

    # FIX 2: "Load to staging table" button
    # Disabled if this file is already loaded in staging; truncation gate already satisfied
    # since we returned early if neither truncate_ok nor proceed was True
    load_button_disabled = already_loaded

    if already_loaded:
        st.info("This file is already loaded in staging. Upload a different file to load again.")

    if st.button("Load to staging table", disabled=load_button_disabled):
        try:
            conn = _get_connection()
            cursor = conn.cursor()
            conn.autocommit = False

            # Create table if not exists
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
                st.error(f"Upload failed — {len(failed_rows)} row(s) had errors. Rolling back all changes.")
                for idx, err in failed_rows:
                    st.write(f"  Row {idx}: {err}")
                conn.rollback()
            elif success_count > 0:
                conn.commit()

                # FIX 2: Store staged state in session
                st.session_state[staged_key] = {
                    "file_hash": file_hash,
                    "row_count": success_count,
                    "staged_at": datetime.now().isoformat(),
                }

                st.success(f"Uploaded {success_count}/{total_rows} rows successfully to '{table_name}'.")
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

    # FIX 1: Show consolidate section if staged state exists for current file
    staged_state = st.session_state.get(staged_key)
    current_file_hash_stored = staged_state.get("file_hash") if staged_state else None
    consolidated_key = f"_consolidated_{consolidate_target}_{table_name}"
    already_consolidated = st.session_state.get(consolidated_key) == file_hash

    if staged_state and current_file_hash_stored == file_hash:
        st.write(f"Staged rows: {staged_state.get('row_count', 0)}")

        if consolidate_target:
            if already_consolidated:
                st.info(f"Consolidated {staged_state.get('row_count', 0)} rows into '{consolidate_target}' already.")
                st.button(f"Consolidate to {consolidate_target}", disabled=True)
            else:
                st.write(f"Consolidate to '{consolidate_target}'?")

                if st.button(f"Consolidate to {consolidate_target}"):
                    # Use fresh connection inside button handler
                    try:
                        conn = _get_connection()
                        cursor = conn.cursor()
                        conn.autocommit = False

                        # Check staging row count concurrency guard
                        cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
                        actual_count = cursor.fetchone()[0]

                        expected_count = staged_state.get("row_count")
                        if actual_count != expected_count:
                            st.error(
                                f"Staging count mismatch: {actual_count} rows in table, "
                                f"expected {expected_count}. Refusing consolidation. "
                                "Someone may have modified the staging table."
                            )
                        else:
                            # Run consolidation INSERT INTO target SELECT FROM staging
                            cols = ", ".join(db_columns)
                            consolidate_query = f"""
                            INSERT INTO {consolidate_target} ({cols}) 
                            SELECT {cols} FROM {table_name}
                            """
                            cursor.execute(consolidate_query)

                            if cursor.rowcount == expected_count:
                                conn.commit()
                                st.success(f"Consolidated {expected_count} rows into '{consolidate_target}' successfully!")

                                # After success store consolidated_<table> = file_hash
                                st.session_state[consolidated_key] = file_hash

                                # Disable the button for that same file hash so a second click
                                # in the same session cannot insert again
                                st.session_state[f"_consolidated_disabled_{consolidate_target}_{file_hash}"] = True
                            else:
                                st.error(
                                    f"Consolidation failed: expected {expected_count} rows, "
                                    f"got {cursor.rowcount}. Rolling back..."
                                )
                                conn.rollback()

                    except Exception as e:
                        st.error(f"Consolidation failed: {e}")
                        if 'conn' in locals():
                            try:
                                conn.rollback()
                            except Exception:
                                pass
                    finally:
                        if 'conn' in locals():
                            conn.close()
