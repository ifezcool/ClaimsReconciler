The Fix — Two Places Per Module
Fix 1: Compilation files (appeals_page.py, telemedicine.py, ambulance.py)
Inside the compile_*_files function, locate the column_mapping loop where data is copied from original columns into standardized_df. It currently looks like this conceptually:
if new_col in ['S_N']:
    # numeric handling
else:
    standardized_df[new_col] = df_clean[original_col].astype(str)
Expand the condition that handles numeric-looking columns to also cover BATCH_NUMBER and PROVIDER_CODE. The logic for those two columns should be: if the value is not null and the string representation (after stripping .0) is all digits, store it as a plain integer string (no decimal). Otherwise fall back to a regular string conversion. This is the same pattern already used for S_N — just extend the condition to include BATCH_NUMBER and PROVIDER_CODE in the same branch.
Apply this change in all three compilation files: appeals_page.py, telemedicine.py, and ambulance.py.
Fix 2: Upload files (AppealsUpload.py, TelemedicineUpload.py, AmbulanceUpload.py)
Inside the clean_value function, before the final return str(val).strip() line, add a defensive check specifically for BATCH_NUMBER and PROVIDER_CODE: if the column is one of those two, attempt to convert the value to float then to int then to string, so that "12345.0" becomes "12345". If that conversion fails for any reason (e.g. the value is genuinely non-numeric like a code with letters), fall through to the regular str(val).strip() return. This acts as a safety net in case any decimal values slip through from the compiled file.
Apply this change in all three upload files: AppealsUpload.py, TelemedicineUpload.py, and AmbulanceUpload.py.