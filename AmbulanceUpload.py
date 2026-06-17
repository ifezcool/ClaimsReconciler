from db_upload_common import render_generic_upload

COLUMN_MAPPING = {
    'S_N': 'S_N', 'CLAIM_TYPE': 'CLAIM_TYPE', 'BATCH_NUMBER': 'BATCH_NUMBER',
    'HOSPITAL': 'HOSPITAL', 'NUMBER_OF_CLAIMS': 'NUMBER_OF_CLAIMS',
    'ENCOUNTER_MONTH': 'ENCOUNTER_MONTH', 'DATE_OF_RECEIPT': 'DATE_OF_RECEIPT',
    'APPROVED_PA_VALUE_N': 'APPROVED_PA_VALUE_N',
    'AMOUNT_RECOMMENDED_FOR_PAYMENT_N': 'AMOUNT_RECOMMENDED_FOR_PAYMENT_N',
    'VARIANCE': 'VARIANCE', 'VARIANCE1': 'VARIANCE1', 'NARRATION': 'NARRATION',
    'Source_File': 'Source_File', 'PROVIDER_CODE': 'PROVIDER_CODE',
    'Paiddate': 'Paiddate', 'SCH_NO': 'SCH_NO', 'APPEAL_NO': 'APPEAL_NO', 'SCH_NUM': 'SCH_NUM',
}

def render_ambulance_upload():
    render_generic_upload(
        table_name='ambulancetbl',
        column_mapping=COLUMN_MAPPING,
        date_columns=['DATE_OF_RECEIPT', 'Paiddate'],
        numeric_columns={
            'APPROVED_PA_VALUE_N': 'DECIMAL(18,2)',
            'AMOUNT_RECOMMENDED_FOR_PAYMENT_N': 'DECIMAL(18,2)',
            'VARIANCE': 'DECIMAL(18,2)',
            'VARIANCE1': 'DECIMAL(18,2)',
            'NUMBER_OF_CLAIMS': 'INT',
        },
        consolidate_target='Compiled_ambulance',
        file_label="Ambulance (Please truncate ambulancetbl before uploading)",
    )
