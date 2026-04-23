import pandas as pd
import datetime
import holidays  # jeśli chcesz uwzględnić święta
import traceback
import logging

from helper_functions import (
    get_supplier_sap_numbers,
    get_zek103_data,
    update_excel_with_quantities,
    list_excel_files,
    get_mb52_data,
    filter_mb52_data,
    update_excel_with_dataframe,
    get_past_workdays,
    filter_z_mat_consumption_data,
    extract_unique_mb51_files,
    load_grouped_mb51_dataframes,
    retrieve_supplier_name,
    retrieve_plant
)


zek103_dtypes = {
    'Werk': 'string',
    'Mat': 'string',
}

mb52_dtypes = {
    'Material': 'string',
    'Lagerort': 'string',
    'Werk': 'string',
}

mb51_consumption_dtypes = {
    'Material': 'string',
    'Werk': 'string',
}

MB51_new_col_names = {
    'Material': 'Material',
    'Buchungsdatum': 'date',
    'Menge': 'quantity',
    'Basis-ME': 'unit',
    'Werk': 'plant'
}

supplier_files_dict = {
    "SHC": ["CONS_246", "CONS_246_2"],
    "HESSE": ["CONS_246", "CONS_246_2"],
    "ROTO_ELZETT": ["CONS_246", "CONS_246_2"],
    "TOKOZ": ["CONS_246", "CONS_246_2"],
    "HOHAGE": ["CONS_246", "CONS_246_2"],
    "KROSNO": ["CONS_246", "CONS_246_2"],
    "FRANZEN": ["CONS_246", "CONS_246_2"],
    "ROZTOCZE": ["CONS_246", "CONS_246_2"],
    "BELATRONIC": ["CONS_246", "CONS_246_2"],
    "DEVENTER": ["CONS_246", "CONS_246_2"],
    "NMC": ["CONS_246", "CONS_246_2"],
    "NEHER": ["CONS_223", "CONS_223_2"],
    "EJOT": ["CONS_224", "CONS_224_2"],
    "SPAX": ["CONS_224", "CONS_224_2"],
    "PROFINE": ["CONS_224", "CONS_224_2"],
    "STOROPACK": ["CONS_234", "CONS_234_2"],
    "YUYAO_JILO": ["CONS_234", "CONS_234_2"],
    "BRETTHAUER": ["CONS_234", "CONS_234_2"],
    "LEO_FRANCOIS": ["CONS_234", "CONS_234_2"],
    "GALLARDO": ["CONS_234", "CONS_234_2"],
    "REISSER": ["CONS_234", "CONS_234_2"],
    "ABC_COLORE": ["CONS_228", "CONS_228_2"],
    "BACCARAT": ["CONS_228", "CONS_228_2"],
    "DOMICET": ["CONS_239", "CONS_239_2"],
    "BIZEA": ["CONS_239", "CONS_239_2"],
    "DAFA": ["CONS_239", "CONS_239_2"],
    "RETECH": ["CONS_239", "CONS_239_2"],
    "KABEL_MAIER": ["CONS_239", "CONS_239_2"],
    "GRUNEFELD": ["CONS_239", "CONS_239_2"],
    "HAFELE": ["CONS_234", "CONS_234_2"],
    "BECKER": ["CONS_337", "CONS_337_2"],
    "INTERNATIO": ["CONS_337", "CONS_337_2"],
    "RHEINZINK": ["CONS_337", "CONS_337_2"],
    "SCHOLLGLAS": ["CONS_337", "CONS_337_2"],
    "WALTRON": ["CONS_337", "CONS_337_2"],
    "UNION": ["CONS_234", "CONS_234_2"],
}


zek103_file_path = r"\\rfmesrv5\connect\DST_SAP_Transfer\P11\PPS_LUB\05_PURCHASING_AUTOMATION\ZEK103_PUR_LUB_002.xlsx"
mb52_file_path = r"\\rfmesrv5\connect\DST_SAP_Transfer\P11\PPS_LUB\05_PURCHASING_AUTOMATION\MB52_PUR_LUB_003.xlsx"
z_mat_consumption_filepath = r"\\rfmesrv5\connect\DST_SAP_Transfer\P11\PPS_LUB\05_PURCHASING_AUTOMATION\Z_MAT_CONSUMPTION.xlsx"
mb51_consumption_filepath = r"\\rfmesrv5\connect\DST_SAP_Transfer\P11\PPS_LUB\05_PURCHASING_AUTOMATION\MB51_usage.XLSX"

main_excel_path = r'P:\Technisch\PLANY PRODUKCJI\PLANIŚCI\PP_TOOLS_TEMP_FILES\12_PURCHASE_AUTOMATION\PurchAutomationTemplate.xlsm'
supplier_files_directory_path_test = r'P:\Technisch\PLANY PRODUKCJI\PLANIŚCI\PP_TOOLS_TEMP_FILES\12_PURCHASE_AUTOMATION\supplier_files'
supplier_files_directory_path = r'P:\Zakupy\O\SupplierAutomation\supplier_files'
export_files_directory_path = r"\\rfmesrv5\connect\DST_SAP_Transfer\P11\PPS_LUB\05_PURCHASING_AUTOMATION"

ERROR_LOG_PATH = r"P:\Zakupy\O\SupplierAutomation\error.log"

# Updating the open orders data

movement_types = ('261', '313')

logging.basicConfig(
    filename=ERROR_LOG_PATH,
    level=logging.ERROR,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

try:
    zek103_content = pd.read_excel(zek103_file_path, dtype=zek103_dtypes)
    mb52df = get_mb52_data(mb52_file_path, mb52_dtypes)

    # DONE: Get dfs for all exported files
    unique_mb51_files_dict = extract_unique_mb51_files(supplier_files_dict)
    MB551_dfs = load_grouped_mb51_dataframes(
        grouped_files=unique_mb51_files_dict,
        base_path=export_files_directory_path,
        extension=".xlsx",
        dtypes_mb51=mb51_consumption_dtypes,
        col_names_mb51=MB51_new_col_names
    )


    # TODO: Change path
    supplier_files_paths = list_excel_files(supplier_files_directory_path)

    storage_locs = ('0003', '0004', '0005', '0007', '0710')

    current_year = datetime.date.today().year
    years_list = [current_year - 1, current_year, current_year + 1]
    pl_holidays = holidays.Poland(years=years_list)

    today = datetime.datetime.today().date()
    year_ago = today - datetime.timedelta(days=365)

    past_20_days_now = get_past_workdays(today, 20, 'last', pl_holidays)
    past_20_days_year_ago = get_past_workdays(year_ago, 20, 'last', pl_holidays)
    next_20_days_year_ago = get_past_workdays(year_ago, 20, 'next', pl_holidays)
    past_60_days_now = get_past_workdays(today, 60, 'last', pl_holidays)
    past_60_days_year_ago = get_past_workdays(year_ago, 60, 'last', pl_holidays)

    usage_parameters = [(next_20_days_year_ago, 'C'),
                        (past_20_days_now, 'D'),
                        (past_20_days_year_ago, 'E'),
                        (past_60_days_now, 'F'),
                        (past_60_days_year_ago, 'G'),
    ]

    for file_path in supplier_files_paths:
        plant = retrieve_plant(file_path)
        excel_sheet = f"data"

        # DONE: Retrieve supplier name from file path and get corresponding df
        supplier_name = retrieve_supplier_name(file_path)
        z_mat_or_mb51_consumption_df = MB551_dfs[supplier_files_dict[supplier_name][0]]

        sap_list = get_supplier_sap_numbers(file_path, excel_sheet)
        zek103_output = get_zek103_data(zek103_content, plant, sap_list)
        update_excel_with_quantities(file_path, zek103_output, 'SAP', 'CONSUMPTION', excel_sheet, True)

        mb52_stock = filter_mb52_data(mb52df, sap_list, plant, storage_locs, 'Frei verwendbar')
        mb52_pqm = filter_mb52_data(mb52df, sap_list, plant, storage_locs, 'In QualPrüfung')
        # Wywołanie funkcji
        update_excel_with_dataframe(
            file_path=file_path,
            dataframe=mb52_stock,
            sap_column='A',  # Kolumna z numerami SAP w Excelu
            frei_column='I',  # Kolumna, do której wpisywane są dane
            header_start='Stock',  # Nagłówek wskazujący początek
            header_end='S.C',       # Nagłówek wskazujący koniec
            sheet_name=excel_sheet,
        )

        update_excel_with_dataframe(
            file_path=file_path,
            dataframe=mb52_pqm,
            sap_column='A',  # Kolumna z numerami SAP w Excelu
            frei_column='J',  # Kolumna, do której wpisywane są dane
            header_start='Stock',  # Nagłówek wskazujący początek
            header_end='S.C',       # Nagłówek wskazujący koniec
            sheet_name=excel_sheet,
            col_name= 'In QualPrüfung'
        )

        for parameter in usage_parameters:
            z_mat_consumption_grouped = filter_z_mat_consumption_data(z_mat_or_mb51_consumption_df, sap_list, parameter[0], plant)

            # wpisanie danych do Excela
            update_excel_with_dataframe(
                file_path=file_path,
                dataframe=z_mat_consumption_grouped,
                sap_column='A',  # Kolumna z numerami SAP w Excelu
                frei_column=parameter[1],  # Kolumna, do której wpisywane są dane
                header_start='Consumption',  # Nagłówek wskazujący początek
                header_end='Stock',       # Nagłówek wskazujący koniec
                sheet_name=excel_sheet,
                col_name='quantity'
            )

except Exception as e:
    logging.error("Error occurred", exc_info=True)
    error_details = traceback.format_exc()
    print(f"Wystąpił błąd:\n{error_details}")
    input("Press Enter to continue...")