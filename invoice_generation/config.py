import gspread

FREQUENCY_SCRIPT_SEC = 10

CREDS_JSON_PATH = '' # replace this with your own service account's json. 

def get_config_sheet_id(config_sheet_id, config_sheet_name, cell_address):
    gc = gspread.service_account(CREDS_JSON_PATH)
    config_sheet = gc.open_by_key(config_sheet_id).worksheet(config_sheet_name)
    return config_sheet.acell(cell_address).value

CONFIG_SHEET_ID = '' #added a configuration sheet where it can access the values of sheet ID. 
CONFIG_SHEET_NAME = 'Var'


# LAB_DB_FILE_ID = ""
# LAB_DB_FILE_ID = "" # Updated when changed to ops3@co-lab.cc
LAB_DB_FILE_ID = get_config_sheet_id(CONFIG_SHEET_ID, CONFIG_SHEET_NAME, 'B1')
LAB_DB_SHEET = 'Labs Database'


# SHEET_ID = '' #second part of the google sheets URL
SHEET_ID = "" 
INVOICE_TRACKER_SHEET = "Invoice_Tracker"
SAMPLES_DATA_SHEET = "Samples"

INVOICE_STATUS_SHEET_COLNUMBER = 2
INVOICE_SHEET_ERROR_DESC_COLNUMBER = 3

DRIVE_FOLDER_ID = '' 
WORD_DOCX_MIMETYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
XLSX_MIMETYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

INVOICE_ID_COL = "Invoice ID"
SAMPLE_LABELS_COL = "Sample Labels"
SAMPLE_DESCRIPTION_COL = "Sample Description"
LAB_ID_COL = "Lab ID"   
LAB_TEST_CATEGORY_COL = "Lab Test Category"
MRC_COL = "MRC"
METHOD_REFERENCE_CODE_COL = "Method Reference Code (MRC)"
METHOD_REFERENCE_COL = "Method Reference"
NUMBER_OF_SAMPLES_COL = "Number of Samples"
UNITS_OF_MEASUREMENT_COL = "Units of Measurement"
LOD_COL = "LOD (ppm)"

TEST_PARAM_COL = "Test Param [External]"
TEST_DETAILS_COL = "Test Details"
TEST_PARAM_DETAILS_COL = "Test Param Details"
TEST_PARAMS_COL = "Test Parameters"

COST_COL = "Cost"
PRICE_COL = "Price"
PKG_COST_COL = "Package Cost"
PKG_PRICE_COL = "Package Price"
BASE_PRICE_COL = "Base Price"
BASE_COST_COL = "Base Cost"