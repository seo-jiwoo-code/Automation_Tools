from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import pandas as pd

import os
from config import *

def get_sheet_data(sheet_name, sheet_id=SHEET_ID):
    gc = gspread.service_account(CREDS_JSON_PATH)
    sheet_samples = gc.open_by_key(sheet_id).worksheet(sheet_name)
    return sheet_samples

def save_to_drive(folder_id, filename, filepath, mimetype):
    # Authenticate using the service account
    scopes = ['https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDS_JSON_PATH, scopes)
    drive_service = build('drive', 'v3', credentials=creds)

    # Define the file metadata, like the name and MIME type
    file_metadata = {
        'name': filename,
        'mimeType': mimetype,
        'parents': [folder_id]
    }

    # Specify the file to upload
    media = MediaFileUpload(filepath, mimetype=mimetype)

    # Create the file on Google Drive
    file = drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id'
    ).execute()

    print(f'File ID: {file.get("id")}')

def upload_docx_todrive(localfilepath, filename):
    folder_id = DRIVE_FOLDER_ID
    # file_metadata = {
    #     'name': filename,  # Your desired filename here
    #     'parents': [folder_id]
    # }  
    save_to_drive(DRIVE_FOLDER_ID, filename, localfilepath, mimetype=WORD_DOCX_MIMETYPE)
    # drive_service = build('drive', 'v3', credentials=get_credentials()) # Assuming you've defined get_credentials() earlier
    # file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    # print(f"File uploaded with ID: {file.get('id')}")

def upload_xlsx_todrive(localfilepath, filename):
    folder_id = DRIVE_FOLDER_ID
    save_to_drive(DRIVE_FOLDER_ID, filename, localfilepath, mimetype=XLSX_MIMETYPE)

def get_lab_db():
    rows = get_sheet_data(LAB_DB_SHEET, LAB_DB_FILE_ID).get_all_records(head=3)
    lab_df = pd.DataFrame(rows)
    # print(lab_df.columns)
    # lab_df.to_csv("lab_df.csv")
    # print(lab_df['Price'])
    for numeric_col in [PKG_PRICE_COL, BASE_PRICE_COL, PKG_COST_COL, BASE_COST_COL]:
        lab_df[numeric_col] = pd.to_numeric(lab_df[numeric_col], errors = 'coerce')

    lab_df = lab_df[lab_df[TEST_PARAM_COL] != ""]
    return lab_df

def mark_invoice_as_done(invoice_id):
    invoice_sheet = get_sheet_data(INVOICE_TRACKER_SHEET)
    invoice_records = invoice_sheet.get_all_records()
    # Replace 'status_column_number' with the actual column number of 'Status' in your Google Sheets
    i = 2
    for invoice in invoice_records:
        if invoice[INVOICE_ID_COL] == invoice_id:
            invoice_sheet.update_cell(i, INVOICE_STATUS_SHEET_COLNUMBER, 'Done')
        i += 1

def mark_invoice_as_error(invoice_id, error=""):
    invoice_sheet = get_sheet_data(INVOICE_TRACKER_SHEET)
    invoice_records = invoice_sheet.get_all_records()
    # Replace 'status_column_number' with the actual column number of 'Status' in your Google Sheets
    i = 2
    for invoice in invoice_records:
        if invoice[INVOICE_ID_COL] == invoice_id:
            invoice_sheet.update_cell(i, INVOICE_STATUS_SHEET_COLNUMBER, 'Error')
            invoice_sheet.update_cell(i, INVOICE_SHEET_ERROR_DESC_COLNUMBER, error)
            
        i += 1

def get_pending_invoices():
    invoice_records = get_sheet_data(INVOICE_TRACKER_SHEET).get_all_records() 
    sample_df = pd.DataFrame(get_sheet_data(SAMPLES_DATA_SHEET).get_all_records())
    # shelf_life_data = get_sheet_data("Shelf life").get_all_values()
    # shelf_life_df = pd.DataFrame(shelf_life_data[1:], columns = shelf_life_data[1])
    shelf_life_df = pd.DataFrame(get_sheet_data("Shelf life").get_all_records())

    # other_cost_data = get_sheet_data("Shelf life").get_all_values()
    # other_cost_df = pd.DataFrame(other_cost_data[1:], columns = other_cost_data[1])
    other_cost_df = pd.DataFrame(get_sheet_data("Other cost details").get_all_records())

    pending_invoices = []
    for invoice in invoice_records:
        invoice_id = invoice[INVOICE_ID_COL]
        if invoice["Status"] == "Pending" and invoice[INVOICE_ID_COL]:
            test_param_details = sample_df[sample_df[INVOICE_ID_COL] == invoice_id]
            test_details = []
            for idx, row in test_param_details.iterrows():
                raw_df = sample_df[sample_df[INVOICE_ID_COL] == invoice_id]
                sample_dict = {
                    key : row[key] for key in [
                        TEST_PARAM_COL, 
                        LAB_ID_COL, 
                        LAB_TEST_CATEGORY_COL, 
                        METHOD_REFERENCE_CODE_COL, 
                        COST_COL, 
                        NUMBER_OF_SAMPLES_COL,
                        SAMPLE_LABELS_COL,
                        SAMPLE_DESCRIPTION_COL
                    ]
                }
                for key in sample_dict:
                    if type(sample_dict[key]) == str:
                        sample_dict[key] = sample_dict[key].strip()
                # sample_dict[SAMPLE_LABELS_COL] = row[SAMPLE_LABELS_COL].split(",")
                # sample_dict[SAMPLE_DESCRIPTION_COL] = row[SAMPLE_DESCRIPTION_COL]
                test_details += [sample_dict]

            shelf_life_details = []
            for shelf_life_data in shelf_life_df.columns:
                if invoice_id in shelf_life_data:
                    shelf_life_dict = {}
                    for col in shelf_life_df["PARAMETERS"]:
                        shelf_life_dict[col] = shelf_life_df[shelf_life_df["PARAMETERS"] == col][shelf_life_data].values[0]
                    shelf_life_details += [shelf_life_dict]  

            other_cost_details = {
                "Discount" : ""
            }
            for other_cost_data in other_cost_df.columns:
                if invoice_id == other_cost_data:
                    other_cost_dict = {}
                    for col in other_cost_df["Other Costs"]:
                        other_cost_dict[col] = other_cost_df[other_cost_df["Other Costs"] == col][other_cost_data].values[0]
                    other_cost_details = other_cost_dict
                    break


            pending_invoices += [{
                INVOICE_ID_COL : invoice_id,
                "Samples" : test_details,
                "Shelf-Life" : shelf_life_details,
                "Other Costs" : other_cost_details
                # TEST_DETAILS_COL : test_details
            }]
    return pending_invoices

def get_all_invoices():
    invoice_records = get_sheet_data(INVOICE_TRACKER_SHEET).get_all_records() 
    sample_df = pd.DataFrame(get_sheet_data(SAMPLES_DATA_SHEET).get_all_records())
    # shelf_life_data = get_sheet_data("Shelf life").get_all_values()
    # shelf_life_df = pd.DataFrame(shelf_life_data[1:], columns = shelf_life_data[0])
    shelf_life_df = pd.DataFrame(get_sheet_data("Shelf life").get_all_records())

    # other_cost_data = get_sheet_data("Shelf life").get_all_values()
    # other_cost_df = pd.DataFrame(other_cost_data[1:], columns = other_cost_data[0])
    other_cost_df = pd.DataFrame(get_sheet_data("Other cost details").get_all_records())

    pending_invoices = []
    for invoice in invoice_records:
        invoice_id = invoice[INVOICE_ID_COL]

        test_param_details = sample_df[sample_df[INVOICE_ID_COL] == invoice_id]
        test_details = []
        for idx, row in test_param_details.iterrows():
            raw_df = sample_df[sample_df[INVOICE_ID_COL] == invoice_id]
            sample_dict = {
                key : row[key] for key in [
                    TEST_PARAM_COL, 
                    LAB_ID_COL, 
                    LAB_TEST_CATEGORY_COL, 
                    METHOD_REFERENCE_CODE_COL, 
                    COST_COL, 
                    NUMBER_OF_SAMPLES_COL,
                    SAMPLE_LABELS_COL,
                    SAMPLE_DESCRIPTION_COL
                ]
            }
            for key in sample_dict:
                if type(sample_dict[key]) == str:
                    sample_dict[key] = sample_dict[key].strip()
            # sample_dict[SAMPLE_LABELS_COL] = row[SAMPLE_LABELS_COL].split(",")
            # sample_dict[SAMPLE_DESCRIPTION_COL] = row[SAMPLE_DESCRIPTION_COL]
            test_details += [sample_dict]

        shelf_life_details = []
        for shelf_life_data in shelf_life_df.columns:
            print(invoice_id, shelf_life_data)
            if invoice_id in shelf_life_data:
                shelf_life_dict = {}
                for col in shelf_life_df["PARAMETERS"]:
                    shelf_life_dict[col] = shelf_life_df[shelf_life_df["PARAMETERS"] == col][shelf_life_data].values[0]
                shelf_life_details += [shelf_life_dict]  

        other_cost_details = {
            "Discount" : ""
        }
        for other_cost_data in other_cost_df.columns:
            if invoice_id == other_cost_data:
                other_cost_dict = {}
                for col in other_cost_df["Other Costs"]:
                    other_cost_dict[col] = other_cost_df[other_cost_df["Other Costs"] == col][other_cost_data].values[0]
                other_cost_details = other_cost_dict
                break


        pending_invoices += [{
            INVOICE_ID_COL : invoice_id,
            "Samples" : test_details,
            "Shelf-Life" : shelf_life_details,
            "Other Costs" : other_cost_details
            # TEST_DETAILS_COL : test_details
        }]
    print(pending_invoices)

    return pending_invoices


if __name__ == "__main__":
    import gen_invoice
    import json
    invoices = get_all_invoices()
    labs_df = get_lab_db()
    print(labs_df)
    for invoice in invoices:
        if invoice[INVOICE_ID_COL] == "GSAE006":
            print("-----------------------")
            print(json.dumps(invoice, indent=4))  
            # test_details = invoice[TEST_DETAILS_COL]
            # invoice_id = invoice[INVOICE_ID_COL]
            std_test_details_json = gen_invoice.create_std_test_details_json(invoice, labs_df)
            # print(json.dumps(std_test_details_json, indent=4))
            print("*************************")
            std_test_charges_json = gen_invoice.create_std_test_charges_json(invoice, labs_df)
            # print(json.dumps(std_test_charges_json, indent=4))
            shelf_life_details_json = gen_invoice.create_shelf_life_test_details_json(invoice)
            shelf_life_charges_json = gen_invoice.create_shelf_life_test_charges_json(invoice)
            total_lab_test_charges_json = gen_invoice.create_total_lab_test_charges_json(shelf_life_charges_json, std_test_charges_json)
            food_service_charges_json = gen_invoice.create_food_service_test_charges_json(invoice)
            total_cost_json = gen_invoice.create_total_cost_json(total_lab_test_charges_json, food_service_charges_json)
            
        # except Exception as e:
        #     print("Error in Invoice: ", invoice[INVOICE_ID_COL], e)
        #     continue


            invoice_json = {
                "Standard Test Details" : std_test_details_json,
                "Shelf-Life Details" : shelf_life_details_json,
                "Standard Test Charges" : std_test_charges_json,
                "Shelf-Life Charges" : shelf_life_charges_json,
                "Total Lab Test Charges" : total_lab_test_charges_json,
                "Total Food Service Charges" : food_service_charges_json,
                "Total Cost" : total_cost_json  
            }
            print(json.dumps(invoice_json, indent=4))
    # print(invoices)
    # get_lab_db()

