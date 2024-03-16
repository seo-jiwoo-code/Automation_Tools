
import pandas as pd
import os 
from google_access import get_sheet_data, get_pending_invoices, get_lab_db, upload_docx_todrive, mark_invoice_as_done, mark_invoice_as_error, get_all_invoices, upload_xlsx_todrive
from gen_invoice import create_docx_invoice
import random, string
import time
from adobe_client import AdobeClient
import traceback

from config import *

def get_random_string(length):
    # choose from all lowercase letter
    letters = string.ascii_lowercase
    result_str = ''.join(random.choice(letters) for i in range(length))
    return result_str

def create_filenames(invoice_id):
    invoice_id = ''.join(e for e in invoice_id if e.isalnum())
    #XXXXX Invoice Order P2_jkweclisejbv
    output_xlsx = f'#{invoice_id}_Invoice_Order_P2_{get_random_string(10)}.xlsx'
    output_docx = f'#{invoice_id}_Invoice_Order_P2_{get_random_string(10)}.docx'
    localoutput_docx = os.path.join('/tmp', output_docx)
    localoutput_xlsx = os.path.join('/tmp', output_xlsx)
    return output_docx, localoutput_docx, localoutput_xlsx

def remove_file(filepath):
    if os.path.exists(filepath):
        os.remove(filepath)
    else:
        print("The file does not exist")

def main():
    try:
        pending_invoices = get_pending_invoices()
        if len(pending_invoices) == 0:
            print("No Pending Invoices")
            time.sleep(FREQUENCY_SCRIPT_SEC)
            return 0
    except Exception as e:
        print("Error in getting pending invoices", e)
        time.sleep(FREQUENCY_SCRIPT_SEC)
        return 0
    
    try:
        print("Getting Lab DB Data: ", )
        lab_df = get_lab_db()
    except Exception as e:
        print("Error in getting lab db data", e)
        time.sleep(FREQUENCY_SCRIPT_SEC)
        return 0
    

    for invoice in pending_invoices:
        try:
            invoice_id = invoice[INVOICE_ID_COL]
            output_docx, localoutput_docx, localoutput_xlsx = create_filenames(invoice_id)
            create_docx_invoice(invoice, localoutput_docx, lab_df)
            print(f"{invoice_id} : Invoice Details: {invoice}")
            print(f"{invoice_id} : Created Docx")
            # client = AdobeClient()
            # client.create_xlsx(localoutput_docx, localoutput_xlsx)
            upload_docx_todrive(localfilepath=localoutput_docx, filename=output_docx)
            # upload_xlsx_todrive(localfilepath=localoutput_xlsx, filename=output_xlsx)
            print(f"{invoice_id} : Uploaded Invoice:")
            mark_invoice_as_done(invoice_id)
            print(f"{invoice_id} : Marked as Done")
            # remove_file(localoutput_xlsx)
            remove_file(localoutput_docx)
            print(f"{invoice_id} : Remove File")
            print(f"{invoice_id} : Completed processing:")
        except Exception as e:
            invoice_id = invoice[INVOICE_ID_COL]
            error_trace = traceback.format_exc()
            print(f"Process Error for {invoice} : {error_trace}")
            # print(traceback.format_exc())
            mark_invoice_as_error(invoice_id, error=error_trace)
            return 0
        # except Exception as e:
        #     print(f"Process Error for {invoice} : {")
        

while True:
    start_time = time.time()
    # while time. time() - start_time < 60*10 - FREQUENCY_SCRIPT_SEC:
    while True:
        try:
            print("Starting main")
            main()
            time.sleep(FREQUENCY_SCRIPT_SEC)
        except:
            print("Error in main")
            time.sleep(FREQUENCY_SCRIPT_SEC)
            continue



