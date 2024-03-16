   
import pandas as pd
import os 
from google_access import get_sheet_data, get_pending_invoices, get_lab_db, upload_docx_todrive, mark_invoice_as_done, mark_invoice_as_error, get_all_invoices, upload_xlsx_todrive
from gen_invoice import create_docx_invoice
import random, string
import time
from adobe_client import AdobeClient
import json
import traceback
from config import *


def get_random_string(length):
    # choose from all lowercase letter
    letters = string.ascii_lowercase
    result_str = ''.join(random.choice(letters) for i in range(length))
    return result_str

def create_test_filenames(invoice_id):
    invoice_id = ''.join(e for e in invoice_id if e.isalnum())
    output_xlsx = f'#{invoice_id}_Invoice_Order_P2_{get_random_string(10)}.xlsx'
    output_docx = f'#{invoice_id}_Invoice_Order_P2_{get_random_string(10)}.docx'
    localoutput_docx = os.path.join('docs', output_docx)
    localoutput_xlsx = os.path.join('docs', output_xlsx)
    return output_xlsx, localoutput_docx, localoutput_xlsx

def remove_file(filepath):
    if os.path.exists(filepath):
        os.remove(filepath)
    else:
        print("The file does not exist")

def test():
    all_invoices = get_all_invoices()
    start_time = time.time()
    if len(all_invoices) == 0:
        print("No Invoices")
        return 0

    print("Invoices Detected: ", )
    lab_df = get_lab_db()
    lab_df.to_csv("lab_df.csv")
        
    for invoice in all_invoices:
        print("Invoice ID", invoice[INVOICE_ID_COL])
        # if invoice[INVOICE_ID_COL] == "GSAE006":
        try:
            if invoice[INVOICE_ID_COL] == "TFRC002":
                # test_details = invoice[TEST_DETAILS_COL]
                invoice_id = invoice[INVOICE_ID_COL]
                # print(f"{invoice_id} : Invoice: {invoice}")
                output_xlsx, localoutput_docx, localoutput_xlsx = create_test_filenames(invoice_id)
                # print("Test Details", test_details)
                # print("Local Filepath", localoutput_docx, localoutput_xlsx)
                # print(json.dumps(invoice, indent=4))
                create_docx_invoice(invoice, localoutput_docx, lab_df)
                # print(f"{invoice_id} : Test Details: {test_details}")
                # print(f"{invoice_id} : Created Docx")
                # client = AdobeClient()
                # client.create_xlsx(localoutput_docx, localoutput_xlsx)
                # upload_docx_todrive(localfilepath=localoutput_xlsx, filename=output_xlsx)
                # print(f"{invoice_id} : Uploaded Invoice:")
                # mark_invoice_as_done(invoice_id)
                # print(f"{invoice_id} : Marked as Done")
                # remove_file(localoutput_filepath)
                # print(f"{invoice_id} : Remove File")
                # print(f"{invoice_id} : Completed processing:")
        except Exception as e:
            invoice_id = invoice[INVOICE_ID_COL]
            error_trace = traceback.format_exc()
            print(f"Process Error for {invoice} : {error_trace}")
            print(traceback.format_exc())
            mark_invoice_as_error(invoice_id, error=error_trace)
            return 0    
        #     break
        # except Exception as e:
        #     print(f"Process Error for {invoice[INVOICE_ID_COL]} : {e}")
    # time.sleep(FREQUENCY_SCRIPT_SEC)

test()