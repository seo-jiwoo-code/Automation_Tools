import requests
import time
import urllib.request


# CLIENT_ID = '3295644b6bdc4312a19629c0bc0a452d'
# CLIENT_SECRET = 'p8e-2jdS60yCn2CRlFWM1TElIUHfrHvQMi8T'

CLIENT_ID = "f9f732ec8fe44663b4a06c8c303bc0b2"
CLIENT_SECRET = "p8e-7koYYGW8Vg1s1uWbb6s9QonA_YyDcJLJ"
MIMETYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

class AdobeClient:
    def __init__(self):
        self.cred = {
            'client_id': CLIENT_ID,
            'client_secret': CLIENT_SECRET,
        }
        self.access_token = self.get_access_token()
        self.json_header = {
            'X-API-Key': CLIENT_ID,
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json',
        }
        self.basic_header = {
            'X-API-Key': CLIENT_ID,
            'Authorization': f'Bearer {self.access_token}',
        }

    def remove_asset(self, asset_id):
        response = requests.delete(f'https://pdf-services.adobe.io/assets/{asset_id}', headers=self.basic_header)
        # print(response.json())


    def get_access_token(self):
        response = requests.post('https://pdf-services.adobe.io/token', data=self.cred)
        access_token = response.json()['access_token']
        return access_token

    def create_xlsx(self, input_docx_filepath, output_xlsx_filepath):
        docx_asset_id = self.upload_docx_todrive(input_docx_filepath)
        pdf_asset_id = self.docx_to_pdf(docx_asset_id)
        xlsx_asset_id = self.pdf_to_xlsx(pdf_asset_id, output_xlsx_filepath)
        self.remove_asset(docx_asset_id)
        self.remove_asset(pdf_asset_id)
        self.remove_asset(xlsx_asset_id)

    
    def upload_docx_todrive(self, filepath):
        json_data = {
            'mediaType': MIMETYPE,
        }
        response = requests.post('https://pdf-services.adobe.io/assets', headers=self.json_header, json=json_data)
        upload_uri = response.json()['uploadUri']
        asset_id = response.json()['assetID']
        # print(response.json())

        with open(filepath, 'rb') as f:
            data = f.read()

        response = requests.put(
            upload_uri,
            headers={
                'Content-Type': MIMETYPE,
            },
            data=data,
        )

        return asset_id

    def docx_to_pdf(self, asset_id):
        
        json_data = {
            'assetID': asset_id,
            "documentLanguage": "en-US"
        }

        response = requests.post('https://pdf-services.adobe.io/operation/createpdf', headers=self.json_header, json=json_data)
        job_status_uri = response.headers["location"]

        while True:
            # print("JOB Status URI:", job_status_uri)
            response = requests.get(job_status_uri, headers=self.json_header)
            # print(response.json())
            job_status = response.json()["status"]
            # print("Job Status:", job_status)
            if job_status == "done":
                asset = response.json()["asset"]
                # print("Output Asset:", asset)
                pdf_asset_id = asset["assetID"]
                break
            time.sleep(5)
        
        urllib.request.urlretrieve(asset["downloadUri"], "output24.pdf")

        return pdf_asset_id

    def pdf_to_xlsx(self, asset_id, output_xlsx_filepath):
        json_data = {
            "assetID": asset_id,
            "targetFormat": "xlsx",
            "ocrLang": "en-US"
        }

        response = requests.post('https://pdf-services.adobe.io/operation/exportpdf', headers=self.json_header, json=json_data)
        job_status_uri = response.headers["location"]

        while True:
            # print("JOB Status URI:", job_status_uri)
            response = requests.get(job_status_uri, headers=self.json_header)
            # print(response.json())
            job_status = response.json()["status"]
            # print("Job Status:", job_status)
            if job_status == "done":
                asset = response.json()["asset"]
                # print("Output XLSX Asset:", asset)
                asset_id = asset["assetID"]
                break
            time.sleep(5)

        urllib.request.urlretrieve(asset["downloadUri"], output_xlsx_filepath)
        urllib.request.urlretrieve(asset["downloadUri"], "output24.xlsx")

        return asset_id

if __name__ == "__main__":
    FILEPATH = "docs/JGA001_kkccnkenwf.docx"
    client = AdobeClient()
    client.create_xlsx(FILEPATH, "output2.xlsx")
