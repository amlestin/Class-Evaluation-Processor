# Python 3

import requests
import zipfile
import json
import io


#                              #
#                              #
#                              #
# FILL IN API CREDENTIALS HERE #
#                              #
#                              #

api_token = ""
data_center = ""



def download_survey(api_token, survey_id, data_center):

  file_format = "csv"

  # Setting static parameters
  request_check_progress = 0
  progress_status = "in progress"
  base_url = "https://{0}.qualtrics.com/API/v3/responseexports/".format(data_center)
  headers = {
      "content-type": "application/json",
      "x-api-token": api_token,
      }

  # Step 1: Creating Data Export
  download_request_url = base_url
  download_request_payload = '{"format":"' + file_format + '","survey_id":"' + survey_id + '"}'
  download_request_response = requests.request("POST", download_request_url, data=download_request_payload, headers=headers)
  progress_id = download_request_response.json()["result"]["id"]
  print(download_request_response.text)

  # Step 2: Checking on Data Export Progress and waiting until export is ready
  while request_check_progress < 100 and progress_status is not "complete":
      request_check_url = base_url + progress_id
      request_check_response = requests.request("GET", request_check_url, headers=headers)
      request_check_progress = request_check_response.json()["result"]["percentComplete"]
      print("Download is " + str(request_check_progress) + " complete")

  # Step 3: Downloading file
  request_download_url = base_url + progress_id + '/file'
  request_download = requests.request("GET", request_download_url, headers=headers, stream=True)

  # Step 4: Unzipping the file
  zipfile.ZipFile(io.BytesIO(request_download.content)).extractall()
  print('Complete')

cart = []

while True:
  s = input("Please enter the next survey ID (press enter when done): ")
  if s == "":
    break
  else:
    cart.append(s)

for s in cart:
  survey_id = s
  download_survey(api_token, survey_id, data_center)
