import requests
import csv
import time
import json
from pathlib import Path
import win32com.client

SLEEP = 0  # Time in seconds the script should wait between requests
url_list = []
url_statuscodes = []
url_statuscodes.append(["url", "status_code"])  # set the file header for output


def getStatuscode(url):
    try:

        session_obj = requests.Session()
        response = session_obj.get(url, headers={"User-Agent": "chrome"})

        # response = requests.head(url, verify=False, timeout=5)  # it is faster to only request the header
        return (response.status_code)
    except:
        return -1


# Url checks from file Input
# use one url per line that should be checked
with open('urls.csv', newline='') as f:
    reader = csv.reader(f)
    for row in reader:
        url_list.append(row[0])
# Loop over full list
for url in url_list:
    print(url)
    check = [url, getStatuscode(url)]
    time.sleep(SLEEP)
    url_statuscodes.append(check)
# Save file
with open("urls_withStatusCode.csv", "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerows(url_statuscodes)


def csv_to_json(csv_file_path, json_file_path):
    # create a dictionary
    data_dict = {}

    # Step 2
    # open a csv file handler
    with open(csv_file_path, encoding='utf-8') as csv_file_handler:
        csv_reader = csv.DictReader(csv_file_handler)

        # convert each row into a dictionary
        # and add the converted data to the data_variable

        for rows in csv_reader:
            # assuming a column named 'No'
            # to be the primary key
            key = rows['url']
            data_dict[key] = rows

    # open a json file handler and use json.dumps
    # method to dump the data
    # Step 3
    with open(json_file_path, 'w', encoding='utf-8') as json_file_handler:
        # Step 4
        json_file_handler.write(json.dumps(data_dict, indent=4))


csv_file_path = Path("urls_withStatusCode.csv")
json_file_path = Path("urls.json")
csv_to_json(csv_file_path, json_file_path)


def send_mail(write_body):
    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'chaitanya.k@techsophy.com'
    mail.Subject = 'Bad Request '
    mail.Body = write_body
    # mail.CC = 'sandeep.k@techsophy.com'
    mail.Send()


f = open("urls.json")
data = json.load(f)
for i in data.values():
    if i.get("status_code") != "200":
        send_mail("unable to reach the site  " + i.get("url") + "  failed with error code  " + i.get("status_code"))
f.close()


