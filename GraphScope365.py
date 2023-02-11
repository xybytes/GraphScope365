#!/usr/bin/env python

import requests
import argparse
import pandas as pd
from tqdm import tqdm


# Create an ArgumentParser object to parse command-line arguments
parser = argparse.ArgumentParser(formatter_class=argparse.ArgumentDefaultsHelpFormatter)

# Add command-line arguments for the module, access token, and filter
parser.add_argument("-m", "--module", help="outlook,onedrive,sharepoint")
parser.add_argument("-jwt", "--accessToken", help="Microsoft Graph access token")
parser.add_argument("-f", "--filter", help="Search Specific Keyword", default="*")

# Parse the command-line arguments
args = parser.parse_args()
config = vars(args)

# Define column names for the different modules
columns_sharepoint = ["File Name","Size", "File type", "Shared", "URL", "Created Date Time", "Last Modified Date Time", "Created By", "Last Modified By"] 
columns_outlook = ["Created Date Time" , "From", "To" , "CC", "Subject", "Body Preview", "URL", "Attachments"]
columns_onedrive = ["File Name","Size", "File type", "Created Date Time", "Last Modified Date Time", "Created By", "Last Modified By", "URL"]

# Define the headers for the HTTP request
headers = {
"Host": "graph.microsoft.com",
"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:109.0) Gecko/20100101 Firefox/109.0",
"Accept": "*/*",
"Accept-Language": "en-US,en;q=0.5",
"Accept-Encoding": "gzip, deflate",
"Referer": "https://developer.microsoft.com/",
"Sdkversion": "GraphExplorer/4.0, graph-js/3.0.2 (featureUsage=6)",
"Cache-Control": "no-cache",
"Pragma": "no-cache",
"Authorization": "Bearer " + config["accessToken"],
"Origin": "https://developer.microsoft.com",
"Sec-Fetch-Dest": "empty",
"Sec-Fetch-Mode": "cors",
}


def http_api(url, keyword=None):
    # Function to make an HTTP GET request to the Microsoft Graph API
    payload = {"search": keyword}
    try:
        response = requests.get(url, headers=headers, params=payload)
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"An error occurred while making the request: {e}")
        return None
    return response

def get_site_id():
    # Function to get site id
    response = http_api("https://graph.microsoft.com/v1.0/sites", config["filter"])
    if response is None:
        return None
    list_site_id = []
    with tqdm(total=None, desc="Getting Sharepoint IDs") as pbar:
        while True:
            try:
                data = response.json()
            except ValueError as e:
                print(f"An error occurred while parsing the response JSON: {e}")
                return None
            for item in data["value"]:
                list_site_id.append(item["id"])
            if '@odata.nextLink' in data:
                url = data['@odata.nextLink']
                response = requests.get(url, headers=headers)
                pbar.update(1)
            else:
                break
        return list_site_id

def get_site_list(list_site_id):
    # Function to get site list id
    if list_site_id is None:
        return None
    list_site_data = []
    for site_id in list_site_id:
        response = http_api("https://graph.microsoft.com/v1.0/sites/{}/lists".format(site_id))
        if response is None:
            continue
        try:
            data = response.json()
        except ValueError as e:
            print(f"An error occurred while parsing the response JSON: {e}")
            return None
        for item in data["value"]:
            if item["name"] == "Shared Documents":
                list_site_data.append({"site_id":site_id, "list_id":item["id"], "name":item["name"]})
            else:
                pass
    return list_site_data

def get_file(site_data):
    # Function to get files array
    if site_data is None:
        return None
    for n in site_data:
        response = http_api("https://graph.microsoft.com/v1.0/sites/{}/lists/{}/items?$expand=driveItem".format(n["site_id"],n["list_id"]))
        if response is None:
            return None
        list_files = []
        with tqdm(total=None, desc="Getting Sharepoint Files") as pbar:
            while True:
                try:
                    data = response.json()
                except ValueError as e:
                    print(f"An error occurred while parsing the response JSON: {e}")
                    return None
                for file in data["value"]:
                    file_data = []
                    if file["contentType"]["name"] == "Document":
                        file_data.append(file["fields"]["LinkFilename"])
                        file_data.append(file["fields"]["FileSizeDisplay"])
                        file_data.append(file["fields"]["DocIcon"])
                        file_data.append(file["driveItem"]["shared"]["scope"])
                        file_data.append(file["webUrl"])
                        file_data.append(file["createdDateTime"])
                        file_data.append(file["lastModifiedDateTime"])
                        file_data.append(file["createdBy"]["user"]["email"])
                        file_data.append(file["lastModifiedBy"]["user"]["email"])
                        list_files.append(file_data)
                    else:
                        pass
                if '@odata.nextLink' in data:
                    url = data['@odata.nextLink']
                    response = requests.get(url, headers=headers)
                    if response is None:
                        continue
                    pbar.update(1)
                else:
                    break
        return list_files

def get_emails():
    # Function to dump email
    response = http_api("https://graph.microsoft.com/v1.0/me/messages", config["filter"])
    if response is None:
        return None
    list_emails = []
    with tqdm(total=None, desc="Getting Outlook Emails") as pbar:
        while True:
            try:
                data = response.json()
            except ValueError as e:
                print(f"An error occurred while parsing the response JSON: {e}")
                return None
            for email in data["value"]:
                email_data = []
                email_data.append(email["createdDateTime"])
                email_data.append(email["from"]["emailAddress"]["address"])
                toRecipients = ""
                for recepient in email["toRecipients"]:
                    toRecipients += recepient["emailAddress"]["address"]
                    toRecipients += ";"
                email_data.append(toRecipients)
                ccRecipients = ""
                for recepient in email["ccRecipients"]:
                    ccRecipients += recepient["emailAddress"]["address"]
                    ccRecipients += ";"
                email_data.append(ccRecipients)
                email_data.append(email["subject"])
                email_data.append(email["bodyPreview"])
                email_data.append(email["webLink"])
                email_data.append(email["hasAttachments"])
                list_emails.append(email_data)
            if '@odata.nextLink' in data:
                url = data['@odata.nextLink']
                response = requests.get(url, headers=headers)
                if response is None:
                    continue
                pbar.update(1)
            else:
                break
        return list_emails

def get_onedrive():
    # Function to get file in onedrive
    response = http_api("https://graph.microsoft.com/v1.0/me/drive/root/search(q='{}')".format(config["filter"]))
    if response is None:
        return None
    list_files = []
    with tqdm(total=None, desc="Getting OneDrive Files") as pbar:
        while True:
            try:
                data = response.json()
            except ValueError as e:
                print(f"An error occurred while parsing the response JSON: {e}")
                return None
            for file in data["value"]:
                file_data = []
                if "file" in file:
                    file_data.append(file["name"])
                    file_data.append(file["size"])
                    file_data.append(file["file"]["mimeType"])
                    file_data.append(file["createdDateTime"])
                    file_data.append(file["lastModifiedDateTime"])
                    file_data.append(file["createdBy"]["user"]["email"])
                    file_data.append(file["lastModifiedBy"]["user"]["email"])
                    file_data.append(file["webUrl"])
                    list_files.append(file_data)
                else:
                    pass
            if '@odata.nextLink' in data:
                url = data['@odata.nextLink']
                response = requests.get(url, headers=headers)
                if response is None:
                    continue
                pbar.update(1)
            else:
                break
    return list_files

def export_data(data_array,fields,file_name):
    # Function to export data in a xlsx file
    try:
        df = pd.DataFrame(data_array, columns = fields)
        df.to_excel(excel_writer = file_name, index=False)
        print("Data exported successfully to the file:", file_name)
    except Exception as e:
        print("Error exporting data to the file:", file_name)
        print("Error message:", str(e))

def main():
    if config["module"] == "outlook":
        output_emails = get_emails()
        export_data(output_emails,columns_outlook,"outolok_emails.xlsx")
    elif config["module"] == "onedrive":
        output_onedrive = get_onedrive()
        export_data(output_onedrive,columns_onedrive,"onedrive_files.xlsx")
    elif config["module"] == "sharepoint":
        list_site_id = get_site_id()
        array_site = get_site_list(list_site_id)
        output_files = get_file(array_site)
        export_data(output_files,columns_sharepoint,"sharepoint_files.xlsx")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nExiting program due to KeyboardInterrupt")

