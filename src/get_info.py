from src.utlis import http_api
from tqdm import tqdm


# -----------------------------------------------------------
# This code is a Python script that uses the Microsoft Graph         
# API to retrieve data from Outlook, OneDrive, and Sharepoint.       
# It takes command-line arguments for the module, access token,     
# and filter to search for specific keywords. It defines column     
# names for different modules and headers for the HTTP request.      
# It contains functions to make an HTTP GET request to the          
# Microsoft Graph API, get site IDs, get site lists, get files       
# from Sharepoint, dump emails from Outlook, get files from          
# OneDrive.
# -----------------------------------------------------------


def get_site_id(headers, filter=None):
    # Function to get site id
    response = http_api("https://graph.microsoft.com/v1.0/sites", True, headers, filter)
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
                response = http_api(url, False, headers)
                pbar.update(1)
            else:
                break
        return list_site_id

def get_site_list(list_site_id, headers):
    # Function to get site list id
    if list_site_id is None:
        return None
    list_site_data = []
    for site_id in list_site_id:
        response = http_api("https://graph.microsoft.com/v1.0/sites/{}/lists".format(site_id), False, headers)
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

def get_file(site_data, headers):
    # Function to get files array
    if site_data is None:
        return None
    for n in site_data:
        response = http_api("https://graph.microsoft.com/v1.0/sites/{}/lists/{}/items?$expand=driveItem".format(n["site_id"],n["list_id"]), False, headers)
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
                    response = http_api(url, False, headers)
                    if response is None:
                        continue
                    pbar.update(1)
                else:
                    break
        return list_files

def get_emails(headers, filter):
    # Function to dump email
    print(filter)
    if filter != "":
        response = http_api("https://graph.microsoft.com/v1.0/me/messages", True, headers, filter)
    else:
        response = http_api("https://graph.microsoft.com/v1.0/me/messages", False, headers)
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
                response = http_api(url, False, headers)
                if response is None:
                    continue
                pbar.update(1)
            else:
                break
        return list_emails

def get_onedrive(headers,filter):
    # Function to get file in onedrive
    response = http_api("https://graph.microsoft.com/v1.0/me/drive/root/search(q='{}')".format(filter), False, headers)
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
                response = http_api(url, False, headers)
                if response is None:
                    continue
                pbar.update(1)
            else:
                break
    return list_files
