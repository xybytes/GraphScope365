import requests
import argparse
import pandas as pd
from tqdm import tqdm


parser = argparse.ArgumentParser(formatter_class=argparse.ArgumentDefaultsHelpFormatter)
parser.add_argument("-jwt", "--accessToken", help="Microsoft Graph access token")
parser.add_argument("-f", "--filter", help="Search Specific Keyword", default="*")
args = parser.parse_args()
config = vars(args)

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
    payload = {"search": keyword}
    response = requests.get(url, headers=headers, params=payload)
    return response

def get_site_id():
    response = http_api("https://graph.microsoft.com/v1.0/sites", config["filter"])
    list_site_id = []
    with tqdm(total=None, desc="Getting Sharepoint IDs") as pbar:
        while True:
            data = response.json()
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
    list_site_data = []
    for site_id in list_site_id:
        response = http_api("https://graph.microsoft.com/v1.0/sites/"+site_id+"/lists")
        data = response.json()
        for item in data["value"]:
            if item["name"] == "Shared Documents":
                list_site_data.append({"site_id":site_id, "list_id":item["id"], "name":item["name"]})
            else:
                pass
    return list_site_data

def get_file(site_data):
    for n in site_data:
        response = http_api("https://graph.microsoft.com/v1.0/sites/"+n["site_id"]+"/lists/"+n["list_id"]+"/items?$expand=driveItem")
        list_files = []
        with tqdm(total=None, desc="Getting Sharepoint Files") as pbar:
            while True:
                data = response.json()
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
                    pbar.update(1)
                else:
                    break
        return list_files

def get_emails():
    response = http_api("https://graph.microsoft.com/v1.0/me/messages", config["filter"])
    list_emails = []
    with tqdm(total=None, desc="Getting Outlook Emails") as pbar:
        while True:
            data = response.json()
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
                pbar.update(1)
            else:
                break
        return list_emails


def get_onedrive():
    response = http_api("https://graph.microsoft.com/v1.0/me/drive/root/search(q='')")
    list_files = []
    with tqdm(total=None, desc="Getting OneDrive Files") as pbar:
        while True:
            data = response.json()
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
                pbar.update(1)
            else:
                break
    return list_files

def export_data(data_array,fields,file_name):
    df = pd.DataFrame(data_array, columns = fields)
    df.to_excel(excel_writer = file_name, index=False)

list_site_id = get_site_id()
array_site = get_site_list(list_site_id)
output_files = get_file(array_site)
columns_sharepoint = ["File Name","Size", "File type", "Shared", "URL", "Created Date Time", "Last Modified Date Time", "Created By", "Last Modified By"] 
export_data(output_files,columns_sharepoint,"sharepoint.xlsx")


#output_emails = get_emails()
#columns_outlook = ["Created Date Time" , "From", "To" , "CC", "Subject", "Body Preview", "URL", "Attachments"]
#export_data(output_emails,columns_outlook,"outolok_emails.xlsx")

#output_onedrive = get_onedrive()
#columns_onedrive = ["File Name","Size", "File type", "Created Date Time", "Last Modified Date Time", "Created By", "Last Modified By", "URL"] 
#export_data(output_onedrive,columns_onedrive,"onedrive.xlsx")

