#!/usr/bin/env python

#import requests
from src.get_info import get_site_id,get_site_list,get_file,get_emails,get_onedrive
import argparse
import pandas as pd
from tqdm import tqdm


# Create an ArgumentParser object to parse command-line arguments
parser = argparse.ArgumentParser(formatter_class=argparse.ArgumentDefaultsHelpFormatter)

# Add command-line arguments for the module, access token, and filter
parser.add_argument("-m", "--module", help="outlook,onedrive,sharepoint")
parser.add_argument("-jwt", "--accessToken", help="Microsoft Graph access token")
parser.add_argument("-f", "--filter", help="Search Specific Keyword", default="")

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
        output_emails = get_emails(headers, config["filter"])
        export_data(output_emails,columns_outlook,"outolok_emails.xlsx")
    elif config["module"] == "onedrive":
        output_onedrive = get_onedrive(headers, config["filter"])
        export_data(output_onedrive,columns_onedrive,"onedrive_files.xlsx")
    elif config["module"] == "sharepoint":
        list_site_id = get_site_id(headers, config["filter"])
        array_site = get_site_list(list_site_id, headers)
        output_files = get_file(array_site, headers)
        export_data(output_files,columns_sharepoint,"sharepoint_files.xlsx")
    else:
        print("Command not Found")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nExiting program due to KeyboardInterrupt")

