import requests


def http_api(url, option ,headers, keyword=None):
    # Function to make an HTTP GET request to the Microsoft Graph API
    if option is True:
        payload = {"search": keyword}
        try:
            response = requests.get(url, headers=headers, params=payload)
            response.raise_for_status()
        except requests.exceptions.RequestException as e:
            print(f"An error occurred while making the request: {e}")
            return None
        return response
    else:
        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
        except requests.exceptions.RequestException as e:
            print(f"An error occurred while making the request: {e}")
            return None
        return response        