import requests
from requests import Response
import json
from urllib.parse import urlencode
import pandas as pd
from openpyxl import load_workbook

baseURL = "https://api.messefrankfurt.com/service/esb_api/exhibitor-service/api/2.1/public/exhibitor/search"

params = {
    "pageNumber": "1",
    "pageSize": "10000",
    "findEventVariable": "AUTOMECHANIKA",
}

url = f"{baseURL}?{urlencode(params)}"

headers = {
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en-GB;q=0.9,en;q=0.8",
    "apikey": "LXnMWcYQhipLAS7rImEzmZ3CkrU033FMha9cwVSngG4vbufTsAOCQQ==",
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.5 Safari/605.1.15",
}


def omit_keys(original_dict: dict, keys_to_omit: list[str]):
    # Normalize the keys to lowercase for case-insensitive comparison
    normalized_keys_to_omit = [key.lower() for key in keys_to_omit]

    # Create a new dictionary excluding the specified keys
    filtered_dict = {
        key: value
        for key, value in original_dict.items()
        if key.lower() not in normalized_keys_to_omit
    }

    return filtered_dict


def saveJSON(response: Response):
    data = response.json()

    with open("response.json", "w") as file:
        json.dump(data, file, indent=4)


def fetchResponse():
    print("fetching...")
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        print("Success!")
        saveJSON(response)
        return response
    else:
        print("An error occurred:", response.status_code)
        exit(1)


def getResponse():
    try:
        with open("response.json") as file:
            return json.loads(file.read())
    except:
        return fetchResponse().json()


def parseLinkedInProfile(social: list[dict]):
    if social is not None:
        filtered_list = [
            theDict for theDict in social if theDict.get("network") == "linkedin"
        ]

        if len(filtered_list) > 0:
            return filtered_list[0]["url"]


def main():
    response = getResponse()

    hits: list = response["result"]["hits"]

    objects: list[dict] = []

    for hit in hits:
        exhibitor = hit["exhibitor"]
        address = exhibitor["address"]
        obj = {
            "name": exhibitor["name"],
            "address": omit_keys(
                address, ["email", "pob", "pobcity", "pobzip", "tel", "fax"]
            ),
            "email": address["email"],
            "phone": address["tel"],
            "linkedIn": parseLinkedInProfile(exhibitor["social"]),
            "url": exhibitor["href"],
            "aboutUs": exhibitor["description"]["text"],
        }

        objects.append(obj)

    print(f"total hits: {len(hits)}")

    with open("results.json", "w") as resultsFile:
        json.dump(objects, resultsFile, indent=4)

    excel_file_path = "results.xlsx"
    df = pd.DataFrame(objects)
    df.to_excel(excel_file_path, index=False)

    wb = load_workbook(excel_file_path)
    ws = wb.active

    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter  # Get the column letter

        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass

        adjusted_width = max_length + 2  # Add some extra space for aesthetics
        ws.column_dimensions[column_letter].width = adjusted_width

    wb.save(excel_file_path)


if __name__ == "__main__":
    main()
