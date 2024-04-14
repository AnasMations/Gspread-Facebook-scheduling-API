import gspread
from oauth2client.service_account import ServiceAccountCredentials
import requests
import datetime
import time

def load_facebook_credentials(filename="Facebook token.txt"):
    """
    Load Facebook credentials from a file.

    Args:
    - filename (str): The name of the file containing the Facebook credentials.

    Returns:
    - tuple: Facebook page ID and access token.
    """
    with open(filename, "r") as f:
        fb_page_id = f.readline().strip("\n")
        fb_token = f.readline().strip("\n")
    return fb_page_id, fb_token

def post_on_facebook(fb_page_id, fb_token, confession, epoch_time):
    """
    Post a confession to a Facebook page.

    Args:
    - fb_page_id (str): Facebook page ID.
    - fb_token (str): Facebook access token.
    - confession (str): The confession message to post.
    - epoch_time (int): The scheduled publish time in epoch format.

    Returns:
    - requests.Response: The response object from the Facebook API.
    """
    params = {
        "published": "false",
        "message": confession,
        "scheduled_publish_time": str(int(epoch_time)),
        "access_token": fb_token
    }
    response = requests.post(f"https://graph.facebook.com/{fb_page_id}/feed", params=params)
    return response

def load_google_sheet(sheet_name):
    """
    Load a Google Sheet and its worksheets.

    Args:
    - sheet_name (str): The name of the Google Sheet.

    Returns:
    - gspread.Spreadsheet: The Google Sheet object.
    - gspread.Worksheet: The first worksheet of the Google Sheet.
    """
    scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
    client = gspread.authorize(creds)
    sheet = client.open(sheet_name).sheet1
    sheet2 = client.open(sheet_name).worksheet("Sheet2")
    return sheet, sheet2

def get_time_list(sheet2):
    """
    Extract the time list from the Google Sheet.

    Args:
    - sheet2 (gspread.Worksheet): The Google Sheet worksheet.

    Returns:
    - list: The list of scheduled times.
    """
    time_list = []
    for i in range(1, 100):
        value = sheet2.acell(f"D{i}").value
        if value:
            time_list.append(value)
        else:
            break
    return time_list

def post_confessions(sheet, sheet2, fb_page_id, fb_token):
    """
    Process and post confessions to Facebook.

    Args:
    - sheet (gspread.Worksheet): The Google Sheet worksheet containing confessions.
    - sheet2 (gspread.Worksheet): The Google Sheet worksheet containing metadata.
    - fb_page_id (str): Facebook page ID.
    - fb_token (str): Facebook access token.
    """
    conf_num = int(sheet2.acell("A1").value)
    start_point = int(sheet2.acell("B1").value)

    print(f"Number of new confessions: {len(sheet.get_all_values()) - start_point + 1}\n")

    time_list = get_time_list(sheet2)

    try:
        post_time = sheet2.acell("C1").value
        time_index = time_list.index(post_time.split(" ")[1].lower())
    except ValueError:
        print("Error within last scheduled time in sheet2!")
        start_point = 9999999

    print(post_time)
    day, month, year = map(int, post_time.split(" ")[0].split("-"))

    end_point = int(input("Enter end point in sheet - 0 to post all: "))
    if end_point == 0:
        end_point = len(sheet.get_all_values())

    print()

    batch_updates = []
    
    for i in range(start_point, end_point + 1):
        if i % 5 == 0:
            time.sleep(30)

        row = sheet.row_values(i)

        if not row[0][0].isdigit():
            row.reverse()

        confession = f"%23{conf_num}\n" + ", ".join(filter(None, row[2:])) + f":\n\n{row[1]}"
        confession = confession.replace("#", "%23")

        print(confession)
        print()

        is_skipped = sheet.acell(f"H{i}").value
        if is_skipped == "1":
            sheet.format(f"B{i}", {"backgroundColor": {"red": 1.0, "green": 0.0, "blue": 0.0}})
            print("Skipped")
        else:
            time_index += 1
            if time_index >= len(time_list):
                print("Schedule time overflow!")
                date = datetime.datetime(year, month, day) + datetime.timedelta(days=1)
                day, month, year = date.day, date.month, date.year
                time_index = 0

            post_time_parts = time_list[time_index].split(":")
            hour = int(post_time_parts[0])
            minute = int(post_time_parts[1][:-2])
            if post_time_parts[1][-2:].lower() == "pm":
                hour += 12
            date = datetime.datetime(year, month, day, hour, minute)
            epoch_time = date.timestamp()

            response = post_on_facebook(fb_page_id, fb_token, confession, epoch_time)
            print(response.json())
            if "error" in str(response.json()):
                continue

            batch_updates.append({
                "range": f"G{i}",
                "values": [[str(conf_num)]]
            })
            conf_num += 1

            new_post_time = f"{day}-{month}-{year} {post_time_parts[0]}:{post_time_parts[1]}"
            sheet2.update("C1", [[new_post_time]])
            batch_updates.append({
                "range": f"F{i}",
                "values": [[new_post_time]]
            })
            batch_updates.append({
                "range": f"B{i}",
                "format": {
                    "backgroundColor": {
                        "red": 0.0,
                        "green": 1.0,
                        "blue": 0.0
                    }
                }
            })
            print("Posted successfully!")

        batch_updates.append({
            "range": "A1",
            "values": [[str(conf_num)]]
        })
        batch_updates.append({
            "range": "B1",
            "values": [[str(i+1)]]
        })

        print("--------------------------------------------------\n")

    print("Updating sheet...")
    sheet.batch_update(batch_updates)
    print("Sheet updated!")

def main():
    """
    Main function to process confessions and post them on Facebook.
    """
    print("Loading Facebook credentials...")
    fb_page_id, fb_token = load_facebook_credentials()

    print("Loading spreadsheet...")
    sheet, sheet2 = load_google_sheet("NU Confessions (Responses)")

    post_confessions(sheet, sheet2, fb_page_id, fb_token)

    print("Program End!")
    input("")

if __name__ == "__main__":
    main()
