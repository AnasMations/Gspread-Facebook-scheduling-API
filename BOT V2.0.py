import gspread
from oauth2client.service_account import ServiceAccountCredentials
import requests
import datetime
import time

#Current bug: can't post a confession that includes hash sign #

f = open("Facebook token.txt", "r+")
fbPageid = f.readline().strip("\n")
fbToken = f.readline().strip("\n")
f.close()

def postOnFacebook(confession, epochTime):
    response = requests.post("https://graph.facebook.com/" + fbPageid + "/feed" +
                            "?published=false" +
                            "&message=" + confession +
                            "&scheduled_publish_time=" + str(int(epochTime)) +
                            "&access_token=" + fbToken)
    return response

print("Loading spreadsheet...")

scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
client = gspread.authorize(creds)
sheet = client.open("NU Confessions (Responses)").sheet1  # Open the spreadhseet
sheet2 = client.open("NU Confessions (Responses)").worksheet("Sheet2")

confNum = int(sheet2.acell("A1").value)
startPoint = int(sheet2.acell("B1").value)

print("Number of new confessions: "+str(len(sheet.get_all_values())-startPoint+1))
print()

#Schedule times
#timeList = ["12:30am", "1:00am", "1:30am", "2:00am", "7:00pm", "7:30pm", "8:00pm", "8:30pm", "9:00pm", "9:30pm", "10:00pm", "10:30pm", "11:00pm", "11:30pm", "11:59pm"]
timeList=[]
for i in range(1, 100):
    value = sheet2.acell("D"+str(i)).value
    if(value == None):
        break
    else:
        timeList.append(value)

try:
    postTime = sheet2.acell("C1").value
    timeIndex = timeList.index(postTime.split(" ")[1].lower()) #Get index of current time to post
except:
    print("Error within last scheduled time in sheet2!")
    startPoint = 9999999

print(postTime)
day = int(postTime.split(" ")[0].split("-")[0])
month = int(postTime.split(" ")[0].split("-")[1])
year = int(postTime.split(" ")[0].split("-")[2])

endPoint = int(input("Enter end point in sheet - 0 to post all: "))
if endPoint == 0:
    endPoint = len(sheet.get_all_values())

print()

c=1
for i in range(startPoint, endPoint + 1):
    #delay for Gspread api write request quota
    if(c%5==0):
        time.sleep(30)

    row = sheet.row_values(i)

    #Reverse arabic confession
    if not (row[0][0].isdigit()):
        row.reverse()

    confession = "#" + str(confNum) + "\n"
    for j in range(2, len(row)):
        if j == len(row)-1:
            confession = confession + row[j] + ": \n\n"
        else:
            if (row[j] == None) or (row[j] == "") or (row[j] == "Prefer not to say"):
                continue
            confession = confession + row[j] + ", "
    confession = confession + row[1]
    confession = confession.replace("#", "") #remove all hash symbols form string
    print(confession)
    print()

    isSkipped = sheet.acell("H"+str(i)).value
    if isSkipped=="1":
        sheet.format("B"+str(i), {"backgroundColor": {"red": 1.0, "green": 0.0, "blue": 0.0}})
        print("Skipped")
    else:
        timeIndex += 1
        if(timeIndex >= len(timeList)):
            print("Schedule time overflow!")
            date = datetime.datetime(year, month, day)
            date += datetime.timedelta(days=1)
            day = date.day
            month = date.month
            year = date.year
            timeIndex = 0

        postTime = timeList[timeIndex].split(":")
        hour = int(postTime[0])
        minute = int(postTime[1][:-2])
        if ((postTime[1][-2] + postTime[1][-1]).lower() == "pm"):
            hour += 12
        date = datetime.datetime(year, month, day, hour, minute)
        epochTime = date.timestamp()

        #Facebook scheduling request
        response = postOnFacebook(confession, epochTime)
        print(str(response.json()))
        if str(response.json())[2:7]=="error":
            break

        sheet.update("G" + str(i),str(confNum))
        confNum += 1

        sheet2.update("C1", str(day) + "-" + str(month) + "-" + str(year) + " " + postTime[0] + ":" + postTime[1])
        sheet.update("F"+str(i), str(day)+"-"+str(month)+"-"+str(year)+" "+postTime[0]+":"+postTime[1])
        sheet.format("B" + str(i), {"backgroundColor": {"red": 0.0, "green": 1.0, "blue": 0.0}})
        print("Posted successfully!")

    sheet2.update("A1", str(confNum))
    sheet2.update("B1", str(i+1))
    c=c+1

    print("--------------------------------------------------\n")

print("Program End!")
input("")
