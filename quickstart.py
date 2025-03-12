import datetime
import os.path
import time
import psutil
from datetime import datetime

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from pkg_resources import non_empty_lines

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/calendar"]


def main():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=8080)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    #Basic check to make sure the api is working
    try:
        service = build("calendar", "v3", credentials=creds)

        # Call the Calendar API
        now = datetime.utcnow().isoformat() + "Z"  # 'Z' indicates UTC time
        print("Getting the upcoming 10 events")
        events_result = (
            service.events()
            .list(
                calendarId="primary",
                timeMin=now,
                maxResults=10,
                singleEvents=True,
                orderBy="startTime",
            )
            .execute()
        )
        events = events_result.get("items", [])

        if events:
            getTime(service)

        if not events:
            print("No upcoming events found.")
            return

        # Prints the start and name of the next 10 events
        for event in events:
            start = event["start"].get("dateTime", event["start"].get("date"))
            print(start, event["summary"])

    except HttpError as error:
        print(f"An error occurred: {error}")


def get_running_processes():
    # Returning names for all processes
    return {p.name() for p in psutil.process_iter(attrs=['name'])}


def getTime(service):
    #Enter program name here which will be the one tracked
    #Program name can be found through task manager
    target_program = "Taskmgr.exe"

    print("Monitoring Program...")

    program_running = False
    start_time = None

    while True:
        running_processes = get_running_processes()

        if target_program in running_processes:
            if not program_running:
                #Checks when the program first opens and gets the time and date
                #Will then uddate the program_running variable to let the rest of the code know it has started
                start_time = datetime.now()
                print(f"{target_program} started at {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
                converted_start_time = start_time.isoformat(timespec="seconds")
                program_running = True
        else:
            if program_running:
                #Executes when program closes, calculates end time
                end_time = datetime.now()
                print(f"{target_program} closed at {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
                duration = end_time - start_time
                print(f"Total run time: {duration}")
                converted_end_time = end_time.isoformat(timespec="seconds")
                print(converted_start_time)
                print(converted_end_time)

                # Creating the dictionary to send off to the library api
                event = {
                    #Change summary depending on what you want the name to be on the calendar event
                    "summary": "Coding",
                    "start": {
                        "dateTime": converted_start_time,
                        "timeZone": "Europe/London",
                    },
                    "end": {
                        "dateTime": converted_end_time,
                        "timeZone": "Europe/London",
                    },
                }

                calendar_id = "primary"
                #Sending off the dictionary through the api to update on the google calendar
                event = service.events().insert(calendarId=calendar_id, body=event).execute()
                program_running = False

        time.sleep(5)


if __name__ == "__main__":
    main()
