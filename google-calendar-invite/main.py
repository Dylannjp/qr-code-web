"""from flask import Flask, request, render_template

import datetime as dt
import os.path
import time

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/calendar"]


def main():
  creds = None

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
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  try:
    service = build("calendar", "v3", credentials=creds)

    event = {
      "summary": "Fun Board Game Night",
      "location": "Cunningham 4th Floor",
      "description": "Play board games!!!",
      "colorID": 8,
      "start": {
        "dateTime": "2024-11-15T09:00:00.000Z",
        "timeZone": "America/Chicago"
      },
      "end": {
        "dateTime": "2024-11-15T11:30:00.000Z",
        "timeZone": "America/Chicago"
      },
      "attendees": [
        { "email": "dputra@hawk.iit.edu"},
      ]
    }
    event = service.events().insert(calendarId="primary", body=event).execute()
    print(f"Event creation completed, event link: {event.get('htmlLink')}")
    
    time.sleep(5)
    event_id = event['id']
    updated_attendees = event.get("attendees", [])
    updated_attendees.append({"email": "dylanjm001@gmail.com"})

    # Use the .update method to update the event with the new attendee
    updated_event = service.events().update(
        calendarId="primary",
        eventId=event_id,
        body={
                "summary": event["summary"],
                "location": event["location"],
                "description": event["description"],
                "start": event["start"],
                "end": event["end"],
                "attendees": updated_attendees
            }
    ).execute()

    print(f"Event updated with new attendee: {updated_event.get('htmlLink')}")
              
  except HttpError as error:
    print(f"An error occurred: {error}")

if __name__ == "__main__":
  main()

"""

from flask import Flask, request, render_template, redirect, url_for, send_file
from io import BytesIO
import datetime as dt
import os.path
import qrcode
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


main = Flask(__name__)

SCOPES = ["https://www.googleapis.com/auth/calendar"]

def get_google_calendar_service():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return build("calendar", "v3", credentials=creds)

@main.route('/')
def index():
    return render_template('index.html')

@main.route('/event-created/<event_id>', methods=['GET'])
def event_created(event_id):
    return render_template('event_created.html', event_id=event_id)

@main.route('/create-event', methods=['POST'])
def create_event():
    data = request.form
    summary = data['summary']
    location = data.get('location', '')
    description = data.get('description', '')
    start = data['start']
    end = data['end']
    timezone = data['timezone']
    attendees_emails = data.get('attendees', '').split(',')

    attendees = [{'email': email.strip()} for email in attendees_emails if email.strip()]

    event = {
        "summary": summary,
        "location": location,
        "description": description,
        "start": {
            "dateTime": start,
            "timeZone": timezone,
        },
        "end": {
            "dateTime": end,
            "timeZone": timezone,
        },
        "attendees": attendees,
    }

    try:
        service = get_google_calendar_service()
        created_event = service.events().insert(calendarId="primary", body=event).execute()
        event_id = created_event.get('id')
        event_link = created_event.get('htmlLink')

        return render_template(
            'event_created.html',
            event_id=event_id,
            event_link=event_link
        )
    except HttpError as error:
        return f"An error occurred: {error}"

@main.route('/qr-code/<event_id>', methods=['GET'])
def qr_code(event_id):
    # URL for the add_attendees page with the specific event ID
    add_attendees_url = f"{request.host_url}add-attendees/{event_id}"

    # Generate the QR code
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(add_attendees_url)
    qr.make(fit=True)

    # Save QR code to a BytesIO object
    buffer = BytesIO()
    img = qr.make_image(fill_color="black", back_color="white")
    img.save(buffer, format="PNG")
    buffer.seek(0)

    # Serve the QR code as a downloadable file
    return send_file(
        buffer,
        mimetype="image/png",
        as_attachment=True,
        download_name="qr_code.png"
    )

import pandas as pd
from flask import Flask, render_template, request, redirect, url_for
import os

app = Flask(__name__)

# Route to add attendees
@main.route('/add-attendees/<event_id>', methods=['GET', 'POST'])
def add_attendees(event_id):
    if request.method == 'POST':
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
        service = build("calendar", "v3", credentials=creds)
        # Get form data
        name = request.form.get('name')
        number = request.form.get('number')
        email = request.form.get('email')
        a_number = request.form.get('a_number')
        year = request.form['year']

        # Data to be saved to Excel
        attendee_data = {
            "Name": [name],
            "Phone Number": [number],
            "Email": [email],
            "A Number": [a_number],
            "College Year": [year]
        }

        # Create or load the Excel file
        excel_file = 'attendees.xlsx'
        if os.path.exists(excel_file):
            df = pd.read_excel(excel_file)
        else:
            df = pd.DataFrame(columns=["Name", "Phone Number", "Email", "A Number", "College Year"])

        # Append the new attendee to the DataFrame
        new_attendee = pd.DataFrame(attendee_data)
        df = pd.concat([df, new_attendee], ignore_index=True)

        # Save the updated DataFrame back to Excel
        df.to_excel(excel_file, index=False)

        try:
            # Retrieve the existing event
            event = service.events().get(calendarId="primary", eventId=event_id).execute()

            # Add new attendee to the event
            attendees = event.get("attendees", [])
            attendees.append({"email": email})
            event["attendees"] = attendees

            # Update the event with the new attendee
            updated_event = service.events().update(
                calendarId="primary",
                eventId=event_id,
                body=event
            ).execute()

            return render_template('add_attendees.html', event_id=event_id, success=True)

        except HttpError as error:
            print(f"An error occurred: {error}")
            return render_template('add_attendees.html', event_id=event_id, error=str(error)) 

    return render_template('add_attendees.html', event_id=event_id)

if __name__ == '__main__':
    main.run(debug=True)
