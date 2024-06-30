import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os

# Define the scope and credentials to access the Google Sheets API
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('/credentials1.json', scope)  # Replace 'credentials.json' with your credentials file

# Authorize the credentials and open the Google Sheet
gc = gspread.authorize(credentials)
worksheet = gc.open('test').sheet1  # Replace with your Google Sheet name

# Get tomorrow's date (the day before the ticket scan day)
tomorrow = datetime.date.today() + datetime.timedelta(days=1)

# Find the row index corresponding to tomorrow's date in column B
row_index = None
dates = worksheet.col_values(2)
for index, date in enumerate(dates):
    if date == tomorrow.strftime('%-m/%-d/%Y'):  # Adjusted date format for comparison
        row_index = index + 1  # Adding 1 because index starts from 0, but row numbering starts from 1
        break

if row_index is None:
    print("Tomorrow's date not found in the sheet.")
else:
    # Get the email address of the person assigned for tomorrow from columns C to K
    def get_column_letter(column_index):
        letters = ''
        while column_index > 0:
            column_index, remainder = divmod(column_index - 1, 26)
            letters = chr(65 + remainder) + letters
        return letters

word = "Scan"
print(row_index)
row_values = worksheet.row_values(row_index)

column_alphabet = None
for col_index, cell_value in enumerate(row_values, start=1):
    if isinstance(cell_value, str) and word in cell_value:
        column_alphabet = get_column_letter(col_index)
        print(column_alphabet)
        break

if column_alphabet is None:
    print("Column with 'Scan' not found.")
else:
    Email_index = 1
    Email_reference = f"{column_alphabet}{Email_index}"
    print(Email_reference)
    default_email_prefix = "@gmail.com"
    assigned_person_email_name = worksheet.acell(Email_reference).value
    assigned_person_email = assigned_person_email_name + default_email_prefix
    print(assigned_person_email)

    if assigned_person_email is None:
        print("Assigned person's email not found.")
    else:
        # Create a function to generate the Google Calendar event link
        def generate_google_calendar_link():
            event_title = "Cleaning Task Reminder"
            event_description = "Kind reminder,You have been allocated to do the Cleaning Task for tomorrow"
            event_start_date = tomorrow.strftime('%Y%m%d')  # Format as YYYYMMDD
            event_start_time = "030000"  # Start time is 03:00 AM in UTC
            event_end_time = "040000"    # End time is 04:00 AM in UTC

            google_calendar_link = f"https://www.google.com/calendar/render?action=TEMPLATE&text={event_title}&details={event_description}&dates={event_start_date}T{event_start_time}Z/{event_start_date}T{event_end_time}Z"
            return google_calendar_link

        # Send an email to the assigned person
        sender_email = os.environ.get("emailaddress")
        sender_password = os.environ.get("emailpass") 

        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = assigned_person_email
        message['Subject'] = 'Cleaning Task Reminder Tomorrow'

        google_calendar_link = generate_google_calendar_link()
        button_html = f'<a href="{google_calendar_link}" target="_blank"><button style="background-color: #2196F3; color: white; border: none; border-radius: 4px; padding: 10px 20px; text-align: center; text-decoration: none; display: inline-block; font-size: 16px; margin: 4px 2px; cursor: pointer;">Add to Google Calendar</button></a>'
        body = f'Hello,<br> Kind reminder,You have been allocated to do Cleaning Task for tomorrow.<br> {button_html} <br><br> <i>This is an auto-generated email.</i>'
        message.attach(MIMEText(body, 'html'))

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(message)
        server.quit()
