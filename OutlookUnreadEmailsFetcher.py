import win32com.client
import datetime

# Create an instance of Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Access the Inbox
inbox = outlook.GetDefaultFolder(6) # 6 is the folder number for Inbox in Outlook

# Define the time range (last 15 minutes)
time_limit = datetime.datetime.now() - datetime.timedelta(minutes=15)

# Loop through the items in the Inbox
for message in inbox.Items:
    try:
        if message.UnRead and message.CreationTime > time_limit:
            print("Subject:", message.Subject)
            print("Received at:", message.CreationTime)
            print("Sender:", message.Sender)
            print("Body:", message.Body[:200]) # Print first 200 characters of the body
            print("-" * 50)
    except Exception as e:
        print("An error occurred:", e)

