import imaplib
import email
from email.header import decode_header
from datetime import datetime
from dateutil.relativedelta import relativedelta
from bs4 import BeautifulSoup

# your credentials
username = "YOUR_USERNAME@hotmail.com"
password = "YOUR_PASSWORD"

# keywords to look for
keywords = ["remote", "C2C", "W2"]

# connect to the server
mail = imaplib.IMAP4_SSL("imap-mail.outlook.com")

# authenticate
mail.login(username, password)

# select the mailbox you want to delete in
mailbox = "inbox"
mail.select(mailbox)

# Define your date range
now = datetime.now()

print("Select a date range: ")
print("1: Today")
print("2: Yesterday")
print("3: Last 7 days")
print("4: Last month")
choice = input()

if choice == '1':
    start_date = now.date()  # Only today's emails
    end_date = now.date()  # Only today's emails
elif choice == '2':
    start_date = (now - relativedelta(days=1)).date()  # Only yesterday's emails
    end_date = (now - relativedelta(days=1)).date()  # Only yesterday's emails
elif choice == '3':
    start_date = (now - relativedelta(days=7)).date()  # Last 7 days
    end_date = now.date()
elif choice == '4':
    start_date = (now - relativedelta(months=1)).date()  # Last month
    end_date = now.date()

result, data = mail.uid("search", None, '(SINCE "{0}" BEFORE "{1}")'.format(
    start_date.strftime("%d-%b-%Y"), (end_date + relativedelta(days=1)).strftime("%d-%b-%Y")))

id_list = data[0].split()
keyword_count = {keyword: [] for keyword in keywords}

# function to get the body of the email
def get_body(msg):
    if msg.is_multipart():
        return get_body(msg.get_payload(0))
    else:
        return msg.get_payload(None, True)

for num in id_list:
    result, email_data = mail.uid('fetch', num, '(BODY.PEEK[])')

    raw_email = email_data[0][1].decode("utf-8")
    email_message = email.message_from_string(raw_email)

    # get the email subject
    decoded_header = decode_header(email_message['Subject'])[0]
    subject = decoded_header[0]
    if isinstance(subject, bytes):
        charset = decoded_header[1]
        if charset:
            subject = subject.decode(charset)
        else:
            subject = subject.decode('utf-8', errors='ignore')

    # get the sender
    from_header = decode_header(email_message['From'])[0]
    sender = from_header[0]
    if isinstance(sender, bytes):
        charset = from_header[1]
        if charset:
            sender = sender.decode(charset)
        else:
            sender = sender.decode('utf-8', errors='ignore')

    # get the date
    date_header = decode_header(email_message['Date'])[0]
    date = date_header[0]
    if isinstance(date, bytes):
        charset = date_header[1]
        if charset:
            date = date.decode(charset)
        else:
            date = date.decode('utf-8', errors='ignore')

    # get the email body
    body = get_body(email_message)
    if body is not None:
        body = body.decode('utf-8', errors='ignore')
        soup = BeautifulSoup(body, "html.parser")
        body = soup.get_text()

    # check if keyword is in email subject or body
    for keyword in keywords:
        if keyword in subject or (body and keyword in body):
            keyword_count[keyword].append((sender, subject, date))

# print keyword count
for keyword in keyword_count:
    print("Keyword: ", keyword)
    print("Count: ", len(keyword_count[keyword]))
    print("Emails: ")
    for email_info in keyword_count[keyword]:
        print("Sender: ", email_info[0])
        print("Subject: ", email_info[1])
        print("Date: ", email_info[2])
        print("-------------------------")
