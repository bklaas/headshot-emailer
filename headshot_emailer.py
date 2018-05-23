"""headshot_emailer.py
    Reads a spreadsheet of names, attendee numbers
    (which are also image names), and email addresses.
    Uses yagmail, a library specific for interacting with
    gmail's SMTP server, to send a specific, custom email
    and attachment to each email address.
"""
import yagmail
import pandas as pd
import sys

sys.exit
df = pd.read_csv("BenTest.csv")
cols_needed = ['ATTENDEE_NUM', 'LAST', 'FIRST', 'EMAIL']

all_there = all(col in df.columns for col in cols_needed)
if not all_there:
    print("ERROR!")
    print("Missing column in your spreadsheet. Make sure these are all present:")
    print(cols_needed)
    sys.exit()
    
df = df[cols_needed]
records = df.to_dict(orient= 'records')

# get an SMTP object
yag = yagmail.SMTP('ben@benklaas.com')

# pull the contents from the template file
with open("simple.html", 'r') as f:
        template = f.read()

# XXX Create/Append a log file that has emails of successfully sent emails
# XXX Do NOT send emails to successfully sent email addresses

# LOOP through rows and insert dynamic content
for r in records:
    print("---------")
    print("Sending email to", r['EMAIL'])
    img_name = r['ATTENDEE_NUM'] + ".jpg"
    print("Attaching:", img_name)

    content = template.format(FIRST=r['FIRST'])
    # attach the image as an attachment
    yag.send(
             to = r['EMAIL'],
             subject='Prototype Automated Email for Sending Headshots',
             attachments=img_name,
             contents=content
    )

