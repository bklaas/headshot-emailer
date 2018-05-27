"""headshot_emailer.py
    Uses pandas to read a spreadsheet of names,
    attendee numbers (which are also prefix of image names),
    and email addresses.
    Uses smtplib and email libraries from stdlib
"""
import pandas as pd

import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate

from pathlib import Path
import argparse
import sys

def warn_and_exit(msg):
    print("ERROR!", msg)
    sys.exit()

def get_opts():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(description="Automated headshot emailer")
    parser.add_argument('imagedir')
    parser.add_argument(
        '-d', '--dry_run', action='store_true', dest='dryrun',
        required=False, default=False,
        help="Don't send the email, just display what would have happened."
    )
    parser.add_argument(
        '-f', '--file', action='store', dest='mailing_list',
        required=False, default=mailing_list,
        help="Specify alternate location of mailing list spreadsheet. Default: " + mailing_list + " (in imagedir)"
    )
    return parser.parse_args()

def get_list(opts):
    """Read the mailing list, confirm it exists and has correct columns."""
    if not mailing_list.exists():
        warn_and_exit("Could not find " + str(mailing_list))

    # XXX change to excel
    df = pd.read_excel(mailing_list, dtype=str)
    df.columns = [col.upper() for col in df.columns]
    cols_needed = ['ATTENDEE_NUM', 'LAST', 'FIRST', 'EMAIL']

    all_there = all(col in df.columns for col in cols_needed)
    if not all_there:
        warn_and_exit("Missing column in your spreadsheet. Required columns: ", ', '.join(cols_needed))
    return df[cols_needed]

def send_individual_email(r):
    """If it needs sending, send the email."""

    send_from = "ben@benklaas.com"
    send_to = r['EMAIL']
    subject='Prototype Automated Email for Sending Headshots'

    # if we already sent it, don't do it
    if send_to in already_sent and \
       already_sent[send_to]['ATTENDEE_NUM'] == r['ATTENDEE_NUM']:
        print(send_to, "has already received", r['IMAGE_PATH'])
        return

    content = template.format(FIRST=r['FIRST'])
    print("To:", send_to)
    print("Subject:", subject)
    print("Attachment:", r['IMAGE_PATH'])


    if opts.dryrun:
        print("(dry run, no email sent)")
    else:
        msg = MIMEMultipart()
        msg['From'] = send_from
        msg['To'] = send_to
        msg['Date'] = formatdate(localtime=True)
        msg['Subject'] = subject
        msg.attach(MIMEText(content, 'html'))

        with open(str(r['IMAGE_PATH']), 'rb') as i:
            part = MIMEApplication(
                    i.read(),
                    Name=r['IMAGE_PATH'].name
                  )
        part['Content-Disposition'] = 'attachment; filename="%s"' % r['IMAGE_PATH'].name
        msg.attach(part)

        smtp.sendmail(send_from, send_to, msg.as_string())

        ## attach the image as an attachment
        add_to_sent_log(r)

def add_to_sent_log(r):
    # open sent_email_log for append
    with open(str(sent_email_log), 'a') as f:
        f.write(','.join([r['EMAIL'], r['ATTENDEE_NUM']]))
        f.write("\n")

def get_sent_emails():
    if sent_email_log.exists():
        df = pd.read_csv(sent_email_log, header=None)
        df.columns = ['EMAIL', 'ATTENDEE_NUM']
        df = df.set_index('EMAIL')
        return df.to_dict(orient ='index')
    else:
        return {}
    
mailing_list = "mailing_list.xlsx"
sent_email_list = "sent_emails.csv"
opts = get_opts()
imagedir = Path(opts.imagedir)
mailing_list = imagedir/mailing_list
sent_email_log = imagedir/sent_email_list

df = get_list(opts)
records = df.to_dict(orient= 'records')

# get an SMTP object
if not opts.dryrun:
    #smtp = smtplib.SMTP('smtp.gmail.com', 587)
    smtp = smtplib.SMTP('smtp.office365.com', 587)
    smtp.ehlo()
    ready = smtp.starttls()
    print(ready)
    status = smtp.login('eventmedia@ladenburg.com', 'XXXX')
    print(status)

# pull the contents from the template file
with open("simple.html", 'r') as f:
        template = f.read()

# Do NOT send emails to successfully sent email addresses
already_sent = get_sent_emails()

# create a list of records to send emails to
# 1. is a .jpg
# 2. image name starts with ATTENDEE_NUM
# 3. when match is made, add that image path to the record
to_send_list = []
pattern = "*.jpg"
for f in imagedir.glob(pattern):  
    for r in records:
        if f.name.startswith(r['ATTENDEE_NUM']):
            r['IMAGE_PATH'] = f
            to_send_list.append(r)
            print(f)

# LOOP through rows and insert dynamic content
for idx, r in enumerate(to_send_list, start=1):
    if idx == 1:
        print("--------------------------------------")
        print("Email", idx, "of", len(to_send_list))
        send_individual_email(r)

