"""headshot_emailer.py
    Uses pandas to read a spreadsheet of names,
    attendee numbers (which are also prefix of image names),
    and email addresses.
    Uses smtplib and email libraries from stdlib
"""
import pandas as pd

import time
import smtplib
import getpass
import configparser
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate

from pathlib import Path
import argparse
import csv
import sys

## Change these as needed
send_from = 'eventmedia@ladenburg.com'
smtpserver = 'smtp.office365.com'
smtpport = 587

def warn_and_exit(msg):
    print("ERROR!", msg)
    sys.exit()

def get_opts():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(description="Automated headshot emailer")
    parser.add_argument('imagedir', action='store',
                        help="Name of directory where headshots for this run are stored")
    parser.add_argument('--conference', action='store', dest='conference',
                        help="Name of conference, with specifics mapped out in config.ini")
    parser.add_argument(
        '-d', '--dry_run', action='store_true', dest='dryrun',
        required=False, default=False,
        help="Don't send the email, just display what would have happened."
    )
    parser.add_argument(
        '-f', '--file', action='store', dest='mailing_list',
        required=False, default='mailing_list.xlsx',
        help="Specify alternate location of mailing list spreadsheet. Default: <conference>_mailing_list.xlsx"
    )
    return parser.parse_args()

def get_list(opts):
    """Read the mailing list, confirm it exists and has correct columns."""
    if not mailing_list.exists():
        warn_and_exit("Could not find " + str(mailing_list))

    df = pd.read_excel(mailing_list, dtype=str)
    df.columns = [col.upper() for col in df.columns]
    cols_needed = ['CONFIRMATION NUMBER', 'LAST NAME', 'FIRST NAME', 'EMAIL ADDRESS']

    all_there = all(col in df.columns for col in cols_needed)
    if not all_there:
        warn_and_exit("Missing column in your spreadsheet. Required columns: " + ', '.join(cols_needed))
    return df

def send_individual_email(r):
    """If it needs sending, send the email."""

    send_to = r['EMAIL ADDRESS']

    # if we already sent it, don't do it
    if send_to in already_sent and \
       already_sent[send_to]['CONFIRMATION NUMBER'] == r['CONFIRMATION NUMBER']:
        print(send_to, "has already received", r['IMAGE_PATH'])
        return

    subject = config['subject']
    content = template.format(FIRST=r['FIRST NAME'])
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
    make_headers = False
    if not sent_email_log.exists():
        make_headers = True
    with open(str(sent_email_log), 'a') as f:
        writer = csv.writer(f)
        if make_headers:
            writer.writerow([k for k in sorted(r)])
        writer.writerow([str(r[k]) for k in sorted(r)])

def get_sent_emails():
    if sent_email_log.exists():
        df = pd.read_csv(sent_email_log)
        df = df.set_index('EMAIL ADDRESS')
        return df.to_dict(orient ='index')
    else:
        return {}
    
def get_config(c):
    config = configparser.ConfigParser()
    config_file = config.read("config.ini")
    return {k: v for k, v in config[c].items()}

sent_email_list = "sent_emails.csv"
opts = get_opts()
config = get_config(opts.conference)
mailing_list = Path(opts.conference + "_mailing_list.xlsx")

if opts.mailing_list:
    mailing_list = Path(opts.mailing_list)
imagedir = Path(opts.imagedir)
sent_email_log = imagedir/sent_email_list

df = get_list(opts)
records = df.to_dict(orient= 'records')

# Do NOT send emails to successfully sent email addresses
already_sent = get_sent_emails()

# create a list of records to send emails to
# 1. is a .jpg
# 2. image name starts with CONFIRMATION NUMBER
# 3. when match is made, add that image path to the record
to_send_list = []
pattern = "*.jpg"
for f in imagedir.glob(pattern):  
    for r in records:
        if f.name.startswith(r['CONFIRMATION NUMBER']):
            r['IMAGE_PATH'] = f
            to_send_list.append(r)

if len(to_send_list) > 0:
    # get an SMTP object
    if not opts.dryrun:
        print("Connecting to", smtpserver, "on port", smtpport)
        smtp = smtplib.SMTP(smtpserver, smtpport)
        smtp.ehlo()
        ready = smtp.starttls()
        print(ready[1])

        pswd = getpass.getpass('Password:')
        try:
            status = smtp.login(send_from, pswd)
            print(str(status[1]))
        except smtplib.SMTPAuthenticationError as e:
            print("There was a problem authenticating:")
            print(str(e))
            sys.exit()

# pull the contents from the template file
with open(config['htmlfile'], 'r') as f:
        template = f.read()

# LOOP through rows and insert dynamic content
for idx, r in enumerate(to_send_list, start=1):
    print("--------------------------------------")
    print("Email", idx, "of", len(to_send_list))
    send_individual_email(r)

