"""headshot_emailer.py
    Reads a spreadsheet of names, attendee numbers
    (which are also image names), and email addresses.
    Uses yagmail, a library specific for interacting with
    gmail's SMTP server, to send a specific, custom email
    and attachment to each email address.
"""
import yagmail
import pandas as pd

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

    to_email = r['EMAIL']

    # if we already sent it, don't do it
    if to_email in already_sent and \
       already_sent[to_email]['ATTENDEE_NUM'] == r['ATTENDEE_NUM']:
        print(to_email, "has already received", r['IMAGE_PATH'])
        return

    subject='Prototype Automated Email for Sending Headshots'
    content = template.format(FIRST=r['FIRST'])
    print("To:", to_email)
    print("Subject:", subject)
    print("Attachment:", r['IMAGE_PATH'])

    if opts.dryrun:
        print("(dry run, no email sent)")
    else:
        # attach the image as an attachment
        yag.send(
             to = r['EMAIL'],
             subject=subject,
             attachments=str(r['IMAGE_PATH']),
             contents=content
        )
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
for r in records:
    img_name = r['ATTENDEE_NUM'] + '.jpg'
    r['IMAGE_PATH'] = imagedir/img_name

# get an SMTP object
if not opts.dryrun:
    yag = yagmail.SMTP('ben@benklaas.com')

# pull the contents from the template file
with open("simple.html", 'r') as f:
        template = f.read()

# Do NOT send emails to successfully sent email addresses
already_sent = get_sent_emails()

# for those that need sending, confirm every image exists
for r in records:
    if r['EMAIL'] not in already_sent:
        if not r['IMAGE_PATH'].exists():
            warn_and_exit("Missing image:" + str(r['IMAGE_PATH']))

# LOOP through rows and insert dynamic content
for idx, r in enumerate(records, start=1):
    print("--------------------------------------")
    print("Email", idx, "of", len(records))
    send_individual_email(r)

