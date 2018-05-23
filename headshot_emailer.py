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

mailing_list = "mailing_list.csv"

def warn_and_exit(msg_list):
    print("ERROR!", msg_list)
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
    if not Path(opts.mailing_list).exists():
        warn_and_exit("Could not find", str(Path(opts.mailing_list)))

    df = pd.read_csv(opts.mailing_list, dtype=str)
    df.columns = [col.upper() for col in df.columns]
    cols_needed = ['ATTENDEE_NUM', 'LAST', 'FIRST', 'EMAIL']

    all_there = all(col in df.columns for col in cols_needed)
    if not all_there:
        warn_and_exit("Missing column in your spreadsheet. Required columns:", cols_needed)
    return df[cols_needed]

opts = get_opts()
df = get_list(opts)

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

