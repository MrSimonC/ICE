import csv
import os
import datetime
import time
import sys
from custom_modules.outlook import Outlook


def email_out(row, html_file, subject, attach_folder, from_address):
    # Mass email out to lots of people, one at a time.
    raw_html = open(html_file, 'r').read()
    html = raw_html.replace('$firstname', row['firstName'])
    html = html.replace('$surname', row['surname'])
    html = html.replace('$username', row['username'])
    html = html.replace('$email', row['email'])
    attachments = []
    if os.path.exists(attach_folder):
        attachments = [os.path.join(os.getcwd(), attach_folder, fn) for fn in os.listdir(attach_folder)]
    o = Outlook()
    o.send(True, row['email'], subject, '', html, attachments=attachments, account_to_send_from=from_address)


def email_individually_from_file(csv_path, out_file, html_file, subject, attach_folder, from_address):
    users_to_email = csv.DictReader(open(csv_path))
    for no, row in enumerate(users_to_email):
        print('---\nEmailing ' + str(no) + ': ' + row['firstName'] + ' ' + row['surname'] + ', ' + row['email'])
        email_out(row, html_file, subject, attach_folder, from_address)
        csv_file_write = open(out_file, 'a', newline='')
        writer = csv.DictWriter(csv_file_write, users_to_email.fieldnames)
        comment = 'Emailed {0} {1} , {2} at '.format(row['firstName'], row['surname'], row['email'])
        row['Status'] = comment + datetime.datetime.now().strftime('%d %b %Y %H:%M')  # 01 Jan 1900 19:00
        writer.writerow(row)
        csv_file_write.close()
        time.sleep(0.1)  # be nice to the poor email server


csv_path = 'ICE_to_email.csv'
out_file = 'ICE_email_output.csv'
html_file = 'mainMessage.htm'
subject = 'VPLS switch off - requirement for NBT ICE Account'
attach_folder = 'mainAttachments'
from_address = 'PathICEAccounts@nbt.nhs.uk'

if hasattr(sys, "frozen"):
    containing_folder = os.path.dirname(sys.executable)
else:
    containing_folder = os.path.dirname(os.path.realpath(__file__))

csv_path = os.path.join(containing_folder, csv_path)
out_file = os.path.join(containing_folder, out_file)
html_file = os.path.join(containing_folder, html_file)
attach_folder = os.path.join(containing_folder, attach_folder)
from_address = os.path.join(containing_folder, from_address)

email_individually_from_file(csv_path, out_file, html_file, subject, attach_folder, from_address)

# Compile:
# from custom_modules.compile_helper import CompileHelp
# c = CompileHelp(r'C:\simon_files_compilation_zone\ICE_email')
# c.create_env()
# c.freeze(r'K:\Coding\Python\nbt work\outlook_email_many.py', copy_to=r'\\nbsvr175\Scripts\ice\outlook_email_many.exe')