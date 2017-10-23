# To use:
# Download Reporter from http://help.apple.com/itc/contentreporterguide/en.lproj/static.html#apda86f89da5
#       and move to desired download location
# Place this file inside the Reporter folder.
# Fill in the missing information below. (for help finding your account_number and vendor number see the above URL )
# Run the script!

# IMPORTS
import os
from datetime import date, timedelta
import subprocess
import gzip
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


# Enter Your Info:
base_file_path = "PATH TO YOUR REPORTER FOLDER"
account_number = "ACCOUNT NUMBER HERE"
vendor_number = "VENDOR NUMBER HERE"

me = "YOUR EMAIL HERE"
my_password = r"YOUR PASSWORD HERE"


# Checks if a key-value pair exists for the given key inside of the given dict. If it does it adds the value to the
# previous value, if not it adds the key-value pair to the dict.
def add_to_dict(dictionary, key, value):
    if key in dictionary:
        prev_total = dictionary[key]
        dictionary[key] = prev_total + value
    else:
        dictionary[key] = value


# returns the HTML string for a table representing the given dictionary, with the given title
def get_html_for_dict(dictionary, title):
    html_str = '<p><table style="width:100%"><caption>' + title + '</caption><tr><th>Name</th><th>Units</th></tr>'
    for key, value in dictionary():
        html_str += '<tr><td>' + key + "</td><td>" + str(value) + '</td></tr>'
    html_str += "</table></p><b /><b />"
    return html_str


# set current directory
os.system("cd " + base_file_path)

# get a string version of yesterdays date
yesterday = date.today() - timedelta(1)
yesterday_str = yesterday.strftime("%Y%m%d")

# request report from itc
output = subprocess.check_output("java -jar Reporter.jar p=Reporter.properties a=" + account_number +
                                 " Sales.getReport " + vendor_number + ", Sales, Summary, Daily, " + yesterday_str,
                                 shell=True)

# open downloaded itc report file and unzip it.
zip_filename = str(output.decode('UTF-8').split()[-1])
zip_ref = gzip.open(base_file_path + zip_filename, 'r')
file_contents = zip_ref.read().decode("UTF-8").split('\n')
zip_ref.close()

# parse contents
# get header row, find desired indices
headers = file_contents[0].split('\t')
title_index = headers.index('Title')
units_index = headers.index('Units')
download_type_index = headers.index('Product Type Identifier')

# drop header row
file_contents = file_contents[1:]

# iterate through data rows, getting desired data.
new_download_types = ["1", "1E", "1EP", "1EU", "1F", "1T", "F1"]
redownload_types = ["3", "3F", "3T", "F3"]
update_types = ["7", "7F", "F7"]
iap_types = ["IA1-M", "IA1", "IA9", "IA9-M", "IAC", "IAC-M", "IAY", "IAY-M", "FI1"]

new_downloads = dict()
redownloads = dict()
updates = dict()
iaps = dict()
for row in file_contents:
    row_array = row.split("\t")
    if len(row_array) < max(download_type_index, title_index, units_index):
        break

    product_title = row_array[title_index]
    units = int(row_array[units_index])

    download_type = row_array[download_type_index]

    # Downloads
    if download_type in new_download_types:
        add_to_dict(new_downloads, key=product_title, value=units)

    # redownloads
    if download_type in redownload_types:
        add_to_dict(redownloads, key=product_title, value=units)

    # Downloads
    if download_type in update_types:
        add_to_dict(updates, key=product_title, value=units)

    # Downloads
    if download_type in iap_types:
        add_to_dict(iaps, key=product_title, value=units)

# create email
message = get_html_for_dict(new_downloads, title='New Downloads')
message += get_html_for_dict(iaps, title='New IAPs')
message += get_html_for_dict(updates, title='New Updates')
message += get_html_for_dict(redownloads, title='Redownloads')

# skipped your comments for readability
msg = MIMEMultipart('alternative')
msg['Subject'] = "ITC Daily Downloads!"
msg['From'] = me
msg['To'] = me

html = '<html><head><style>table { font-family: arial, sans-serif; border-collapse: collapse; width: 100%;}' \
       'td, th {border: 1px solid #dddddd;text-align: left;padding: 8px;}' \
       'tr:nth-child(even) {background-color: #dddddd;}' \
       '</style></head><body>' + message + '</body></html>'
part2 = MIMEText(html, 'html')

msg.attach(part2)

# Send the message via gmail's regular server, over SSL - passwords are being sent, after all
s = smtplib.SMTP_SSL('smtp.gmail.com', 465)
# uncomment if interested in the actual smtp conversation
# s.set_debuglevel(1)
# do the smtp auth; sends ehlo if it hasn't been sent already
s.login(me, my_password)

s.sendmail(me, me, msg.as_string())
s.quit()
