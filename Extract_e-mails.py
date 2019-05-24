import re
import os
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename


def get_file():
    # Prompt the user for file and allow them to browse directory for path
    Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
    return askopenfilename()  # show an "Open" dialog box and return the path to the selected file


def get_path(file):
    tup = os.path.split(os.path.abspath(file))  # returns a tuple with path and filename
    path = tup[0]
    org_filename = tup[1].split('.')[0]  # removes file extension
    return path, org_filename


def file_to_string(file):
    # TODO: Check whether file is csv, txt, or something else. Try (csv||txt), else raise
    # Converts an entire txt/csv file into string
    with open(file) as f:
        return f.read().lower()


def sort_out_email(list_of_emails):
    # TODO: Split into username, company, country etc. Need to find a good way to handle domains like .net, pop3 etc.
    list_of_users = []
    list_of_domains = []
    for email in list_of_emails:
        (user, domain) = email.split('@')
        list_of_users.append(user)
        list_of_domains.append(domain)
    return list_of_users, list_of_domains


def extract_emails(line):
    addresses = re.findall(r'[\w\.-]+@[\w\.-]+', line)
    return addresses


def save_to_excel(dataframe, path):
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(path, engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object.
    dataframe.to_excel(writer, index=False, sheet_name='List of extracted emails')

    # Format content
    worksheet = writer.sheets['List of extracted emails']
    worksheet.set_column('A:A', 53)  # Last parameter sets the width of the column for better presentation
    worksheet.set_column('B:B', 39)
    worksheet.set_column('C:C', 26)

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()


if __name__ == '__main__':
    print("Loading... Preparing to select file")
    file = get_file()
    (path, filename) = get_path(file)
    print("Processing...")

    # TODO: Use Generator to read and analyse CSV file in chunks to save memory in case of big files
    # Reads entire CSV file to a string
    text = file_to_string(file)

    email_addresses = (extract_emails(text))
    usernames, domains = sort_out_email(email_addresses)

    df = pd.DataFrame({'E-mail address': email_addresses, 'Username': usernames, 'Domain': domains})
    df = df.drop_duplicates(subset=['E-mail address'])
    df = df.sort_values('Domain', axis=0)  # Sort by domain to group company emails together
    df = df[['E-mail address', 'Username', 'Domain']]  # Sort columns to get emails first

    new_filename = path + '/' 'Extracted E-mails_' + str(filename) + '.xlsx'
    save_to_excel(df, new_filename)
    print("Complete!")

    # TODO: Add sorting feature to excel sheet when saving to excel
    # TODO: Filter out do-not-reply? or postmaster? post? sales etc.?
