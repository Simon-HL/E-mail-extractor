import re
import os
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def getFile():
    # # Snippet to prompt the user for file and allow them to browse directory for path
    Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
    return askopenfilename()  # show an "Open" dialog box and return the path to the selected file

def getPath(file):
    tup = os.path.split(os.path.abspath(file)) #returns a tuple with path and filename
    path = tup[0]
    orgfilename=tup[1].split('.')[0] #removes file extension
    return (path, orgfilename)

#TODO: Sjekk om det er csv eller txt fil, eller noe annet. Try (csv||txt), else raise
def fileToString(file):
    #Converts an entire txt/csv file into string
    with open(file) as f:
        return f.read().lower()

def sortOutEmail(listofemails):
    #TODO: Split into username, company, countrycode etc. Need to find a ogod way to handle domains like .net, pop3 and others.
    # altcountry=''
    # listofcompanies = []
    # listofcountries = []

    listofusers = []
    listofdomains = []

    for email in listofemails:
        (user, domain) = email.split('@')
        listofusers.append(user)
        listofdomains.append(domain)
    return listofusers, listofdomains
        # try:
        #     (company, country) = domain.split('.')
        # except ValueError:
        #     (company, country, altcountry) = domain.split('.')
        # if altcountry != '':
        #     country = country + altcountry
    #     listofusers.append(user)
    #     listofcompanies.append(company)
    #     listofcountries.append(country)
    # return (listofusers, listofcompanies, listofcountries)

def extractEmails(line):
    addresses = re.findall(r'[\w\.-]+@[\w\.-]+', line)
    return addresses
    # for address in addresses:
    #     print(address)

def lookupCountry(countrycode):
    # TODO: Implement table for comparing domain to country to output the country of the email
    #Use dictionaries?
    pass

def lookupCompany(companyabbr):
    # TODO: Implement table for comparing company abbreviation to company to output the company of the email
    #Use dictionaries?
    pass

def saveToExcel(dataframe, path):
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(path, engine='xlsxwriter')
    # Convert the dataframe to an XlsxWriter Excel object.
    dataframe.to_excel(writer, index=False, sheet_name='List of extracted emails')

    #Format content
    worksheet = writer.sheets['List of extracted emails']
    worksheet.set_column('A:A', 53)
    worksheet.set_column('B:B', 39)
    worksheet.set_column('C:C', 26)

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()

file = getFile()
(path, filename) = getPath(file)

#TODO: Use Generator to read and analyse CSV file in chunks to save memory in case of big files
#Reads entire CSV file to a string
text=fileToString(file)

eMailAddresses = (extractEmails(text))
usernames, domains = sortOutEmail(eMailAddresses)

df = pd.DataFrame({'E-mail address': eMailAddresses, 'Username':usernames,'Domain':domains})
df = df.drop_duplicates(subset=['E-mail address'])
df = df.sort_values('Domain', axis=0) #Sort by domain to group company emails together
df = df[['E-mail address','Username','Domain']] #Sort columns to get emails first

newFileName =path+'/' 'Extracted E-mails_' + filename + '.xlsx'
saveToExcel(df,newFileName)

#TODO: Legg til sorteringsmulighet i excel når man lagrer til excel.
#TODO: Filter out do-not-reply? or postmaster? post? sales etc.?