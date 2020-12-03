# Before starting working on the code go to https://developers.google.com/sheets/api/quickstart/python
# in order to be familiar with the first part of the program since it was taken from there.

# Authentication Module Imports
from __future__ import print_function
import pickle
import os.path

from google.auth.transport import requests
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# Imports being use for Verify Sheet Module.
import random
import re
# always define the argument before using it.
auth = '0'

# The original header that should never be change unless told otherwise. The program should use this to determine if
# someone change the table by mistake or just did nor tell the others.
original_header = (
    'Room', 'User', 'Room Count', 'Computer Make', 'Computer Model', 'Specifications', 'DBT #', 'PO #',
    'OS Version', 'Serial ID', 'Room Type', 'Status', 'Verified?', 'Station Type', 'ACQ Date',
    'Lease End Date', 'Custodian or User', 'Notes')


# authentication is giving the program access to the google sheet.
def authentication():
    # If modifying these scopes, delete the file token.pickle.
    scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly']

    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    cred = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            cred = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not cred or not cred.valid:
        if cred and cred.expired and cred.refresh_token:
            cred.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', scopes)
            cred = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(cred, token)
    # calling verify function
    verify_sheet(cred)


# verify_sheet will go row by row checking that the data is where is suppose to be and if it's not look where it
# suppose to be at and move it there. The current code can only check each row and its value to see if they match
# what is suppose to be there based on the pattern found on each column. In order to move the column I found
# something that may be of some help but was not able to work on it so don't know if it will work.
# https://developers.google.com/sheets/api/videos There are multiple videos and a codelab(on the left of the page)
# that should be able to show how to move the columns.

# The Regular Expressions of the Specifications and Serial ID were writen to work around the Unicode since it shows up
# when the code searches the pattern.

# Ask about row 902 column Room(0), the meaning of ? that shows up in some parts of the table and what to do about it,
# there amy be a typo with the Latitude/Lattitude in Computer Model, and the Unicode that shows up in
# Specifications and Serial ID the only way I found to take them out is by hard typing
# the values because they came from copying/pasting.

def verify_sheet(auth_cred):
    # Defining it in order to add the current header of the table to it.
    current_document_header = []

    # The ID and range of a spreadsheet.
    # The ID is needed to be able to read and write data on the google sheet (can be find in the URL of the table).
    # The Range is the name of the sheet given bellow (Default name is Sheet 1) fallow by ! and the size of
    # the sheet (from left to right).
    sample_spreadsheet_id = '18C-kSKaZKNSOfovI8Gps041muvd8yayV75tzbdtzR3k'
    sample_range_name = 'Inventory Details!A1:R1'

    service = build('sheets', 'v4', credentials=auth_cred)

    # Call the Sheets API & Dump header from current inventory sheet.
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=sample_spreadsheet_id,
                                range=sample_range_name).execute()
    integrity_check = result.get('values', [])

    # If access was not granted it will not show the data on the table
    if not integrity_check:
        print('No data found.')
    else:
        for row in integrity_check:
            # Print columns A and R, which correspond to indices 0 through 19.
            print('%s' % (row[0:19]))
            # End of google api example ends

            # Give the values of the first row of the table to the variable current_document_header.
            current_document_header = row[0:19]

    # Confirming if the original header matches the current header if not let the user know in order to be check
    # by a employee
    if list(original_header) == current_document_header:
        print('All Elements are equal.')
    else:
        print('Error! It would seem that the document has been modified incorrectly.')
        # if they don't match show both the current and original in order to see how are they different and fix it.
        print(list(original_header))
        print(current_document_header)

    # Generate random numbers to get access to different row on the google sheet to verify the data
    # is in the correct place
    for x in range(2):

        # try catch in order for the program to continue running even if a error shows
        # if a new constant error show up like IndexError add an except with a message to let others know
        # what is making the error show up.
        try:
            # Generate random number between 1 and 1000
            # For testing purposes I did not make it the length of sheet feel free to change it to what works for you.
            # If a row would be giving problem I would just comment out the random generator and replace it with the row
            # giving problems to figure out it out and then uncomment out the random to see if it would find another
            # unexpected problem.
            random_samples = random.randint(1, 1000)  # make as the length of sheet

            # Access the correct row base on the number generated
            # In order to write the random number into the Range they need to be a string so use str on
            # the random_sample to add them where it needs to go to be able to access the row based
            # on the num generated.
            rows = sheet.values().get(spreadsheetId=sample_spreadsheet_id,
                                      range='Inventory Details!A' + str(random_samples) + ':R' +
                                            str(random_samples)).execute()
            # check is getting the values and the next part will use it to print the information it got.
            check = rows.get('values', [])
            for row in check:
                # This part will always display the full row with all data in it before doing anything else
                # and will also show the parts that are empty but only if it is in the columns 0-19 (A-R).
                print('%s' % (row[0:19]))
                print(random_samples)
                print()

            # In order to verify if the data is in the correct place we are using Regular expression (RE).
            # This website only gives some example but with all the commands. Study it and figure out how they work
            # before continuing on to the next part https://www.w3schools.com/python/python_regex.asp

            # Each if statement has the same structure what to do if it finds a pattern, if it's empty and if it did
            # not find any pattern and the column is not empty.

            # verify Room using the pattern given to identify it
            # The pattern is 2 capital letters follow by - and 3 numbers. Some columns are  differently are a bit
            # different ask about them or write a RE to be able to recognize it but the main pattern is the one given.
            # The | will separate each possible pattern that may be encounter in that column while the row[0] tells
            # the program what column to access int his case column A (0 or Room).
            if re.search("^[A-Z][A-Z]-[0-9][0-9][0-9]$|^[A-Z][A-Z]-[0-9][0-9][0-9] $|"
                         "^[A-Z][A-Z]-.*[0-9][0-9][0-9].*[A-Z]$", row[0]):

                # All RE are writen twice per search the first one is to see if its finding it in order to run the
                # if statement while the second is to show us what is getting to see if there is something that need to
                # be fix
                print(re.search("^[A-Z][A-Z]-[0-9][0-9][0-9]$|^[A-Z][A-Z]-[0-9][0-9][0-9] $|"
                                "^[A-Z][A-Z]-.*[0-9][0-9][0-9].*[A-Z]$", row[0]))
                print("Match in Room")
                print()

            # If the column is empty it will display a message that is empty.
            elif row[0] == '':
                print('No value in Column Room')
                print()

            # If no pattern was not found it will display an error message.
            else:
                print('Error in Room')
                print()

            # Verify the User using the pattern given yo identify it
            # In this case there is no pattern just names that can be inserted instead of possible patterns if
            # a new user is created just add a | with the name of the new user.
            if re.search("Classroom|Closet|Math|Has IPADS", row[1], re.IGNORECASE):
                print(re.search("Classroom|Closet|Math|Has IPADS", row[1], re.IGNORECASE))
                print("Match in User")
                print()
            # if User column is empty let the user know
            elif row[1] == '':
                print('No value in column User')
                print()
            else:
                print('Error in User')
                print()

            # verify Room Count using the pattern given to identify it
            # For the room count the current pattern only finds single and double digits since it's unlikely it goes to
            # triple digits
            if re.search("^[0-9][0-9]$|^[0-9]$", row[2]):
                print(re.search("^[0-9][0-9]$|^[0-9]$", row[2]))
                print('Match in Room Count')
                print()
            # if Room Count column is empty let the user know
            elif row[2] == '':
                print('No value in column Room Count')
                print()
            else:
                print('Error in Room Count')
                print()

            # verify Computer Make using the pattern given to identify it
            # Those are currents brands being use if any new brand shows up needs to be add to the list
            if re.search('Dell|Lenovo|Apple|HP', row[3]):
                print(re.search('Dell|Lenovo|Apple|HP', row[3]))
                print('Match in Computer Make')
                print()
            # if Computer Make column is empty let the user know
            elif row[3] == '':
                print('No value in column Computer Make')
                print()
            else:
                print('Error in Computer Make')
                print()

            # verify Computer Model using the pattern given to identify it
            # This pattern can be improve, for now it's only looking for some key words to be able to detect it.
            if re.search('Optiplex .*[0-9][0-9][0-9][0-9].*|Optiplex .*[0-9][0-9][0-9].*|Precision .*[0-9][0-9]|'
                         'Latitude [A-Z].*[0-9][0-9][0-9][0-9]|Latitude [0-9][0-9][0-9][0-9]|Mac [0-9][0-9].*|'
                         'iMac|iMac Pro .*|iPad|HP Compaq .*', row[4]):
                print(re.search('Optiplex .*[0-9][0-9][0-9][0-9].*|Optiplex .*[0-9][0-9][0-9].*|'
                                'Precision .*[0-9][0-9]|Latitude [A-Z].*[0-9][0-9][0-9][0-9]|'
                                'Latitude [0-9][0-9][0-9][0-9]|iMac [0-9][0-9].*|iMac|Mac Pro .*|'
                                'iPad|HP Compaq .*', row[4]))
                print('Match in Computer Model')
                print()
            # if computer model column is empty let user know
            elif row[4] == '':
                print('No value in column Computer Model')
                print()
            else:
                print('Error in Computer Model')
                print()

            # verify Specifications using the pattern given to identify it
            # This is one of the columns that display the info with unicode which messes up the search and because of
            # that this search was build to work around that. It is working but can also be improve.
            if re.search('.*Hz.*|.*gHz.*hd|.*gHz.*gb|.*ghz.*tb', row[5], re.IGNORECASE):
                print(re.search('.*Hz.*|.*gHz.*hd|.*gHz.*gb|.*gHz.*tb', row[5], re.IGNORECASE))
                print('Match in Specifications')
                print()
            # if Specifications column is empty let user know
            elif row[5] == '':
                print('No value in column Specification')
                print()
            else:
                print('Error in Specifications')
                print()

            # verify DBT # using the pattern given to identify it
            #
            if re.search('^[A-Z][0-9][0-9][0-9][0-9] / [A-Z][0-9][0-9][0-9][0-9]|'
                         '^[A-Z][0-9][0-9][0-9][0-9]/[A-Z][0-9][0-9][0-9][0-9]|'
                         '^[A-Z][0-9][0-9][0-9][0-9]$|^[0-9][0-9].*[0-9][0-9]|'
                         '^[A-Z][0-9][0-9][0-9][0-9][0-9]$', row[6]):
                print(re.search('^[A-Z][0-9][0-9][0-9][0-9] / [A-Z][0-9][0-9][0-9][0-9]|'
                                '^[A-Z][0-9][0-9][0-9][0-9]/[A-Z][0-9][0-9][0-9][0-9]|'
                                '^[A-Z][0-9][0-9][0-9][0-9]$|^[0-9][0-9].*[0-9][0-9]|'
                                '^[A-Z][0-9][0-9][0-9][0-9][0-9]$', row[6]))
                print('Match in DBT #')
                print()
            # if DBT # column is empty let the user know
            elif row[6] == '':
                print('No value in column DBT #')
                print()
            else:
                print('Error, in DBT #')
                print()

            # verify PO# using the pattern given to identify it
            # For this one it only looks for number but need to be careful in case a new pattern that may include
            # something other things than just numbers
            if re.search('Purchased|^[0-9][0-9][0-9][0-9][0-9]$|^[0-9][0-9][0-9][0-9][0-9][0-9]$|'
                         '^[0-9][0-9][0-9][0-9]$|^[0-9][0-9][0-9][0-9][0-9][0-9][0-9]$', row[7]):
                print(re.search('Purchased|^[0-9][0-9][0-9][0-9][0-9]$|^[0-9][0-9][0-9][0-9][0-9][0-9]$|'
                                '^[0-9][0-9][0-9][0-9]$|^[0-9][0-9][0-9][0-9][0-9][0-9][0-9]$', row[7]))
                print('Match in PO #')
                print()
            # if PO# column is empty let user know
            elif row[7] == '':
                print('No value in column PO #')
                print()
            else:
                print('Error in PO #')
                print()

            # verify OS Version using the pattern given to identify it
            # This one works like the Model it looks for some key words in order to get all the info
            if re.search('^Windows [0-9][0-9]|^Windows [0-9]|^OS .*[a-z]|[0-9].[0-9].[0-9]|'
                         '^Windows XP|^Windows[0-9][0-9]|^Windows[0-9]', row[8]):
                print(re.search('^Windows [0-9][0-9]|^Windows [0-9]|^OS .*[a-z]|[0-9].[0-9].[0-9]|'
                                '^Windows XP|^Windows[0-9][0-9]|^Windows[0-9]', row[8]))
                print('Match in OS Version')
                print()
            # if OS Version column is empty let user know
            elif row[8] == '':
                print('No value in column OS Version')
                print()
            else:
                print('Error in  OS Version')
                print()

            # verify Serial ID
            # This is the other one that display the unicode but is does not give trouble compare
            # with the specifications.
            # This is the last one I was working on it's not finish. I would recommend to rewrite it is RE.
            # The pattern is hard to see but there is one.
            if re.search('[A-Z].*[0-9].*[A-Z].*[0-9]|[0-9].*[0-9].*[A-Z].*[A-Z]|'
                         '[A-Z].*[0-9].*[A-Z].*[A-Z]|[0-9].*[0-9].*[A-Z].*[0-9]|'
                         '[0-9].*[A-Z].*[A-Z].*[0-9]|[A-Z].*[A-Z].*[A-Z].*[0-9]|'
                         '[A-Z].*[0-9].*[0-9].*[0-9]|'
                         '[A-Z][0-9].*[A-Z].*[0-9]|[0-9][0-9].*[A-Z].*[A-Z]|'
                         '[A-Z][0-9].*[A-Z].*[A-Z]|[0-9][0-9].*[A-Z].*[0-9]|'
                         '[0-9][A-Z].*[A-Z].*[0-9]|[A-Z][A-Z].*[A-Z].*[0-9]|'
                         '[A-Z][0-9].*[0-9].*[0-9]', row[9]):
                print(re.search('[A-Z].*[0-9].*[A-Z].*[0-9]|[0-9].*[0-9].*[A-Z].*[A-Z]|'
                                '[A-Z].*[0-9].*[A-Z].*[A-Z]|[0-9].*[0-9].*[A-Z].*[0-9]|'
                                '[0-9].*[A-Z].*[A-Z].*[0-9]|[A-Z].*[A-Z].*[A-Z].*[0-9]|'
                                '[A-Z].*[0-9].*[0-9].*[0-9]|'
                                '[A-Z][0-9].*[A-Z].*[0-9]|[0-9][0-9].*[A-Z].*[A-Z]|'
                                '[A-Z][0-9].*[A-Z].*[A-Z]|[0-9][0-9].*[A-Z].*[0-9]|'
                                '[0-9][A-Z].*[A-Z].*[0-9]|[A-Z][A-Z].*[A-Z].*[0-9]|'
                                '[A-Z][0-9].*[0-9].*[0-9]', row[9]))
                print('Match in Serial ID')
                print()

            elif row[9] == '':
                print('No value in column Serial ID')
                print()
            else:
                print('Error in Serial ID')
                print()

            # verify Room Type using the pattern given to identify it
            # May need to be updated base on if a new room is added or taken out.
            if re.search('Lab|Testing|Staff|Lecture|Office', row[10]):
                print(re.search('Lab|Testing|Staff|Lecture|Office', row[10]))
                print('Match in Room Type')
                print()
            # if Room Type column is empty let user know
            elif row[10] == '':
                print('No value in column Room Type')
                print()
            else:
                print('Error in Room Type')
                print()

            # verify Status using the pattern given to identify it
            # Ask to see if there is a different way they would put this to be on the safe side.
            if re.search('Active|Not Active', row[11]):
                print(re.search('Active|Not Active', row[11]))
                print('Match in Status')
                print()
            # if Status column is empty let user know
            elif row[11] == '':
                print('No value in column Status')
                print()
            else:
                print('Error in Status')
                print()

            # verify Verified? using the patter given to identify it
            # Ask in order to know if it's not verified it will be blank or they plan to classify it as 'not verified'
            if re.search('Verified', row[12]):
                print(re.search('Verified', row[12]))
                print('Match in Verified?')
                print()
            # if Verified? column is empty let the user know
            elif row[12] == '':
                print('No value in column Verified/Not Verified')
                print()
            else:
                print('Error in Verified?')
                print()

            # verify Station Type using the pattern given to identify it
            # May need to be updated base on if a new station is added or taken out.
            if re.search('Student|Instructor|Server|Staff|Bunker', row[13]):
                print(re.search('Student|Instructor|Server|Staff|Bunker', row[13]))
                print('Match in Station Type')
                print()
            # if Station Type column is empty let user know
            elif row[13] == '':
                print('No value in column Station Type')
                print()
            else:
                print('Error in Station Type')
                print()

            # verify ACQ Date using the pattern given to identify it
            # I thought iw will be asy for this one to be able to see the pattern it looking for ratter than saying it.
            #               xx/xx/xxxx                              x/xx/xxxx
            if re.search('^[0-1][0-2]/[0-3][0-9]/[0-2]0[0-9][0-9]$|^[0-9]/[0-3][0-9]/[0-2]0[0-9][0-9]$|'
            #               xx/x/xxxx                               x/x/xxxx
                         '^[0-1][0-2]/[0-9]/[0-2]0[0-9][0-9]$|^[0-9]/[0-9]/[0-2]0[0-9][0-9]$', row[14]):

                print(re.search('^[0-1][0-2]/[0-3][0-9]/[0-2]0[0-9][0-9]$|^[0-9]/[0-3][0-9]/[0-2]0[0-9][0-9]$|'
                                '^[0-1][0-2]/[0-9]/[0-2]0[0-9][0-9]$|^[0-9]/[0-9]/[0-2]0[0-9][0-9]$', row[14]))
                print('Match in ACQ Date')
                print()
            # if ACQ Date column is empty let user know
            elif row[14] == '':
                print('No value in column ACQ Date')
                print()
            else:
                print('Error in ACQ Date')
                print()

            # verify Lease End Date using the pattern given to identify it
            # For this ir looking for a month abbreviation with the first letter capitalize a - and two digits number
            # or it is was purchased or leased
            if re.search('Purchased|Leased|[A-Z][a-z][a-z]-[0-9][0-9]', row[15]):
                print(re.search('Purchased|Leased|[A-Z][a-z][a-z]-[0-9][0-9]', row[15]))
                print('Match in Lease End Date')
                print()
            # if Lease End Date column is empty let user know
            elif row[15] == '':
                print('No value in column Lease End Date')
                print()
            else:
                print('Error, Lease End Date')
                print()

            #row.insert(18, 'hello')
            #print(row[18])
            #data = {'values': [rows[:18] for rows in row]}
            #print(row)
            #sheet.values().update(spreadsheetId=sample_spreadsheet_id, range='A1', body=data).execute()
        except IndexError:
            print('Multiple columns are empty one after another until the end of the table'
                  ' and is making the program run into errors')
            continue
        except:
            print('Something else went wrong')


# After finishing the verify this is the next step if nothing changes between when it was written and now.
def scan_sheet():
    print('Sheet Verification goes here')


authentication()


# From this point on is the first version of the code that was later change to the one above. This is here in case there
# is a need to reference it.
'''

# Editing Here
    wrong_header = [original_header[4], original_header[9], ]

    current_document_header = []

    example2 = input('string that came in from document')
    example = example2.split()

    print('Sheet Verification goes here')
    if list(original_header) == current_document_header:
        print('All Elements are equal.')
    else:
        print('Error! It would seem that the document has been modified incorrectly. Attempting to revert the document')
    #if content



# End Editing for now

#####################################################################################################################

# Section 2 Editing


sheet_b_dbt = 0
sheet_b_room = 0

# Open Documents
sheet_main_a = client.open("Main(A)").sheet1
sheet_inventory_b = client.open("Inventory(B)").sheet1

# Search for the column with the name DBT#
for column in range(1, 5):
    DBTLoc = sheet_inventory_b.column(1, column).value
    if DBTLoc == "DBT #":
        sheet_b_dbt = column
        var1 = 5

# Search for the column with the name Room`
for column in range(1, 5):
    DBTLoc = sheet_inventory_b.column(1,column).value
    if DBTLoc == "Room ":
        sheet_b_room = column

print(var1)
# test value in Inventory sheet / DBT#
# the number 6 is the row and the number 2 is the column
scan_sheet_b = sheet_inventory_b.column(6, sheet_b_dbt).value
scan_sheet_b2 = sheet_inventory_b.column(6, sheet_b_room).value


# scan the Main sheet to see if it as the same DBT# on both sheets
# the range represents the row and the number 7 represents the column AKA the DBT# in main sheet
# if both sheets have the same DBT# the program will print found
for rows in range(2, 8):
    dbt_sheet_a = sheet_main_a.column(rows,7).value
    if scan_sheet_b == dbt_sheet_a:  # Compare Scan sheet to DBT's on Sheet A
        # logger.info('Found a match')
        print('found')  # Can we print information to a logfile rather than to the screen ex. print(DBT, 'Matches ')
        RoomNum = sheet_main_a.column(rows, 1).value

        if scan_sheet_b2 == RoomNum:
            print('same')

'''
