import gspread
from oauth2client.service_account import ServiceAccountCredentials
from tinydb import TinyDB, Query
from prettytable import PrettyTable
from datetime import date

"""IMPORT FROM GOOGLE SHEETS"""

# List provides the necessary links to connect to the API
scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

# Function opens the Json file named parade state
creds = ServiceAccountCredentials.from_json_keyfile_name("parade-state-350612-4de42593bc63.json", scope)

# This give authority for using the file and opening the API
client = gspread.authorize(creds)

# Opens the google-docs file named 'Parade State'
sheet = client.open("Parade state").sheet1

# Gets all the data in the Google Sheets
data = sheet.get_all_records()

# The variables needed when managing and importing the Tiny Database
User = Query()
db = TinyDB('db.json')

"""DATA STRUCTURES"""
""" Dictionary will be used to store data for the name and their statues. Choosing Dictionary also allow this
system to remove any duplicated names. Users that keyed the same name twice but with a different status, will
update his status with the latest submitted one """
current_dict = {}

# The pretty table module that display the parade state in a properly formatted table
ps_Table = PrettyTable(["No.", "Personnel", "status"])
sum_Table = PrettyTable(["Status", "Total Number"])
""" This helps to format the raw data into a understandable dictionary called current_dict.
It is done by getting data from the first and second question """
for personnel in data:
    current_dict[personnel.get('Full Name (Name keyed in the database)')] = personnel.get('Status')

# This list will be used to display the parade state
display = []
# dictionary helps to provide the summary of the parade state
present_day = {"Present": [], "Work From Home": [], "Outside Stationed": [], "Attached Out": [],
               "On Course": [], "Day Off": [], "Local Leave": [],
               "Overseas Leave": [], "Medical Leave": [], "Medical/Dental Appointment": [],
               "Report Sick Inside": [], "Report Sick Outside": [],
               "Hospitalized / Sickbay": [], "AWOL": [], "OTHERS": []}


"""DATABASE FUNCTIONS"""


# function to enter details into the database
def insert(rank, name):
    db.insert({'rank': rank, 'name': name})


# function to remove a personnel details from the database
def delete_by_name(name):
    db.remove(User.name == name)


# Function ask user to enter the details of the personnel to add into the database
def add_database():
    name: str = str(input("what\'s your name: "))
    rank = str(input("what\'s your rank: "))
    insert(rank.lower(), name.lower())


# Function ask user to enter the name of the personnel to remove from database
def remove_by_name():
    name = str(input('what\'s your name: '))
    delete_by_name(name.lower())


"""SUBMITTING AND CALCULATING PARADE STATE FUNCTIONS"""


def submit():
    number = 1
    for key, value in current_dict.items():
        result = db.search(User.name == key)
        for match in result:
            ps_Table.add_row([str(number), str(match.get('rank').upper() + ' ' + match.get('name').upper()),
                              str(value)])
            present_day[value] += [match.get('name')]
            number += 1


def total_strength():
    strength = 0
    absent = 0
    item = list(present_day.keys())
    for word in item[0:5]:
        strength += len(present_day.get(word))
    for word in item[5:]:
        absent += len(present_day.get(word))
    sum_Table.add_row(['TOTAL STRENGTH: ', str(strength)])
    sum_Table.add_row(['TOTAL ABSENTEES: ', str(absent)])
    print(sum_Table)


"""DISPLAYING FUNCTIONS"""


def display_ps():
    print('Parade State ' + str(date.today()) + "\nTotal Present: " + str(len(present_day['Present']))+'/'+str(len(db)))
    print(ps_Table)


def display_sum():
    print('Parade State Summary ' + str(date.today()))
    for key, value in present_day.items():
        sum_Table.add_row([str(key), str(len(value))])
    total_strength()


"""MAIN FUNCTION"""

while True:
    question = str(input('Add or Remove personnel in Database/Submit Parade State (D/S): '))
    if question == 'D':
        edit = str(input('(Add/Remove) from database?: '))
        if edit == 'Add':
            while True:
                add_database()
                ask = str(input('Do you still want to add?(Yes/No): '))
                if ask == 'No':
                    break
        elif edit == 'Remove':
            remove_by_name()
            break
        else:
            print("Please key in from the following options mentioned above")
    elif question == 'S':
        submit()
        print('Parade State has been submitted')
        while True:
            second = str(input('View Parade State or Summary (P/S): '))
            if second == 'P':
                display_ps()
            elif second == 'S':
                display_sum()
            else:
                break
    else:
        last_check = str(input('Before leaving, would you like to deleted today\'s parade state? (Y/N): '))
        if last_check == 'Y':
            count = len(data)
            sheet.delete_rows(2, count + 1)
            print('Parade State deleted')
            print('Thank you and have a nice day')
            break
        else:
            if last_check == 'N':
                print('Thank you and have a nice day')
                break
