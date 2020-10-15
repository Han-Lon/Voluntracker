import tkinter
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import httplib2
import csv
from PIL import Image
from PIL import ImageTk
from tkinter.filedialog import askopenfilename
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from openpyxl import load_workbook
import platform


# A GUI-based application for managing volunteer hours for the Delta Sigma Pi fraternity. Pull backups from an
# online Google Sheets csv file, save them to your local drive, edit the membership roster, view organization
# metrics, and format a backup file into a submit-ready Excel spreadsheet.

# TODO When running PyInstaller, use this terminal command -> pyinstaller --onefile --windowed --icon=[link_to_,ico_image] --clean Voluntracker.py

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# Holds the link to the spreadsheet (a unique identifier from the URL) and the range we want to pull
SpreadURL = ''
SAMPLE_RANGE_NAME = 'B:F'

# Paths to all of the necessary files required for Voluntracker to run properly
# Default path syntax is Windows/DOS specific - there is a check in __main__ that will change paths to appropriate
# Linux/Mac OS/Unix syntax if executed on a non-Windows OS
BACKUP_PATH = 'Backups'

ROSTER_PATH = 'Configuration\\Roster.csv'

LOGO_PATH = 'Configuration\\DSP_logo.png'

PICKLE_PATH = 'Configuration\\token.pickle'

CREDS_PATH = 'Configuration\\credentials.json'

SPRERL_PATH = 'Configuration\\spreadurl.txt'

TEMPLATE_URL = 'Configuration\\BaseTemplate.xlsx'

# Global variable to hold values from the spreadsheet
values = []


# Handles creation of the initial window that the user sees upon program start
class MainWindow(tkinter.Tk):
    def __init__(self):
        tkinter.Tk.__init__(self)
        self.geometry("500x400")
        self.title("Voluntracker")

        # Imports the image to use on the top of the screen
        raw_logo = Image.open(LOGO_PATH)
        logo = ImageTk.PhotoImage(raw_logo)

        photo_label = tkinter.Label(self, image=logo, bg='white', width=self.winfo_screenwidth())
        photo_label.image = logo

        toolbar = tkinter.Frame(master=self, bg='#228f54', width=self.winfo_screenwidth(), height=2)
        toolbar_label = tkinter.Label(toolbar, width=toolbar.winfo_screenwidth(), bg='#228f54', fg='white', height=2)

        # For spacing out the widgets
        label2 = tkinter.Label(self, width=self.winfo_screenwidth(), height=2)

        # The welcome message
        greet_label = tkinter.Label(self, text='Welcome to Voluntracker, Executive!', font=('Sans-Serif'))

        # All of the buttons on the toolbar
        backupbutton = tkinter.Button(toolbar, text='Pull Backups', command=create_backup_window,
                                      height=2, width=15)

        rosterbutton = tkinter.Button(toolbar, text='Check Roster', command=create_roster_window,
                                      height=2, width=15)

        metricbutton = tkinter.Button(toolbar, text="Metrics", command=create_metrics_window, height=2, width=15)

        submitbutton = tkinter.Button(toolbar, text="Finalize Submission", command=create_submission_window,
                                      height=2, width=15)

        # For identifying the current spreadsheet's URL and allowing the user to change it
        url_text = 'Currently using the spreadsheet at: ' + SpreadURL
        url_label = tkinter.Label(self, text=url_text)
        url_button = tkinter.Button(self, text='Change Spreadsheet', command=lambda: change_url(self),
                                    height=2, width=15)

        # Places all of the widgets on the tkinter window
        photo_label.pack()
        toolbar.pack()
        toolbar_label.pack()
        label2.pack()
        backupbutton.pack(side='left', padx=2, pady=2)
        rosterbutton.pack(side='left', padx=2, pady=2)
        metricbutton.pack(side='left', padx=2, pady=2)
        submitbutton.pack(side='left', padx=2, pady=2)
        greet_label.pack(side='top', padx=2, pady=2)
        url_button.pack(side='bottom', padx=2, pady=2)
        url_label.pack(side='left', padx=2, pady=2)
        self.mainloop()

    # For displaying new info on a tkinter window
    def refresh(self):
        self.destroy()
        self.__init__()


# Handles creation of all subsequent windows triggered from the main window
class NewWindow(tkinter.Toplevel):
    def __init__(self):
        tkinter.Toplevel.__init__(self)
        self.geometry("800x600")
        self.title("Voluntracker")

        # Imports the image to use on the top of the screen
        raw_logo = Image.open(LOGO_PATH)
        logo = ImageTk.PhotoImage(raw_logo)

        photo_label = tkinter.Label(self, image=logo, bg='white', width=self.winfo_screenwidth())
        photo_label.image = logo

        toolbar = tkinter.Frame(master=self, bg='#228f54', width=self.winfo_screenwidth())
        toolbar_label = tkinter.Label(toolbar, width=toolbar.winfo_screenwidth(), bg='#228f54', fg='white')

        # For spacing out the widgets
        label2 = tkinter.Label(self, width=self.winfo_screenwidth(), height=1)

        # All of the buttons on the toolbar
        backupbutton = tkinter.Button(toolbar, text='Pull Backups', command=create_backup_window,
                                      height=2, width=15)

        rosterbutton = tkinter.Button(toolbar, text='Check Roster', command=create_roster_window,
                                      height=2, width=15)

        metricbutton = tkinter.Button(toolbar, text="Metrics", command=create_metrics_window, height=2, width=15)

        submitbutton = tkinter.Button(toolbar, text="Finalize Submission", command=create_submission_window,
                                      height=2, width=15)

        # Places all of the widgets on the tkinter window
        photo_label.pack()
        toolbar.pack()
        toolbar_label.pack()
        label2.pack()
        backupbutton.pack(side='left', padx=2, pady=2)
        rosterbutton.pack(side='left', padx=2, pady=2)
        metricbutton.pack(side='left', padx=2, pady=2)
        submitbutton.pack(side='left', padx=2, pady=2)


# Creates a backup of the linked Google Sheet
def pull_backup():
    global values
    global BACKUP_PATH

    if platform.system() == 'Windows':
        slash = '\\'
    else:
        slash = '/'

    # The backups directory currently exists -- create a new backup file
    if os.path.exists(BACKUP_PATH):
        # Reads files currently in the Backups folder in order to determine new filename
        cur_files = 0
        recent_file = ''

        with os.scandir(path=BACKUP_PATH) as directory:
            for file in directory:
                cur_files += 1
                recent_file = file.name.split('p', 1)[0]

        if cur_files == 0:
            new_backup = 'hours_backup.csv'
        else:
            new_backup = recent_file + 'p' + str(cur_files) + '.csv'

        # Writes the results obtained from create_backup_window to a csv file in the given directory
        with open(BACKUP_PATH + slash + new_backup, 'w+', newline='\n', encoding='utf-8') as csvfile:
            csvfile.truncate()
            valuewriter = csv.writer(csvfile)
            for item in values:
                valuewriter.writerow(item)
    # The backups directory does not exist -- Create a new one and save an initial backup file
    else:
        os.mkdir(BACKUP_PATH)
        # Writes the results obtained from create_backup_window to a csv file in the given directory
        with open(BACKUP_PATH + slash + 'hours_backup.csv', 'w+', newline='\n', encoding='utf-8') as csvfile:
            csvfile.truncate()
            valuewriter = csv.writer(csvfile)
            for item in values:
                valuewriter.writerow(item)


# Creates a new window for pulling the current data in Google Sheets and creating backups from it
def create_backup_window():
    # Google API login and setup
    creds = None
    global values

    if(SpreadURL == 'NO CURRENT SPREADSHEET SELECTED'):
        create_error_window('ERROR: Please input a valid Google Sheets URL first')
    else:
        try:
            # The file token.pickle stores the user's access and refresh tokens, and is
            # created automatically when the authorization flow completes for the first
            # time.
            if os.path.exists(PICKLE_PATH):
                with open(PICKLE_PATH, 'rb') as token:
                    creds = pickle.load(token)
            # If there are no (valid) credentials available, let the user log in.
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        CREDS_PATH, SCOPES)
                    creds = flow.run_local_server()
                # Save the credentials for the next run
                with open(PICKLE_PATH, 'wb') as token:
                    pickle.dump(creds, token)

            service = build('sheets', 'v4', credentials=creds)

            # Call the Sheets API
            sheet = service.spreadsheets()
            result = sheet.values().get(spreadsheetId=SpreadURL,
                                        range=SAMPLE_RANGE_NAME).execute()
            values = result.get('values', [])
        except httplib2.ServerNotFoundError:
            print("Something Went Wrong")

        # Creates new window for viewing and interacting with Backup Management
        TKbackup = NewWindow()
        TKbackup.title("Backup Management")

        # Creates the button that will allow users to pull new backups
        backupbutton = tkinter.Button(TKbackup)
        # Button for creating backups
        backupbutton["text"] = "Click here \n to create a backup!"
        backupbutton["command"] = pull_backup

        # Creating the textbox and accompanying scrollbar that will show current values
        scroll = tkinter.Scrollbar(TKbackup)
        textbox = tkinter.Listbox(TKbackup, yscrollcommand=scroll.set)
        # TODO 'item' is a list of strings -- might want to change this for formatting purposes
        for item in values:
            if item == values[0]:
                item = '  '.join(item)
                textbox.insert('end', item)
            else:
                item = '  '.join(item)
                textbox.insert('end', item)

        # Placing everything where it belongs
        backupbutton.pack(side='bottom')
        textbox.pack(side='left', fill='both', expand=True)
        scroll.pack(side='right', fill='y')

        scroll.config(command=textbox.yview)
        TKbackup.mainloop()


# The Roster Window -- for managing the roster, such as entering or deleting members, as well as viewing the
# current roster
def create_roster_window():
    TKroster = NewWindow()
    TKroster.title("Roster")

    rosterdata = []

    # Refreshes the window by essentially destroying the current window and creating it again with new data
    def refresh():
        TKroster.destroy()
        create_roster_window()

    if os.path.exists(ROSTER_PATH):
        with open(ROSTER_PATH, 'r', newline='\n', encoding='utf-8') as rosterfile:
            rostreader = csv.reader(rosterfile)
            for member in rostreader:
                member = ''.join(member)  # Converts list to a single string
                rosterdata.append(member)
    else:
        with open(ROSTER_PATH, 'w+', newline='\n', encoding='utf-8') as rosterfile:
            rosterfile.truncate()

    scroll = tkinter.Scrollbar(TKroster)
    textbox = tkinter.Listbox(TKroster)

    if len(rosterdata) != 0:
        for member in rosterdata:
            textbox.insert('end', member)
    else:
        textbox.insert('end', 'No members in roster file!')

    editmem = tkinter.Button(TKroster, command=edit_members, text='Edit Members')
    refresh = tkinter.Button(TKroster, command=refresh, text='Refresh')

    refresh.pack(side='top')
    editmem.pack(side='bottom')
    textbox.pack(side='left', fill='both', expand=True)
    scroll.pack(side='right', fill='y')

    TKroster.mainloop()


# Editing members -- adding and deleting them
def edit_members():
    TKedit = NewWindow()
    TKedit.geometry('600x200')
    TKedit.title('Edit Member Roster')

    # Reading currently existing members off the roster file
    membersList = ['--Select a member--']
    with open(ROSTER_PATH, 'r', encoding='utf-8', newline='\n') as rosterfile:
        rostread = csv.reader(rosterfile)
        for member in rostread:
            member = ''.join(member)
            membersList.append(member)

    # Enter a new member into the roster
    def callback():
        with open(ROSTER_PATH, 'a', newline='\n', encoding='utf-8') as rosterfile:
            rostwrite = csv.writer(rosterfile)
            rostwrite.writerow([field.get()])
        field.select_clear()

    # Delete a member from the roster
    def deletemem():
        deletion = defaultVar.get()
        newRoster = []
        for mem in membersList:
            if mem == deletion:
                continue
            elif mem == '--Select a member--':
                continue
            else:
                newRoster.append(mem)

        with open(ROSTER_PATH, 'w', encoding='utf-8', newline='\n') as rosterfile:
            rostwriter = csv.writer(rosterfile)
            for member in newRoster:
                rostwriter.writerow([member])

        TKedit.destroy()

    # Build all of the buttons and whatnot into the TK window
    defaultVar = tkinter.StringVar(TKedit)
    defaultVar.set('--Select a member--')

    label1 = tkinter.Label(TKedit, text='Add a new member here: ')
    field = tkinter.Entry(TKedit)
    button1 = tkinter.Button(TKedit, text='Submit Changes', command=callback)
    memSelect = tkinter.OptionMenu(TKedit, defaultVar, *membersList)
    button2 = tkinter.Button(TKedit, text='Delete Member', command=deletemem)

    label1.pack(side='left')
    field.pack(side='left')
    button1.pack(side='left')
    button2.pack(side='bottom')
    memSelect.pack(side='bottom')

    TKedit.mainloop()


# Creating a window to show metrics on members in the fraternity and their volunteer hours
def create_metrics_window():
    TKmetrics = NewWindow()
    TKmetrics.title("Metrics")
    TKmetrics.geometry(str(TKmetrics.winfo_screenwidth()) + 'x' + str(TKmetrics.winfo_screenheight()))

    memberMetrics = []
    try:
        FILEPATH = askopenfilename(title='Select the backup file to use!')
    except FileNotFoundError:
        create_error_window('ERROR: You have to select a valid backup file to use')
        TKmetrics.quit()

    # Check backup file for current members and their volunteer hours
    with open(FILEPATH, 'r', encoding='utf-8', newline='\n') as backupfile:
        backupread = csv.reader(backupfile)
        for event in backupread:
            if event[0] == 'What is your name?':
                continue
            else:
                memberEvent = []
                memberEvent.append(event[0])
                memberEvent.append(float(event[3]))
                memberMetrics.append(memberEvent)
    memberMetrics.sort()

    # All of the monstrous code here is for adding together hours for members who have multiple events listed.
    # If this wasn't present, and you just used memberMetrics, it wouldn't properly add together a member's total hours
    cleanData = list()
    duplicate = False
    pos = 0

    # We need to check every event in memberMetrics
    for event in memberMetrics:
        # First member needs to be added
        if len(cleanData) == 0:
            cleanData.append(memberMetrics[0])
        # Then everyone else
        else:
            # And then if they're already present in the new list
            for entry in cleanData:
                if event[0] == entry[0]:
                    duplicate = True
                    value = event[1]
            # If they're in the new list, add their new event to the old one
            if duplicate:
                cleanData[pos][1] = value + cleanData[pos][1]
                duplicate = False
            # If they aren't in the new list, add them and their hours from their event
            else:
                cleanData.append(event)
                pos = pos + 1

    # Check for members who have yet to submit hours (wouldn't show up if we just check backups/current)
    with open(ROSTER_PATH, 'r', encoding='utf-8', newline='\n') as rosterfile:
        rostread = csv.reader(rosterfile)
        match = False
        for member in rostread:
            member = ''.join(member)
            for x in cleanData:
                if member == x[0]:
                    match = True
            if match:
                match = False
            else:
                addition = []
                addition.append(member)
                addition.append(0)
                cleanData.append(addition)

    # Using matplotlib to create a bar chart with members' volunteer hours
    f = Figure(figsize=(4, 5), dpi=100)
    names = []
    hours = []
    for member in cleanData:
        names.append(member[0])
        hours.append(member[1])

    f.add_subplot(111).bar(names, hours)

    canvas = FigureCanvasTkAgg(f, TKmetrics)
    canvas.draw()
    canvas.get_tk_widget().pack(side='top', fill='both', expand=True)

    TKmetrics.mainloop()


# The submissions window -- for taking the data from a backup file and entering it into a formalized Excel workbook
# This is probably the most complex part of this program
def create_submission_window():

    TKsubmit = NewWindow()
    TKsubmit.title("Submission")

    members = []
    with open(ROSTER_PATH, 'r', encoding='utf-8', newline='\n') as roster:
        rostreader = csv.reader(roster)
        for member in rostreader:
            member = ''.join(member)
            members.append(member)

    filename = askopenfilename()
    eventData = []

    # We need to get all of the data from the selected backup file
    with open(filename, 'r', encoding='utf-8', newline='\n') as rawdata:
        rawreader = csv.reader(rawdata)
        for i in rawreader:
            if i[0] == 'What is your name?':
                continue
            else:
                event = []
                volunteer = i[0]
                event.append(volunteer)
                duration = i[3]
                event.append(duration)
                location = i[2]
                event.append(location)
                eventData.append(event)
    eventData.sort()

    # Variables for going through the columns and ensuring duplicates are properly handled
    presentMem = []
    wb = load_workbook(TEMPLATE_URL)
    num = 6
    pos = 'B' + str(num)

    # Used to track if the below try-catch block ran into an error, and therefore the Excel spreadsheet shouldn't
    # be saved.
    success = True

    # For each of the unique events in eventData, find out if the volunteer is already listed. If so, add the event
    # to their row. If not, give them a new row and put in their first event.
    for event in eventData:
        try:
            # Member not already recorded. Give them a new row and put in their first event
            if event[0] not in presentMem:
                wb.active[pos] = event[0]
                presentMem.append(event[0])
                tempCol = 'C'
                pos = tempCol + str(num)
                wb.active[pos] = event[2]
                tempCol = 'D'
                pos = tempCol + str(num)
                wb.active[pos] = float(event[1])
                num += 1
                pos = 'B' + str(num)
            # Member already recorded. Move on to the next pair of columns to put in new event data
            elif event[0] in presentMem:
                # Member has one activity already recorded, new one will be added to the second pair of columns (E+F)
                if wb.active['E' + str(num-1)].value is None:
                    pos = 'E' + str(num - 1)
                    wb.active[pos] = event[2]
                    pos = 'F' + str(num - 1)
                    wb.active[pos] = float(event[1])
                    pos = 'B' + str(num)
                # Member has two activities already recorded, new one will be added to the third pair of columns (G+H)
                elif wb.active['G' + str(num-1)].value is None:
                    pos = 'G' + str(num - 1)
                    wb.active[pos] = event[2]
                    pos = 'H' + str(num - 1)
                    wb.active[pos] = float(event[1])
                    pos = 'B' + str(num)
                # Member already has three activies recorded, new one will be added to the fourth pair of columns (I+J)
                elif wb.active['I' + str(num-1)].value is None:
                    pos = 'I' + str(num - 1)
                    wb.active[pos] = event[2]
                    pos = 'J' + str(num - 1)
                    wb.active[pos] = float(event[1])
                    pos = 'B' + str(num)
                else:
                    errName = event[0]
                    raise ValueError('Too many events for one member: ' + event[0])
        except ValueError:
            success = False
            create_error_window('Too many events for one member: ' + errName)

    # If we didn't run into any errors, we have to save the new Excel file and notify user of success
    if success:
        wb.save('FINALIZED_SUBMISSION.xlsx')
        create_error_window('Successfully created a formalized volunteer spreadsheet! \n'
                            'It\'s located at: ' + os.path.abspath('FINALIZED_SUBMISSION.xlsx'))

    TKsubmit.mainloop()


# Create a window for error messages, defined in a way that's (hopefully) useful to the end user
def create_error_window(errText):
    TKError = NewWindow()
    TKError.geometry('400x300')
    errLabel = tkinter.Label(TKError, text=errText, wraplength=400)

    errLabel.pack()

    TKError.mainloop()


# Change the Google Sheets URL -- needs to be the FULL URL straight from the address bar when on the Sheet in a browser
def change_url(mainwindow):
    TKurl = NewWindow()
    TKurl.geometry('600x200')
    TKurl.title('Change Google Sheet')
    descriptor = tkinter.Label(TKurl, text='Paste the URL here:')
    field = tkinter.Entry(TKurl)

    # Executes when the user clicks the 'Submit' button -- changes the Spreadsheet URL
    def callback():
        global SpreadURL
        try:
            # The URL should be submitted as a full copy-paste of the URL with the Google Sheets file open in the
            # browser
            raw_input = str(field.get())
            input = raw_input.split('/d/')[1]
            input = input.split('/')[0]
            SpreadURL = input
            with open(SPRERL_PATH, 'w', encoding='utf-8') as pathfile:
                pathfile.truncate()
                pathfile.write(input)
            field.select_clear()
            TKurl.destroy()
            mainwindow.refresh()
        except IndexError:
            create_error_window('INDEX ERROR: Please make sure you properly copied the URL into the textbox')

    change = tkinter.Button(TKurl, text='Submit', command=callback)

    descriptor.pack(side='left')
    field.pack(side='left')
    change.pack(side='left')
    TKurl.mainloop()


# The starting/main execution point for Voluntracker.py
if __name__ == '__main__':
    # If we aren't executing on Windows, we need to change our file path syntax for the OS
    if platform.system() != 'Windows':
        BACKUP_PATH = 'Backups'
        ROSTER_PATH = 'Configuration/Roster.csv'
        LOGO_PATH = 'Configuration/DSP_logo.png'
        PICKLE_PATH = 'Configuration/token.pickle'
        CREDS_PATH = 'Configuration/credentials.json'
        SPRERL_PATH = 'Configuration/spreadurl.txt'
        TEMPLATE_URL = 'Configuration/BaseTemplate.xlsx'

    # Check to see if there's a current spreadsheet URL file in Configuration
    if os.path.exists(SPRERL_PATH):
        with open(SPRERL_PATH, 'r', encoding='utf-8') as pathfile:
            SpreadURL = pathfile.read()
    # If there isn't, create a new spreadsheet URL file and store a default value in it
    else:
        with open(SPRERL_PATH, 'w', encoding='utf-8') as pathfile:
            pathfile.write('NO CURRENT SPREADSHEET SELECTED')
            SpreadURL = 'NO CURRENT SPREADSHEET SELECTED'

    TK = MainWindow()

