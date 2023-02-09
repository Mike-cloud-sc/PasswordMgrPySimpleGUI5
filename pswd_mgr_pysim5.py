import PySimpleGUI as sg
import hashlib
import pandas as pd
import os


#  You must install (pip) xlsxWriter if you want to create an Excel file, no import required

# 1/9/23 I used "Cmd +" and "CMD -" to hide functions

# Create an app using py2app
# pip install py2app
# % py2applet --make-setup pswd_mgr_pysim5.py   (create a setup file needed for compliation)
# % python3 setup.py py2app -A  (create the app in test mode
# If errors, debug as per instructions in Python Coding Angela Yu.docx
# % python3 setup.py py2app  (create a portable .app)
'''
The password popup and password hash came from Demo.Password_Login.py downloaded
from PySimple course from Mike and Udemy.
'''


# ----------------------------- Paste this code into your program / script -----------------------------
# determine if a password matches the secret password by comparing SHA1 hash codes
def PasswordMatches(password, a_hash):
    password_utf = password.encode('utf-8')
    sha1hash = hashlib.sha1()
    sha1hash.update(password_utf)
    password_hash = sha1hash.hexdigest()
    return password_hash == a_hash

# I generated this hash using script: Demo_Password_Login.py from Udemy course PySimpleGUI
# The script is included in this project folder.  May never need to use it unless I change
# the password.
login_password_hash = '35f95f5551b94168e8cb7fb2c5adfa851e7476c1'  # Mike

sg.theme('light green')

# while True:
#     password = sg.popup_get_text('Enter password)',   password_char='*', font=('Helvetica 16'))
#     if password and PasswordMatches(password, login_password_hash):
#         sg.Window('', [[sg.T('SUCCESS!!!')], [sg.T('Password Validated', font=('Helvetica 16'))],  [sg.OK(s=10), ]],  location=(1000, 90), disable_close=True).read(close=True)
#         break
#     else:
#         answer = sg.popup_get_text('Password Failed, Do you want to try again? Y or N',  background_color='red', font=('Helvetica 16'), location=(1000, 90))
#         if answer == 'n' or answer == 'N':
#             exit()

def search(search_arg):
    my_list1=[]
    my_list2 = []
    count_list_items = 0  # used to determine if there were any matches during search
    # Search all fields(Name, UserID, Password, Notes) for a specific set of characters
    for index, row in df.iterrows():
        row.UserID = str(row.UserID)  # convert numeric values to strings
        row.Notes = str(row.Notes)  # convert numeric values to strings

        if search_arg in str(row.Name.lower()) or search_arg in row.UserID.lower()\
                or search_arg in row.Notes.lower() or search_arg in row.Password.lower():
            count_list_items += 1

            # Write each account info to listbox
            index = str(index).zfill(3)
            my_list1.append(f'{""} {index} {" " *3}  {row.Name} \n')
            my_list2.append(f'{""} {index}   {row.UserID} \n')

    if count_list_items > 0:
        window['-LBOX1-'].update(values=my_list1)
        window['-LBOX2-'].update(values=my_list2)

    else:
        # window['-LBOX-'].update('')  # clear all items in the listbox
        sg.popup('No Record Found', any_key_closes=True, background_color='red', font=('Helvetica 16'), location=(1000,90))
    return


    #  If user clicks on a different line in the list box (-LBOX-), then repopulate the fields

def list_changed(item):  # User method
    # if item_index.isdigit():
    #    item_index = int(item_index)  # index# must be integer to delete records

    # df2 = {'Name': name, 'UserID': user_id, 'Password': pwd, 'Notes': notes}
    name = df.at[item, 'Name']
    userid= df.at[item, 'UserID']
    password = df.at[item, 'Password']
    notes= df.at[item, 'Notes']

    # Display names:  '-ACCOUNT NAME-', '-USERID-', '-PASSWORD-', '-NOTES-'

    window['-ACCOUNT NAME-'].update(name)
    window['-USERID-'].update(userid)
    window['-PASSWORD-'].update(password)
    window['-NOTES-'].update(notes)
    return


def clear_input_output_fields():
    window['-ACCOUNT NAME-'].update('')
    window['-USERID-'].update('')
    window['-PASSWORD-'].update('')
    window['-NOTES-'].update('')
    window['-LBOX1-'].update('')
    window['-LBOX2-'].update('')
    window['-lineEdit_search-'].update('')

def add_record():
    global df
    name = window['-ACCOUNT NAME-'].get()
    user_id = window['-USERID-'].get()
    pwd = window['-PASSWORD-'].get()
    notes = window['-NOTES-'].get()

    answer= sg.popup_get_text(f'Do you want to add Name: {name}, Y or N: ', font=('Helvetica 16'),  location=(1000,90))
    answer = answer.lower()

    if answer == 'y':
        # Add (append) new password info to DataFrame
        df2 = {'Name': name, 'UserID': user_id, 'Password': pwd, 'Notes': notes}
        df = df.append(df2, ignore_index=True) # will be droped from pandas July 2022
        # save dataframe
        df.to_pickle(password_file)
        sg.popup("DataFrame was saved to disk", font=('Helvetica 16'), location=(1000,90))
        # Clear the input fields after record added to dataframe
        clear_input_output_fields()  # callfunction to clear fields
    else:
        sg.popup("New record was not added",  font=('Helvetica 16'), location=(1000,90))

def delete():
    answer= sg.popup_get_text(f'Do you really want to delete record#: {item}, Y or N: ', font=('Helvetica 16'), location=(1000,90))
    if answer == 'y':
        df.drop(item, inplace=True)  # default for row is axis=0
        # save dataframe
        df.to_pickle(password_file)
        sg.popup(f"Record# {item} was deleted and dataframe was saved to disk", font=('Helvetica 16'),  background_color= 'yellow', text_color= 'black', location=(1000,90))
        clear_input_output_fields()
    else:
        sg.popup(f"Record# {item} was not deleted", font=('Helvetica 16'), location=(1000, 90))



def df_update():
    global df
    name = window['-ACCOUNT NAME-'].get()
    user_id = window['-USERID-'].get()
    pwd = window['-PASSWORD-'].get()
    notes = window['-NOTES-'].get()

    answer = sg.popup_get_text(f'Do you want to update record#: {item}, Y or N: ', font=('Helvetica 16'), location=(1000, 90))
    if answer == 'y':
        # Write all data fields to dataframe. Name, UserID, Password and Notes are column names in dataframe
        # item_index is the row number of the dataframe
        df.loc[item, "Name"] = name
        df.loc[item, "UserID"] = user_id
        df.loc[item, "Password"] = pwd
        df.loc[item, "Notes"] = notes

        # save dataframe
        df.to_pickle(password_file)

        sg.popup("Record updated and Dataframe was saved to disk",  font=('Helvetica 16'), location=(1000,90))

        # Clear the input fields after record updated and dataframe saved to disk
        clear_input_output_fields()  # callfunction to clear fields
    else:
        sg.popup("Record not updated. ",  font=('Helvetica 16'), location=(1000,90))
    return

def display_spreadsheet():
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('Passwords.xlsx', engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object.
    df.to_excel(writer, sheet_name='Sheet1')  # create Excel spreadsheet

    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    # Get worksheet
    worksheet = writer.sheets['Sheet1']

    cell_format = workbook.add_format({'text_wrap': True})  # wrap cell and auto adjust height of cell
    cell_format.set_font_size(16)

    # Set the column width and format.
    worksheet.set_column('A:A', 3, cell_format)
    worksheet.set_column('B:B', 40, cell_format)
    worksheet.set_column('C:C', 40, cell_format)
    worksheet.set_column('D:D', 40, cell_format)
    worksheet.set_column('E:E', 50, cell_format)
    worksheet.set_column('A:E', None, cell_format)

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()

    # Open Excel File
    os.system('open -a "/Applications/Microsoft Excel.app" "Passwords.xlsx"')

#  1/10/23  This script shares the passwords.pkl file with PasswordManagerGUILoginWindow
#  script to insure integrity of data

password_file = '/Users/michaelpauken/PythonActiveProjects/PasswordMgrPySimpleGUI5/passwords.pkl'
df = pd.read_pickle(password_file)  # create dataframe of passwords from pickle file

empty_list = []  # integrity

# sg.Print('', font='Default 18', keep_on_top=True, size=(40,30), location=(1520,100))  #  Debug Window

sg.theme('light green')

# 2 - Layout
layout = [
    [sg.Text('Enter Search Argument'), sg.Input(key='-lineEdit_search-', size=(30,1), expand_x = True), sg.Button('Search',   bind_return_key=True)],
    [sg.T('REC#', size=(10, 1)), sg.T('ACCOUNT NAME', size=(40, 1)), sg.T('USER ID', size=(30, 1))],
    [sg.Listbox(empty_list, size=(40, 10), select_mode=sg.LISTBOX_SELECT_MODE_EXTENDED, no_scrollbar=False, horizontal_scroll=False, expand_x=True, expand_y=True, enable_events=True, k='-LBOX1-'),
    sg.Listbox(empty_list, size=(40, 10), select_mode=sg.LISTBOX_SELECT_MODE_EXTENDED, no_scrollbar=False, horizontal_scroll=False, expand_x=True, expand_y=True, enable_events=False, k='-LBOX2-')],
    [sg.Text('Account Name', size=(11)), sg.Push(), sg.Input(key='-ACCOUNT NAME-', size=(70), expand_x=True)],
    [sg.Text('UserID', size=(11)), sg.Push(), sg.Input(key='-USERID-', size=(70), expand_x=True)],
    [sg.Text('Password', size=(11)), sg.Push(), sg.Input(key='-PASSWORD-', size=(70), expand_x=True)],
    [sg.Text('Notes', size=(11)), sg.Push(), sg.Multiline(size=(70, 3), key='-NOTES-', disabled=False, autoscroll=True, horizontal_scroll=True, expand_x=True, expand_y=True)],
    [sg.Button('Add'), sg.Button('Delete'), sg.Button('Update'), sg.Button('Clear Input Fields'), sg.Button('Exit'), sg.Button('Excel')]
]

# Display names:  '-ACCOUNT NAME-', '-USERID-', '-PASSWORD-', '-NOTES-'

# 3 - Window
window = sg.Window('Password Manager', layout, font='Default 18',  finalize=True, resizable=True)

while True:
    # 4 - Event loop / handling
    event, values = window.read()

    # sg.Print(f'event = {event}', c='white on red', erase_all=True)
    # sg.Print(*[f'   {k} = {values[k]}' for k in values], sep='\n')

    if event == sg.WIN_CLOSED or event == 'Exit':
        exit()

    # start building the password manager window here

    if event == 'Search':
        search_arg = window['-lineEdit_search-'].get().lower()
        my_list =search(search_arg)  # goto search function

    try:
        if event == '-LBOX1-':
            item = window['-LBOX1-'].get()
            item = item[0]
            item = int(item.split()[0])  # get the first character in the string and convert to integer
            list_changed(item)  # call function to populate the update fields

    except Exception as e:
        sg.popup('No Record Found', any_key_closes=True, background_color='red', font=('Helvetica 16'),
                 location=(1000, 90))

        list_changed(item)  #  call function to populate the update fields

    if event == 'Clear Input Fields':
        clear_input_output_fields()  #  callfunction to clear fields

    if event == 'Add':
        add_record()

    if event == 'Delete':
        delete()   # function for deleting record

    if event == 'Update':
        df_update()

    if event == 'Excel':
        display_spreadsheet()
    #
    # sg.Print(f'event = {event}', c='white on red', erase_all=True)
    # sg.Print(*[f'   {k} = {values[k]}' for k in values], sep='\n')  # prints the line selected
    # 5 - Close
    # window.close()



# if __name__ == '__main__':
#     sg.theme('DarkPurple5')
#     # main_login()


# password = sg.popup_get_text(
#             'This is the next window')
#         if password == '' or password == 'stop':
#             break

# sg.theme('Dark Purple 4')
# sg.popup(sg.theme_background_color('#FF0000'))  # red hex code
# exit()