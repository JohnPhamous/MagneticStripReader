# Cross Campus Entrepreneur's Club Mag Strip Reader
# Based off Kyle Minshall's Mag Strip Reader
# https://github.com/acm-ucr/Card-Swiper/blob/master/app.py

import re, openpyxl, getpass, os

#Sets up .xlsx file with inputs
new_wb = openpyxl.Workbook()
new_wb_sheet = new_wb.get_active_sheet()
new_wb_sheet.title = 'CCE Logins'
new_wb_sheet['A1'] = 'First Name'
new_wb_sheet['B1'] = 'Last Name'
new_wb_sheet['C1'] = 'Email'
new_wb_sheet['D1'] = 'SID'

enter_again = True
counter = 2

while enter_again == True:
    os.system('clear')
    print('Hello! You are signing in for the Cross Campus Entrepreneur Club.')
    card_info = getpass.getpass('Swipe your card: ', stream=None)

    if card_info == 'end':
        new_wb.save('CEE Logins.xlsx')
        enter_again = False

    else:

        #locates first name
        regex_firstname = '/([^\s]+)'
        firstname_match = re.findall(regex_firstname, card_info)

        #locates last name
        regex_lastname = '\^(.*?)/'
        lastname_match = re.findall(regex_lastname, card_info)

        #locates SID
        regex_sid = '000000([0-9]{9}).*'
        sid_match = re.findall(regex_sid, card_info)
        
        #Prompts to user to check if information is correct
        print('Your first name is ' + firstname_match[0])
        print('Your last name is ' + lastname_match[0])
        #print('Your SID is ' + sid_match[0])
        
        valid_email = False
        while valid_email == False:
            email_info = input('Enter your email: ')
            print('Is your email ' + email_info + '?')
            good = input('(y/n) ')
            if good == 'n':
                valid_email = False
            else:
                valid_email = True
                new_wb_sheet['A' + str(counter)] = str(firstname_match[0])
                new_wb_sheet['B' + str(counter)] = str(lastname_match[0])
                new_wb_sheet['C' + str(counter)] = str(email_info).upper()
                new_wb_sheet['D' + str(counter)] = str(sid_match[0]).upper()
                os.system('clear')
                counter += 1
                print(' ')        
print('Finished')

