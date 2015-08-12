__version__ = '1.0'
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait  # available since 2.4.0
from selenium.webdriver.support import expected_conditions as EC  # available since 2.26.0
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException

import time
import csv
from datetime import datetime
import win32com.client
import os
from custom_modules.outlook import Outlook


class Automator:
    """
    Add Users or Reset Passwords in ICE.
    Requires a csv file with headers (in any order):
    firstName, surname, username, description, role, location, newPassword, Status
    Status column will be updated on output. For resetting users it will search on username first, then full name
    It will output input_fileName_output.csv as the results file. Make sure that file is ok to be wiped!

    http://selenium-python.readthedocs.org/en/latest/getting-started.html
    https://gist.github.com/huangzhichong/3284966
    """
    def __init__(self, username, password, url_login):
        self.username = username
        self.password = password
        self.url_login = url_login

    def login(self):
        """
        Logs into ICE, sets self.driver
        """
        try:
            self.driver = webdriver.Ie()
            self.driver.get(self.url_login)
            self.driver.maximize_window()
            self.driver.find_element_by_id('txtName').send_keys(self.username)
            self.driver.find_element_by_id('txtPassword').send_keys(self.password)
            self.driver.execute_script('frmLogin.action = "login.aspx?action=login";frmLogin.submit();')
            # Add/Edit User
            wait = WebDriverWait(self.driver, 10)
            wait.until(EC.element_to_be_clickable((By.ID, 'a151'))) # Add edit user button
            time.sleep(1)
            return
        except Exception:
            raise

    def add_user(self, details):
        """
        Add a single user into ICE
        Assumes login() has been called
        details = {'firstName': '', 'surname': '', 'username': '', 'description': '', 'role': '', 'location': '', 'newPassword: ''}
        Returns string: comment of what happened
        """
        try:
            self.driver.switch_to.default_content()   # Jump to the top of the frames hierachy
            self.driver.find_element_by_id('a151').click()   # Add/Edit user button
            wait = WebDriverWait(self.driver, 10)
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, 'Right')))
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, 'appFrame')))
            # Add new user menu
            self.driver.find_element_by_id('AddButton').click()
            self.driver.find_element_by_id('usernameTextBox').send_keys(details['username'])
            self.driver.find_element_by_id('PasswordTextBox').send_keys(details['newPassword'])
            self.driver.find_element_by_id('ConfirmPasswordTextBox').send_keys(details['newPassword'])
            if not self.driver.find_element_by_id('ChangePasswordCheckBox').is_selected():   # Should always be unticked on load
                self.driver.find_element_by_id('ChangePasswordCheckBox').click()
            self.driver.find_element_by_id('FullnameTextBox').send_keys(details['firstName'] + ' ' + details['surname'])
            self.driver.find_element_by_id('InitialsTextbox').send_keys(details['firstName'][:1] + details['surname'][:1])
            self.driver.find_element_by_id('DescriptionTextBox').send_keys(details['description'])  # Description/Job title
            Select(self.driver.find_element_by_id('RoleList')).select_by_visible_text(details['role'])  # Role dropdown
            # Locations Profile
            wait.until(EC.element_to_be_clickable((By.ID, 'imgLP')))
            self.driver.find_element_by_id('imgLP').click()
            Select(self.driver.find_element_by_id('LocationListBox')).select_by_visible_text(details['location'])  #All Locations dropdown
            self.driver.find_element_by_id('AddButton').click()
        except:
            return "There was a problem filling in the page. Can you check the role/location etc?"
        try:
            self.driver.find_element_by_id('btnCommand').click()     # Save user
            time.sleep(1)
            # Alert will display if a duplicate is found in the system
            alert = Alert(self.driver)
            alertText = alert.text
            alert.accept()
            wait.until(EC.element_to_be_clickable((By.ID, 'btnCommand')))   # Wait for Save User button
            self.driver.find_element_by_id('btnGoToIndex').click()
            if alertText[:13] == "Create failed" and alertText[-30:] == "already exists; cannot create.":
                return "Duplicate person found in the system"
            else:
                return alertText
        except NoAlertPresentException:
            # If you have a success message
            try:
                if self.driver.find_element_by_id('messageDisplay').text.strip() == \
                    'The user has been successfully updated.'\
                    or self.driver.find_element_by_id('messageDisplay').text.strip() == \
                        'The user has been successfully added.':
                    return 'User added successfully'
                else:
                    return self.driver.find_element_by_id('messageDisplay').text.strip()
            except NoSuchElementException:
                # You are now stuck on the page unable to save with an error (usually unrecoverable for add user)
                # Password problem?
                try:
                    if self.driver.find_element_by_id('PasswordValidator').text == 'You have used this password before in your last three passwords.':
                        return "Couldn't save the user as password has been used before."
                    else:
                        return self.driver.find_element_by_id('PasswordValidator').text
                except NoSuchElementException:
                    # Location correction
                    try:
                        if self.driver.find_element_by_id('spanLocationError').text == "There must be at least one location in the user's profile.":
                            Select(self.driver.find_element_by_id('LocationListBox')).select_by_visible_text(details['location'])  #All Locations dropdown
                            self.driver.find_element_by_id('AddButton').click()
                            self.driver.find_element_by_id('btnCommand').click()     # Save user
                            time.sleep(1)
                            try:  # If you have a success message
                                if self.driver.find_element_by_id('messageDisplay').text.strip() == "The user has been successfully updated.":
                                    return "Success (& location updated)"
                            except NoSuchElementException:
                                pass
                    except:
                        pass
            return "Couldn't save the user for some reason I can't determine."

    def add_users_from_file(self, input_file, out_file):
        """
        Add users from input_file
        Write input_file out to out_file with a Status column added
        input_file should have headers:
        firstName, surname, username, description, role, location, newPassword, Status
        """

        csvFileRead = open(input_file, 'r')
        dictOfRows = csv.DictReader(csvFileRead)

        # Process file entries, appending to the file one at a time
        for row in dictOfRows:
            csvFileWrite = open(out_file, 'a', newline='')
            writer = csv.DictWriter(csvFileWrite, dictOfRows.fieldnames)
            print('---\nProcessing: ' + row['firstName'] + ' ' + row['surname'])
            if not self.password_validates(row['newPassword']):
                Comment = "ICE won't accept this password even if i try it!"
            else:
                Comment = self.addUser(row)
            print(Comment)
            i = datetime.now()
            row['Status'] = Comment + ' (%s/%s/%s %s:%s)' % (i.day, i.month, i.year, i.hour, i.minute)
            writer.writerow(row)
            csvFileWrite.close()

        csvFileRead.close()
        self.driver.quit()

    def add_users_from_fileProcessInbox(self,
                                     input_file,
                                     out_file,
                                     look_for="yes",
                                     emailDetails={}):
        """
        Monitor default Outlook Inbox for messages with (e.g.) "Yes" in Subject line
        If found, set up user in ICE taking details from input_file, (email end user), append result to out_file
        emailDetails = { 'fromAddress': '',
                         'userSubject': '',
                         'userHTMLFile': '',
                         'userAttachFolder': '',
                         'passSubject': '',
                         'passHTMLFile': '',
                         'passAttachFolder': ''
                         'UHBAddress': ''}
        """
        ol = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = ol.GetDefaultFolder(6) #olFolderInbox
        messages = inbox.Items
        for msg in messages:
            # If look_for found in subject line
            if msg.Subject.lower().find(look_for.lower()) != -1:
                csvFileRead = open(input_file, 'r')
                dictOfRows = csv.DictReader(csvFileRead)
                for row in dictOfRows:
                    try:
                        # Exchange users are a pain - you have to look up their email
                        if msg.SenderEmailType == "EX":
                            sender = msg.Sender
                            exchUser = sender.GetExchangeUser()
                            frmAddr = exchUser.PrimarySmtpAddress
                        else:
                            frmAddr = msg.SenderEmailAddress
                    except:
                        # msg.SenderEmailType raises error on (Message Recall mails i think)
                        continue
                    if frmAddr.lower().strip() == row['email'].lower().strip():
                        # Add to ICE, write to outputFile
                        csvFileWrite = open(out_file, 'a', newline='')
                        writer = csv.DictWriter(csvFileWrite, dictOfRows.fieldnames)
                        print('---\nProcessing: ' + row['firstName'] + ' ' + row['surname'])
                        if not self.password_validates(row['newPassword']):
                            Comment = "ICE won't accept this password even if i try it!"
                            # Shouldn't ever get here if input_file is validated, but might want to consider emailing someone the fail here?
                        else:
                            try:
                                self.login()
                            except:
                                print('Warning - cant log into ICE, I need to tell someone!')
                                return
                            Comment = self.add_user(row)
                        print(Comment)
                        i = datetime.now()
                        row['Status'] = Comment + ' (%s/%s/%s %s:%s)' % (i.day, i.month, i.year, i.hour, i.minute)
                        writer.writerow(row)
                        csvFileWrite.close()
                        # Email end user
                        Comment = 'User added successfully'
                        if Comment == 'User added successfully':
                            # Email username
                            self.process_html_and_email(row, emailDetails['userHTMLFile'], emailDetails['userSubject'], emailDetails['userAttachFolder'], emailDetails['fromAddress'])
                            time.sleep(30)
                            # Email password
                            self.process_html_and_email(row, emailDetails['passHTMLFile'], emailDetails['passSubject'], emailDetails['passAttachFolder'], emailDetails['fromAddress'])
                            # Email UHB
                            fwdMsg = msg.Forward()
                            fwdMsg.To = emailDetails['UHBAddress']
                            fwdMsg.Send()
                            return
                # Person not found in spreadsheet here - what do we do?

    def process_html_and_email(self, row, htmlFile, subject, attachFolder, fromAddress):
        """
        Will process out $username, $password in HTML and send out email
        """
        htmlRaw = open(htmlFile, 'r').read()
        html = htmlRaw.replace('$firstname', row['firstName'])
        html = html.replace('$firstname', row['surname'])
        html = html.replace('$username', row['username'])
        html = html.replace('$password', row['newPassword'])
        attachments = []
        if os.path.exists(attachFolder):
            attachments = [os.path.join(attachFolder, fn) for fn in os.listdir(attachFolder)]
        o = Outlook()
        o.send(True, row['email'], subject, '', html, attachments=attachments, accountToSendFrom=fromAddress)

    def reset_password(self, details):
        """
        Reset password for and individual user
        details = {'firstname': 'myfirstname', 'surname': 'mysurname' etc}
        Returns string: comment of what happened
        """
        self.driver.switch_to.default_content()     # Jump to the top of the frames hierachy
        self.driver.find_element_by_id('a151').click()   # Add/Edit user button
        wait = WebDriverWait(self.driver, 15)
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, 'Right')))
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, 'appFrame')))

        if details['username'] != '':
            searchText = details['username'].strip()
            htmlColumnName = 'username'
            htmlSearchDropDown = 'username'
        else:
            searchText = details['firstName'].strip() + ' ' + details['surname'].strip()
            htmlColumnName = 'Full Name'
            htmlSearchDropDown = 'full name'

        Select(self.driver.find_element_by_id('SearchTypeList')).select_by_visible_text(htmlSearchDropDown)  # Search by
        time.sleep(1)
        self.driver.find_element_by_id('SearchTextBox').send_keys(searchText)
        self.driver.find_element_by_id('SearchButton').click()
        time.sleep(2)   # Give time for results to appear

        # Search html results: a specific column for text
        rowNo = self.find_result_row(searchText, htmlColumnName)

        # Return if there is no match or no results or a duplicate
        if rowNo == -1:  # No match or no results
            if not self.driver.find_element_by_id('chkIncludeInactive').is_selected():
                self.driver.find_element_by_id('chkIncludeInactive').click() # Show inactive users
            self.driver.find_element_by_id('SearchButton').click()
            rowNo = self.find_result_row(searchText, htmlColumnName)
            if rowNo == -1:
                return "Can't find person"
            elif rowNo == -2:
                return "Duplicate person found in the system"
        elif rowNo == -2:
            return "Duplicate person found in the system"

        # Result should be found. click result's full name entry (in theory it should always be the first in the list, but doesn't hurt to check)
        self.driver.find_elements_by_css_selector('.dataGridRow')[rowNo].find_elements_by_css_selector('td')[1].click()

        # Set password
        self.driver.find_element_by_id('PasswordTextBox').send_keys(details['newPassword'])
        self.driver.find_element_by_id('ConfirmPasswordTextBox').send_keys(details['newPassword'])
        if not self.driver.find_element_by_id('ChangePasswordCheckBox').is_selected():   # Should always be unticked on load
            self.driver.find_element_by_id('ChangePasswordCheckBox').click()
        self.driver.find_element_by_id('btnCommand').click()     # Save user
        time.sleep(1)
        try:  # If you have a success message
            if self.driver.find_element_by_id('messageDisplay').text.strip() == 'The user has been successfully updated.':
                return 'Password Updated'
            else:
                return self.driver.find_element_by_id('messageDisplay').text.strip()
        except NoSuchElementException:
            try:    # Password problem?
                if self.driver.find_element_by_id('PasswordValidator').text == 'You have used this password before in your last three passwords.':
                    return "Couldn't save the user as password has been used before."
                else:
                    return self.driver.find_element_by_id('PasswordValidator').text
            except NoSuchElementException:
                try:    # Location correction
                    if self.driver.find_element_by_id('spanLocationError').text == "There must be at least one location in the user's profile.":
                        Select(self.driver.find_element_by_id('LocationListBox')).select_by_visible_text(details['location'])  #All Locations dropdown
                        self.driver.find_element_by_id('AddButton').click()
                        self.driver.find_element_by_id('btnCommand').click()     # Save user
                        time.sleep(1)
                        try:  # If you have a success message
                            if self.driver.find_element_by_id('messageDisplay').text.strip() == "The user has been successfully updated.":
                                return "Password Updated (& location updated)"
                        except NoSuchElementException:
                            pass
                except:
                    pass
        return "Couldn't save the user for some reason I can't determine."

    def reset_passwords_from_file(self, input_file, out_file):
        """
        Reset users passwords from a file
        Write input_file out to out_file with a Status column added
        input_file should have headers:
        firstName, surname, username, description, role, location, newPassword, Status
        """
        csvFileRead = open(input_file, 'r')
        dictOfRows = csv.DictReader(csvFileRead)
        csvFileWrite = open(out_file, 'w', newline='')
        writer = csv.DictWriter(csvFileWrite, dictOfRows.fieldnames)

        # Process file entries
        writer.writeheader()
        for row in dictOfRows:
            if not self.password_validates(row['newPassword']):
                Comment = "ICE won't accept this password even if i try it!"
            else:
                Comment = self.reset_password(row)
            print(Comment)
            i = datetime.now()
            row['Status'] = Comment + ' (%s/%s/%s %s:%s)' % (i.day, i.month, i.year, i.hour, i.minute)
            writer.writerow(row)

        csvFileRead.close()
        csvFileWrite.close()
        time.sleep(3)
        self.driver.quit()

    def find_result_row(self, searchText, htmlColumnName):
        """
        Search html page results for searchText in the html column name in the results screen
        Return: row of result, or -1 if not found, or -2 for duplicate
        """
        headers = []
        results = []
        htmlHeader = self.driver.find_elements_by_css_selector('.header')
        for header in htmlHeader:
            headers.append(header.text)

        htmlResults = self.driver.find_elements_by_css_selector('.dataGridRow')
        for id, row in enumerate(htmlResults):
            results.append({headers[id]:item.text for id, item in enumerate(row.find_elements_by_css_selector('td'))})

        if results:
            # Find result row, 0 = first entry
            for id, person in enumerate(results):
                if searchText.lower() == person[htmlColumnName].lower():
                    # Match found, but check for duplicates
                    temp = []
                    for people in results:
                        if people[htmlColumnName] in temp:
                            return -2   # Duplicate found
                        else:
                            temp.append(people[htmlColumnName])
                    return id
        return -1    # No result or no match

    def password_validates(self, password):
        """
        Check password is > 5 characters, has at least 1 alpha and at least 1 number
        """
        if any(char.isdigit() for char in password) \
                and any(char.isalpha() for char in password) \
                and len(password) > 5:
            return True
        else:
            return False

def demo():
    #url_login = "http://ocs/icedesktop/dotnet/icedesktop/login.aspx"     #live
    url_login = "http://ocstrain/icedesktop/dotnet/icedesktop/login.aspx"    #test
    ice = Automator('<USERNAME>', '<PASSWORD>', url_login)
    emailDetails = {
        'fromAddress': 'simon.crouch@nbt.nhs.uk',
        'userSubject': 'Here are your new ICE login details',
        'userHTMLFile': r'K:\Coding\Python\supporting files\ICE\usernameDetails.htm',
        'userAttachFolder': r'K:\Coding\Python\supporting files\ICE\userAttachments',
        'passSubject': 'Here is your new ICE password',
        'passHTMLFile': r'K:\Coding\Python\supporting files\ICE\passwordDetails.htm',
        'passAttachFolder': r'K:\Coding\Python\supporting files\ICE\passwordAttachments',
        'UHBAddress': 'simon.crouch@nhs.net'
    }
    ice.add_users_from_fileProcessInbox(r'K:\Coding\Python\supporting files\ICE\ICE.csv', r'K:\Coding\Python\supporting files\ICE\ICE_output.csv', "!ICETest", emailDetails)

print('ICE User Addition Automation (Simon Crouch April 2015)\nServer is running.')
demo()
