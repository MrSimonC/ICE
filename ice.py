import time
import csv
import os
import sys
import win32com.client
from custom_modules.outlook import Outlook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
from datetime import datetime
__version__ = '1.1.12'
# v1.1.9 = added in for no in range(inbox.Items.Count-1, -1, -1)
# v1.1 has logout

# Set working path i.e. make all file references relative to this main module or .exe path
# TODO remove_no this terrible way of coding:
os.chdir(os.path.dirname(sys.executable)) if hasattr(sys, "frozen") \
    else os.chdir(os.path.dirname(os.path.realpath(__file__)))


class Automate:
    """
    Add Users or Reset Passwords in ICE.
    Requires a csv file with headers (in any order):
    firstName, surname, username, description, role, location, newPassword, Status
    Status column will be updated on output. For resetting users it will search on username first, then full name
    It will output input_fileName_output.csv as the results file. Make sure that file is ok to be wiped!

    ***SETUP
    Ensure "Enable Protected Mode" is Disabled for all zones in IE
    Ensure IEDriver is in same directory

    http://selenium-python.readthedocs.org/en/latest/getting-started.html
    https://gist.github.com/huangzhichong/3284966
    """
    def __init__(self, username, password, url_login):
        self.username = username
        self.password = password
        self.url_login = url_login
        self.driver = None  # instantiated at login
        self.wait = None  # instantiated at login

    def login(self):
        """
        Logs into ICE, sets self.driver
        """
        try:
            self.driver = webdriver.Ie()
            self.driver.maximize_window()
            self.driver.get(self.url_login)
            self.driver.find_element_by_id('txtName').send_keys(self.username)
            self.driver.find_element_by_id('txtPassword').send_keys(self.password)
            self.driver.execute_script('frmLogin.action = "login.aspx?action=login";frmLogin.submit();')
            # Add/Edit User
            self.wait = WebDriverWait(self.driver, 10)
            self.wait.until(EC.element_to_be_clickable((By.ID, 'a151')))  # Add edit user button
            time.sleep(1)
            return
        except Exception:
            raise

    def add_user(self, details):
        """
        Add a single user into ICE
        Assumes login() has been called
        :param details: = {'firstName': '', 'surname': '', 'username': '', 'description': '', 'role': '', 'location': '', 'newPassword: ''}
        :return string: comment of what happened
        """
        try:
            self.driver.switch_to.default_content()   # Jump to the top of the frames hierachy
            self.driver.find_element_by_id('a151').click()   # Add/Edit user button
            wait = WebDriverWait(self.driver, 10)
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, 'Right')))
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, 'appFrame')))
            # Add new user menu
            self.driver.find_element_by_id('AddButton').click()
            # self.wait.until(EC.element_to_be_clickable((By.ID, 'UsernameTextBox')))
            self.driver.find_element_by_id('UsernameTextBox').send_keys(details['username'])
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
            alert_text = alert.text
            alert.accept()
            wait.until(EC.element_to_be_clickable((By.ID, 'btnCommand')))   # Wait for Save User button
            self.driver.find_element_by_id('btnGoToIndex').click()
            if alert_text[:13] == "Create failed" and alert_text[-30:] == "already exists; cannot create.":
                return "Duplicate person found in the system"
            else:
                return alert_text
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
                    if self.driver.find_element_by_id('PasswordValidator').text == \
                            'You have used this password before in your last three passwords.':
                        return "Couldn't save the user as password has been used before."
                    else:
                        return self.driver.find_element_by_id('PasswordValidator').text
                except NoSuchElementException:
                    # Location correction
                    try:
                        if self.driver.find_element_by_id('spanLocationError').text == \
                                "There must be at least one location in the user's profile.":
                            Select(self.driver.find_element_by_id('LocationListBox')).\
                                select_by_visible_text(details['location'])  # All Locations dropdown
                            self.driver.find_element_by_id('AddButton').click()
                            self.driver.find_element_by_id('btnCommand').click()     # Save user
                            time.sleep(1)
                            try:  # If you have a success message
                                if self.driver.find_element_by_id('messageDisplay').text.strip() == \
                                        "The user has been successfully updated.":
                                    return "Success (& location updated)"
                            except NoSuchElementException:
                                pass
                    except:
                        pass
            return "Couldn't save the user for some reason I can't determine."

    def log_off(self):
        # self.driver.execute_script('javascript:closeDown();')  # produces javascript error
        time.sleep(1)
        self.driver.quit()

    def add_users_from_file(self, input_file, out_file):
        """
        Add users from input_file
        Write input_file out to out_file with a Status column added
        :param input_file: should have headers: firstName, surname, username, description, role,
        location, newPassword, Status
        :param out_file: file to append results to
        """

        csv_file_read = open(input_file, 'r')
        rows_dict = csv.DictReader(csv_file_read)

        # Process file entries, appending to the file one at a time
        for row in rows_dict:
            csv_file_write = open(out_file, 'a', newline='')
            writer = csv.DictWriter(csv_file_write, rows_dict.fieldnames)
            print('---\nProcessing: ' + row['firstName'] + ' ' + row['surname'])
            if not self.password_validates(row['newPassword']):
                comment = "ICE won't accept this password even if i try it!"
            else:
                comment = self.add_user(row)
            print(comment)
            # i = datetime.now()
            # row['Status'] = comment + ' (%s/%s/%s %s:%s)' % (i.day, i.month, i.year, i.hour, i.minute)
            row['Status'] = comment + ' (' + datetime.now().strftime('%d %b %Y %H:%M') + ')'  # 01 Jan 1900 19:00
            writer.writerow(row)
            csv_file_write.close()

        csv_file_read.close()
        self.driver.quit()

    def add_users_from_file_process_inbox(self,
                                          input_file,
                                          out_file,
                                          look_for="yes",
                                          email_details=None):
        """
        Monitor default Outlook Inbox for messages with (e.g.) "Yes" in Subject line
        If found, set up user in ICE taking details from input_file, (email end user), append result to out_file
        :param input_file: csv file path containing headers: firstName, surname, username, description, role,
        location, newPassword, email, Status. Ensure there are no email address duplicates
        :param out_file: csv path for audit trail. Writes headers: firstName, surname, username, description, role,
        location, newPassword, email, Status
        :param look_for: string to look for in subject
        :param email_details: { 'fromAddress': '',
                                'userSubject': '',
                                'userHTMLFile': '',
                                'userAttachFolder': '',
                                'passSubject': '',
                                'passHTMLFile': '',
                                'passAttachFolder': ''
                                'UHBAddress': ''}
        """
        if email_details is None:
            email_details = {}
        if not os.access(out_file, os.W_OK):
            print('Can''t write to output file. Please close ' + str(out_file))
            return
        ol = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = ol.GetDefaultFolder(6)  # 6=olFolderInbox
        for no in range(inbox.Items.Count-1, -1, -1):
            msg = inbox.Items[no]
            # If look_for found in subject line
            if msg.Subject.lower().find(look_for.lower()) != -1:
                csv_file_read = open(input_file, 'r')
                approved_users_csv = csv.DictReader(csv_file_read)
                found_email_in_spreadsheet = False
                for row in approved_users_csv:
                    try:
                        # Exchange users are a pain - you have to look up their email
                        if msg.SenderEmailType == "EX":
                            from_address = msg.Sender.GetExchangeUser().PrimarySmtpAddress
                        else:
                            from_address = msg.SenderEmailAddress
                    except:
                        # msg.SenderEmailType raises error on (Message Recall mails i think)
                        continue
                    if from_address.lower().strip() == row['email'].lower().strip():
                        found_email_in_spreadsheet = True
                        # Add to ICE, write to outputFile
                        csv_file_write = open(out_file, 'a', newline='')
                        writer = csv.DictWriter(csv_file_write, approved_users_csv.fieldnames)
                        print('---\nProcessing: Firstname: ' + row['firstName'] + ' Surname: ' + row['surname'] +
                              ' Email: ' + row['email'].lower().strip())
                        if not self.password_validates(row['newPassword']):
                            comment = "ICE won't accept this password even if i try it!"
                            # Shouldn't ever get here if input_file is validated, but might want to
                            # consider emailing someone the fail here?
                        else:
                            try:
                                self.login()
                            except:
                                print('Warning - cant log into ICE, I need to tell someone!')
                                # TODO Email simon
                                return
                            comment = self.add_user(row)
                            self.log_off()
                        print(comment)
                        # i = datetime.now()
                        # row['Status'] = comment + ' (%s/%s/%s %s:%s)' % (i.day, i.month, i.year, i.hour, i.minute)
                        row['Status'] = comment + ' (' + datetime.now().strftime('%d %b %Y %H:%M') + ')'  # 01 Jan 1900 19:00
                        writer.writerow(row)
                        csv_file_write.close()
                        # Email end user
                        if comment == 'User added successfully':
                            # Email username
                            self.email_out(row, email_details['userHTMLFile'], email_details['userSubject'],
                                           email_details['userAttachFolder'], email_details['fromAddress'])
                            time.sleep(30)
                            # Email password
                            self.email_out(row, email_details['passHTMLFile'], email_details['passSubject'],
                                           email_details['passAttachFolder'], email_details['fromAddress'])
                            # Email UHB
                            # forward_message = msg.Forward()
                            # forward_message.To = email_details['UHBAddress']
                            # forward_message.Send()
                            # Move original email
                            msg.Move(inbox.Folders(email_details['processed_folder']))
                        else:  # Failed to add successfully
                            msg.Move(inbox.Folders(email_details['failed_folder']))
                if not found_email_in_spreadsheet:
                    msg.Subject += ' [email address not found in spreadsheet]'
                    msg.Save()
                    msg.Move(inbox.Folders(email_details['failed_folder']))
            else:  # search term not found in subject line
                msg.Subject += ' [search term not found in subject line]'
                msg.Save()
                msg.Move(inbox.Folders(email_details['failed_folder']))

    @staticmethod
    def email_out(row, html_file, subject, attach_folder, from_address):
        """
        Will process out $username, $password in HTML and send out email
        """
        raw_html = open(html_file, 'r').read()
        html = raw_html.replace('$firstname', row['firstName'])
        html = html.replace('$surname', row['surname'])
        html = html.replace('$username', row['username'])
        html = html.replace('$password', row['newPassword'])
        html = html.replace('email', row['email'])
        attachments = []
        if os.path.exists(attach_folder):
            attachments = [os.path.join(os.getcwd(), attach_folder, fn) for fn in os.listdir(attach_folder)]
        o = Outlook()
        o.send(True, row['email'], subject, '', html, attachments=attachments, account_to_send_from=from_address)

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
            search_text = details['username'].strip()
            html_column_name = 'username'
            html_search_dropdown = 'username'
        else:
            search_text = details['firstName'].strip() + ' ' + details['surname'].strip()
            html_column_name = 'Full Name'
            html_search_dropdown = 'full name'

        Select(self.driver.find_element_by_id('SearchTypeList')).select_by_visible_text(html_search_dropdown)  # Search by
        time.sleep(1)
        self.driver.find_element_by_id('SearchTextBox').send_keys(search_text)
        self.driver.find_element_by_id('SearchButton').click()
        time.sleep(2)   # Give time for results to appear

        # Search html results: a specific column for text
        row_no = self.find_result_row(search_text, html_column_name)

        # Return if there is no match or no results or a duplicate
        if row_no == -1:  # No match or no results
            if not self.driver.find_element_by_id('chkIncludeInactive').is_selected():
                self.driver.find_element_by_id('chkIncludeInactive').click() # Show inactive users
            self.driver.find_element_by_id('SearchButton').click()
            row_no = self.find_result_row(search_text, html_column_name)
            if row_no == -1:
                return "Can't find person"
            elif row_no == -2:
                return "Duplicate person found in the system"
        elif row_no == -2:
            return "Duplicate person found in the system"

        # Result should be found. click result's full name entry (in theory it should always be the first in the list,
        # but doesn't hurt to check)
        self.driver.find_elements_by_css_selector('.dataGridRow')[row_no].find_elements_by_css_selector('td')[1].click()

        # Set password
        self.driver.find_element_by_id('PasswordTextBox').send_keys(details['newPassword'])
        self.driver.find_element_by_id('ConfirmPasswordTextBox').send_keys(details['newPassword'])
        if not self.driver.find_element_by_id('ChangePasswordCheckBox').is_selected():  # Should always be unticked on load
            self.driver.find_element_by_id('ChangePasswordCheckBox').click()
        self.driver.find_element_by_id('btnCommand').click()  # Save user
        time.sleep(1)
        try:  # If you have a success message
            if self.driver.find_element_by_id('messageDisplay').text.strip() == 'The user has been successfully updated.':
                return 'Password Updated'
            else:
                return self.driver.find_element_by_id('messageDisplay').text.strip()
        except NoSuchElementException:
            try:    # Password problem?
                if self.driver.find_element_by_id('PasswordValidator').text == \
                        'You have used this password before in your last three passwords.':
                    return "Couldn't save the user as password has been used before."
                else:
                    return self.driver.find_element_by_id('PasswordValidator').text
            except NoSuchElementException:
                try:    # Location correction
                    if self.driver.find_element_by_id('spanLocationError').text == \
                            "There must be at least one location in the user's profile.":
                        Select(self.driver.find_element_by_id('LocationListBox')).select_by_visible_text(details['location'])  #All Locations dropdown
                        self.driver.find_element_by_id('AddButton').click()
                        self.driver.find_element_by_id('btnCommand').click()  # Save user
                        time.sleep(1)
                        try:  # If you have a success message
                            if self.driver.find_element_by_id('messageDisplay').text.strip() == \
                                    "The user has been successfully updated.":
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
        csv_file_read = open(input_file, 'r')
        dict_of_rows = csv.DictReader(csv_file_read)
        csv_file_write = open(out_file, 'w', newline='')
        writer = csv.DictWriter(csv_file_write, dict_of_rows.fieldnames)

        # Process file entries
        writer.writeheader()
        for row in dict_of_rows:
            if not self.password_validates(row['newPassword']):
                comment = "ICE won't accept this password even if i try it!"
            else:
                comment = self.reset_password(row)
            print(comment)
            # i = datetime.now()
            # row['Status'] = comment + ' (%s/%s/%s %s:%s)' % (i.day, i.month, i.year, i.hour, i.minute)
            row['Status'] = comment + ' (' + datetime.now().strftime('%d %b %Y %H:%M') + ')'  # 01 Jan 1900 19:00
            writer.writerow(row)

        csv_file_read.close()
        csv_file_write.close()
        time.sleep(3)
        self.driver.quit()

    def find_result_row(self, search_text, html_column_name):
        """
        Search html page results for searchText in the html column name in the results screen
        Return: row of result, or -1 if not found, or -2 for duplicate
        """
        headers = []
        results = []
        html_header = self.driver.find_elements_by_css_selector('.header')
        for header in html_header:
            headers.append(header.text)

        html_results = self.driver.find_elements_by_css_selector('.dataGridRow')
        for id, row in enumerate(html_results):
            results.append({headers[id]:item.text for id, item in enumerate(row.find_elements_by_css_selector('td'))})

        if results:
            # Find result row, 0 = first entry
            for id, person in enumerate(results):
                if search_text.lower() == person[html_column_name].lower():
                    # Match found, but check for duplicates
                    temp = []
                    for people in results:
                        if people[html_column_name] in temp:
                            return -2   # Duplicate found
                        else:
                            temp.append(people[html_column_name])
                    return id
        return -1    # No result or no match

    @staticmethod
    def password_validates(password):
        """
        Check password is > 5 characters, has at least 1 alpha and at least 1 number
        """
        if any(char.isdigit() for char in password) \
                and any(char.isalpha() for char in password) \
                and len(password) > 5:
            return True
        else:
            return False


def start():
    ice_url_login = 'http://ocs/icedesktop/dotnet/icedesktop/login.aspx'  # live
    ice = Automate('nbf1707', 'MYPASSWORDFORLIVE', ice_url_login)  # live
    # ice_url_login = 'http://ocstrain/icedesktop/dotnet/icedesktop/login.aspx'  # test
    # ice = Automate('nbf1707', 'MYPASSWORDFORTEST', ice_url_login)  # test

    # Testing:
    # email_details = {
    #     'fromAddress': 'simon.crouch@nbt.nhs.uk',
    #     'userSubject': 'Here are your new ICE login details',
    #     'userHTMLFile': r'K:\Coding\Python\supporting files\ICE\usernameDetails.htm',
    #     'userAttachFolder': r'K:\Coding\Python\supporting files\ICE\userAttachments',
    #     'passSubject': 'Here is your new ICE password',
    #     'passHTMLFile': r'K:\Coding\Python\supporting files\ICE\passwordDetails.htm',
    #     'passAttachFolder': r'K:\Coding\Python\supporting files\ICE\passwordAttachments',
    #     'UHBAddress': 'simon.crouch@nhs.net',
    #     'processed_folder': 'Test'
    # }
    # ice.add_users_from_file_process_inbox(r'K:\Coding\Python\supporting files\ICE\ICE.csv',
    #                                       r'K:\Coding\Python\supporting files\ICE\ICE_output.csv',
    #                                       '!ICETest',
    #                                       email_details)

    email_details = {
        'fromAddress': 'MYFROM@ADDRESS.COM',
        'userSubject': 'NBT ICE Account Username',
        'userHTMLFile': 'usernameDetails.htm',
        'userAttachFolder': 'userAttachments',
        'passSubject': 'NBT ICE Account Password',
        'passHTMLFile': 'passwordDetails.htm',
        'passAttachFolder': 'passwordAttachments',
        'UHBAddress': 'SEPARATEINBOX@TOEMAIL.COM',
        'processed_folder': 'Processed',
        'failed_folder': 'Failed'
    }
    ice.add_users_from_file_process_inbox('ICE.csv',
                                          'ICE_output.csv',
                                          'yes',
                                          email_details)

if __name__ == '__main__':
    print('ICE User Addition Automation (Simon Crouch created April 2015)\nServer is running.\nv' + __version__)
    start()

"""
Test cases:
In spreadsheet, Yes = ok v'1.1.8'
In spreadsheet, no Yes = ok v'1.1.8'
Not in spreadsheet, Yes = ok v'1.1.8'
Not in spreadsheet, no Yes = ok v'1.1.8'
"""
# Compiling:
# from custom_modules.compile_helper import CompileHelp
# c = CompileHelp(r'C:\simon_files_compilation_zone\ICE')
# # c.create_env('pypiwin32 requests')
# c.freeze(r'K:\Coding\Python\custom_modules\ice.py', copy_to=r'\\nbsvr175\Scripts\ice\ice.exe')
