from custom_modules.file import Tools
import custom_modules.ice as ICE
import os
import getpass
import sys


def Main():
    print("ICE Automator Program v" + ICE.__version__ + " (by Simon Crouch, IM&T Mar2015). \nPress Ctrl+c to cancel at anytime.\n")
    inputFile = input("File path to csv file to process [ICE.csv]: ") or r'ICE.csv'
    inputFile = inputFile.replace("\"", "") #strip "s. Necessary for windows!
    fh = Tools()
    if not fh.check_file(inputFile):
        print("Can't read  " + inputFile + ". Do you have it open? Exiting.")
        sys.exit(2)
    fileName, fileExtension = os.path.splitext(inputFile)   #output
    outFile = fileName + '_output.csv'
    if not fh.check_file(outFile, True):
        print("Can't write to " + outFile)
        sys.exit(2)
    userName = input("Your ICE login username [" + getpass.getuser() + "]: ") or getpass.getuser()
    password = getpass.getpass("Your ICE login password: ")
    if not userName or not password:
        print("No email/password given. Exiting.")
        sys.exit(2)
    operation = input("Operation (Type 'reset' or 'add'): ").lower()
    if operation != 'add' and operation != 'reset':
        print("Couldn't work out operation. Try 'add' or 'reset'.")
        sys.exit(2)
    liveOrTest = input("Which environment? (Type 'live' or '[train]'): ") or 'train'
    if liveOrTest == 'live':
        urlLogin = "http://ocs/icedesktop/dotnet/icedesktop/login.aspx"     #live
    else:
        urlLogin = "http://ocstrain/icedesktop/dotnet/icedesktop/login.aspx"    #train/test
    ice = ICE.Automator(userName, password, urlLogin)
    if operation == 'reset':
        try:
            ice.login()
        except:
            print("Login failed. Is your password correct?")
            sys.exit(0)
        ice.reset_passwords_from_file(inputFile, outFile)
        input("Press Enter to exit.")
        sys.exit(0)
    elif operation == 'add':
        try:
            ice.login()
        except:
            print("Login failed. Is your password correct?")
            sys.exit(0)
        ice.add_users_from_file(inputFile, outFile)
        input("Press Enter to exit.")
        sys.exit(0)

Main()

import threading
def runEverySixtySeconds():
  threading.Timer(60.0, runEverySixtySeconds).start()
  # Do something
