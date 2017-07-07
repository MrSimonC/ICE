import ice
import os
import getpass
import sys
import threading


def main():
    print("ICE Automate Program v" + ice.__version__ + " (by Simon Crouch, IM&T Mar2015). "
                                                                      "\nPress Ctrl+c to cancel at anytime.\n")
    input_file = input("File path to csv file to process [ICE.csv]: ") or r'ICE.csv'
    input_file = input_file.replace("\"", "")  # strip "s. Necessary for windows!
    if not os.access(input_file, os.R_OK):
        print("Can't read  " + input_file + ". Do you have it open? Exiting.")
        sys.exit(2)
    file_name, file_extension = os.path.splitext(input_file)   #output
    out_file = file_name + '_output.csv'
    if not os.access(out_file, os.W_OK):
        print("Can't write to " + out_file)
        sys.exit(2)
    user_name = input("Your ICE login username [" + getpass.getuser() + "]: ") or getpass.getuser()
    password = getpass.getpass("Your ICE login password: ")
    if not user_name or not password:
        print("No email/password given. Exiting.")
        sys.exit(2)
    operation = input("Operation (Type 'reset' or 'add'): ").lower()
    if operation != 'add' and operation != 'reset':
        print("Couldn't work out operation. Try 'add' or 'reset'.")
        sys.exit(2)
    live_or_test = input("Which environment? (Type 'live' or '[train]'): ") or 'train'
    if live_or_test == 'live':
        url_login = "http://ocs/icedesktop/dotnet/icedesktop/login.aspx"     #live
    else:
        url_login = "http://ocstrain/icedesktop/dotnet/icedesktop/login.aspx"    #train/test
    ice = ice.Automate(user_name, password, url_login)
    if operation == 'reset':
        try:
            ice.login()
        except:
            print("Login failed. Is your password correct?")
            sys.exit(0)
        ice.reset_passwords_from_file(input_file, out_file)
        input("Press Enter to exit.")
        sys.exit(0)
    elif operation == 'add':
        try:
            ice.login()
        except:
            print("Login failed. Is your password correct?")
            sys.exit(0)
        ice.add_users_from_file(input_file, out_file)
        input("Press Enter to exit.")
        sys.exit(0)

main()


def run_every_60_seconds():
    threading.Timer(60.0, run_every_60_seconds).start()
    # Do something
