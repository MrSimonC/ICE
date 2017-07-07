# ICE _(Sunquest ICE Automation)_
Made for North Bristol Trust Back Office Team, whilst working as a Senior Clinical Systems Analyst.

## Background
Adding users to [Sunquest ICE](http://www.sunquestinfo.com/products-solutions/integrated-clinical-environment) is via a GUI and has no batch upload or API. Using Selenium broswer control, this program can setup multiple users in ICE from a csv file.

Using this functionality I've also created a program to email multiple users asking them to reply with "Yes" in the subject line. Then, it will monitor an Outlook inbox, lookup each sender against a .csv, if found, setup the user in ICE and email them back with a username and password.

## Prerequisite
On the server PC, install "Visual C++ Redistributable for Visual Studio 2015 x86.exe" (on 32-bit, or x64 on 64-bit) which allows Python 3.5 dlls to work, found here:
https://www.microsoft.com/en-gb/download/details.aspx?id=48145
- Ensure "Enable Protected Mode" is Disabled for all zones in IE, Tools, Options
- Ensure `IEDriver.exe` is in same directory as any .exe you run (or put in your PATH)

### Installation and Running
- "ICE Command Line" folder: run a one-off batch of users - either add/reset passwords
  - fill in `ICE.csv`
  - run `ice_cmd.exe`
- "ICE send out emails and monitor responses": email users, monitor inbox
  - to email many users:
    - fill in `ICE_to_email.csv`
    - fill in `mainMessage.htm`
    - put any attachments to the outgoing email in "mainAttachments" folder
    - run `outlook_email_many.exe`
    - results will be output to: `ICE_email_output.csv`
  - to setup users when they respond back with "Yes" in the subject line:
    - fill in `ICE.csv`
    - edit `usernameDetails.htm` with your message
    - edit `passwordDetails.htm` with your message
    - run `ice.exe`
    - results will be output to: `ICE_output.csv`
    - optionally use Windows Task Scheduler to run `ice.exe` periodically (create a new schedule, or import `ice.xml` and use that as your base task)
- Or use [ice.py](ice.py) to create your own Python program

_Written by:_  
_Simon Crouch, April 2015 in Python 3.5_