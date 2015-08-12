__version__ = 1.1
import win32com.client
import datetime

class Outlook:
    def __init__(self):
        # http://msdn.microsoft.com/EN-US/library/microsoft.office.interop.outlook.mailitem_members(v=office.14).aspx
        self.outlook = win32com.client.Dispatch("Outlook.Application")

    def get_first_message_body(self):
        # http://stackoverflow.com/questions/5077625/reading-e-mails-from-outlook-with-python-through-mapi
        mapi = self.outlook.GetNamespace("MAPI")
        inbox = mapi.GetDefaultFolder(6) #olFolderInbox
        messages = inbox.Items
        #message = messages.GetFirst()
        #message = messages.GetLast()
        for message in messages:
            print(message.SenderEmailAddress)
            print(message.Subject)
        #body_content = message.subject + "\n" + message.body
        #return body_content

    def move_to_folder(self, olMailItem, folderName):   #sac- does this work?
        olMailItem.Move(self.outlook.Folders(folderName))

    # Sends an email. Returns: True/False.
    def send(self, Send,
        to,
        subject,
        bodyClearText,
        bodyHTML="",
        bodyFormat=2,
        fromEmail="",
        cc="",
        bcc="",
        replyRecipients="",
        flagText="",
        reminderTFalse=False,
        reminderDateTime="",
        importance=1,
        readReceipt=False,
        deferredDeliveryDateTime="",
        accountToSendFrom="",
        attachments=[]):
        """
        e.g.
        #o = Outlook()
        #mobile = "07948548409"
        #o.send(True, mobile + "@sms.drdoctor.co.uk", "SMS sent to patient mobile: " + mobile, "Hey dad - just testing my program from work! Have fun!--end")
        #o.send(False, "simon.crouch@nbt.nhs.uk", "Test subject", "Test sms message programmatic 04--end") #, attachments=[r'C:\\Users\nbf1707\Desktop\Galaxy SOP setup.txt', r'C:\simon files\pythonFromK\dist\ICE.csv'])
        #o.getFirstMessageBody()
        """
        item = self.outlook.CreateItem(0)    # olMailItem
        item.To = to
        item.Subject = subject
        if bodyClearText:   # default signature is kept if you don't change the .Body property
            item.Body = bodyClearText
        if bodyHTML:
            item.HTMLBody = bodyHTML
        item.BodyFormat = bodyFormat	# 1=plain text, 2=html, 3=rich text
        item.SentOnBehalfOfName = fromEmail		# sets the "From" field. "" is ok as Outlook just uses default account
        item.CC = cc		# == Recipient1 = sacComObjectReply.Recipients.Add("a@a.com") then Recipient1.Type = 2 #To=1 Cc=2 Bcc=3
        item.BCC = bcc
        item.FlagRequest = flagText 	# sets follow up flag *for recipients*! very cool. "" is ok
        item.ReplyRecipientNames = replyRecipients		# sac, depreciated, should use item.ReplyRecipients.Add("simon crouch")
        if reminderDateTime:
            item.ReminderTime = reminderDateTime
        item.ReminderSet = reminderTFalse
        item.ReadReceiptRequested = readReceipt
        item.Importance = importance	# 2=high, 1=med, 0=low
        if deferredDeliveryDateTime:    # '%d/%m/%y %H:%M'
            item.DeferredDeliveryTime = deferredDeliveryDateTime
        if accountToSendFrom:
            item.SendUsingAccount = item.Application.Session.Accounts.Item(accountToSendFrom)	#sac- this took me ages. You can also use a #
        if attachments:
            [item.Attachments.Add(attachment) for attachment in attachments]
        if Send:
            item.Send()
        else:
            item.Display()
        return True

    def outlook_repeat_delay_email(self, to, sub, message, delay_date, repeatCount=1, daysApart=0):
        delay_date = datetime.datetime.strptime(delay_date, '%d/%m/%y %H:%M')
        for i in range(0, repeatCount):
            dateFormatted = delay_date.strftime('%d/%m/%y %H:%M')
            sub = sub + " (" + dateFormatted + ")"
            self.send(True, to, sub, message, deferredDeliveryDateTime=dateFormatted)
            delay_date += datetime.timedelta(days=daysApart)

    def appointments_before_0930(self, days_forward=7):
        """
        e.g.
        for item in apptsBefore930(7):
            print(item.Subject)

        # AppointmentItem Members (Outlook)
        # https://msdn.microsoft.com/EN-US/library/office/ff869026(v=office.14).aspx
        """
        mapi = self.outlook.GetNamespace("MAPI")
        oCalendar = mapi.GetDefaultFolder("9")   # olFolderCalendar https://msdn.microsoft.com/en-us/library/office/ff861868.aspx
        oItems = oCalendar.Items
        oItems.IncludeRecurrences = True
        oItems.Sort("[Start]")

        startDate = datetime.datetime.now().strftime("%d/%m/%Y %H:%M %p")
        endDate = (datetime.datetime.now() + datetime.timedelta(days=days_forward)).strftime("%d/%m/%Y %H:%M %p")
        oItems = oItems.Restrict("[Start] >= '" + startDate + "' AND [Start] <= '" + endDate + "'")

        # Find next week appointments, earlier than 09.30am
        for item in oItems:
            if item.Start.time() < datetime.time(9, 30) \
                    and not item.Start.time() == datetime.time(0, 0)\
                    and not item.Categories == "FreeTime":
                #print(item.Start, ' ', item.Subject, ' ', item.Categories, ' ', item.Location)
                yield item
