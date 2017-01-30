# this is a script that connects to Exchange to grab mail attachments
# written by MC 30.01.2017 during my time with Nokia - dub2sauce[at]gmail[dot]com
#  make sure you have @on.nokia.com domain, otherwise it won't work

user='someone@somewhere.com'
password='topsecret'					
SaveLocation = 'C:/Users/admin/Downloads/dist/exchangelib-1.7.6/test/'			# where you want the file(s) saved
SubjectSearch = 'FIR';																					# it will search all emails containing this string

from exchangelib.configuration import Configuration
from exchangelib import DELEGATE, IMPERSONATION, Account, Credentials, EWSDateTime, EWSTimeZone, Configuration, NTLM, CalendarItem, Q
from exchangelib.folders import Folder
import logging,os.path

config = Configuration(
	server='x.x.x.x',
    credentials = Credentials(
    username=user,
    password=password),
    verify_ssl=False
	)

account = Account(
    primary_smtp_address='mailbox@company.com',
    autodiscover=False, 
	  config=config,
    access_type=DELEGATE)

for item in account.inbox.all():
    if SubjectSearch in item.subject:
        for attachment in item.attachments:
            local_path = os.path.join(SaveLocation, attachment.name)
            with open(local_path, 'wb') as f:
                 f.write(attachment.content)
                 print('Saved attachment to', local_path)
