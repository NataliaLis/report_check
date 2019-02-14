""" File with variables that can be changed by the user

dir_name - directory, where TVC reports are stored
sent_from - personal details of the e-mail sender
email_intro - part of e-mail body - before output from the script
email_greetings - end of e-mail body
today - date for which script makes TVC calculations, default - current day
subject - subject of the e-mail
archive - if True, script will write output to te TVC report (separate sheet), default - False
"""

from datetime import date, datetime


"""Path to the directory where TVC reports are stored

The most current file is chosen
Important: Do NOT change the original name of the file"""
DIR_NAME = '/home/plp76889/Desktop/work/1_Skrypty/report_check'

"""How the file name starts"""
FILE_NAME = 'Weekly_'

"""Credentials of person who will send e-mails"""
SENT_FROM = ('tl@example.com', 'TL Name')

"""Message to the e-mail receiver

It will appear at the beginning of the e-mail"""
EMAIL_INTRO = """Hi! 
Some of your TVC data is missing: 

"""

"""Greeting at the end of the e-mail message"""
email_greetings = """

Sincerly,
%s
""" % SENT_FROM[1]

"""This variable is used in subject and tvc calculations - means date of running script

Hint: if you use old script (eg. from yesterday, you should change this variable accordingly,
Otherwise script calculates wrong data
In this case change "date.today()" to
datetime(YYYY,M,D).date(), where
                            YYYY - year
                            M - month
                            D - day"""
today = date.today() # datetime(2018,8,22).date() #date.today()

"""E-mail subject"""
subject = f'[To do] Missing TVC summary ({today.strftime("%d-%m-%Y")})'

"""Archiving option

If you want to have output of this script archived (in the next sheet of your Weekly TVC report),
assign 'True' to this variable. Otherwise use 'False' option.
Hint: It must be capitalized"""
ARCHIVE = False
