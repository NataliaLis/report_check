import os
import sys
from datetime import datetime
from datetime import timedelta
import operator
from operator import itemgetter
import itertools

from prettytable import PrettyTable
import pprint

import report_read_write as rw
import tvc_calc as calc
from tvc_edit import DIR_NAME, subject, EMAIL_INTRO, email_greetings, ARCHIVE, SENT_FROM, FILE_NAME
from tvc_edit import today as today
import send_email.send_email as send
from logging_tvc import logger


def colorfy_str(msg, color):
    '''Add color to string

    Keyword arguments:
    msg -- string that need to be coloured
    color -- choosen color (red, blue, green, grey, cyan)'''

    colors = {'red': '\033[0;31m', 'blue': '\033[0;34m', 'green': '\033[0;32m', 'grey': '\033[0;90m',
              'cyan': '\033[0;96m', 'reset': '\033[0;0m'}
    return colors.get(color, f'\nNo such color available\nChoose one from: {",".join(colors.keys())}\n') \
            + msg + colors['reset']

def file_newest(dir_path: str, file_name: str, ext: str = 'xlsx'):
    """Find the latest file in the directory

    Keyword arguments:
    dir_path -- path to directory when files are stored
    file_name -- how the name of the file starts
    ext -- file extension"""
    if '.' not in ext: ext + '.'
    
    file_list = [os.path.join(dir_path, file) for file in os.listdir(dir_path)
                 if file.startswith(file_name) and file.endswith(ext)]
    
    file_latest = max(file_list, key=os.path.getctime)
    return file_latest


def str_to_date(date_str: str):
    """Changes string into datetime variable"""
    date_from_str = [int(d) for d in date_str.split('-')]
    return datetime(*date_from_str)


def find_mon_fri(friday):
    """Return tuple (or list of tuples) that contains dates of monday and friday"""
    if isinstance(friday, str):
        friday = str_to_date(friday)
        week = (friday.date() - timedelta(days=4), friday.date())
        return week
    elif isinstance(friday, datetime):
        week = (friday.date() - timedelta(days=4), friday.date())
        return week
    elif isinstance(friday, list):
        fridays = [str_to_date(fri) for fri in friday]
        weeks = [(fri.date() - timedelta(days=4), fri.date()) for fri in fridays]
        return weeks


COL_CONSOLE = ['Person', 'Claimed', 'Required', 'Missing']
COL_EMAIL = ['Week', 'Claimed', 'Required', 'Missing']

logger.info(f'Calculations made for {today}')
logger.info(f'File path: {DIR_NAME}')

try:
    os.path.exists(DIR_NAME)

    report_latest = file_newest(DIR_NAME, FILE_NAME)
    print('\n' + 'You are working with the latest report: '.rjust(45) + f'{os.path.basename(report_latest)}')
    print('Calculations made for: '.rjust(45) + today.strftime("%d-%m-%Y") + '\n')

    if report_latest:
        report_data, report_fridays = rw.read_excel(report_latest)

        pp = pprint.PrettyPrinter(indent=1)
        logger.info('Latest report:\n' + str(pp.pprint(report_data)) + '\n')

        report_weeks = find_mon_fri(report_fridays)
        tvc_data_miss = calc.tvc_miss(data=report_data, weeks=report_weeks)

        if tvc_data_miss:
            for week, items in itertools.groupby(tvc_data_miss, operator.itemgetter('week')):

                print('\n' + f'+ {week -timedelta(days=4)} - {week} +')
                tvc_by_week = PrettyTable(COL_CONSOLE)

                for item in list(items):
                    tvc_by_week.add_row([item['username'], str(item['tvc_claim']), str(item['tvc_val']),
                                             str(item['hour_miss']) + ' h ' + str(item['min_miss']) + ' m'])
                print(tvc_by_week)
        else:
            print(colorfy_str('\nNo missing TVC data - we are ', 'info') + colorfy_str('green', 'ok') + colorfy_str('!', 'blue'))

        if ARCHIVE:
            print(colorfy_str('\nArchiving in progress', 'grey'))
            rw.tvc_archive(report_latest, tvc_data_miss)
            print(colorfy_str('Done', 'grey'))

        try:
            email_ok = input(colorfy_str('\nDo you want to inform user(s) about missing TVC data via e-mail? ', 'blue'))

            while True:
                if email_ok.lower() == 'y':
                    tvc_data_miss = sorted(tvc_data_miss, key=itemgetter('email'))

                    for email, items in itertools.groupby(tvc_data_miss, operator.itemgetter('email')):
                        logger.info('Send to: ' + email)
                        tvc_summary = PrettyTable(COL_EMAIL)
                        logger.info(type(items))
                        for i in list(items):
                            logger.info('Missing in week ' + str(i['week']) + ': ' + str(i['hour_miss']) + ' hours '
                                    + str(i['min_miss']) + ' minutes\n')
                            tvc_summary.add_row([str(i['week']), str(i['tvc_claim']), str(i['tvc_val']),
                                                     str(i['hour_miss']) + ' h ' + str(i['min_miss']) + ' m'])
                        email_message = f'{EMAIL_INTRO} {tvc_summary} {email_greetings}'
                        logger.info(f'E-mail subject: {subject}')
                       # send.send_email(SENT_FROM[0], [email.lower()], subject, [SENT_FROM[0]], email_message,
                       #                     'text')
                        logger.info(f'E-mail body:\n{email_message}')
                    print(colorfy_str('\nDone\n', 'grey'))
                    break
                elif email_ok.lower() == 'n':
                    print(colorfy_str('\nNo data sent on user request\n', 'red'))
                    break
                else:
                    email_ok = input(colorfy_str('Choose proper option: y or n: ', 'blue'))
        except (KeyboardInterrupt, SystemExit):
            sys.stderr.write(colorfy_str('\n\nKeyboard interrupt\n\n', 'red'))

except (OSError, FileNotFoundError, FileExistsError):
    logger.exception(colorfy_str(f'\nNo Reports available\nCheck DIR_NAME variable ({DIR_NAME})'
                             f' or file name in TVC_edit.py\n\n', 'red'))
except (ValueError, TypeError):
    logger.exception(colorfy_str(f'\nNo TVC reports in {DIR_NAME}.\n'
                             f'Check if report name start with "{FILE_NAME}"\n', 'red'))
