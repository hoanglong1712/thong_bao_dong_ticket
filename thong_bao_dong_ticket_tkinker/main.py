# import win10toast

from openpyxl import Workbook
import datetime
import openpyxl
import sys
import os
import time
import schedule
import time
from tkinter import *
from sys import exit

# flag turns to False when there isn't any uncompleted task in input.xlsx
ongoing = True


def notification(message, title):
    """
        create a  popup windows
    :param message:
    :param title: title of the window
    """
    root = Tk()
    label = Label(root, text=message)
    label.pack()

    root.title(title)

    root.geometry('400x50+700+500')
    root.mainloop()
    pass


def turn(file_name):
    """

    :param file_name: the excel file
    """
    book = openpyxl.load_workbook(file_name)
    sheet = book.active
    row_index = 1

    # column definitions
    user_col = 'A'
    time_col = 'B'
    code_col = 'C'
    desc_col = 'D'
    status_col = 'E'

    # get current time
    current_time = datetime.datetime.now()

    t_time = datetime.datetime.now()

    user_str = ""
    global ongoing
    ongoing = False

    # iterate over the sheet
    while row_index <= sheet.max_row:
        data_time = sheet[f'{time_col}{row_index}'].value
        status = sheet[f'{status_col}{row_index}'].value
        tmp = sheet[f'{user_col}{row_index}'].value

        # update user string
        if tmp is not None:
            user_str = tmp
            pass
        # if the cell of time is there and
        # the work hasn't been done
        if data_time is not None and status is None:
            # get it expected time to be closed
            t_time = t_time.replace(hour=data_time.hour, minute=data_time.minute) + datetime.timedelta(hours=2)
            # if it is the right time to close the ticket
            if current_time >= t_time:
                code_str = sheet[f'{code_col}{row_index}'].value
                desc_str = sheet[f'{desc_col}{row_index}'].value
                # show a notification
                notification(f'{desc_str}', f'{code_str} - {user_str}')
                ongoing = True
                pass
            pass

        row_index += 1
        pass

    pass


def main(a_file_name):
    """
        main function
    :param a_file_name:
    """
    # job run every 1 minute
    job = schedule.every(1).minutes.do(turn, file_name=a_file_name)
    global ongoing
    while True:
        # pending
        schedule.run_pending()
        time.sleep(1)
        # cancel the job when all ticket has been solved
        if not ongoing:
            schedule.cancel_job(job)
            break
    pass


if __name__ == '__main__':
    main(sys.argv[1])
    try:
        # main(sys.argv[1])
        # smain('input.xlsx')
        pass
    except:
        print('thong_bao_dong_ticket input.xlsx')
        pass
