from plyer import notification

# pyinstaller --onefile --hidden-import plyer.platforms.win.notification main.py
# pyinstaller --onefile --hidden-import plyer.platforms.linux.notification main.py

from openpyxl import Workbook
import datetime
import openpyxl
import sys
import os
import time
import schedule
import time


def turn(file_name):
    book = openpyxl.load_workbook(file_name)
    sheet = book.active
    row_index = 1

    user_col = 'A'
    time_col = 'B'
    code_col = 'C'
    desc_col = 'D'
    status_col = 'E'

    # get current time
    current_time = datetime.datetime.now()

    t_time = datetime.datetime.now()

    user_str = ""
    while row_index <= sheet.max_row:
        data_time = sheet[f'{time_col}{row_index}'].value
        status = sheet[f'{status_col}{row_index}'].value
        if data_time is not None and status is None:
            t_time = t_time.replace(hour=data_time.hour, minute=data_time.minute) + datetime.timedelta(hours=2)
            if current_time >= t_time:
                tmp = sheet[f'{user_col}{row_index}'].value
                if tmp is not None:
                    user_str = tmp
                code_str = sheet[f'{code_col}{row_index}'].value
                desc_str = sheet[f'{desc_col}{row_index}'].value
                # create an object to ToastNotifier class

                notification.notify(
                    title=f'{code_str}: {user_str}',
                    message=f'{desc_str}',
                    # displaying time
                    timeout=60
                )
                time.sleep(5)
                pass
            pass

        row_index += 1
        pass
    pass



def main(a_file_name):
    schedule.every(1).minutes.do(turn, file_name=a_file_name)
    while True:
        schedule.run_pending()
        time.sleep(1)
    pass

if __name__ == '__main__':
    main(sys.argv[1])
    try:
        #main(sys.argv[1])
        #smain('input.xlsx')
        pass
    except:
        print('thong_bao_dong_ticket input.xlsx')
        pass

