# All links in this file were erased in sight of corporate security reasons.
# TODO items with '------V' mean 'done'.


import requests
from requests.auth import HTTPBasicAuth
import json
import xlsxwriter
from datetime import datetime
import re
import win32com.client
import win32com
import win32con
import win32api
import getpass
import os
import sys
import logging


# TODO
# =========
# 1. Put all temp files in separate folder. And work with them out of it. -----V
# 1.1 Delete all temp files before programm exit.    -----V
# 2. Input duty_report account info. Make input password hidden.    -----V
# 2.1 Store duty_report account info in hidden file in data folder.    -----V
# 2.2 Cypher account info.
# 2.3 Check authorisation on start and input new data if false.    -----V
# 3. Input name of mail folder.    -----V
# 3.1 Check existence of such folder.    -----V
# 3.2 Store folder name in data folder.    -----V
# 3.2.1! Store folder name in separate file (or remake settings file).    -----V
# 3.3 Check if folder name is in data and existence of such folder. Input name 
# if false.    -----V
# 3.3.2 Check folder name if many spaces. (Пробелы в названии папки сбивают считывание
# opl_name из файла.)    -----V
# 3.4! Check folders inlay in outlook.    -----V
# 3.5! Filter emails by date. If many mails outlook_load() works too long.
# 3.6! Exclude forward mails and replies and other trash.    -----V
# 4. Add subject and start_data info in end_output file.    -----V
# 4.1 Add hyperlinks to opl numbers.    -----V
# 4.2 Add hyperlinks to gp numbers.    -----V
# 5. Make menu.
# 5.1 Reset account info and enter new.
# 6. Improve logging.
# 7. Store all data in final file and don't rewrite it.
# 7.1 Delete duplicates.

# *optional changes

def outlook_load(what_to_do):
    global opl_folder
    global ok_folder

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    accounts = win32com.client.Dispatch("Outlook.Application").Session.Accounts;

    def email_all(folder):
        today = datetime.today()
        # start_of_day = today + ' 09:00:00'
        # start_of_day = datetime.strptime(start_of_day, '%Y-%m-%d %H:%M:%S')

        with open('./data/inbox.txt', 'w+') as f:
            messages = folder.Items
            messages.Sort('SentOn', True)
            count = 0
            for message in messages:
                try:
                    subject = message.Subject
                    sender = message.Sender
                except AttributeError:
                    continue
                sender_check = 'Operation Log Notification'
                check = 'Уведомление о назначении согласующих на плановую работу'
                check2 = 'Уведомление о создании плановой работы'
                if sender_check in str(sender) and check in str(subject):
                    body = message.Body
                    end_time = re.search(r'Дата окончания\s((\d){2}.(\d){2}.(\d){4} (\d){2}:(\d){2}:(\d){2})', body).group(1)
                    x = datetime.strptime(end_time, '%d.%m.%Y %H:%M:%S')
                    if x >= today:
                        print(f'Processing {subject} ', end = '...')
                        count += 1
                        f.write('=================================\n')
                        f.write(str(sender) + ": ")
                        f.write(subject + '\n')
                        f.write('----------------\n')
                        for i in body.splitlines():
                            try:
                                if len(i) > 2:
                                    f.write(i + '\n')
                            except:
                                print("can't write")
                        print('OK')
                    else:
                        print('Some old opl. Pass.')
            f.write('=================================\n')
            f.write('Всего сообщений: ' + str(count))


    def find_folder(folders, item):
        check = False
        def folders_walk(folders_x, item_x):
            nonlocal check
            for folder in folders_x:
                if folder.name == item_x:
                    print(f'Found {folder.name}')
                    check = folder
                if len(folder.folders) > 0:
                    folders_walk(folder.folders, item_x)
        folders_walk(folders, item)
        return check


    account = accounts[0]
    inbox = outlook.Folders(account.DeliveryStore.DisplayName)
    folders = inbox.Folders

    if what_to_do == 'check':
        while ok_folder != True:
            opl_folder = input('Outlook OPL-mail folder:')
            ok_folder = find_folder(folders, opl_folder)
            if ok_folder == False:
                print('Incorrect folder name or no such folder.')
            else:
                print('Success!')
                break
    elif what_to_do == 'load':
        folder_to_load = find_folder(folders, opl_folder)
        print('Loading mail...')
        email_all(folder_to_load)


def all_opl_sort():
    print('Sorting opl...')
    today = datetime.today()

    # class OPL():    #TODO: make opl in OOP, not in dict.
    #   def __init__(self, number, start_time, end_time):
    #       self.number = number
    #       self.start_time = start_time
    #       self.end_time = end_time

    print(datetime.strftime(today, "%Y.%m.%d %H:%M:%S"))

    with open('./data/inbox.txt', 'r') as f:
        opl_data = f.read()

    splitline = '================================='
    opl_list = opl_data.split(splitline)

    result = []

    for i in opl_list[1:-1]:
        test = i.replace('\t', ' ')
        opl_dict = {
            'number': re.search(r' Notification: ((\d)*) ', test).group(1),
            'subject': re.search(r'Предмет работы (.*)', test).group(1),
            # 'impact': '',
            'branch': re.search(r'Место проведения (.*)', test).group(1),
            'start_time': re.search(r'Дата начала ((\d){2}.(\d){2}.(\d){4} (\d){2}:(\d){2}:(\d){2})', test).group(1),
            'end_time': re.search(r'Дата окончания ((\d){2}.(\d){2}.(\d){4} (\d){2}:(\d){2}:(\d){2})', test).group(1)
        }
        result.append(opl_dict)

    print(len(result))

    with open('./data/opl_search.txt', 'w+') as f:
        for i in result:
            f.write(i['number'] + '\t' + i['start_time'] + '\t' \
             + i['end_time']  + '\t' + i['branch']  + '\t' + i['subject'] + '\n')


def to_excel(dict_to):
    print('Creating excel file...')
    today = datetime.today()
    today_time = datetime.strftime(today, "%Y.%m.%d_%H.%M")
    excel_file = 'data_' + str(today_time) + '.xlsx'
    workbook = xlsxwriter.Workbook(excel_file)
    worksheet = workbook.add_worksheet()

    row = 0
    col = 0

    d = dict_to

    worksheet.write(0, 0, 'OPL')
    worksheet.write(0, 1, 'HD')
    worksheet.write(0, 2, 'start_datetime')
    worksheet.write(0, 3, 'end_datetime')
    worksheet.write(0, 4, 'branch')
    worksheet.write(0, 5, 'subject')

    for key in d.keys():
        hyper_opl = 'erased_link' + key + 'erased_link'
        hyper_gp = 'erased_link' + str(d[key][1])
        row += 1
        worksheet.write_url(row, col, hyper_opl, string = key)
        if type(d[key][1]) == int:
            worksheet.write_url(row, col + 1, hyper_gp, string = str(d[key][0]))
        else:
            worksheet.write(row, col + 1, d[key][0])
        worksheet.write(row, col + 2, d[key][2])
        worksheet.write(row, col + 3, d[key][3])
        worksheet.write(row, col + 4, d[key][4])
        worksheet.write(row, col + 5, d[key][5])
    print('Work complete.', excel_file, 'created.')
    workbook.close()


def test_connection():
    global user_name
    global password
    global ok
    while ok != 200:
        user_name = input('UserName:')
        password = getpass.getpass('PassWord:')
        ok = duty_connection(user_name, password)
        if ok == 401:
            print('Wrong username or password.')
        elif ok == 200:
            print('Success!')
            break


def duty_connection(user_name, password):
    global duty_session
    url = 'erased_link'
    duty_session = requests.Session()
    r = duty_session.get(url, auth=(user_name, password))
    return (r.status_code)


def duty_load():
    print('Checking HD...')
    url = 'erased_link'

    payload_fttb = {
        'groupFilter': 'duty_fttb',
        'typeFilter': '%D0%9F%D0%BB%D0%B0%D0%BD%D0%BE%D0%B2%D1%8B%D0%B5+%D1%80%D0%B0%D0%B1%D0%BE%D1%82%D1%8B',
        'regionFilter': '%D0%92%D1%81%D0%B5',
        'archive': 'false',
        'source': 'index',
        'search': '',
        'order': 'asc'
    }

    payload_mag = {
        'groupFilter': 'duty_mag',
        'typeFilter': '%D0%9F%D0%BB%D0%B0%D0%BD%D0%BE%D0%B2%D1%8B%D0%B5+%D1%80%D0%B0%D0%B1%D0%BE%D1%82%D1%8B',
        'regionFilter': '%D0%92%D1%81%D0%B5',
        'archive': 'false',
        'source': 'index',
        'search': '',
        'order': 'asc'
    }

    duty_gp_fttb = duty_session.get(url, params=payload_fttb, auth=(user_name, password))
    gp_data = duty_gp_fttb.json()
    duty_gp_mag = duty_session.get(url, params=payload_mag, auth=(user_name, password))
    gp_mag = duty_gp_mag.json()

    gp_data.extend(gp_mag)

    # with open('nest_data', 'w+') as write_file:
    #     json.dump(gp_data, write_file)

    global total_dict
    total_dict = {}

    with open('./data/opl_search.txt', 'r') as f:
        opl_list = f.readlines()

        for i in opl_list:
            opl_info = i.strip().split('\t')
            search = opl_info[0]

            # REFACTOR total_dict creation. stop at some point.

            for i in gp_data:
                if search == i['extNumber'] or search in i['description']:
                    total_dict[search] = [i['number']]
                    total_dict[search].append(i['helpdeskID'])
            if search not in total_dict.keys():
                total_dict[search] = ['Нет ГП']
                total_dict[search].append('string')
            total_dict[search].append(opl_info[1])
            total_dict[search].append(opl_info[2])
            total_dict[search].append(opl_info[3])
            total_dict[search].append(opl_info[4])


# for i in total_dict:
#   print(f"OPL: {i} - HD: {total_dict[i]}")

def check_data():
    global ok
    global user_name
    global password
    global opl_folder
    global ok_folder

    # REFACTOR authentication and opl check. Devide them.

    def write_data():
        global ok
        global ok_folder
        try:
            os.remove(acc_file)
        except FileNotFoundError:
            pass
        with open(acc_file, 'w+') as f:
            win32api.SetFileAttributes(acc_file, win32con.FILE_ATTRIBUTE_HIDDEN)
            ok = 401
            test_connection()
            ok_folder = False
            outlook_load('check')
            f.write(user_name + ' ' + password + '\n' + opl_folder)


    acc_file = os.path.join(data_path, 'acc.id')
    current_user = os.getlogin()
    if not os.path.exists(acc_file) or os.path.getsize(acc_file) == 0:
        write_data()
    else:
        with open(acc_file, 'r') as f:
            usr_settings = f.readlines()
            user_name, password = usr_settings[0].strip().split(' ')
            opl_folder = usr_settings[1]
        if user_name != current_user:
            print(f"Impostor! You're not {user_name}! Relogin!")
            write_data()
    ok = duty_connection(user_name, password)
    if ok == 401:
        write_data()


def main():
    global data_path
    global logger
    user = os.getlogin()
    data_dir = 'data'
    data_path = os.path.join(os.getcwd(), data_dir)
    if not os.path.exists(data_path):
        os.mkdir(data_path)

    logger = logging.getLogger("opl_checker")
    logger.setLevel(logging.INFO)

    fh = logging.FileHandler('./data/app.log')

    formatter = logging.Formatter('%(asctime)s - ' + user + ' - %(name)s - %(levelname)s - %(message)s')
    fh.setFormatter(formatter)

    logger.addHandler(fh)

    launch_args = sys.argv
    diag = '-diag' in launch_args

    try:
        check_data()
        outlook_load('load')
    except (KeyboardInterrupt, SystemExit):
        logger.info('Mail gathering ended by user.')
        print('Mail gathering ended by user.')
    except Exception as e:
        logger.exception(str(e))
        print('Error occured! Details are in "app.log"')
    try:
        all_opl_sort()
        duty_load()
        to_excel(total_dict)
    except Exception as e:
        logger.exception(str(e))
        print('Error occured! Details are in "app.log"')


    if not diag:
        try:
            os.remove('./data/inbox.txt')
            os.remove('./data/opl_search.txt')
        except FileNotFoundError:
            pass
    else:
        logger.info('Diag mod started!')


if __name__ == "__main__":
    main()


input("Press Enter")