import imaplib
from imapclient import imap_utf7
from time import sleep
from tqdm import tqdm, contrib
from colorama import init, Fore, Style, just_fix_windows_console
import csv
import json

# Настройки

auto_delete_list = ['spam', 'trash', 'drafts|template', 'спам', 'удаленные']  # Папки для исключения из копирования
copy_mode = "auto"  # auto или manual
since_date = "01-Jan-1990"  # Дата для начала копирования
before_date = "11-Apr-2023"  # Дата окончания копирования

init()
just_fix_windows_console()

# загружаем csv файл с учётными записями
csv_file_path = 'data.csv'
json_log_file = 'log.json'

with open(csv_file_path, encoding="UTF-8") as file:
    file_reader = csv.reader(file, delimiter=",")

    accounts_list = []
    for row in file_reader:
        accounts_list.append(row)


def write_log(account, mailbox, last_email):
    data = read_log()

    if not data:
        data = {"accounts": {account: {'last_mail': {mailbox: last_email}}}}
    else:
        if account not in data['accounts']:
            data['accounts'].update({account: {'last_mail': {mailbox: last_email}}})
        else:
            data['accounts'][account]['last_mail'].update({mailbox: last_email})   

    json_data = json.dumps(data)
    with open(json_log_file, 'w') as jwf:
        jwf.write(json_data)


def read_log():
    try:
        with open(json_log_file, 'r') as jrf:
            data = jrf.read()
            if not data:
                return None
            else:
                new_data = json.loads(data)

            return new_data
    except FileNotFoundError:
        print(f'Не существует лог файл...Создаём...')
        open(json_log_file, 'a')


# Проверка логинов и паролей на возможность подключения к серверу
def account_check(accounts_list):

    for account in accounts_list:
        src_server = account[0]
        src_username = account[1]
        src_password = account[2]
        dest_server = account[3]
        dest_username = account[4]
        dest_password = account[5]
        try: 
            src_conn = imaplib.IMAP4_SSL(src_server)
            dest_conn = imaplib.IMAP4_SSL(dest_server)            # src_conn.login(src_username, src_password)
            print(Fore.YELLOW + f'Подключение к целевому аккаунту {src_username} :', src_conn.login(src_username, src_password)[0] + Style.RESET_ALL)            
            print(Fore.BLUE + f'Подключение к удалённому аккаунту {dest_username} :', dest_conn.login(dest_username, dest_password)[0] + Style.RESET_ALL)
            src_conn.logout()
            dest_conn.logout()
        except Exception as e:
            print(Fore.RED + f'Произошла ошибка при проверке аккаунта {src_username}, {dest_username}: {str(e)}' + Style.RESET_ALL)


def copy_emails(src_server, src_username, src_password, dest_server, dest_username, dest_password, folder, copy_start):
    # Установка соединения с исходным IMAP-сервером
    src_conn = imaplib.IMAP4_SSL(src_server)
    src_conn.login(src_username, src_password)

    # Установка соединения с целевым IMAP-сервером
    dest_conn = imaplib.IMAP4_SSL(dest_server)
    dest_conn.login(dest_username, dest_password)

    try:

        # Выбор папки на исходном сервере
        src_conn.select(folder)

        # Проверка существования папки на удалённом сервере
        check = dest_conn.select(folder)[0]
        if check != 'OK':
            print("Не найдена папка на удалённом сервере. Создаём....")
            dest_conn.create(bytes(folder, "utf-8"))
            print(f'Папка {folder_decode(bytes(folder, "utf-8"))} создана')
        else:
            print(f' Папка {folder_decode(bytes(folder, "utf-8"))} найдена на удалённом сервере')
            print("Копируем из папки:", folder_decode(bytes(folder, "utf-8")))
        # Получение списка всех писем в папке на исходном сервере
        typ, data = src_conn.search(None, f'(BEFORE "{before_date}")')
        # Вывод информации о количестве писем в папке
        email_count = len(data[0].split())
        print(f'Найдено {email_count} писем для копирования.')
        copy_start = 0
        # Итерация по всем письмам в папке на исходном сервере и копирование их на целевой сервер
        for i, num in contrib.tenumerate(data[0].split(), ncols=80, ascii=True, desc=f'Копируем из {folder_decode(bytes(folder, "utf-8"))}', start=copy_start):
            typ, msg_data = src_conn.fetch(num, '(RFC822)')
            dest_conn.append(folder, None, None, msg_data[0][1])
            # Записываем в лог файл
            write_log(src_username, folder, int(num))

    except Exception as e:
        print(f'Произошла ошибка при копировании: {str(e)}')

    finally:
        # Закрытие соединений с обоими серверами
        src_conn.select(folder)
        src_conn.close()
        src_conn.logout()
        dest_conn.select()
        dest_conn.close()
        dest_conn.logout()


def folder_decode(folder):
    folder = imap_utf7.decode(folder).replace('"', '')
    return folder


# Функция получения списка папок для копирования
def folders_list(src_server, src_username, src_password):
    # Установка соединения с исходным IMAP-сервером
    src_conn = imaplib.IMAP4_SSL(src_server)
    src_conn.login(src_username, src_password)

    try:
        # Выбор папки на исходном сервере
        new_folders = []
        for folder in (src_conn.list()[1]):

            n_folder = folder.decode().split('"|" ')[-1:]
            new_folders.append(*n_folder)

        print(Fore.LIGHTGREEN_EX + 'Получен список папок с целевого сервера:' + Style.RESET_ALL)
        print(f'Ищем письма с {since_date} по {before_date}')
        for i, folder in enumerate(new_folders, start=1):
            # Получение списка всех писем в папке на исходном сервере
            src_conn.select(folder)
            typ, data = src_conn.search(None, f'(BEFORE {before_date})')
            # Вывод информации о количестве писем в папке
            email_count = len(data[0].split())

            print(Fore.GREEN + f'{i}. В папке {folder_decode(bytes(folder, "utf-8"))} найдено {email_count} писем для копирования.' + Style.RESET_ALL)

    except Exception as e:
        print(f'Произошла ошибка при получении списка папок: {str(e)}')
    finally:
        # Закрытие соединений с обоими серверами
        src_conn.close()
        src_conn.logout()
        return (new_folders)


def folder_choise(folder_list):
    if copy_mode == "auto":
        # Исключаем папки для копирования
        new_folder_list = []

        for folder in folder_list:
            if folder_decode(bytes(folder, "utf-8")).lower() in auto_delete_list:
                print(Fore.RED + f'Папка {folder_decode(bytes(folder, "utf-8"))} не будет скопирована' + Style.RESET_ALL)
            else:
                new_folder_list.append(folder)

        return new_folder_list
    elif copy_mode == "manual":
        # выбираем папки для копирования
        new_list = {}
        folders_for_migration = []
        for i, j in enumerate(folder_list, start=1):
            new_list[i] = j

        delete_list = list(input("Укажите номера папок для исключения из копирования: ").split())

        if not delete_list:
            print('Будут скопированы все папки!')
        else:
            print(Fore.RED + 'Данные папки не будут скопированы' + Style.RESET_ALL)
            for i in delete_list:
                print(Fore.RED + folder_decode(new_list.get(int(i))) + Style.RESET_ALL)
                new_list.pop(int(i))
            for folder in new_list.values():
                folders_for_migration.append(folder)

        return folders_for_migration


account_check(accounts_list)    # Проверяем коректность данных для авторизации
for account in accounts_list:
    src_server = account[0]
    src_username = account[1]
    src_password = account[2]

    dest_server = account[3]
    dest_username = account[4]
    dest_password = account[5]

    print(Fore.CYAN + f'Готовимся копировать аккаунт {src_username}' + Style.RESET_ALL)
    sleep(2)
    try:
        new_folders_list = folder_choise(folders_list(src_server, src_username, src_password))
        sleep(3)
        try:

            for new_folder in new_folders_list:
                data_log = read_log()
                copy_start = 0

                if not data_log:
                    copy_emails(src_server, src_username, src_password, dest_server, dest_username, dest_password, new_folder, copy_start)
                else:
                    if src_username in data_log['accounts']:
                        if new_folder in data_log['accounts'][src_username]['last_mail']:
                            copy_start = int(data_log['accounts'][src_username]['last_mail'][new_folder])
                            copy_emails(src_server, src_username, src_password, dest_server, dest_username, dest_password, new_folder, copy_start)
                        else:
                            copy_emails(src_server, src_username, src_password, dest_server, dest_username, dest_password, new_folder, copy_start)
                    else:
                        copy_emails(src_server, src_username, src_password, dest_server, dest_username, dest_password, new_folder, copy_start)

            print("Копирование завершено!")
        except Exception as e:
            print(f'Произошла ошибка при выборе папки для копирования на удалённом сервере: {str(e)}')
    except Exception as e:
        print(f'Произошла ошибка при обработке списка папок: {str(e)}')
