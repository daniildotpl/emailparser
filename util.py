import os
import re
import datetime
from django.conf import settings
from django.contrib import messages
from django.shortcuts import redirect


import csv
import shutil
import imaplib
import email
import email.message
import email.header
from bs4 import BeautifulSoup




class UpdateCsvGetMixin:
    # -----------------------------------------------
    # --- переопределяем get ------------------------
    def get(self, request, *args, **kwargs):
        if self.write(self.get_rec_list()):
            return super().get(request, *args, **kwargs)
        else:
            return redirect('error')

    # -----------------------------------------------
    # --- функция записи в файл ---------------------
    def write(self, rec_list):
        # просто назову путь к фалу 'XXX'
        file_path = 'XXX'

        # записываем наш массив данных в файл
        with open(file_path, "w", encoding='utf8', newline="") as file:
            writer = csv.writer(file, delimiter = ";")
            if writer.writerows(rec_list):
                messages.success(self.request, 'Обновлено')

        return True

    # -----------------------------------------------
    # --- функция парсинга почты и создания списка -- 
    def get_rec_list(self):
        
        # параметры соединения
        # просто заменю всё на 'XXX'
        imap_host = 'XXX'
        imap_port = 'XXX'
        imap_user = 'XXX'
        imap_pass = 'XXX'
        folder = 'XXX'
        
        # список "тегов", по которыми ищем необходимую информацию
        # он же - шапка для таблицы с заявками
        names = [
            'Дата получения', 
            'Дата создания', 
            'Заявка номер',
            'Тема', 
            'Адрес', 
            # и другие поля для поиска
        ]

        # сразу записываем их в выходной массив
        out = [names]
        
        # создаем соединение
        try:
            imap = imaplib.IMAP4_SSL(imap_host)
            if imap:
                # что б на верняка посмотрим сообщение в консоле 
                print('We got object IMAP4_SSL')
        except:
            messages.error(self.request, 'Что-то не получается')
            # если можно, посмотрим сообщение в консоле
            print('we can not IMAP4_SSL')

        # логинимся
        try:
            if imap.login(imap_user, imap_pass):
                # что б на верняка посмотрим сообщение в консоле
                print('We logined')
        except:
            messages.error(self.request, 'Что-то не получается')
            # если можно, посмотрим сообщение в консоле
            print('we can not LOGUN')
        
        # выбираем папку входящие
        try:
            sel = imap.select(folder, readonly=True)
            if sel[0] == 'OK':
                # и в консоль
                print('Folder {} selected'.format(folder))
            else:
                # в консоль
                print('Folder {} NOT FOUND'.format(folder))
        except:
            messages.error(self.request, 'Что-то не получается')
            # в консоль
            print('We can`t SELECT folder')

        # делаем дату с которой будем искать письма
        today = (datetime.date.today()).strftime("%d-%b-%Y")
        dayago = (datetime.date.today() - datetime.timedelta(1)).strftime("%d-%b-%Y")
        twodayago = (datetime.date.today() - datetime.timedelta(2)).strftime("%d-%b-%Y")

        # ищем все номера писем за сегодня и вчера (у нас 15 попыток)
        for i in range(15):
            typ, data = imap.search(None
                , '(SINCE {0})'.format(dayago)
            )
            if typ == 'OK':
                break
            else:
                pass

        # переводим номера в список
        data = data[0].split()

        # переворачиваем список в обратную сторону
        data = data[::-1]

        # проход по списку номеров писем с получением полезной нагрузки, заголовков и тд
        parts = [] # пустой список для записи туда заявок
        for i in data:
            status, data = imap.fetch(i, '(RFC822)')
            msg = data[0][1]
            msg = email.message_from_bytes(msg)
            # преобразование темы письма в нужные нам данные (order и theme)
            sub = msg["Subject"]
            # там получится какая-то дичь поэтому перекодируем её
            sub = str(email.header.make_header(email.header.decode_header(sub)))
            try:
                sub = sub.split(':')
                order = sub[0]
                theme = sub[1]
            except:
                pass
            # преобразование времени получения письма в нужные данные (time и day)
            time = msg["Date"]
            day = time.split(' ')
            day = day[1:4]
            day = '-'.join(day)
            if msg.is_multipart():
                for part in msg.walk():
                    payload = part.get_payload(decode=True).decode('utf-8')
                    parts.append(payload)
            else:
                payload = msg.get_payload(decode=True).decode('utf-8')
                parts.append(payload)
            
            # готовим суп из полезноц нагрузки
            html = str(payload) 
            html = html.replace('<br>', ' ') 
            html = html.replace('<br/>', ' ') 
            html = html.replace("\r"," ") 
            html = html.replace("\n"," ") 
            html = html.replace(";",",") 
            soup = BeautifulSoup(html, "html.parser")

            # ищем номер заявки
            try:
                number = soup.find('a')
                number = number.get_text()
            except:
                messages.error(self.request, 'Что-то не получается')
            
            # объявляем массив для строчки
            info_row = []

            # ищем подходящие нам поля в супе
            for name in names:
                # создаем и приводим в порядок инфу после заголовков
                info = ''
                # почистим заголовки от лишнего
                name = name.strip()
                try:
                    bold = soup.find('b', text=re.compile(name))
                    info = bold.next_sibling
                    try:
                        info = info.get_text()
                    except:
                        pass
                except:
                    pass
                # очищаем строки
                info = info.strip()
                info = info.replace(':', '', 1)

                # записываем всё в список
                info_row.append(info)

            # а потом этот список в ещё один список
            out.append(info_row)

            # по условию ж мы собираем заявки за текущие сутки + последняя заявка за вчера
            # отбиваемся если дошли до последнего вчерашнего письма
            # в заявке число в датах фомата с числами до 10 - 01.01.2022 а у нас даты формата - 01.01.2022
            # поэтому - костыль для подрезки 0 из даты
            dayago_exp = list(dayago)
            if dayago_exp[0] == '0':
                dayago_exp = ''.join(dayago_exp[1:])
                dayago = dayago_exp
            if day == dayago: break

        # закрываем соединение
        imap.close()

        # print(out)
        return out
