import time
import win32com
from openpyxl import load_workbook
import warnings
from win32com import client
from datetime import datetime

from telethon import TelegramClient, sync, events, functions, types
from telethon.tl.functions.contacts import ImportContactsRequest
from telethon.tl.types import PeerUser, PeerChat, PeerChannel, InputPhoneContact



#получаем номер телефона из текста сообщения
def phone_output(order):
    phone_index = order.find('Телефон:')
    phone_output = order[(phone_index + 9):(phone_index + 22)]
    return phone_output

#получаем значение длины забора из текста сообщения
def length_output(order):
    length_index = order.find('Длина забора')
    length_slice = order[length_index:]
    length_output = int(length_slice.split()[3])
    return length_output

#текущее время для использования в названии pdf-файла
def current_time():
    current_datetime = datetime.now()
    time = f'{current_datetime.day}-{current_datetime.month}' \
           f'-{current_datetime.year}--{current_datetime.hour}' \
           f'-{current_datetime.minute}-{current_datetime.second}'
    return time

#вставляем значение длины из заявки в эксель файл
def set_excel_value_length(length_output_):
    wb = load_workbook('смета.xlsx')
    sheet = wb['штакетник']
    sheet['J1'] = length_output_
    wb.save('смета.xlsx')

#выделяем диапазон со сметой и сохраняем его в pdf
def excel_to_pdf(time, phone_output, length_output):
    # Open Microsoft Excel
    excel = win32com.client.Dispatch("Excel.Application")
    # Read Excel File
    sheets = excel.Workbooks.Open('C:/python_projects/rpa_orders/смета.xlsx')
    work_sheets = sheets.Worksheets['штакетник']
    work_sheets.Range("A1:E27").ExportAsFixedFormat(0, f'C:/python_projects/rpa_orders/смета_{time}_'
                                                       f'{phone_output[1:]}_'
                                                       f'{length_output}_м.pdf', OpenAfterPublish=True)
    sheets.Close(True)
    excel.Quit()

#Отправка и получение сообщений из телеграм посредством бибилотеки Telethon
# Вставляем api_id и api_hash
api_id = '*********'
api_hash = '***********'
#запуск клиента
client = TelegramClient('session_name', api_id, api_hash)
client.start()
#реагируем только на сообщение с чата с заявками, "****number*** - заменить на актуальный"
orders_channel_entity = client.get_entity(PeerChannel(****number***))

#установка на событие - новое сообщение из
@client.on(events.NewMessage(orders_channel_entity))
async def normal_handler(event):
    #получаем текст сообщения из чата
    order = event.message.to_dict()['message']
    #номер телефона
    phone_number = phone_output(order)
    #длина забора
    length = length_output(order)

    #создание нового контакта
    result = await client(ImportContactsRequest([InputPhoneContact(
        client_id=1,  # любой ид
        phone=phone_number,
        first_name=phone_number, last_name=str(length),
    )]))

    time.sleep(2)
    #отправляем значение в эксель
    set_excel_value_length(length)
    #текущее время для названия pdf-файла
    time_file_name = current_time()
    #сохранение диапазона эксель таблицы в pdf
    excel_to_pdf(time_file_name, phone_number, length)
    time.sleep(5)
    #получаем сущность для созданного контакта, чтобы потом по ней к нему обратиться и отправить сообщение
    contact_added = await client.get_entity(phone_number)
    print(f'{contact_added=}')
    #отправка текстового сообщения
    await client.send_message(contact_added, "Здравствуйте. Расчет стоимости Вашего забора:")
    #отправка pdf-файла
    await client.send_file(contact_added, f'смета_{time_file_name}_'
                                          f'{phone_number[1:]}_'
                                          f'{length}_м.pdf')

#запуск клиента в цикле пока не отключим
client.run_until_disconnected()


#узнаем доступы к конкретному чату
# @client.on(events.NewMessage())
# async def handler_all(event):
#     chat_id = event.chat_id  # ID чата
#
#     sender_id = event.sender_id  # Получаем ID Юзера
#     msg_id = event.id  # Получаем ID сообщения
#
#     sender = await event.get_sender()  # получаем имя юзера
#     name = utils.get_display_name(sender)  # Имя Юзера
#
#     chat_from = event.chat if event.chat else (await event.get_chat())  # получаем имя группы
#     chat_title = utils.get_display_name(chat_from)  # получаем имя группы
#
#     print(f"ID: {chat_id} {chat_title} >>  (ID: {sender_id})  {name} - (ID: {msg_id}) {event.text}")
#
#
# with client:
#     client.run_until_disconnected()