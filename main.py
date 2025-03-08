from __future__ import print_function
import vk_api
from vk_api.bot_longpoll import VkBotEventType, VkBotLongPoll
from vk_api.keyboard import VkKeyboard, VkKeyboardColor
import openai
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import re
from textblob import TextBlob

# Настройки Google Sheets
scope = ['https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name(
    'zhenya-pidor-5f22b73bcae0.json', scope
)

# Подключение к OpenAI
openai.api_key = "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"

# Подключение к VK API
vk_session = vk_api.VkApi(
    token="5658b273c2d41e7942219e7990e9e991df9a5874b4fbae86751549e25d44364d988abb11e4f4375bc0ea7"
)
longpoll = VkBotLongPoll(vk_session, 199826105)

print('start')

def sender(id, text):
    vk_session.method('messages.send', {
        'user_id': id,
        'message': text,
        'random_id': 0
    })

def chat_sender(id, text):
    vk_session.method('messages.send', {
        'chat_id': id,
        'message': text,
        'random_id': 0
    })

def extract_info_from_message(msg):
    # Попробуем извлечь Никнейм, Дату и Количество дней
    name_match = re.search(r"([A-Za-z0-9_]+)\s+(\d{2}\.\d{2})\s+(\d+)", msg)
    if name_match:
        name = name_match.group(1)
        date = name_match.group(2)
        days = int(name_match.group(3))
        
        # Преобразуем дату в объект datetime
        try:
            date_obj = datetime.strptime(date, "%d.%m")
            return name, date_obj, days
        except Exception as e:
            print(f"Ошибка при обработке даты: {e}")
            return None, None, None
    else:
        return None, None, None
today = datetime.now()
# Функция для добавления данных в таблицу Google Sheets
def parse_message_and_respond(msg):
    # Извлекаем информацию из сообщения
    name, date, days = extract_info_from_message(msg)

    if name and date and days:
        # Если все данные корректны
        print(f"Распознанные данные: Ник={name}, Дата={date.strftime('%d.%m')}, Количество дней={days}")
        formatted_entry = [name, f"{today.strftime("%d.%m")} - {date.strftime('%d.%m')}", days]  # Добавляем 100 по умолчанию, если нет

        # Подключение к Google Sheets
        doci = "1tvNArJyC1oU-byYv4s3VzxHRy3F5miLomQjX_G7sjuw"
        client = gspread.authorize(credentials)
        spreadshe = client.open_by_key(doci)
        worksheetName = "Лог неактивов"
        worksheet = spreadshe.worksheet(worksheetName)

        worksheet.append_row(formatted_entry)
        print(f"Добавлено в таблицу: {formatted_entry}")
        chat_sender(id, f"Добавлено в таблицу: {formatted_entry}")
        return True
    else:
        return False

# Проверка и форматирование строк в таблице
def check_and_format_rows():
    doci = "1tvNArJyC1oU-byYv4s3VzxHRy3F5miLomQjX_G7sjuw"
    client = gspread.authorize(credentials)
    spreadshe = client.open_by_key(doci)
    worksheetName = "Лог неактивов"
    worksheet = spreadshe.worksheet(worksheetName)
    current_date = datetime.now().strftime("%d.%m")

    # Получаем все данные из таблицы
    all_data = worksheet.get_all_values()

    # Обрабатываем каждую строку
    for i, row in enumerate(all_data):
        if len(row) > 1:  # Проверяем, что есть значение во втором столбце
            date_range = row[1]  # Диапазон дат
            match = re.search(r"\d{2}\.\d{2} - (\d{2}\.\d{2})", date_range)

            if match:
                end_date = match.group(1)  # Конечная дата из диапазона

                # Сравниваем текущую дату с конечной датой
                if end_date == current_date:
                    color = {
                        "red": 255,
                        "green": 0,
                        "blue": 0
                    }  # Красный цвет (RGB)
                elif end_date > current_date:
                    color = {
                        "red": 0,
                        "green": 255,
                        "blue": 0
                    }  # Зеленый цвет (RGB)
                else:
                    continue  # Пропускаем, если дата не подходит

                # Форматируем строку
                row_index = i + 1  # Индекс строки в Google Sheets начинается с 1
                worksheet.format(f"A{row_index}:C{row_index}", {
                    "backgroundColor": {
                        "red": color["red"] / 255,
                        "green": color["green"] / 255,
                        "blue": color["blue"] / 255
                    }
                })

while True:
    try:
        for event in longpoll.listen():
            if event.type == VkBotEventType.MESSAGE_NEW:
                if event.from_chat:
                    msg = event.object.text
                    id = event.chat_id

                    if id == 23:
                        success = parse_message_and_respond(msg)
                        check_and_format_rows()

                        if success:
                            chat_sender(id, f"//////")
                        else:
                            # Проверяем сообщение на формат "Pusii_Reborn 08.03 100"
                            if parse_message_and_respond(msg):
                                chat_sender(id, "Данные успешно обработаны!")
                            else:
                                chat_sender(id, "Неверный формат сообщения. Используй: Никнейм ДД.ММ Количество дней или отправь заявление.")
    except Exception as err:
        chat_sender(event.chat_id, f"Ошибка: {err}")

