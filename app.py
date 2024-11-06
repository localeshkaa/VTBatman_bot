import asyncio
import os

from aiogram import Bot, Dispatcher
from aiogram.types import BotCommandScope, BotCommandScopeAllPrivateChats

from dotenv import find_dotenv, load_dotenv

load_dotenv(find_dotenv())

from handlers.user_group_admin import user_group_admin
from handlers.user_private import user_private_router
from handlers.admin_private import admin_router
from common.bot_cmds_list import private

ALLOWED_UPDATES = ['message, edited_message']

bot = Bot(token=os.getenv('TOKEN'))
bot.my_admins_list = []

dp = Dispatcher()

dp.include_routers(user_private_router, admin_router, user_group_admin)


async def main():
    await bot.delete_webhook(drop_pending_updates=True)
    await bot.set_my_commands(commands=private, scope=BotCommandScopeAllPrivateChats())
    await dp.start_polling(bot, allowed_updates=ALLOWED_UPDATES)


asyncio.run(main())


#
#
# # Обработчик команды /start
# @dp.message_handler(commands=['start'])
# async def process_start_command(message: Message):
#     await message.reply("Добро пожаловать!")
#
#
# # Обработчик команды /help
# @dp.message_handler(commands=['help'])
# async def process_help_command(message: Message):
#     await message.reply("Этот бот позволяет создавать отчетность для ВТБ")
#
#
# # Обработчик кнопки "Добавить запись" и "Внести изменения"
# @dp.message_handler(lambda message: message.text == "Добавить запись" or message.text == "Внести изменения")
# async def handle_action(message: Message):
#     if message.text == "Внести изменения" and message.chat.id in allowed_chat_id:
#         await edit_data(message)
#     elif message.text == "Добавить запись":
#         await message.reply("Функция добавления записи еще не реализована.")  # Заглушка для добавления записи
#
#
# # Функция обработки изменения данных
# async def edit_data(message: Message):
#     await message.reply("Какие изменения необходимо внести?")
#
#     # Создаем клавиатуру для редактирования данных
#     markup_edit_data = ReplyKeyboardMarkup(resize_keyboard=True)
#
#     # Добавляем кнопки
#     create_file_button = KeyboardButton("Создать файл")
#     show_data_button = KeyboardButton("Отобразить таблицу")
#     add_person_button = KeyboardButton("Добавить человека")
#     remove_person_button = KeyboardButton("Удалить человека")
#     create_group_button = KeyboardButton("Создать группу")
#     back_button = KeyboardButton("Назад")
#
#     markup_edit_data.add(create_file_button, show_data_button, add_person_button,
#                          remove_person_button, create_group_button, back_button)
#
#     user_states[message.chat.id] = ['edit_data']  # Устанавливаем состояние пользователя
#
#     await message.reply("Выберите действие:", reply_markup=markup_edit_data)
#
#
# @dp.message_handler(func=lambda message: message.text == "Добавить человека" and message.chat.id in allowed_chat_id)
# def add_person(message):
#     bot.send_message(message.chat.id, "Введите фамилию и имя человека")
#
#     user_states[message.chat.id] = ['add_person']  # Устанавливаем состояние пользователя
#
#     # Ожидаем ответ с именем
#     bot.register_next_step_handler(message, get_name)
#
#
# @dp.message_handler(func=lambda message: message.text == "Назад")
# def back_button(message):
#     if user_states.get(message.chat.id) == ['edit_data']:
#         start(message)  # Возвращаемся к главному меню
#     else:
#         bot.send_message(chat_id=message.chat.id, text='Вы находитесь в главном меню.',
#                          reply_markup=start_markup)
#
#
# # Функция для обработки имени (пример)
# def get_names(message):
#     name = message.text
#     bot.send_message(message.chat.id, f"Вы добавили человека: {name}")
#
#
# @dp.message_handler(commands=['stop'])
# # Закрытие файла для записи
# def stop(message):
#     bot.send_message(message.chat.id, text="Файл для записи закрыт!")
#
#
# @dp.message_handler(commands=['help'])
# def help(message):
#     bot.send_message(message.chat.id, text="Этот бот умеет записывать данные в .xlsx формат")
#
#
# @dp.message_handler(func=lambda message: message.text == "Создать файл" and message.chat.id in allowed_chat_id)
# def create_file_buttons(message):
#     # Получаем сегодняшнюю дату
#     today_date = datetime.date.today()
#     # Форматируем имя файла
#     file_name = today_date.strftime("%Y%m%d") + ".xlsx"  # Добавляем расширение .xlsx
#
#     # Загружаем рабочую книгу из шаблона
#     workbook = load_workbook('shablon.xlsx')
#
#     # Создаем 6 листов (если это необходимо)
#     # for i in range(1, 7):
#     #     if i == 1:
#     #         sheet = workbook.active
#     #         sheet.title = f"Лист {i}"
#     #     else:
#     #         workbook.create_sheet(title=f"Лист {i}")
#
#     # Сохраняем рабочую книгу с новым именем
#     workbook.save(file_name)
#
#     # Отправляем сообщение пользователю о том, что файл создан
#     bot.send_message(message.chat.id, f"Файл '{file_name}' успешно создан!")


#
#
# @dp.message_handler(func=lambda message: message.text == "Отобразить таблицу" and message.chat.id in allowed_chat_id)
# def show_data(message):  # Отображение таблицы
#     # Загружаем файл с сегодняшней датой
#     file_name = datetime.date.today().strftime("%Y%m%d") + ".xlsx"
#
#     try:
#         file = load_workbook(file_name)
#     except FileNotFoundError:
#         bot.send_message(message.chat.id, text="Файл не найден.")
#         return
#
#     bot.send_message(message.chat.id, text="Введите номер группы")
#
#     @dp.message_handler(func=lambda msg: msg.chat.id == message.chat.id)
#     def handle_user_input(user_message):
#         user_input = user_message.text.strip()  # Получаем ввод пользователя
#
#         if user_input in file.sheetnames:
#             sheet = file[user_input]  # Получаем нужный лист
#
#             # Собираем данные из листа
#             data = []
#             for row in sheet.iter_rows(values_only=True):
#                 data.append(row)  # Добавляем строки в список
#
#             # Отправляем данные пользователю
#             for row in data:
#                 bot.send_message(message.chat.id, text=', '.join(map(str, row)))  # Отправляем каждую строку
#
#             bot.send_message(message.chat.id, text="Данные успешно отправлены!")
#         else:
#             bot.send_message(message.chat.id, text="Лист с таким номером группы не найден.")
#     # with open(datetime.date.today().strftime("%Y%m%d") + ".xlsx", "rb") as file:
#
#
# @dp.message_handler(func=lambda message: message.text == "Добавить запись")
# def handle_text(message):
#     bot.send_message(message.chat.id, text="""Сообщение должно выглядеть в формате:
#         ФИО:
#         /* остальные данные */""")
#     user_input = message.text
#
#     # Регулярные выражения для извлечения данных
#     name = re.search(r'ФИО (.+)', user_input).group(1) if re.search(r'ФИО (.+)', user_input) else None
#     report_date = re.search(r'Отчёт (.+)', user_input).group(1) if re.search(r'Отчёт (.+)', user_input) else None
#     dk_given = re.search(r'Дк выдано: (.+)', user_input).group(1) if re.search(r'Дк выдано: (.+)', user_input) else None
#     kk_given = re.search(r'Кк выдано: (.+)', user_input).group(1) if re.search(r'Кк выдано: (.+)', user_input) else None
#     urgent_request = re.search(r'Срочная заявка: (.+)', user_input).group(1) if re.search(r'Срочная заявка: (.+)',
#                                                                                           user_input) else None
#     referral_links = re.search(r'Реферальнын ссылки: (.+)', user_input).group(1) if re.search(
#         r'Реферальнын ссылки: (.+)', user_input) else None
#     additional_cards_14 = re.search(r'Доп\. Карты от 14 лет: (.+)', user_input).group(1) if re.search(
#         r'Доп\. Карты от 14 лет: (.+)', user_input) else None
#     children_cards = re.search(r'Детские карты: (.+)', user_input).group(1) if re.search(r'Детские карты: (.+)',
#                                                                                          user_input) else None
#     request = re.search(r'Заявок - (.+)', user_input).group(1) if re.search(r'Заявок - (.+)', user_input) else None
#     sold_kk = re.search(r'Продано - (.+)', user_input).group(1) if re.search(r'Продано - (.+)', user_input) else None
#     errors = re.search(r'Ошибка - (.+)', user_input).group(1) if re.search(r'Ошибка - (.+)', user_input) else None
#     apib = re.search(r'Апиб - (.+)', user_input).group(1) if re.search(r'Апиб - (.+)', user_input) else None
#     apik = re.search(r'Апик - (.+)', user_input).group(1) if re.search(r'Апик - (.+)', user_input) else None
#     skk = re.search(r'Скк - (.+)', user_input).group(1) if re.search(r'Скк - (.+)', user_input) else None
#     in_moment = re.search(r'В моменте - (.+)', user_input).group(1) if re.search(r'В моменте - (.+)',
#                                                                                  user_input) else None
#     all_connect = re.search(r'Всего подключено - (.+)', user_input).group(1) if re.search(r'Всего подключено - (.+)',
#                                                                                           user_input) else None
#     sold_com = re.search(r'Продано - (.+)', user_input).group(1) if re.search(r'Продано - (.+)', user_input) else None
#     sold_ksp = re.search(r'Продано - (.+)', user_input).group(1) if re.search(r'Продано - (.+)', user_input) else None
#     ss_ap = re.search(r'СС - (.+)', user_input).group(1) if re.search(r'СС - (.+)', user_input) else None
#     zhku_ap = re.search(r'ЖКУ - (.+)', user_input).group(1) if re.search(r'ЖКУ - (.+)', user_input) else None
#     third_pers_card = re.search(r'Карты на 3-е лицо: (.+)', user_input).group(1) if re.search(
#         r'Карты на 3-е лицо: - (.+)', user_input) else None
#     pens = re.search(r'Пенс - (.+)', user_input).group(1) if re.search(r'Пенс - (.+)', user_input) else None
#     ns = re.search(r'НС: (\d+)', user_input).group(1) if re.search(r'НС: (\d+)', user_input) else None
#     pos = re.search(r'POS: (\d+)', user_input).group(1) if re.search(r'POS: (\d+)', user_input) else None
#     additional_card = re.search(r'Доп карта: (\d+)', user_input).group(1) if re.search(r'Доп карта: (\d+)',
#                                                                                        user_input) else None
#     sticker = re.search(r'Стикер: (\d+)', user_input).group(1) if re.search(r'Стикер: (\d+)', user_input) else None
#
#     # Создаем DataFrame из данных в виде столбцов
#     data = {
#         "Параметр": [
#             "ФИО"
#             "Отчёт",
#             "Дк выдано",
#             "Кк выдано",
#             "Срочная заявка",
#             "Реферальных ссылки",
#             "Доп. Карты от 14 лет",
#             "Детские карты",
#             "Заявок",
#             "Продано",
#             "Ошибка",
#             "Апиб",
#             "Апик",
#             "Скк",
#             "В моменте",
#             "Всего подключено",
#             "Продано",
#             "Продано",
#             "СС",
#             "ЖКУ",
#             "Карты на 3-е лицо",
#             "Пенс",
#             "НС",
#             "POS",
#             "Доп карта",
#             "Стикер"
#         ],
#         "Значение": [
#             name,
#             report_date,
#             dk_given,
#             kk_given,
#             urgent_request,
#             referral_links,
#             additional_cards_14,
#             children_cards,
#             request,
#             sold_kk,
#             errors,
#             apib,
#             apik,
#             skk,
#             in_moment,
#             all_connect,
#             sold_com,
#             sold_ksp,
#             ss_ap,
#             zhku_ap,
#             third_pers_card,
#             pens,
#             ns,
#             pos,
#             additional_card,
#             sticker
#         ]
#     }
#
#     df = pd.DataFrame(data)
#
#     # Сохраняем DataFrame в Excel в памяти с расширением .xlsx
#     output = BytesIO()
#     with pd.ExcelWriter(output, engine='openpyxl') as writer:
#         df.to_excel(writer, index=False, sheet_name='Отчет')
#
#     output.seek(0)
#
#     # Сохраняем временный файл на диске
#     temp_filename = 'report.xlsx'
#     with open(temp_filename, 'wb') as f:
#         f.write(output.getvalue())
#
#     # Отправляем Excel-файл пользователю с правильным именем файла
#     with open(temp_filename, 'rb') as f:
#         bot.send_document(message.chat.id, f, caption='Ваши данные в Excel')
#
#     # Удаляем временный файл после отправки
#     os.remove(temp_filename)
