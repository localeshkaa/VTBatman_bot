from aiogram.types import ReplyKeyboardMarkup, KeyboardButton
from aiogram.utils.keyboard import ReplyKeyboardBuilder


def get_keyboard(
    *btns: str,
    placeholder: str = None,
    sizes: tuple[int] = (2,),
    back_button: bool = False
) -> ReplyKeyboardMarkup:
    """
    Универсальная функция для создания клавиатуры с поддержкой
    динамических размеров и опциональной кнопкой «Назад».

    :param btns: Тексты кнопок
    :param placeholder: Текст-заполнитель для поля ввода
    :param sizes: Кортеж с размером кнопок в строке
    :param back_button: Добавляет кнопку "Назад" в конец клавиатуры
    :return: Сформированная клавиатура
    """
    keyboard = ReplyKeyboardBuilder()

    # Добавляем основные кнопки
    for text in btns:
        keyboard.add(KeyboardButton(text=text))

    # Добавляем кнопку "Назад", если указано
    if back_button:
        keyboard.add(KeyboardButton(text="Назад"))

    return keyboard.adjust(*sizes).as_markup(
        resize_keyboard=True, input_field_placeholder=placeholder
    )



# start_kb = ReplyKeyboardMarkup(
#     keyboard=[
#         [
#             KeyboardButton(text='Добавить запись'),
#             KeyboardButton(text='Внести изменения')
#         ],
#     ],
#     resize_keyboard=True,
#     input_field_placeholder='Выберите функцию'
# )
#
# edit_kb = ReplyKeyboardMarkup(
#     keyboard=[
#         [
#             KeyboardButton(text='Создать файл'),
#             KeyboardButton(text='Отобразить таблицу'),
#         ],
#         [
#             KeyboardButton(text='Добавить человека'),
#             KeyboardButton(text='Удалить человека'),
#         ],
#         [
#             KeyboardButton(text='Создать группу')
#         ],
#         KeyboardButton(text='Назад')
#
#     ],
#     resize_keyboard=True,
#     input_field_placeholder='Выберите функцию'
# )
# delete_kb = ReplyKeyboardRemove()
