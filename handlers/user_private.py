import re
from aiogram import types, Router, F
from aiogram.filters import CommandStart
from aiogram.fsm.context import FSMContext
from openpyxl.reader.excel import load_workbook

from filters.chat_types import ChatTypeFilter
from edit_table import current_file, add_note_xlsx
from kbds.reply import get_keyboard
from states import EditTable

user_private_router = Router()
user_private_router.message.filter(ChatTypeFilter(chat_types=['private']))


@user_private_router.message(CommandStart())
async def start_cmd(message: types.Message):
    await message.answer('Started',
                         reply_markup=get_keyboard(
                             "Добавить запись",
                             "О Боте",
                             placeholder="Выберите соответствующий раздел",
                             sizes=(2,)
                         ),
                         )
    await message.answer(str(message.chat.id))


@user_private_router.message(EditTable.add_note, F.text == "Добавить запись")
async def note_add(message: types.Message, state: FSMContext):
    # Проверяем, что файл создан администратором
    if not current_file:
        await message.answer("Администратор должен сначала создать файл.")
        return

    # Сохраняем ID чата в состоянии для добавления записи
    await state.update_data(note=message.chat.id)

    # Добавляем запись в файл
    try:
        await add_note_xlsx(message)
        await message.answer("Запись добавлена.")
    except Exception as e:
        await message.answer(f"Произошла ошибка при добавлении записи: {str(e)}")
