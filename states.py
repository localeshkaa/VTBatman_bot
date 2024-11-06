from aiogram.fsm.state import StatesGroup, State


class EditTable(StatesGroup):
    file_created = State()
    waiting_for_data = State()
    awaiting_back = State()
    show_data = State()
    creating_file = State()
    add_name = State()
    add_group_id = State()
    delete_name = State()
    add_note = State()
    day_plan = State()
