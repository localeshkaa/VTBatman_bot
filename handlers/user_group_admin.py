from aiogram import types, Bot, F, Router
from aiogram.filters import Command

from filters.chat_types import ChatTypeFilter

user_group_admin = Router()
user_group_admin.message.filter(ChatTypeFilter(['group', 'supergroup']))
user_group_admin.edited_message.filter(ChatTypeFilter(['group', 'supergroup']))


@user_group_admin.message(Command("admin"))
async def get_admin(message: types.Message, bot: Bot):
    chat_id = message.chat.id
    admins_list = await bot.get_chat_administrators(chat_id)

    admins_list = [
        member.user.id
        for member in admins_list
        if member.status == 'administrator' or member.status == 'creator'
    ]
    bot.my_admins_list = admins_list
    if message.from_user.id in admins_list:
        await message.delete()
