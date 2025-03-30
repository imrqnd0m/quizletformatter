import asyncio
import xlsxwriter
from aiogram import Bot, Dispatcher, types, F
from aiogram.types import FSInputFile
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from datetime import datetime

TOKEN = "8109396622:AAHh3Gnm__jDkfv_p1fvQh0yvnZARw_OSNM"
bot = Bot(token=TOKEN)
storage = MemoryStorage()
dp = Dispatcher(storage=storage)


class Form(StatesGroup):
    waiting_for_text = State()
    waiting_for_separator = State()
    waiting_for_numbering = State()


@dp.message(Command("start"))
async def start(message: types.Message):
    await message.answer("Привет! Отправь команду /form чтобы я создал таблицу из списка слов")


@dp.message(Command("form"))
async def request_text(message: types.Message, state: FSMContext):
    await message.answer("Пришли список слов")
    await state.set_state(Form.waiting_for_text)


@dp.message(Form.waiting_for_text, F.text)
async def request_separator(message: types.Message, state: FSMContext):
    await state.update_data(text_data=message.text.strip().split("\n"))
    await message.answer(
        "Какой разделитель используется между словом и переводом? Отправь только символ-разделитель")
    await state.set_state(Form.waiting_for_separator)


@dp.message(Form.waiting_for_separator, F.text)
async def request_numbering(message: types.Message, state: FSMContext):
    await state.update_data(separator=message.text.strip())
    await message.answer("Присутствует ли нумерация перед словами? (Да/Нет)")
    await state.set_state(Form.waiting_for_numbering)


@dp.message(Form.waiting_for_numbering, F.text)
async def process_text(message: types.Message, state: FSMContext):
    user_data = await state.get_data()
    text_data = user_data["text_data"]
    separator = " " + user_data["separator"] + " "
    remove_numbering = message.text.strip().lower() == "да"

    filename = datetime.today().strftime('%Y-%m-%d') + ".xlsx"
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()

    ind = 0
    for line in text_data:
        ind += 1
        if remove_numbering:
            line = line.split(". ", 1)[-1]
        parts = line.split(separator)
        if len(parts) == 2:
            worksheet.write(f'A{ind}', parts[0].strip())
            worksheet.write(f'B{ind}', parts[1].strip())

    workbook.close()

    doc = FSInputFile(filename)
    await message.answer_document(doc)
    await state.clear()


async def main():
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())