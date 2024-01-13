from telebot import types, TeleBot
from random import randint

import threading
import pathlib
import os

import cd_parser
import xls

TOKEN = "6598793377:AAHLUTUCdnNZd-vxrxjTT5bIM2PXyGaaRwU"
bot = TeleBot(TOKEN, parse_mode=None)


@bot.message_handler(commands=['start'])
def send_welcome(message: types.Message):
    bot.send_message(message.chat.id, "Введите кадастровый номер или вставьте excel файл Р1Р7", reply_markup=create_buttons())        
 

@bot.message_handler(content_types=['document'])
def send_obj_info_by_doc(doc: types.Message):
    file_name = doc.document.file_name
    file_type = pathlib.Path(file_name).suffix
    file_info = bot.get_file(doc.document.file_id)
    downloaded_file = bot.download_file(file_info.file_path)
    
    
    if file_type == ".xls" or file_type == ".xlsx":
        file_dir = os.path.join(os.getcwd(), f"{randint(1, 99999)}{file_type}")
        with open(file_dir, "wb") as file:
            file.write(downloaded_file)
        file.close()
        xls.read_xls(file_dir, bot, doc)
    else:
        bot.send_message(doc.chat.id, 'Не поддерживаемый тип файла!\nПожалуйста загрузите файлы с расширением: <b>.xls</b> или <b>.xlsx</b>', parse_mode='html')



@bot.message_handler(content_types=['text'])
def btns_handler(message: types.Message):
    if(message.text == 'Парсинг карточек ЕГРН'):
        bot.send_message(message.chat.id, "Отправте excel файл Р1Р7")
    
    elif(message.text == 'Ввод кад. номера'):
        bot.send_message(message.chat.id, "Введите кадастровый номер")
    else:
        cad_num = message.text
        threading.Thread(target=cd_parser.parse_txt, args=(cad_num, bot, message), daemon=True).start()


def create_buttons():
    markup = types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
    excel = types.KeyboardButton('Парсинг карточек ЕГРН')
    text = types.KeyboardButton('Ввод кад. номера')
    markup.add(excel, text)
    return markup



bot.infinity_polling()


