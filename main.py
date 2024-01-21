from telebot import types, TeleBot
from dotenv import load_dotenv
from random import randint

import pathlib
import os
import json
import re

from xls import Excel

from cd_parser import ParseTxt, ParseExcel


load_dotenv()
TOKEN = os.getenv("BOT_TOKEN")

bot = TeleBot(TOKEN, parse_mode=None)

parseTxt = ParseTxt(bot)
parseExcel = ParseExcel(bot)
xls = Excel(bot, parseExcel)

@bot.message_handler(commands=['start'])
def send_welcome(message: types.Message):
    bot.send_message(message.chat.id, "Введите кадастровый номер или вставьте excel файл Р1Р7", reply_markup=create_buttons())        
 
 
@bot.message_handler(content_types=['document'])
def send_obj_info_by_doc(message: types.Message):
    if(parseExcel.isWork):
        bot.send_message(message.chat.id, "Идет обработка файла/карточки! Пожалуйста повторите попытку позже.") 
        return
    
    jsonPath = os.path.join(os.getcwd(), 'cadNums.JSON')
    if(os.path.exists(jsonPath)):
        os.remove(jsonPath)
    
    fileName = message.document.file_name
    fileType = pathlib.Path(fileName).suffix
    fileInfo = bot.get_file(message.document.file_id)
    downloaded_file = bot.download_file(fileInfo.file_path)
    
    
    if fileType == ".xls" or fileType == ".xlsx":
        filePath = os.path.join(os.getcwd(), f"{randint(1, 99999)}{fileType}")
        with open(filePath, "wb") as file:
            file.write(downloaded_file)
        file.close()
        xls.read(filePath, message)
    else:
        bot.send_message(message.chat.id, 'Не поддерживаемый тип файла!\nПожалуйста загрузите файлы с расширением: <b>.xls</b> или <b>.xlsx</b>', parse_mode='html')


@bot.message_handler(content_types=['text'])
def btns_handler(message: types.Message):
    if(parseTxt.isWork or parseExcel.isWork):
        bot.send_message(message.chat.id, "Идет обработка файла/карточки! Пожалуйста повторите попытку позже.") 
        
    elif(message.text == 'Парсинг в файл Р1Р7'):
        bot.send_message(message.chat.id, "Отправте excel файл Р1Р7")
    
    elif(message.text == 'Карточка ЕГРН по КН'):
        bot.send_message(message.chat.id, "Введите кадастровый номер")
    
    else:
        cadNum = message.text
        cadNumREGEXP = r"\d{2}:\d{2}"
        if(re.match(cadNumREGEXP, cadNum)):
            parseTxt.parse(cadNum, message)
        else:
            bot.send_message(message.chat.id, "Не корректный кадастровый номер!")
        


@bot.callback_query_handler(func=lambda call: True)
def stopParse(call: types.CallbackQuery):
    if(call.data == 'stop'):
        parseExcel.stop = True
        
    elif(call.data == 'download'):
        objPath = os.path.join(os.getcwd(), 'cadNums.JSON')
        if(os.path.exists(objPath)):    
            with open(objPath, 'r') as file:    
                obj = json.load(file)
                file.close()
            xls.write(obj['fp'], obj, call.message.chat.id)
            os.remove(objPath)
            bot.delete_message(call.message.chat.id, call.message.message_id)
        else:
            bot.send_message(call.message.chat.id, "Ошибка! Не удалось найти файл")
            
    elif(call.data == 'continue'):
        objPath = os.path.join(os.getcwd(), 'cadNums.JSON')
        if(os.path.exists(objPath)):   
            with open(objPath, 'r') as file:    
                obj = json.load(file)
                file.close()
            parseExcel.parse(obj['data'], obj["fp"], call.message, xls)
            os.remove(objPath)
        else:
            bot.send_message(call.message.chat.id, "Ошибка! Не удалось найти файл")
        
def create_buttons():
    markup = types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
    excel = types.KeyboardButton('Парсинг в файл Р1Р7')
    text = types.KeyboardButton('Карточка ЕГРН по КН')
    markup.add(excel, text)
    return markup

bot.infinity_polling()


