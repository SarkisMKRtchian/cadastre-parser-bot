import pandas
from openpyxl import load_workbook, styles
from pprint import pprint as pp
import cd_parser
import telebot
from telebot import types
import pathlib
import anticaptcha
import os
import re
import log
def read_xls(doc: str, bot: telebot.TeleBot, mess: types.Message):
    try:
        sheet_1 = pandas.read_excel(doc, sheet_name='Раздел 1').to_dict()
        sheet_2_dict = pandas.read_excel(doc, sheet_name='Раздел 7').to_dict()
        
        sheet_2 = []
        for key in sheet_2_dict:
            sheet_2.append(sheet_2_dict[key])
        
        cad_nums = [{
                "col": 3, 
                "row": 2, 
                "sheet": "Раздел 1", 
                "cad_num": sheet_1['Unnamed: 1'][0], 
                "mess": sheet_1['Unnamed: 2'][0] if (len(sheet_1) >= 3) else '', 
                "adress": sheet_1['Unnamed: 1'][4]
                }]
        
        if(sheet_1['Unnamed: 1'][13] != "данные отсутствуют"): 
            cad_nums.append({
                    "col": 3, 
                    "row": 15, 
                    "sheet": "Раздел 1", 
                    "cad_num": sheet_1['Unnamed: 1'][13], 
                    "mess": sheet_1['Unnamed: 2'][13] if (len(sheet_1) >= 3) else ''
                })
        
        for num in sheet_2[1]:        
            cad_nums.append({
                "col": 8,
                "row": num + 2,
                "sheet": "Раздел 7",
                "cad_num": sheet_2[1][num],
                "mess": str(sheet_2[8][num]) if (len(sheet_2) == 9) else ''
            })
        
        
        
        data = cd_parser.parser_excel(cad_nums, bot, message_id=mess.message_id, chat_id=mess.chat.id, filename=doc)
        if (data != False): write_excel(doc, data, bot, mess.chat.id, mess.message_id)
    except KeyError as err:
        log.write(f"Key error: {err} | {__file__}")
        bot.send_message(mess.chat.id, "Ошибка при обротке файла. Провертье файл на корректность")
        os.remove(doc)


def write_excel(file, data, bot: telebot.TeleBot, chat_id, mess_id):
    try:
        wb = load_workbook(file, data_only=True)
        
        for item in data['data']:
            ws = wb[item["sheet"]]
            if(item['sheet'] == 'Раздел 1'):
                ws.cell(item['row'], item['col']).value = item['mess']
                ws.cell(item['row'], item['col']).alignment = styles.Alignment(horizontal="left", vertical="top")
            else:
                ws.cell(item['row'], item['col']).value = data['data'][0]['cad_num']
                ws.cell(item['row'], item['col'] + 1).value = item['mess']
                ws.cell(item['row'], item['col']).alignment = styles.Alignment(horizontal="left", vertical="top")
                ws.cell(item['row'], item['col'] + 1).alignment = styles.Alignment(horizontal="left", vertical="top")
            
        wb.save(file)
        wb.close()
        adress: str = re.sub(r"[,|.| ]", '-', data['data'][0]['adress'])
        cad_num: str = re.sub(r"[:]", "-", data['data'][0]['cad_num'])
        
        excel_file = os.path.join(os.getcwd(), f"Р1Р7-{cad_num}-{adress}{pathlib.Path(file).suffix}")
        
        os.rename(file, excel_file)
        
        with open(excel_file, "rb") as f:
            errs = ''
            if(data['errs'] != ''):
                errs = f"\nПри обработке следуйющих кад номеров произошла ошибка: \n{data['errs']}"
        
            bot.send_document(chat_id, f, caption=f"Начало обработки: {data['start']}\nКонец обработки: {data['end']}\nБаланс антикапчи: {anticaptcha.get_balance()} ${errs}")
            f.close()
        os.remove(excel_file)
        bot.delete_message(chat_id, mess_id - 2)
        bot.delete_message(chat_id, mess_id - 1)
        bot.delete_message(chat_id, mess_id)
        bot.delete_message(chat_id, mess_id + 1)
        bot.delete_message(chat_id, mess_id + 2)
    
    except OSError as err:
        bot.send_message(chat_id, "Во время обробтки файла произошла ошибка, повторитье попытку")
        
    except telebot.apihelper.ApiTelegramException as err:
        bot.send_message(chat_id, "Ошибка при работе бота. Повторите попытку")
        log.write(f"{err.description} | {__file__}")
        os.remove(excel_file)
        
    except TypeError as err:
        bot.send_message(chat_id, "Не удалось сформировать exel документ. Повторитье попытку")
        log.write(f"{err} | {__file__}")
        os.remove(file)
    
    

