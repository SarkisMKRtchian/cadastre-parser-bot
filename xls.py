from openpyxl import load_workbook, styles
from pprint import pprint as pp
import cd_parser
import telebot
from telebot import types
import pathlib
import pandas
import os
import re
import log
def read_xls(doc: str, bot: telebot.TeleBot, mess: types.Message):
    try:
        # Читает 1 и 7 раздел и конвертирует их в словарь
        sheet_1 = pandas.read_excel(doc, sheet_name='Раздел 1').to_dict()
        sheet_2_dict = pandas.read_excel(doc, sheet_name='Раздел 7').to_dict()
        
        # Собираем новый массив удаляя ключи словаря, потому что они могу отличатся
        sheet_2 = []
        for key in sheet_2_dict:
            sheet_2.append(sheet_2_dict[key])
        
        # Информация по главному кн
        cad_nums = [{
                "col": 3, 
                "row": 2, 
                "sheet": "Раздел 1", 
                "cad_num": sheet_1['Unnamed: 1'][0], 
                "mess": sheet_1['Unnamed: 2'][0] if (len(sheet_1) >= 3) else '', 
                "adress": sheet_1['Unnamed: 1'][4]
            }]
        
        # Информация по зем. участку
        if(sheet_1['Unnamed: 1'][13] != "данные отсутствуют"): 
            cad_nums.append({
                    "col": 3, 
                    "row": 15, 
                    "sheet": "Раздел 1", 
                    "cad_num": sheet_1['Unnamed: 1'][13], 
                    "mess": sheet_1['Unnamed: 2'][13] if (len(sheet_1) >= 3) else ''
                })
        
        # Проходимт по все кн из раздела 7 и добавляет в массив. Так же проверят если уже заполнены данные их тоже пишет. col и row это их места в xls файле
        for num in sheet_2[1]:        
            cad_nums.append({
                "col": 8,
                "row": num + 2,
                "sheet": "Раздел 7",
                "cad_num": sheet_2[1][num],
                "mess": str(sheet_2[8][num]) if (len(sheet_2) == 9) else ''
            })
        
        # Парсит росреестр (см. cd_parser.py)
        cd_parser.parse_excel(cad_nums, bot, mess, doc)
        
        data = cd_parser.parser_excel(cad_nums, bot, message_id=mess.message_id, chat_id=mess.chat.id, filename=doc)
        if (data != False): write_excel(doc, data, bot, mess.chat.id, mess.message_id)
    except KeyError as err:
        log.write(f"Key error: {err} | {__file__}")
        bot.send_message(mess.chat.id, "Ошибка при обротке файла. Провертье корректность файла")
        os.remove(doc)


def write_excel(file: str, data: dict, bot: telebot.TeleBot, chat_id: int):
    # Пишет xls файл
    try:
        # открывает файл
        wb = load_workbook(file, data_only=True)
        
        # Проходит по всем данным и записывает их в свои колонки
        for item in data['data']:
            ws = wb[item["sheet"]]
            if(item['sheet'] == 'Раздел 1'):
                ws.cell(item['row'], item['col']).value = item['mess']
                # Прекрипляет к верху колонки
                ws.cell(item['row'], item['col']).alignment = styles.Alignment(horizontal="left", vertical="top")
            else:
                ws.cell(item['row'], item['col']).value = data['data'][0]['cad_num']
                ws.cell(item['row'], item['col'] + 1).value = item['mess']
                ws.cell(item['row'], item['col']).alignment = styles.Alignment(horizontal="left", vertical="top")
                ws.cell(item['row'], item['col'] + 1).alignment = styles.Alignment(horizontal="left", vertical="top")
        
        # Создает новое имя файла
        adress: str = re.sub(r"[,|.| ]", '-', data['data'][0]['adress'])
        cad_num: str = re.sub(r"[:]", "-", data['data'][0]['cad_num'])
        f_name = f"Р1Р7-{cad_num}-{adress}{pathlib.Path(file).suffix}"
        # Сохраняет файл и закрывает
        wb.save(f_name)
        wb.close()
        
        # Открывает файл
        with open(f_name, "rb") as f:
            errs = 'При обработке следующих кад номеров произошла ошибка:\n'
            # Еасли во время парсинга какие-то кн не удалось обрабротать, то записываем их в сообщение (см. cd_parser.py)
            if(len(data['processed_failure']) > 0):
                for i in data['processed_failure']:
                    errs = f"{data['processed_failure'][i]}\n"
            # Отправляет файл
            bot.send_document(chat_id, f, caption=f"{adress}\nC {data['start']} по {data['end']} = {data['time_for_one_card']} сек/КН\n{data['processed']}\n${errs}")
            f.close()
        # Удаляет файл
        os.remove(f_name)
    
    except OSError as err:
        bot.send_message(chat_id, "Во время обробтки файла произошла ошибка, повторитье попытку")
        log.write(f"{err} | {__file__}")
        os.remove(f_name)
        
    except TypeError as err:
        bot.send_message(chat_id, "Не удалось сформировать exel документ. Повторитье попытку")
        log.write(f"{err} | {__file__}")
        os.remove(f_name)
    
    except Exception as err:
        log.write(f"{err} | {__file__}")
        bot.send_message(chat_id, "Не удалось сформировать exel документ. Повторитье попытку")
        os.remove(f_name)
    

