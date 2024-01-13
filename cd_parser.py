import selenium.common.exceptions as sl_exps 
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from telebot import types
import telebot
from datetime import datetime
import anticaptcha
import os
import time
import xls
from random import randint
import log

def parser_excel(cad_num: str, bot: telebot.TeleBot, chat_id: int, message_id: int, filename: str):
    # dv_dir = os.path.join(os.getcwd(), "driver", "win", "geckodriver.exe") -> win
    # bw_dir = r"C:\Program Files\Mozilla Firefox\firefox.exe" -> win
    dv_dir = os.path.join(os.getcwd(), "driver", "linux", "geckodriver")
    bw_dir = r"usr/bin/firefox"
    cp_dir = os.path.join(os.getcwd())
    
    service = Service(executable_path=dv_dir, port=randint(1000, 10000))
    
    options = webdriver.FirefoxOptions()
    options.binary_location = bw_dir
    options.add_argument('--headless')
    
    driver = webdriver.Firefox(service=service, options=options)
    
    bot.send_message(chat_id, "Начинаю обработку. Ожидайте...")
    
    URL = "https://lk.rosreestr.ru/eservices/real-estate-objects-online"
    driver.maximize_window()
    driver.get(URL)
    date_start = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    mess_id = message_id + 2
    errs = ''
    try:
        i = 0
        bot.send_message(chat_id, f"Обработано кад. номеров: {i} из {len(cad_num)}")
        time.sleep(10)
        for numbers in cad_num:
            if (numbers['mess'] != ''): 
                if(numbers['mess'] != 'nan'):
                    i += 1
                    bot.edit_message_text(f"Обработано кад. номеров: {i} из {len(cad_num)}\nКад. номер: {numbers['cad_num']}", chat_id, mess_id)
                    continue
            
            
            
            with open(cp_dir + "/cp.png", 'wb') as file:
                l = driver.find_element(By.XPATH, '//*[@alt="captcha"]')
                file.write(l.screenshot_as_png)
            
            captcha = anticaptcha.solve_captcha(cp_dir + "/cp.png", bot, chat_id)
            if(captcha == False):
                return False
            
            captch_input = driver.find_element(By.ID, "captcha")
            captch_input.clear()
            captch_input.send_keys(captcha)
            
            
            cad_input = driver.find_element(By.ID, "query")
            cad_input.clear()
            cad_input.send_keys(numbers["cad_num"])
            
            err = driver.find_elements(By.CLASS_NAME, "rros-ui-lib-message--error")
            if(len(err) != 0):
                i += 1
                err += f"{numbers['cad_num']}\n"
                continue
                    
            time.sleep(1)
            sch_btn = driver.find_element(By.ID, "realestateobjects-search")
                                
            sch_btn.click()
            
            time.sleep(5)
            
            card_btn = driver.find_elements(By.CLASS_NAME, "realestateobjects-wrapper__results__cadNumber")
            if(len(card_btn) == 0):
                    numbers['mess'] = ''
                    i += 1
                    err += f"{numbers['cad_num']}\n"
                    continue
                
            card_btn[0].click()
            
            time.sleep(1)
            
            card = driver.find_elements(By.CLASS_NAME, "build-card-wrapper__info")
            
            mess = ""
            for elem in card:
                h3 = elem.find_element(By.TAG_NAME, "h3").text
                ul = elem.find_element(By.TAG_NAME, "ul")
                li = ul.find_elements(By.TAG_NAME, "li")
                mess += h3
                for item in li:
                    name = item.find_element(By.TAG_NAME, "span").text
                    value = item.find_elements(By.CLASS_NAME, "build-card-wrapper__info__ul__subinfo__options__item__line")
                    if(len(value) == 1): mess += name + value[0].text
                    else:
                        mess += name
                        for val in value:
                            mess += val.text
                    
            numbers['mess'] = mess
            back_btn = driver.find_element(By.CLASS_NAME, "realestate-object-modal__btn")
            back_btn.click()
            
            
            i += 1
            
            bot.edit_message_text(f"Обработано кад. номеров: {i} из {len(cad_num)}\nКад. номер: {numbers['cad_num']}", chat_id, mess_id)
            time.sleep(3)
        
        return {
            "data": cad_num,
            "start": date_start,
            "end": datetime.now().strftime("%d.%m.%Y %H:%M:%S"),
            "errs": errs
            }
    except Exception as ex:
        markup = types.InlineKeyboardMarkup()
        dow_btn = types.InlineKeyboardButton('Скачать', callback_data='download')
        con_btn = types.InlineKeyboardButton("Продолжить", callback_data='continue')
        markup.add(dow_btn, con_btn)
        
        bot.send_message(chat_id, "Во время получения данных из росреестра произошла ошибка!\n\nПродолжить дописать файл?", reply_markup=markup)
        
        @bot.callback_query_handler(func=lambda call: True)
        def callback_handler(callback: types.CallbackQuery):
            if callback.data == "download":
                xls.write_excel(file=filename, bot=bot, chat_id=chat_id, data=cad_num)
            else:
                parser_excel(cad_num, bot, chat_id,  callback.message.message_id, filename)
                
        bot.delete_message(chat_id, mess_id + 1)
        print(ex)
        return False
        
    finally:
        driver.close()
        driver.quit()


def parse_txt(cad_num: str, bot: telebot.TeleBot, message: types.Message):
    # os.path.join(os.getcwd(), "driver", "win", "geckodriver.exe") -> windows path
    dv_dir = os.path.join(os.getcwd(), "driver", "linux", "geckodriver")
    # r"C:\Program Files\Mozilla Firefox\firefox.exe" -> windows paht
    # bw_dir = r"usr/bin/firefox"
    cp_dir = os.path.join(os.getcwd())
    
    service = Service(executable_path=dv_dir, port=randint(1000, 10000))
    
    options = webdriver.FirefoxOptions()
    # options.binary_location = bw_dir
    options.add_argument('--headless')
    
    driver = webdriver.Firefox(service=service, options=options)
    URL = "https://lk.rosreestr.ru/eservices/real-estate-objects-online"
    driver.maximize_window()
    driver.get(URL)
    
    chat_id = message.chat.id
    bot.send_message(chat_id, "Начинаю обработку. Ожидайте...")
    mess_id = message.message_id + 1
    date_start = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    try: 
        time.sleep(10)
    
        with open(cp_dir + "/cp.png", 'wb') as file:
            l = driver.find_element(By.XPATH, '//*[@alt="captcha"]')
            file.write(l.screenshot_as_png)
        
        captcha = anticaptcha.solve_captcha(cp_dir + "/cp.png", bot, chat_id)
        if(captcha == False):
            return
        captch_input = driver.find_element(By.ID, "captcha")
        captch_input.clear()
        captch_input.send_keys(captcha)
        
        
        cad_input = driver.find_element(By.ID, "query")
        cad_input.clear()
        cad_input.send_keys(cad_num)
        
        err = driver.find_elements(By.CLASS_NAME, "rros-ui-lib-message--error")
        if(len(err) != 0):
            bot.edit_message_text("Ошибка! Не удалось решить капчу....\n\nПовторяю попытку...", chat_id, mess_id)
            driver.quit()
            driver.close()
            message.message_id + 2
            parse_txt(cad_num, bot, message)
            return False
                
        time.sleep(1)
        sch_btn = driver.find_element(By.ID, "realestateobjects-search")
                            
        sch_btn.click()
        
        time.sleep(4)
        
        card_btn = driver.find_elements(By.CLASS_NAME, "realestateobjects-wrapper__results__cadNumber")
        if(len(card_btn) == 0):
                bot.edit_message_text(f"По кад. номеру: {cad_num} нет информации.\nПроверте пожалуйста корректность ввода кад. номера", chat_id, mess_id)
                return
            
        card_btn[0].click()
        
        time.sleep(1)
        
        card = driver.find_elements(By.CLASS_NAME, "build-card-wrapper__info")
        
        mess = ""
        for elem in card:
            h3 = elem.find_element(By.TAG_NAME, "h3").text
            ul = elem.find_element(By.TAG_NAME, "ul")
            li = ul.find_elements(By.TAG_NAME, "li")
            mess += f"<b>{h3}\n</b>"
            for item in li:
                name = item.find_element(By.TAG_NAME, "span").text
                value = item.find_elements(By.CLASS_NAME, "build-card-wrapper__info__ul__subinfo__options__item__line")
                if(len(value) == 1): mess += f"    {name}: <b>{value[0].text}</b>\n"
                else:
                    mess += f"    {name}: "
                    for val in value:
                        mess += f"<b>{val.text}</b>\n"
        
        mess += f"""\nНачало обработки: {date_start}\nКонец обработки: {datetime.now().strftime("%d.%m.%Y %H:%M:%S")}\nБаланс антикапчи: {anticaptcha.get_balance()} $"""
        return mess.strip()
    except sl_exps.TimeoutException as err:
        bot.send_message(chat_id, "Ошибка! Сайт росреестра не отвечает", chat_id)
        log.write(f"{err.msg} | {__file__}")
        
    except sl_exps.NoSuchElementException as err:
        bot.send_message(chat_id, "Ошибка! Сайт росреестра не отвечает", chat_id)
        log.write(f"{err.msg} | {__file__}")
    except telebot.apihelper.ApiTelegramException as err:
        bot.send_message(chat_id, "Ошибка при работе бота. Повторите попытку...")
        log.write(f"{err.description} | {__file__}")
        
    finally:
        driver.close()
        driver.quit()
    

