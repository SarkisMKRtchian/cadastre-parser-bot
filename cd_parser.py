import selenium.common.exceptions as sl_exps 
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from telebot import types
import telebot
from datetime import datetime
import anticaptcha
import os
import time
import os

import xls
import log

def parser_excel(cad_num: str, bot: telebot.TeleBot, chat_id: int, message_id: int, filename: str):
    # dv_dir = os.path.join(os.getcwd(), "driver", "win", "geckodriver.exe") -> win
    # bw_dir = r"C:\Program Files\Mozilla Firefox\firefox.exe" -> win
    dv_dir = os.path.join(os.getcwd(), "driver", "linux", "geckodriver")
    bw_dir = r"usr/bin/firefox"
    cp_dir = os.path.join(os.getcwd())
    
    service = Service(executable_path=dv_dir, port=randint(6000, 7000))
    
    # Настройка браузера
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
        
        time.sleep(10)
        for numbers in cad_num:
            if (numbers['mess'] != ''): 
                if(numbers['mess'] != 'nan'):
                    i += 1
                    bot.edit_message_text(f"Обработано кад. номеров: {i} из {len(cad_num)}\nКад. номер: {numbers['cad_num']}", chat_id, mess_id)
                    continue
            
            
            
            with open(cp_dir + "/cp.png", 'wb') as file:
                l = driver.find_element(By.XPATH, '//*[@alt="captcha"]')
                f.write(l.screenshot_as_png)
            
            captcha = anticaptcha.solve_captcha(cp_dir + "/cp.png", bot, chat_id)
            if(captcha == False):
                return False
            
            captch_input = driver.find_element(By.ID, "captcha")
            captch_input.clear()
            captch_input.send_keys(captcha)
            
            
            # Вводит кад. номер
            cad_input = driver.find_element(By.ID, "query")
            cad_input.clear()
            cad_input.send_keys(cad_num["cad_num"])
            
            err = driver.find_elements(By.CLASS_NAME, "rros-ui-lib-message--error")
            if(len(err) != 0):
                i += 1
                processed_failure.append(cad_num['cad_num'])
                driver.find_element(By.CLASS_NAME, "rros-ui-lib-captcha-content-reload-btn").click()
                time.sleep(3)
                continue
                    
            time.sleep(1)
            
            # Нажимает на кнопку поиск
            sch_btn = driver.find_element(By.ID, "realestateobjects-search")                    
            sch_btn.click()
            
            time.sleep(3)
            
            # Проверка ошибок на сайте росреестра
            search_err = driver.find_elements(By.CLASS_NAME, "rros-ui-lib-error-title")
            if(len(search_err) > 0):
                remove_messages(bot, chat_id, remove_messages_id)
                bot.send_message(chat_id, f"Внимание! Справка Росреестра выдаёт ошибку. Проверьте работу справки по ссылке\n{URL}\nОшибка: {search_err[0].text}")
                os.remove(file)
                return
            
            # Проверяет наличие обьектов по кад номеру, если их нет то пропускает иттерацию
            card_btn = driver.find_elements(By.CLASS_NAME, "realestateobjects-wrapper__results__cadNumber")
            if(len(card_btn) == 0):
                    cad_num['mess'] = ''
                    i += 1
                    processed_failure.append(cad_num['cad_num'])
                    continue
                
            card_btn[0].click()
            
            time.sleep(1)
            
            # Открывает карточку обьекта
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
             
            # Если данныз нет или их не удалось собрать то добавляет кн в массив неудачных кн
            if(mess == ''): processed_failure.append(cad_num['cad_num'])
            
            cad_num['mess'] = mess
            
            # Нажимает на кнопку назад
            back_btn = driver.find_element(By.CLASS_NAME, "realestate-object-modal__btn")
            back_btn.click()
            
            
            i += 1
            cad_nums_processed += 1
            bot.edit_message_text(f"Обработано кад. номеров: {i} из {len(cad_nums)}\nКад. номер: {cad_num['cad_num']}", chat_id, pr_id)
            
            time.sleep(3)
        
        return {
            "data": cad_num,
            "start": date_start,
            "end": datetime.now().strftime("%d.%m.%Y %H:%M:%S"),
            "errs": errs
            }
    except Exception as ex:
        driver.close()
        driver.quit()
        
        log.write(f"{ex} | {__file__}")
        
        err_mess = bot.send_message(chat_id, "Во время получения данных из росреестра произошла неизсветная ошибка!")
        
        remove_messages(bot, chat_id, remove_messages_id)
        if(cad_nums_processed > 0): 
            parse_excel_restart(file, cad_nums, bot, message, cad_nums_processed)
            bot.delete_message(chat_id, err_mess.message_id)
        
    finally:
        driver.close()
        driver.quit()


def parse_txt(cad_num: str, bot: telebot.TeleBot, message: types.Message):
    # os.path.join(os.getcwd(), "driver", "win", "geckodriver.exe") -> windows path
    dv_dir = os.path.join(os.getcwd(), "driver", "linux", "geckodriver")
    # r"C:\Program Files\Mozilla Firefox\firefox.exe" -> windows paht
    # bw_dir = r"usr/bin/firefox"
    cp_dir = os.path.join(os.getcwd())
    
    service = Service(executable_path=dv_dir, port=randint(6000, 7000))
    
    # Настройка браузера
    options = webdriver.FirefoxOptions()
    options.binary_location = bw_dir
    options.add_argument('--headless')
    
    driver = webdriver.Firefox(service=service, options=options)
    driver.maximize_window()
    # Открывает сайт росреестра
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
            file.close()
        
        captcha = anticaptcha.solve_captcha(cp_dir + "/cp.png", bot, chat_id)
        if(captcha == False):
            os.remove(cp_name)
            return
        
        # Вводит капчу
        captch_input = driver.find_element(By.ID, "captcha")
        captch_input.clear()
        captch_input.send_keys(captcha)
        # удаляет изб. капчи
        os.remove(cp_name)
        
        
        cad_input = driver.find_element(By.ID, "query")
        cad_input.clear()
        cad_input.send_keys(cad_num)
        
        # Проверка валид. капчи, если капча не правильная выходит из фнукции
        err = driver.find_elements(By.CLASS_NAME, "rros-ui-lib-message--error")
        if(len(err) != 0):
            bot.edit_message_text("Ошибка! Не удалось решить капчу....\n\nПовторите попытку...", chat_id, start_work.message_id)
            
            remove_messages(bot, chat_id, [message.message_id - 2, message.message_id - 1, message.message_id])
            
            return
                
        time.sleep(1)
        
        # Нажимает на кнопку поиск
        sch_btn = driver.find_element(By.ID, "realestateobjects-search")      
        sch_btn.click()
        
        time.sleep(3)
        
        # Проверка ошибок на сайте росреестра
        search_err = driver.find_elements(By.CLASS_NAME, "rros-ui-lib-error-title")
        if(len(search_err) > 0):
            remove_messages(bot, chat_id, remove_messages_id)
            bot.send_message(chat_id, f"Внимание! Справка Росреестра выдаёт ошибку. Проверьте работу справки по ссылке\n{URL}\nОшибка: {search_err[0].text}")
            os.remove(file)
            return
        
        # Проверяет наличие обьектов по кад номеру, если их нет то завершает фнукцию
        card_btn = driver.find_elements(By.CLASS_NAME, "realestateobjects-wrapper__results__cadNumber")
        if(len(card_btn) == 0):
            bot.edit_message_text(f"По кад. номеру: {cad_num} нет информации.\nПроверте пожалуйста корректность ввода кад. номера и повторите попытку", chat_id, start_work.message_id)
            
            remove_messages(bot, chat_id, [message.message_id - 2, message.message_id - 1, message.message_id])
            
            return
            
        # Открывает карточку обьекта
        card_btn[0].click()
        time.sleep(1)
        card = driver.find_elements(By.CLASS_NAME, "build-card-wrapper__info")
        
        # Забирает информацию из карточки
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
        driver.close()
        driver.quit()
        
        log.write(f"{err} | {__file__}")
        
        bot.send_message(chat_id, "Ошибка! Сайт росреестра не отвечает", chat_id)
        
        remove_messages(bot, chat_id, remove_messages_id)
        
    except sl_exps.NoSuchElementException as err:
        driver.close()
        driver.quit()
        
        log.write(f"{err} | {__file__}")
        
        bot.send_message(chat_id, "Ошибка! Сайт росреестра не отвечает", chat_id)
        log.write(f"{err.msg} | {__file__}")
    except telebot.apihelper.ApiTelegramException as err:
        bot.send_message(chat_id, "Ошибка при работе бота. Повторите попытку...")
        log.write(f"{err.description} | {__file__}")
        
    finally:
        driver.close()
        driver.quit()
    

# Проходит по всем id и удаляет их
def remove_messages(bot: telebot.TeleBot, chat_id: int, message_ids: list[int]):
    for i in range(0, len(message_ids)):
        try:
            bot.delete_message(chat_id, message_ids[i])
        except telebot.apihelper.ApiTelegramException as err:
            log.write(f"{err.description} | mess_id: {message_ids[i]} | {__file__}")
            continue

# Если во время обработки возникла ошибка, но при этом удалось собрать хоть какие-то данные, то отправляет пользователью сообщение, что может либо продолжить либо собрать xls файл
def parse_excel_restart(file: str, cad_nums: dict, bot: telebot.TeleBot, message: types.Message, processed: int):
    markup = types.InlineKeyboardMarkup()
    dow_btn = types.InlineKeyboardButton('Скачать', callback_data='download')
    con_btn = types.InlineKeyboardButton("Продолжить", callback_data='continue')
    markup.add(dow_btn, con_btn)
    
    mess = bot.send_message(message.chat.id, f"Бот успел обработать {processed} из {len(cad_nums)} КН. Хотите продолжить или скачать файл?", reply_markup=markup)
    
    @bot.callback_query_handler(func=lambda call: True)
    def callback_handler(callback: types.CallbackQuery):
        if callback.data == "download":
            xls.write_excel(file, bot, message.chat.id, cad_nums)
        else:
            parse_excel(cad_nums, bot, callback.message, file)
        time.sleep(1)
        bot.delete_message(message.chat.id, mess.message_id)
            
            