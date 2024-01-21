from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from urllib3.exceptions import MaxRetryError
from requests.exceptions import ReadTimeout
from telebot import TeleBot, types
from typing import Tuple
from random import randint

import selenium.common.exceptions as sl_exps 
import anticaptcha
import telebot
import time
import json
import os
import re

import log

class Parse:
    
    def __init__(self, bot: TeleBot) -> None:
        
            
        self.bot = bot
        self.isWork = False
        self.stop = False
        
    def removeMessages(self, chatId: int, messageId: list[int]):
        
        for i in range(len(messageId)):
            try:
                self.bot.delete_message(chatId, messageId[i])
            except telebot.apihelper.ApiTelegramException as err:
                log.write(f"{err.description} | mess_id: {messageId[i]} | {__file__}")
                continue
            
    def browser(self):
        self.isWork = True
        dv_dir = os.path.join(os.getcwd(), "driver", "linux", "geckodriver")
        bw_dir = r"/usr/lib/firefox-esr/firefox-esr"
        
        service = Service(executable_path=dv_dir, port=randint(6000, 7000))
        
        options = webdriver.FirefoxOptions()
        options.binary_location = bw_dir
        options.add_argument('--headless')
        
        driver = webdriver.Firefox(service=service, options=options)
        driver.maximize_window()
        driver.get(self.url())
        return driver
    
    def closeBrowser(self, browser: webdriver.Firefox):
        
        browser.close()
        browser.quit()
        self.isWork = False
    
    def browserWait(self, browser: webdriver.Firefox,by: Tuple[str, str], attr: str):
        
        wait = WebDriverWait(browser, 500)
        wait.until(EC.element_attribute_to_include(by, attr))

    def url(self):
        
        return "https://lk.rosreestr.ru/eservices/real-estate-objects-online"
    
    def sloveCaptcha(self, browser: webdriver.Firefox, chatId: int):
        
        cpDir = os.path.join(os.getcwd(), f"{randint(1, 10)}.png") 
        with open(cpDir, 'wb') as file:
            cp = browser.find_element(By.XPATH, '//*[@alt="captcha"]')
            file.write(cp.screenshot_as_png)
            file.close()
        
        
        captcha = anticaptcha.solve_captcha(cpDir, self.bot, chatId)
        os.remove(cpDir)
        return captcha
    
    def parseReestr(self, browser: webdriver.Firefox, cadNum: str, chatId: int, isArray: bool = False):
       
        if(isArray):
            backBtn = browser.find_elements(By.CLASS_NAME, "realestate-object-modal__btn")
            if(len(backBtn) > 0):
                backBtn[0].click()
            
        self.browserWait(browser, (By.XPATH, '//*[@alt="captcha"]'), 'src')
        reestrErr = browser.find_elements(By.CLASS_NAME, "rros-ui-lib-error-title")
        if(len(reestrErr) > 0):
            return {
                "status": False,
                "mess": f"Внимание! Справка Росреестра выдаёт ошибку. Проверьте работу справки по ссылке\n{self.url()}\nОшибка: {reestrErr[0].text}"
            }
        
        captcha = self.sloveCaptcha(browser, chatId)
        if(captcha == False):
            return {
                "status": False,
                "mess": "captcha input err"
            }
        
        inputCaptcha = browser.find_element(By.ID, "captcha")
        inputCaptcha.clear()
        inputCaptcha.send_keys(captcha)
            
            
        inputCadNum = browser.find_element(By.ID, "query")
        inputCadNum.clear()
        inputCadNum.send_keys(cadNum)

                
        
        browser.find_element(By.TAG_NAME, "body").click()
        
        schBtn = browser.find_element(By.ID, "realestateobjects-search")      
        
        schBtn.click()
        
        
        errInputCaptcha = browser.find_elements(By.CLASS_NAME, "rros-ui-lib-message--error")
        if(len(errInputCaptcha) != 0):
            if(errInputCaptcha[0].text == "Текст введен неверно"):
                return {
                    "status": False,
                    "mess": "Не удалось решить каптчу"
                }
        
        
        searching = True
        while searching:
            spinner = browser.find_elements(By.CLASS_NAME, "rros-ui-lib-spinner__wrapper")
            if(len(spinner) == 0):
                searching = False
            time.sleep(0.1)
        
        
        searchErr = browser.find_elements(By.CLASS_NAME, "rros-ui-lib-error-title")
        if(len(searchErr) > 0):
            return {
                "status": False,
                "mess": f"Внимание! Справка Росреестра выдаёт ошибку. Проверьте работу справки по ссылке\n{self.url()}\nОшибка: {searchErr[0].text}"
            }
        
        
        cardBtn = browser.find_elements(By.CLASS_NAME, "realestateobjects-wrapper__results__cadNumber")
        if(len(cardBtn) == 0):
            return {
                "status": False,
                "mess": f"По кад. номеру: {cadNum} нет информации.\nПроверте пожалуйста корректность ввода кад. номера и повторите попытку"
            }
            
        
        cardBtn[0].click()
        card = browser.find_elements(By.CLASS_NAME, "build-card-wrapper__info")
        
        
        message = ""
        for elem in card:
            h3 = elem.find_element(By.TAG_NAME, "h3").text
            ul = elem.find_element(By.TAG_NAME, "ul")
            li = ul.find_elements(By.TAG_NAME, "li")
            message += f"<b>{h3}\n</b>"
            for item in li:
                name = item.find_element(By.TAG_NAME, "span").text
                value = item.find_elements(By.CLASS_NAME, "build-card-wrapper__info__ul__subinfo__options__item__line")
                if(len(value) == 1): message += f"    {name}: <b>{value[0].text}</b>\n"
                else:
                    message += f"    {name}: "
                    for val in value:
                        message += f"<b>{val.text}</b>\n"
            
            
        return {
            "status": True,
            "mess": re.sub(r"[<b>|</b>|\n|    ]", '',message).strip() if isArray else message 
        }
            

class ParseTxt(Parse):
    
    def __init__(self, bot: TeleBot) -> None:
        super().__init__(bot)
        
    
    def parse(self, cadNum, message: types.Message):
        
        try:
            startTime = time.time()
            dltMessages = [message.message_id - 2, message.message_id - 1, message.message_id]
            balance = anticaptcha.get_balance()
            startProcessingMess = self.bot.send_message(message.chat.id, f"Начинаю обработку. Ожидайте...\nДата начала обработки: {time.strftime('%H:%M:%S', time.localtime(startTime))}\nБаланс антикапчи: {balance} $")
            browser = self.browser()
            parse = self.parseReestr(browser, cadNum, message.chat.id)
            
            if(parse['status']):
                result = parse['mess']
                endTime = time.time()
                resMessage = f"{result}\n\nОбработка карточки заняло: {round(endTime - startTime, 1)} сек"
                self.bot.send_message(message.chat.id, resMessage, parse_mode='html')
            else:
                if(parse['mess'] != 'captcha input err'):
                    self.bot.send_message(message.chat.id, f"КН: {cadNum}\n{parse['mess']}")

        except sl_exps.TimeoutException as ex:
            log.write(f"{ex} | {__file__}")
            
            self.bot.send_message(message.chat.id, "Ошибка! Сайт росреестра не отвечает")
        
        except sl_exps.NoSuchElementException as ex:
            
            log.write(f"{ex} | {__file__}")
            
            self.bot.send_message(message.chat.id, "Ошибка! Сайт росреестра не отвечает")
            
        except ReadTimeout as ex:
            log.write(f"{ex} | {__file__}")
            
            self.bot.send_message(message.chat.id, "Ошибка! Сайт росреестра не отвечает из-за слабого интернет соеденения")
            
        except MaxRetryError as err:
            log.write(f"{err} | {__file__}")
            
            edt = self.bot.edit_message_text(message.chat.id, "Ошибка! Росреестр разорвал подключение! Ождиаю 5 секунд и поторяю попытку...")
            
            time.sleep(5)
            
            self.parse(cadNum, edt)
            
        except Exception as err:
            log.write(f"{err} | {__file__}")
            
            self.bot.send_message(message.chat.id, "Во время получения данных из росреестра произошла неизсветная ошибка! Повторите попытку")
        finally:
            dltMessages.append(startProcessingMess.message_id)
            self.removeMessages(message.chat.id, dltMessages)
            self.closeBrowser(browser)
            
class ParseExcel(Parse):
    
    def __init__(self, bot: TeleBot) -> None:
        super().__init__(bot)
        
    def parse(self, cadNums: dict, filePath: str, message: types.Message, xlsx):
        
        try:
            startTime = time.time()
            stopBtn = types.InlineKeyboardMarkup().add(types.InlineKeyboardButton('Отменить', callback_data='stop'))
            balance = anticaptcha.get_balance()
            startProcessingMess = self.bot.send_message(message.chat.id, f"Начинаю обработку. Ожидайте...\nДата начала обработки: {time.strftime('%H:%M:%S', time.localtime(startTime))}\nБаланс антикапчи: {balance} $", reply_markup=stopBtn)
            browser = self.browser()
            
            processed = 0
            itteration = 0
            captchaSloved = 0
            processedFailure = []
            processedMess = self.bot.send_message(message.chat.id, f"Обработано кад. номеров: {itteration} из {len(cadNums)}")
            
            dltMessages = [message.message_id - 2, message.message_id - 1, message.message_id, startProcessingMess.message_id, processedMess.message_id]
            for cadNum in cadNums:
                if(self.stop):
                    break
                if (cadNum['mess'] != ''):
                    if(cadNum['mess'] != 'nan'): 
                        itteration += 1
                        processed += 1
                        self.bot.edit_message_text(f"Обработано кад. номеров: {itteration} из {len(cadNums)}\nКад. номер: {cadNum['cadNum']}", message.chat.id, processedMess.message_id)
                        continue
                    
                parse = self.parseReestr(browser, cadNum['cadNum'], message.chat.id, True)
                
                if(parse['status']):
                    cadNum['mess'] = parse['mess']
                    processed += 1
                    itteration += 1
                    captchaSloved += 1
                else:
                    if(parse['mess'] != 'captcha input err'):
                        captchaSloved += 1
                    processedFailure.append(cadNum['cadNum'])
                    itteration += 1
                    
                self.bot.edit_message_text(f"Обработано кад. номеров: {itteration} из {len(cadNums)}\nКад. номер: {cadNum['cadNum']}", message.chat.id, processedMess.message_id)
            
            
            data = self.createDataObj(cadNums, startTime, time.time(), processedFailure, processed, filePath, balance, captchaSloved)
            
            xlsx.write(filePath, data, message.chat.id)
            
        except MaxRetryError as ex:
            errMess = self.bot.edit_message_text(chat_id=message.chat.id, message_id=processedMess.message_id, text="Ошибка! Росреестр разорвал подключение! Ождиаю 5 секунд и поторяю попытку...")
            
            time.sleep(5)
            
            self.parse(cadNums, filePath, errMess)
        
        except ReadTimeout as ex:
            log.write(f"{ex} | {__file__}")
            
            self.bot.send_message(message.chat.id, "Ошибка! Сайт росреестра не отвечает из-за слабого интернет соеденения")
            
        except sl_exps.TimeoutException as ex:
            log.write(f"{ex.msg} | {__file__}")
            
            errMess = self.bot.send_message(message.chat.id, "Ошибка! Сайт росреестра не отвечает")
            
            if(processed > 0): 
                self.bot.send_message(message.chat.id, f"Бот успел обработать {processed} из {len(cadNums)} КН. Хотите продолжить или скачать файл?", reply_markup=self.parseErrButtons())
                data = self.createDataObj(cadNums, startTime, time.time(), processedFailure, processed, filePath, balance, captchaSloved)
                self.crateJson(data)
            
        except sl_exps.NoSuchElementException as ex:
            
            log.write(f"{ex.msg} | {__file__}")
            
            errMess = self.bot.send_message(message.chat.id, "Ошибка! Сайт росреестра не отвечает")
            
            
            if(processed > 0): 
                self.bot.send_message(message.chat.id, f"Бот успел обработать {processed} из {len(cadNums)} КН. Хотите продолжить или скачать файл?", reply_markup=self.parseErrButtons())
                data = self.createDataObj(cadNums, startTime, time.time(), processedFailure, processed, filePath, balance, captchaSloved)
                self.crateJson(data)
                
        except Exception as ex:
            log.write(f"{ex} | {__file__}")
            errMess = self.bot.send_message(message.chat.id, "Во время получения данных из росреестра произошла неизсветная ошибка!")
            
            if(processed > 0): 
                self.bot.send_message(message.chat.id, f"Бот успел обработать {processed} из {len(cadNums)} КН. Хотите продолжить или скачать файл?", reply_markup=self.parseErrButtons())
                data = self.createDataObj(cadNums, startTime, time.time(), processedFailure, processed, filePath, balance, captchaSloved)
                self.crateJson(data)
            
        finally: 
            self.stop = False
            self.isWork = False
            self.removeMessages(message.chat.id, dltMessages)
            self.closeBrowser(browser)
            
    def parseErrButtons(self):
        """
        Создает кнопки для бота, если при парсинге возникла ошибка
        """
        markup = types.InlineKeyboardMarkup()
        dow_btn = types.InlineKeyboardButton('Скачать', callback_data='download')
        con_btn = types.InlineKeyboardButton("Продолжить", callback_data='continue')
        markup.add(dow_btn, con_btn)
        
        return markup
    
    def crateJson(self, data: dict):
        
        dir = os.path.join(os.getcwd(), 'cadNums.JSON')
        with open(dir, 'w') as file:
            json.dump(data, file)
            
    def createDataObj(self, cadNums: dict, startTime: float, endTime: float, processedFailure: list[str], processed: int, filePath: str, balance: float, captchSloved: int):
        
        afterBalance = anticaptcha.get_balance()
        timeForOneCard = round((endTime - startTime) / processed, 1) if processed != 0 else round(endTime - startTime, 1)
        cost = balance - afterBalance
        return {
            "data": cadNums,
            "stoped": self.stop,
            "start": startTime,
            "end": endTime,
            "processedFailure": processedFailure,
            "processed": f"Обработано КН: {processed} из {len(cadNums)} (1кн: {timeForOneCard}c./{round((cost / captchSloved), 5)}$)",
            "cost": round(cost, 5),
            "fp": filePath
        }