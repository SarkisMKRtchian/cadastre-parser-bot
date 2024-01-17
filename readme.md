# **cadastre-parser-bot** - Телеграм бот для получения информации по кадастровому номеру из онлайн-справки Росреестра в виде карточки

### *Что умеет бот?* 
- [x] получить по кадастровому номеру информацию из онлайн-справки Росреестра в виде карточки
- [x] заполнить excel файл Раздел1 и Раздел7 карточками на всё здание


### *Бот* (Инициализация бота)
```python
# main.py
from telebot import types
import telebot
from random import randint
import pathlib
import os

from xls import read_xls # xls.py
import cd_parser # cd_parser.py
import log # log.py

# Инициализируте бота
TOKEN = "YOUR_TOKEN"
bot = telebot.TeleBot(TOKEN, parse_mode=None)
```

### *Антикапча* (Инициализация антикапчи)
**API_KEY** Можно получить [здесь](https://anti-captcha.com/)
```python
# anticaptch.py
def solve_captcha(captcha_src: str, bot: telebot.TeleBot, chat_id: int):
    API_KEY = "YOUR_API_KEY"
    solver = imagecaptcha()
    
    solver.set_verbose(1)
    solver.set_key(API_KEY)

    # Specify softId to earn 10% commission with your app.
    # Get your softId here: https://anti-captcha.com/clients/tools/devcenter
    solver.set_soft_id(0)

    balance = solver.get_balance()
    print(balance)
    if balance <= 0:
        balance = solver.get_balance()
        if balance <= 0:
            bot.send_message(chat_id, 'Пополните баланс антикапчи на https://anti-captcha.com/')
            return False

    captcha_text = solver.solve_and_return_solution(captcha_src)
    if captcha_text != 0:
        return captcha_text
    else:
        print("task finished with error " + solver.error_code)
```

### *Работа бота*
Откройте телеграм бота и нажмите страт. После выберите что нужно сделать **Парсинг карточек ЕГРН** или **Ввод кад. номера**

### **Парсинг карточек ЕГРН** 
Отправтье боту **excel Р1Р7** файл и ожидайте. Бот соберет всю информацию о **здании**, **обьекте** итд... и отправит вам отработаный excel файл.

### **Ввод кад. номера**
Отправтье боту кадастровый номер и ождийте. Бот соберет всю информацию о **здании**, **обьекте** итд... и отправит вам информацию в виде текста

