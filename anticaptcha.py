from anticaptchaofficial.imagecaptcha import *
from telebot import TeleBot
from dotenv import load_dotenv

import os

import log


def solve_captcha(captcha_src: str, bot: TeleBot, chat_id: int):
    load_dotenv()
    API_KEY = os.getenv("CAPTCHA_API_KEY")
    solver = imagecaptcha()
    
    solver.set_verbose(1)
    solver.set_key(API_KEY)

    
    solver.set_soft_id(0)

    balance = solver.get_balance()
    print(balance)
    if balance <= 0:
        balance = solver.get_balance()
        if balance <= 0:
            bot.send_message(chat_id, 'Недостаточно средств на балансе антикапчи!\nПополните баланс антикапчи на https://anti-captcha.com/\nОперация прервано')
            return False

    captcha_text = solver.solve_and_return_solution(captcha_src)
    if captcha_text != 0:
        return captcha_text
    else:
        print("task finished with error " + solver.error_code)
        log.write(f"{solver.error_code} | {__file__}")
        if(solver.error_code == "ERROR_NO_SLOT_AVAILABLE"):
            bot.send_message(chat_id, "Ошибка антикапчи: Свободных работников на данный момент нет. Пожалуйста, попробуйте немного позже или увеличьте максимальную ставку в меню «Настройки» — «Настройка API» в зоне клиентов Anti-Captcha.\nОперация прервано")
        return False

def get_balance():
    solver = imagecaptcha()
    solver.set_verbose(1)
    solver.set_key("178480a4f2942503ca93d0836d11f2cb")

    # Specify softId to earn 10% commission with your app.
    # Get your softId here: https://anti-captcha.com/clients/tools/devcenter
    solver.set_soft_id(0)

    balance = solver.get_balance()
    return balance
