# Use this token to access the HTTP API:
# 746235750:AAF_68KwRrstDiSAoY0EtY8DJwRRYDlS98g
import telepot
from telepot.namedtuple import ReplyKeyboardMarkup, KeyboardButton
token='746235750:AAF_68KwRrstDiSAoY0EtY8DJwRRYDlS98g'
basic_auth = ('userid80ar', 'password')
SetProxy = telepot.api.set_proxy("https://46.101.110.133:80")
bot = telepot.Bot(token)


from telepot.loop import MessageLoop
from telepot.namedtuple import InlineKeyboardMarkup, InlineKeyboardButton


def handle(msg):

    content_type, chat_type, chat_id = telepot.glance(msg)
    text = msg["text"]
    if text =="Евгений":
        bot.sendMessage(chat_id, 'testing custom keyboard',
                        reply_markup=ReplyKeyboardMarkup(
                            keyboard=[
                                [KeyboardButton(text="Yes"), KeyboardButton(text="No")]
                            ]
                        ))
    elif text=="/start":
        bot.sendMessage(chat_id, "Введите номер площадки и описание работы через пробел")
    else:
        bot.sendMessage(chat_id, "Взяли в обработку, подождите...")

MessageLoop(bot, handle).run_as_thread()

# Keep the program running.
while True:
    n = input('To stop enter "stop":')
    if n.strip() == 'stop':
        break


# print(bot.handle())
# data=bot.getUpdates()
# content_type, chat_type, chat_id = telepot.glance(msg)
# bot.sendMessage("89613392","Я это сделал")
