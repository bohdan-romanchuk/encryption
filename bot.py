import telebot, xlrd, xlwt
from xlutils.copy import copy

TOKEN = ""
    
bot = telebot.TeleBot(TOKEN)

# Таблиця, у якій зберігаються шифри
rb = xlrd.open_workbook('./crypt.xls')
wb = copy(rb)

sheet = rb.sheet_by_index(0)
w_sheet = wb.get_sheet(0)

row_number = sheet.nrows


# Функція шифрування
def encrypt(textForEncrypt, typeOfEncrypt, key):

    encryptedMessage = ""

    # Шифр Цезаря
    if (typeOfEncrypt == 0):
        for i in textForEncrypt:
            if (ord(i) + key > ord("я")):
                encryptedMessage += chr(ord(i) + key - 32)
            else:
                encryptedMessage += chr(ord(i) + key)

        encryptedMessage += "#" + str(key)
    
    # Трансліт / Недоазбука Морзе :) / Смайлики 
    else:
        for i in textForEncrypt:
            for j in range(row_number):
                if (i == str(sheet.cell(j, 0).value)):
                    encryptedMessage += str(sheet.cell(j, typeOfEncrypt).value)

    # Якщо тип шифрування - не Трансліт - додати ключ для можливості дешифрування
    if (typeOfEncrypt != 1):
        encryptedMessage += "#" + str(typeOfEncrypt) + "#" 

    return encryptedMessage



# Функція дешифрування
def decrypt(textForDecrypt):
    
    decryptedMessage = ""
    
    # Дешифрування Шифру Цезаря
    if (textForDecrypt[len(textForDecrypt) - 3:] == "#0#"):

        key = int(str(textForDecrypt[len(textForDecrypt) - 5 : len(textForDecrypt) - 3]))

        for i in textForDecrypt:
            if (ord(i) - key < ord("а") and ord(i) > 900):
                decryptedMessage += chr(ord(i) - key + 32)
            else:
                decryptedMessage += chr(ord(i) - key)

        return decryptedMessage[: len(textForDecrypt) - 6]

    # Дешифрування (майже)Азбуки Морзе
    elif (textForDecrypt[len(textForDecrypt) - 3:] == "#2#"):
        for i in range(int((len(textForDecrypt) - 3) / 13)):
            for j in range(row_number):
                if textForDecrypt[13 * i : 13 * (i + 1)] == str(sheet.cell(j, 2).value):
                    decryptedMessage += str(sheet.cell(j, 0).value)
        
        return decryptedMessage

    # Дешифрування Смайликів
    elif (textForDecrypt[len(textForDecrypt) - 3:] == "#3#"):
        for i in textForDecrypt:
            for j in range(row_number):
                if i == str(sheet.cell(j, 3).value):
                    decryptedMessage += str(sheet.cell(j, 0).value)

        return decryptedMessage

    else:
        return "На жаль, неможливо розшифрувати :("



@bot.message_handler(commands=["help"])
def start(message):
    bot.send_message(message.chat.id, "/start - початок використання\n/help - допомога\n/encrypt - вибір методу шифрування\n/decrypt - дешифрування")

@bot.message_handler(commands=['start'])
def handle_start(message):
    user_markup = telebot.types.ReplyKeyboardMarkup(True, False)
    user_markup.row('Шифрування', 'Дешифрування')
    bot.send_message(message.from_user.id, 'Обери, що ти хочеш.', reply_markup=user_markup) 

@bot.message_handler(content_types=['text'])
def get_text_messages(message):

    if (str(message.text) == "Шифрування" or str(message.text) == "/encrypt"):
        user_markup = telebot.types.ReplyKeyboardMarkup(True, False)
        user_markup.row('Шифр Цезаря', 'Трансліт')
        user_markup.row('Азбука Морзе', 'Смайли')
        bot.send_message(message.from_user.id, 'Вибери тип шифрування.', reply_markup=user_markup)
        

    elif (str(message.text) == "Дешифрування" or str(message.text) == "/crypt"):
        bot.send_message(message.chat.id, "Введи текст для дешифрування.", reply_markup=telebot.types.ReplyKeyboardRemove())  

    elif (str(message.text) == "Шифр Цезаря"):
        w_sheet.write(0, 5, "0")
        wb.save('crypt.xls')

        bot.send_message(message.chat.id, "Введи число від 1 до 31.", reply_markup=telebot.types.ReplyKeyboardRemove())
    
    elif (str(message.text) == "Трансліт"):
        w_sheet.write(0, 5, "1")
        wb.save('crypt.xls')
        bot.send_message(message.chat.id, "Введи текст для шифрування", reply_markup=telebot.types.ReplyKeyboardRemove())
    
    elif (str(message.text) == "Азбука Морзе"):
        w_sheet.write(0, 5, "2")
        wb.save('crypt.xls')

        bot.send_message(message.chat.id, "Введи текст для шифрування", reply_markup=telebot.types.ReplyKeyboardRemove())
    
    elif (str(message.text) == "Смайли"):
        w_sheet.write(0, 5, "3")
        wb.save('crypt.xls')

        bot.send_message(message.chat.id, "Введи текст для шифрування (не бажано використовувати смайлики))", reply_markup=telebot.types.ReplyKeyboardRemove())
                
    elif (len(str(message.text)) <= 2 and int(message.text) < 32):            
        w_sheet.write(0, 4, str(message.text))
        wb.save('crypt.xls')

        bot.send_message(message.chat.id, "Введи текст для шифрування")

    elif (str(message.text)[len(str(message.text)) - 1] == "#"):
        mess = decrypt(str(message.text))
        bot.send_message(message.from_user.id, mess)
        
        user_markup = telebot.types.ReplyKeyboardMarkup(True, False)
        user_markup.row('Шифрування', 'Дешифрування')
        bot.send_message(message.from_user.id, 'Обери, що ти хочеш.', reply_markup=user_markup)
                
    else:
        ob = xlrd.open_workbook('./crypt.xls')
        sh = ob.sheet_by_index(0)  

        key = int(sh.cell(0, 4).value)
        typeOfEncr = int(sh.cell(0, 5).value) 

        mess = encrypt(str(message.text), typeOfEncr, key)        
        bot.send_message(message.from_user.id, mess)

        w_sheet.write(0, 4, "0")
        w_sheet.write(0, 5, "1")
        wb.save('crypt.xls')

        user_markup = telebot.types.ReplyKeyboardMarkup(True, False)
        user_markup.row('Шифрування', 'Дешифрування')
        bot.send_message(message.from_user.id, 'Обери, що ти хочеш.', reply_markup=user_markup)


if (__name__ == "__main__"):
    bot.polling()
