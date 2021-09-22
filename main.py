from openpyxl import load_workbook
import telebot
from telebot import types
from os import walk

token="1986086924:AAFNbyaH3lHwpIu9H_a_LmOEuqlrIrdKU8M"
base_path="catalog/"
table_path="main.xlsx"
table="main"
electronic=base_path+"Электронные сигареты/"
delivery=["Москва","Обнинск","РАНХиГС","РЭУ","РНИМУ","Бауманка"]
dan_id=1588645954
doctors_id=808525546 #степан id


def append_in_xlsx(path,table_name,row):
    try:
        wb = load_workbook(path)
        ws = wb[str(table_name)]
        ws.append(row)
        wb.save(path)
        wb.close()
        return True
    except:
        return False

def get_description(path):
    f=open(str(path)+'main.txt',"r",encoding="utf-8")
    array=f.readlines()
    Name=str(array[0]).replace("\n",'')
    Description=str(array[1]).replace("\n",'').replace("&",'\n')
    Taste=str(array[2]).replace("\n","").split(":")
    Price=str(array[3])
    Tastes=''
    for i in Taste:
        Tastes+=i+'  '
    return [Name+"\n"+Description+"\n"+"Цена: "+Price+" ₽", Name,Description,Taste,Price,path+"main.png"]

def get_dir(path):
    f = []
    for (dirpath, dirnames, filenames) in walk(path):
        f.extend(dirnames)
        f.extend(filenames)
        break
    return f

def Keyboard_Generator(buttons):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    for i in buttons:
        button = types.KeyboardButton(text=i)
        keyboard.add(button)
    return keyboard

def Inline_Keyboard_Generator(buttons):
    keyboard = types.InlineKeyboardMarkup(row_width=2)
    for i in buttons:
        button = types.InlineKeyboardButton(i[0],callback_data=i[1])
        keyboard.add(button)
    return keyboard

def command_worker(message,chat_id):
    global electronic
    if message=="/Электронные сигареты":
        mass=[]
        for i in get_dir(electronic):
            mass.append("/"+str(i))
        mass.append("На главную")
        bot.send_message(chat_id, "Выбери производителя👇", reply_markup=Keyboard_Generator(mass))

    if str(message).replace("/",'') in get_dir(electronic):
        bot.send_message(chat_id, "Выбери товар",reply_markup=Keyboard_Generator(["Назад к производителям"]))
        manufacturer=str(message).replace('/','')
        for i in get_dir(electronic+str(message).replace("/",'')+"/"):
            path=electronic + str(message).replace("/", '') + "/"+str(i)+"/"
            print(path)
            i=get_description(path)
            mass = []
            for t in i[3]:
                mass.append([f"{t}", f"E:{i[1]}:{manufacturer}:{t}"])
            keyboard = Inline_Keyboard_Generator(mass)
            bot.send_photo(chat_id, photo=open(i[5], 'rb'), caption=i[0], reply_markup=keyboard)


bot =telebot.TeleBot(token)

@bot.message_handler(commands=['start'])
def welcome(message):
    bot.send_message(message.chat.id, f"Добро пожаловать, {message.from_user.first_name} ! 🤍"
                                      f"\nТебя приветствует онлайн магазин Табачка у Дома!🔥\n"
                                      f"Наши соц.сети, где регулярно проходят акции:\n"
                                      f"Instagram: https://instagram.com/tabachka_777_40?utm_medium=copy_link\n"
                                      f"Делай заказ прямо сейчас, в удобном боте! 👇",
                     reply_markup=Keyboard_Generator(["/Электронные сигареты"]))


@bot.message_handler(content_types=['text'])
def get_message(message):
    chat_id=message.chat.id
    text=message.text
    print(text,"h",message.chat.id, message.from_user.first_name)
    if "/" in text:
        if text=="/admin_get_table" and (chat_id==dan_id or chat_id==doctors_id):
            db=open(table_path,"rb")
            bot.send_document(chat_id,db)

        command_worker(text,chat_id)
        print("/")

    if text=="Назад к производителям":
        mass = []
        for i in get_dir(electronic):
            mass.append("/" + str(i))
        mass.append("На главную")
        bot.send_message(chat_id, "Выбери производителя👇", reply_markup=Keyboard_Generator(mass))

    if text=="На главную":
        bot.send_message(message.chat.id, "Выбери товар",
                         reply_markup=Keyboard_Generator(["/Электронные сигареты"]))

@bot.callback_query_handler(func=lambda call: True)
def Callback_inline(call):
    if call.message:
        chat_id = call.message.chat.id
        data=str(call.data)
        print(data,"i")
        if data.split(":")[0]=="E":
            mass=[]
            for i in delivery:
                mass.append([i,"D:"+str(data)+":"+i])
            bot.send_message(chat_id,"Выберите доствку:",reply_markup=Inline_Keyboard_Generator(mass))

        if data.split(":")[0] == "D":
            name=data.split(":")[2]
            manufacturer=data.split(":")[3]
            taste=data.split(":")[4]
            place = data.split(":")[5]
            bot.edit_message_text(chat_id=call.message.chat.id,
                                  message_id=call.message.message_id,
                                  text=f"Спасибо. Мы приняли ваш заказ:\n"
                                       f"✅ Производитель: {manufacturer}\n"
                                       f"✅ Название: {name}\n"
                                       f"✅ Вкус: {taste}\n"
                                       f"✅ Доставка: {place}\n",
                                  reply_markup=None)
            bot.send_message(chat_id,
                             "Скоро мы с вами свяжемся! ⚡",
                             reply_markup=Keyboard_Generator(["/Электронные сигареты"]))

            #-----------------------------------------------export------------------------------
            username=call.message.chat.username
            first_name=call.message.chat.first_name
            mass=(manufacturer,name,taste,place,username,first_name,chat_id)
            try:
                append_in_xlsx(table_path,table,mass)
                bot.send_message(dan_id,f"Новый заказ:\n{manufacturer}\n{name}\n{taste}\n{place}\n{username}\n{first_name}\n{chat_id}")
            except:
                bot.send_message(dan_id,f"WRITE ERRR\nНовый заказ:\n{manufacturer}\n{name}\n{taste}\n{place}\n{username}\n{first_name}\n{chat_id}")



bot.polling(none_stop=True)