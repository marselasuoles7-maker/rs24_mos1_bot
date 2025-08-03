
from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import Updater, CommandHandler, MessageHandler, Filters, ConversationHandler, CallbackContext
import pandas as pd
import smtplib
import random
import string
from datetime import datetime, time
import os

from config import BOT_TOKEN, ADMINS, role_config

FILIAl, DEPARTMENT, FULLNAME, PHONE, CODE, MAIN_MENU, MEETING_TYPE, CLIENT_NAME, VENDOR_NAME, SHIPMENT = range(10)

user_sessions = {}
meeting_data = []
confirmation_codes = {}

def load_structure_from_excel():
    df = pd.read_excel("structure_phone.xlsx")
    structure = {}
    phone_lookup = {}
    for _, row in df.iterrows():
        filial = str(row["Филиал"]).strip()
        dept = str(row["Отдел"]).strip()
        fio = str(row["ФИО"]).strip()
        phone = ''.join(filter(str.isdigit, str(row["Телефон"])))
        role = str(row["Должность"]).strip()

        structure.setdefault(filial, {}).setdefault(dept, {})[phone] = {
            "fio": fio,
            "role": role
        }
        phone_lookup[phone] = {
            "filial": filial,
            "department": dept,
            "fio": fio,
            "role": role
        }
    return structure, phone_lookup

structure_data, phone_lookup = load_structure_from_excel()
FILIALS = list(structure_data.keys())

def start(update: Update, context: CallbackContext):
    update.message.reply_text("Введите номер телефона для авторизации:")
    return PHONE

def get_phone(update: Update, context: CallbackContext):
    phone = ''.join(filter(str.isdigit, update.message.text))
    user_info = phone_lookup.get(phone)
    if not user_info:
        update.message.reply_text("Телефон не найден. Обратитесь к администратору.")
        return ConversationHandler.END
    code = ''.join(random.choices(string.digits, k=6))
    confirmation_codes[update.effective_user.id] = (phone, code)
    context.user_data["user_info"] = user_info
    update.message.reply_text(f"Ваш код: {code} (в проде сюда вместо этого отправляется SMS)")
    return CODE

def get_code(update: Update, context: CallbackContext):
    user_id = update.effective_user.id
    input_code = update.message.text.strip()
    phone, correct_code = confirmation_codes.get(user_id, (None, None))
    if input_code == correct_code:
        user_info = context.user_data["user_info"]
        user_sessions[user_id] = {
            "phone": phone,
            **user_info
        }
        reply_keyboard = [["Отчет"]]
        update.message.reply_text("Успешно авторизованы.", reply_markup=ReplyKeyboardMarkup(reply_keyboard, resize_keyboard=True))
        return MAIN_MENU
    else:
        update.message.reply_text("Неверный код. Попробуйте снова:")
        return CODE

def main_menu(update: Update, context: CallbackContext):
    text = update.message.text
    if text == "Отчет":
        reply_keyboard = [["Менеджер–Клиент"], ["Менеджер–Клиент–Вендор"], ["📦 Отгрузка за день"]]
        update.message.reply_text("Выберите тип встречи:", reply_markup=ReplyKeyboardMarkup(reply_keyboard, resize_keyboard=True))
        return MEETING_TYPE
    else:
        update.message.reply_text("Пожалуйста, выберите действие из меню.")
        return MAIN_MENU

def meeting_type(update: Update, context: CallbackContext):
    user_sessions[update.effective_user.id]["meeting_type"] = update.message.text
    update.message.reply_text("Введите название клиента:", reply_markup=ReplyKeyboardRemove())
    return CLIENT_NAME

def client_name(update: Update, context: CallbackContext):
    user_sessions[update.effective_user.id]["client"] = update.message.text
    if user_sessions[update.effective_user.id]["meeting_type"] == "Менеджер–Клиент–Вендор":
        update.message.reply_text("Введите название вендора:")
        return VENDOR_NAME
    else:
        return save_meeting(update, context)

def vendor_name(update: Update, context: CallbackContext):
    user_sessions[update.effective_user.id]["vendor"] = update.message.text
    return save_meeting(update, context)

def save_meeting(update: Update, context: CallbackContext):
    session = user_sessions[update.effective_user.id]
    meeting = {
        "user_id": update.effective_user.id,
        "phone": session.get("phone"),
        "fio": session.get("fio"),
        "filial": session.get("filial"),
        "department": session.get("department"),
        "role": session.get("role"),
        "meeting_type": session.get("meeting_type"),
        "client": session.get("client"),
        "vendor": session.get("vendor", "-"),
        "shipment": session.get("shipment", "-"),
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M")
    }
    meeting_data.append(meeting)
    update.message.reply_text("Данные сохранены.", reply_markup=ReplyKeyboardMarkup([["Отчет"]], resize_keyboard=True))
    return MAIN_MENU

def shipment_handler(update: Update, context: CallbackContext):
    update.message.reply_text("Введите объем отгрузки за день (в рублях):", reply_markup=ReplyKeyboardRemove())
    return SHIPMENT

def shipment_value(update: Update, context: CallbackContext):
    user_sessions[update.effective_user.id]["shipment"] = update.message.text
    return save_meeting(update, context)

def send_excel_report(context: CallbackContext):
    if not meeting_data:
        return
    df = pd.DataFrame(meeting_data)

    for user_id, session in user_sessions.items():
        role = session.get("role")
        if role not in role_config:
            continue
        scope = role_config[role]
        if scope == "all":
            df_filtered = df
        elif scope == "filial":
            df_filtered = df[df["filial"] == session["filial"]]
        elif scope == "department":
            df_filtered = df[df["department"] == session["department"]]
        else:
            continue

        if df_filtered.empty:
            continue

        filename = f"report_{user_id}_{datetime.now().strftime('%Y%m%d')}.xlsx"
        df_filtered.to_excel(filename, index=False)
        context.bot.send_document(chat_id=user_id, document=open(filename, "rb"))
        os.remove(filename)

def cancel(update: Update, context: CallbackContext):
    update.message.reply_text("Выход из бота.", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END

def main():
    updater = Updater(BOT_TOKEN, use_context=True)
    dp = updater.dispatcher

    job = updater.job_queue
    job.run_daily(send_excel_report, time=time(17, 58))

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            PHONE: [MessageHandler(Filters.text & ~Filters.command, get_phone)],
            CODE: [MessageHandler(Filters.text & ~Filters.command, get_code)],
            MAIN_MENU: [MessageHandler(Filters.text & ~Filters.command, main_menu)],
            MEETING_TYPE: [MessageHandler(Filters.text & ~Filters.command, meeting_type)],
            CLIENT_NAME: [MessageHandler(Filters.text & ~Filters.command, client_name)],
            VENDOR_NAME: [MessageHandler(Filters.text & ~Filters.command, vendor_name)],
            SHIPMENT: [MessageHandler(Filters.text & ~Filters.command, shipment_value)],
        },
        fallbacks=[CommandHandler("cancel", cancel)]
    )

    dp.add_handler(conv_handler)
    updater.start_polling()
    updater.idle()

if __name__ == '__main__':
    main()
