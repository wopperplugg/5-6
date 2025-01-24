#!/usr/bin/env python
# pylint: disable=unused-argument
import logging
import pandas as pd
import numpy as np
from telegram import ReplyKeyboardMarkup, ReplyKeyboardRemove, Update
from telegram.ext import (
    Application,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    filters,
)

# Enable logging
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO
)
# set higher logging level for httpx to avoid all GET and POST requests being logged
logging.getLogger("httpx").setLevel(logging.WARNING)

logger = logging.getLogger(__name__)

GENDER, ANALYZE = range(2)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Начинает беседу и спрашивает пользователя о его поле."""
    reply_keyboard = [["отчет по студентам", "% выполненого дз"]]

    await update.message.reply_text(
        "Привет! Меня зовут exel Бот. Я помогу вам с exel документами. "
        "Отправьте /cancel, чтобы прекратить общение со мной.\n\n",
        reply_markup=ReplyKeyboardMarkup(
            reply_keyboard, one_time_keyboard=True, input_field_placeholder="отчет по студентам или % выполненого дз?"
        ),
    )

    return GENDER


async def analitic(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Сохраняет выбранный пол и спрашивает о фото."""
    user = update.message.from_user
    logger.info("стадия пользователя %s: %s", user.first_name, update.message.text)
    await update.message.reply_text(
        "Понял! Пожалуйста, подождите"
    )
    try:
        otchet = pd.read_excel('C:/python/pythonProject/5.xlsx', skiprows=0)

        if update.message.text == "отчет по студентам":
            if 'Average score' in otchet.columns:
                otchet['Average score'] = pd.to_numeric(otchet['Average score'], errors='coerce')
                low_score_students = otchet[otchet['Average score'] < 3]

                if not low_score_students.empty:
                    response = "Студенты с низким средним баллом (ниже 3):\n"
                    result = low_score_students[['FIO', 'Группа']]
                    formatted_rows = [f"{row['FIO']:<40} {row['Группа']}" for _, row in result.iterrows()]
                    response += "\n".join(formatted_rows)
                else:
                    response = "Все студенты имеют средний балл 3 и выше."
            else:
                response = "Ошибка: Ожидаемый столбец 'Average score' отсутствует в отчете."

        elif update.message.text == "% выполненого дз":
            if 'Percentage Homework' in otchet.columns:
                low_homework_students = otchet[otchet['Percentage Homework'] < 50]
                if not low_homework_students.empty:
                    response = "Студенты с процентом выполненного домашнего задания ниже 50%:\n"
                    result = low_homework_students['FIO']  # Отбираем только 'FIO'
                    response += "\n".join(result)  # Соединяем ФИО в строку
                else:
                    response = "Все студенты выполнили более 50% домашнего задания."
            else:
                response = "Ошибка: Ожидаемый столбец 'Percentage Homework' отсутствует в отчете."

        # Разбиваем сообщение на части, если оно слишком длинное
        max_length = 4096
        while len(response) > max_length:
            await update.message.reply_text(response[:max_length])
            response = response[max_length:]
        await update.message.reply_text(response)

    except Exception as e:
        logger.error("Ошибка при чтении или обработке отчета: %s", e)
        response = "Произошла ошибка при анализе отчета. Пожалуйста, проверьте файл."
        await update.message.reply_text(response)

    return ConversationHandler.END

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Отменяет и завершает разговор."""
    user = update.message.from_user
    logger.info("Пользователь %s отменил разговор.", user.first_name)
    await update.message.reply_text(
        "До свидания! Надеюсь, мы сможем поговорить снова когда-нибудь.", reply_markup=ReplyKeyboardRemove()
    )

    return ConversationHandler.END

def main() -> None:
    """Run the bot."""
    application = Application.builder().token("7659203514:AAGcVvd-9Q_5e5IaE74tXhokykxBDiSd3C0").build()

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            GENDER: [MessageHandler(filters.Regex("^(отчет по студентам|% выполненого дз)$"), analitic)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    application.add_handler(conv_handler)
    application.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()