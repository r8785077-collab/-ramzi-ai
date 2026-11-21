
import os
import logging
from io import BytesIO

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, CallbackQueryHandler,
    MessageHandler, ContextTypes, filters
)
from openai import AsyncOpenAI
from docx import Document

logging.basicConfig(level=logging.INFO)

BOT_TOKEN = os.getenv("BOT_TOKEN")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

client = AsyncOpenAI(api_key=OPENAI_API_KEY)

MAIN_MENU = InlineKeyboardMarkup([
    [InlineKeyboardButton("📚 إنشاء بحث مدرسي", callback_data="research")],
    [InlineKeyboardButton("📄 إنشاء سيرة ذاتية CV", callback_data="cv")],
    [InlineKeyboardButton("📝 تلخيص درس", callback_data="summary")],
    [InlineKeyboardButton("ℹ️ مساعدة", callback_data="help")]
])

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text(
        "🤖 أهلاً بك في RAMZI AI\n\n"
        "مساعدك الذكي لإنشاء البحوث المدرسية والسيرة الذاتية.\n\n"
        "اختر الخدمة:",
        reply_markup=MAIN_MENU
    )

async def menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    context.user_data["mode"] = query.data

    if query.data == "research":
        msg = "📚 اكتب عنوان البحث والمستوى الدراسي واللغة.\nمثال: بحث عن التلوث، السنة الرابعة متوسط، بالعربية."
    elif query.data == "cv":
        msg = "📄 أرسل معلوماتك لإنشاء CV:\nالاسم، الدراسة أو الشهادة، الخبرة، المهارات، اللغات، ورقم الهاتف إن أردت."
    elif query.data == "summary":
        msg = "📝 أرسل النص أو الدرس الذي تريد تلخيصه."
    else:
        msg = "RAMZI AI يساعدك على إنشاء البحوث والسيرة الذاتية وتلخيص الدروس."

    await query.message.reply_text(msg)

async def generate_text(prompt):
    response = await client.responses.create(
        model="gpt-4.1-mini",
        input=prompt
    )
    return response.output_text

async def create_docx(title, content):
    doc = Document()
    doc.add_heading(title, level=1)

    for paragraph in content.split("\n"):
        if paragraph.strip():
            doc.add_paragraph(paragraph)

    file = BytesIO()
    doc.save(file)
    file.seek(0)
    file.name = "RAMZI_AI.docx"
    return file

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    mode = context.user_data.get("mode")
    user_text = update.message.text

    if mode not in ["research", "cv", "summary"]:
        await update.message.reply_text(
            "اختر خدمة من القائمة أولاً.",
            reply_markup=MAIN_MENU
        )
        return

    await update.message.reply_text("⏳ جاري إعداد طلبك، انتظر قليلاً...")

    if mode == "research":
        prompt = (
            "اكتب بحثاً مدرسياً منظماً باللغة المطلوبة بناءً على الطلب التالي. "
            "اجعله يحتوي على عنوان ومقدمة وعناصر واضحة وخاتمة ومراجع مقترحة. "
            "لا تختلق مراجع مؤكدة.\n\n" + user_text
        )
        title = "بحث مدرسي"

    elif mode == "cv":
        prompt = (
            "أنشئ سيرة ذاتية احترافية ومنظمة بناءً على المعلومات التالية. "
            "لا تختلق معلومات غير مقدمة. اكتب باللغة المناسبة.\n\n" + user_text
        )
        title = "السيرة الذاتية"

    else:
        prompt = (
            "لخص الدرس التالي بطريقة واضحة وسهلة للطالب، "
            "مع أهم الأفكار والمصطلحات.\n\n" + user_text
        )
        title = "ملخص درس"

    try:
        result = await generate_text(prompt)
        file = await create_docx(title, result)

        await update.message.reply_document(
            document=file,
            filename="RAMZI_AI.docx",
            caption="✅ تم إعداد ملفك بواسطة RAMZI AI"
        )

        await update.message.reply_text(
            "اختر خدمة أخرى:",
            reply_markup=MAIN_MENU
        )

    except Exception as e:
        logging.exception(e)
        await update.message.reply_text(
            "حدث خطأ أثناء إنشاء الملف. حاول مرة أخرى."
        )

def main():
    if not BOT_TOKEN or not OPENAI_API_KEY:
        raise RuntimeError("BOT_TOKEN و OPENAI_API_KEY مطلوبان.")

    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(menu))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    app.run_polling()

if __name__ == "__main__":
    main()
