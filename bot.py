import os
import io
import json
import time
import base64
import logging
import tempfile
from pathlib import Path
from datetime import datetime, date
from typing import Optional

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.constants import ChatAction
from telegram.ext import (
    Application, CommandHandler, MessageHandler, CallbackQueryHandler,
    ContextTypes, filters
)
from openai import OpenAI
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_RIGHT, TA_CENTER
from reportlab.lib.units import cm
import fitz  # PyMuPDF
import arabic_reshaper
from bidi.algorithm import get_display

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
log = logging.getLogger("RAMZI_AI")

BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "").strip()
ADMIN_ID = os.getenv("ADMIN_ID", "").strip()
MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini").strip()
DAILY_LIMIT = int(os.getenv("DAILY_LIMIT", "5"))
MAX_FILE_MB = int(os.getenv("MAX_FILE_MB", "15"))
DATA_FILE = Path(os.getenv("DATA_FILE", "ramzi_data.json"))

if not BOT_TOKEN or not OPENAI_API_KEY:
    raise RuntimeError("يرجى ضبط BOT_TOKEN و OPENAI_API_KEY في متغيرات البيئة.")

client = OpenAI(api_key=OPENAI_API_KEY)
user_states = {}
data = {"usage": {}, "users": {}}

def load_data():
    global data
    try:
        if DATA_FILE.exists():
            loaded = json.loads(DATA_FILE.read_text(encoding="utf-8"))
            if isinstance(loaded, dict):
                data.update(loaded)
    except Exception as exc:
        log.warning("Could not load data: %s", exc)

def save_data():
    try:
        DATA_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as exc:
        log.warning("Could not save data: %s", exc)

def is_admin(user_id: int) -> bool:
    return bool(ADMIN_ID) and str(user_id) == ADMIN_ID

def register_user(user):
    uid = str(user.id)
    data["users"][uid] = {
        "name": user.full_name,
        "username": user.username or "",
        "last_seen": datetime.now().isoformat(timespec="seconds")
    }
    save_data()

def used_today(user_id: int) -> int:
    key = f"{date.today().isoformat()}:{user_id}"
    return int(data["usage"].get(key, 0))

def can_use(user_id: int) -> bool:
    return is_admin(user_id) or used_today(user_id) < DAILY_LIMIT

def record_use(user_id: int):
    key = f"{date.today().isoformat()}:{user_id}"
    data["usage"][key] = used_today(user_id) + 1
    save_data()

def rtl(text: str) -> str:
    # ReportLab requires Arabic shaping and bidi ordering.
    return get_display(arabic_reshaper.reshape(text))

def find_arabic_font() -> Optional[str]:
    candidates = [
        "fonts/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/noto/NotoNaskhArabic-Regular.ttf",
        "/usr/share/fonts/truetype/noto/NotoSansArabic-Regular.ttf",
        "/usr/share/fonts/truetype/freefont/FreeSans.ttf",
    ]
    for p in candidates:
        if Path(p).exists():
            return p
    return None

def make_docx(title: str, body: str, path: str):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title)
    r.bold = True
    r.font.size = Pt(18)

    for line in body.splitlines():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.space_after = Pt(7)
        run = p.add_run(line if line.strip() else " ")
        run.font.name = "Arial"
        run.font.size = Pt(13)
    doc.save(path)

def make_pdf(title: str, body: str, path: str):
    font_path = find_arabic_font()
    if not font_path:
        raise RuntimeError(
            "لم أجد خطاً عربياً لإنشاء PDF. أضف ملف fonts/DejaVuSans.ttf إلى المشروع ثم أعد التشغيل."
        )
    if "RAMZIArabic" not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont("RAMZIArabic", font_path))

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "ArabicTitle", parent=styles["Title"], fontName="RAMZIArabic",
        fontSize=18, leading=25, alignment=TA_CENTER, spaceAfter=18
    )
    body_style = ParagraphStyle(
        "ArabicBody", parent=styles["BodyText"], fontName="RAMZIArabic",
        fontSize=12, leading=20, alignment=TA_RIGHT, spaceAfter=8
    )
    story = [Paragraph(rtl(title), title_style), Spacer(1, 0.3 * cm)]
    for line in body.splitlines():
        safe = (line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
        story.append(Paragraph(rtl(safe) if safe.strip() else " ", body_style))
    pdf = SimpleDocTemplate(
        path, pagesize=A4, rightMargin=2*cm, leftMargin=2*cm,
        topMargin=2*cm, bottomMargin=2*cm
    )
    pdf.build(story)

def ask_ai(prompt: str, system: str = "") -> str:
    response = client.chat.completions.create(
        model=MODEL,
        messages=[
            {"role": "system", "content": system or (
                "أنت RAMZI AI، مساعد تعليمي عربي. أنشئ محتوى تعليمياً دقيقاً ومنظماً. "
                "لا تختلق المراجع أو المعلومات. إذا كانت المعلومة غير مؤكدة فاذكر ذلك. "
                "اكتب بلغة واضحة مناسبة للمستوى الدراسي، مع مقدمة وعناصر وخاتمة ومراجع مقترحة عند الحاجة."
            )},
            {"role": "user", "content": prompt}
        ],
        temperature=0.5,
    )
    return response.choices[0].message.content.strip()

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    register_user(user)
    keyboard = [
        [InlineKeyboardButton("📚 إنشاء بحث", callback_data="research")],
        [InlineKeyboardButton("📝 تلخيص درس", callback_data="summary"),
         InlineKeyboardButton("❓ أسئلة للمراجعة", callback_data="quiz")],
        [InlineKeyboardButton("🔎 استخراج نص من ملف/صورة", callback_data="extract")],
        [InlineKeyboardButton("ℹ️ المساعدة", callback_data="help")]
    ]
    await update.message.reply_text(
        f"مرحباً {user.first_name} بك في RAMZI AI 🎓\n\n"
        "مساعدك لإنشاء البحوث المدرسية، تلخيص الدروس، واستخراج النصوص.\n"
        f"الحصة المجانية: {DAILY_LIMIT} طلبات يومياً.\n\n"
        "اختر خدمة من القائمة أو أرسل عنوان بحثك مباشرة.",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📌 طريقة الاستخدام:\n"
        "• أرسل عنوان البحث مع المستوى الدراسي وعدد الصفحات.\n"
        "• أرسل صورة أو PDF لاستخراج النص.\n"
        "• استخدم /research لبدء بحث جديد.\n"
        "• استخدم /admin للمسؤول فقط.\n\n"
        "بعد إنشاء المحتوى ستظهر أزرار تنزيل Word وPDF."
    )

async def admin_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    if not is_admin(uid):
        await update.message.reply_text("هذه القائمة مخصصة للمسؤول.")
        return
    today_prefix = date.today().isoformat() + ":"
    today_uses = sum(v for k, v in data["usage"].items() if k.startswith(today_prefix))
    await update.message.reply_text(
        "🛠 لوحة تحكم RAMZI AI\n"
        f"عدد المستخدمين المسجلين: {len(data['users'])}\n"
        f"طلبات اليوم: {today_uses}\n"
        f"الحد المجاني: {DAILY_LIMIT} طلبات/يوم\n"
        f"النموذج: {MODEL}\n\n"
        "للتغيير: عدّل DAILY_LIMIT من إعدادات الاستضافة."
    )

async def set_mode(update: Update, context: ContextTypes.DEFAULT_TYPE, mode: str):
    context.user_data["mode"] = mode
    labels = {
        "research": "اكتب عنوان البحث، المستوى الدراسي، وعدد الصفحات المطلوب.",
        "summary": "أرسل نص الدرس أو صورة/ملف الدرس الذي تريد تلخيصه.",
        "quiz": "اكتب اسم الدرس والمستوى الدراسي لأعدّ لك أسئلة للمراجعة.",
        "extract": "أرسل صورة أو ملف PDF أو Word لاستخراج النص منه."
    }
    await update.effective_message.reply_text(labels[mode])

async def buttons(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    mode = q.data
    if mode == "help":
        await q.message.reply_text("أرسل عنوان البحث أو صورة/PDF. يمكنك استعمال /help لعرض التعليمات.")
        return
    await set_mode(update, context, mode)

async def research_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await set_mode(update, context, "research")

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    register_user(user)
    if not can_use(user.id):
        await update.message.reply_text(
            f"وصلت إلى الحد المجاني اليومي ({DAILY_LIMIT} طلبات). حاول غداً."
        )
        return
    text = update.message.text.strip()
    mode = context.user_data.pop("mode", "research")
    await update.message.chat.send_action(ChatAction.TYPING)
    try:
        if mode == "research":
            prompt = (
                "أنشئ بحثاً مدرسياً منظماً بناءً على طلب المستخدم التالي. "
                "استخدم عنواناً ومقدمة وعناوين فرعية وخاتمة ومراجع مقترحة. "
                "لا تضع بيانات شخصية أو اسم تلميذ من عندك.\n\n" + text
            )
            title = text.splitlines()[0][:100] or "بحث مدرسي"
            result = await __import__("asyncio").to_thread(ask_ai, prompt)
        elif mode == "summary":
            title = "ملخص درس"
            result = await __import__("asyncio").to_thread(
                ask_ai, "لخص الدرس التالي في نقاط واضحة، ثم أضف أهم المصطلحات:\n" + text
            )
        elif mode == "quiz":
            title = "أسئلة للمراجعة"
            result = await __import__("asyncio").to_thread(
                ask_ai, "أنشئ 10 أسئلة متنوعة مع الأجوبة في النهاية حول:\n" + text
            )
        else:
            title = "نتيجة المعالجة"
            result = await __import__("asyncio").to_thread(ask_ai, text)
        record_use(user.id)
        context.user_data["last_result"] = {"title": title, "body": result}
        await send_long_text(update, result)
        keyboard = [[
            InlineKeyboardButton("📄 تحميل PDF", callback_data="download_pdf"),
            InlineKeyboardButton("📝 تحميل Word", callback_data="download_docx")
        ]]
        await update.message.reply_text(
            "تم إعداد المحتوى. اختر صيغة التحميل:",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )
    except Exception as exc:
        log.exception("Text processing failed")
        await update.message.reply_text(f"حدث خطأ أثناء إعداد المحتوى: {str(exc)[:300]}")

async def send_long_text(update, text: str):
    limit = 3900
    for i in range(0, len(text), limit):
        await update.message.reply_text(text[i:i+limit])

async def extract_image(file_bytes: bytes, mime: str) -> str:
    b64 = base64.b64encode(file_bytes).decode("utf-8")
    response = client.chat.completions.create(
        model=MODEL,
        messages=[{
            "role": "user",
            "content": [
                {"type": "text", "text": "استخرج النص الظاهر في هذه الصورة كما هو، وحافظ على ترتيب الفقرات قدر الإمكان. إذا كان النص غير واضح فضع [غير واضح]."},
                {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}}
            ]
        }],
        max_tokens=3000
    )
    return response.choices[0].message.content.strip()

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    register_user(user)
    if not can_use(user.id):
        await update.message.reply_text(f"وصلت إلى الحد المجاني اليومي ({DAILY_LIMIT} طلبات). حاول غداً.")
        return
    doc = update.message.document
    if doc.file_size and doc.file_size > MAX_FILE_MB * 1024 * 1024:
        await update.message.reply_text(f"حجم الملف أكبر من الحد المسموح ({MAX_FILE_MB} MB).")
        return
    name = (doc.file_name or "file").lower()
    tgfile = await context.bot.get_file(doc.file_id)
    raw = bytes(await tgfile.download_as_bytearray())
    try:
        await update.message.chat.send_action(ChatAction.TYPING)
        if name.endswith(".pdf"):
            pdf = fitz.open(stream=raw, filetype="pdf")
            extracted = "\n\n".join(page.get_text() for page in pdf).strip()
            if not extracted:
                await update.message.reply_text("الـPDF يبدو ممسوحاً ضوئياً. أرسل الصفحات كصور لاستخراج النص، أو جرّب ملفاً يحتوي نصاً قابلاً للتحديد.")
                return
        elif name.endswith(".docx"):
            with tempfile.NamedTemporaryFile(suffix=".docx") as tmp:
                tmp.write(raw); tmp.flush()
                d = Document(tmp.name)
                extracted = "\n".join(p.text for p in d.paragraphs)
        elif name.endswith((".jpg", ".jpeg", ".png", ".webp")):
            mime = "image/png" if name.endswith(".png") else "image/jpeg"
            extracted = await __import__("asyncio").to_thread(
                lambda: client.chat.completions.create(
                    model=MODEL,
                    messages=[{"role":"user","content":[
                        {"type":"text","text":"استخرج النص من الصورة كما هو، مع الحفاظ على الفقرات."},
                        {"type":"image_url","image_url":{"url":f"data:{mime};base64,{base64.b64encode(raw).decode()}"}}
                    ]}],
                    max_tokens=3000
                ).choices[0].message.content.strip()
            )
        else:
            await update.message.reply_text("الملفات المدعومة: PDF وWord وصور JPG وPNG وWEBP.")
            return
        record_use(user.id)
        context.user_data["last_result"] = {"title": "النص المستخرج", "body": extracted}
        await send_long_text(update, extracted or "لم أتمكن من العثور على نص داخل الملف.")
        keyboard = [[
            InlineKeyboardButton("📄 حفظ PDF", callback_data="download_pdf"),
            InlineKeyboardButton("📝 حفظ Word", callback_data="download_docx")
        ]]
        await update.message.reply_text("يمكنك حفظ النص المستخرج:", reply_markup=InlineKeyboardMarkup(keyboard))
    except Exception as exc:
        log.exception("Document processing failed")
        await update.message.reply_text(f"تعذر معالجة الملف: {str(exc)[:250]}")

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    register_user(user)
    if not can_use(user.id):
        await update.message.reply_text(f"وصلت إلى الحد المجاني اليومي ({DAILY_LIMIT} طلبات). حاول غداً.")
        return
    photo = update.message.photo[-1]
    tgfile = await context.bot.get_file(photo.file_id)
    raw = bytes(await tgfile.download_as_bytearray())
    try:
        await update.message.chat.send_action(ChatAction.TYPING)
        result = await __import__("asyncio").to_thread(
            lambda: client.chat.completions.create(
                model=MODEL,
                messages=[{"role":"user","content":[
                    {"type":"text","text":"استخرج النص من الصورة كما هو، مع الحفاظ على ترتيب الأسطر والفقرات."},
                    {"type":"image_url","image_url":{"url":f"data:image/jpeg;base64,{base64.b64encode(raw).decode()}"}}
                ]}],
                max_tokens=3000
            ).choices[0].message.content.strip()
        )
        record_use(user.id)
        context.user_data["last_result"] = {"title": "النص المستخرج", "body": result}
        await send_long_text(update, result)
        keyboard = [[
            InlineKeyboardButton("📄 حفظ PDF", callback_data="download_pdf"),
            InlineKeyboardButton("📝 حفظ Word", callback_data="download_docx")
        ]]
        await update.message.reply_text("حفظ النص في ملف:", reply_markup=InlineKeyboardMarkup(keyboard))
    except Exception as exc:
        log.exception("Photo processing failed")
        await update.message.reply_text(f"تعذر استخراج النص: {str(exc)[:250]}")

async def download_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    result = context.user_data.get("last_result")
    if not result:
        await q.message.reply_text("لا يوجد محتوى محفوظ. أنشئ بحثاً أو استخرج نصاً أولاً.")
        return
    title, body = result["title"], result["body"]
    suffix = ".pdf" if q.data == "download_pdf" else ".docx"
    try:
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            path = tmp.name
        if suffix == ".pdf":
            make_pdf(title, body, path)
        else:
            make_docx(title, body, path)
        with open(path, "rb") as f:
            await q.message.reply_document(document=f, filename=f"RAMZI_AI{suffix}", caption="تم إنشاء الملف بواسطة RAMZI AI")
        os.unlink(path)
    except Exception as exc:
        log.exception("File generation failed")
        await q.message.reply_text(f"تعذر إنشاء الملف: {str(exc)[:250]}")

def main():
    load_data()
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler("research", research_cmd))
    app.add_handler(CommandHandler("admin", admin_cmd))
    app.add_handler(CallbackQueryHandler(download_file, pattern="^download_(pdf|docx)$"))
    app.add_handler(CallbackQueryHandler(buttons))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    log.info("RAMZI AI is starting...")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
