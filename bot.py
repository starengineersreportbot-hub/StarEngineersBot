import logging
import json
import os
import threading
import io
from datetime import datetime
from http.server import SimpleHTTPRequestHandler, HTTPServer
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
import google.generativeai as genai
from docx.shared import Inches
from docxtpl import DocxTemplate

# --- הגדרות ---
TELEGRAM_TOKEN = "8810122605:AAFkA97_VY3KV172CFf-7BleyDhMQgj4yYM"
MY_CHAT_ID = 8251059616 
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
TEMPLATE_PATH = "template.docx"

genai.configure(api_key=GEMINI_API_KEY)
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)

FIELD_LABELS = {
    "project_num": "מספר פרויקט",
    "letter_num": "מספר מכתב",
    "client_name": "לכבוד (שם הלקוח)",
    "contact_person": "לידי (איש קשר)",
    "client_email": "במייל (כתובת אימייל)",
    "structure_name": "שם המבנה / הפרויקט",
    "inspection_subject": "במהלך הסיור בוצע פיקוח ל...",
    "star_present": "נוכח מטעם סטאר מהנדסים",
    "inspector_name": "שם המפקח באתר",
    "execution_team": "נציגי הביצוע / קבלן",
    "author_initials": "ראשי תיבות של כותב הדוח",
    "work_status": "תיאור מצב העבודה הנוכחי"
}

PROMPT_INSTRUCTIONS = """
אתה עוזר משרדי של "סטאר מהנדסים". חלץ מידע מתוך טקסט מהשטח עבור טופס פיקוח עליון והחזר JSON בלבד:
{
  "project_num": "", "letter_num": "", "client_name": "", "contact_person": "", "client_email": "",
  "structure_name": "", "inspection_subject": "", "star_present": "", "inspector_name": "",
  "execution_team": "", "author_initials": "", "work_status": ""
}
אם פרט חסר, השאר את הערך ריק.
"""

def generate_report_bytes(data):
    doc = DocxTemplate(TEMPLATE_PATH)
    rlm = "\u200f"
    current_date = datetime.now().strftime("%d/%m/%Y")
    
    txts = ["יש להסיר שאריות בטון ישן מתחתית ברזלי הזיון.", "יש לדאוג טרם היציקה שמשטח היציקה נקי.", "יש לשמור על עובי כיסוי עפ\"י התכניות.", "ניתן להמשיך עבודות לאחר אישור המפקח.", "לשאלות נוספות ניתן לפנות אלינו."]
    gen_remarks = [f"{rlm}5.{i+1}.{rlm} {t}{rlm}" for i, t in enumerate(txts)]
    cc_default = "1. בוריס בקלמן/ישראל קנר – סטאר מהנדסים\n2. ראש צוות - סטאר מהנדסים\n3. מנהל פרויקט\n4. תיק פרויקט\n5. תיק כללי"
    cc_final = [f"{rlm}{line.strip()}{rlm}" for line in cc_default.split('\n') if line.strip()]

    context = {
        'report_date': current_date,
        'project_num': data.get("project_num", ""),
        'letter_num': data.get("letter_num", ""),
        'client_name': data.get("client_name", ""),
        'contact_person': data.get("contact_person", ""),
        'client_email': data.get("client_email", ""),
        'structure_name': data.get("structure_name", ""),
        'visit_date': current_date,
        'inspection_subject': data.get("inspection_subject", ""),
        'star_present': rlm + data.get('star_present', 'הח"מ') + rlm,
        'inspector_name': data.get("inspector_name", ""),
        'execution_team': data.get("execution_team", ""),
        'author_initials': data.get("author_initials", "A.K"),
        'work_status': "\n".join([f"{rlm}{line}{rlm}" for line in data.get("work_status", "").split("\n")]),
        'signature_image': None,
        'specific_remarks_list': [],
        'general_remarks_list': gen_remarks,
        'cc_final_list': cc_final
    }
    doc.render(context)
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_chat.id != MY_CHAT_ID: return
    context.user_data.clear()
    context.user_data['state'] = 'IDLE'
    await update.message.reply_text("🏗️ בוט פיקוח עליון - סטאר מהנדסים\nשלח הודעה מהשטח ואני אבנה את הדוח.")

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_chat.id != MY_CHAT_ID: return
    
    current_state = context.user_data.get('state', 'IDLE')
    user_input = update.message.text

    if current_state == 'IDLE':
        await update.message.reply_text("🔄 מנתח נתונים...")
        model = genai.GenerativeModel('gemini-2.0-flash')
        response = model.generate_content(f"{PROMPT_INSTRUCTIONS}\n\n{user_input}")
        
        extracted_data = json.loads(response.text.replace("```json", "").replace("```", "").strip())
        context.user_data['report_data'] = extracted_data
        missing = [key for key, val in extracted_data.items() if not val.strip()]
        
        if missing:
            context.user_data['state'] = 'ASKING_FIELDS'
            context.user_data['missing_fields'] = missing
            await update.message.reply_text(f"⚠️ חסרים פרטים. אנא הזן: *{FIELD_LABELS[missing[0]]}*")
        else:
            await send_finished_report(update, context)

    elif current_state == 'ASKING_FIELDS':
        missing_fields = context.user_data.get('missing_fields', [])
        current_field = missing_fields.pop(0)
        context.user_data['report_data'][current_field] = user_input
        
        if missing_fields:
            context.user_data['missing_fields'] = missing_fields
            await update.message.reply_text(f"❓ הבא בתור: *{FIELD_LABELS[missing_fields[0]]}*")
        else:
            await send_finished_report(update, context)

async def send_finished_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("🎉 כל הפרטים מלאים! מפיק את הדוח...")
    data = context.user_data['report_data']
    bio = generate_report_bytes(data)
    await update.message.reply_document(document=bio, filename=f"{data.get('project_num','REPORT')}.docx")
    context.user_data.clear()
    context.user_data['state'] = 'IDLE'

def main():
    threading.Thread(target=lambda: HTTPServer(('0.0.0.0', int(os.environ.get("PORT", 10000))), SimpleHTTPRequestHandler).serve_forever(), daemon=True).start()
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    app.run_polling()

if __name__ == '__main__':
    main()