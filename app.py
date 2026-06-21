import logging
import json
import os
import threading
import io
import time
from datetime import datetime
from http.server import SimpleHTTPRequestHandler, HTTPServer
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
import google.generativeai as genai
from google.api_core import exceptions
from docxtpl import DocxTemplate

# --- הגדרות ---
TELEGRAM_TOKEN = "8810122605:AAFkA97_VY3KV172CFf-7BleyDhMQgj4yYM"
MY_CHAT_ID = 8251059616 

# משיכת המפתח מתוך ה-Environment של Render
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
TEMPLATE_PATH = "template.docx"

genai.configure(api_key=GEMINI_API_KEY)
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)

FIELD_LABELS = {
    "project_num": "מספר פרויקט", "letter_num": "מספר מכתב", "client_name": "לכבוד (שם הלקוח)",
    "contact_person": "לידי (איש קשר)", "client_email": "במייל (כתובת אימייל)",
    "structure_name": "שם המבנה / הפרויקט", "inspection_subject": "במהלך הסיור בוצע פיקוח ל...",
    "star_present": "נוכח מטעם סטאר מהנדסים", "inspector_name": "שם המפקח באתר",
    "execution_team": "נציגי הביצוע / קבלן", "author_initials": "ראשי תיבות של כותב הדוח",
    "work_status": "תיאור מצב העבודה הנוכחי"
}

def get_ai_response(prompt):
    """פונקציה המשתמשת ב-Gemini 2.0 עם מנגנון המתנה חכם"""
    model = genai.GenerativeModel('gemini-2.0-flash')
    retries = 3
    for i in range(retries):
        try:
            return model.generate_content(prompt)
        except exceptions.ResourceExhausted:
            if i < retries - 1:
                logging.warning(f"Quota exceeded для Gemini 2.0, retrying in {20 * (i + 1)} seconds...")
                time.sleep(20 * (i + 1))
                continue
            raise
    return None

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
        'project_num': data.get("project_num", ""), 'letter_num': data.get("letter_num", ""),
        'client_name': data.get("client_name", ""), 'contact_person': data.get("contact_person", ""),
        'client_email': data.get("client_email", ""), 'structure_name': data.get("structure_name", ""),
        'visit_date': current_date, 'inspection_subject': data.get("inspection_subject", ""),
        'star_present': rlm + data.get('star_present', 'הח"מ') + rlm,
        'inspector_name': data.get("inspector_name", ""), 'execution_team': data.get("execution_team", ""),
        'author_initials': data.get("author_initials", "A.K"),
        'work_status': "\n".join([f"{rlm}{line}{rlm}" for line in data.get("work_status", "").split("\n")]),
        'signature_image': None, 'specific_remarks_list': [],
        'general_remarks_list': gen_remarks, 'cc_final_list': cc_final
    }
    doc.render(context)
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # הדפסה ללוג כדי לוודא את ה-Chat ID שלך במקרה של חוסר תגובה
    logging.info(f"Received message from Chat ID: {update.effective_chat.id}")
    
    if update.effective_chat.id != MY_CHAT_ID: 
        return
    
    current_state = context.user_data.get('state', 'IDLE')
    
    if current_state == 'IDLE':
        await update.message.reply_text("🔄 מנתח נתונים מהשטח עם Gemini 2.0, נא להמתין...")
        try:
            prompt = f"אתה עוזר של סטאר מהנדסים. חלץ מידע מתוך הטקסט הבא ל-JSON:\n{update.message.text}"
            response = get_ai_response(prompt)
            data = json.loads(response.text.replace("```json", "").replace("```", "").strip())
            context.user_data['report_data'] = data
            missing = [k for k, v in data.items() if not v or not str(v).strip()]
            
            if missing:
                context.user_data['state'] = 'ASKING_FIELDS'
                context.user_data['missing_fields'] = missing
                await update.message.reply_text(f"⚠️ חסרים פרטים. אנא הזן: *{FIELD_LABELS[missing[0]]}*")
            else:
                await send_finished_report(update, context)
        except Exception as e:
            await update.message.reply_text(f"❌ שגיאה בניתוח: {e}")
    
    elif current_state == 'ASKING_FIELDS':
        missing = context.user_data.get('missing_fields', [])
        field = missing.pop(0)
        context.user_data['report_data'][field] = update.message.text
        if missing:
            context.user_data['missing_fields'] = missing
            await update.message.reply_text(f"❓ הבא: *{FIELD_LABELS[missing[0]]}*")
        else:
            await send_finished_report(update, context)

async def send_finished_report(update, context):
    await update.message.reply_text("🎉 כל הפרטים מלאים! מפיק את הדוח...")
    bio = generate_report_bytes(context.user_data['report_data'])
    await update.message.reply_document(document=bio, filename="Report.docx")
    context.user_data.clear()
    context.user_data['state'] = 'IDLE'

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    logging.info(f"Start command triggered by Chat ID: {update.effective_chat.id}")
    if update.effective_chat.id != MY_CHAT_ID: 
        return
    context.user_data['state'] = 'IDLE'
    await update.message.reply_text("🏗️ בוט פיקוח עליון (Gemini 2.0) - סטאר מהנדסים\nמוכן! שלח לי הודעה מהשטח ואני אבנה את הדוח.")

def main():
    threading.Thread(target=lambda: HTTPServer(('0.0.0.0', int(os.environ.get("PORT", 10000))), SimpleHTTPRequestHandler).serve_forever(), daemon=True).start()
    
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    app.run_polling()

if __name__ == '__main__': main()