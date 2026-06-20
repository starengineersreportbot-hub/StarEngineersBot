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

# --- הגדרות קבועות ומאובטחות ---
TELEGRAM_TOKEN = "8810122605:AAFkA97_VY3KV172CFf-7BleyDhMQgj4yYM"
MY_CHAT_ID = 8251059616 

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
TEMPLATE_PATH = "template.docx"

# חיבור ל-Gemini
genai.configure(api_key=GEMINI_API_KEY)

logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)

# מיפוי השדות לעברית בשביל השאלות של הבוט
FIELD_LABELS = {
    "project_num": "מספר פרויקט",
    "letter_num": "מספר מכתב",
    "client_name": "לכבוד (שם הלקוח)",
    "contact_person": "לידי (איש קשר / מנהל הפרויקט)",
    "client_email": "במייל (כתובת אימייל הלקוח)",
    "structure_name": "שם המבנה / הפרויקט",
    "inspection_subject": "במהלך הסיור בוצע פיקוח ל...",
    "star_present": "נוכח מטעם סטאר מהנדסים",
    "inspector_name": "שם המפקח באתר",
    "execution_team": "נציגי הביצוע / קבלן",
    "author_initials": "ראשי תיבות של כותב הדוח",
    "work_status": "תיאור מצב העבודה הנוכחי באתר"
}

PROMPT_INSTRUCTIONS = """
אתה עוזר משרדי עבור חברת קונסטרוקציות בשם "סטאר מהנדסים". תפקידך הוא לקחת טקסט חופשי שנכתב או תומלל על ידי מהנדס בשטח, ולשלוף ממנו את המידע עבור טופס פיקוח עליון.
עליך להחזיר פלט בפורמט JSON בלבד, ללא שום טקסט נוסף לפני או אחרי, לפי המבנה הבא:
{
  "project_num": "מספר פרויקט אם הוזכר, אחרת תשאיר ריק",
  "letter_num": "מספר מכתב/דוח אם הוזכר, אחרת תשאיר ריק",
  "client_name": "שם הלקוח לכבודו נכתב הדוח, אחרת תשאיר ריק",
  "contact_person": "לידי מי המכתב מיועד, אחרת תשאיר ריק",
  "client_email": "כתובת מייל אם הוזכרה, אחרת תשאיר ריק",
  "structure_name": "שם המבנה או הפרויקט, אחרת תשאיר ריק",
  "inspection_subject": "מה בוצע במהלך הסיור (למשל: בדיקת זיון תקרת קומה א), אחרת תשאיר ריק",
  "star_present": "מי נוכח מטעם סטאר מהנדסים (למשל: הח\"מ), אחרת תשאיר ריק",
  "inspector_name": "שם המפקח באתר, אחרת תשאיר ריק",
  "execution_team": "נציגי הביצוע/קבלן, אחרת תשאיר ריק",
  "author_initials": "ראשי תיבות של המהנדס, אחרת תשאיר ריק",
  "work_status": "תיאור קצר ומקצועי של מצב העבודה הנוכחי באתר לפי דברי המהנדס, אחרת תשאיר ריק"
}
השתמש בשפה מקצועית של מהנדסי בניין. אם פרט מסוים לא הוזכר בטקסט בשום צורה, השאר אותו כריק ("").
"""

def generate_report_bytes(data):
    """פונקציה שמייצרת את קובץ ה-Word בדיוק לפי הלוגיקה של ה-Streamlit שלך"""
    doc = DocxTemplate(TEMPLATE_PATH)
    rlm = "\u200f" # סימן כיווניות מימין לשמאל
    current_date = datetime.now().strftime("%d/%m/%Y")
    
    # הגדרות ברירת מחדל להערות והעתקים (כמו ב-Streamlit)
    txts = ["יש להסיר שאריות בטון ישן מתחתית ברזלי הזיון.", "יש לדאוג טרם היציקה שמשטח היציקה נקי.", "יש לשמור על עובי כיסוי עפ\"י התכניות.", "ניתן להמשיך עבודות לאחר אישור המפקח.", "לשאלות נוספות ניתן לפנות אלינו."]
    # בבוט כרגע נכניס את ההערות הכלליות כריקות או מלאות כברירת מחדל - ניתן להרחיב בהמשך
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
        'star_present': f"{rlm}{data.get('star_present', 'הח\"מ')}{rlm}",
        'inspector_name': data.get("inspector_name", ""),
        'execution_team': data.get("execution_team", ""),
        'author_initials': data.get("author_initials", "A.K"),
        'work_status': "\n".join([f"{rlm}{line}{rlm}" for line in data.get("work_status", "").split("\n")]),
        'signature_image': None,
        'specific_remarks_list': [], # יורחב בהמשך עם שליחת תמונות
        'general_remarks_list': gen_remarks,
        'cc_final_list': cc_final
    }
    
    doc.render(context)
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_chat.id != MY_CHAT_ID:
        return
    # איפוס הזיכרון של השיחה
    context.user_data.clear()
    context.user_data['state'] = 'IDLE'
    await update.message.reply_text("🏗️ בוט פיקוח עליון אינטראקטיבי של סטאר מהנדסים מוכן!\n\nשלח לי את הודעת השטח החופשית שלך (בטקסט או בהקלטה), ואני אדאג לחלץ את הנתונים ולשאול אותך על מה שפספסת.")

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_chat.id != MY_CHAT_ID:
        return

    current_state = context.user_data.get('state', 'IDLE')
    user_input = update.message.text

    # --- מצב 1: הודעה ראשונית מהשטח ---
    if current_state == 'IDLE':
        await update.message.reply_text("🔄 ה-AI מנתח את הדיווח מהשטח, אנא המתן...")
        
        try:
            model = genai.GenerativeModel('gemini-2-flash')
            response = model.generate_content(f"{PROMPT_INSTRUCTIONS}\n\nהטקסט מהשטח:\n{user_input}")
            
            clean_json = response.text.replace("```json", "").replace("```", "").strip()
            extracted_data = json.loads(clean_json)
            
            # שמירת הנתונים שמצאנו בזיכרון של המשתמש
            context.user_data['report_data'] = extracted_data
            
            # בדיקה מה חסר
            missing = [key for key, val in extracted_data.items() if not val.strip()]
            
            if missing:
                context.user_data['state'] = 'ASKING_FIELDS'
                context.user_data['missing_fields'] = missing
                
                # מציגים למשתמש מה כן נקלט
                found_summary = "\n".join([f"🔹 *{FIELD_LABELS[k]}:* {v}" for k, v in extracted_data.items() if v.strip()])
                welcome_back = "✅ *הנתונים הבאים נקלטו בהצלחה:*\n" + found_summary if found_summary else " לא הצלחתי לחלץ נתונים מזהים."
                
                await update.message.reply_text(f"{welcome_back}\n\n⚠️ *חסרים לנו פרטים חובה לדוח.*")
                
                # שואלים את השאלה הראשונה
                next_field = missing[0]
                await update.message.reply_text(f"❓ אנא הזן: *{FIELD_LABELS[next_field]}*")
            else:
                # הכל נמצא כבר במכה הראשונה!
                await send_finished_report(update, context)
                
        except Exception as e:
            await update.message.reply_text(f"❌ שגיאה בפענוח הנתונים: {e}\nנסה לשלוח שוב את הודעת השטח.")

    # --- מצב 2: השלמת שדות חסרים בפינג-פונג ---
    elif current_state == 'ASKING_FIELDS':
        missing_fields = context.user_data.get('missing_fields', [])
        if not missing_fields:
            context.user_data['state'] = 'IDLE'
            return
            
        # שמירת התשובה לשדה הנוכחי
        current_field = missing_fields.pop(0)
        context.user_data['report_data'][current_field] = user_input
        
        # אם נשארו עוד שדות חסרים, שואלים את הבא בתור
        if missing_fields:
            context.user_data['missing_fields'] = missing_fields
            next_field = missing_fields[0]
            await update.message.reply_text(f"👍 נקלט. \n❓ אנא הזן: *{FIELD_LABELS[next_field]}*")
        else:
            # סיימנו את כל השאלות! מפיקים את הדוח
            await send_finished_report(update, context)

async def send_finished_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("🎉 כל שדות החובה מלאים! מייצר את קובץ ה-Word...")
    try:
        data = context.user_data['report_data']
        project_num = data.get("project_num", "REPORT")
        letter_num = data.get("letter_num", "01")
        
        # יצירת הקובץ בזיכרון השרת
        report_stream = generate_report_bytes(data)
        output_filename = f"{project_num}-{letter_num}.docx"
        
        # שליחה למשתמש
        await update.message.reply_document(
            document=report_stream, 
            filename=output_filename, 
            caption="📝 הנה דוח הפיקוח העליון הרשמי שלך מוכן ומעוצב!"
        )
        
        # איפוס הבוט לסבב הבא
        context.user_data.clear()
        context.user_data['state'] = 'IDLE'
        await update.message.reply_text("✨ הבוט מאופס ומוכן לדוח הבא שלך!")
        
    except Exception as e:
        await update.message.reply_text(f"❌ שגיאה בהפקת קובץ הוורד: {e}")
        context.user_data['state'] = 'IDLE'

def run_dummy_server():
    port = int(os.environ.get("PORT", 10000))
    server = HTTPServer(('0.0.0.0', port), SimpleHTTPRequestHandler)
    print(f"Dummy server running on port {port}...")
    server.serve_forever()

def main():
    threading.Thread(target=run_dummy_server, daemon=True).start()
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    print("Bot is running with full dynamic question flow...")
    app.run_polling()

if __name__ == '__main__':
    main()
