import logging
import json
import os
from datetime import datetime
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
import google.generativeai as genai
from docx import Document

# --- מפתחות קבועים ומאובטחים ---
TELEGRAM_TOKEN = "8810122605:AAFkA97_VY3KV172CFf-7BleyDhMQgj4yYM"
MY_CHAT_ID = 8251059616 

# --- משיכת מפתח ה-AI מהגדרות השרת ב-Render ---
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
TEMPLATE_PATH = "template.docx"  # ודא שקובץ הוורד שלך ב-GitHub נקרא בדיוק כך!

# חיבור ל-Gemini
genai.configure(api_key=GEMINI_API_KEY)

logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)

PROMPT_INSTRUCTIONS = """
אתה עוזר משרדי עבור חברת קונסטרוקציות בשם "סטאר מהנדסים". תפקידך הוא לקחת טקסט חופשי שנכתב או תומלל על ידי מהנדס בשטח, ולשלוף ממנו את המידע עבור טופס פיקוח עליון.
עליך להחזיר פלט בפורמט JSON בלבד, ללא שום טקסט נוסף לפני או אחרי, לפי המבנה הבא:
{
  "project_num": "מספר פרויקט אם הוזכר, אחרת תשאיר ריק",
  "letter_num": "מספר מכתב/דוח אם הוזכר, אחרת תשאיר ריק",
  "client_name": "שם הלקוח לכבודו נכתב הדוח",
  "contact_person": "לידי מי המכתב מיועד",
  "client_email": "כתובת מייל אם הוזכרה",
  "structure_name": "שם המבנה או הפרויקט",
  "inspection_subject": "מה בוצע במהלך הסיור (למשל: בדיקת זיון תקרת קומה א)",
  "star_present": "מי נוכח מטעם סטאר מהנדסים (למשל: הח\"מ)",
  "inspector_name": "שם המפקח באתר",
  "execution_team": "נציגי הביצוע/קבלן",
  "author_initials": "ראשי תיבות של המהנדס",
  "work_status": "תיאור קצר ומקצועי של מצב העבודה הנוכחי באתר לפי דברי המהנדס"
}
השתמש בשפה מקצועית של מהנדסי בניין. אם פרט מסוים לא הוזכר בטקסט, השאר אותו כריק ("").
"""

def create_word_report(data, template_path, output_path):
    """פונקציה שמחליפה את התגיות במסמך הוורד בנתונים שהתקבלו מה-AI"""
    doc = Document(template_path)
    current_date = datetime.now().strftime("%d/%m/%Y")
    
    replacements = {
        "{project_num}": data.get("project_num", ""),
        "{letter_num}": data.get("letter_num", ""),
        "{report_date}": current_date,
        "{client_name}": data.get("client_name", ""),
        "{contact_person}": data.get("contact_person", ""),
        "{client_email}": data.get("client_email", ""),
        "{structure_name}": data.get("structure_name", ""),
        "{visit_date}": current_date,
        "{inspection_subject}": data.get("inspection_subject", ""),
        "{star_present}": data.get("star_present", ""),
        "{inspector_name}": data.get("inspector_name", ""),
        "{execution_team}": data.get("execution_team", ""),
        "{author_initials}": data.get("author_initials", ""),
        "{work_status}": data.get("work_status", "")
    }
    
    # החלפה בפסקאות רגילות
    for paragraph in doc.paragraphs:
        for placeholder, value in replacements.items():
            if placeholder in paragraph.text:
                paragraph.text = paragraph.text.replace(placeholder, value)
                
    # החלפה בתוך טבלאות
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for placeholder, value in replacements.items():
                        if placeholder in paragraph.text:
                            paragraph.text = paragraph.text.replace(placeholder, value)
                            
    doc.save(output_path)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_chat.id != MY_CHAT_ID:
        return
    await update.message.reply_text("🏗️ בוט פיקוח עליון ישיר של סטאר מהנדסים פעיל!\nשלח לי הודעה חופשית (טקסט או קול) ותקבל קובץ Word מוכן חזרה.")

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_chat.id != MY_CHAT_ID:
        return

    user_text = update.message.text
    await update.message.reply_text("🔄 ה-AI מפענח את הנתונים ומייצר את קובץ ה-Word, אנא המתן...")

    output_filename = f"פיקוח_עליון_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"

    try:
        # פנייה ל-Gemini
        model = genai.GenerativeModel('gemini-1.5-flash')
        response = model.generate_content(f"{PROMPT_INSTRUCTIONS}\n\nהטקסט מהשטח:\n{user_text}")
        
        clean_json = response.text.replace("```json", "").replace("```", "").strip()
        data = json.loads(clean_json)
        
        # יצירת הקובץ מהתבנית
        create_word_report(data, TEMPLATE_PATH, output_filename)
        
        # שליחת הקובץ המוכן חזרה לטלגרם
        with open(output_filename, 'rb') as doc_file:
            await update.message.reply_document(document=doc_file, filename=output_filename, caption="📝 הנה דוח הפיקוח העליון המוכן שלך!")
            
        os.remove(output_filename)
        
    except Exception as e:
        await update.message.reply_text(f"❌ התרחשה שגיאה בהפקת המסמך: {e}")

def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    print("Bot is running...")
    app.run_polling()

if __name__ == '__main__':
    main()