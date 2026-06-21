import os
import telebot
from openai import OpenAI
from docxtpl import DocxTemplate
from datetime import datetime
import json

# טעינת משתני הסביבה
TOKEN = os.environ.get('TELEGRAM_TOKEN')
OPENAI_API_KEY = os.environ.get('OPENAI_API_KEY')

bot = telebot.TeleBot(TOKEN)
client = OpenAI(api_key=OPENAI_API_KEY)

# מילון לשמירת מצב השיחה של המשתמשים
user_states = {}

def get_missing_fields_prompt(text):
    """שולח את הטקסט ל-OpenAI כדי לבדוק אם חסרים פרטים לפי התבנית המדויקת"""
    prompt = f"""
    אתה עוזר מקצועי של חברת "סטאר מהנדסים" שתפקידו לעבד נתונים מהשטח ולהכין דוח פיקוח עליון מובנה.
    נתח את הטקסט הבא ששלח המהנדס מהשטח וזהה איזה מהפרטים הבאים חסרים:
    1. project_num (מספר פרויקט)
    2. letter_num (מספר מכתב)
    3. client_name (שם הלקוח)
    4. contact_person (איש קשר)
    5. client_email (כתובת אימייל)
    6. structure_name (שם המבנה)
    7. inspection_subject (נושא הפיקוח, למשל: "יציקת רצפה", "זיון קירות")
    8. inspector_name (שם המפקח באתר)
    9. execution_team (נציגי הביצוע / קבלן)
    10. work_status (תיאור מצב העבודה הנוכחי מהשטח)

    שים לב לחוקים הבאים:
    - star_present (נציג סטאר מהנדסים) הוא תמיד "הח"מ" כברירת מחדל, לכן אל תסמן אותו כחסר.
    - אל תמציא פרטים! אם פרט לא מופיע, הוא חסר.

    החזר אך ורק תשובת JSON תקינה בפורמט הבא, ללא שום טקסט נוסף לפני או אחרי:
    {{
        "extracted_data": {{
            "project_num": "הערך שנמצא או null",
            "letter_num": "הערך שנמצא או null",
            "client_name": "הערך שנמצא או null",
            "contact_person": "הערך שנמצא או null",
            "client_email": "הערך שנמצא או null",
            "structure_name": "הערך שנמצא או null",
            "inspection_subject": "הערך שנמצא או null",
            "inspector_name": "הערך שנמצא או null",
            "execution_team": "הערך שנמצא או null",
            "work_status": "הערך שנמצא או null"
        }}
    }}

    הטקסט מהשטח:
    "{text}"
    """
    
    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"},
            temperature=0
        )
        return json.loads(response.choices[0].message.content)
    except Exception as e:
        print(f"Error calling OpenAI: {e}")
        return None

def improve_description(text_to_improve):
    """משתמש ב-OpenAI כדי לנסח את תיאור המצב בשפה הנדסית גבוהה"""
    if not text_to_improve:
        return ""
    prompt = f"""
    אתה מהנדס מבנים בכיר ב"סטאר מהנדסים". נסח מחדש את תיאור מצב העבודה הבא שנכתב כהערות קצרות מהשטח, 
    והפוך אותו לפסקה מנוסחת היטב בשפה הנדסית רשמית, מקצועית ומדויקת המתאימה לדוח פיקוח עליון.
    
    הנתונים הגולמיים מהשטח: "{text_to_improve}"
    """
    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3
        )
        return response.choices[0].message.content.strip()
    except:
        return text_to_improve

def create_report(data, template_path="template.docx", output_path="output.docx"):
    """משבץ את הנתונים לתוך תבנית הוורד באמצעות docxtpl"""
    doc = DocxTemplate(template_path)
    
    # שיפור הניסוח של התיאור הקיים ב-work_status
    professional_status = improve_description(data.get('work_status', ''))
    
    # בניית ההקשר (Context) שיוזרק ישירות לסוגריים המסולסלים בוורד
    context = {
        "report_date": datetime.now().strftime("%d/%m/%Y"),
        "visit_date": datetime.now().strftime("%d/%m/%Y"), # כברירת מחדל תאריך הסיור הוא היום
        "project_num": str(data.get("project_num") or ""),
        "letter_num": str(data.get("letter_num") or ""),
        "client_name": str(data.get("client_name") or ""),
        "contact_person": str(data.get("contact_person") or ""),
        "client_email": str(data.get("client_email") or ""),
        "structure_name": str(data.get("structure_name") or ""),
        "inspection_subject": str(data.get("inspection_subject") or ""),
        "inspector_name": str(data.get("inspector_name") or ""),
        "execution_team": str(data.get("execution_team") or ""),
        "star_present": "הח\"מ",
        "work_status": professional_status,
        "specific_remarks_list": [], # כרגע ריק כדי לא לשבור את הלולאות בתבניתของคุณ
        "general_remarks_list": [],
        "cc_final_list": [],
        "signature_image": ""
    }
    
    doc.render(context)
    doc.save(output_path)
    return output_path

@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    bot.reply_to(message, "שלום אביב! שלח לי את הנתונים מהסיור בשטח, ואני אפיק דוח פיקוח עליון מותאם לתבנית של סטאר מהנדסים.")

@bot.message_handler(func=lambda message: True)
def handle_text(message):
    chat_id = message.chat.id
    text = message.text

    # תהליך השלמת פרטים חסרים
    if chat_id in user_states and user_states[chat_id].get("waiting_for"):
        state = user_states[chat_id]
        missing_field_key = state["waiting_for"]
        
        state["data"][missing_field_key] = text
        state["missing_list"].pop(0)
        
        if state["missing_list"]:
            next_field = state["missing_list"][0]
            prompts_map = {
                "project_num": "מספר פרויקט", "letter_num": "מספר מכתב",
                "client_name": "שם הלקוח (לכבוד)", "contact_person": "איש קשר (לידי)",
                "client_email": "כתובת אימייל (במייל)", "structure_name": "שם המבנה",
                "inspection_subject": "נושא הפיקוח (למשל: יציקת רצפה)",
                "inspector_name": "שם המפקח באתר", "execution_team": "נציגי הביצוע / קבלן"
            }
            field_display = prompts_map.get(next_field, next_field)
            state["waiting_for"] = next_field
            bot.send_message(chat_id, f"מעולה. עכשיו, מהו **{field_display}**?")
            return
        else:
            bot.send_message(chat_id, "כל הפרטים נאספו בהצלחה! מייצר עבורך את קובץ הוורד המעודכן...")
            try:
                doc_file = create_report(state["data"])
                with open(doc_file, 'rb') as f:
                    bot.send_document(chat_id, f, caption="הנה דוח הפיקוח העליון המוכן שלך! 📝")
                user_states.pop(chat_id, None)
            except Exception as e:
                bot.send_message(chat_id, f"אירעה שגיאה ביצירת הקובץ: {e}")
            return

    # תהליך התחלתי - ניתוח ראשוני
    bot.send_message(chat_id, "מנתח את הנתונים ששלחת ב-OpenAI, רק רגע...")
    analysis = get_missing_fields_prompt(text)
    
    if not analysis:
        bot.send_message(chat_id, "היה קושי קטן בתקשורת, אנא נסה לשלוח שוב.")
        return
        
    extracted_data = analysis.get("extracted_data", {})
    actual_missing = [k for k, v in extracted_data.items() if v is None or v == "null" or v == "None"]
    
    if actual_missing:
        user_states[chat_id] = {
            "data": extracted_data,
            "missing_list": actual_missing,
            "waiting_for": actual_missing[0]
        }
        
        prompts_map = {
            "project_num": "מספר פרויקט", "letter_num": "מספר מכתב",
            "client_name": "שם הלקוח (לכבוד)", "contact_person": "איש קשר (לידי)",
            "client_email": "כתובת אימייל (במייל)", "structure_name": "שם המבנה",
            "inspection_subject": "נושא הפיקוח (למשל: יציקת רצפה)",
            "inspector_name": "שם המפקח באתר", "execution_team": "נציגי הביצוע / קבלן",
            "work_status": "תיאור מצב העבודה הנוכחי"
        }
        
        first_missing = actual_missing[0]
        field_display = prompts_map.get(first_missing, first_missing)
        bot.send_message(chat_id, f"הבנתי חלק מהפרטים, מהו **{field_display}**?")
    else:
        bot.send_message(chat_id, "כל הפרטים קיימים! מפיק את קובץ הוורד...")
        try:
            doc_file = create_report(extracted_data)
            with open(doc_file, 'rb') as f:
                bot.send_document(chat_id, f, caption="הנה דוח הפיקוח העליון המוכן שלך! 📝")
        except Exception as e:
            bot.send_message(chat_id, f"אירעה שגיאה ביצירת הקובץ: {e}")

if __name__ == '__main__':
    print("Bot is running perfectly with docxtpl...")
    bot.infinity_polling()
