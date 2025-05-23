full_code = '''import os
import json
from flask import Flask, request
from docxtpl import DocxTemplate
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx import Document
from pydub import AudioSegment
from email.message import EmailMessage
import smtplib
from openai import OpenAI
import requests

WEBEX_BOT_EMAIL = "FRN.ENG@webex.bot"

# === Configuration ===
WEBEX_BOT_TOKEN = os.environ["WEBEX_BOT_TOKEN"]
OPENAI_KEY = os.environ["OPENAI_KEY"]
EMAIL_SENDER = os.environ["EMAIL_SENDER"]
EMAIL_PASSWORD = os.environ["EMAIL_PASSWORD"]
DEFAULT_EMAIL_RECEIVER = "frnreports@gmail.com"

client = OpenAI(api_key=OPENAI_KEY)
app = Flask(__name__)

# === Investigator Names ===
investigator_names = [
    "المقدم محمد علي القاسم",
    "النقيب عبدالله راشد ال علي",
    "النقيب سليمان محمد الزرعوني",
    "الملازم أول أحمد خالد الشامسي",
    "العريف راشد محمد بن حسين",
    "المدني محمد ماهر العلي",
    "المدني امنه خالد المازمي",
    "المدني حمده ماجد ال علي"
]

expected_fields = [
    "Date", "Briefing", "LocationObservations",
    "Examination", "Outcomes", "TechincalOpinion"
]

field_prompts = {
    "Date": "🎙️ أرسل تاريخ الواقعة.",
    "Briefing": "🎙️ أرسل موجز الواقعة.",
    "LocationObservations": "🎙️ أرسل معاينة الموقع.",
    "Examination": "🎙️ أرسل نتيجة الفحص الفني.",
    "Outcomes": "🎙️ أرسل النتيجة.",
    "TechincalOpinion": "🎙️ أرسل الرأي الفني."
}

field_names_ar = {
    "Date": "التاريخ",
    "Briefing": "موجز الواقعة",
    "LocationObservations": "معاينة الموقع",
    "Examination": "نتيجة الفحص الفني",
    "Outcomes": "النتيجة",
    "TechincalOpinion": "الرأي الفني"
}

user_state = {}

# === Utilities ===
def transcribe(file_path):
    audio = AudioSegment.from_file(file_path)
    audio.export("converted.wav", format="wav")
    with open("converted.wav", "rb") as f:
        result = client.audio.transcriptions.create(model="whisper-1", file=f, language="ar")
    return result.text

def enhance_with_gpt(field_name, user_input):
    if field_name == "TechincalOpinion":
        prompt = f"يرجى إعادة صياغة ({field_name}) التالية بطريقة مهنية وتحليلية، وباستخدام لغة رسمية وعربية فصحى:\n\n{user_input}"
    elif field_name == "Date":
        prompt = f"يرجى صياغة تاريخ الواقعة بالتنسيق التالي فقط: 20/مايو/2025. النص:\n\n{user_input}"
    else:
        prompt = f"يرجى إعادة صياغة التالي ({field_name}) باستخدام أسلوب مهني وعربي فصيح، مع تجنب المشاعر:\n\n{user_input}"

    response = client.chat.completions.create(
        model="gpt-4",
        messages=[{"role": "user", "content": prompt}]
    )
    return response.choices[0].message.content.strip()

def format_report_doc(path):
    doc = Document(path)
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run.font.name = "Dubai"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "Dubai")
            run.font.size = Pt(13)
    doc.save(path)

def generate_report(data):
    filename = f"تقرير الفحص {data['Investigator'].replace(' ', '_')}.docx"
    doc = DocxTemplate("police_report_template.docx")
    doc.render(data)
    doc.save(filename)
    format_report_doc(filename)
    return filename

def send_email(file_path, recipient, investigator_name):
    msg = EmailMessage()
    msg["Subject"] = "تقرير فحص تلقائي"
    msg["From"] = EMAIL_SENDER
    msg["To"] = recipient
    msg.set_content(f"📎 يرجى مراجعة التقرير المرفق.\n\nمع تحيات فريق العمل، {investigator_name}.")
    with open(file_path, "rb") as f:
        msg.add_attachment(f.read(), maintype="application", subtype="vnd.openxmlformats-officedocument.wordprocessingml.document", filename=os.path.basename(file_path))
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
        smtp.send_message(msg)

def send_webex_message(room_id, message):
    headers = {
        "Authorization": f"Bearer {WEBEX_BOT_TOKEN}",
        "Content-Type": "application/json"
    }
    payload = {"roomId": room_id, "markdown": message}
    requests.post("https://webexapis.com/v1/messages", headers=headers, json=payload)

def send_investigator_card(room_id):
    card = {
        "roomId": room_id,
        "markdown": "🧑‍✈️ الرجاء اختيار اسم الفاحص:",
        "attachments": [{
            "contentType": "application/vnd.microsoft.card.adaptive",
            "content": {
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "type": "AdaptiveCard",
                "version": "1.4",
                "body": [{
                    "type": "TextBlock",
                    "text": "🧑‍✈️ الرجاء اختيار اسم الفاحص:",
                    "wrap": True
                }],
                "actions": [
                    {
                        "type": "Action.Submit",
                        "title": name,
                        "data": {"investigator": name}
                    } for name in investigator_names
                ]
            }
        }]
    }
    headers = {
        "Authorization": f"Bearer {WEBEX_BOT_TOKEN}",
        "Content-Type": "application/json"
    }
    requests.post("https://webexapis.com/v1/messages", headers=headers, json=card)

@app.route("/")
def index():
    return "Bot is running", 200

@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.json
    room_id = data["data"]["roomId"]
    message_id = data["data"]["id"]
    person_id = data["data"]["personId"]

    headers = {"Authorization": f"Bearer {WEBEX_BOT_TOKEN}"}
    msg_response = requests.get(f"https://webexapis.com/v1/messages/{message_id}", headers=headers)
    msg_data = msg_response.json()

    if msg_data.get("personEmail") == WEBEX_BOT_EMAIL:
        return "OK"

    user_state.setdefault(person_id, {})
    if user_state[person_id].get("message_id_handled") == message_id:
        return "OK"
    user_state[person_id]["message_id_handled"] = message_id

    if "data" in data and "investigator" in data["data"]:
        selected = data["data"]["investigator"]
        user_state[person_id] = {"step": 0, "data": {"Investigator": selected}, "message_id_handled": message_id}
        send_webex_message(room_id, f"✅ تم اختيار الفاحص: {selected}\n{field_prompts[expected_fields[0]]}")
        return "OK"

    if person_id not in user_state or "data" not in user_state[person_id] or "Investigator" not in user_state[person_id]["data"]:
        send_investigator_card(room_id)
        return "OK"

    state = user_state[person_id]
    step = state["step"]

    if "files" in msg_data:
        file_url = msg_data["files"][0]
        audio = requests.get(file_url, headers=headers)
        with open("voice.mp4", "wb") as f:
            f.write(audio.content)

        transcribed = transcribe("voice.mp4")
        current_field = expected_fields[step]
        enhanced = enhance_with_gpt(field_names_ar[current_field], transcribed)

        state["data"][current_field] = enhanced
        state["step"] += 1

        if state["step"] < len(expected_fields):
            next_field = expected_fields[state["step"]]
            send_webex_message(room_id, f"✅ تم تسجيل {field_names_ar[current_field]}.\n{field_prompts[next_field]}")
        else:
            send_webex_message(room_id, "✅ تم استلام جميع البيانات. جاري إعداد التقرير...")
            filename = generate_report(state["data"])
            send_email(filename, DEFAULT_EMAIL_RECEIVER, state["data"]["Investigator"])
            send_webex_message(room_id, f"📩 تم إرسال التقرير إلى {DEFAULT_EMAIL_RECEIVER}")
            user_state.pop(person_id)

    else:
        send_webex_message(room_id, "🎙️ الرجاء إرسال تسجيل صوتي.")

    return "OK"

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
'''

with open("/mnt/data/main.py", "w", encoding="utf-8") as f:
    f.write(full_code)

"/mnt/data/main.py has been saved with the updated code."
