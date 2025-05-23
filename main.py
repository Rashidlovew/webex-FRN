# FILE: main.py

import os
from flask import Flask, request
from docxtpl import DocxTemplate
from docx.shared import Pt
from docx.oxml.ns import qn
from docx import Document
from pydub import AudioSegment
from email.message import EmailMessage
import smtplib
from openai import OpenAI
import requests
import json

WEBEX_BOT_EMAIL = "FRN.ENG@webex.bot"
WEBEX_BOT_TOKEN = os.environ["WEBEX_BOT_TOKEN"]
OPENAI_KEY = os.environ["OPENAI_KEY"]
EMAIL_SENDER = os.environ["EMAIL_SENDER"]
EMAIL_PASSWORD = os.environ["EMAIL_PASSWORD"]
DEFAULT_EMAIL_RECEIVER = os.environ["EMAIL_RECEIVER"]

client = OpenAI(api_key=OPENAI_KEY)
app = Flask(__name__)
user_state = {}

investigator_names = [
    "المقدم محمد علي القاسم", "النقيب عبدالله راشد ال علي", "النقيب سليمان محمد الزرعوني",
    "الملازم أول أحمد خالد الشامسي", "العريف راشد محمد بن حسين",
    "المدني محمد ماهر العلي", "المدني امنه خالد المازمي", "المدني حمده ماجد ال علي"
]

expected_fields = ["Investigator", "Date", "Briefing", "LocationObservations", "Examination", "Outcomes", "TechincalOpinion"]
field_prompts = {
    "Investigator": "🧑‍✈️ يرجى اختيار اسم الفاحص من القائمة.",
    "Date": "🎙️ أرسل تاريخ الواقعة.",
    "Briefing": "🎙️ أرسل موجز الواقعة.",
    "LocationObservations": "🎙️ أرسل معاينة الموقع.",
    "Examination": "🎙️ أرسل نتيجة الفحص الفني.",
    "Outcomes": "🎙️ أرسل النتيجة.",
    "TechincalOpinion": "🎙️ أرسل الرأي الفني."
}
field_names_ar = {
    "Investigator": "الفاحص", "Date": "التاريخ", "Briefing": "موجز الواقعة",
    "LocationObservations": "معاينة الموقع", "Examination": "نتيجة الفحص الفني",
    "Outcomes": "النتيجة", "TechincalOpinion": "الرأي الفني"
}

def transcribe(file_path):
    audio = AudioSegment.from_file(file_path)
    audio.export("converted.wav", format="wav")
    with open("converted.wav", "rb") as f:
        result = client.audio.transcriptions.create(model="whisper-1", file=f, language="ar")
    return result.text

def enhance_with_gpt(field, text):
    prompt = f"يرجى إعادة صياغة التالي ({field}) باستخدام أسلوب مهني وعربي فصيح:\n\n{text}" if field != "التاريخ" else f"يرجى صياغة تاريخ الواقعة بهذا الشكل: 25/مايو/2025. النص:\n\n{text}"
    response = client.chat.completions.create(model="gpt-4", messages=[{"role": "user", "content": prompt}])
    return response.choices[0].message.content.strip()

def format_report_doc(path):
    doc = Document(path)
    for p in doc.paragraphs:
        for run in p.runs:
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

def send_email(file_path, to, name):
    msg = EmailMessage()
    msg["Subject"] = "تقرير فحص تلقائي"
    msg["From"] = EMAIL_SENDER
    msg["To"] = to
    msg.set_content(f"📎 يرجى مراجعة التقرير المرفق.\n\nمع تحيات فريق العمل، {name}.")
    with open(file_path, "rb") as f:
        msg.add_attachment(f.read(), maintype="application", subtype="vnd.openxmlformats-officedocument.wordprocessingml.document", filename=os.path.basename(file_path))
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
        smtp.send_message(msg)

def send_webex_message(room_id, message):
    headers = {"Authorization": f"Bearer {WEBEX_BOT_TOKEN}", "Content-Type": "application/json"}
    data = {"roomId": room_id, "markdown": message}
    requests.post("https://webexapis.com/v1/messages", headers=headers, json=data)

def send_investigator_card(room_id):
    card = {
        "roomId": room_id,
        "markdown": "يرجى اختيار اسم الفاحص:",
        "attachments": [{
            "contentType": "application/vnd.microsoft.card.adaptive",
            "content": {
                "type": "AdaptiveCard", "version": "1.2",
                "body": [{
                    "type": "Input.ChoiceSet", "id": "investigator", "style": "expanded",
                    "choices": [{"title": name, "value": name} for name in investigator_names]
                }],
                "actions": [{"type": "Action.Submit", "title": "إرسال"}]
            }
        }]
    }
    headers = {"Authorization": f"Bearer {WEBEX_BOT_TOKEN}", "Content-Type": "application/json"}
    requests.post("https://webexapis.com/v1/messages", headers=headers, json=card)

@app.route("/")
def index():
    return "Bot is running", 200

@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.json
    person_id = data["data"]["personId"]
    room_id = data["data"]["roomId"]

    if "attachmentActionId" in data["data"]:
        action_id = data["data"]["attachmentActionId"]
        if user_state.get(person_id, {}).get("handled_action") == action_id:
            return "OK"

        headers = {"Authorization": f"Bearer {WEBEX_BOT_TOKEN}"}
        res = requests.get(f"https://webexapis.com/v1/attachment/actions/{action_id}", headers=headers)
        inputs = res.json()["inputs"]
        name = inputs.get("investigator")

        if name:
            user_state[person_id] = {"step": 1, "data": {"Investigator": name}, "handled_action": action_id}
            send_webex_message(room_id, f"✅ تم اختيار {name}.\n{field_prompts['Date']}")
        return "OK"

    msg_id = data["data"]["id"]
    headers = {"Authorization": f"Bearer {WEBEX_BOT_TOKEN}"}
    msg = requests.get(f"https://webexapis.com/v1/messages/{msg_id}", headers=headers).json()
    text = msg.get("text", "").strip()

    if msg.get("personEmail") == WEBEX_BOT_EMAIL:
        return "OK"

    if text == "/reset":
        user_state.pop(person_id, None)
        send_webex_message(room_id, "🔄 تم إعادة ضبط الجلسة. أرسل رسالة جديدة للبدء.")
        return "OK"

    if person_id not in user_state:
        send_webex_message(room_id, "👋 مرحباً بك في بوت إعداد تقارير الفحص الخاص بقسم الهندسة الجنائية.\n🎙️ سيتم إدخال البيانات عبر تسجيلات صوتية.\n🔄 لإعادة البدء أرسل /reset")
        send_investigator_card(room_id)
        return "OK"

    if "files" in msg:
        state = user_state[person_id]
        step = state["step"]
        field = expected_fields[step]
        file_url = msg["files"][0]
        audio = requests.get(file_url, headers=headers)
        with open("voice.mp4", "wb") as f:
            f.write(audio.content)
        raw = transcribe("voice.mp4")
        refined = enhance_with_gpt(field_names_ar[field], raw)
        state["data"][field] = refined
        state["step"] += 1
        if state["step"] < len(expected_fields):
            send_webex_message(room_id, f"✅ تم تسجيل {field_names_ar[field]}.\n{field_prompts[expected_fields[state['step']]]}")
        else:
            send_webex_message(room_id, "✅ تم استلام جميع البيانات. جاري إعداد التقرير...")
            report = generate_report(state["data"])
            send_email(report, DEFAULT_EMAIL_RECEIVER, state["data"]["Investigator"])
            send_webex_message(room_id, f"📩 تم إرسال التقرير إلى {DEFAULT_EMAIL_RECEIVER}")
            user_state.pop(person_id)
    else:
        send_webex_message(room_id, "🎙️ الرجاء إرسال تسجيل صوتي.")

    return "OK"

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
