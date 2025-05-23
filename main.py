# Save the full updated main.py including flood protection and commands

full_main_py = """
import os
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

# === Configuration ===
WEBEX_BOT_TOKEN = os.environ["WEBEX_BOT_TOKEN"]
OPENAI_KEY = os.environ["OPENAI_KEY"]
EMAIL_SENDER = os.environ["EMAIL_SENDER"]
EMAIL_PASSWORD = os.environ["EMAIL_PASSWORD"]
DEFAULT_EMAIL_RECEIVER = os.environ["EMAIL_RECEIVER"]

client = OpenAI(api_key=OPENAI_KEY)
app = Flask(__name__)

# === Investigator emails ===
investigator_emails = {
    "المقدم محمد علي القاسم": "mohammed@example.com",
    "النقيب عبدالله راشد ال علي": "abdullah@example.com",
    "النقيب سليمان محمد الزرعوني": "sulaiman@example.com",
}

# === Field structure ===
expected_fields = [
    "Date", "Briefing", "LocationObservations",
    "Examination", "Outcomes", "TechincalOpinion", "Investigator"
]
field_prompts = {
    "Date": "🎙️ أرسل تاريخ الواقعة.",
    "Briefing": "🎙️ أرسل موجز الواقعة.",
    "LocationObservations": "🎙️ أرسل معاينة الموقع.",
    "Examination": "🎙️ أرسل نتيجة الفحص الفني.",
    "Outcomes": "🎙️ أرسل النتيجة.",
    "TechincalOpinion": "🎙️ أرسل الرأي الفني.",
    "Investigator": "🧑‍✈️ أرسل اسم المحقق."
}
field_names_ar = {
    "Date": "التاريخ",
    "Briefing": "موجز الواقعة",
    "LocationObservations": "معاينة الموقع",
    "Examination": "نتيجة الفحص الفني",
    "Outcomes": "النتيجة",
    "TechincalOpinion": "الرأي الفني",
    "Investigator": "المحقق"
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
    prompt = (
        f"يرجى إعادة صياغة التالي ({field_name}) باستخدام أسلوب مهني وعربي فصيح، "
        f"مع تجنب المشاعر :\\n\\n{user_input}"
    )
    response = client.chat.completions.create(
        model="gpt-4",
        messages=[{"role": "user", "content": prompt}]
    )
    return response.choices[0].message.content.strip()

def format_report_doc(path):
    doc = Document(path)
    for paragraph in doc.paragraphs:
        paragraph.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
        paragraph._element.set(qn("w:rtl"), "1")
        for run in paragraph.runs:
            run.font.name = "Dubai"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "Dubai")
            run.font.size = Pt(13)
    doc.save(path)

def generate_report(data):
    filename = f"تقرير_التحقيق_{data['Investigator'].replace(' ', '_')}.docx"
    doc = DocxTemplate("police_report_template.docx")
    doc.render(data)
    doc.save(filename)
    format_report_doc(filename)
    return filename

def send_email(file_path, recipient, investigator_name):
    msg = EmailMessage()
    msg["Subject"] = "تقرير تحقيق تلقائي"
    msg["From"] = EMAIL_SENDER
    msg["To"] = recipient
    msg.set_content(f"📎 يرجى مراجعة التقرير المرفق.\\n\\nمع تحيات فريق العمل، {investigator_name}.")
    with open(file_path, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=os.path.basename(file_path)
        )
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
        smtp.send_message(msg)

def send_webex_message(room_id, message):
    headers = {
        "Authorization": f"Bearer {WEBEX_BOT_TOKEN}",
        "Content-Type": "application/json"
    }
    payload = {
        "roomId": room_id,
        "markdown": message
    }
    requests.post("https://webexapis.com/v1/messages", headers=headers, json=payload)

@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.json
    room_id = data["data"]["roomId"]
    message_id = data["data"]["id"]
    person_id = data["data"]["personId"]

    headers = {"Authorization": f"Bearer {WEBEX_BOT_TOKEN}"}
    msg_response = requests.get(f"https://webexapis.com/v1/messages/{message_id}", headers=headers)
    msg_data = msg_response.json()

    # Ignore bot messages
    if msg_data.get("personId") == person_id:
        return "OK"

    # Avoid repeated responses to same message
    if "message_id_handled" in user_state.get(person_id, {}) and user_state[person_id]["message_id_handled"] == message_id:
        return "OK"

    user_state.setdefault(person_id, {})["message_id_handled"] = message_id
    message_text = msg_data.get("text", "").strip()

    if message_text == "/start":
        user_state[person_id] = {"step": 0, "data": {}, "message_id_handled": message_id}
        send_webex_message(room_id, "👋 مرحباً بك في بوت إعداد تقارير الفحص.\\nسنجمع المعلومات خطوة بخطوة.\\n🟢 للبدء، أرسل تسجيل صوتي يحتوي على:")
        send_webex_message(room_id, field_prompts[expected_fields[0]])
        return "OK"

    elif message_text == "/reset":
        user_state.pop(person_id, None)
        send_webex_message(room_id, "🔄 تم إعادة ضبط الجلسة. أرسل /start للبدء من جديد.")
        return "OK"

    elif message_text == "/help":
        help_msg = (
            "📌 أوامر البوت:\\n"
            "/start – بدء إدخال تقرير جديد\\n"
            "/reset – إعادة ضبط الجلسة\\n"
            "/help – عرض هذه التعليمات\\n"
            "🎙️ أرسل تسجيل صوتي في كل خطوة"
        )
        send_webex_message(room_id, help_msg)
        return "OK"

    if person_id not in user_state or "step" not in user_state[person_id]:
        send_webex_message(room_id, "⚠️ لم يتم بدء جلسة بعد. أرسل /start للبدء.")
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
            email_to = investigator_emails.get(state["data"]["Investigator"], DEFAULT_EMAIL_RECEIVER)
            send_email(filename, email_to, state["data"]["Investigator"])
            send_webex_message(room_id, f"📩 تم إرسال التقرير إلى {email_to}")
            user_state.pop(person_id)
    else:
        send_webex_message(room_id, "🎙️ الرجاء إرسال تسجيل صوتي.")

    return "OK"

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
"""

with open("/mnt/data/main.py", "w", encoding="utf-8") as f:
    f.write(full_main_py)

"/mnt/data/main.py"

