from pathlib import Path

# Final version of the Webex bot with:
# - Enhanced debug logging for investigator selection
# - Proper handling of file creation for user_state.json
# - Handling Adaptive Card response
final_debugged_bot_code = """
import os
import json
import base64
import tempfile
import requests
from flask import Flask, request
from openai import OpenAI
from docxtpl import DocxTemplate
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import smtplib
from email.message import EmailMessage
from pydub import AudioSegment

# === Configuration ===
WEBEX_BOT_TOKEN = os.environ["WEBEX_BOT_TOKEN"]
OPENAI_KEY = os.environ["OPENAI_KEY"]
EMAIL_SENDER = os.environ["EMAIL_SENDER"]
EMAIL_PASSWORD = os.environ["EMAIL_PASSWORD"]
DEFAULT_EMAIL_RECEIVER = "frnreports@gmail.com"
BOT_EMAIL = "FRN.ENG@webex.bot"
TEMPLATE_FILE = "police_report_template.docx"
STATE_FILE = "/mnt/data/user_state.json"

# Ensure /mnt/data exists and initialize user state
if not os.path.exists("/mnt/data"):
    os.makedirs("/mnt/data")
if os.path.exists(STATE_FILE):
    with open(STATE_FILE, "r", encoding="utf-8") as f:
        user_state = json.load(f)
        print("✅ Loaded existing user_state.")
else:
    user_state = {}
    print("🔄 Initialized new user_state.")

def save_user_state():
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(user_state, f, ensure_ascii=False, indent=2)
        print("💾 Saved user_state.")

app = Flask(__name__)
client = OpenAI(api_key=OPENAI_KEY)

investigator_names = [
    "المقدم محمد علي القاسم", "النقيب عبدالله راشد ال علي",
    "النقيب سليمان محمد الزرعوني", "الملازم أول أحمد خالد الشامسي",
    "العريف راشد محمد بن حسين", "المدني محمد ماهر العلي",
    "المدني امنه خالد المازمي", "المدني حمده ماجد ال علي"
]

field_steps = ["Investigator", "Date", "Briefing"]

field_prompts = {
    "Date": "🗓️ الرجاء إرسال التاريخ كملاحظة صوتية.",
    "Briefing": "📝 الرجاء إرسال ملخص الفحص كملاحظة صوتية."
}

field_labels = {
    "Date": "تم تسجيل التاريخ",
    "Briefing": "تم تسجيل الملخص",
    "Investigator": "تم اختيار المحقق"
}

def format_paragraph(p):
    run = p.runs[0]
    run.font.name = 'Dubai'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Dubai')
    run.font.size = Pt(13)
    p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    p.paragraph_format.right_to_left = True

def format_report_doc(doc):
    for para in doc.paragraphs:
        format_paragraph(para)

def generate_report(data, file_path):
    tpl = DocxTemplate(TEMPLATE_FILE)
    tpl.render(data)
    format_report_doc(tpl.docx)
    tpl.save(file_path)

def transcribe_audio(file_path):
    audio = AudioSegment.from_file(file_path)
    wav_path = tempfile.mktemp(suffix=".wav")
    audio.export(wav_path, format="wav")
    with open(wav_path, "rb") as audio_file:
        transcript = client.audio.transcriptions.create(
            file=audio_file,
            model="whisper-1",
            language="ar"
        )
    return transcript.text

def enhance_field(text, field):
    if field == "Date":
        prompt = f"صيغة التاريخ التالية غير رسمية: '{text}'. صيغه ليكون بصيغة رسمية كاملة."
    elif field == "Briefing":
        prompt = f"لخص المقطع الصوتي التالي بأسلوب مهني لتقريره الفني: '{text}'"
    else:
        return text
    chat = client.chat.completions.create(
        model="gpt-4",
        messages=[
            {"role": "system", "content": "أنت مساعد محترف تكتب تقارير فنية بأسلوب رسمي."},
            {"role": "user", "content": prompt}
        ]
    )
    return chat.choices[0].message.content.strip()

def send_email(subject, body, to, attachment_path):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = EMAIL_SENDER
    msg["To"] = to
    msg.set_content(body)
    with open(attachment_path, "rb") as f:
        msg.add_attachment(f.read(), maintype="application", subtype="vnd.openxmlformats-officedocument.wordprocessingml.document", filename=os.path.basename(attachment_path))
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
        smtp.send_message(msg)

def send_message(person_id, text):
    requests.post("https://webexapis.com/v1/messages", headers={
        "Authorization": f"Bearer {WEBEX_BOT_TOKEN}",
        "Content-Type": "application/json"
    }, json={"toPersonId": person_id, "markdown": text})

def send_adaptive_card(person_id):
    buttons = [{"type": "Action.Submit", "title": name, "data": {"investigator": name}} for name in investigator_names]
    card = {
        "type": "AdaptiveCard",
        "version": "1.3",
        "body": [{"type": "TextBlock", "text": "👤 اختر اسم المحقق:", "weight": "bolder"}],
        "actions": buttons
    }
    requests.post("https://webexapis.com/v1/messages", headers={
        "Authorization": f"Bearer {WEBEX_BOT_TOKEN}",
        "Content-Type": "application/json"
    }, json={
        "toPersonId": person_id,
        "markdown": "اختر اسم المحقق:",
        "attachments": [{"contentType": "application/vnd.microsoft.card.adaptive", "content": card}]
    })

@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.json
    if "data" not in data:
        return "ok"
    user_id = data["data"]["personId"]
    email = data["data"].get("personEmail", "")
    if email == BOT_EMAIL:
        return "ok"

    if data["resource"] == "messages":
        msg_id = data["data"]["id"]
        msg = requests.get(f"https://webexapis.com/v1/messages/{msg_id}", headers={"Authorization": f"Bearer {WEBEX_BOT_TOKEN}"}).json()
        if "files" in msg:
            file_url = msg["files"][0]
            audio_data = requests.get(file_url, headers={"Authorization": f"Bearer {WEBEX_BOT_TOKEN}"}).content
            tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".ogg")
            tmp_file.write(audio_data)
            tmp_file.close()
            field = user_state.get(user_id, {}).get("step")
            text = transcribe_audio(tmp_file.name)
            enhanced = enhance_field(text, field)
            user_state.setdefault(user_id, {}).setdefault("data", {})[field] = enhanced
            next_idx = field_steps.index(field) + 1
            if next_idx < len(field_steps):
                next_field = field_steps[next_idx]
                user_state[user_id]["step"] = next_field
                send_message(user_id, f"{field_labels[field]} ✅\\n{field_prompts[next_field]}")
            else:
                data = user_state[user_id]["data"]
                doc_path = f"/mnt/data/report_{data['Investigator']}.docx"
                generate_report(data, doc_path)
                send_email("تم إنشاء التقرير", f"شكرًا {data['Investigator']}، تم إرسال التقرير بالبريد.", DEFAULT_EMAIL_RECEIVER, doc_path)
                send_message(user_id, f"📄 تم إنشاء التقرير بنجاح وإرساله عبر البريد.\nشكراً لك {data['Investigator']}!")
                user_state.pop(user_id)
            save_user_state()
        else:
            if user_id not in user_state:
                user_state[user_id] = {"step": "Investigator", "data": {}}
                send_message(user_id, "👋 مرحباً بك في بوت إعداد تقارير الفحص الخاص بقسم الهندسة الجنائية.\n📌 أرسل ملاحظة صوتية عند كل طلب.")
                send_adaptive_card(user_id)

    elif data["resource"] == "attachmentActions":
        action_id = data["data"]["id"]
        action_data = requests.get(f"https://webexapis.com/v1/attachment/actions/{action_id}", headers={"Authorization": f"Bearer {WEBEX_BOT_TOKEN}"}).json()
        print("✅ Received Adaptive Card selection:", action_data, flush=True)
        selection = action_data["inputs"]["investigator"]
        user_state[user_id] = {"step": "Date", "data": {"Investigator": selection}}
        send_message(user_id, f"تم اختيار المحقق: {selection} ✅\\n{field_prompts['Date']}")
        save_user_state()
    return "ok"

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
"""

# Write the updated code to a file
final_path = Path("/mnt/data/final_debugged_webex_bot.py")
final_path.write_text(final_debugged_bot_code.strip(), encoding="utf-8")

final_path
