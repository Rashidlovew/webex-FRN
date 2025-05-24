from pathlib import Path

updated_code = """
import os
import json
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

# Configuration
WEBEX_BOT_TOKEN = os.environ["WEBEX_BOT_TOKEN"]
OPENAI_KEY = os.environ["OPENAI_KEY"]
EMAIL_SENDER = os.environ["EMAIL_SENDER"]
EMAIL_PASSWORD = os.environ["EMAIL_PASSWORD"]
DEFAULT_EMAIL_RECEIVER = "frnreports@gmail.com"
BOT_EMAIL = "FRN.ENG@webex.bot"
TEMPLATE_FILE = "police_report_template.docx"
STATE_FILE = "user_state.json"

# Load or initialize state
if os.path.exists(STATE_FILE):
    with open(STATE_FILE, "r", encoding="utf-8") as f:
        user_state = json.load(f)
        print("ℹ️ Loaded previous state")
else:
    user_state = {}
    print("ℹ️ No previous state found.")

def save_user_state():
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(user_state, f, ensure_ascii=False, indent=2)

app = Flask(__name__)
client = OpenAI(api_key=OPENAI_KEY)

investigator_names = [
    "المقدم محمد علي القاسم", "النقيب عبدالله راشد ال علي",
    "النقيب سليمان محمد الزرعوني", "الملازم أول أحمد خالد الشامسي",
    "العريف راشد محمد بن حسين", "المدني محمد ماهر العلي",
    "المدني امنه خالد المازمي", "المدني حمده ماجد ال علي"
]

expected_fields = [
    "Date", "Briefing", "LocationObservations",
    "Examination", "Outcomes", "TechincalOpinion"
]

field_prompts = {
    "Date": "🎙️ أرسل تاريخ الواقعة.",
    "Briefing": "🎙️ أرسل موجز الواقعة.",
    "LocationObservations": "🎙️ أرسل معاينة الموقع حيث بمعاينة موقع الحادث تبين ما يلي .....",
    "Examination": "🎙️ أرسل نتيجة الفحص الفني ... حيث بفحص موضوع الحادث تبين ما يلي .....",
    "Outcomes": "🎙️ أرسل النتيجة حيث أنه بعد المعاينة و أجراء الفحوص الفنية اللازمة تبين ما يلي:.",
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

def format_paragraph(p):
    if p.runs:
        run = p.runs[0]
        run.font.name = 'Dubai'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Dubai')
        run.font.size = Pt(17)

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

def enhance_with_gpt(field_name, user_input):
    if field_name == "TechincalOpinion":
        prompt = (
            f"يرجى إعادة صياغة ({field_name}) التالية بطريقة مهنية وتحليلية، "
            f"وباستخدام لغة رسمية وعربية فصحى:\n\n{user_input}"
        )
    elif field_name == "Date":
        prompt = (
            f"يرجى صياغة تاريخ الواقعة بالتنسيق التالي فقط: 20/مايو/2025. النص:\n\n{user_input}"
        )
    else:
        prompt = (
            f"يرجى إعادة صياغة التالي ({field_name}) باستخدام أسلوب مهني وعربي فصيح، "
            f"مع تجنب المشاعر :\n\n{user_input}"
        )

    response = client.chat.completions.create(
        model="gpt-4",
        messages=[{"role": "user", "content": prompt}]
    )
    return response.choices[0].message.content.strip()

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

def send_message(person_id, text, parent_id=None):
    payload = {"toPersonId": person_id, "markdown": text}
    if parent_id:
        payload["parentId"] = parent_id
    requests.post("https://webexapis.com/v1/messages", headers={
        "Authorization": f"Bearer {WEBEX_BOT_TOKEN}",
        "Content-Type": "application/json"
    }, json=payload)

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
    parent_id = data["data"]["id"]
    if email == BOT_EMAIL:
        return "ok"

    if data["resource"] == "messages":
        msg = requests.get(f"https://webexapis.com/v1/messages/{parent_id}", headers={"Authorization": f"Bearer {WEBEX_BOT_TOKEN}"}).json()
        if "files" in msg:
            file_url = msg["files"][0]
            audio_data = requests.get(file_url, headers={"Authorization": f"Bearer {WEBEX_BOT_TOKEN}"}).content
            tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".ogg")
            tmp_file.write(audio_data)
            tmp_file.close()
            step = user_state.get(user_id, {}).get("step")
            if not step:
                send_message(user_id, "❗ لم يتم تحديد الخطوة الحالية. أرسل /start للبدء.", parent_id)
                return "ok"
            text = transcribe_audio(tmp_file.name)
            result = enhance_with_gpt(step, text)
            user_state.setdefault(user_id, {}).setdefault("data", {})[step] = result
            next_index = expected_fields.index(step) + 1
            if next_index < len(expected_fields):
                next_step = expected_fields[next_index]
                user_state[user_id]["step"] = next_step
                send_message(user_id, f"{field_names_ar[step]} ✅\\n{field_prompts[next_step]}", parent_id)
            else:
                data_dict = user_state[user_id]["data"]
                report_file = f"report_{data_dict['Investigator']}.docx"
                generate_report(data_dict, report_file)
                send_email("تم إنشاء التقرير", f"شكرًا {data_dict['Investigator']}، تم إرسال التقرير بالبريد.", DEFAULT_EMAIL_RECEIVER, report_file)
                send_message(user_id, f"📄 تم إنشاء التقرير بنجاح وإرساله عبر البريد.\nشكراً لك {data_dict['Investigator']}!", parent_id)
                user_state.pop(user_id)
            save_user_state()
        else:
            if user_id not in user_state:
                user_state[user_id] = {"step": "Investigator", "data": {}}
                send_message(user_id, "👋 مرحباً بك في بوت إعداد تقارير الفحص.\n📌 أرسل ملاحظة صوتية عند كل طلب.", parent_id)
                send_adaptive_card(user_id)
    elif data["resource"] == "attachmentActions":
        action_id = data["data"]["id"]
        action_data = requests.get(f"https://webexapis.com/v1/attachment/actions/{action_id}", headers={"Authorization": f"Bearer {WEBEX_BOT_TOKEN}"}).json()
        selection = action_data["inputs"]["investigator"]
        user_state[user_id] = {"step": expected_fields[0], "data": {"Investigator": selection}}
        send_message(user_id, f"تم اختيار المحقق: {selection} ✅\\n{field_prompts[expected_fields[0]]}", parent_id)
        save_user_state()
    return "ok"

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
"""

#final_path = Path("/mnt/data/final_bot_with_custom_enhancement.py")
#final_path.write_text(updated_code.strip(), encoding="utf-8")
#final_path
