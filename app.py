from flask import Flask, request, jsonify, send_from_directory
import imaplib, email, socket
from email.header import decode_header, make_header

IMAP_SERVER = "outlook.office365.com"
IMAP_PORT = 993

app = Flask(__name__, static_folder='.', static_url_path='')

def decode_str(s):
    if not s:
        return ''
    try:
        return str(make_header(decode_header(s)))
    except:
        try:
            return s.decode('utf-8', errors='ignore') if isinstance(s, bytes) else str(s)
        except:
            return str(s)

def get_preview(msg, length=200):
    try:
        if msg.is_multipart():
            for part in msg.walk():
                ctype = part.get_content_type()
                disp = str(part.get('Content-Disposition') or '')
                if ctype == "text/plain" and 'attachment' not in disp:
                    payload = part.get_payload(decode=True)
                    if payload:
                        return payload.decode(errors='ignore').strip().replace('\r','').replace('\n',' ')[:length]
        else:
            payload = msg.get_payload(decode=True)
            if payload:
                return payload.decode(errors='ignore').strip().replace('\r','').replace('\n',' ')[:length]
    except Exception:
        pass
    return ''

def fetch_messages(username, password, folder="INBOX", unseen_only=True, limit=20, filter_from=None, filter_subject=None):
    try:
        socket.create_connection((IMAP_SERVER, IMAP_PORT), timeout=6)
    except Exception as e:
        return {"error": "لا يمكن الاتصال بخادم IMAP: " + str(e)}

    try:
        m = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        m.login(username, password)
        m.select(folder)

        criteria = '(UNSEEN)' if unseen_only else 'ALL'
        status, data = m.search(None, criteria)
        if status != 'OK':
            return {"error": "خطأ أثناء البحث: " + str(status)}

        ids = data[0].split()
        ids = ids[-limit:] if ids else []

        messages = []
        for num in reversed(ids):
            typ, msg_data = m.fetch(num, '(RFC822)')
            if typ != 'OK':
                continue
            raw = msg_data[0][1]
            msg = email.message_from_bytes(raw)

            subject = decode_str(msg.get('Subject', ''))
            from_raw = decode_str(msg.get('From', ''))
            date = msg.get('Date', '')
            preview = get_preview(msg, length=240)

            if filter_from and filter_from.lower() not in from_raw.lower():
                continue
            if filter_subject and filter_subject.lower() not in subject.lower():
                continue

            messages.append({
                "subject": subject,
                "from": from_raw,
                "date": date,
                "preview": preview
            })
        m.logout()
        return {"messages": messages}
    except imaplib.IMAP4.error as e:
        return {"error": "IMAP Error: " + str(e)}
    except Exception as ex:
        return {"error": "خطأ غير متوقع: " + str(ex)}

@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

@app.route('/check', methods=['POST'])
def check_api():
    data = request.get_json(force=True)
    username = data.get('username')
    password = data.get('password')
    if not username or not password:
        return jsonify({"error": "ادخل البريد و كلمة المرور"}), 400

    folder = data.get('folder', 'INBOX')
    unseen_only = bool(data.get('unseen_only', True))
    limit = int(data.get('limit') or 20)
    filter_from = data.get('filter_from') or None
    filter_subject = data.get('filter_subject') or None

    result = fetch_messages(username, password, folder, unseen_only, limit, filter_from, filter_subject)
    if "error" in result:
        return jsonify(result), 500
    return jsonify(result)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
