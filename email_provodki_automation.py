#!/usr/bin/env python3
"""
Автоматизация: Почта (Gmail) → Excel «Проводки» → Отчёт → Telegram
Расписание: будни 12:30 МСК
"""

import imaplib, email, os, json, sys, tempfile, re, time
from datetime import datetime, timedelta
from email.header import decode_header
from pathlib import Path
import openpyxl, pandas as pd, requests
try:
    import pyxlsb; HAS_PYXLSB = True
except ImportError:
    HAS_PYXLSB = False

LOCK_FILE = Path(__file__).parent / ".bot_lock_provodki"
LAST_REPORT_FILE = Path(__file__).parent / ".last_report_provodki.txt"
CONFIG_PATH = Path(__file__).parent / "email_provodki_config.json"

def _is_locked():
    if not LOCK_FILE.exists(): return False
    age = time.time() - LOCK_FILE.stat().st_mtime
    if age > 300:
        try: LOCK_FILE.unlink()
        except: pass
        return False
    return True

def _acquire_lock(): LOCK_FILE.write_text(str(time.time()))
def _release_lock():
    try: LOCK_FILE.unlink(missing_ok=True)
    except: pass

def _tg_keyboard():
    return {"inline_keyboard": [
        [{"text": "🔄 Проверить сейчас", "callback_data": "check_now"},
         {"text": "📊 Последний отчёт", "callback_data": "last_report"}],
        [{"text": "⚙️ Статус", "callback_data": "status"}]
    ]}

def load_config():
    if not CONFIG_PATH.exists():
        template = {
            "gmail": {"email": "polypro2005@gmail.com", "app_password": "ВСТАВЬТЕ_ПАРОЛЬ_ПРИЛОЖЕНИЯ_GMAIL", "imap_server": "imap.gmail.com", "imap_port": 993},
            "telegram": {"bot_token": "ВСТАВЬТЕ_TOKEN_TELEGRAM_BOT", "chat_id": "ВСТАВЬТЕ_CHAT_ID", "telegram_user_id": ""},
            "search": {"filename_keyword": "проводки", "file_extensions": [".xlsx", ".xls", ".xlsb"], "max_emails_to_check": 20},
            "schedule": {"time": "12:30", "timezone": "Europe/Moscow", "weekdays_only": True, "monday_3days": False}
        }
        with open(CONFIG_PATH, "w", encoding="utf-8") as f: json.dump(template, f, ensure_ascii=False, indent=2)
        print(f"⚠️  Создан шаблон: {CONFIG_PATH}"); sys.exit(0)
    with open(CONFIG_PATH, "r", encoding="utf-8") as f: return json.load(f)

# ── Почта ──
def decode_mime_header(header_value):
    if header_value is None: return ""
    decoded_parts = decode_header(header_value)
    result = []
    for part, charset in decoded_parts:
        if isinstance(part, bytes):
            try: result.append(part.decode(charset or "utf-8", errors="replace"))
            except: result.append(part.decode("utf-8", errors="replace"))
        else: result.append(part)
    return " ".join(result)

def find_matching_attachment(msg, keyword, extensions):
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_disposition() == "attachment":
                filename = part.get_filename()
                if not filename: continue
                filename_decoded = decode_mime_header(filename).lower()
                for ext in extensions:
                    if keyword.lower() in filename_decoded and filename_decoded.endswith(ext.lower()):
                        return (filename, decode_mime_header(filename), part.get_payload(decode=True))
    return None

def get_latest_excel(config):
    gmail_cfg, search_cfg = config["gmail"], config["search"]
    mail = imaplib.IMAP4_SSL(gmail_cfg["imap_server"], gmail_cfg["imap_port"])
    try: mail.login(gmail_cfg["email"], gmail_cfg["app_password"])
    except imaplib.IMAP4.error as e: print(f"❌ Ошибка входа: {e}"); return None, None, None
    mail.select("INBOX")
    status, messages = mail.search(None, "ALL")
    if status != "OK": mail.logout(); return None, None, None
    msg_ids = messages[0].split() if messages[0] else []
    max_check = min(search_cfg["max_emails_to_check"], len(msg_ids))
    for i in range(1, max_check + 1):
        msg_id = msg_ids[-i]
        status, msg_data = mail.fetch(msg_id, "(RFC822)")
        if status != "OK": continue
        msg = email.message_from_bytes(msg_data[0][1])
        result = find_matching_attachment(msg, keyword=search_cfg["filename_keyword"], extensions=search_cfg["file_extensions"])
        if result:
            raw_name, decoded_name, content = result
            email_info = {"subject": decode_mime_header(msg["Subject"]), "from": decode_mime_header(msg["From"]), "date": msg["Date"], "file": decoded_name}
            mail.logout(); return email_info, content, decoded_name
    mail.logout(); return None, None, None

# ── Форматирование ──
def fmt_num(n):
    try: return f"{int(float(str(n))):,}".replace(",", " ")
    except: return "0"

def short_name(name, max_len=18):
    if not name: return "".ljust(max_len)
    n = name.strip()
    KNOWN = {
        "ПАО СБЕРБАНК": "Сбербанк", "Сбербанк ПАО": "Сбербанк",
        "ПАО СБЕРБАНК Г.МОСКВА": "Сбербанк", "СИБИРСКИЙ БАНК ПАО СБЕРБАНК": "Сбербанк",
        "КАЛУЖСКОЕ ОТДЕЛЕНИЕ N8608 ПАО СБЕРБАНК": "Сбербанк",
        "БАНК ВТБ (ПАО) г. Москва": "ВТБ", "БАНК ГПБ (АО), г.Москва": "ГПБ",
        "МИ ФНС России по управлению долгом": "ФНС", "МИ ФНС России по управлен": "ФНС",
        "УФК по Калужской области (ОСФР по Калужской области)": "УФК",
        "УФК ПО КАЛУЖСКОЙ ОБЛАСТИ (ОСП ПО ВАШ": "УФК", "УФК по Калужской области (УМВД России": "УФК",
        "УФССП РОССИИ ПО КАЛУЖСКОЙ ОБЛАСТИ": "ФССП", "УФК соц.страх": "ФСС",
        "Газпром газораспределение Калуга": "Газпром", "Газпром межрегионгаз Калуга": "Газпром",
        "Калугаоблводоканал": "Водоканал",
    }
    if n in KNOWN: return KNOWN[n][:max_len].ljust(max_len)
    for key, val in KNOWN.items():
        if key in n or n in key: return val[:max_len].ljust(max_len)
    for suffix in [" ООО", " НАО", " АО", " ЗАО"]:
        if n.endswith(suffix): n = n[:-len(suffix)]
    for prefix in ['ООО "', 'ЗАО "', 'НАО "', 'АО "']:
        if n.startswith(prefix):
            n = n[len(prefix):]
            if n.endswith('"'): n = n[:-1]
    if n.startswith("ИП "):
        parts = n[3:].strip().split()
        if parts:
            if len(parts) > 2: n = parts[0].title() + " " + parts[1][0] + "." + parts[2][0] + "."
            elif len(parts) > 1: n = parts[0].title() + " " + parts[1][0] + "."
            else: n = parts[0].title()
    n = " ".join(n.split()).strip()
    if len(n) <= max_len: return n.ljust(max_len)
    words = n.split()
    important = [w for w in words if w.upper() == w and len(w) > 2]
    if important:
        result = " ".join(important)
        if len(result) <= max_len: return result.ljust(max_len)
    return n[:max_len-3] + "..."

# ── Генерация отчёта ──
def generate_report(sheets_data):
    """Генерирует отчёт: по каждому юрлицу за его крайнюю дату."""
    from datetime import datetime as dt
    wd = ["пн","вт","ср","чт","пт","сб","вс"]
    
    # Первый проход: глобальные ширины
    max_nw, max_nn = 6, 10
    for sk in ["дпл","трио","тдп","ип пох"]:
        rows = sheets_data.get(sk, [])
        if not rows: continue
        for row in rows:
            if not row or not row[0]: continue
            d = str(row[0]).strip()
            if len(d) != 10 or d[2] != '.' or d[5] != '.': continue
            if len(row) > 1 and row[1]:
                sn = short_name(str(row[1]), 18).strip()
                if len(sn) > max_nn: max_nn = len(sn)
            for c in [6,7,8,9,10,11]:
                try:
                    if len(row) > c and row[c]:
                        fl = len(fmt_num(row[c]))
                        if fl > max_nw: max_nw = fl
                except: pass
    max_nn = min(max_nn, 18)
    
    sep = "─" * (max_nn + 1 + max_nw)
    all_lines = ["📋 ПРОВОДКИ БАНКОВСКИЕ", ""]
    
    for sk, cp_name in [("дпл","ДПЛ"), ("трио","ТриоПласт"), ("тдп","Торговый дом полимер"), ("ип пох","ИП Похоменко")]:
        rows = sheets_data.get(sk, [])
        if not rows: continue
        
        # Крайняя дата
        last_date = None
        for row in rows:
            if not row or not row[0]: continue
            d = str(row[0]).strip()
            if len(d) == 10 and d[2] == '.' and d[5] == '.':
                try:
                    date = dt.strptime(d, "%d.%m.%Y")
                    if last_date is None or date > last_date: last_date = date
                except: pass
        if last_date is None: continue
        ds = last_date.strftime("%d.%m.%Y")
        
        # Сбор данных за крайнюю дату
        ostatok = 0
        for row in rows:
            if row and len(row) > 6 and row[4] and "остаток на начало" in str(row[4]).lower():
                try: ostatok = float(str(row[6]))
                except: pass
                break
        
        post_by_contr = {}
        spis_by_cat = {}
        total_post = 0.0
        total_spis = 0.0
        
        for row in rows:
            if not row or not row[0]: continue
            if str(row[0]).strip() != ds: continue
            
            contr = str(row[1]).strip() if len(row) > 1 and row[1] else ""
            cat = str(row[4]).strip() if len(row) > 4 and row[4] else "прочие"
            if not cat or cat.lower() in ("статьи затрат","остаток на начало",""): cat = "прочие"
            if not contr: contr = "(без названия)"
            
            post = sum(float(str(row[c])) for c in [6,8,10] if len(row) > c and row[c]) if any(len(row) > c and row[c] for c in [6,8,10]) else 0.0
            spis = sum(float(str(row[c])) for c in [7,9,11] if len(row) > c and row[c]) if any(len(row) > c and row[c] for c in [7,9,11]) else 0.0
            
            # FIX: calculate properly
            post = 0.0; spis = 0.0
            for c in [6,8,10]:
                try:
                    if len(row) > c and row[c]: post += float(str(row[c]))
                except: pass
            for c in [7,9,11]:
                try:
                    if len(row) > c and row[c]: spis += float(str(row[c]))
                except: pass
            
            total_post += post; total_spis += spis
            if post > 0:
                if contr not in post_by_contr: post_by_contr[contr] = 0.0
                post_by_contr[contr] += post
            if spis > 0:
                if cat not in spis_by_cat: spis_by_cat[cat] = {}
                if contr not in spis_by_cat[cat]: spis_by_cat[cat][contr] = 0.0
                spis_by_cat[cat][contr] += spis
        
        if total_post == 0 and total_spis == 0: continue
        
        real_balance = None
        itog_rows = sheets_data.get("итог", [])
        for row in itog_rows:
            if row and row[0] and str(row[0]).strip().lower() == sk:
                try: real_balance = float(str(row[1]))
                except: pass
                if len(row) > 6 and row[6]:
                    try: real_balance += float(str(row[6]))
                    except: pass
                break
        
        all_lines.append(f"🏭 {cp_name}")
        all_lines.append(f"📅 {ds} ({wd[last_date.weekday()]})")
        all_lines.append(f"💰 Приход: {fmt_num(total_post)}")
        all_lines.append(f"💸 Расход: {fmt_num(total_spis)}")
        if real_balance is not None:
            all_lines.append(f"🏦 Остаток на р/с: {fmt_num(real_balance)}")
        
        # ── Детализация ──
        code_rows = []
        sep_line = "─" * (max_nn + 1 + max_nw)
        # ПРИХОД
        if post_by_contr:
            code_rows.append(f"<code>ПРИХОД{'': >{(max_nn + 1 + max_nw) - 6}}</code>")
            code_rows.append(f"<code>{sep_line}</code>")
            sp = sorted(post_by_contr.items(), key=lambda x: x[1], reverse=True)
            for c, v in sp:
                cs = short_name(c, max_nn)[:max_nn]
                code_rows.append(f"<code>{cs} {fmt_num(v).rjust(max_nw)}</code>")
            code_rows.append(f"<code>{sep_line}</code>")
            code_rows.append(f"<code>{'ИТОГО'.ljust(max_nn)} {fmt_num(total_post).rjust(max_nw)}</code>")
        # РАСХОД
        if spis_by_cat:
            code_rows.append(f"<code>РАСХОД{'': >{(max_nn + 1 + max_nw) - 6}}</code>")
            code_rows.append(f"<code>{sep_line}</code>")
            sc = sorted(spis_by_cat.items(), key=lambda x: sum(x[1].values()), reverse=True)
            for cat, cd in sc:
                sc2 = sorted(cd.items(), key=lambda x: x[1], reverse=True)
                code_rows.append(f"<code><b>{cat.capitalize().ljust(max_nn)}</b> {fmt_num(sum(cd.values())).rjust(max_nw)}</code>")
                for c, v in sc2:
                    cs = short_name(c, max_nn)[:max_nn]
                    code_rows.append(f"<code>{cs} {fmt_num(v).rjust(max_nw)}</code>")
            code_rows.append(f"<code>{sep_line}</code>")
            code_rows.append(f"<code>{'ИТОГО'.ljust(max_nn)} {fmt_num(total_spis).rjust(max_nw)}</code>")
        if code_rows:
            all_lines.append("<pre>" + "\n".join(code_rows) + "</pre>")
        
        all_lines.append("")
    
    if len(all_lines) <= 2:
        all_lines.append("📭 Нет данных за последнюю дату.")
    # Выравниваем все <code>-блоки до единой ширины
    import re as _re
    def _visible_len(s):
        """Длина текста без HTML-тегов"""
        return len(_re.sub(r'<[^>]+>', '', s))
    _widths = []
    for _ln in all_lines:
        for _m in _re.finditer(r'<code>(.*?)</code>', _ln):
            _widths.append(_visible_len(_m.group(1)))
    if _widths:
        _max_w = max(_widths)
        _new = []
        for _ln in all_lines:
            _matches = list(_re.finditer(r'<code>(.*?)</code>', _ln))
            if _matches:
                _result = _ln
                for _m in reversed(_matches):
                    _content = _m.group(1)
                    _visible = _visible_len(_content)
                    _need = _max_w - _visible
                    if _need > 0:
                        # Вставляем пробелы перед закрывающим </code>
                        _padded = _content[:len(_content)] + " " * _need
                        _result = _result[: _m.start()] + f"<code>{_padded}</code>" + _result[_m.end() :]
                    else:
                        _result = _result[: _m.start()] + f"<code>{_content}</code>" + _result[_m.end() :]
                _new.append(_result)
            else:
                _new.append(_ln)
        all_lines = _new

    return "\n".join(all_lines)

# ── Telegram ──
def send_telegram(text, config):
    tg_cfg = config["telegram"]
    bt, cid = tg_cfg["bot_token"], tg_cfg["chat_id"]
    if "ВСТАВЬТЕ" in bt or "ВСТАВЬТЕ" in cid: print("⚠️  Telegram не настроен"); return False
    MAX = 3800
    if len(text) <= MAX: return _tg_send(bt, cid, text)
    for i in range(0, len(text), MAX):
        if not _tg_send(bt, cid, text[i:i+MAX]): return False
    return True

def _tg_send(token, chat_id, text):
    try:
        r = requests.post(f"https://api.telegram.org/bot{token}/sendMessage",
            json={"chat_id": chat_id, "text": text, "parse_mode": "HTML", "reply_markup": _tg_keyboard()}, timeout=30)
        if r.status_code == 200:
            print(f"✅ Отправлено ({len(text)} символов)")
            with open(LAST_REPORT_FILE, "w", encoding="utf-8") as f: f.write(text)
            return True
        else: print(f"❌ {r.status_code} {r.text[:150]}"); return False
    except Exception as e: print(f"❌ {e}"); return False

def _tg_send_simple(token, chat_id, text):
    try: requests.post(f"https://api.telegram.org/bot{token}/sendMessage",
        json={"chat_id": chat_id, "text": text, "parse_mode": "HTML"}, timeout=30)
    except: pass

def _log(msg):
    try:
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(Path(__file__).parent / ".bot_polling.log", "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {msg}\n")
    except: pass

def run_polling():
    config = load_config()
    tg_cfg = config["telegram"]
    bt, cid = tg_cfg["bot_token"], int(tg_cfg["chat_id"])
    uid_allowed = str(tg_cfg.get("telegram_user_id", ""))
    _log("🤖 Polling started (проводки)")
    offset = 0
    while True:
        try:
            r = requests.post(f"https://api.telegram.org/bot{bt}/getUpdates",
                json={"offset": offset, "timeout": 30}, timeout=35)
            if r.status_code != 200: time.sleep(5); continue
            for upd in r.json().get("result", []):
                offset = upd["update_id"] + 1
                if "callback_query" in upd:
                    cb = upd["callback_query"]
                    uid = str(cb["from"]["id"])
                    cd = cb["data"]
                    _log(f"callback from={uid} data={cd}")
                    try: requests.post(f"https://api.telegram.org/bot{bt}/answerCallbackQuery",
                        json={"callback_query_id": cb["id"]})
                    except: pass
                    if uid_allowed and uid != uid_allowed:
                        _tg_send_simple(bt, cid, f"⛔ Нет прав (id: <code>{uid}</code>)")
                        continue
                    if cd == "check_now": _handle_check_now(config, bt, cid)
                    elif cd == "last_report": _handle_last_report(bt, cid)
                    elif cd == "status": _handle_status(config, bt, cid)
                elif "message" in upd:
                    msg = upd["message"]
                    txt = msg.get("text", "")
                    if txt == "/start":
                        uid = str(msg["from"]["id"])
                        _tg_send_simple(bt, cid,
                            f"👋 Твой Telegram ID: <code>{uid}</code>\nДобавь в config для авторизации.")
        except Exception as e: _log(f"polling loop ex: {e}"); time.sleep(5)

def _handle_check_now(config, bot_token, chat_id):
    if _is_locked(): _tg_send_simple(bot_token, chat_id, "⏳ Уже выполняется"); return
    _acquire_lock()
    try:
        _log("check_now: started")
        _tg_send_simple(bot_token, chat_id, "⏳ Проверяю почту…")
        email_info, content, filename = get_latest_excel(config)
        if not content: _tg_send_simple(bot_token, chat_id, "📭 Нового файла нет."); return
        suffix = Path(filename).suffix.lower()
        sheets_data = {}
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp: tmp.write(content); tp = tmp.name
        try:
            xls = pd.ExcelFile(tp, engine="openpyxl")
            for sn in xls.sheet_names:
                try:
                    df = pd.read_excel(xls, sheet_name=sn, dtype=str, header=None)
                    sheets_data[sn] = df.fillna("").values.tolist()
                except: pass
            report = generate_report(sheets_data)
            send_telegram(report, config)
            out_dir = Path(__file__).parent / "reports"; out_dir.mkdir(exist_ok=True)
            with open(out_dir / f"report_provodki_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt", "w", encoding="utf-8") as f: f.write(report)
        finally:
            try: os.unlink(tp)
            except: pass
        _log("check_now: done")
    finally: _release_lock()

def _handle_status(config, bot_token, chat_id):
    parts = ["🤖 <b>Бот проводок</b>: ✅ работает"]
    g = config["gmail"]
    parts.append("📧 <b>Почта</b>: ✅" if "ВСТАВЬТЕ" not in g.get("app_password","") else "📧 <b>Почта</b>: ❌")
    t = config["telegram"]
    parts.append("📨 <b>Telegram</b>: ✅" if "ВСТАВЬТЕ" not in t.get("bot_token","") else "📨 <b>Telegram</b>: ❌")
    if LAST_REPORT_FILE.exists():
        mtime = datetime.fromtimestamp(LAST_REPORT_FILE.stat().st_mtime)
        parts.append(f"📊 Последний отчёт: {(datetime.now()-mtime).seconds//60} мин назад")
    else: parts.append("📊 Последний отчёт: нет")
    _tg_send_simple(bot_token, chat_id, "\n".join(parts))

def _handle_last_report(bot_token, chat_id):
    if not LAST_REPORT_FILE.exists(): _tg_send_simple(bot_token, chat_id, "📭 Нет отчёта."); return
    _tg_send(bot_token, chat_id, LAST_REPORT_FILE.read_text(encoding="utf-8"))

def main():
    print("="*50); print("💳 Проводки → Telegram")
    config = load_config()
    email_info, content, filename = get_latest_excel(config)
    if not content: print("❌ Файл не найден"); send_telegram("📭 Файл «Проводки» не найден.", config); return
    print(f"✅ Найден: {filename}")
    suffix = Path(filename).suffix.lower()
    sheets_data = {}
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp: tmp.write(content); tp = tmp.name
    try:
        xls = pd.ExcelFile(tp, engine="openpyxl")
        for sn in xls.sheet_names:
            try:
                df = pd.read_excel(xls, sheet_name=sn, dtype=str, header=None)
                sheets_data[sn] = df.fillna("").values.tolist()
            except: pass
        report = generate_report(sheets_data)
        print(f"\n📊 Отчёт ({len(report)}):\n{report}")
        send_telegram(report, config)
        out_dir = Path(__file__).parent / "reports"; out_dir.mkdir(exist_ok=True)
        rf = out_dir / f"report_provodki_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        with open(rf, "w", encoding="utf-8") as f: f.write(report)
        print(f"💾 {rf}")
    finally:
        try: os.unlink(tp)
        except: pass
    print("✅ Готово!")

if __name__ == "__main__":
    if "--polling" in sys.argv: run_polling()
    else: main()
