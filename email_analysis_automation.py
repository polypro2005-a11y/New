#!/usr/bin/env python3
"""
Автоматизация: Почта (Gmail) → Excel «Анализ производства» → Отчёт → Telegram

Режимы:
  - Будни (Вт–Пт): отчёт за последние сутки
  - Понедельник:    3 отчёта за (Пт, Сб, Вс)
  - Расписание:     11:00 МСК, только будни
"""

import imaplib
import email
import os
import json
import sys
import tempfile
import re
import time
from datetime import datetime, timedelta
from email.header import decode_header
from pathlib import Path

import openpyxl
import pandas as pd
import requests

try:
    import pyxlsb
    HAS_PYXLSB = True
except ImportError:
    HAS_PYXLSB = False

# ── Lock / State ──
LOCK_FILE = Path(__file__).parent / ".bot_lock"
LAST_REPORT_FILE = Path(__file__).parent / ".last_report.txt"

def _is_locked():
    if not LOCK_FILE.exists():
        return False
    age = time.time() - LOCK_FILE.stat().st_mtime
    if age > 300:  # 5 min stale timeout
        try:
            LOCK_FILE.unlink()
        except:
            pass
        return False
    return True

def _acquire_lock():
    LOCK_FILE.write_text(str(time.time()))

def _release_lock():
    try:
        LOCK_FILE.unlink(missing_ok=True)
    except:
        pass

def _tg_keyboard():
    return {
        "inline_keyboard": [
            [
                {"text": "🔄 Проверить сейчас", "callback_data": "check_now"},
                {"text": "📊 Последний отчёт", "callback_data": "last_report"}
            ],
            [
                {"text": "⚙️ Статус", "callback_data": "status"}
            ]
        ]
    }

CONFIG_PATH = Path(__file__).parent / "email_report_config.json"

# ── Конфигурация ──────────────────────────────────────────────────

def load_config():
    if not CONFIG_PATH.exists():
        template = {
            "gmail": {
                "email": "polypro2005@gmail.com",
                "app_password": "ВСТАВЬТЕ_ПАРОЛЬ_ПРИЛОЖЕНИЯ_GMAIL",
                "imap_server": "imap.gmail.com",
                "imap_port": 993
            },
            "telegram": {
                "bot_token": "ВСТАВЬТЕ_TOKEN_TELEGRAM_BOT",
                "chat_id": "ВСТАВЬТЕ_CHAT_ID",
                "telegram_user_id": ""
            },
            "search": {
                "filename_keyword": "анализ",
                "file_extensions": [".xlsx", ".xls", ".xlsb"],
                "max_emails_to_check": 20
            },
            "schedule": {
                "time": "11:00",
                "timezone": "Europe/Moscow",
                "weekdays_only": True,        # только будни
                "monday_3days": True,         # в пн — 3 отчёта за выходные
                "monday_days_back": 3
            }
        }
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(template, f, ensure_ascii=False, indent=2)
        print(f"⚠️  Создан шаблон: {CONFIG_PATH}")
        sys.exit(0)

    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


# ── Почта ─────────────────────────────────────────────────────────

def decode_mime_header(header_value):
    if header_value is None:
        return ""
    decoded_parts = decode_header(header_value)
    result = []
    for part, charset in decoded_parts:
        if isinstance(part, bytes):
            try:
                result.append(part.decode(charset or "utf-8", errors="replace"))
            except (LookupError, UnicodeDecodeError):
                result.append(part.decode("utf-8", errors="replace"))
        else:
            result.append(part)
    return " ".join(result)


def find_matching_attachment(msg, keyword, extensions):
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_disposition() == "attachment":
                filename = part.get_filename()
                if not filename:
                    continue
                filename_decoded = decode_mime_header(filename).lower()
                for ext in extensions:
                    if keyword.lower() in filename_decoded and filename_decoded.endswith(ext.lower()):
                        return (filename, decode_mime_header(filename), part.get_payload(decode=True))
    return None


def get_latest_excel(config):
    gmail_cfg = config["gmail"]
    search_cfg = config["search"]
    mail = imaplib.IMAP4_SSL(gmail_cfg["imap_server"], gmail_cfg["imap_port"])
    try:
        mail.login(gmail_cfg["email"], gmail_cfg["app_password"])
    except imaplib.IMAP4.error as e:
        print(f"❌ Ошибка входа: {e}")
        return None, None, None
    mail.select("INBOX")
    status, messages = mail.search(None, "ALL")
    if status != "OK":
        mail.logout()
        return None, None, None
    msg_ids = messages[0].split() if messages[0] else []
    max_check = min(search_cfg["max_emails_to_check"], len(msg_ids))
    for i in range(1, max_check + 1):
        msg_id = msg_ids[-i]
        status, msg_data = mail.fetch(msg_id, "(RFC822)")
        if status != "OK":
            continue
        msg = email.message_from_bytes(msg_data[0][1])
        result = find_matching_attachment(
            msg,
            keyword=search_cfg["filename_keyword"],
            extensions=search_cfg["file_extensions"]
        )
        if result:
            raw_name, decoded_name, content = result
            email_info = {
                "subject": decode_mime_header(msg["Subject"]),
                "from": decode_mime_header(msg["From"]),
                "date": msg["Date"],
                "file": decoded_name,
            }
            mail.logout()
            return email_info, content, decoded_name
    mail.logout()
    return None, None, None


# ── Обработка Excel ──────────────────────────────────────────────

def excel_serial_to_date(serial):
    base = datetime(1899, 12, 30)
    return base + timedelta(days=serial)


def get_weekday_ru(dt):
    days = ["пн", "вт", "ср", "чт", "пт", "сб", "вс"]
    return days[dt.weekday()]


def find_data_start(rows, title_keywords):
    """Находит индекс первой строки с данными (после заголовков)."""
    for i, row in enumerate(rows):
        if not row or not row[0]:
            continue
        first = str(row[0]).strip().lower()
        if first == 'дата':
            return i + 1
        if any(kw in first for kw in title_keywords):
            continue
    return 1


def extract_all_dates(rows, kg_col, title_keywords):
    """
    Извлекает все даты с данными из листа.
    Возвращает список (date_serial, day_sum, night_sum, total),
    отсортированный по дате.
    """
    start_idx = find_data_start(rows, title_keywords)
    date_data = {}

    for row in rows[start_idx:]:
        if not row or not row[0]:
            continue
        first = str(row[0]).strip().lower()
        if first in ('итого', ''):
            continue
        try:
            date_val = float(str(row[0]))
        except ValueError:
            continue
        if not (45000 < date_val < 48000):
            continue

        shift = str(row[1]).strip().lower() if len(row) > 1 and row[1] else ''
        if shift not in ('д', 'н'):
            continue

        try:
            kg = float(str(row[kg_col])) if len(row) > kg_col and row[kg_col] else 0.0
        except (ValueError, IndexError):
            kg = 0.0
        if kg == 0.0:
            continue

        if date_val not in date_data:
            date_data[date_val] = {'д': 0.0, 'н': 0.0}
        date_data[date_val][shift] += kg

    # Сортируем даты
    sorted_dates = sorted(date_data.keys())
    result = []
    for d in sorted_dates:
        day_v = round(date_data[d]['д'])
        night_v = round(date_data[d]['н'])
        result.append((d, day_v, night_v, day_v + night_v))
    return result


def get_data_for_date(all_dates, target_serial):
    """Возвращает данные для конкретной даты из списка."""
    for d, day_v, night_v, total in all_dates:
        if abs(d - target_serial) < 0.001:
            return day_v, night_v, total
    return 0, 0, 0


def find_total_row(rows, kg_col, title_keywords):
    for row in rows:
        if not row or not row[0]:
            continue
        first = str(row[0]).strip().lower()
        if first == 'итого':
            try:
                return float(str(row[kg_col])) if len(row) > kg_col and row[kg_col] else 0.0
            except:
                return 0.0
    return 0.0


def extract_forecast(sheets_data):
    """Извлекает данные прогноза по каждой единице (per-column)."""
    прогноз_rows = sheets_data.get("прогноз", [])

    granula_cur = [0.0] * 5
    granula_prog = [0.0] * 5
    polu_cur = [0.0] * 4
    polu_prog = [0.0] * 4
    peregon = 0

    for row in прогноз_rows:
        if not row:
            continue
        first = str(row[0]).strip().lower() if row[0] else ""

        if first == "текущее":
            for i in range(5):
                col = i + 2
                if len(row) > col and row[col]:
                    try:
                        granula_cur[i] = float(str(row[col]))
                    except:
                        pass
            for i in range(4):
                col = i + 7
                if len(row) > col and row[col]:
                    try:
                        polu_cur[i] = float(str(row[col]))
                    except:
                        pass

        elif first == "прогноз":
            for i in range(5):
                col = i + 2
                if len(row) > col and row[col]:
                    try:
                        granula_prog[i] = float(str(row[col]))
                    except:
                        pass
            for i in range(4):
                col = i + 7
                if len(row) > col and row[col]:
                    try:
                        polu_prog[i] = float(str(row[col]))
                    except:
                        pass

    for row in прогноз_rows:
        if not row or not row[0]:
            continue
        first = str(row[0]).strip().lower()
        if "перегон" in first and len(row) > 1:
            try:
                v = float(str(row[1]))
                peregon = int(v)
            except (ValueError, IndexError):
                peregon = 0
            break

    return {
        "granula_cur": granula_cur,
        "granula_prog": granula_prog,
        "polu_cur": polu_cur,
        "polu_prog": polu_prog,
        "peregon": peregon
    }


def get_kg_col_for_sheet(name):
    """Возвращает индекс колонки кг для листа."""
    if name in ("Э-1", "Э-4"):
        return 3
    return 2  # Э-2, Э-3, Э-5


def get_kg_col_for_polu(name):
    """Возвращает индекс колонки кг для полуфабриката."""
    if name == "мойка Т":
        return 7
    return 4  # Ш-1, Ш-2, мойка Н


def generate_report_for_date(sheets_data, target_date_serial, forecast_data, sheet_name_label=None):
    """
    Генерирует отчёт для одной даты.
    """
    gc_arr = forecast_data["granula_cur"]
    gp_arr = forecast_data["granula_prog"]
    pc_arr = forecast_data["polu_cur"]
    pp_arr = forecast_data["polu_prog"]
    peregon = forecast_data["peregon"]

    # Дата отчёта
    report_date = excel_serial_to_date(target_date_serial)
    date_str = report_date.strftime("%d.%m.%Y")
    weekday_str = get_weekday_ru(report_date)

    # ── ГРАНУЛА ──
    granula_lines = []
    for name in ["Э-1", "Э-2", "Э-3", "Э-4", "Э-5"]:
        rows = sheets_data.get(name, [])
        if not rows:
            continue
        kg_col = get_kg_col_for_sheet(name)
        all_dates = extract_all_dates(rows, kg_col, ['экструдер'])
        day_v, night_v, total_v = get_data_for_date(all_dates, target_date_serial)
        if day_v == 0 and night_v == 0:
            # Нет данных на эту дату — показываем 0, а не эвристику
            granula_lines.append({"name": name, "day": 0, "night": 0, "total": 0})
        else:
            granula_lines.append({"name": name, "day": int(day_v), "night": int(night_v), "total": int(total_v)})

    # ── ПОЛУФАБРИКАТ ──
    polu_lines = []
    for name in ["Ш-1", "Ш-2", "мойка Н", "мойка Т"]:
        rows = sheets_data.get(name, [])
        if not rows:
            continue
        kg_col = get_kg_col_for_polu(name)
        title_kw = ['шредер', 'мойка']
        all_dates = extract_all_dates(rows, kg_col, title_kw)
        day_v, night_v, total_v = get_data_for_date(all_dates, target_date_serial)
        if day_v == 0 and night_v == 0:
            polu_lines.append({"name": name, "day": 0, "night": 0, "total": 0})
        else:
            polu_lines.append({"name": name, "day": int(day_v), "night": int(night_v), "total": int(total_v)})

    # ── Формирование текста ──
    lines = []
    lines.append(f"📋 Анализ производства — {date_str} ({weekday_str})")
    lines.append("")

    def name_w(lst):
        return max(len(l["name"]) for l in lst)

    # ── ГРАНУЛА (таблица) ──
    lines.append("🟢 <b>ГРАНУЛА</b>")
    nw = max(name_w(granula_lines), len("Линия"))
    all_d = [l["day"] for l in granula_lines] + [sum(l["day"] for l in granula_lines)]
    all_n = [l["night"] for l in granula_lines] + [sum(l["night"] for l in granula_lines)]
    all_t = [l["total"] for l in granula_lines] + [sum(l["total"] for l in granula_lines)]
    dw = max(len("день"), max(len(str(x)) for x in all_d))
    nw2_col = max(len("ночь"), max(len(str(x)) for x in all_n))
    tw = max(len("итого"), max(len(str(x)) for x in all_t))

    sd = sn = st = 0
    for l in granula_lines:
        sd += l["day"]; sn += l["night"]; st += l["total"]
    # Заголовок
    h = "Линия".ljust(nw) + "день".rjust(dw+1) + "ночь".rjust(nw2_col+1) + "итого".rjust(tw+1)
    code_lines = [f"<code>{h}</code><code>&lt;/&gt;</code>"]
    for l in granula_lines:
        code_lines.append(f"<code><b>{l['name'].ljust(nw)}</b>{str(l['day']).rjust(dw+1)}{str(l['night']).rjust(nw2_col+1)}{str(l['total']).rjust(tw+1)}</code>")
    code_lines.append(f"<code><b>{'ИТОГО'.ljust(nw)}</b>{str(sd).rjust(dw+1)}{str(sn).rjust(nw2_col+1)}{str(st).rjust(tw+1)}</code>")
    lines.append("\n".join(code_lines))
    lines.append("")

    # ── ПОЛУФАБРИКАТ ──
    lines.append("🔵 <b>ПОЛУФАБРИКАТ</b>")
    nw = max(name_w(polu_lines), len("Линия"))
    all_d = [l["day"] for l in polu_lines] + [sum(l["day"] for l in polu_lines)]
    all_n = [l["night"] for l in polu_lines] + [sum(l["night"] for l in polu_lines)]
    all_t = [l["total"] for l in polu_lines] + [sum(l["total"] for l in polu_lines)]
    dw = max(len("день"), max(len(str(x)) for x in all_d))
    nw2_col = max(len("ночь"), max(len(str(x)) for x in all_n))
    tw = max(len("итого"), max(len(str(x)) for x in all_t))

    sd2 = sn2 = st2 = 0
    for l in polu_lines:
        sd2 += l["day"]; sn2 += l["night"]; st2 += l["total"]
    h = "Линия".ljust(nw) + "день".rjust(dw+1) + "ночь".rjust(nw2_col+1) + "итого".rjust(tw+1)
    code_lines = [f"<code>{h}</code><code>&lt;/&gt;</code>"]
    for l in polu_lines:
        code_lines.append(f"<code><b>{l['name'].ljust(nw)}</b>{str(l['day']).rjust(dw+1)}{str(l['night']).rjust(nw2_col+1)}{str(l['total']).rjust(tw+1)}</code>")
    code_lines.append(f"<code><b>{'ИТОГО'.ljust(nw)}</b>{str(sd2).rjust(dw+1)}{str(sn2).rjust(nw2_col+1)}{str(st2).rjust(tw+1)}</code>")
    lines.append("\n".join(code_lines))
    lines.append("")

    # ── ПРОГНОЗ ──
    lines.append("📊 <b>ПРОГНОЗ НА МЕСЯЦ</b>")
    # 「Полуфабрикат」(12) сокращаю до「П/фабрикат」(10) чтобы влезть в 25 символов
    granula_names = ["Э-1", "Э-2", "Э-3", "Э-4", "Э-5"]
    polu_names = ["Ш-1", "Ш-2", "мойка Н", "мойка Т"]

    # Totals
    gc_total = int(sum(gc_arr))
    gp_total = int(sum(gp_arr))
    pc_total = int(sum(pc_arr))
    pp_total = int(sum(pp_arr))

    # Max widths
    lw = max(len("материал"),
             max(len(n) for n in granula_names),
             max(len(n) for n in polu_names),
             len("Гранула"), len("П/фабрикат"), len("Перегон"))

    # Num widths
    all_nums = []
    for i in range(5):
        all_nums += [int(gc_arr[i]), int(gp_arr[i])]
    for i in range(4):
        all_nums += [int(pc_arr[i]), int(pp_arr[i])]
    all_nums += [gc_total, gp_total, pc_total, pp_total, peregon]

    nw = max(len("текущее"), len("прогноз"),
             max(len(str(x)) for x in all_nums))
    pw = nw
    if pw + lw + 1 + pw > 25:
        pw = 7

    code_lines = []
    code_lines.append(f"<code>{'материал'.ljust(lw)}{'текущее'.rjust(pw)} прогноз</code><code>&lt;/&gt;</code>")

    # Гранула per-line
    for i in range(5):
        cv = int(gc_arr[i]) if gc_arr[i] else 0
        pv = int(gp_arr[i]) if gp_arr[i] else 0
        code_lines.append(f"<code><b>{granula_names[i].ljust(lw)}</b>{str(cv).rjust(pw)} {str(pv).rjust(pw)}</code>")

    # Гранула subtotal
    code_lines.append(f"<code><b>{'Гранула'.ljust(lw)}</b>{str(gc_total).rjust(pw)} {str(gp_total).rjust(pw)}</code>")

    # Полуфабрикат per-line
    for i in range(4):
        cv = int(pc_arr[i]) if pc_arr[i] else 0
        pv = int(pp_arr[i]) if pp_arr[i] else 0
        code_lines.append(f"<code><b>{polu_names[i].ljust(lw)}</b>{str(cv).rjust(pw)} {str(pv).rjust(pw)}</code>")

    # Полуфабрикат subtotal
    code_lines.append(f"<code><b>{'П/фабрикат'.ljust(lw)}</b>{str(pc_total).rjust(pw)} {str(pp_total).rjust(pw)}</code>")

    # Перегон (только текущее)
    code_lines.append(f"<code><b>{'Перегон'.ljust(lw)}</b>{str(peregon).rjust(pw)}</code>")

    lines.append("\n".join(code_lines))

    return "\n".join(lines)


def get_target_dates(sheets_data):
    """
    Определяет, за какие даты формировать отчёт.
    Возвращает список serial-дат.
    """
    today = datetime.now()
    weekday = today.weekday()  # 0=Mon, 6=Sun

    # Собираем все даты из Э-1
    rows = sheets_data.get("Э-1", [])
    all_dates = extract_all_dates(rows, get_kg_col_for_sheet("Э-1"), ['экструдер'])

    if not all_dates:
        # Fallback: вчера
        yesterday = today - timedelta(days=1)
        serial = (yesterday - datetime(1899, 12, 30)).days
        return [serial]

    # Все доступные serial-даты
    serials = [d[0] for d in all_dates]

    # Понедельник → 3 последних даты
    if weekday == 0:  # Monday
        return serials[-3:]

    # Остальные будни → последняя дата
    return [serials[-1]]


# ── Telegram ──────────────────────────────────────────────────────

def send_telegram(text, config):
    tg_cfg = config["telegram"]
    bt = tg_cfg["bot_token"]
    cid = tg_cfg["chat_id"]

    if "ВСТАВЬТЕ" in bt or "ВСТАВЬТЕ" in cid:
        print("⚠️  Telegram не настроен")
        return False

    MAX = 3800
    if len(text) <= MAX:
        return _tg_send(bt, cid, text)
    for i in range(0, len(text), MAX):
        part = text[i:i + MAX]
        ok = _tg_send(bt, cid, part)
        if not ok:
            return False
    return True


def _tg_send(token, chat_id, text):
    try:
        r = requests.post(
            f"https://api.telegram.org/bot{token}/sendMessage",
            json={
                "chat_id": chat_id,
                "text": text,
                "parse_mode": "HTML",
                "reply_markup": _tg_keyboard()
            },
            timeout=30
        )
        if r.status_code == 200:
            print(f"✅ Отправлено ({len(text)} символов)")
            # Save as last report for "Последний отчёт" button
            with open(LAST_REPORT_FILE, "w", encoding="utf-8") as f:
                f.write(text)
            return True
        else:
            print(f"❌ {r.status_code} {r.text[:150]}")
            return False
    except Exception as e:
        print(f"❌ {e}")
        return False


def _tg_send_document(token, chat_id, text):
    """Отправка как .txt файл — моноширинный шрифт, без переносов строк."""
    try:
        import re
        from io import BytesIO
        # Убираем HTML-теги для plain-text файла
        plain = re.sub(r'<[^>]+>', '', text)
        buf = BytesIO()
        buf.write(plain.encode('utf-8'))
        buf.seek(0)
        fname = f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        files = {'document': (fname, buf, 'text/plain; charset=utf-8')}
        r = requests.post(
            f"https://api.telegram.org/bot{token}/sendDocument",
            data={'chat_id': chat_id},
            files=files,
            timeout=60
        )
        if r.status_code == 200:
            print(f"✅ Файл отправлен ({len(text)} символов)")
            return True
        else:
            print(f"❌ {r.status_code} {r.text[:150]}")
            return False
    except Exception as e:
        print(f"❌ {e}")
        return False


def _log(msg):
    """Write to .bot_polling.log (used by polling mode)."""
    try:
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        p = Path(__file__).parent / ".bot_polling.log"
        with open(p, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {msg}\n")
    except:
        pass


def _tg_send_simple(token, chat_id, text):
    """Отправка без клавиатуры — для ответов на кнопки."""
    try:
        r = requests.post(
            f"https://api.telegram.org/bot{token}/sendMessage",
            json={"chat_id": chat_id, "text": text, "parse_mode": "HTML"},
            timeout=30
        )
        if r.status_code != 200:
            _log(f"_tg_send_simple: {r.status_code} {r.text[:200]}")
    except Exception as e:
        _log(f"_tg_send_simple ex: {e}")


# ── Polling (интерактивные кнопки) ────────────────────────────────

def run_polling():
    """Долгий polling: слушает кнопки и /start."""
    config = load_config()
    tg_cfg = config["telegram"]
    bt = tg_cfg["bot_token"]
    cid = int(tg_cfg["chat_id"])
    allowed_uid = str(tg_cfg.get("telegram_user_id", ""))

    BOT_LOG = Path(__file__).parent / ".bot_polling.log"
    _log("🤖 Polling started")
    _log(f"📢 Chat ID: {cid} | allowed_uid: {allowed_uid or '(any)'}")

    offset = 0
    while True:
        try:
            r = requests.post(
                f"https://api.telegram.org/bot{bt}/getUpdates",
                json={"offset": offset, "timeout": 30},
                timeout=35
            )
            if r.status_code != 200:
                time.sleep(5)
                continue

            updates = r.json().get("result", [])
            for update in updates:
                offset = update["update_id"] + 1

                # ── Нажатие кнопки ──
                if "callback_query" in update:
                    cb = update["callback_query"]
                    uid = str(cb["from"]["id"])
                    cdata = cb["data"]

                    _log(f"callback from={uid} data={cdata}")

                    # Гасим индикатор загрузки
                    try:
                        requests.post(
                            f"https://api.telegram.org/bot{bt}/answerCallbackQuery",
                            json={"callback_query_id": cb["id"]}
                        )
                    except Exception as e:
                        _log(f"answerCallbackQuery: {e}")

                    # Проверка прав
                    if allowed_uid and uid != allowed_uid:
                        _log(f"denied uid={uid}")
                        _tg_send_simple(bt, cid,
                            f"⛔ У вас нет прав на эту команду (ваш id: <code>{uid}</code>)")
                        continue

                    if cdata == "check_now":
                        _handle_check_now(config, bt, cid)
                    elif cdata == "last_report":
                        _handle_last_report(bt, cid)
                    elif cdata == "status":
                        _handle_status(config, bt, cid)
                    _log(f"handled: {cdata}")

                # ── Текстовое сообщение (команда /start) ──
                elif "message" in update:
                    msg = update["message"]
                    uid = str(msg["from"]["id"])
                    txt = msg.get("text", "")

                    if txt == "/start":
                        reply = (
                            f"👋 Привет! Твой Telegram ID: <code>{uid}</code>\n"
                            f"Добавь его в поле <code>telegram_user_id</code> в конфиге "
                            f"<code>email_report_config.json</code> для авторизации.\n\n"
                            f"Кнопки — быстрый запуск прямо из чата."
                        )
                        _tg_send_simple(bt, cid, reply)

        except Exception as e:
            _log(f"polling loop ex: {e}")
            time.sleep(5)


def _handle_check_now(config, bot_token, chat_id):
    """🔄 Проверить сейчас — запускает полный цикл."""
    if _is_locked():
        _tg_send_simple(bot_token, chat_id, "⏳ Проверка уже выполняется")
        return

    _acquire_lock()
    try:
        _log("check_now: started")
        _tg_send_simple(bot_token, chat_id, "⏳ Проверяю почту и формирую отчёт…")

        email_info, content, filename = get_latest_excel(config)
        if not content:
            _tg_send_simple(bot_token, chat_id,
                "📭 Нового файла «Анализ производства» в почте нет.")
            return

        suffix = Path(filename).suffix.lower()
        sheets_data = {}
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(content)
            tmp_path = tmp.name

        try:
            if suffix == ".xlsb" and HAS_PYXLSB:
                with pyxlsb.open_workbook(tmp_path) as wb:
                    for sheet_name in list(wb.sheets):
                        try:
                            with wb.get_sheet(sheet_name) as sheet:
                                rows = []
                                for row in sheet.rows():
                                    rows.append([c.v for c in row])
                                sheets_data[sheet_name] = rows
                        except:
                            pass
            else:
                xls = pd.ExcelFile(tmp_path, engine="openpyxl")
                for sheet_name in xls.sheet_names:
                    try:
                        df = pd.read_excel(xls, sheet_name=sheet_name,
                                           dtype=str, header=None)
                        sheets_data[sheet_name] = df.fillna("").values.tolist()
                    except:
                        pass

            forecast_data = extract_forecast(sheets_data)
            target_dates = get_target_dates(sheets_data)

            if not target_dates:
                _tg_send_simple(bot_token, chat_id,
                    "📭 Нет новых данных для формирования отчёта.")
                return

            target_serial = target_dates[-1]
            report = generate_report_for_date(
                sheets_data, target_serial, forecast_data)
            send_telegram(report, config)  # уже с клавиатурой

            # Сохраняем локально
            out_dir = Path(__file__).parent / "reports"
            out_dir.mkdir(exist_ok=True)
            rf = out_dir / (
                f"report_{excel_serial_to_date(target_serial).strftime('%Y%m%d')}_"
                f"{datetime.now().strftime('%H%M%S')}.txt"
            )
            with open(rf, "w", encoding="utf-8") as f:
                f.write(report)

        finally:
            try:
                os.unlink(tmp_path)
            except:
                pass

        _log("check_now: done")

    finally:
        _release_lock()


def _handle_status(config, bot_token, chat_id):
    """⚙️ Статус — проверяет работоспособность."""
    _log("_handle_status start")
    parts = []

    # 1. Бот
    parts.append("🤖 <b>Бот</b>: ✅ работает (polling)")

    # 2. Почта
    gmail_cfg = config["gmail"]
    if "ВСТАВЬТЕ" not in gmail_cfg.get("app_password", ""):
        parts.append("📧 <b>Почта</b>: ✅ настроена")
    else:
        parts.append("📧 <b>Почта</b>: ❌ не настроена")

    # 3. Telegram
    tg_cfg = config["telegram"]
    if "ВСТАВЬТЕ" not in tg_cfg.get("bot_token", ""):
        parts.append("📨 <b>Telegram</b>: ✅ настроен")
    else:
        parts.append("📨 <b>Telegram</b>: ❌ не настроен")

    if tg_cfg.get("telegram_user_id"):
        parts.append(f"🔐 <b>Авторизация</b>: ✅ включена (id: <code>{tg_cfg['telegram_user_id']}</code>)")
    else:
        parts.append("🔐 <b>Авторизация</b>: ❌ отключена (любой может нажимать кнопки)")

    # 4. Последний отчёт
    if LAST_REPORT_FILE.exists():
        mtime = datetime.fromtimestamp(LAST_REPORT_FILE.stat().st_mtime)
        ago = int((datetime.now() - mtime).total_seconds() / 60)
        parts.append(f"📊 <b>Последний отчёт</b>: {ago} мин назад ({mtime.strftime('%H:%M')})")
    else:
        parts.append("📊 <b>Последний отчёт</b>: нет")

    # 5. Блокировка
    if _is_locked():
        parts.append("⏳ <b>Проверка</b>: 🔄 выполняется сейчас")
    else:
        parts.append("⏳ <b>Проверка</b>: ✅ свободно")

    reply = "\n".join(parts)
    _log(f"_handle_status: sending reply ({len(reply)} chars)")
    _tg_send_simple(bot_token, chat_id, reply)
    _log("_handle_status done")


def _handle_last_report(bot_token, chat_id):
    """📊 Последний отчёт — пересылает последний сохранённый."""
    if not LAST_REPORT_FILE.exists():
        _log("last_report: no saved report")
        _tg_send_simple(bot_token, chat_id, "📭 Последний отчёт не найден.")
        return
    text = LAST_REPORT_FILE.read_text(encoding="utf-8")
    _log(f"last_report: resending {len(text)} chars")
    _tg_send(bot_token, chat_id, text)  # с клавиатурой


# ── Главная ───────────────────────────────────────────────────────

def main():
    print("=" * 50)
    print("📬 Анализ производства → Telegram")
    ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f"🕐 {ts}")
    print("=" * 50)

    config = load_config()

    # Шаг 1 — почта
    email_info, content, filename = get_latest_excel(config)
    if not content:
        print("❌ Файл не найден")
        send_telegram("📭 Файл «Анализ производства» не найден в почте.", config)
        return
    print(f"✅ Найден: {filename}")

    # Шаг 2 — парсинг
    suffix = Path(filename).suffix.lower()
    sheets_data = {}
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(content)
        tmp_path = tmp.name

    try:
        if suffix == ".xlsb" and HAS_PYXLSB:
            with pyxlsb.open_workbook(tmp_path) as wb:
                for sheet_name in list(wb.sheets):
                    try:
                        with wb.get_sheet(sheet_name) as sheet:
                            rows = []
                            for row in sheet.rows():
                                rows.append([c.v for c in row])
                            sheets_data[sheet_name] = rows
                    except:
                        pass
        else:
            xls = pd.ExcelFile(tmp_path, engine="openpyxl")
            for sheet_name in xls.sheet_names:
                try:
                    df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str, header=None)
                    sheets_data[sheet_name] = df.fillna("").values.tolist()
                except:
                    pass

        # Прогноз
        forecast_data = extract_forecast(sheets_data)

        # Какие даты отображать
        target_dates = get_target_dates(sheets_data)
        today = datetime.now()
        print(f"📅 Сегодня: {today.strftime('%A')}, целей: {len(target_dates)}")

        for idx, target_serial in enumerate(target_dates):
            report = generate_report_for_date(sheets_data, target_serial, forecast_data)
            label = f"[{idx+1}/{len(target_dates)}]"
            print(f"\n{label} Отчёт для даты {excel_serial_to_date(target_serial).strftime('%d.%m.%Y')}:")
            print(report)
            print(f"({len(report)} символов)")

            # Отправка в Telegram
            ok = send_telegram(report, config)
            if ok:
                print(f"{label} ✅ Отправлено")
            else:
                print(f"{label} ❌ Ошибка")

            # Сохраняем локально
            out_dir = Path(__file__).parent / "reports"
            out_dir.mkdir(exist_ok=True)
            report_file = out_dir / f"report_{excel_serial_to_date(target_serial).strftime('%Y%m%d')}_{datetime.now().strftime('%H%M%S')}.txt"
            with open(report_file, "w", encoding="utf-8") as f:
                f.write(report)
            print(f"💾 {report_file}")

    finally:
        try:
            os.unlink(tmp_path)
        except:
            pass

    print("\n✅ Готово!")


if __name__ == "__main__":
    if "--polling" in sys.argv:
        run_polling()
    else:
        main()
