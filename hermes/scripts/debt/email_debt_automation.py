#!/usr/bin/env python3
"""
Автоматизация: Почта (Gmail) → Excel «Задолженность» → Отчёт → Telegram

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
from datetime import datetime
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
LOCK_FILE = Path(__file__).parent / ".bot_lock_debt"
LAST_REPORT_FILE = Path(__file__).parent / ".last_report_debt.txt"

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

CONFIG_PATH = Path(__file__).parent / "email_debt_config.json"

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
                "filename_keyword": "задолженность",
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

def generate_debt_report(sheets_data):
    """
    Генерирует отчёт по задолженности из Excel.
    Формат файла:
      Row 0-3: заголовки
      Row 4..N-1: данные (наименование, долг, аванс)
      Row N: Итого
    Отчёт: Долг (сорт.↓) + Итого, Аванс (сорт.↓) + Итого
    """
    lines = []
    lines.append("📋 <b>ЗАДОЛЖЕННОСТЬ</b>")
    lines.append("")

    for sheet_name, rows in sheets_data.items():
        if not rows or len(rows) < 5:
            continue

        # Ищем строку с датой (строка 1)
        date_str = ""
        for r in rows:
            if r and r[0]:
                txt = str(r[0]).strip()
                if "задолженность" in txt.lower() and "за " in txt.lower():
                    # Extract date like "24 июля 2026 г."
                    import re
                    m = re.search(r'за\s+(.+?)(?:\.\s*г|г)?\.?\s*$', txt)
                    if m:
                        date_str = m.group(1).strip()
                    break

        # Ищем строку с "Итого"
        total_debt = 0
        total_adv = 0
        for r in rows:
            if r and r[0] is not None and not (isinstance(r[0], float) and str(r[0]) == 'nan'):
                if str(r[0]).strip().lower() == "итого":
                    try:
                        if len(r) > 1 and r[1] is not None and not (isinstance(r[1], float) and str(r[1]) == 'nan'):
                            total_debt = int(float(str(r[1])))
                    except: pass
                    try:
                        if len(r) > 2 and r[2] is not None and not (isinstance(r[2], float) and str(r[2]) == 'nan'):
                            total_adv = int(float(str(r[2])))
                    except: pass
                    break

        # Собираем данные (после строки с "Долг" как маркер начала)
        debt_items = []
        adv_items = []
        started = False
        for r in rows:
            if not r:
                continue
            first = str(r[0]).strip() if r[0] is not None and not (isinstance(r[0], float) and str(r[0]) == 'nan') else ""
            if first.lower() == "итого":
                break
            # Ищем "Долг" в любой колонке
            if not started:
                for cell in r:
                    if cell is not None and isinstance(cell, str) and "долг" in cell.strip().lower():
                        started = True
                        break
                if started:
                    continue

            if not started:
                continue

            # Это строка данных
            name = first
            # Убираем организационно-правовые формы
            for suffix in [" ООО", " ЗАО", " НАО", " АО", " ПАО"]:
                if name.upper().endswith(suffix):
                    name = name[:-len(suffix)].strip()
                    break
            if not name:
                continue

            debt_val = 0
            adv_val = 0
            try:
                if len(r) > 1 and r[1] is not None and not (isinstance(r[1], float) and str(r[1]) == 'nan'):
                    debt_val = int(float(str(r[1])))
            except: pass
            try:
                if len(r) > 2 and r[2] is not None and not (isinstance(r[2], float) and str(r[2]) == 'nan'):
                    adv_val = int(float(str(r[2])))
            except: pass

            if debt_val > 0:
                debt_items.append((name, debt_val))
            if adv_val > 0:
                adv_items.append((name, adv_val))

        if not debt_items and not adv_items:
            continue

        # Сортировка по убыванию
        debt_items.sort(key=lambda x: x[1], reverse=True)
        adv_items.sort(key=lambda x: x[1], reverse=True)

        # Заголовок с датой
        if date_str:
            lines.append(f"📅 <b>за {date_str}</b>")
            lines.append("")

        # ── ДОЛГ ──
        lines.append("💰 <b>ЗАДОЛЖЕННОСТЬ ПОКУПАТЕЛЕЙ</b>")

        # Вычисляем ширину колонок
        name_w = max(len(n) for n, v in debt_items)
        if name_w > 24:
            name_w = 24
        name_w = max(name_w, len("Наименование"))

        # Ширина для чисел
        all_vals = [v for n, v in debt_items] + [total_debt]
        num_w = max(len(f"{v:,}".replace(",", " ")) for v in all_vals) if all_vals else 10
        num_w = max(num_w, len("Долг"))

        pre_lines = []
        # Заголовок таблицы
        pre_lines.append(
            f"<code>{'Наименование'.ljust(name_w)}</code>{'Долг'.rjust(num_w)}"
        )
        # Разделитель
        sep = "─" * (name_w + 1 + num_w)
        pre_lines.append(sep)

        # Данные
        for name, val in debt_items:
            if len(name) > name_w:
                name = name[:name_w-3] + "..."
            val_str = f"{val:,}".replace(",", " ")
            pre_lines.append(
                f"<code>{name.ljust(name_w)}</code>{val_str.rjust(num_w)}"
            )

        # Итого
        total_str = f"{total_debt:,}".replace(",", " ")
        pre_lines.append(sep)
        pre_lines.append(
            f"<code>{'ИТОГО'.ljust(name_w)}</code>{total_str.rjust(num_w)}"
        )

        lines.append("<pre>" + "\n".join(pre_lines) + "</pre>")
        lines.append("")

        # ── АВАНСЫ ──
        if adv_items:
            lines.append("💳 <b>АВАНСЫ</b>")

            name_w2 = max(len(n) for n, v in adv_items)
            if name_w2 > 24:
                name_w2 = 24
            name_w2 = max(name_w2, len("Наименование"))

            all_vals2 = [v for n, v in adv_items] + [total_adv]
            num_w2 = max(len(f"{v:,}".replace(",", " ")) for v in all_vals2) if all_vals2 else 10
            num_w2 = max(num_w2, len("Аванс"))

            pre_lines2 = []
            pre_lines2.append(
                f"<code>{'Наименование'.ljust(name_w2)}</code>{'Аванс'.rjust(num_w2)}"
            )
            sep2 = "─" * (name_w2 + 1 + num_w2)
            pre_lines2.append(sep2)

            for name, val in adv_items:
                if len(name) > name_w2:
                    name = name[:name_w2-3] + "..."
                val_str = f"{val:,}".replace(",", " ")
                pre_lines2.append(
                    f"<code>{name.ljust(name_w2)}</code>{val_str.rjust(num_w2)}"
                )

            total_str2 = f"{total_adv:,}".replace(",", " ")
            pre_lines2.append(sep2)
            pre_lines2.append(
                f"<code>{'ИТОГО'.ljust(name_w2)}</code>{total_str2.rjust(num_w2)}"
            )

            lines.append("<pre>" + "\n".join(pre_lines2) + "</pre>")

    if len(lines) <= 2:
        return "📭 Данные по задолженности не найдены."
    return "\n".join(lines)


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
                "📭 Нового файла «Задолженность» в почте нет.")
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
                # .xls → xlrd, .xlsx → openpyxl
                engine = "xlrd" if suffix == ".xls" else "openpyxl"
                xls = pd.ExcelFile(tmp_path, engine=engine)
                for sheet_name in xls.sheet_names:
                    try:
                        df = pd.read_excel(xls, sheet_name=sheet_name,
                                           dtype=str, header=None)
                        sheets_data[sheet_name] = df.fillna("").values.tolist()
                    except:
                        pass

            report = generate_debt_report(sheets_data)
            if "не найдены" in report:
                _tg_send_simple(bot_token, chat_id, report)
                return
            send_telegram(report, config)  # уже с клавиатурой

            # Сохраняем локально
            out_dir = Path(__file__).parent / "reports"
            out_dir.mkdir(exist_ok=True)
            rf = out_dir / (
                f"debt_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
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
    print("📬 Задолженность → Telegram")
    ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f"🕐 {ts}")
    print("=" * 50)

    config = load_config()

    # Шаг 1 — почта
    email_info, content, filename = get_latest_excel(config)
    if not content:
        print("❌ Файл не найден")
        send_telegram("📭 Файл «Задолженность» не найден в почте.", config)
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
            # .xls → xlrd, .xlsx → openpyxl
            engine = "xlrd" if suffix == ".xls" else "openpyxl"
            xls = pd.ExcelFile(tmp_path, engine=engine)
            for sheet_name in xls.sheet_names:
                try:
                    df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str, header=None)
                    sheets_data[sheet_name] = df.fillna("").values.tolist()
                except:
                    pass

        # Генерация отчёта по задолженности
        report = generate_debt_report(sheets_data)
        print(f"\nОтчёт ({len(report)} символов):")
        print(report)

        # Отправка в Telegram
        ok = send_telegram(report, config)
        if ok:
            print("✅ Отправлено")
        else:
            print("❌ Ошибка отправки")

        # Сохраняем локально
        out_dir = Path(__file__).parent / "reports"
        out_dir.mkdir(exist_ok=True)
        report_file = out_dir / f"debt_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
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
