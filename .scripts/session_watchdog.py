#!/usr/bin/env python3
"""
Hermes Session Watchdog — no_agent скрипт.

Проверяет активные сессии Hermes. Если сессия превысила порог сообщений,
создаёт handoff-файл для переноса в новый чат.

Watchdog pattern: если всё ок — stdout пуст (тихий режим).
Если создан хотя бы один handoff — пишет уведомление.
"""
import sqlite3
import os
import json
import time
from datetime import datetime

# === КОНФИГУРАЦИЯ ===
STATE_DB = os.path.expanduser("~/AppData/Local/hermes/state.db")
WORKSPACE = os.path.expanduser("~/workspace")
THRESHOLD = 100          # сообщений — порог для handoff
REGISTRY_FILE = os.path.join(WORKSPACE, ".handoff_registry.json")
# ====================

def load_registry():
    if os.path.exists(REGISTRY_FILE):
        with open(REGISTRY_FILE, "r") as f:
            return json.load(f)
    return {"handoffs": []}

def save_registry(registry):
    with open(REGISTRY_FILE, "w") as f:
        json.dump(registry, f, indent=2, ensure_ascii=False)

def get_next_number(registry):
    numbers = [h.get("number", 0) for h in registry["handoffs"]]
    return max(numbers + [0]) + 1

def get_session_title(cur, session_id):
    """Пробуем получить тему из первых сообщений сессии."""
    cur.execute(
        "SELECT content FROM messages WHERE session_id=? AND role='user' AND content IS NOT NULL ORDER BY id LIMIT 3",
        (session_id,)
    )
    messages = cur.fetchall()
    if messages:
        for m in messages:
            text = m[0].strip()[:80]
            if text:
                return text
    return "Без названия"

def main():
    if not os.path.exists(STATE_DB):
        return  # тихо

    conn = sqlite3.connect(STATE_DB)
    cur = conn.cursor()
    registry = load_registry()
    existing_ids = {h["session_id"] for h in registry["handoffs"]}
    created = []

    # Ищем активные сессии (без ended_at) с превышением порога
    # Берём ТОЛЬКО САМУЮ длинную за раз, чтобы не плодить кучу файлов
    cur.execute(
        "SELECT id, display_name, message_count, source, model FROM sessions "
        "WHERE ended_at IS NULL AND message_count > ? "
        "ORDER BY message_count DESC LIMIT 1",
        (THRESHOLD,)
    )
    sessions = cur.fetchall()

    for sid, display_name, msg_count, source, model in sessions:
        if sid in existing_ids:
            continue  # уже есть handoff

        # Получаем примерную тему
        topic = get_session_title(cur, sid)

        # Номер handoff
        num = get_next_number(registry)

        # Имя файла
        fname = f"handoff-{num}.md"
        fpath = os.path.join(WORKSPACE, fname)

        # Контент handoff-файла
        content = f"""# Handoff-{num}: {topic[:60]}

**Сессия:** `{sid}`
**Источник:** {source}
**Модель:** {model or '?'}
**Сообщений:** {msg_count}
**Создан:** {datetime.now().strftime('%Y-%m-%d %H:%M')}

---

## Контекст

Автоматически созданный handoff для переноса длинной сессии.

- **Рабочая директория:** `{WORKSPACE}`
- **Сообщений на момент переноса:** {msg_count}

## Что делать

1. Начать новый чат в WebUI
2. Написать: `Прочитай {fname} и продолжи. Название чата: «{topic[:40]} → продолжение»`
3. Продолжить с того же места

---

*Автоматически создано session_watchdog.py*
"""

        with open(fpath, "w", encoding="utf-8") as f:
            f.write(content)

        # Регистрируем
        entry = {
            "number": num,
            "session_id": sid,
            "display_name": display_name or topic[:40],
            "message_count": msg_count,
            "file": fname,
            "created_at": datetime.now().isoformat()
        }
        registry["handoffs"].append(entry)
        created.append(entry)

    if created:
        save_registry(registry)
        lines = []
        for c in created:
            lines.append(f"📄 {c['file']} — сессия {c['session_id'][:12]}... ({c['message_count']} сообщений)")
        print("🌀 **Hermes: авто-handoff**")
        for l in lines:
            print(l)
        print("\nНачни новый чат и напиши: «Прочитай <файл> и продолжи»")
    else:
        # тихо — watchdog pattern, ничего не выводим
        pass

    conn.close()

if __name__ == "__main__":
    main()
