# ────────────────────────── raporty_bot.py (refactor 2025-08) ──────────────────────────
import os
import json
import logging
from datetime import datetime
from typing import Dict, List, Optional, Tuple

from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet

# ───────────── SharePoint (opcjonalny upload) ─────────────
try:
    from office365.sharepoint.client_context import ClientContext
    from office365.runtime.auth.client_credential import ClientCredential
except ModuleNotFoundError:
    ClientContext = ClientCredential = None  # brak biblioteki → upload pomijamy

# ───────────── Telegram ─────────────
from telegram import (
    Update,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    BotCommand,
)
from telegram.ext import (
    ApplicationBuilder,
    Application,
    CommandHandler,
    CallbackQueryHandler,
    MessageHandler,
    ContextTypes,
    ConversationHandler,
    filters,
)

# ──────────────────── konfiguracja ────────────────────
load_dotenv()

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")  # MUST HAVE
WEBHOOK_URL = os.getenv("WEBHOOK_URL", "").rstrip("/")  # np. https://app-xyz.northflank.app
PORT = int(os.getenv("PORT", 8080))  # Northflank zwykle ekspozycja na 8080

# opcjonalne ustawienia SharePoint
SHAREPOINT_SITE = os.getenv("SHAREPOINT_SITE")
SHAREPOINT_DOC_LIB = os.getenv("SHAREPOINT_DOC_LIB")
SHAREPOINT_CLIENT_ID = os.getenv("SHAREPOINT_CLIENT_ID")
SHAREPOINT_CLIENT_SECRET = os.getenv("SHAREPOINT_CLIENT_SECRET")

EXCEL_FILE = "reports.xlsx"
MAPPING_FILE = "report_msgs.json"

# ──────────────────── stałe excela ────────────────────
HEADERS = [
    "ID",        # unikalny klucz: {user_id}_{dd.mm.YYYY}_{idx}
    "Data",
    "Imię",      # <- zamiast "Osoba"
    "Miejsce",
    "Start",
    "Koniec",
    "Zadania",
    "Uwagi",
]
COLS = {name: i + 1 for i, name in enumerate(HEADERS)}  # 1-based indexy kolumn

# ──────────────────── stany konwersacji ────────────────────
PLACE, START_TIME, END_TIME, TASKS, NOTES, ANOTHER = range(6)
# stany edycji
SELECT_ENTRY, SELECT_FIELD, EDIT_VALUE, EDIT_MORE = range(6, 10)

# ──────────────────── funkcje pomocnicze ────────────────────
def load_mapping() -> Dict[str, int]:
    if os.path.exists(MAPPING_FILE):
        with open(MAPPING_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_mapping(mapping: Dict[str, int]) -> None:
    with open(MAPPING_FILE, "w", encoding="utf-8") as f:
        json.dump(mapping, f)


def month_key_from_date(date_str: str) -> str:
    # date_str format: dd.mm.YYYY
    d = datetime.strptime(date_str, "%d.%m.%Y")
    return f"{d.year:04d}-{d.month:02d}"


def ensure_month_sheet(wb: Workbook, month_key: str) -> Worksheet:
    """Zwraca arkusz dla danego miesiąca. Tworzy jeśli nie istnieje i ustawia nagłówki.
    Nowy miesiąc ląduje jako pierwszy arkusz (na górze)."""
    ws: Optional[Worksheet] = wb[month_key] if month_key in wb.sheetnames else None
    if ws is None:
        ws = wb.create_sheet(title=month_key, index=0)  # na górze
        ws.append(HEADERS)
        # jeżeli to pierwszy tworzony arkusz i w zeszycie jest domyślny "Sheet", usuń go
        if "Sheet" in wb.sheetnames and wb["Sheet"].max_row == 1 and wb["Sheet"].max_column == 1:
            std = wb["Sheet"]
            wb.remove(std)
    else:
        # przenieś na początek (góra), jeśli trzeba
        idx = wb.sheetnames.index(month_key)
        if idx != 0:
            wb.move_sheet(ws, offset=-idx)
    return ws


def open_wb() -> Workbook:
    if os.path.exists(EXCEL_FILE):
        return load_workbook(EXCEL_FILE)
    wb = Workbook()
    # zostawimy pusty zeszyt, pierwszy prawdziwy arkusz powstanie przy pierwszym zapisie
    return wb


def report_exists(user_id: int, date: str) -> bool:
    if not os.path.exists(EXCEL_FILE):
        return False
    wb = load_workbook(EXCEL_FILE)
    ws = ensure_month_sheet(wb, month_key_from_date(date))
    prefix = f"{user_id}_{date}_"
    for row in ws.iter_rows(min_row=2, values_only=True):
        if str(row[0]).startswith(prefix):
            return True
    return False


def parse_time(text: str) -> Optional[str]:
    try:
        t = datetime.strptime(text.strip(), "%H:%M")
        return t.strftime("%H:%M")
    except Exception:
        return None


def save_report(entries: List[Dict[str, str]], user_id: int, date: str, name: str) -> None:
    """Zapisuje NOWE wpisy na dany dzień (append). Nie kasuje istniejących wierszy.
    Edycja pola istniejącego wpisu odbywa się funkcją update_report_field()."""
    wb = open_wb()
    ws = ensure_month_sheet(wb, month_key_from_date(date))

    # policz ile pozycji już istnieje dla tego usera i daty, żeby nadać kolejne indeksy
    prefix = f"{user_id}_{date}_"
    existing_idxs = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        rid = str(row[0])
        if rid.startswith(prefix):
            try:
                existing_idxs.append(int(rid.split("_")[-1]))
            except Exception:
                pass
    next_idx = (max(existing_idxs) + 1) if existing_idxs else 1

    for idx_offset, e in enumerate(entries, start=0):
        idx = next_idx + idx_offset
        ws.append([
            f"{user_id}_{date}_{idx}",
            date,
            name,
            e["place"],
            e["start"],
            e["end"],
            e["tasks"],
            e["notes"],
        ])

    wb.save(EXCEL_FILE)
    _maybe_upload_sharepoint()


def _maybe_upload_sharepoint() -> None:
    if all([ClientContext, SHAREPOINT_SITE, SHAREPOINT_DOC_LIB, SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET]):
        try:
            ctx = ClientContext(SHAREPOINT_SITE).with_credentials(
                ClientCredential(SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET)
            )
            folder = ctx.web.get_folder_by_server_relative_url(SHAREPOINT_DOC_LIB)
            with open(EXCEL_FILE, "rb") as f:
                folder.upload_file(os.path.basename(EXCEL_FILE), f).execute_query()
        except Exception as e:
            logging.warning("SharePoint upload failed: %s", e)


def read_entries_for_day(user_id: int, date: str) -> List[Dict[str, str]]:
    if not os.path.exists(EXCEL_FILE):
        return []
    wb = load_workbook(EXCEL_FILE)
    ws = ensure_month_sheet(wb, month_key_from_date(date))
    prefix = f"{user_id}_{date}_"
    out: List[Dict[str, str]] = []
    for row in ws.iter_rows(min_row=2, values_only=False):
        rid = str(row[0].value)
        if rid and rid.startswith(prefix):
            out.append(
                {
                    "rid": rid,
                    "row": row[0].row,  # numer wiersza w arkuszu
                    "date": row[COLS["Data"] - 1].value,
                    "name": row[COLS["Imię"] - 1].value,
                    "place": row[COLS["Miejsce"] - 1].value or "",
                    "start": row[COLS["Start"] - 1].value or "",
                    "end": row[COLS["Koniec"] - 1].value or "",
                    "tasks": row[COLS["Zadania"] - 1].value or "",
                    "notes": row[COLS["Uwagi"] - 1].value or "",
                }
            )
    # sortuj po indeksie wpisu rosnąco
    out.sort(key=lambda e: int(e["rid"].split("_")[-1]))
    return out


def update_report_field(user_id: int, date: str, rid: str, field: str, new_value: str) -> None:
    """Nadpisuje dokładnie jedną komórkę (bez usuwania/insertów)."""
    wb = load_workbook(EXCEL_FILE)
    ws = ensure_month_sheet(wb, month_key_from_date(date))

    col_name_map = {
        "place": "Miejsce",
        "start": "Start",
        "end": "Koniec",
        "tasks": "Zadania",
        "notes": "Uwagi",
    }
    target_col = COLS[col_name_map[field]]

    # wyszukaj wiersz po RID
    target_row = None
    for row in ws.iter_rows(min_row=2, values_only=False):
        if str(row[0].value) == rid:
            target_row = row[0].row
            break
    if not target_row:
        raise RuntimeError("Nie znaleziono wiersza do edycji.")

    ws.cell(row=target_row, column=target_col, value=new_value)
    wb.save(EXCEL_FILE)
    _maybe_upload_sharepoint()


def format_report(entries: List[Dict[str, str]], date: str, name: str) -> str:
    lines = [f"📄 Raport dzienny – {date}", f"👤 Imię: {name}", ""]
    for i, e in enumerate(entries, start=1):
        lines.extend(
            [
                f"#{i}",
                f"📍 Miejsce: {e['place']}",
                f"⏰ {e['start']} – {e['end']}",
                "📝 Zadania:",
                str(e["tasks"]) or "-",
                "💬 Uwagi:",
                str(e["notes"]) or "-",
                "",
            ]
        )
    return "\n".join(lines)


# ──────────────────── handlery ────────────────────
async def show_menu(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    context.user_data.clear()
    context.user_data["msg_ids"] = []
    date = datetime.now().strftime("%d.%m.%Y")
    uid = update.effective_user.id

    create_text = "📋 Stwórz raport" if not report_exists(uid, date) else "✏️ Edytuj raport"
    cb_data = "create" if not report_exists(uid, date) else "edit"

    kb = [
        [InlineKeyboardButton(create_text, callback_data=cb_data)],
        [InlineKeyboardButton("📥 Eksportuj", callback_data="export")],
    ]

    m = await update.effective_chat.send_message("Wybierz opcję:", reply_markup=InlineKeyboardMarkup(kb))
    context.user_data["msg_ids"].append(m.message_id)


async def export_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.callback_query:
        await update.callback_query.answer()
        chat_id = update.callback_query.message.chat.id
        await update.callback_query.edit_message_reply_markup(reply_markup=None)
    else:
        chat_id = update.effective_chat.id

    if not os.path.exists(EXCEL_FILE):
        msg = await context.bot.send_message(chat_id, "⚠️ Brak pliku raportów.")
        context.user_data.setdefault("msg_ids", []).append(msg.message_id)
    else:
        with open(EXCEL_FILE, "rb") as f:
            doc = await context.bot.send_document(
                chat_id, f, filename=EXCEL_FILE, caption="Arkusz z raportami (miesięczne arkusze u góry)"
            )
            context.user_data.setdefault("msg_ids", []).append(doc.message_id)
    return ConversationHandler.END


async def menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data

    if data == "export":
        return await export_handler(update, context)

    date = datetime.now().strftime("%d.%m.%Y")
    context.user_data.update({
        "date": date,
        "name": query.from_user.first_name,
        "uid": query.from_user.id,
        "msg_ids": context.user_data.get("msg_ids", []),
    })

    await query.edit_message_reply_markup(reply_markup=None)

    if data == "create":
        # start nowego raportu (append)
        msg = await query.message.chat.send_message("📍 Podaj miejsce wykonywania pracy:")
        context.user_data["msg_ids"].append(msg.message_id)
        return PLACE

    # tryb edycji: wybór wpisu
    entries = read_entries_for_day(query.from_user.id, date)
    if not entries:
        kb = [[InlineKeyboardButton("📋 Stwórz raport", callback_data="create")]]
        msg = await query.message.chat.send_message(
            "Nie znaleziono dzisiejszego raportu. Chcesz stworzyć nowy?", reply_markup=InlineKeyboardMarkup(kb)
        )
        context.user_data["msg_ids"].append(msg.message_id)
        return ConversationHandler.END

    context.user_data["edit_entries"] = entries
    # usuń poprzednie menu
    await query.edit_message_text("Wybierz pozycję do edycji:")

    kb_rows = []
    for idx, e in enumerate(entries, start=1):
        label = f"#{idx} {e['place']}  {e['start']}-{e['end']}"
        kb_rows.append([InlineKeyboardButton(label, callback_data=f"entry:{idx-1}")])
    kb_rows.append([InlineKeyboardButton("↩️ Anuluj", callback_data="cancel_edit")])

    msg = await query.message.chat.send_message("Wybierz pozycję:", reply_markup=InlineKeyboardMarkup(kb_rows))
    context.user_data["msg_ids"].append(msg.message_id)
    return SELECT_ENTRY


# ────────────── FLOW: CREATE (bez zmian poza nagłówkami/Imię) ──────────────
async def place(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["msg_ids"].append(update.message.message_id)
    await update.message.delete()
    place_txt = (update.message.text or "").strip()

    if not place_txt:
        err = await update.effective_chat.send_message("Podaj poprawne miejsce.")
        context.user_data["msg_ids"].append(err.message_id)
        return PLACE

    context.user_data["place"] = place_txt
    ask = await update.effective_chat.send_message("⏰ Podaj godzinę rozpoczęcia pracy (HH:MM):")
    context.user_data["msg_ids"].append(ask.message_id)
    return START_TIME


async def start_time(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["msg_ids"].append(update.message.message_id)
    await update.message.delete()
    t = parse_time(update.message.text or "")
    last_end = context.user_data["entries"][-1]["end"] if context.user_data.get("entries") else None

    if not t or (last_end and t <= last_end):
        err = await update.effective_chat.send_message("⏰ Błędna godzina. Spróbuj ponownie.")
        context.user_data["msg_ids"].append(err.message_id)
        return START_TIME

    context.user_data["start"] = t
    ask = await update.effective_chat.send_message("⏰ Podaj godzinę zakończenia pracy (HH:MM):")
    context.user_data["msg_ids"].append(ask.message_id)
    return END_TIME


async def end_time(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["msg_ids"].append(update.message.message_id)
    await update.message.delete()
    t = parse_time(update.message.text or "")

    if not t or t <= context.user_data.get("start"):
        err = await update.effective_chat.send_message("⏰ Błędna godzina. Spróbuj ponownie.")
        context.user_data["msg_ids"].append(err.message_id)
        return END_TIME

    context.user_data["end"] = t
    ask = await update.effective_chat.send_message("📝 Opisz wykonane prace:")
    context.user_data["msg_ids"].append(ask.message_id)
    return TASKS


async def tasks(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["msg_ids"].append(update.message.message_id)
    await update.message.delete()
    txt = (update.message.text or "").strip()

    if not txt:
        err = await update.effective_chat.send_message("📝 Lista zadań nie może być pusta.")
        context.user_data["msg_ids"].append(err.message_id)
        return TASKS

    context.user_data["tasks"] = txt
    ask = await update.effective_chat.send_message("💬 Dodaj uwagi lub wpisz '-' jeśli brak:")
    context.user_data["msg_ids"].append(ask.message_id)
    return NOTES


async def notes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["msg_ids"].append(update.message.message_id)
    await update.message.delete()
    txt = (update.message.text or "").strip()

    if not txt:
        err = await update.effective_chat.send_message("💬 Uwagi nie mogą być puste.")
        context.user_data["msg_ids"].append(err.message_id)
        return NOTES

    entry = {
        "place": context.user_data.pop("place"),
        "start": context.user_data.pop("start"),
        "end": context.user_data.pop("end"),
        "tasks": context.user_data.pop("tasks"),
        "notes": txt,
    }
    context.user_data.setdefault("entries", []).append(entry)

    kb = [
        [InlineKeyboardButton("➕ Dodaj kolejne miejsce", callback_data="again")],
        [InlineKeyboardButton("✅ Zakończ raport", callback_data="finish")],
    ]
    msg = await update.effective_chat.send_message("Co dalej?", reply_markup=InlineKeyboardMarkup(kb))
    context.user_data["msg_ids"].append(msg.message_id)
    return ANOTHER


async def another(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == "again":
        await query.edit_message_reply_markup()
        ask = await query.message.chat.send_message("📍 Podaj miejsce wykonywania pracy:")
        context.user_data["msg_ids"].append(ask.message_id)
        return PLACE

    # finish → usuń prompty
    for mid in context.user_data.get("msg_ids", []):
        try:
            await query.message.chat.delete_message(mid)
        except Exception:
            pass

    # zapisz do excela
    save_report(
        context.user_data.get("entries", []),
        context.user_data["uid"] if "uid" in context.user_data else query.from_user.id,
        context.user_data["date"],
        context.user_data.get("name", query.from_user.first_name),
    )

    # odczytaj i wyświetl końcowy raport (już z pliku)
    final_entries = read_entries_for_day(query.from_user.id, context.user_data["date"])
    rpt = format_report(final_entries, context.user_data["date"], context.user_data.get("name", query.from_user.first_name))

    msg = await query.message.chat.send_message(rpt)

    mapping = load_mapping()
    mapping[f"{query.from_user.id}_{context.user_data['date']}"] = msg.message_id
    save_mapping(mapping)
    return ConversationHandler.END


# ────────────── FLOW: EDIT ──────────────
async def select_entry(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == "cancel_edit":
        await query.edit_message_reply_markup(reply_markup=None)
        await query.message.edit_text("Anulowano edycję.")
        return ConversationHandler.END

    # entry:{idx}
    idx = int(query.data.split(":")[1])
    context.user_data["edit_idx"] = idx
    e = context.user_data["edit_entries"][idx]

    await query.edit_message_reply_markup(reply_markup=None)
    await query.message.edit_text(
        f"Wybrano: #{idx+1} {e['place']} {e['start']}-{e['end']}\nCo edytować?"
    )

    kb = [
        [InlineKeyboardButton("Miejsce", callback_data="field:place")],
        [InlineKeyboardButton("Godzina start", callback_data="field:start")],
        [InlineKeyboardButton("Godzina koniec", callback_data="field:end")],
        [InlineKeyboardButton("Zadania", callback_data="field:tasks")],
        [InlineKeyboardButton("Uwagi", callback_data="field:notes")],
        [InlineKeyboardButton("↩️ Wybierz inną pozycję", callback_data="back_to_entries")],
    ]
    msg = await query.message.chat.send_message("Wybierz pole:", reply_markup=InlineKeyboardMarkup(kb))
    context.user_data.setdefault("msg_ids", []).append(msg.message_id)
    return SELECT_FIELD


async def select_field(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == "back_to_entries":
        # pokaż ponownie listę wpisów
        entries = context.user_data.get("edit_entries", [])
        kb_rows = []
        for idx, e in enumerate(entries, start=1):
            label = f"#{idx} {e['place']}  {e['start']}-{e['end']}"
            kb_rows.append([InlineKeyboardButton(label, callback_data=f"entry:{idx-1}")])
        kb_rows.append([InlineKeyboardButton("↩️ Anuluj", callback_data="cancel_edit")])
        await query.edit_message_text("Wybierz pozycję:")
        msg = await query.message.chat.send_message("Wybierz pozycję:", reply_markup=InlineKeyboardMarkup(kb_rows))
        context.user_data["msg_ids"].append(msg.message_id)
        return SELECT_ENTRY

    field = query.data.split(":")[1]
    context.user_data["edit_field"] = field

    prompt = {
        "place": "📍 Podaj nowe *miejsce*:",
        "start": "⏰ Podaj *nową godzinę start* (HH:MM):",
        "end": "⏰ Podaj *nową godzinę koniec* (HH:MM):",
        "tasks": "📝 Podaj nowe *zadania*:",
        "notes": "💬 Podaj nowe *uwagi* (lub '-' jeśli brak):",
    }[field]

    await query.edit_message_reply_markup(reply_markup=None)
    ask = await query.message.chat.send_message(prompt, parse_mode="Markdown")
    context.user_data["msg_ids"].append(ask.message_id)
    return EDIT_VALUE


async def edit_value(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # zbierz i skasuj wiadomość użytkownika
    context.user_data["msg_ids"].append(update.message.message_id)
    await update.message.delete()

    val = (update.message.text or "").strip()
    field = context.user_data.get("edit_field")
    idx = context.user_data.get("edit_idx")
    date = context.user_data.get("date")
    uid = context.user_data.get("uid")

    entries = context.user_data.get("edit_entries", [])
    e = entries[idx]

    # walidacja czasu
    if field in ("start", "end"):
        t = parse_time(val)
        if not t:
            err = await update.effective_chat.send_message("⏰ Błędny format. Użyj HH:MM.")
            context.user_data["msg_ids"].append(err.message_id)
            return EDIT_VALUE
        # sprawdź relację start < end
        start = t if field == "start" else str(e["start"]) or t
        end = t if field == "end" else str(e["end"]) or t
        if start and end and start >= end:
            err = await update.effective_chat.send_message("⏰ Start musi być < koniec.")
            context.user_data["msg_ids"].append(err.message_id)
            return EDIT_VALUE
        val = t

    # aktualizacja w excelu
    try:
        update_report_field(uid, date, e["rid"], field, val)
    except Exception as ex:
        err = await update.effective_chat.send_message(f"❌ Błąd zapisu: {ex}")
        context.user_data["msg_ids"].append(err.message_id)
        return EDIT_VALUE

    # odśwież lokalny cache wpisów (żeby kolejne edycje operowały na aktualnych danych)
    context.user_data["edit_entries"] = read_entries_for_day(uid, date)

    kb = [
        [InlineKeyboardButton("Edytuj inne pole tej pozycji", callback_data="again_same")],
        [InlineKeyboardButton("Edytuj inną pozycję", callback_data="again_other")],
        [InlineKeyboardButton("Pokaż raport i zakończ", callback_data="finish_edit")],
    ]
    msg = await update.effective_chat.send_message("Zmieniono. Edytować coś jeszcze?", reply_markup=InlineKeyboardMarkup(kb))
    context.user_data["msg_ids"].append(msg.message_id)
    return EDIT_MORE


async def edit_more(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == "again_same":
        # pokaż wybór pola dla tej samej pozycji
        idx = context.user_data.get("edit_idx")
        e = context.user_data.get("edit_entries", [])[idx]
        await query.edit_message_reply_markup(reply_markup=None)
        await query.message.chat.send_message(
            f"Wybrano: #{idx+1} {e['place']} {e['start']}-{e['end']}\nCo edytować?"
        )
        kb = [
            [InlineKeyboardButton("Miejsce", callback_data="field:place")],
            [InlineKeyboardButton("Godzina start", callback_data="field:start")],
            [InlineKeyboardButton("Godzina koniec", callback_data="field:end")],
            [InlineKeyboardButton("Zadania", callback_data="field:tasks")],
            [InlineKeyboardButton("Uwagi", callback_data="field:notes")],
            [InlineKeyboardButton("↩️ Wybierz inną pozycję", callback_data="back_to_entries")],
        ]
        msg = await query.message.chat.send_message("Wybierz pole:", reply_markup=InlineKeyboardMarkup(kb))
        context.user_data.setdefault("msg_ids", []).append(msg.message_id)
        return SELECT_FIELD

    if query.data == "again_other":
        # lista pozycji
        entries = context.user_data.get("edit_entries", [])
        kb_rows = []
        for idx, e in enumerate(entries, start=1):
            label = f"#{idx} {e['place']}  {e['start']}-{e['end']}"
            kb_rows.append([InlineKeyboardButton(label, callback_data=f"entry:{idx-1}")])
        kb_rows.append([InlineKeyboardButton("↩️ Anuluj", callback_data="cancel_edit")])
        await query.edit_message_reply_markup(reply_markup=None)
        msg = await query.message.chat.send_message("Wybierz pozycję:", reply_markup=InlineKeyboardMarkup(kb_rows))
        context.user_data["msg_ids"].append(msg.message_id)
        return SELECT_ENTRY

    # finish_edit → sprzątnij prompty i pokaż gotowy raport
    for mid in context.user_data.get("msg_ids", []):
        try:
            await query.message.chat.delete_message(mid)
        except Exception:
            pass

    date = context.user_data["date"]
    uid = context.user_data["uid"]
    entries = read_entries_for_day(uid, date)
    name = entries[0]["name"] if entries else context.user_data.get("name", query.from_user.first_name)
    rpt = format_report(entries, date, name)
    msg = await query.message.chat.send_message(rpt)

    mapping = load_mapping()
    mapping[f"{uid}_{date}"] = msg.message_id
    save_mapping(mapping)
    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.effective_chat.send_message("Anulowano.")
    return ConversationHandler.END


async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.effective_chat.send_message(
        "Użyj /start aby otworzyć menu. Tryby: tworzenie raportu (dodaje pozycje) oraz edycja pola bez przepisywania całości."
    )


# ──────────────────── PTB Application ────────────────────
async def on_startup(app: Application) -> None:
    await app.bot.set_my_commands(
        [
            BotCommand("start", "Otwórz menu raportów"),
            BotCommand("export", "Eksportuj raporty"),
            BotCommand("help", "Pomoc"),
        ]
    )


def build_app() -> Application:
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).post_init(on_startup).build()

    # komendy
    app.add_handler(CommandHandler("start", show_menu))
    app.add_handler(CommandHandler("export", export_handler))
    app.add_handler(CommandHandler("help", help_cmd))

    # conversation
    conv = ConversationHandler(
        entry_points=[CallbackQueryHandler(menu_handler, pattern=r"^(create|edit|export)$")],
        states={
            # tworzenie
            PLACE: [MessageHandler(filters.TEXT & ~filters.COMMAND, place)],
            START_TIME: [MessageHandler(filters.TEXT & ~filters.COMMAND, start_time)],
            END_TIME: [MessageHandler(filters.TEXT & ~filters.COMMAND, end_time)],
            TASKS: [MessageHandler(filters.TEXT & ~filters.COMMAND, tasks)],
            NOTES: [MessageHandler(filters.TEXT & ~filters.COMMAND, notes)],
            ANOTHER: [CallbackQueryHandler(another, pattern=r"^(again|finish)$")],
            # edycja
            SELECT_ENTRY: [CallbackQueryHandler(select_entry, pattern=r"^(entry:\d+|cancel_edit)$")],
            SELECT_FIELD: [CallbackQueryHandler(select_field, pattern=r"^(field:(place|start|end|tasks|notes)|back_to_entries)$")],
            EDIT_VALUE: [MessageHandler(filters.TEXT & ~filters.COMMAND, edit_value)],
            EDIT_MORE: [CallbackQueryHandler(edit_more, pattern=r"^(again_same|again_other|finish_edit)$")],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        per_chat=True,
        per_user=True,
        per_message=False,
    )
    app.add_handler(conv)
    return app


# ──────────────────── main ────────────────────
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s")

    if not TELEGRAM_TOKEN:
        raise SystemExit("Brak TELEGRAM_TOKEN w env.")

    bot_app = build_app()

    if WEBHOOK_URL:
        # produkcja – webhook (Northflank)
        bot_app.run_webhook(
            listen="0.0.0.0",
            port=PORT,
            url_path=TELEGRAM_TOKEN,
            webhook_url=f"{WEBHOOK_URL}/{TELEGRAM_TOKEN}",
        )
    else:
        # lokalnie – polling
        bot_app.run_polling(allowed_updates=Update.ALL_TYPES)
