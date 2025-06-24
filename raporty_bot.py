# ────────────────────────── raporty_bot.py ──────────────────────────
import os
import json
import asyncio
import logging
from datetime import datetime
from typing import Dict, List, Optional

from dotenv import load_dotenv
from flask import Flask, request
from openpyxl import Workbook, load_workbook

# ───────────── SharePoint (opcjonalny upload) ─────────────
try:
    from office365.sharepoint.client_context import ClientContext
    from office365.runtime.auth.client_credential import ClientCredential
except ModuleNotFoundError:
    ClientContext = ClientCredential = None  # biblioteka nieobecna → upload pomijamy

# ───────────── Telegram ─────────────
from telegram import (
    Bot,
    Update,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    BotCommand,
)
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    CallbackQueryHandler,
    MessageHandler,
    ContextTypes,
    ConversationHandler,
    filters,
    Application,
)

# ──────────────────── konfiguracja ────────────────────
load_dotenv()

TELEGRAM_TOKEN            = os.getenv("TELEGRAM_TOKEN")
WEBHOOK_URL               = os.getenv("WEBHOOK_URL")          # pusty → polling lokalny
SHAREPOINT_SITE           = os.getenv("SHAREPOINT_SITE")      # opcjonalnie
SHAREPOINT_DOC_LIB        = os.getenv("SHAREPOINT_DOC_LIB")   # opcjonalnie
SHAREPOINT_CLIENT_ID      = os.getenv("SHAREPOINT_CLIENT_ID") # opcjonalnie
SHAREPOINT_CLIENT_SECRET  = os.getenv("SHAREPOINT_CLIENT_SECRET")

EXCEL_FILE   = "reports.xlsx"
MAPPING_FILE = "report_msgs.json"

# ──────────────────── stany konwersacji ────────────────────
PLACE, START_TIME, END_TIME, TASKS, NOTES, ANOTHER = range(6)

# ──────────────────── funkcje pomocnicze ────────────────────
def load_mapping() -> Dict[str, int]:
    if os.path.exists(MAPPING_FILE):
        with open(MAPPING_FILE, "r") as f:
            return json.load(f)
    return {}

def save_mapping(mapping: Dict[str, int]) -> None:
    with open(MAPPING_FILE, "w") as f:
        json.dump(mapping, f)

def report_exists(user_id: int, date: str) -> bool:
    if not os.path.exists(EXCEL_FILE):
        return False
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
    prefix = f"{user_id}_{date}_"
    return any(str(r[0]).startswith(prefix) for r in ws.iter_rows(min_row=2, values_only=True))

def parse_time(text: str) -> Optional[str]:
    try:
        t = datetime.strptime(text.strip(), "%H:%M")
        return t.strftime("%H:%M")
    except ValueError:
        return None

def save_report(entries: List[Dict[str, str]],
                user_id: int,
                date: str,
                name: str,
                edit: bool = False) -> None:

    # Excel
    if os.path.exists(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.append(["ID", "Data", "Osoba", "Miejsce",
                   "Start", "Koniec", "Zadania", "Uwagi"])

    if edit:
        prefix = f"{user_id}_{date}_"
        rows = list(ws.iter_rows(min_row=2))
        idxs = [row[0].row for row in rows if str(row[0].value).startswith(prefix)]
        for i in sorted(idxs, reverse=True):
            ws.delete_rows(i)

    for idx, e in enumerate(entries, start=1):
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

    # SharePoint upload (jeśli biblioteka i zmienne są dostępne)
    if all([ClientContext,
            SHAREPOINT_SITE, SHAREPOINT_DOC_LIB,
            SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET]):
        ctx = ClientContext(SHAREPOINT_SITE).with_credentials(
            ClientCredential(SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET)
        )
        folder = ctx.web.get_folder_by_server_relative_url(SHAREPOINT_DOC_LIB)
        with open(EXCEL_FILE, "rb") as f:
            folder.upload_file(os.path.basename(EXCEL_FILE), f).execute_query()

def format_report(entries: List[Dict[str, str]],
                  date: str,
                  name: str) -> str:
    lines = [f"📄 Raport dzienny – {date}", f"👤 Osoba: {name}", ""]
    for e in entries:
        lines.extend([
            f"📍 Miejsce: {e['place']}",
            f"⏰ {e['start']} – {e['end']}",
            "📝 Zadania:",
            e["tasks"],
            "💬 Uwagi:",
            e["notes"],
            "",
        ])
    return "\n".join(lines)

# ──────────────────── handlery (bez zmian w logice) ────────────────────
async def show_menu(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    context.user_data.clear()
    context.user_data["msg_ids"] = []
    date = datetime.now().strftime("%d.%m.%Y")
    uid  = update.effective_user.id

    create_text = "📋 Stwórz raport" if not report_exists(uid, date) else "✏️ Edytuj raport"
    cb_data     = "create"           if not report_exists(uid, date) else "edit"

    kb = [
        [InlineKeyboardButton(create_text, callback_data=cb_data)],
        [InlineKeyboardButton("📥 Eksportuj",  callback_data="export")],
    ]

    m = await update.effective_chat.send_message(
        "Wybierz opcję:",
        reply_markup=InlineKeyboardMarkup(kb),
    )
    context.user_data["msg_ids"].append(m.message_id)

async def export_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
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
                chat_id,
                f,
                filename=EXCEL_FILE,
                caption=f"Raporty za {datetime.now().strftime('%m.%Y')}",
            )
            context.user_data.setdefault("msg_ids", []).append(doc.message_id)
    return ConversationHandler.END

async def menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()
    data = query.data

    if data == "export":
        return await export_handler(update, context)

    edit_flag = data == "edit"
    date = datetime.now().strftime("%d.%m.%Y")
    key  = f"{query.from_user.id}_{date}"

    mapping = load_mapping()
    if edit_flag and key in mapping:
        try:
            await context.bot.delete_message(query.message.chat.id, mapping[key])
        except Exception:
            pass

    context.user_data.update(
        {
            "entries": [],
            "edit": edit_flag,
            "date": date,
            "name": query.from_user.first_name,
            "msg_ids": context.user_data.get("msg_ids", []),
        }
    )
    await query.edit_message_reply_markup(reply_markup=None)
    msg = await context.bot.send_message(
        chat_id=query.message.chat.id,
        text="📍 Podaj miejsce wykonywania pracy:",
        allow_sending_without_reply=True,
    )
    context.user_data["msg_ids"].append(msg.message_id)
    return PLACE

async def place(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["msg_ids"].append(update.message.message_id)
    await update.message.delete()
    place = update.message.text.strip()

    if not place:
        err = await update.effective_chat.send_message("Podaj poprawne miejsce.")
        context.user_data["msg_ids"].append(err.message_id)
        return PLACE

    context.user_data["place"] = place
    ask = await update.effective_chat.send_message("⏰ Podaj godzinę rozpoczęcia pracy (HH:MM):")
    context.user_data["msg_ids"].append(ask.message_id)
    return START_TIME

async def start_time(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["msg_ids"].append(update.message.message_id)
    await update.message.delete()
    t = parse_time(update.message.text or "")
    last_end = context.user_data["entries"][-1]["end"] if context.user_data["entries"] else None

    if not t or (last_end and t <= last_end):
        err = await update.effective_chat.send_message("⏰ Błędna godzina. Spróbuj ponownie.")
        context.user_data["msg_ids"].append(err.message_id)
        return START_TIME

    context.user_data["start"] = t
    ask = await update.effective_chat.send_message("⏰ Podaj godzinę zakończenia pracy (HH:MM):")
    context.user_data["msg_ids"].append(ask.message_id)
    return END_TIME

async def end_time(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["msg_ids"].append(update.message.message_id)
    t = parse_time(update.message.text or "")

    if not t or t <= context.user_data.get("start"):
        err = await update.effective_chat.send_message("⏰ Błędna godzina. Spróbuj ponownie.")
        context.user_data["msg_ids"].append(err.message_id)
        return END_TIME

    context.user_data["end"] = t
    ask = await update.effective_chat.send_message("📝 Opisz wykonane prace:")
    context.user_data["msg_ids"].append(ask.message_id)
    return TASKS

async def tasks(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["msg_ids"].append(update.message.message_id)
    txt = update.message.text.strip()

    if not txt:
        err = await update.effective_chat.send_message("📝 Lista zadań nie może być pusta.")
        context.user_data["msg_ids"].append(err.message_id)
        return TASKS

    context.user_data["tasks"] = txt
    ask = await update.effective_chat.send_message("💬 Dodaj uwagi lub wpisz '-' jeśli brak:")
    context.user_data["msg_ids"].append(ask.message_id)
    return NOTES

async def notes(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["msg_ids"].append(update.message.message_id)
    txt = update.message.text.strip()

    if not txt:
        err = await update.effective_chat.send_message("💬 Uwagi nie mogą być puste.")
        context.user_data["msg_ids"].append(err.message_id)
        return NOTES

    entry = {
        "place": context.user_data.pop("place"),
        "start": context.user_data.pop("start"),
        "end":   context.user_data.pop("end"),
        "tasks": context.user_data.pop("tasks"),
        "notes": txt,
    }
    context.user_data["entries"].append(entry)

    kb = [
        [InlineKeyboardButton("Dodaj kolejne miejsce", callback_data="again")],
        [InlineKeyboardButton("Zakończ raport",         callback_data="finish")],
    ]
    msg = await update.effective_chat.send_message("Co dalej?", reply_markup=InlineKeyboardMarkup(kb))
    context.user_data["msg_ids"].append(msg.message_id)
    return ANOTHER

async def another(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()

    if query.data == "again":
        await query.edit_message_reply_markup()
        ask = await query.message.chat.send_message(
            "📍 Podaj miejsce wykonywania pracy:",
            allow_sending_without_reply=True,
        )
        context.user_data["msg_ids"].append(ask.message_id)
        return PLACE

    # finish
    for mid in context.user_data.get("msg_ids", []):
        try:
            await query.message.chat.delete_message(mid)
        except Exception:
            pass

    save_report(
        context.user_data["entries"],
        query.from_user.id,
        context.user_data["date"],
        context.user_data["name"],
        edit=context.user_data.get("edit", False),
    )

    rpt = format_report(context.user_data["entries"],
                        context.user_data["date"],
                        context.user_data["name"])

    msg = await query.message.chat.send_message(rpt)

    mapping = load_mapping()
    mapping[f"{query.from_user.id}_{context.user_data['date']}"] = msg.message_id
    save_mapping(mapping)
    return ConversationHandler.END

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.effective_chat.send_message("Anulowano.")
    return ConversationHandler.END

async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.effective_chat.send_message(
        "Użyj /start do menu, /export do pobrania raportów lub /help po pomoc."
    )

# ──────────────────── PTB Application ────────────────────
async def on_startup(app: Application) -> None:
    await app.bot.set_my_commands([
        BotCommand("start",  "Otwórz menu raportów"),
        BotCommand("export", "Eksportuj raporty"),
        BotCommand("help",   "Pomoc"),
    ])
    if WEBHOOK_URL:  # webhook tylko w produkcji
        await app.bot.set_webhook(f"{WEBHOOK_URL}/{TELEGRAM_TOKEN}")

def build_app() -> Application:
    app = (
        ApplicationBuilder()
        .token(TELEGRAM_TOKEN)
        .post_init(on_startup)
        .build()
    )

    # komendy
    app.add_handler(CommandHandler("start",  show_menu))
    app.add_handler(CommandHandler("export", export_handler))
    app.add_handler(CommandHandler("help",   help_cmd))

    # conversation
    conv = ConversationHandler(
        entry_points=[CallbackQueryHandler(menu_handler, pattern="^(create|edit|export)$")],
        states={
            PLACE:      [MessageHandler(filters.TEXT & ~filters.COMMAND, place)],
            START_TIME: [MessageHandler(filters.TEXT & ~filters.COMMAND, start_time)],
            END_TIME:   [MessageHandler(filters.TEXT & ~filters.COMMAND, end_time)],
            TASKS:      [MessageHandler(filters.TEXT & ~filters.COMMAND, tasks)],
            NOTES:      [MessageHandler(filters.TEXT & ~filters.COMMAND, notes)],
            ANOTHER:    [CallbackQueryHandler(another, pattern="^(again|finish)$")],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        per_chat=True,
        per_user=True,
        per_message=False,
    )
    app.add_handler(conv)
    return app

# ──────────────────── Flask + stała pętla zdarzeń ────────────────────
flask_app = Flask(__name__)
bot_app    = build_app()
bot: Bot   = bot_app.bot  # alias

# Jedna pętla zdarzeń na cały worker Gunicorna
loop = asyncio.new_event_loop()
asyncio.set_event_loop(loop)
loop.run_until_complete(bot_app.initialize())

@flask_app.route(f"/{TELEGRAM_TOKEN}", methods=["POST"])
def telegram_webhook() -> str:
    """Odbiera Update z Telegrama i przekazuje do PTB w istniejącej pętli."""
    update = Update.de_json(request.get_json(force=True), bot)
    asyncio.run_coroutine_threadsafe(bot_app.process_update(update), loop)
    return "OK"

@flask_app.route("/")
def index() -> str:
    return "Bot działa!"

# Gunicorn na Renderze spodziewa się zmiennej `app`
app = flask_app

# ────────── lokalnie (brak WEBHOOK_URL) ──────────
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)

    if not WEBHOOK_URL:                     # tryb developerski – polling
        bot_app.run_polling(allowed_updates=Update.ALL_TYPES)
    else:                                   # test webhooka lokalnie (np. ngrok)
        bot_app.bot.set_webhook(f"{WEBHOOK_URL}/{TELEGRAM_TOKEN}")
        flask_app.run(host="0.0.0.0",
                      port=int(os.getenv("PORT", 5000)))
