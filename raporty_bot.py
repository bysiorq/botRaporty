# ────────────────────────── raporty_bot.py ──────────────────────────
import os
import json
import asyncio
import logging
import threading
from datetime import datetime
from typing import Dict, List, Optional

from dotenv import load_dotenv
from flask import Flask, request
from openpyxl import Workbook, load_workbook

# ───────────── SharePoint (opcjonalny upload) ─────────────
try:
    from office365.sharepoint.client_context import ClientContext
    from office365.runtime.auth.client_credential import ClientCredential
except ModuleNotFoundError:              # brak biblioteki → pomijamy
    ClientContext = ClientCredential = None

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

TELEGRAM_TOKEN           = os.getenv("TELEGRAM_TOKEN")
WEBHOOK_URL              = os.getenv("WEBHOOK_URL")            # pusty → polling
SHAREPOINT_SITE          = os.getenv("SHAREPOINT_SITE")
SHAREPOINT_DOC_LIB       = os.getenv("SHAREPOINT_DOC_LIB")
SHAREPOINT_CLIENT_ID     = os.getenv("SHAREPOINT_CLIENT_ID")
SHAREPOINT_CLIENT_SECRET = os.getenv("SHAREPOINT_CLIENT_SECRET")

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
    return any(str(r[0]).startswith(prefix)
               for r in ws.iter_rows(min_row=2, values_only=True))

def parse_time(text: str) -> Optional[str]:
    try:
        return datetime.strptime(text.strip(), "%H:%M").strftime("%H:%M")
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
        for row in sorted((r[0].row for r in rows
                           if str(r[0].value).startswith(prefix)),
                          reverse=True):
            ws.delete_rows(row)

    for idx, e in enumerate(entries, 1):
        ws.append([
            f"{user_id}_{date}_{idx}", date, name,
            e["place"], e["start"], e["end"], e["tasks"], e["notes"]
        ])
    wb.save(EXCEL_FILE)

    # SharePoint upload (jeżeli skonfigurowany)
    if all([ClientContext, SHAREPOINT_SITE, SHAREPOINT_DOC_LIB,
            SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET]):
        ctx = ClientContext(SHAREPOINT_SITE).with_credentials(
            ClientCredential(SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET)
        )
        folder = ctx.web.get_folder_by_server_relative_url(SHAREPOINT_DOC_LIB)
        with open(EXCEL_FILE, "rb") as f:
            folder.upload_file(os.path.basename(EXCEL_FILE), f).execute_query()

def format_report(entries: List[Dict[str, str]], date: str, name: str) -> str:
    out = [f"📄 Raport dzienny – {date}", f"👤 Osoba: {name}", ""]
    for e in entries:
        out.extend([
            f"📍 Miejsce: {e['place']}",
            f"⏰ {e['start']} – {e['end']}",
            "📝 Zadania:", e["tasks"],
            "💬 Uwagi:",   e["notes"], ""
        ])
    return "\n".join(out)

# ──────────────────── handlery ────────────────────
async def show_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    context.user_data["msg_ids"] = []
    date = datetime.now().strftime("%d.%m.%Y")
    uid  = update.effective_user.id

    new = not report_exists(uid, date)
    kb  = [[InlineKeyboardButton("📋 Stwórz raport" if new else "✏️ Edytuj raport",
                                 callback_data="create" if new else "edit")],
           [InlineKeyboardButton("📥 Eksportuj", callback_data="export")]]

    m = await update.effective_chat.send_message("Wybierz opcję:",
                                                 reply_markup=InlineKeyboardMarkup(kb))
    context.user_data["msg_ids"].append(m.message_id)

async def export_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.callback_query:
        await update.callback_query.answer()
        cid = update.callback_query.message.chat.id
        await update.callback_query.edit_message_reply_markup(reply_markup=None)
    else:
        cid = update.effective_chat.id

    if not os.path.exists(EXCEL_FILE):
        msg = await context.bot.send_message(cid, "⚠️ Brak pliku raportów.")
        context.user_data.setdefault("msg_ids", []).append(msg.message_id)
    else:
        with open(EXCEL_FILE, "rb") as f:
            doc = await context.bot.send_document(
                cid, f, filename=EXCEL_FILE,
                caption=f"Raporty za {datetime.now().strftime('%m.%Y')}")
            context.user_data.setdefault("msg_ids", []).append(doc.message_id)
    return ConversationHandler.END

async def menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()

    if q.data == "export":
        return await export_handler(update, context)

    edit   = q.data == "edit"
    date   = datetime.now().strftime("%d.%m.%Y")
    key    = f"{q.from_user.id}_{date}"
    mapping = load_mapping()

    if edit and key in mapping:
        try:
            await context.bot.delete_message(q.message.chat.id, mapping[key])
        except Exception:
            pass

    context.user_data.update({"entries": [], "edit": edit, "date": date,
                              "name": q.from_user.first_name,
                              "msg_ids": context.user_data.get("msg_ids", [])})

    await q.edit_message_reply_markup(reply_markup=None)
    m = await context.bot.send_message(q.message.chat.id,
                                       "📍 Podaj miejsce wykonywania pracy:",
                                       allow_sending_without_reply=True)
    context.user_data["msg_ids"].append(m.message_id)
    return PLACE

async def place(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["msg_ids"].append(update.message.message_id)
    await update.message.delete()

    txt = update.message.text.strip()
    if not txt:
        err = await update.effective_chat.send_message("Podaj poprawne miejsce.")
        context.user_data["msg_ids"].append(err.message_id)
        return PLACE

    context.user_data["place"] = txt
    ask = await update.effective_chat.send_message(
        "⏰ Podaj godzinę rozpoczęcia pracy (HH:MM):")
    context.user_data["msg_ids"].append(ask.message_id)
    return START_TIME

async def start_time(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["msg_ids"].append(update.message.message_id)
    await update.message.delete()

    t = parse_time(update.message.text or "")
    last_end = context.user_data["entries"][-1]["end"] if context.user_data["entries"] else None
    if not t or (last_end and t <= last_end):
        err = await update.effective_chat.send_message("⏰ Błędna godzina. Spróbuj ponownie.")
        context.user_data["msg_ids"].append(err.message_id)
        return START_TIME

    context.user_data["start"] = t
    ask = await update.effective_chat.send_message(
        "⏰ Podaj godzinę zakończenia pracy (HH:MM):")
    context.user_data["msg_ids"].append(ask.message_id)
    return END_TIME

async def end_time(update: Update, context: ContextTypes.DEFAULT_TYPE):
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

async def tasks(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["msg_ids"].append(update.message.message_id)
    txt = update.message.text.strip()
    if not txt:
        err = await update.effective_chat.send_message("📝 Lista zadań nie może być pusta.")
        context.user_data["msg_ids"].append(err.message_id)
        return TASKS

    context.user_data["tasks"] = txt
    ask = await update.effective_chat.send_message(
        "💬 Dodaj uwagi lub wpisz '-' jeśli brak:")
    context.user_data["msg_ids"].append(ask.message_id)
    return NOTES

async def notes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["msg_ids"].append(update.message.message_id)
    txt = update.message.text.strip()
    if not txt:
        err = await update.effective_chat.send_message("💬 Uwagi nie mogą być puste.")
        context.user_data["msg_ids"].append(err.message_id)
        return NOTES

    context.user_data["entries"].append({
        "place": context.user_data.pop("place"),
        "start": context.user_data.pop("start"),
        "end":   context.user_data.pop("end"),
        "tasks": context.user_data.pop("tasks"),
        "notes": txt,
    })

    kb = [[InlineKeyboardButton("Dodaj kolejne miejsce", callback_data="again")],
          [InlineKeyboardButton("Zakończ raport",        callback_data="finish")]]
    m = await update.effective_chat.send_message("Co dalej?",
                                                 reply_markup=InlineKeyboardMarkup(kb))
    context.user_data["msg_ids"].append(m.message_id)
    return ANOTHER

async def another(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()

    if q.data == "again":
        await q.edit_message_reply_markup()
        ask = await q.message.chat.send_message("📍 Podaj miejsce wykonywania pracy:",
                                                allow_sending_without_reply=True)
        context.user_data["msg_ids"].append(ask.message_id)
        return PLACE

    # finish
    for mid in context.user_data.get("msg_ids", []):
        try:
            await q.message.chat.delete_message(mid)
        except Exception:
            pass

    save_report(context.user_data["entries"],
                q.from_user.id,
                context.user_data["date"],
                context.user_data["name"],
                edit=context.user_data.get("edit", False))

    rpt = format_report(context.user_data["entries"],
                        context.user_data["date"],
                        context.user_data["name"])
    msg = await q.message.chat.send_message(rpt)

    mapping = load_mapping()
    mapping[f"{q.from_user.id}_{context.user_data['date']}"] = msg.message_id
    save_mapping(mapping)
    return ConversationHandler.END

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.effective_chat.send_message("Anulowano.")
    return ConversationHandler.END

async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.effective_chat.send_message(
        "Użyj /start do menu, /export do pobrania raportów lub /help po pomoc.")

# ──────────────────── PTB Application ────────────────────
async def on_startup(app: Application):
    await app.bot.set_my_commands([
        BotCommand("start",  "Otwórz menu raportów"),
        BotCommand("export", "Eksportuj raporty"),
        BotCommand("help",   "Pomoc"),
    ])
    if WEBHOOK_URL:
        await app.bot.set_webhook(f"{WEBHOOK_URL}/{TELEGRAM_TOKEN}")

def build_app() -> Application:
    app = (ApplicationBuilder()
           .token(TELEGRAM_TOKEN)
           .post_init(on_startup)
           .build())

    app.add_handler(CommandHandler("start",  show_menu))
    app.add_handler(CommandHandler("export", export_handler))
    app.add_handler(CommandHandler("help",   help_cmd))

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
        per_chat=True, per_user=True, per_message=False)
    app.add_handler(conv)
    return app

# ──────────────────── Flask + async-loop w wątku ────────────────────
flask_app = Flask(__name__)
bot_app   = build_app()
bot: Bot  = bot_app.bot

def _start_async_loop():
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    loop.run_until_complete(bot_app.initialize())
    loop.create_task(bot_app.start())
    bot_app.loop = loop          # przechowujemy referencję
    loop.run_forever()

threading.Thread(target=_start_async_loop, daemon=True).start()

@flask_app.route(f"/{TELEGRAM_TOKEN}", methods=["POST"])
def telegram_webhook():
    update = Update.de_json(request.get_json(force=True), bot)
    asyncio.run_coroutine_threadsafe(bot_app.process_update(update),
                                     bot_app.loop)
    return "OK"

@flask_app.route("/")
def index():
    return "Bot działa!"

# Gunicorn na Renderze importuje zmienną `app`
app = flask_app

# ────────── uruchomienie lokalne ──────────
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    if not WEBHOOK_URL:        # dev: polling
        bot_app.run_polling(allowed_updates=Update.ALL_TYPES)
    else:                      # test webhooka lokalnie (np. ngrok)
        bot_app.bot.set_webhook(f"{WEBHOOK_URL}/{TELEGRAM_TOKEN}")
        flask_app.run(host="0.0.0.0",
                      port=int(os.getenv("PORT", 5000)))
