# ────────────────────────── raporty_bot.py (refactor 2025-08, sticky + lock + DATA_DIR + presets + calendar + exports + backups + validation + tags + summaries) ──────────────────────────
import os
import re
import json
import logging
import shutil
import calendar as cal
from datetime import datetime, date, timedelta
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
from telegram.error import BadRequest

# ───────────── File locking & atomic save ─────────────
# pip install portalocker
import tempfile
import portalocker

# ──────────────────── konfiguracja ────────────────────
load_dotenv()

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")  # MUST HAVE
WEBHOOK_URL = os.getenv("WEBHOOK_URL", "").rstrip("/")  # np. https://app-xyz.northflank.app
PORT = int(os.getenv("PORT", 8080))  # Northflank zwykle 8080

DATA_DIR = os.getenv("DATA_DIR", ".")
os.makedirs(DATA_DIR, exist_ok=True)
BACKUP_DIR = os.path.join(DATA_DIR, "backups")
os.makedirs(BACKUP_DIR, exist_ok=True)
BACKUP_KEEP = int(os.getenv("BACKUP_KEEP", "20"))

# opcjonalne ustawienia SharePoint
SHAREPOINT_SITE = os.getenv("SHAREPOINT_SITE")
SHAREPOINT_DOC_LIB = os.getenv("SHAREPOINT_DOC_LIB")
SHAREPOINT_CLIENT_ID = os.getenv("SHAREPOINT_CLIENT_ID")
SHAREPOINT_CLIENT_SECRET = os.getenv("SHAREPOINT_CLIENT_SECRET")

EXCEL_FILE = os.path.join(DATA_DIR, "reports.xlsx")
MAPPING_FILE = os.path.join(DATA_DIR, "report_msgs.json")
PRESETS_FILE = os.path.join(DATA_DIR, "presets.json")
LOCK_FILE = os.path.join(DATA_DIR, "reports.lock")

ADMIN_IDS = {int(x) for x in os.getenv("ADMIN_IDS", "").split(",") if x.strip().isdigit()}

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
# wybór daty i decyzja o overlapie
DATE_PICK, OVERLAP_DECIDE = range(10, 12)

# ──────────────────── helpers: excel/lock/backup ────────────────────
def _atomic_save_wb(wb: Workbook, path: str) -> None:
    fd, tmp_path = tempfile.mkstemp(dir=os.path.dirname(path), suffix=".tmp")
    os.close(fd)
    wb.save(tmp_path)
    os.replace(tmp_path, path)


def _with_lock(fn, *args, **kwargs):
    with portalocker.Lock(LOCK_FILE, timeout=30):
        return fn(*args, **kwargs)


def _backup_file():
    if not os.path.exists(EXCEL_FILE):
        return
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = os.path.join(BACKUP_DIR, f"reports_{ts}.xlsx")
    try:
        shutil.copy2(EXCEL_FILE, dst)
    except Exception as e:
        logging.warning("Backup failed: %s", e)
    # rotacja
    files = sorted([f for f in os.listdir(BACKUP_DIR) if f.startswith("reports_") and f.endswith(".xlsx")])
    if len(files) > BACKUP_KEEP:
        for old in files[: len(files) - BACKUP_KEEP]:
            try:
                os.remove(os.path.join(BACKUP_DIR, old))
            except Exception:
                pass


def open_wb() -> Workbook:
    def _open():
        if os.path.exists(EXCEL_FILE):
            return load_workbook(EXCEL_FILE)
        wb = Workbook()
        return wb
    return _with_lock(_open)


def month_key_from_date(date_str: str) -> str:
    d = datetime.strptime(date_str, "%d.%m.%Y")
    return f"{d.year:04d}-{d.month:02d}"


def ensure_month_sheet(wb: Workbook, month_key: str) -> Worksheet:
    """Zwraca arkusz dla danego miesiąca. Tworzy jeśli nie istnieje i ustawia nagłówki.
    Nowy miesiąc ląduje jako pierwszy arkusz (na górze)."""
    ws: Optional[Worksheet] = wb[month_key] if month_key in wb.sheetnames else None
    if ws is None:
        ws = wb.create_sheet(title=month_key, index=0)  # na górze
        ws.append(HEADERS)
        # usuń domyślny "Sheet" jeśli pusty
        if "Sheet" in wb.sheetnames and wb["Sheet"].max_row == 1 and wb["Sheet"].max_column == 1:
            wb.remove(wb["Sheet"])
    else:
        # przenieś na początek
        idx = wb.sheetnames.index(month_key)
        if idx != 0:
            wb.move_sheet(ws, offset=-idx)
    return ws


def get_month_sheet_if_exists(wb: Workbook, month_key: str) -> Optional[Worksheet]:
    return wb[month_key] if month_key in wb.sheetnames else None


def report_exists(user_id: int, date_str: str) -> bool:
    if not os.path.exists(EXCEL_FILE):
        return False
    def _exists():
        wb = load_workbook(EXCEL_FILE)
        ws = get_month_sheet_if_exists(wb, month_key_from_date(date_str))
        if not ws:
            return False
        prefix = f"{user_id}_{date_str}_"
        for row in ws.iter_rows(min_row=2, values_only=True):
            if (row and row[0]) and str(row[0]).startswith(prefix):
                return True
        return False
    return _with_lock(_exists)


def save_report(entries: List[Dict[str, str]], user_id: int, date_str: str, name: str) -> None:
    """Dopisuje nowe wpisy na dany dzień (append). Indeksy kontynuowane. Edycje pól robi update_report_field()."""
    def _save():
        wb = open_wb()  # już w locku
        ws = ensure_month_sheet(wb, month_key_from_date(date_str))
        prefix = f"{user_id}_{date_str}_"
        existing_idxs: List[int] = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            rid = str(row[0]) if row and row[0] is not None else ""
            if rid.startswith(prefix):
                try:
                    existing_idxs.append(int(rid.split("_")[-1]))
                except Exception:
                    pass
        next_idx = (max(existing_idxs) + 1) if existing_idxs else 1

        for off, e in enumerate(entries):
            idx = next_idx + off
            ws.append([
                f"{user_id}_{date_str}_{idx}",
                date_str,
                name,
                e["place"],
                e["start"],
                e["end"],
                e["tasks"],
                e["notes"],
            ])
        _backup_file()
        _atomic_save_wb(wb, EXCEL_FILE)
    _with_lock(_save)
    _maybe_upload_sharepoint()


def read_entries_for_day(user_id: int, date_str: str) -> List[Dict[str, str]]:
    if not os.path.exists(EXCEL_FILE):
        return []
    def _read():
        wb = load_workbook(EXCEL_FILE)
        ws = get_month_sheet_if_exists(wb, month_key_from_date(date_str))
        if not ws:
            return []
        prefix = f"{user_id}_{date_str}_"
        out: List[Dict[str, str]] = []
        for row in ws.iter_rows(min_row=2, values_only=False):
            rid = str(row[0].value) if row and row[0] is not None else ""
            if rid and rid.startswith(prefix):
                out.append({
                    "rid": rid,
                    "row": row[0].row,
                    "date": row[COLS["Data"] - 1].value,
                    "name": row[COLS["Imię"] - 1].value,
                    "place": row[COLS["Miejsce"] - 1].value or "",
                    "start": row[COLS["Start"] - 1].value or "",
                    "end": row[COLS["Koniec"] - 1].value or "",
                    "tasks": row[COLS["Zadania"] - 1].value or "",
                    "notes": row[COLS["Uwagi"] - 1].value or "",
                })
        out.sort(key=lambda e: int(e["rid"].split("_")[-1]))
        return out
    return _with_lock(_read)


def read_entries_all_weeks(user_id: int) -> List[Dict[str, str]]:
    if not os.path.exists(EXCEL_FILE):
        return []
    def _read_all():
        wb = load_workbook(EXCEL_FILE)
        out: List[Dict[str, str]] = []
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            # pomiń puste arkusze
            if ws.max_row < 2:
                continue
            for row in ws.iter_rows(min_row=2, values_only=False):
                rid = str(row[0].value) if row and row[0] is not None else ""
                if rid and rid.startswith(f"{user_id}_"):
                    out.append({
                        "rid": rid,
                        "date": row[COLS["Data"] - 1].value,
                        "start": row[COLS["Start"] - 1].value or "",
                        "end": row[COLS["Koniec"] - 1].value or "",
                    })
        return out
    return _with_lock(_read_all)


def update_report_field(user_id: int, date_str: str, rid: str, field: str, new_value: str) -> None:
    def _upd():
        wb = load_workbook(EXCEL_FILE)
        ws = ensure_month_sheet(wb, month_key_from_date(date_str))
        col_name_map = {
            "place": "Miejsce",
            "start": "Start",
            "end": "Koniec",
            "tasks": "Zadania",
            "notes": "Uwagi",
        }
        target_col = COLS[col_name_map[field]]
        target_row = None
        for row in ws.iter_rows(min_row=2, values_only=False):
            if str(row[0].value) == rid:
                target_row = row[0].row
                break
        if not target_row:
            raise RuntimeError("Nie znaleziono wiersza do edycji.")
        ws.cell(row=target_row, column=target_col, value=new_value)
        _backup_file()
        _atomic_save_wb(wb, EXCEL_FILE)
    _with_lock(_upd)
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

# ──────────────────── helpers: mapping/presets/utils ────────────────────
def load_mapping() -> Dict[str, int]:
    if os.path.exists(MAPPING_FILE):
        with open(MAPPING_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_mapping(mapping: Dict[str, int]) -> None:
    with open(MAPPING_FILE, "w", encoding="utf-8") as f:
        json.dump(mapping, f)


def load_presets() -> Dict[str, Dict[str, List[str]]]:
    if os.path.exists(PRESETS_FILE):
        with open(PRESETS_FILE, "r", encoding="utf-8") as f:
            try:
                return json.load(f)
            except Exception:
                return {}
    return {}


def save_presets(presets: Dict[str, Dict[str, List[str]]]) -> None:
    with open(PRESETS_FILE, "w", encoding="utf-8") as f:
        json.dump(presets, f, ensure_ascii=False)


def remember_place(user_id: int, place: str) -> None:
    def _upd():
        presets = load_presets()
        key = str(user_id)
        user = presets.setdefault(key, {"places": [], "tasks": []})
        if place in user["places"]:
            user["places"].remove(place)
        user["places"].insert(0, place)
        user["places"] = user["places"][:5]
        save_presets(presets)
    _with_lock(_upd)


def remember_tasks(user_id: int, tasks: str) -> None:
    def _upd():
        presets = load_presets()
        key = str(user_id)
        user = presets.setdefault(key, {"places": [], "tasks": []})
        t = tasks.strip()
        if not t:
            return
        if t in user["tasks"]:
            user["tasks"].remove(t)
        user["tasks"].insert(0, t)
        user["tasks"] = user["tasks"][:5]
        save_presets(presets)
    _with_lock(_upd)


def get_recent_places(user_id: int) -> List[str]:
    presets = load_presets()
    return presets.get(str(user_id), {}).get("places", [])


def get_task_templates(user_id: int) -> List[str]:
    presets = load_presets()
    return presets.get(str(user_id), {}).get("tasks", [])


def parse_time(text: str) -> Optional[str]:
    try:
        t = datetime.strptime(text.strip(), "%H:%M")
        return t.strftime("%H:%M")
    except Exception:
        return None


def time_to_minutes(t: str) -> int:
    h, m = map(int, t.split(":"))
    return h * 60 + m


def minutes_to_hhmm(m: int) -> str:
    h = m // 60
    mm = m % 60
    return f"{h}h {mm:02d}m"


def extract_tags(text: str) -> List[str]:
    return re.findall(r"#[\wąćęłńóśżźĄĆĘŁŃÓŚŻŹ]+", text or "")


def intervals_overlap(a_start: str, a_end: str, b_start: str, b_end: str) -> bool:
    return max(time_to_minutes(a_start), time_to_minutes(b_start)) < min(time_to_minutes(a_end), time_to_minutes(b_end))


def has_overlap(user_id: int, date_str: str, start: str, end: str, exclude_rid: Optional[str] = None, in_memory: Optional[List[Dict]] = None) -> Tuple[bool, List[Tuple[str, str]]]:
    conflicts = []
    # in-memory (z aktualnej sesji)
    for e in (in_memory or []):
        if intervals_overlap(start, end, e["start"], e["end"]):
            conflicts.append((e["start"], e["end"]))
    # z pliku
    for e in read_entries_for_day(user_id, date_str):
        if exclude_rid and e["rid"] == exclude_rid:
            continue
        if e["start"] and e["end"] and intervals_overlap(start, end, e["start"], e["end"]):
            conflicts.append((e["start"], e["end"]))
    return (len(conflicts) > 0, conflicts)


def compute_daily_minutes(entries: List[Dict[str, str]]) -> int:
    total = 0
    for e in entries:
        if e["start"] and e["end"]:
            total += time_to_minutes(e["end"]) - time_to_minutes(e["start"])
    return total


def compute_week_minutes(user_id: int, any_date_ddmmYYYY: str) -> int:
    d = datetime.strptime(any_date_ddmmYYYY, "%d.%m.%Y").date()
    iso_year, iso_week, _ = d.isocalendar()
    total = 0
    for e in read_entries_all_weeks(user_id):
        try:
            ed = datetime.strptime(e["date"], "%d.%m.%Y").date()
        except Exception:
            continue
        y, w, _ = ed.isocalendar()
        if (y, w) == (iso_year, iso_week) and e["start"] and e["end"]:
            total += time_to_minutes(e["end"]) - time_to_minutes(e["start"])
    return total


def format_report(entries: List[Dict[str, str]], date_str: str, name: str, with_totals: bool = False, uid: Optional[int] = None) -> str:
    lines = [f"📄 Raport dzienny – {date_str}", f"👤 Imię: {name}", ""]
    for i, e in enumerate(entries, start=1):
        tags = ", ".join(extract_tags(e["tasks"]))
        lines.extend([
            f"#{i}",
            f"📍 Miejsce: {e['place']}",
            f"⏰ {e['start']} – {e['end']}",
            "📝 Zadania:",
            str(e["tasks"]) or "-",
        ])
        if tags:
            lines.append(f"🏷 Tagi: {tags}")
        lines.extend([
            "💬 Uwagi:",
            str(e["notes"]) or "-",
            "",
        ])
    if with_totals and uid is not None:
        day_min = compute_daily_minutes(entries)
        week_min = compute_week_minutes(uid, date_str)
        lines.append(f"➕ Razem dziś: {minutes_to_hhmm(day_min)} | Tydzień: {minutes_to_hhmm(week_min)}")
    return "\n".join(lines)

# ──────────────────── helpers: sticky + menu + calendar ────────────────────
async def sticky_set(update_or_ctx, context: ContextTypes.DEFAULT_TYPE, text: str, reply_markup: Optional[InlineKeyboardMarkup] = None):
    chat = update_or_ctx.effective_chat if isinstance(update_or_ctx, Update) else None
    chat_id = chat.id if chat else update_or_ctx.callback_query.message.chat.id
    sticky_id = context.user_data.get("sticky_id")
    if sticky_id:
        try:
            await context.bot.edit_message_text(chat_id=chat_id, message_id=sticky_id, text=text, reply_markup=reply_markup)
            return
        except BadRequest as e:
            if "message is not modified" in str(e).lower():
                return
        except Exception:
            pass
    m = await context.bot.send_message(chat_id, text, reply_markup=reply_markup)
    context.user_data["sticky_id"] = m.message_id


async def sticky_delete(context: ContextTypes.DEFAULT_TYPE, chat_id: int):
    sticky_id = context.user_data.get("sticky_id")
    if sticky_id:
        try:
            await context.bot.delete_message(chat_id, sticky_id)
        except Exception:
            pass
        context.user_data.pop("sticky_id", None)


def today_str() -> str:
    return datetime.now().strftime("%d.%m.%Y")


def yesterday_str() -> str:
    return (datetime.now() - timedelta(days=1)).strftime("%d.%m.%Y")


def to_ddmmyyyy(d: date) -> str:
    return d.strftime("%d.%m.%Y")


def build_main_menu(uid: int, date_str: str) -> InlineKeyboardMarkup:
    create_text = "📋 Stwórz raport" if not report_exists(uid, date_str) else "✏️ Edytuj raport"
    cb_data = "create" if not report_exists(uid, date_str) else "edit"
    kb = [
        [InlineKeyboardButton(f"📅 Data: {date_str}", callback_data="change_date")],
        [InlineKeyboardButton(create_text, callback_data=cb_data)],
        [InlineKeyboardButton("📥 Eksport", callback_data="export"), InlineKeyboardButton("📥 Mój eksport", callback_data="myexport")],
    ]
    return InlineKeyboardMarkup(kb)


def month_kb(year: int, month: int) -> InlineKeyboardMarkup:
    month_name = cal.month_name[month]
    days = cal.monthcalendar(year, month)  # tygodnie z dniami, 0 = brak dnia
    rows = []
    # nagłówek
    rows.append([InlineKeyboardButton(f"{month_name} {year}", callback_data="noop")])
    # dni tygodnia
    rows.append([InlineKeyboardButton(x, callback_data="noop") for x in ["Pn","Wt","Śr","Cz","Pt","So","Nd"]])
    # siatka
    for week in days:
        r = []
        for idx, d in enumerate(week):
            if d == 0:
                r.append(InlineKeyboardButton(" ", callback_data="noop"))
            else:
                ds = to_ddmmyyyy(date(year, month, d))
                r.append(InlineKeyboardButton(str(d), callback_data=f"day:{ds}"))
        rows.append(r)
    # nawigacja
    prev_month = (date(year, month, 1) - timedelta(days=1))
    next_month = (date(year, month, cal.monthrange(year, month)[1]) + timedelta(days=1))
    rows.append([
        InlineKeyboardButton("« Poprzedni", callback_data=f"cal:{prev_month.year}-{prev_month.month:02d}"),
        InlineKeyboardButton("Dziś", callback_data=f"day:{today_str()}"),
        InlineKeyboardButton("Następny »", callback_data=f"cal:{next_month.year}-{next_month.month:02d}"),
    ])
    return InlineKeyboardMarkup(rows)


# ──────────────────── handlery menu/top-level ────────────────────
async def show_menu(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    context.user_data.clear()
    context.user_data["msg_ids"] = set()
    sel_date = today_str()
    context.user_data["date"] = sel_date
    uid = update.effective_user.id
    await sticky_set(update, context, "Wybierz opcję:", build_main_menu(uid, sel_date))


async def change_date_cb(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    # pokaż kalendarz bieżącego miesiąca
    now = datetime.now()
    await sticky_set(update, context, "📅 Wybierz datę:", month_kb(now.year, now.month))
    return DATE_PICK


async def calendar_nav_cb(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    if data.startswith("cal:"):
        y, m = map(int, data.split(":")[1].split("-"))
        await sticky_set(update, context, "📅 Wybierz datę:", month_kb(y, m))
    elif data.startswith("day:"):
        ds = data.split(":")[1]
        context.user_data["date"] = ds
        uid = query.from_user.id
        await sticky_set(update, context, "Wybierz opcję:", build_main_menu(uid, ds))
        return ConversationHandler.END
    return DATE_PICK


async def export_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # /export [YYYY-MM] lub przycisk z menu (bez paramów → bieżący miesiąc z user_data)
    month_arg = None
    if update.callback_query and update.callback_query.data == "export":
        month_arg = month_key_from_date(context.user_data.get("date", today_str()))
    else:
        # z komendy
        args = getattr(context, "args", []) or []
        month_arg = args[0] if args else month_key_from_date(today_str())

    if ADMIN_IDS and update.effective_user.id not in ADMIN_IDS:
        await sticky_set(update, context, "Brak uprawnień do eksportu (tylko admini). Użyj /myexport <YYYY-MM>.")
        return ConversationHandler.END

    path = export_month(month_arg)
    if not path:
        await sticky_set(update, context, f"Brak danych dla {month_arg}.")
        return ConversationHandler.END

    with open(path, "rb") as f:
        await update.effective_chat.send_document(f, filename=os.path.basename(path), caption=f"Eksport {month_arg}")
    try:
        os.remove(path)
    except Exception:
        pass
    return ConversationHandler.END


async def myexport_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # /myexport [YYYY-MM] lub przycisk z menu
    month_arg = None
    if update.callback_query and update.callback_query.data == "myexport":
        month_arg = month_key_from_date(context.user_data.get("date", today_str()))
    else:
        args = getattr(context, "args", []) or []
        month_arg = args[0] if args else month_key_from_date(today_str())

    path = export_month(month_arg, user_id=update.effective_user.id)
    if not path:
        await sticky_set(update, context, f"Brak danych dla {month_arg}.")
        return ConversationHandler.END

    with open(path, "rb") as f:
        await update.effective_chat.send_document(f, filename=os.path.basename(path), caption=f"Mój eksport {month_arg}")
    try:
        os.remove(path)
    except Exception:
        pass
    return ConversationHandler.END


def export_month(month_key: str, user_id: Optional[int] = None) -> Optional[str]:
    if not os.path.exists(EXCEL_FILE):
        return None
    def _exp() -> Optional[str]:
        wb = load_workbook(EXCEL_FILE)
        if month_key not in wb.sheetnames:
            return None
        ws = wb[month_key]
        out = Workbook()
        wso = out.active
        wso.title = month_key
        wso.append(HEADERS)
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            if user_id and not str(row[0]).startswith(f"{user_id}_"):
                continue
            wso.append(list(row))
        tmpf = os.path.join(DATA_DIR, f"export_{month_key}_{user_id or 'ALL'}.xlsx")
        _atomic_save_wb(out, tmpf)
        return tmpf
    return _with_lock(_exp)


# ──────────────────── FLOW: CREATE ────────────────────
async def menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data

    if data in {"export", "myexport"}:
        # obsłużą dedykowane handlery
        return

    date_str = context.user_data.get("date", today_str())
    context.user_data.update({
        "name": query.from_user.first_name,
        "uid": query.from_user.id,
    })

    if data == "create":
        # pokaż presety miejsc
        places = get_recent_places(query.from_user.id)
        kb = []
        if places:
            kb.extend([[InlineKeyboardButton(p, callback_data=f"place_preset:{i}")] for i, p in enumerate(places)])
        kb.append([InlineKeyboardButton("Wpisz ręcznie", callback_data="place_manual")])
        await sticky_set(update, context, "📍 Podaj miejsce wykonywania pracy:", InlineKeyboardMarkup(kb))
        return PLACE

    if data == "edit":
        entries = read_entries_for_day(query.from_user.id, date_str)
        if not entries:
            await sticky_set(update, context, "Brak wpisów dla tej daty.")
            return ConversationHandler.END
        context.user_data["edit_entries"] = entries
        kb_rows = []
        for idx, e in enumerate(entries, start=1):
            label = f"#{idx} {e['place']}  {e['start']}-{e['end']}"
            kb_rows.append([InlineKeyboardButton(label, callback_data=f"entry:{idx-1}")])
        kb_rows.append([InlineKeyboardButton("↩️ Anuluj", callback_data="cancel_edit")])
        await sticky_set(update, context, "Wybierz pozycję do edycji:", InlineKeyboardMarkup(kb_rows))
        return SELECT_ENTRY


# PLACE: obsługa presetów i wpisu ręcznego
async def place(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # jeśli to callback z presetem
    if update.callback_query:
        q = update.callback_query
        await q.answer()
        if q.data.startswith("place_preset:"):
            idx = int(q.data.split(":")[1])
            places = get_recent_places(q.from_user.id)
            if idx < len(places):
                context.user_data["place"] = places[idx]
                await sticky_set(update, context, "⏰ Podaj godzinę rozpoczęcia pracy (HH:MM):")
                return START_TIME
        elif q.data == "place_manual":
            await sticky_set(update, context, "📍 Wyślij nazwę miejsca (tekst):")
            return PLACE
        return PLACE

    # wiadomość tekstowa
    try:
        await update.message.delete()
    except Exception:
        pass
    place_txt = (update.message.text or "").strip()
    if not place_txt:
        await sticky_set(update, context, "Podaj poprawne miejsce.")
        return PLACE
    context.user_data["place"] = place_txt
    await sticky_set(update, context, "⏰ Podaj godzinę rozpoczęcia pracy (HH:MM):")
    return START_TIME


async def start_time(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        await update.message.delete()
    except Exception:
        pass
    t = parse_time(update.message.text or "")
    last_end = context.user_data.get("entries", [])[-1]["end"] if context.user_data.get("entries") else None
    if not t or (last_end and t <= last_end):
        await sticky_set(update, context, "⏰ Błędna godzina. Spróbuj ponownie.")
        return START_TIME
    context.user_data["start"] = t
    await sticky_set(update, context, "⏰ Podaj godzinę zakończenia pracy (HH:MM):")
    return END_TIME


async def end_time(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        await update.message.delete()
    except Exception:
        pass
    t = parse_time(update.message.text or "")
    if not t or t <= context.user_data.get("start"):
        await sticky_set(update, context, "⏰ Błędna godzina. Spróbuj ponownie.")
        return END_TIME
    # walidacja nakładania
    uid = context.user_data.get("uid")
    date_str = context.user_data.get("date", today_str())
    start = context.user_data.get("start")
    overlap, conflicts = has_overlap(uid, date_str, start, t, in_memory=context.user_data.get("entries", []))
    if overlap:
        context.user_data["pending_end"] = t
        kb = [
            [InlineKeyboardButton("Kontynuuj mimo to", callback_data="overlap_ok")],
            [InlineKeyboardButton("Zmień godziny", callback_data="overlap_fix")],
        ]
        msg = "⚠️ Wykryto nakładanie z przedziałami: " + ", ".join([f"{a}-{b}" for a,b in conflicts])
        await sticky_set(update, context, msg, InlineKeyboardMarkup(kb))
        return OVERLAP_DECIDE

    context.user_data["end"] = t
    # pokaż szablony zadań
    templates = get_task_templates(uid)
    kb = []
    if templates:
        for i, tplt in enumerate(templates):
            label = tplt[:40] + ("…" if len(tplt) > 40 else "")
            kb.append([InlineKeyboardButton(label, callback_data=f"tpl_task:{i}")])
    kb.append([InlineKeyboardButton("Wpisz ręcznie", callback_data="tpl_manual")])
    await sticky_set(update, context, "📝 Wybierz szablon zadań lub wpisz ręcznie:", InlineKeyboardMarkup(kb))
    return TASKS


async def overlap_decide(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    if q.data == "overlap_ok":
        context.user_data["end"] = context.user_data.pop("pending_end")
        # przejdź do wyboru zadań
        templates = get_task_templates(context.user_data.get("uid"))
        kb = []
        if templates:
            for i, tplt in enumerate(templates):
                label = tplt[:40] + ("…" if len(tplt) > 40 else "")
                kb.append([InlineKeyboardButton(label, callback_data=f"tpl_task:{i}")])
        kb.append([InlineKeyboardButton("Wpisz ręcznie", callback_data="tpl_manual")])
        await sticky_set(update, context, "📝 Wybierz szablon zadań lub wpisz ręcznie:", InlineKeyboardMarkup(kb))
        return TASKS
    # popraw godziny
    await sticky_set(update, context, "⏰ Podaj godzinę rozpoczęcia pracy (HH:MM):")
    return START_TIME


async def tasks(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # obsługa wyboru szablonu
    if update.callback_query:
        q = update.callback_query
        await q.answer()
        if q.data.startswith("tpl_task:"):
            idx = int(q.data.split(":")[1])
            templates = get_task_templates(q.from_user.id)
            if idx < len(templates):
                context.user_data["tasks"] = templates[idx]
                await sticky_set(update, context, "💬 Dodaj uwagi lub wpisz '-' jeśli brak:")
                return NOTES
        elif q.data == "tpl_manual":
            await sticky_set(update, context, "📝 Wyślij listę zadań (tekst):")
            return TASKS
        return TASKS

    try:
        await update.message.delete()
    except Exception:
        pass
    txt = (update.message.text or "").strip()
    if not txt:
        await sticky_set(update, context, "📝 Lista zadań nie może być pusta.")
        return TASKS
    context.user_data["tasks"] = txt
    await sticky_set(update, context, "💬 Dodaj uwagi lub wpisz '-' jeśli brak:")
    return NOTES


async def notes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        await update.message.delete()
    except Exception:
        pass
    txt = (update.message.text or "").strip()
    if not txt:
        await sticky_set(update, context, "💬 Uwagi nie mogą być puste.")
        return NOTES

    entry = {
        "place": context.user_data.pop("place"),
        "start": context.user_data.pop("start"),
        "end": context.user_data.pop("end"),
        "tasks": context.user_data.pop("tasks"),
        "notes": txt,
    }
    context.user_data.setdefault("entries", []).append(entry)

    # zapamiętaj presety
    remember_place(context.user_data.get("uid"), entry["place"])
    remember_tasks(context.user_data.get("uid"), entry["tasks"])

    kb = [
        [InlineKeyboardButton("➕ Dodaj kolejne miejsce", callback_data="again")],
        [InlineKeyboardButton("✅ Zakończ raport", callback_data="finish")],
    ]
    await sticky_set(update, context, "Co dalej?", InlineKeyboardMarkup(kb))
    return ANOTHER


async def another(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == "again":
        # powrót do miejsca z presetami
        places = get_recent_places(query.from_user.id)
        kb = []
        if places:
            kb.extend([[InlineKeyboardButton(p, callback_data=f"place_preset:{i}")] for i, p in enumerate(places)])
        kb.append([InlineKeyboardButton("Wpisz ręcznie", callback_data="place_manual")])
        await sticky_set(update, context, "📍 Podaj miejsce wykonywania pracy:", InlineKeyboardMarkup(kb))
        return PLACE

    # finish → usuń sticky i pokaż finalny raport
    chat_id = query.message.chat.id

    save_report(
        context.user_data.get("entries", []),
        context.user_data.get("uid", query.from_user.id),
        context.user_data.get("date", today_str()),
        context.user_data.get("name", query.from_user.first_name),
    )

    final_entries = read_entries_for_day(query.from_user.id, context.user_data.get("date", today_str()))
    rpt = format_report(final_entries, context.user_data.get("date", today_str()), context.user_data.get("name", query.from_user.first_name), with_totals=True, uid=query.from_user.id)

    await sticky_delete(context, chat_id)
    msg = await query.message.chat.send_message(rpt)

    mapping = load_mapping()
    mapping[f"{query.from_user.id}_{context.user_data.get('date', today_str())}"] = msg.message_id
    save_mapping(mapping)
    context.user_data.clear()
    return ConversationHandler.END


# ────────────── FLOW: EDIT ──────────────
async def select_entry(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == "cancel_edit":
        await sticky_set(update, context, "Anulowano edycję.")
        return ConversationHandler.END

    idx = int(query.data.split(":" )[1])  # entry:{idx}
    context.user_data["edit_idx"] = idx
    e = context.user_data["edit_entries"][idx]

    kb = [
        [InlineKeyboardButton("Miejsce", callback_data="field:place")],
        [InlineKeyboardButton("Godzina start", callback_data="field:start")],
        [InlineKeyboardButton("Godzina koniec", callback_data="field:end")],
        [InlineKeyboardButton("Zadania", callback_data="field:tasks")],
        [InlineKeyboardButton("Uwagi", callback_data="field:notes")],
        [InlineKeyboardButton("↩️ Wybierz inną pozycję", callback_data="back_to_entries")],
    ]
    await sticky_set(update, context, f"Wybrano: #{idx+1} {e['place']} {e['start']}-{e['end']}\nCo edytować?", InlineKeyboardMarkup(kb))
    return SELECT_FIELD


async def select_field(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == "back_to_entries":
        entries = context.user_data.get("edit_entries", [])
        kb_rows = []
        for idx, e in enumerate(entries, start=1):
            label = f"#{idx} {e['place']}  {e['start']}-{e['end']}"
            kb_rows.append([InlineKeyboardButton(label, callback_data=f"entry:{idx-1}")])
        kb_rows.append([InlineKeyboardButton("↩️ Anuluj", callback_data="cancel_edit")])
        await sticky_set(update, context, "Wybierz pozycję:", InlineKeyboardMarkup(kb_rows))
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

    await sticky_set(update, context, prompt)
    return EDIT_VALUE


async def edit_value(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        await update.message.delete()
    except Exception:
        pass

    val = (update.message.text or "").strip()
    field = context.user_data.get("edit_field")
    idx = context.user_data.get("edit_idx")
    date_str = context.user_data.get("date", today_str())
    uid = context.user_data.get("uid")

    e = context.user_data.get("edit_entries", [])[idx]

    if field in ("start", "end"):
        t = parse_time(val)
        if not t:
            await sticky_set(update, context, "⏰ Błędny format. Użyj HH:MM.")
            return EDIT_VALUE
        start = t if field == "start" else str(e["start"]) or t
        end = t if field == "end" else str(e["end"]) or t
        if start and end and start >= end:
            await sticky_set(update, context, "⏰ Start musi być < koniec.")
            return EDIT_VALUE
        # sprawdź nakładanie z innymi wpisami
        overlap, conflicts = has_overlap(uid, date_str, start, end, exclude_rid=e["rid"]) 
        if overlap:
            await sticky_set(update, context, "⚠️ Przedziały czasu nakładają się: " + ", ".join([f"{a}-{b}" for a,b in conflicts]) + ". Zmień godziny.")
            return EDIT_VALUE
        val = t

    try:
        update_report_field(uid, date_str, e["rid"], field, val)
    except Exception as ex:
        await sticky_set(update, context, f"❌ Błąd zapisu: {ex}")
        return EDIT_VALUE

    # odśwież lokalny cache wpisów
    context.user_data["edit_entries"] = read_entries_for_day(uid, date_str)

    kb = [
        [InlineKeyboardButton("Edytuj inne pole tej pozycji", callback_data="again_same")],
        [InlineKeyboardButton("Edytuj inną pozycję", callback_data="again_other")],
        [InlineKeyboardButton("Pokaż raport i zakończ", callback_data="finish_edit")],
    ]
    await sticky_set(update, context, "Zmieniono. Edytować coś jeszcze?", InlineKeyboardMarkup(kb))
    return EDIT_MORE


async def edit_more(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == "again_same":
        idx = context.user_data.get("edit_idx")
        e = context.user_data.get("edit_entries", [])[idx]
        kb = [
            [InlineKeyboardButton("Miejsce", callback_data="field:place")],
            [InlineKeyboardButton("Godzina start", callback_data="field:start")],
            [InlineKeyboardButton("Godzina koniec", callback_data="field:end")],
            [InlineKeyboardButton("Zadania", callback_data="field:tasks")],
            [InlineKeyboardButton("Uwagi", callback_data="field:notes")],
            [InlineKeyboardButton("↩️ Wybierz inną pozycję", callback_data="back_to_entries")],
        ]
        await sticky_set(update, context, f"Wybrano: #{idx+1} {e['place']} {e['start']}-{e['end']}\nCo edytować?", InlineKeyboardMarkup(kb))
        return SELECT_FIELD

    if query.data == "again_other":
        entries = context.user_data.get("edit_entries", [])
        kb_rows = []
        for idx, e in enumerate(entries, start=1):
            label = f"#{idx} {e['place']}  {e['start']}-{e['end']}"
            kb_rows.append([InlineKeyboardButton(label, callback_data=f"entry:{idx-1}")])
        kb_rows.append([InlineKeyboardButton("↩️ Anuluj", callback_data="cancel_edit")])
        await sticky_set(update, context, "Wybierz pozycję:", InlineKeyboardMarkup(kb_rows))
        return SELECT_ENTRY

    # finish_edit → usuń sticky i pokaż gotowy raport z podsumowaniami
    chat_id = query.message.chat.id
    date_str = context.user_data.get("date", today_str())
    uid = context.user_data.get("uid")
    entries = read_entries_for_day(uid, date_str)
    name = entries[0]["name"] if entries else context.user_data.get("name", query.from_user.first_name)
    rpt = format_report(entries, date_str, name, with_totals=True, uid=uid)

    await sticky_delete(context, chat_id)
    msg = await query.message.chat.send_message(rpt)

    mapping = load_mapping()
    mapping[f"{uid}_{date_str}"] = msg.message_id
    save_mapping(mapping)
    context.user_data.clear()
    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        await sticky_delete(context, update.effective_chat.id)
    except Exception:
        pass
    await update.effective_chat.send_message("Anulowano.")
    context.user_data.clear()
    return ConversationHandler.END


async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.effective_chat.send_message(
        "Użyj /start aby otworzyć menu. W menu: zmiana daty (kalendarz), "
        "tworzenie, edycja, eksporty. Tworzenie ma presety miejsc i szablony zadań. "
        "Walidacja nakładania godzin. Po zakończeniu pokazuję sumy dzienne i tygodniowe."
    )


# ──────────────────── PTB Application ────────────────────
async def on_startup(app: Application) -> None:
    await app.bot.set_my_commands([
        BotCommand("start", "Otwórz menu raportów"),
        BotCommand("export", "Eksport (admin): /export YYYY-MM"),
        BotCommand("myexport", "Mój eksport: /myexport YYYY-MM"),
        BotCommand("help", "Pomoc"),
    ])


def build_app() -> Application:
    app = (
        ApplicationBuilder()
        .token(TELEGRAM_TOKEN)
        .post_init(on_startup)
        .build()
    )

    # komendy
    app.add_handler(CommandHandler("start", show_menu))
    app.add_handler(CommandHandler("export", export_handler))
    app.add_handler(CommandHandler("myexport", myexport_handler))
    app.add_handler(CommandHandler("help", help_cmd))

    # menu/top-level callbacks (poza conversation)
    app.add_handler(CallbackQueryHandler(change_date_cb, pattern=r"^change_date$"))
    app.add_handler(CallbackQueryHandler(calendar_nav_cb, pattern=r"^(cal:\d{4}-\d{2}|day:\d{2}\.\d{2}\.\d{4})$"))

    # conversation
    conv = ConversationHandler(
        entry_points=[CallbackQueryHandler(menu_handler, pattern=r"^(create|edit|export|myexport)$")],
        states={
            # wybór daty (kalendarz inline)
            DATE_PICK: [
                CallbackQueryHandler(calendar_nav_cb, pattern=r"^(cal:\d{4}-\d{2}|day:\d{2}\.\d{2}\.\d{4})$")
            ],
            # tworzenie
            PLACE: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, place),
                CallbackQueryHandler(place, pattern=r"^(place_preset:\d+|place_manual)$"),
            ],
            START_TIME: [MessageHandler(filters.TEXT & ~filters.COMMAND, start_time)],
            END_TIME:   [MessageHandler(filters.TEXT & ~filters.COMMAND, end_time)],
            OVERLAP_DECIDE: [CallbackQueryHandler(overlap_decide, pattern=r"^(overlap_ok|overlap_fix)$")],
            TASKS: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, tasks),
                CallbackQueryHandler(tasks, pattern=r"^(tpl_task:\d+|tpl_manual)$"),
            ],
            NOTES:   [MessageHandler(filters.TEXT & ~filters.COMMAND, notes)],
            ANOTHER: [CallbackQueryHandler(another, pattern=r"^(again|finish)$")],
            # edycja
            SELECT_ENTRY: [CallbackQueryHandler(select_entry, pattern=r"^(entry:\d+|cancel_edit)$")],
            SELECT_FIELD: [CallbackQueryHandler(select_field, pattern=r"^(field:(place|start|end|tasks|notes)|back_to_entries)$")],
            EDIT_VALUE:   [MessageHandler(filters.TEXT & ~filters.COMMAND, edit_value)],
            EDIT_MORE:    [CallbackQueryHandler(edit_more, pattern=r"^(again_same|again_other|finish_edit)$")],
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
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
    )

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
