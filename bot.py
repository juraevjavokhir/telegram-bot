"""
Telegram-бот "Заявка для Китая"
================================
Собирает у пользователя список запчастей (фото + наименование + артикул +
количество + единица измерения) и по кнопке "Готово" формирует Excel-файл
заявки (колонки: №, Фото, Наименование, Артикул, Количество, Ед. изм.,
Цена, Сумма) и сразу отправляет его в чат.

Запуск:
    pip install -r requirements.txt
    export BOT_TOKEN="<токен от @BotFather>"
    python bot.py

Подробности — в README.md.
"""

import logging
import math
import os
import shutil
import tempfile
import uuid
from datetime import datetime

from PIL import Image as PILImage
from PIL import ImageOps
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from telegram import (
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    ReplyKeyboardRemove,
    Update,
)
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    filters,
)

# --------------------------------------------------------------------------
# Настройки
# --------------------------------------------------------------------------

BOT_TOKEN = os.environ.get("BOT_TOKEN", "")
# Если api.telegram.org заблокирован в вашей сети — задайте PROXY_URL,
# например: PROXY_URL=socks5://127.0.0.1:1080 или http://127.0.0.1:8080
PROXY_URL = os.environ.get("PROXY_URL", "")

# Состояния диалога
WAITING_PHOTO, WAITING_NAME, WAITING_ARTICLE, WAITING_QTY, WAITING_UNIT = range(5)

DEFAULT_UNIT = "шт"

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)
# httpx пишет в INFO каждый URL запроса, а в нём — токен бота
logging.getLogger("httpx").setLevel(logging.WARNING)


# --------------------------------------------------------------------------
# Вспомогательные функции для работы с сессией пользователя
# --------------------------------------------------------------------------

def _session(context: ContextTypes.DEFAULT_TYPE) -> dict:
    """Возвращает (создавая при необходимости) данные текущей заявки."""
    if "items" not in context.user_data:
        context.user_data["items"] = []
    if "tmp_dir" not in context.user_data:
        context.user_data["tmp_dir"] = tempfile.mkdtemp(prefix="zayavka_")
    return context.user_data


def _current_item(context: ContextTypes.DEFAULT_TYPE) -> dict:
    if "draft" not in context.user_data:
        context.user_data["draft"] = {}
    return context.user_data["draft"]


def _cleanup(context: ContextTypes.DEFAULT_TYPE) -> None:
    """Удаляет временные файлы (фото) и очищает сессию."""
    tmp_dir = context.user_data.get("tmp_dir")
    if tmp_dir and os.path.isdir(tmp_dir):
        shutil.rmtree(tmp_dir, ignore_errors=True)
    context.user_data.clear()


DONE_KEYBOARD = InlineKeyboardMarkup(
    [
        [
            InlineKeyboardButton("➕ Добавить ещё", callback_data="add_more"),
            InlineKeyboardButton("✅ Готово", callback_data="finish"),
        ]
    ]
)


# --------------------------------------------------------------------------
# Диалог: /start -> фото -> название -> количество -> ед.изм. -> повтор/готово
# --------------------------------------------------------------------------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    _cleanup(context)
    _session(context)
    await update.message.reply_text(
        "Собираем заявку для партнёра из Китая.\n\n"
        "Отправьте фото 1-й запчасти.\n"
        "В любой момент можно отменить командой /cancel.",
        reply_markup=ReplyKeyboardRemove(),
    )
    return WAITING_PHOTO


async def add_more(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()
    await query.edit_message_reply_markup(reply_markup=None)
    await query.message.reply_text("Отправьте фото следующей запчасти.")
    return WAITING_PHOTO


async def receive_photo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    session = _session(context)
    photo = update.message.photo[-1]  # самое большое разрешение
    tg_file = await photo.get_file()

    local_path = os.path.join(session["tmp_dir"], f"{uuid.uuid4().hex}.jpg")
    await tg_file.download_to_drive(local_path)

    draft = _current_item(context)
    draft.clear()
    draft["photo_path"] = local_path

    await update.message.reply_text("Введите наименование детали:")
    return WAITING_NAME


async def photo_missing(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text(
        "Пожалуйста, отправьте именно фото (как изображение, не файлом)."
    )
    return WAITING_PHOTO


async def receive_name(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    name = (update.message.text or "").strip()
    if not name:
        await update.message.reply_text("Наименование не может быть пустым. Введите ещё раз:")
        return WAITING_NAME

    draft = _current_item(context)
    draft["name"] = name
    await update.message.reply_text(
        "Введите код товара (артикул).\n"
        "Если артикула нет — отправьте «-»."
    )
    return WAITING_ARTICLE


async def receive_article(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    article = (update.message.text or "").strip()
    if article == "-":
        article = ""

    draft = _current_item(context)
    draft["article"] = article

    await update.message.reply_text("Введите количество (числом):")
    return WAITING_QTY


async def receive_qty(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    raw = (update.message.text or "").strip().replace(",", ".")
    try:
        qty = float(raw)
        # nan/inf проходят float(), но ломают int(qty) ниже
        if not math.isfinite(qty) or qty <= 0:
            raise ValueError
    except ValueError:
        await update.message.reply_text(
            "Не похоже на количество. Введите число, например: 10"
        )
        return WAITING_QTY

    qty_display = int(qty) if qty == int(qty) else qty
    draft = _current_item(context)
    draft["qty"] = qty_display

    await update.message.reply_text(
        f"Единица измерения? (например: шт, компл, м)\n"
        f"Отправьте «-», чтобы поставить значение по умолчанию «{DEFAULT_UNIT}»."
    )
    return WAITING_UNIT


async def receive_unit(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    unit = (update.message.text or "").strip()
    if unit in ("", "-"):
        unit = DEFAULT_UNIT

    draft = _current_item(context)
    draft["unit"] = unit

    session = _session(context)
    session["items"].append(dict(draft))
    draft.clear()

    n = len(session["items"])
    await update.message.reply_text(
        f"Добавлено. Всего позиций в заявке: {n}.",
        reply_markup=DONE_KEYBOARD,
    )
    return WAITING_PHOTO


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    _cleanup(context)
    await update.message.reply_text(
        "Заявка отменена. Чтобы начать заново — /start.",
        reply_markup=ReplyKeyboardRemove(),
    )
    return ConversationHandler.END


# --------------------------------------------------------------------------
# Формирование и отправка Excel
# --------------------------------------------------------------------------

HEADER_FILL = PatternFill(start_color="1FA398", end_color="1FA398", fill_type="solid")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=11)
BODY_FONT = Font(name="Arial", size=10)
THIN = Side(style="thin", color="B7B7B7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

HEADERS = ["№", "Фото", "Наименование", "Артикул", "Количество", "Ед. изм.", "Цена", "Сумма"]
COL_WIDTHS = {"A": 6, "C": 34, "D": 16, "E": 12, "F": 10, "G": 12, "H": 12}
N_COLS = len(HEADERS)  # 8

# Точный размер фото в ячейке: 3 см (ширина) x 2.5 см (высота)
CM_TO_PX = 96 / 2.54          # пикселей в 1 см при 96 dpi (Excel/openpyxl)
CM_TO_PT = 72 / 2.54          # points в 1 см (высота строки задаётся в points)
PHOTO_W_CM = 3.0
PHOTO_H_CM = 2.5
PHOTO_W_PX = PHOTO_W_CM * CM_TO_PX          # ~113.4 px — реальный размер в Excel
PHOTO_H_PX = PHOTO_H_CM * CM_TO_PX          # ~94.5 px
ROW_HEIGHT_PT = PHOTO_H_CM * CM_TO_PT       # ~70.9 pt — высота строки под фото
COL_B_WIDTH = (PHOTO_W_PX - 5) / 7          # ширина колонки B в Excel-единицах
UPSCALE = 4                                  # запас разрешения для чёткости при печати/зуме
THUMB_W_PX = round(PHOTO_W_PX * UPSCALE)
THUMB_H_PX = round(PHOTO_H_PX * UPSCALE)


def build_workbook(items: list[dict], out_path: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Заявка"

    # Заголовок листа
    last_col_letter = get_column_letter(N_COLS)
    ws.merge_cells(f"A1:{last_col_letter}1")
    ws["A1"] = f"Заявка на запчасти — {datetime.now().strftime('%d.%m.%Y')}"
    ws["A1"].font = Font(name="Arial", bold=True, size=13)
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 24

    header_row = 2
    for col_idx, title in enumerate(HEADERS, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=title)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = BORDER

    for col_letter, width in COL_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width
    ws.column_dimensions["B"].width = COL_B_WIDTH  # под фото ровно 3 см

    first_data_row = header_row + 1
    for i, item in enumerate(items):
        row = first_data_row + i

        ws.cell(row=row, column=1, value=i + 1)
        # B: фото — заполняется картинкой ниже
        ws.cell(row=row, column=3, value=item["name"])
        ws.cell(row=row, column=4, value=item.get("article", ""))
        ws.cell(row=row, column=5, value=item["qty"])
        ws.cell(row=row, column=6, value=item["unit"])
        # G: цена — пусто, заполнит партнёр
        ws.cell(row=row, column=7, value=None)
        # H: сумма — формула, посчитается сама, когда впишут цену
        ws.cell(row=row, column=8, value=f"=E{row}*G{row}")

        for col in range(1, N_COLS + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = BORDER
            cell.alignment = CENTER if col != 3 else LEFT
            cell.font = BODY_FONT
        ws.cell(row=row, column=7).number_format = "#,##0.00"
        ws.cell(row=row, column=8).number_format = "#,##0.00"

        ws.row_dimensions[row].height = ROW_HEIGHT_PT

        # Вставляем фото ровно 3x2.5 см, кадрируя по центру под нужные
        # пропорции и сохраняя высокое разрешение для чёткости при печати.
        photo_path = item.get("photo_path")
        if photo_path and os.path.exists(photo_path):
            try:
                with PILImage.open(photo_path) as im:
                    im = ImageOps.exif_transpose(im).convert("RGB")
                    im = ImageOps.fit(
                        im,
                        (THUMB_W_PX, THUMB_H_PX),
                        method=PILImage.LANCZOS,
                    )
                    resized_path = photo_path + ".thumb.jpg"
                    im.save(resized_path, format="JPEG", quality=95)
                xl_img = XLImage(resized_path)
                xl_img.width = PHOTO_W_PX
                xl_img.height = PHOTO_H_PX
                ws.add_image(xl_img, f"B{row}")
            except Exception:
                logger.exception("Не удалось вставить фото для строки %s", row)

    # Итоговая строка
    total_row = first_data_row + len(items)
    ws.cell(row=total_row, column=3, value="Итого:")
    ws.cell(row=total_row, column=3).font = Font(name="Arial", bold=True, size=10)
    ws.cell(row=total_row, column=3).alignment = LEFT
    total_cell = ws.cell(
        row=total_row,
        column=8,
        value=f"=SUM(H{first_data_row}:H{total_row - 1})" if items else 0,
    )
    total_cell.font = Font(name="Arial", bold=True, size=10)
    total_cell.number_format = "#,##0.00"
    for col in range(1, N_COLS + 1):
        ws.cell(row=total_row, column=col).border = BORDER

    # Легенда для партнёра
    legend_row = total_row + 2
    ws.cell(
        row=legend_row,
        column=1,
        value="Жёлтым — заполняет партнёр (цена за ед.). Сумма считается автоматически.",
    ).font = Font(name="Arial", italic=True, size=9, color="808080")
    ws.merge_cells(start_row=legend_row, start_column=1, end_row=legend_row, end_column=N_COLS)

    yellow = PatternFill(start_color="FFF2A8", end_color="FFF2A8", fill_type="solid")
    for i in range(len(items)):
        ws.cell(row=first_data_row + i, column=7).fill = yellow

    ws.freeze_panes = "A3"
    wb.save(out_path)


async def finish(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()
    await query.edit_message_reply_markup(reply_markup=None)

    session = _session(context)
    items = session.get("items", [])

    if not items:
        await query.message.reply_text(
            "Список пуст — нечего сохранять. Отправьте фото первой детали или /cancel."
        )
        return WAITING_PHOTO

    await query.message.reply_text("Формирую Excel-файл…")

    out_path = os.path.join(session["tmp_dir"], "zayavka.xlsx")
    build_workbook(items, out_path)

    date_str = datetime.now().strftime("%Y-%m-%d")
    with open(out_path, "rb") as f:
        await query.message.reply_document(
            document=f,
            filename=f"Заявка_{date_str}.xlsx",
            caption=f"Готово ✅ Позиций: {len(items)}",
        )

    _cleanup(context)
    return ConversationHandler.END


async def unknown_text_in_photo_state(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text(
        "Жду фото следующей детали, либо нажмите «Готово» выше, либо /cancel."
    )
    return WAITING_PHOTO


# --------------------------------------------------------------------------
# main
# --------------------------------------------------------------------------

def main() -> None:
    if not BOT_TOKEN:
        raise SystemExit(
            "Не задан BOT_TOKEN. Установите переменную окружения BOT_TOKEN "
            "(токен от @BotFather) и запустите снова."
        )

    builder = Application.builder().token(BOT_TOKEN)
    if PROXY_URL:
        logger.info("Использую прокси: %s", PROXY_URL)
        builder = builder.proxy(PROXY_URL).get_updates_proxy(PROXY_URL)
    application = builder.build()

    conv = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            WAITING_PHOTO: [
                MessageHandler(filters.PHOTO, receive_photo),
                CallbackQueryHandler(add_more, pattern="^add_more$"),
                CallbackQueryHandler(finish, pattern="^finish$"),
                MessageHandler(filters.TEXT & ~filters.COMMAND, unknown_text_in_photo_state),
                MessageHandler(~filters.PHOTO & ~filters.TEXT & ~filters.COMMAND, photo_missing),
            ],
            WAITING_NAME: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_name),
            ],
            WAITING_ARTICLE: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_article),
            ],
            WAITING_QTY: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_qty),
            ],
            WAITING_UNIT: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_unit),
            ],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        allow_reentry=True,  # /start посреди диалога начинает заявку заново
    )

    application.add_handler(conv)

    logger.info("Бот запущен, ожидаю сообщения…")
    application.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
