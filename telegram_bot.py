"""
Telegram Bot for Site Data Lookup (PythonAnywhere Ready)
========================================================
"""

import os
import logging
from pathlib import Path

import openpyxl
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    filters,
    ContextTypes,
)

# ──────────────────────────────────────────────
# CONFIG — Token directly set for easy deployment
# ──────────────────────────────────────────────
BOT_TOKEN = "8713734794:AAGGfGJkANJlzPp139__NXEsh_0VvgayVP4"
EXCEL_FILE = Path(__file__).parent / "site_data.xlsx"
HEADER_ROW = 3
DATA_START_ROW = 4

# Column indices (0-based)
COL_SN = 0
COL_GENERIC_ID = 1
COL_2G_SITE_ID = 2
COL_4G_SITE_ID = 3
COL_ADDRESS = 5
COL_THANA = 6
COL_KEY_PROFILE = 9
COL_LOCK = 12
COL_LAT = 14
COL_LONG = 15

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)


# ──────────────────────────────────────────────
# LOAD EXCEL DATA
# ──────────────────────────────────────────────
def load_site_data() -> list[dict]:
    wb = openpyxl.load_workbook(EXCEL_FILE, read_only=True, data_only=True)
    ws = wb["Site Data"]
    sites = []
    for row in ws.iter_rows(min_row=DATA_START_ROW, values_only=True):
        cells = list(row)
        if not cells or cells[COL_GENERIC_ID] is None:
            continue
        site = {
            "sn": str(cells[COL_SN] or ""),
            "generic_id": str(cells[COL_GENERIC_ID] or "").strip(),
            "site_id_2g": str(cells[COL_2G_SITE_ID] or "").strip(),
            "site_id_4g": str(cells[COL_4G_SITE_ID] or "").strip(),
            "address": str(cells[COL_ADDRESS] or "").strip(),
            "thana": str(cells[COL_THANA] or "").strip(),
            "key_profile": str(cells[COL_KEY_PROFILE] or "").strip(),
            "lock": str(cells[COL_LOCK] or "").strip(),
            "lat": str(cells[COL_LAT] or "").strip(),
            "long": str(cells[COL_LONG] or "").strip(),
        }
        sites.append(site)
    wb.close()
    logger.info("Loaded %d sites from Excel", len(sites))
    return sites


SITES: list[dict] = []


def search_sites(query: str) -> list[dict]:
    q = query.strip().upper()
    results = []
    for site in SITES:
        if (
            q in site["generic_id"].upper()
            or q in site["site_id_2g"].upper()
            or q in site["site_id_4g"].upper()
            or q in site["address"].upper()
            or q in site["thana"].upper()
        ):
            results.append(site)
    return results


# ──────────────────────────────────────────────
# FORMAT SITE INFO
# ──────────────────────────────────────────────
def format_site(site: dict) -> str:
    lines = [
        "\U0001f4e1 *Site Information*",
        "\u2501" * 20,
        f"\U0001f194 *Generic ID:* `{site['generic_id']}`",
        f"\U0001f4f6 *2G Site ID:* `{site['site_id_2g']}`",
        f"\U0001f4f6 *4G Site ID:* `{site['site_id_4g']}`",
        f"\U0001f4cd *Address:* {site['address']}",
        f"\U0001f3d8\ufe0f *Thana:* {site['thana']}",
        f"\U0001f511 *Key Profile:* `{site['key_profile']}`",
        f"\U0001f512 *Lock:* `{site['lock']}`",
        f"\U0001f310 *Latitude:* `{site['lat']}`",
        f"\U0001f310 *Longitude:* `{site['long']}`",
        "\u2501" * 20,
    ]
    return "\n".join(lines)


def get_map_button(site: dict) -> InlineKeyboardMarkup | None:
    lat = site["lat"]
    lng = site["long"]
    if lat and lng and lat != "None" and lng != "None":
        url = f"https://www.google.com/maps?q={lat},{lng}"
        keyboard = [
            [InlineKeyboardButton("\U0001f5fa\ufe0f Open in Google Maps", url=url)],
            [InlineKeyboardButton("\U0001f4cd Send Location", callback_data=f"loc_{lat}_{lng}")],
        ]
        return InlineKeyboardMarkup(keyboard)
    return None


# ──────────────────────────────────────────────
# HANDLERS
# ──────────────────────────────────────────────
async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    welcome = (
        "\U0001f916 *Site Data Bot*\n\n"
        "Assalamu Alaikum! I am Site Data Bot.\n\n"
        "\U0001f4cc *How to use:*\n"
        "\u2022 Send any *Site ID* (Generic / 2G / 4G)\n"
        "  Example: `DHK0005` or `DHK_X0005`\n\n"
        "\u2022 Search by *Address or Thana*\n"
        "  Example: `Gulshan` or `Banani`\n\n"
        "\U0001f4ca *Commands:*\n"
        "/start - Start bot\n"
        "/help - Help\n"
        "/total - Total sites\n"
        "/search `<keyword>` - Search site\n\n"
        f"\U0001f4e6 Database: *{len(SITES)}* sites loaded"
    )
    await update.message.reply_text(welcome, parse_mode="Markdown")


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    help_text = (
        "\u2139\ufe0f *Help*\n\n"
        "\U0001f50d *How to search:*\n"
        "1\ufe0f\u20e3 Type Site ID: `DHK0005`\n"
        "2\ufe0f\u20e3 Use /search: `/search Gulshan`\n"
        "3\ufe0f\u20e3 Type Thana name: `Banani`\n\n"
        "\U0001f4cd *Map Features:*\n"
        "\u2022 \U0001f5fa\ufe0f *Open in Google Maps* - Opens map in browser\n"
        "\u2022 \U0001f4cd *Send Location* - Sends Telegram location pin\n\n"
        "\u26a1 Max 10 results shown per search"
    )
    await update.message.reply_text(help_text, parse_mode="Markdown")


async def total_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        f"\U0001f4ca Total *{len(SITES)}* sites in database.",
        parse_mode="Markdown",
    )


async def search_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not context.args:
        await update.message.reply_text(
            "\u274c Please provide a search term.\nExample: `/search DHK0005`",
            parse_mode="Markdown",
        )
        return
    query = " ".join(context.args)
    await handle_search(update, query)


async def text_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.message.text.strip()
    if query:
        await handle_search(update, query)


async def handle_search(update: Update, query: str):
    results = search_sites(query)
    if not results:
        await update.message.reply_text(
            f"\u274c *No site found*\n\nSearch: `{query}`\n\n"
            "\U0001f4a1 Try a full or partial Site ID or Address.",
            parse_mode="Markdown",
        )
        return

    total_found = len(results)
    results = results[:10]

    if total_found > 10:
        await update.message.reply_text(
            f"\U0001f50d *{total_found}* sites found. Showing first *10*...\n"
            "Use a more specific keyword to narrow results.",
            parse_mode="Markdown",
        )

    for site in results:
        msg = format_site(site)
        keyboard = get_map_button(site)
        await update.message.reply_text(msg, parse_mode="Markdown", reply_markup=keyboard)


async def location_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    if data and data.startswith("loc_"):
        parts = data.split("_")
        if len(parts) >= 3:
            try:
                lat = float(parts[1])
                lng = float(parts[2])
                await query.message.reply_location(latitude=lat, longitude=lng)
            except (ValueError, IndexError):
                await query.message.reply_text("Invalid location data")


# ──────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────
def main():
    global SITES
    SITES = load_site_data()

    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start_command))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("total", total_command))
    app.add_handler(CommandHandler("search", search_command))
    app.add_handler(CallbackQueryHandler(location_callback, pattern=r"^loc_"))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_message))

    print("Bot is running...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
