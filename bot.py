#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import re, json, logging, os
from pathlib import Path
from datetime import datetime
import pandas as pd
import pdfplumber
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from telegram import Update, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes

# ─── الإعدادات ──────────────────────────────────────────────────────────
BOT_TOKEN  = "8792112088:AAHPOQKqDcz2ULmm-uH3o4GgG73iAPGMCP4"
ADMIN_ID   = 45138916
BASE_DIR   = Path(__file__).parent
EXCEL      = BASE_DIR / "ارقام_الدواجن___شركات_-_افراد__.xlsx"
USERS_FILE = BASE_DIR / "users.json"
TEMP_DIR   = BASE_DIR / "temp"
TEMP_DIR.mkdir(exist_ok=True)

logging.basicConfig(level=logging.INFO)

# تنسيقات الإكسيل
FILL_ROW  = PatternFill("solid", fgColor="F1F8E9")
FONT_MAIN = Font(name="Arial", size=11, color="212121")
ALIGN_C   = Alignment(horizontal="center", vertical="center")
ALIGN_R   = Alignment(horizontal="right",  vertical="center")

# ─── إدارة المستخدمين والقوائم ──────────────────────────────────────────
def get_users():
    if USERS_FILE.exists():
        try: return json.loads(USERS_FILE.read_text(encoding="utf-8"))
        except: return [ADMIN_ID]
    return [ADMIN_ID]

def ok(uid):
    return uid in get_users()

def get_main_menu():
    keyboard = [
        [KeyboardButton("/stats"), KeyboardButton("/getfile")],
        [KeyboardButton("/duplicates"), KeyboardButton("/users")],
        [KeyboardButton("/help")]
    ]
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

# ─── محرك الاستخراج والجرد ─────────────────────────────────────────────

def get_phone(text):
    if not text or text == 'nan': return None
    clean_text = re.sub(r'\D', '', str(text))
    m = re.search(r'(?:966|0)?(5\d{8})', clean_text)
    if m: return '0' + m.group(1)
    return None

def is_co(text):
    keywords = ["شركة","مؤسسة","مصنع","توزيع","تسويق","مندوب","مبيعات","وكيل","موزع","مشرف","للتجارة","المحدودة"]
    return any(k in str(text) for k in keywords)

# ─── تحديث الإكسيل والملخص (الجرد الشامل) ──────────────────────────────

def update_excel(records, source_type):
    wb = load_workbook(EXCEL)
    ws_a, ws_s, ws_m = wb['أفراد'], wb['شركات'], wb['الملخص']
    
    # منع التكرار
    exist = set()
    for row in ws_a.iter_rows(min_row=2, max_col=2, values_only=True):
        if row[1]: exist.add(str(row[1])[-9:])
    for row in ws_s.iter_rows(min_row=2, max_col=4, values_only=True):
        if row[3]: exist.add(str(row[3])[-9:])

    new_a, new_s, dups = 0, 0, 0
    for r in records:
        ph = r['phone']
        if ph[-9:] in exist:
            dups += 1
            continue
        
        exist.add(ph[-9:])
        if is_co(r['text']):
            row_idx = ws_s.max_row + 1
            ws_s.append([row_idx - 1, r['text'][:40], "مستخرج", ph])
            new_s += 1
        else:
            row_idx = ws_a.max_row + 1
            ws_a.append([row_idx - 1, ph, r['text'][:40], "دواجن", "غير محدد", "", r['date'], r['text'][:100], source_type])
            new_a += 1

    # الجرد الفعلي للملف كامل
    total_a = ws_a.max_row - 1
    total_s = ws_s.max_row - 1
    total_all = total_a + total_s

    # تحديث شيت الملخص
    ws_m['C5'] = total_a
    ws_m['C6'] = total_s
    ws_m['C7'] = total_all

    wb.save(EXCEL)
    return new_a, new_s, dups, total_all

# ─── معالجة الرسائل والملفات ──────────────────────────────────────────

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not ok(update.effective_user.id): return
    msg = (
        "🐔 *بوت دواجن*\n\n"
        "أرسل ZIP أو TXT تصدير واتساب، أو ملفات Excel/PDF!\n\n"
        "استخدم القائمة بالأسفل للتحكم 👇"
    )
    await update.message.reply_text(msg, parse_mode="Markdown", reply_markup=get_main_menu())

async def handle_any_input(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not ok(update.effective_user.id): return
    
    source_name = "نص مباشر"
    raw_text = ""

    if update.message.document:
        doc = update.message.document
        source_name = doc.file_name
        f = await ctx.bot.get_file(doc.file_id)
        fp = TEMP_DIR / doc.file_name
        await f.download_to_drive(fp)
        
        if fp.suffix.lower() == '.pdf':
            with pdfplumber.open(fp) as pdf:
                raw_text = "\n".join([p.extract_text() or "" for p in pdf.pages])
        else:
            raw_text = fp.read_text(encoding='utf-8', errors='ignore')
        fp.unlink()
    else:
        raw_text = update.message.text

    recs = [{'phone': p, 'text': line, 'date': datetime.now().strftime('%d.%m.%Y')} 
            for line in raw_text.splitlines() if (p := get_phone(line))]

    if recs:
        na, ns, d, total = update_excel(recs, source_name)
        res = (
            "✅ *تم الاستخراج بنجاح!*\n\n"
            f"📝 النوع: {source_name[:20]}\n"
            f"👤 أفراد: {na} | 🏢 شركات: {ns}\n"
            f"🔁 مكرر: {d} | 📊 الإجمالي: {total:,}"
        )
        await update.message.reply_text(res, parse_mode="Markdown", reply_markup=get_main_menu())
    else:
        await update.message.reply_text("⚠️ لم يتم العثور على أرقام جديدة.")

async def cmd_getfile(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if ok(update.effective_user.id):
        await update.message.reply_document(document=open(EXCEL, 'rb'), caption="📊 قاعدة البيانات المحدثة")

async def cmd_stats(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if ok(update.effective_user.id):
        wb = load_workbook(EXCEL)
        total = (wb['أفراد'].max_row - 1) + (wb['شركات'].max_row - 1)
        await update.message.reply_text(f"📈 *إحصائيات حالية:*\nالإجمالي في الملف: {total:,} رقم", parse_mode="Markdown")

# ─── التشغيل ──────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("getfile", cmd_getfile))
    app.add_handler(CommandHandler("stats", cmd_stats))
    app.add_handler(MessageHandler(filters.ALL & (~filters.COMMAND), handle_any_input))
    print("🚀 البوت يعمل بالقائمة الجديدة...")
    app.run_polling()
