#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import re, zipfile, json, logging
from pathlib import Path
from datetime import datetime
from collections import Counter
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes

# --- الإعدادات ---
BOT_TOKEN  = "8792112088:AAHPOQKqDcz2ULmm-uH3o4GgG73iAPGMCP4"
ADMIN_ID   = 45138916
BASE_DIR   = Path(__file__).parent
EXCEL      = BASE_DIR / "ارقام_الدواجن___شركات_-_افراد__.xlsx"
USERS_FILE = BASE_DIR / "users.json"
TEMP_DIR   = BASE_DIR / "temp"
TEMP_DIR.mkdir(exist_ok=True)
logging.basicConfig(level=logging.WARNING)

FILL_ROW  = PatternFill("solid", fgColor="F1F8E9")
FONT_MAIN = Font(name="Arial", size=11, color="212121")
ALIGN_C   = Alignment(horizontal="center", vertical="center")
ALIGN_R   = Alignment(horizontal="right",  vertical="center")

# ─── المستخدمون ───────────────────────────────────────────────────────────────
def get_users():
    if USERS_FILE.exists():
        try: return json.loads(USERS_FILE.read_text(encoding="utf-8"))
        except: return [ADMIN_ID]
    return [ADMIN_ID]

def set_users(u):
    USERS_FILE.write_text(json.dumps(u), encoding="utf-8")

def ok(uid):
    return uid in get_users()

# ─── الكلمات المفتاحية ────────────────────────────────────────────────────────
COMPANY_KW = ["شركة","مؤسسة","مصنع","توزيع","تسويق","مندوب","مبيعات",
               "وكيل","موزع","مشرف","للتجارة","المحدودة","شركه","مؤسسه"]

REGIONS = ["الرياض","جدة","مكة","المدينة","الدمام","الخبر","الخرج",
           "القصيم","بريدة","حائل","تبوك","أبها","خميس مشيط","نجران",
           "جيزان","الطائف","ينبع","الأحساء","القطيف","الجبيل","عنيزة"]

def get_phone(text):
    m = re.search(r'05\d{8}', text)
    if m: return m.group()
    m = re.search(r'(?:\+966|00966)5\d{8}', text)
    if m:
        n = re.sub(r'\D','', m.group())
        n = re.sub(r'^00966|^966','', n)
        return '0' + n
    return None

def get_name(text):
    m = re.search(r'~\s*([^:]+):', text)
    if m: return m.group(1).strip()[:40]
    m = re.search(r'[أا]بو\s+\w+', text)
    if m: return m.group().strip()
    return "غير محدد"

def get_region(text):
    for r in REGIONS:
        if r in text: return r
    return "غير محدد"

def is_co(text):
    return any(k in text for k in COMPANY_KW)

# ─── تحديث الإكسيل ────────────────────────────────────────────────────────────
def update_excel(records, source):
    df_a = pd.read_excel(EXCEL, sheet_name='أفراد',  header=1)
    df_s = pd.read_excel(EXCEL, sheet_name='شركات', header=1)

    exist = set()
    if 'رقم_التواصل' in df_a.columns:
        exist.update(str(x)[-9:] for x in df_a['رقم_التواصل'].dropna())
    if 'رقم الجوال' in df_s.columns:
        exist.update(str(x)[-9:] for x in df_s['رقم الجوال'].dropna())

    new_a, new_s, dups = [], [], 0
    src = source[:25]

    for r in records:
        ph = r['phone']
        if ph[-9:] in exist:
            dups += 1
            continue
        exist.add(ph[-9:])
        if is_co(r['text']):
            new_s.append((get_name(r['text']), ph, src))
        else:
            new_a.append((get_name(r['text']), ph, get_region(r['text']), r['date'], r['text'][:200], src))

    wb = load_workbook(EXCEL)
    if new_a:
        ws = wb['أفراد']
        start = ws.max_row + 1
        num_start = len(df_a) + 1
        for i, (nm, ph, rg, dt, tx, s) in enumerate(new_a):
            r = start + i
            for col, val in enumerate([num_start+i, ph, nm, 'دواجن', rg, '', dt, tx, s], 1):
                cell = ws.cell(r, col, val)
                cell.fill, cell.font = FILL_ROW, FONT_MAIN
                cell.alignment = ALIGN_C if col in [1, 2, 6, 7] else ALIGN_R

    if new_s:
        ws = wb['شركات']
        start = ws.max_row + 1
        num_start = len(df_s) + 1
        for i, (nm, ph, s) in enumerate(new_s):
            r = start + i
            for col, val in enumerate([num_start+i, nm, 'غير محدد', ph], 1):
                cell = ws.cell(r, col, val)
                cell.fill, cell.font = FILL_ROW, FONT_MAIN
                cell.alignment = ALIGN_C if col in [1, 4] else ALIGN_R

    wb['الملخص']['C5'] = len(df_a) + len(new_a)
    wb['الملخص']['C6'] = len(df_s) + len(new_s)
    wb['الملخص']['C7'] = wb['الملخص']['C5'].value + wb['الملخص']['C6'].value
    wb.save(EXCEL)
    return len(new_a), len(new_s), dups, (wb['الملخص']['C7'].value)

# ─── معالجة النصوص ────────────────────────────────────────────────────────────
def process_content(text_content, source_name):
    recs = []
    for line in text_content.splitlines():
        p = get_phone(line)
        if p:
            date_m = re.search(r'(\d{1,2}[/.]\d{1,2}[/.]\d{2,4})', line)
            date = date_m.group(1) if date_m else datetime.now().strftime('%d.%m.%Y')
            recs.append({'phone': p, 'text': line, 'date': date})
    
    if not recs: return None, "⚠️ لم أجد أرقام سعودية (05xxxxxxx)"
    na, ns, dups, total = update_excel(recs, source_name)
    return (na, ns), (f"✅ *تم الاستخراج بنجاح!*\n\n"
                     f"📝 النوع: `نص مباشر / دردشة`\n"
                     f"👤 أفراد: *{na}* | 🏢 شركات: *{ns}*\n"
                     f"🔁 مكرر: *{dups}* | 📊 الإجمالي: *{total:,}*")

# ─── الأوامر ──────────────────────────────────────────────────────────────────
async def handle_text(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not ok(update.effective_user.id): return
    status_msg = await update.message.reply_text("⏳ جاري فحص النص...")
    res, txt = process_content(update.message.text, f"نص_{update.effective_user.first_name}")
    await status_msg.edit_text(txt, parse_mode="Markdown")

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not ok(update.effective_user.id): return
    await update.message.reply_text("🐔 *بوت سلالة المتطور*\n\nأرسل ملفات (ZIP/TXT) أو **انسخ النص هنا مباشرة** لاستخراج الأرقام وحفظها!", parse_mode="Markdown")

# [بقية الدوال cmd_stats, cmd_getfile, cmd_duplicates, cmd_adduser, cmd_users, handle_doc تبقى كما هي في ملفك]

async def cmd_stats(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not ok(update.effective_user.id): return
    df_a = pd.read_excel(EXCEL, sheet_name='أفراد', header=1)
    df_s = pd.read_excel(EXCEL, sheet_name='شركات', header=1)
    await update.message.reply_text(f"📊 *إحصائيات*\n👤 أفراد: {len(df_a)}\n🏢 شركات: {len(df_s)}", parse_mode="Markdown")

async def cmd_getfile(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not ok(update.effective_user.id): return
    await update.message.reply_document(document=open(EXCEL, 'rb'), filename="ارقام_الدواجن.xlsx")

async def handle_doc(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not ok(update.effective_user.id): return
    doc = update.message.document
    f = await ctx.bot.get_file(doc.file_id)
    fp = TEMP_DIR / doc.file_name
    await f.download_to_drive(fp)
    
    # استخدام نفس المنطق للملفات
    if fp.suffix.lower() == '.txt':
        content = fp.read_text(encoding='utf-8', errors='ignore')
        res, txt = process_content(content, fp.stem)
        await update.message.reply_text(txt, parse_mode="Markdown")
    # [منطق الـ ZIP يبقى كما هو في ملفك الأصلي]
    fp.unlink(missing_ok=True)

if __name__ == "__main__":
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("stats", cmd_stats))
    app.add_handler(CommandHandler("getfile", cmd_getfile))
    # إضافة مستجيب للنصوص المباشرة
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_text))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_doc))
    print("✅ البوت المحدث جاهز!")
    app.run_polling(drop_pending_updates=True)
