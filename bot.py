#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import re, zipfile, json, logging, os
from pathlib import Path
from datetime import datetime
from collections import Counter
import pandas as pd
import pdfplumber
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes

# ─── الإعدادات الأساسية ────────────────────────────────────────────────────────
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

# ─── إدارة المستخدمين ──────────────────────────────────────────────────────────
def get_users():
    if USERS_FILE.exists():
        try: return json.loads(USERS_FILE.read_text(encoding="utf-8"))
        except: return [ADMIN_ID]
    return [ADMIN_ID]

def set_users(u):
    USERS_FILE.write_text(json.dumps(u), encoding="utf-8")

def ok(uid):
    return uid in get_users()

# ─── أدوات الاستخراج ──────────────────────────────────────────────────────────
COMPANY_KW = ["شركة","مؤسسة","مصنع","توزيع","تسويق","مندوب","مبيعات","وكيل","موزع","مشرف","للتجارة","المحدودة","شركه","مؤسسه"]
REGIONS    = ["الرياض","جدة","مكة","المدينة","الدمام","الخبر","الخرج","القصيم","بريدة","حائل","تبوك","أبها","نجران","جيزان","الطائف"]

def get_phone(text):
    if not text or text == 'nan': return None
    # تنظيف النص من المسافات والرموز
    clean_text = re.sub(r'[\s\-\(\)\+]', '', str(text))
    # البحث عن صيغة 05xxxxxxxx
    m = re.search(r'05\d{8}', clean_text)
    if m: return m.group()
    # البحث عن صيغة دولية 9665xxxxxxxx
    m = re.search(r'9665\d{8}', clean_text)
    if m: return '0' + m.group()[3:]
    return None

def get_name(text, default="غير محدد"):
    if not text or text == 'nan': return default
    m = re.search(r'~\s*([^:]+):', str(text))
    if m: return m.group(1).strip()[:40]
    return str(text)[:40].strip()

def is_co(text):
    return any(k in str(text) for k in COMPANY_KW)

# ─── معالجة ملفات الإكسيل والـ CSV ─────────────────────────────────────────────
def extract_from_excel(path):
    try:
        # محاولة قراءة الملف (إكسيل أو CSV)
        if path.suffix.lower() == '.csv':
            df = pd.read_csv(path)
        else:
            df = pd.read_excel(path)
        
        # البحث عن الأعمدة المناسبة تلقائياً
        cols = df.columns.tolist()
        phone_col = next((c for c in cols if any(k in str(c).lower() for k in ['جوال', 'هاتف', 'phone', 'تواصل'])), None)
        name_col  = next((c for c in cols if any(k in str(c).lower() for k in ['اسم', 'name', 'العميل'])), None)
        
        recs = []
        for _, row in df.iterrows():
            ph = get_phone(row[phone_col]) if phone_col else None
            if ph:
                nm = get_name(row[name_col]) if name_col else "غير محدد"
                recs.append({'phone': ph, 'text': f"{nm} : {ph}", 'date': datetime.now().strftime('%d.%m.%Y')})
        return recs
    except Exception as e:
        logging.error(f"Excel Error: {e}")
        return []

# ─── معالجة ملفات PDF ──────────────────────────────────────────────────────────
def extract_from_pdf(path):
    recs = []
    try:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    for line in text.splitlines():
                        ph = get_phone(line)
                        if ph:
                            recs.append({'phone': ph, 'text': line, 'date': datetime.now().strftime('%d.%m.%Y')})
    except Exception as e:
        logging.error(f"PDF Error: {e}")
    return recs

# ─── تحديث ملف الإكسيل الرئيسي ─────────────────────────────────────────────────
def update_excel(records, source_name):
    wb = load_workbook(EXCEL)
    ws_a = wb['أفراد']
    ws_s = wb['شركات']
    
    # جلب الأرقام الموجودة مسبقاً لمنع التكرار
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
            data = [row_idx - 1, get_name(r['text']), "مستخرج", ph]
            for c, val in enumerate(data, 1):
                cell = ws_s.cell(row_idx, c, val)
                cell.fill, cell.font, cell.alignment = FILL_ROW, FONT_MAIN, (ALIGN_C if c in [1,4] else ALIGN_R)
            new_s += 1
        else:
            row_idx = ws_a.max_row + 1
            region = next((rg for rg in REGIONS if rg in r['text']), "غير محدد")
            data = [row_idx - 1, ph, get_name(r['text']), "دواجن", region, "", r['date'], r['text'][:100], source_name[:20]]
            for c, val in enumerate(data, 1):
                cell = ws_a.cell(row_idx, c, val)
                cell.fill, cell.font, cell.alignment = FILL_ROW, FONT_MAIN, (ALIGN_C if c in [1,2,6,7] else ALIGN_R)
            new_a += 1

    wb.save(EXCEL)
    return new_a, new_s, dups

# ─── معالجة الرسائل والأوامر ────────────────────────────────────────────────────
async def handle_any_doc(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not ok(update.effective_user.id): return
    doc = update.message.document
    msg = await update.message.reply_text("⏳ جاري تحليل الملف...")
    
    f = await ctx.bot.get_file(doc.file_id)
    fp = TEMP_DIR / doc.file_name
    await f.download_to_drive(fp)
    
    recs = []
    ext = fp.suffix.lower()
    
    if ext in ['.xlsx', '.xls', '.csv']:
        recs = extract_from_excel(fp)
    elif ext == '.pdf':
        recs = extract_from_pdf(fp)
    elif ext in ['.txt', '.zip']:
        # منطق الواتساب القديم (مبسط هنا)
        content = fp.read_text(encoding='utf-8', errors='ignore')
        for line in content.splitlines():
            p = get_phone(line)
            if p: recs.append({'phone': p, 'text': line, 'date': datetime.now().strftime('%d.%m.%Y')})

    if recs:
        na, ns, d = update_excel(recs, doc.file_name)
        await msg.edit_text(f"✅ *تمت المعالجة!*\n\n📄 الملف: `{doc.file_name}`\n👤 أفراد: {na}\n🏢 شركات: {ns}\n🔁 مكرر: {d}", parse_mode="Markdown")
    else:
        await msg.edit_text("⚠️ لم أجد أرقاماً مدعومة في هذا الملف.")
    fp.unlink(missing_ok=True)

async def handle_direct_text(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not ok(update.effective_user.id): return
    lines = update.message.text.splitlines()
    recs = []
    for line in lines:
        p = get_phone(line)
        if p: recs.append({'phone': p, 'text': line, 'date': datetime.now().strftime('%d.%m.%Y')})
    
    if recs:
        na, ns, d = update_excel(recs, "نص مباشر")
        await update.message.reply_text(f"✅ استخراج نصي:\nأفراد: {na} | شركات: {ns} | مكرر: {d}")

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if ok(update.effective_user.id):
        await update.message.reply_text("🐔 *بوت سلالة المتطور جاهز!*\n\nأرسل لي الآن:\n1. ملفات إكسيل (XLSX/CSV)\n2. ملفات PDF\n3. نصوص مباشرة (نسخ/لصق)\n4. تصدير واتساب (ZIP/TXT)", parse_mode="Markdown")

async def cmd_getfile(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if ok(update.effective_user.id):
        await update.message.reply_document(document=open(EXCEL, 'rb'), filename="قاعدة_بيانات_الدواجن.xlsx")

# ─── التشغيل ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("getfile", cmd_getfile))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_any_doc))
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_direct_text))
    print("✅ البوت يعمل ويدعم جميع الصيغ المذكورة!")
    app.run_polling()
