#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import re, json, logging, os
from pathlib import Path
from datetime import datetime
import pandas as pd
import pdfplumber
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from telegram import Update
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

# ─── إدارة المستخدمين ──────────────────────────────────────────────────
def get_users():
    if USERS_FILE.exists():
        try: return json.loads(USERS_FILE.read_text(encoding="utf-8"))
        except: return [ADMIN_ID]
    return [ADMIN_ID]

def ok(uid):
    return uid in get_users()

# ─── محرك الاستخراج الذكي ─────────────────────────────────────────────

def get_phone(text):
    """دالة مرنة جداً لاستخراج الرقم السعودي من أي نص أو جدول"""
    if not text or text == 'nan': return None
    
    # 1. تنظيف النص من كل شيء ماعدا الأرقام
    clean_text = re.sub(r'\D', '', str(text))
    
    # 2. البحث عن نمط الجوال (يبدأ بـ 5 ويتبعه 8 أرقام) مع احتمالية وجود 966 أو 0 في البداية
    # الصيغ المدعومة: 05xxxxxxxx, 9665xxxxxxxx, 5xxxxxxxx
    m = re.search(r'(?:966|0)?(5\d{8})', clean_text)
    if m:
        return '0' + m.group(1)
    
    return None

def get_name(text, default="غير محدد"):
    if not text or text == 'nan': return default
    text = str(text).strip()
    # تنظيف الأسماء المستخرجة من الواتساب التي تبدأ بـ ~
    m = re.search(r'~\s*([^:]+):', text)
    if m: return m.group(1).strip()[:40]
    return text[:40]

COMPANY_KW = ["شركة","مؤسسة","مصنع","توزيع","تسويق","مندوب","مبيعات","وكيل","موزع","مشرف","للتجارة","المحدودة","شركه","مؤسسه"]
REGIONS    = ["الرياض","جدة","مكة","المدينة","الدمام","الخبر","الخرج","القصيم","بريدة","حائل","تبوك","أبها","نجران","جيزان","الطائف"]

def is_co(text):
    return any(k in str(text) for k in COMPANY_KW)

# ─── معالجة الملفات (Excel, PDF, Text) ────────────────────────────────

def extract_from_excel(path):
    try:
        df = pd.read_csv(path) if path.suffix.lower() == '.csv' else pd.read_excel(path)
        cols = df.columns.tolist()
        # البحث عن أعمدة الجوال والأسماء بمرونة
        p_col = next((c for c in cols if any(k in str(c).lower() for k in ['جوال', 'هاتف', 'phone', 'تواصل'])), None)
        n_col = next((c for c in cols if any(k in str(c).lower() for k in ['اسم', 'name', 'العميل'])), None)
        
        recs = []
        for _, row in df.iterrows():
            ph = get_phone(row[p_col]) if p_col else None
            if ph:
                nm = get_name(row[n_col]) if n_col else "غير محدد"
                recs.append({'phone': ph, 'text': f"{nm} : {ph}", 'date': datetime.now().strftime('%d.%m.%Y')})
        return recs
    except: return []

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
    except: pass
    return recs

# ─── تحديث الإكسيل الرئيسي ─────────────────────────────────────────────

def update_excel(records, source_name):
    wb = load_workbook(EXCEL)
    ws_a, ws_s = wb['أفراد'], wb['شركات']
    
    # فحص المكرر
    exist = set()
    for row in ws_a.iter_rows(min_row=2, max_col=2, values_only=True):
        if row[1]: exist.add(str(row[1])[-9:])
    for row in ws_s.iter_rows(min_row=2, max_col=4, values_only=True):
        if row[3]: exist.add(str(row[3])[-9:])

    na, ns, dups = 0, 0, 0
    for r in records:
        ph = r['phone']
        if ph[-9:] in exist:
            dups += 1
            continue
        
        exist.add(ph[-9:])
        if is_co(r['text']):
            row_idx = ws_s.max_row + 1
            cells = [row_idx - 1, get_name(r['text']), "مستخرج", ph]
            for c, val in enumerate(cells, 1):
                cell = ws_s.cell(row_idx, c, val)
                cell.fill, cell.font, cell.alignment = FILL_ROW, FONT_MAIN, (ALIGN_C if c in [1,4] else ALIGN_R)
            ns += 1
        else:
            row_idx = ws_a.max_row + 1
            reg = next((rg for rg in REGIONS if rg in r['text']), "غير محدد")
            cells = [row_idx - 1, ph, get_name(r['text']), "دواجن", reg, "", r['date'], r['text'][:100], source_name[:20]]
            for c, val in enumerate(cells, 1):
                cell = ws_a.cell(row_idx, c, val)
                cell.fill, cell.font, cell.alignment = FILL_ROW, FONT_MAIN, (ALIGN_C if c in [1,2,6,7] else ALIGN_R)
            na += 1

    wb.save(EXCEL)
    return na, ns, dups

# ─── واجهة البوت ─────────────────────────────────────────────────────

async def handle_docs(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not ok(update.effective_user.id): return
    doc = update.message.document
    status = await update.message.reply_text("⏳ جاري تحليل الملف ذكياً...")
    
    f = await ctx.bot.get_file(doc.file_id)
    fp = TEMP_DIR / doc.file_name
    await f.download_to_drive(fp)
    
    ext = fp.suffix.lower()
    recs = []
    
    if ext in ['.xlsx', '.xls', '.csv']: recs = extract_from_excel(fp)
    elif ext == '.pdf': recs = extract_from_pdf(fp)
    elif ext in ['.txt', '.zip']:
        content = fp.read_text(encoding='utf-8', errors='ignore')
        recs = [{'phone': p, 'text': line, 'date': datetime.now().strftime('%d.%m.%Y')} 
                for line in content.splitlines() if (p := get_phone(line))]

    if recs:
        na, ns, d = update_excel(recs, doc.file_name)
        await status.edit_text(f"✅ *تمت المعالجة بنجاح*\n\n📄 الملف: `{doc.file_name}`\n👤 أفراد: {na}\n🏢 شركات: {ns}\n🔁 مكرر محذوف: {d}", parse_mode="Markdown")
    else:
        await status.edit_text("⚠️ لم أتمكن من العثور على أرقام جوال (تبدأ بـ 05) في هذا الملف.")
    
    if fp.exists(): fp.unlink()

async def handle_text(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not ok(update.effective_user.id): return
    recs = [{'phone': p, 'text': line, 'date': datetime.now().strftime('%d.%m.%Y')} 
            for line in update.message.text.splitlines() if (p := get_phone(line))]
    if recs:
        na, ns, d = update_excel(recs, "نص مباشر")
        await update.message.reply_text(f"✅ تم الاستخراج من النص:\nأفراد: {na} | شركات: {ns} | مكرر: {d}")

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if ok(update.effective_user.id):
        await update.message.reply_text("🐔 *بوت سلالة المطور*\nأرسل ملفات PDF، Excel، أو نصوص مباشرة وسأحفظها لك فوراً.", parse_mode="Markdown")

async def cmd_getfile(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if ok(update.effective_user.id):
        await update.message.reply_document(document=open(EXCEL, 'rb'), filename="قاعدة_البيانات.xlsx")

# ─── التشغيل ──────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("getfile", cmd_getfile))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_docs))
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_text))
    print("🚀 البوت المطور يعمل الآن...")
    app.run_polling()
