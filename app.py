import streamlit as st
from groq import Groq
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io, os, smtplib, re, json, base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import date
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

# ReportLab
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                 TableStyle, Image as RLImage, HRFlowable,
                                 PageBreak, KeepTogether)

# ══════════════════════════════════════════════════════════════
#  CONFIG
# ══════════════════════════════════════════════════════════════
GMAIL_USER      = "Wijdan.psyc@gmail.com"
GMAIL_PASS      = "rias eeul lyuu stce"
RECIPIENT_EMAIL = "Wijdan.psyc@gmail.com"
LOGO_PATH       = os.path.join(os.path.dirname(__file__), "logo.png")

DEEP_BLUE     = "#1B3A6B"
MID_BLUE      = "#2E6DB4"
GOLD          = "#C8922A"
LIGHT_BG      = "#F5F7FA"

DEEP_BLUE_RGB = RGBColor(0x1B, 0x3A, 0x6B)
MID_BLUE_RGB  = RGBColor(0x2E, 0x6D, 0xB4)
GOLD_RGB      = RGBColor(0xC8, 0x92, 0x2A)
DARK_RGB      = RGBColor(0x1A, 0x1A, 0x2E)

# PDF palette
PDF_DEEP    = colors.HexColor('#1B3A6B')
PDF_MID     = colors.HexColor('#2E6DB4')
PDF_GOLD    = colors.HexColor('#C8922A')
PDF_LIGHT   = colors.HexColor('#F5F7FA')
PDF_BORDER  = colors.HexColor('#C5D3F5')
PDF_CREAM   = colors.HexColor('#EEF2FF')

# ══════════════════════════════════════════════════════════════
#  SB-5 STRUCTURE
# ══════════════════════════════════════════════════════════════
IQ_SCORES     = ["FSIQ","NVIQ","VIQ"]
FACTOR_SCORES = ["FR","KN","QR","VS","WM"]

IQ_LABELS = {
    "FSIQ": {"en":"Full Scale IQ",          "ar":"درجة الذكاء الكلية"},
    "NVIQ": {"en":"Nonverbal IQ",           "ar":"درجة المجال غير اللفظي"},
    "VIQ":  {"en":"Verbal IQ",              "ar":"درجة المجال اللفظي"},
}
FACTOR_LABELS = {
    "FR": {"en":"Fluid Reasoning",          "ar":"الاستدلال التحليلي"},
    "KN": {"en":"Knowledge",               "ar":"المعلومات"},
    "QR": {"en":"Quantitative Reasoning",  "ar":"الاستدلال الكمي"},
    "VS": {"en":"Visual-Spatial",          "ar":"المعالجة البصرية المكانية"},
    "WM": {"en":"Working Memory",          "ar":"الذاكرة العاملة"},
}

CLASSIFICATIONS = [
    (145,160,"Very Gifted",         "موهوب بشدة",        "#1565C0"),
    (130,144,"Gifted",              "موهوب",             "#1976D2"),
    (120,129,"Superior",            "متفوق",             "#0288D1"),
    (110,119,"High Average",        "متوسط مرتفع",       "#00897B"),
    (90, 109,"Average",             "متوسط",             "#388E3C"),
    (80,  89,"Low Average",         "متوسط منخفض",       "#F9A825"),
    (70,  79,"Borderline",          "الفئة البينية",     "#EF6C00"),
    (55,  69,"Mild Impairment",     "إعاقة بسيطة",      "#E53935"),
    (40,  54,"Moderate Impairment", "إعاقة متوسطة",     "#B71C1C"),
    (0,   39,"Severe Impairment",   "إعاقة شديدة",      "#880E4F"),
]

def classify(ss):
    if ss is None: return "Unknown","غير محدد","#888888"
    for lo,hi,en,ar,col in CLASSIFICATIONS:
        if lo <= ss <= hi: return en,ar,col
    return "Unknown","غير محدد","#888888"

def percentile_from_ss(ss):
    if ss is None: return 0
    from math import erf,sqrt
    z = (ss-100)/15.0
    p = 0.5*(1+erf(z/sqrt(2)))
    return max(1,min(99,round(p*100)))

# ══════════════════════════════════════════════════════════════
#  STEP 1: EXTRACT TEXT FROM UPLOADED WORD DOCUMENT (.doc or .docx)
# ══════════════════════════════════════════════════════════════
def extract_word_text(uploaded_file) -> str:
    """
    Extract all text from an uploaded Word file.
    Handles both modern .docx (zip-based) and legacy .doc (binary OLE) formats.
    Uses multiple strategies with fallbacks.
    """
    uploaded_file.seek(0)
    raw_bytes = uploaded_file.read()

    # ── Strategy 1: try python-docx (works for real .docx) ──
    try:
        from docx import Document as DocxDocument
        doc = DocxDocument(io.BytesIO(raw_bytes))
        parts = []
        for para in doc.paragraphs:
            t = para.text.strip()
            if t:
                parts.append(t)
        for table in doc.tables:
            for row in table.rows:
                cells = [c.text.strip() for c in row.cells if c.text.strip()]
                if cells:
                    parts.append(" | ".join(cells))
        text = "\n".join(parts)
        if len(text.strip()) > 100:
            return text
    except Exception:
        pass

    # ── Strategy 2: extract UTF-16-LE text (common in .doc binary) ──
    try:
        text_utf16 = raw_bytes.decode('utf-16-le', errors='ignore')
        # Keep Arabic Unicode block + printable ASCII
        cleaned = re.sub(r'[^\u0600-\u06FF\u0020-\u007E\u00A0-\u00FF\n\r\t]+', ' ', text_utf16)
        cleaned = re.sub(r'[ \t]{3,}', '  ', cleaned)
        lines = [l.strip() for l in cleaned.split('\n') if len(l.strip()) > 3]
        text = "\n".join(lines)
        if len(text.strip()) > 100:
            return text
    except Exception:
        pass

    # ── Strategy 3: scan for UTF-8 / Latin-1 readable text ──
    try:
        text_latin = raw_bytes.decode('latin-1', errors='ignore')
        # Extract Arabic runs (stored as UTF-16 pairs often appear garbled in latin-1,
        # but numbers and latin text are readable)
        # Collect sequences of printable chars ≥ 4 chars long
        printable_runs = re.findall(r'[\u0020-\u007E\u00A0-\u00FF]{4,}', text_latin)
        text = "\n".join(r.strip() for r in printable_runs if r.strip())
        if len(text.strip()) > 100:
            return text
    except Exception:
        pass

    return ""

# ══════════════════════════════════════════════════════════════
#  STEP 2: PARSE DATA FROM RAW TEXT VIA GROQ
# ══════════════════════════════════════════════════════════════
def extract_data_from_text(raw_text: str) -> dict:
    """Send raw PDF text to Groq and get structured JSON data back."""
    prompt = f"""You are a clinical psychologist assistant. The following is raw text extracted from an Arabic Stanford-Binet 5 (SB5) psychological assessment report.

Extract ALL available information and return ONLY a valid JSON object — no explanation, no markdown, no backticks.

Extract these fields (use null for any missing value):
{{
  "name": "examinee full name",
  "dob": "date of birth EXACTLY as written in the document (day/month/year)",
  "age": "age EXACTLY as written — include years, months, days if shown e.g. '8 years 3 months 12 days'",
  "age_years": integer_or_null,
  "age_months": integer_or_null,
  "age_days": integer_or_null,
  "gender": "male or female (translate from Arabic if needed)",
  "grade": "school grade or level",
  "school": "school or institution name",
  "examiner": "examiner name — look for الفاحص or الأخصائي or المختبر",
  "test_date": "date of testing EXACTLY as written in the document",
  "referral": "referral source",
  "complaints": "reason for referral / presenting complaints",
  "behavioral_obs": "behavioral observations during testing",
  "background": "background information / history",
  "FSIQ": integer_or_null,
  "FSIQ_ci": "confidence interval as string e.g. 91-99 or null",
  "NVIQ": integer_or_null,
  "NVIQ_ci": "string or null",
  "VIQ": integer_or_null,
  "VIQ_ci": "string or null",
  "FR": integer_or_null,
  "FR_ci": "string or null",
  "KN": integer_or_null,
  "KN_ci": "string or null",
  "QR": integer_or_null,
  "QR_ci": "string or null",
  "VS": integer_or_null,
  "VS_ci": "string or null",
  "WM": integer_or_null,
  "WM_ci": "string or null",
  "nv_fr": integer_or_null,
  "v_fr": integer_or_null,
  "nv_kn": integer_or_null,
  "v_kn": integer_or_null,
  "nv_qr": integer_or_null,
  "v_qr": integer_or_null,
  "nv_vs": integer_or_null,
  "v_vs": integer_or_null,
  "nv_wm": integer_or_null,
  "v_wm": integer_or_null
}}

Notes:
- CRITICAL: Copy dob and test_date EXACTLY as they appear in the document — do not reformat
- CRITICAL: Copy age EXACTLY as stated — e.g. "8 سنة 3 شهر 12 يوم" → "8 years 3 months 12 days"
- CRITICAL: الفاحص = examiner. Also check for: الأخصائي، المقيّم، اسم الفاحص
- درجة الذكاء الكلية = FSIQ
- المجال غير اللفظي / غير لفظي = NVIQ
- المجال اللفظي / لفظي = VIQ
- الاستدلال التحليلي / الاستدلال الطلاقة = FR
- المعلومات / المعرفة = KN
- الاستدلال الكمي = QR
- المعالجة البصرية المكانية / البصري المكاني = VS
- الذاكرة العاملة = WM
- غير لفظي = NV (nonverbal subtest)
- لفظي = V (verbal subtest)
- Any number in parentheses after a score is likely the confidence interval
- Percentile ranks (رتبة مئينية) should NOT be confused with standard scores
- Standard scores for IQ are typically 40-160; subtest scaled scores are typically 1-19

RAW REPORT TEXT:
{raw_text[:8000]}
"""
    client = Groq(api_key=st.secrets["GROQ_API_KEY"])
    r = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role":"user","content":prompt}],
        max_tokens=1500,
        temperature=0.1,
    )
    raw_json = r.choices[0].message.content.strip()
    # Strip any accidental markdown fences
    raw_json = re.sub(r"^```[a-z]*\n?","",raw_json)
    raw_json = re.sub(r"\n?```$","",raw_json)
    try:
        return json.loads(raw_json)
    except Exception:
        # Try to extract JSON substring
        m = re.search(r'\{.*\}', raw_json, re.DOTALL)
        if m:
            return json.loads(m.group())
        return {}

# ══════════════════════════════════════════════════════════════
#  STEP 2B: TRANSLATE / TRANSLITERATE ARABIC DEMOGRAPHICS
# ══════════════════════════════════════════════════════════════
def translate_demographics(data: dict) -> dict:
    """
    Translate any Arabic text in demographic fields to English.
    Names are transliterated phonetically. Other fields are translated.
    Returns updated data dict with English-safe values.
    """
    fields = {
        "name":      data.get("name","") or "",
        "dob":       data.get("dob","") or "",
        "age":       data.get("age","") or "",
        "gender":    data.get("gender","") or "",
        "grade":     data.get("grade","") or "",
        "school":    data.get("school","") or "",
        "examiner":  data.get("examiner","") or "",
        "test_date": data.get("test_date","") or "",
        "referral":  data.get("referral","") or "",
    }

    # Check if any field contains Arabic
    has_arabic = any(
        any('\u0600' <= c <= '\u06FF' for c in str(v))
        for v in fields.values()
    )
    if not has_arabic:
        return data  # Nothing to translate

    prompt = f"""You are a translation assistant. The following fields may contain Arabic text from a psychological assessment report.

For each field:
- "name" and "examiner": transliterate phonetically to English (e.g. أحمد → Ahmed, محمد → Mohammed)
- "gender": translate (ذكر → Male, أنثى → Female)
- "dob", "test_date": keep numbers exactly, translate month names if written in Arabic
- "age": translate fully (e.g. ٨ سنوات ٣ أشهر → 8 years 3 months)
- All other fields: translate to English

Return ONLY a valid JSON object with the same keys, values in English. No explanation, no markdown.

Input:
{json.dumps(fields, ensure_ascii=False)}
"""
    try:
        client = Groq(api_key=st.secrets["GROQ_API_KEY"])
        r = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=400,
            temperature=0.1,
        )
        raw = r.choices[0].message.content.strip()
        raw = re.sub(r"^```[a-z]*\n?", "", raw)
        raw = re.sub(r"\n?```$", "", raw)
        translated = json.loads(raw)
        for k, v in translated.items():
            if v and k in data:
                data[k] = v
    except Exception:
        pass  # If translation fails, use original values
    return data


def make_profile_chart(data: dict) -> bytes:
    all_keys   = ["FSIQ","NVIQ","VIQ","FR","KN","QR","VS","WM"]
    all_labels, all_vals, all_colors = [], [], []
    for k in all_keys:
        v = data.get(k)
        if v is None: continue
        lbl = (IQ_LABELS if k in IQ_LABELS else FACTOR_LABELS)[k]["en"]
        _,_,col = classify(v)
        all_labels.append(lbl); all_vals.append(v); all_colors.append(col)
    if not all_vals: return None

    fig, ax = plt.subplots(figsize=(12,7))
    fig.patch.set_facecolor('#FAFBFF'); ax.set_facecolor('#FAFBFF')
    y_pos = np.arange(len(all_labels))

    bands = [(40,54,"#FCE4EC"),(55,69,"#FFEBEE"),(70,79,"#FFF3E0"),
             (80,89,"#FFFDE7"),(90,109,"#F1F8E9"),(110,119,"#E3F2FD"),
             (120,129,"#E1F5FE"),(130,160,"#E8EAF6")]
    for lo,hi,col in bands:
        ax.axvspan(lo,hi,alpha=0.3,color=col,zorder=0)

    bars = ax.barh(y_pos, all_vals, color=all_colors, height=0.6,
                   edgecolor='white', linewidth=1.2, zorder=3)
    n_iq = sum(1 for k in all_keys if k in IQ_LABELS and data.get(k))
    if 0 < n_iq < len(all_labels):
        ax.axhline(y=n_iq-0.5, color='#BDBDBD', linestyle='--', linewidth=1.2, zorder=4)

    ax.axvline(x=100, color='#1B3A6B', linestyle='-', linewidth=2, alpha=0.6, label='Mean (100)')
    ax.axvline(x=85,  color='#EF6C00', linestyle=':', linewidth=1.2, alpha=0.5)
    ax.axvline(x=115, color='#EF6C00', linestyle=':', linewidth=1.2, alpha=0.5)

    for bar_, val in zip(bars, all_vals):
        en_c,_,_ = classify(val)
        pct = percentile_from_ss(val)
        label = f"{val}  ({en_c}, {pct}th %ile)"
        ax.text(bar_.get_width()+1, bar_.get_y()+bar_.get_height()/2,
                label, va='center', ha='left', fontsize=8.5, color='#1A1A2E')

    ax.set_yticks(y_pos); ax.set_yticklabels(all_labels, fontsize=10.5)
    ax.set_xlim(40,175)
    ax.set_xlabel("Standard Score  (Mean=100, SD=15)", fontsize=11, color='#1A1A2E')
    ax.set_title("Stanford-Binet 5 — Full Score Profile", fontsize=13,
                 fontweight='bold', color='#1B3A6B', pad=12)
    ax.legend(fontsize=9, framealpha=0.7)
    ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
    ax.grid(axis='x', linestyle=':', alpha=0.4, zorder=1)
    plt.tight_layout()
    buf=io.BytesIO(); plt.savefig(buf,format='png',dpi=150,bbox_inches='tight')
    plt.close(fig); buf.seek(0); return buf.read()

def make_factor_radar(data: dict) -> bytes:
    labels = [FACTOR_LABELS[k]["en"] for k in FACTOR_SCORES if data.get(k)]
    vals   = [data[k] for k in FACTOR_SCORES if data.get(k)]
    if len(vals) < 3: return None
    N = len(labels)
    angles = np.linspace(0,2*np.pi,N,endpoint=False).tolist()
    vals_  = vals+[vals[0]]; angles_= angles+[angles[0]]
    fig,ax = plt.subplots(figsize=(6,6),subplot_kw=dict(polar=True))
    fig.patch.set_facecolor('#FAFBFF'); ax.set_facecolor('#EEF2FF')
    for ref,col in [(70,'#EF6C00'),(100,'#1B3A6B'),(130,'#0288D1')]:
        ref_n = ref/160.0
        ax.plot(angles_,[ref_n]*len(angles_),'--',color=col,linewidth=0.9,alpha=0.5)
    vals_n = [v/160.0 for v in vals_]
    ax.plot(angles_,vals_n,'o-',linewidth=2.2,color='#2E6DB4',markersize=7)
    ax.fill(angles_,vals_n,alpha=0.28,color='#2E6DB4')
    ax.set_xticks(angles[:-1] if len(angles)==N+1 else angles)
    ax.set_xticklabels(labels,size=9.5)
    ax.set_ylim(0,1)
    ax.set_yticks([70/160,85/160,100/160,115/160,130/160])
    ax.set_yticklabels(['70','85','100','115','130'],size=7.5,color='#666')
    ax.set_title("Factor Index Radar Profile",size=12,fontweight='bold',color='#1B3A6B',pad=18)
    plt.tight_layout()
    buf=io.BytesIO(); plt.savefig(buf,format='png',dpi=150,bbox_inches='tight')
    plt.close(fig); buf.seek(0); return buf.read()

def make_classification_gauge(fsiq: int) -> bytes:
    fig,ax = plt.subplots(figsize=(10,4))
    fig.patch.set_facecolor('#FAFBFF'); ax.set_facecolor('#FAFBFF')
    bands_ordered = [
        (40, 54, "#B71C1C","Moderate\nImpairment"),
        (55, 69, "#E53935","Mild\nImpairment"),
        (70, 79, "#EF6C00","Borderline"),
        (80, 89, "#F9A825","Low\nAverage"),
        (90,109, "#388E3C","Average"),
        (110,119,"#00897B","High\nAverage"),
        (120,129,"#0288D1","Superior"),
        (130,145,"#1565C0","Gifted+"),
    ]
    for lo,hi,col,lbl in bands_ordered:
        ax.barh(0,hi-lo,left=lo,height=0.55,color=col,alpha=0.85,edgecolor='white',linewidth=1.5)
        ax.text((lo+hi)/2,0,lbl,ha='center',va='center',fontsize=7.5,
                color='white',fontweight='bold')
    ax.annotate('',xy=(fsiq,0.35),xytext=(fsiq,0.72),
                arrowprops=dict(arrowstyle='->',color='#1A1A2E',lw=2.8))
    en_c,_,_ = classify(fsiq)
    pct = percentile_from_ss(fsiq)
    ax.text(fsiq,0.88,f"FSIQ = {fsiq}\n{en_c} | {pct}th percentile",
            ha='center',va='bottom',fontsize=10.5,fontweight='bold',color='#1A1A2E')
    ax.set_xlim(40,145); ax.set_ylim(-0.45,1.15)
    ax.set_xlabel("Standard Score",fontsize=10)
    ax.set_yticks([])
    ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.set_title("Full Scale IQ — Classification Gauge",fontsize=12,
                 fontweight='bold',color='#1B3A6B',pad=8)
    plt.tight_layout()
    buf=io.BytesIO(); plt.savefig(buf,format='png',dpi=150,bbox_inches='tight')
    plt.close(fig); buf.seek(0); return buf.read()

def make_subtest_chart(data: dict) -> bytes:
    keys = FACTOR_SCORES
    nv_vals = [data.get(f"nv_{k.lower()}") for k in keys]
    v_vals  = [data.get(f"v_{k.lower()}")  for k in keys]
    if all(v is None for v in nv_vals+v_vals): return None
    labels = [FACTOR_LABELS[k]["en"] for k in keys]
    x = np.arange(len(labels)); w = 0.35
    fig,ax = plt.subplots(figsize=(11,5.5))
    fig.patch.set_facecolor('#FAFBFF'); ax.set_facecolor('#FAFBFF')
    nv_c = [v if v is not None else 0 for v in nv_vals]
    v_c  = [v if v is not None else 0 for v in v_vals]
    b1=ax.bar(x-w/2,nv_c,w,label='Nonverbal',color='#2E6DB4',edgecolor='white',linewidth=0.8)
    b2=ax.bar(x+w/2,v_c, w,label='Verbal',   color='#C8922A',edgecolor='white',linewidth=0.8)
    ax.axhline(y=10,color='#1B3A6B',linestyle='--',linewidth=1.8,alpha=0.6,label='Mean (10)')
    ax.axhline(y=7, color='#EF6C00',linestyle=':',linewidth=1.2,alpha=0.5)
    ax.axhline(y=13,color='#EF6C00',linestyle=':',linewidth=1.2,alpha=0.5)
    for bar_ in list(b1)+list(b2):
        h=bar_.get_height()
        if h>0: ax.text(bar_.get_x()+bar_.get_width()/2.,h+0.15,str(int(h)),
                        ha='center',va='bottom',fontsize=9.5,fontweight='bold')
    ax.set_xticks(x); ax.set_xticklabels(labels,rotation=15,ha='right',fontsize=10)
    ax.set_ylabel("Scaled Score  (Mean=10, SD=3)",fontsize=10)
    ax.set_title("Subtest Scaled Scores — Nonverbal vs. Verbal",fontsize=12,
                 fontweight='bold',color='#1B3A6B')
    ax.set_ylim(0,21); ax.legend(fontsize=10)
    ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
    ax.grid(axis='y',linestyle=':',alpha=0.4)
    plt.tight_layout()
    buf=io.BytesIO(); plt.savefig(buf,format='png',dpi=150,bbox_inches='tight')
    plt.close(fig); buf.seek(0); return buf.read()

def make_discrepancy_chart(data: dict) -> bytes:
    """Bar chart comparing NV vs V IQ and showing discrepancy significance."""
    nviq = data.get("NVIQ"); viq = data.get("VIQ")
    if not nviq or not viq: return None
    disc = abs(nviq - viq)
    fig, axes = plt.subplots(1, 2, figsize=(11, 4.5))
    fig.patch.set_facecolor('#FAFBFF')
    # Left: NV vs V comparison
    ax = axes[0]; ax.set_facecolor('#FAFBFF')
    cats = ["Nonverbal IQ\n(NVIQ)", "Verbal IQ\n(VIQ)"]
    vals = [nviq, viq]
    cols = ["#2E6DB4","#C8922A"]
    bars = ax.bar(cats, vals, color=cols, width=0.45, edgecolor='white', linewidth=1.5)
    ax.axhline(y=100, color='#1B3A6B', linestyle='--', linewidth=1.5, alpha=0.6)
    for bar_,val in zip(bars,vals):
        en_c,_,_ = classify(val)
        ax.text(bar_.get_x()+bar_.get_width()/2., val+1.5,
                f"{val}\n({en_c})", ha='center', va='bottom', fontsize=10, fontweight='bold')
    ax.set_ylim(50, max(vals)+25)
    ax.set_title("NV vs. V IQ Comparison", fontsize=11, fontweight='bold', color='#1B3A6B')
    ax.set_ylabel("Standard Score", fontsize=10)
    ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
    # Right: discrepancy significance
    ax2 = axes[1]; ax2.set_facecolor('#FAFBFF')
    sig_color = "#E53935" if disc >= 23 else "#F9A825" if disc >= 15 else "#388E3C"
    sig_label = "Statistically Significant\n(p < .05)" if disc >= 15 else "Not Significant"
    ax2.barh(0, disc, height=0.4, color=sig_color, edgecolor='white')
    ax2.axvline(x=15, color='#F9A825', linestyle='--', linewidth=1.5, label='Significant (15)')
    ax2.axvline(x=23, color='#E53935', linestyle='--', linewidth=1.5, label='Highly Sig. (23)')
    ax2.text(disc+0.5, 0, f"Discrepancy = {disc}\n{sig_label}",
             va='center', ha='left', fontsize=10, fontweight='bold', color='#1A1A2E')
    ax2.set_xlim(0, max(30, disc+15))
    ax2.set_yticks([])
    ax2.set_xlabel("Point Difference (|NVIQ − VIQ|)", fontsize=10)
    ax2.set_title("NV–V Discrepancy Analysis", fontsize=11, fontweight='bold', color='#1B3A6B')
    ax2.legend(fontsize=9, framealpha=0.7)
    ax2.spines['top'].set_visible(False); ax2.spines['right'].set_visible(False)
    plt.tight_layout()
    buf=io.BytesIO(); plt.savefig(buf,format='png',dpi=150,bbox_inches='tight')
    plt.close(fig); buf.seek(0); return buf.read()

# ══════════════════════════════════════════════════════════════
#  STEP 3: GROQ REPORT GENERATION
# ══════════════════════════════════════════════════════════════
def build_score_summary(data: dict) -> str:
    lines = ["IQ SCORES:"]
    for k in IQ_SCORES:
        v = data.get(k)
        if v:
            en_c,_,_ = classify(v); pct=percentile_from_ss(v); ci=data.get(f"{k}_ci","—")
            lines.append(f"  {IQ_LABELS[k]['en']}: SS={v}, %ile={pct}th, 90%CI={ci}, Classification={en_c}")
    lines.append("\nFACTOR INDEX SCORES:")
    for k in FACTOR_SCORES:
        v = data.get(k)
        if v:
            en_c,_,_ = classify(v); pct=percentile_from_ss(v); ci=data.get(f"{k}_ci","—")
            lines.append(f"  {FACTOR_LABELS[k]['en']}: SS={v}, %ile={pct}th, 90%CI={ci}, Classification={en_c}")
    lines.append("\nSUBTEST SCALED SCORES (NV | V):")
    for k in FACTOR_SCORES:
        nv=data.get(f"nv_{k.lower()}"); v=data.get(f"v_{k.lower()}")
        if nv or v:
            lines.append(f"  {FACTOR_LABELS[k]['en']}: NV={nv or '—'} | V={v or '—'}")
    return "\n".join(lines)

def generate_en_report(data: dict) -> str:
    score_summary = build_score_summary(data)
    name     = data.get("name","the examinee") or "the examinee"
    age      = data.get("age","—") or "—"
    gender   = data.get("gender","—") or "—"
    pronoun  = "he" if "male" in str(gender).lower() else "she"
    examiner = data.get("examiner","—") or "—"
    dob      = data.get("dob","—") or "—"
    tdate    = data.get("test_date","—") or "—"

    prompt = f"""You are a senior licensed psychologist writing a world-class Stanford-Binet 5 (SB5) psychological assessment report.
Use the most current research on SB5 (Roid, 2003; Roid et al., 2016) and CHC theory of intelligence.

EXAMINEE: {name} | AGE: {age} | GENDER: {gender} | EXAMINER: {examiner}
DATE OF BIRTH: {dob}
TEST DATE: {tdate}
REPORT DATE: {date.today().strftime('%B %d, %Y')}

SCORE SUMMARY:
{score_summary}

Write a COMPREHENSIVE, PREMIUM SB5 report. Use formal psychoeducational language. Be very specific to the actual scores.
No markdown symbols (**, ##, ---, etc.). Section titles: ALL CAPS on their own line.
Use {pronoun}/{pronoun}s consistently. Reference specific scores and percentiles throughout.

IMPORTANT — The report must start with this exact header block, filling in the values as given:

STANFORD-BINET INTELLIGENCE SCALES, FIFTH EDITION — PSYCHOLOGICAL REPORT
Name | {name}
Date of Birth | {dob}
Age | {age}
Gender | {gender}
Examiner | {examiner}
Test Date | {tdate}
Report Date | {date.today().strftime('%B %d, %Y')}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Then write these sections IN ORDER — do NOT include Reason for Referral, Background Information, Tests Administered, or Behavioral Observations:

ASSESSMENT RESULTS AND INTERPRETATION

1. FULL SCALE IQ (FSIQ)
Interpret FSIQ deeply: what it represents, confidence interval, percentile, classification, clinical meaning.
Reference CHC theory (Gf-Gc framework).

2. NONVERBAL IQ (NVIQ) AND VERBAL IQ (VIQ)
Compare NV and V domains. Discuss discrepancy (statistically significant at p<.05 if |NV-V|≥15).
Clinical implications of any discrepancy.

3. FLUID REASONING (FR)
Interpret factor score, NV and V subtest scores, inductive/deductive reasoning abilities.

4. KNOWLEDGE (KN)
Interpret crystallized intelligence, vocabulary, acquired knowledge.

5. QUANTITATIVE REASONING (QR)
Interpret numerical and mathematical reasoning abilities.

6. VISUAL-SPATIAL PROCESSING (VS)
Interpret spatial visualization, pattern recognition, visuoconstructive abilities.

7. WORKING MEMORY (WM)
Interpret short-term memory, attention, cognitive control.

8. STRENGTHS AND WEAKNESSES PROFILE
Identify relative strengths (highest scores) and weaknesses (lowest scores).
Discuss intra-cognitive variability. Is the profile consistent or scattered?

9. DIAGNOSTIC IMPRESSIONS
Patterns that emerge. Reference relevant diagnostic considerations as hypotheses only.
Do NOT provide a formal diagnosis.

10. RECOMMENDATIONS
Provide 10–12 specific, evidence-based recommendations across:
a) Educational accommodations and classroom strategies
b) Intervention priorities (cognitive, academic, behavioral)
c) Further evaluation needs
d) Family and home support strategies
e) Therapeutic or clinical referrals if indicated

11. SUMMARY
A concise 2-paragraph executive summary for school teams and specialists.

NOTE TO FAMILY
Write a warm, plain-language 2-paragraph explanation for the family with NO clinical jargon.
Explain what the test found, what it means for their child in daily life, and the 3 most important next steps.
Label this section clearly as "NOTE TO FAMILY" — do not call it "Parent-Friendly Summary" or anything else.
"""
    client = Groq(api_key=st.secrets["GROQ_API_KEY"])
    r = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role":"user","content":prompt}],
        max_tokens=4500,
    )
    return r.choices[0].message.content.strip()

# ══════════════════════════════════════════════════════════════
#  STEP 4: BUILD ENGLISH PDF REPORT
# ══════════════════════════════════════════════════════════════
def _pdf_styles():
    S = {}
    S['title']    = ParagraphStyle('title',    fontName='Helvetica-Bold',  fontSize=17, textColor=PDF_DEEP,  spaceAfter=10,  alignment=TA_CENTER)
    S['subtitle'] = ParagraphStyle('subtitle', fontName='Helvetica',       fontSize=9,  textColor=PDF_GOLD,  spaceAfter=2,  alignment=TA_CENTER)
    S['section']  = ParagraphStyle('section',  fontName='Helvetica-Bold',  fontSize=11, textColor=PDF_DEEP,  spaceBefore=14,spaceAfter=4)
    S['body']     = ParagraphStyle('body',     fontName='Helvetica',       fontSize=9.5,textColor=colors.HexColor('#1A1A2E'),leading=14,spaceAfter=5)
    S['small']    = ParagraphStyle('small',    fontName='Helvetica',       fontSize=8,  textColor=PDF_MID,   leading=11)
    S['bold']     = ParagraphStyle('bold',     fontName='Helvetica-Bold',  fontSize=9.5,textColor=colors.HexColor('#1A1A2E'),leading=14,spaceAfter=3)
    S['parent']   = ParagraphStyle('parent',   fontName='Helvetica',       fontSize=10.5,textColor=colors.HexColor('#1A1A2E'),leading=16,spaceAfter=6,leftIndent=10,rightIndent=10,backColor=colors.HexColor('#FFFDF0'))
    return S

def _hdr_style(): return colors.HexColor('#1B3A6B')
def _shade1():    return colors.HexColor('#EEF2FF')
def _shade2():    return colors.HexColor('#FFF8EE')

def _score_tbl(rows, col_widths, header_cmds=None):
    t = Table(rows, colWidths=col_widths, repeatRows=1)
    cmds = [
        ('BACKGROUND',(0,0),(-1,0), _hdr_style()),
        ('TEXTCOLOR', (0,0),(-1,0), colors.white),
        ('FONTNAME',  (0,0),(-1,0),'Helvetica-Bold'),
        ('FONTSIZE',  (0,0),(-1,-1),8.5),
        ('BOX',       (0,0),(-1,-1),0.5,PDF_BORDER),
        ('INNERGRID', (0,0),(-1,-1),0.3,PDF_BORDER),
        ('VALIGN',    (0,0),(-1,-1),'MIDDLE'),
        ('TOPPADDING',(0,0),(-1,-1),4),
        ('BOTTOMPADDING',(0,0),(-1,-1),4),
        ('LEFTPADDING',(0,0),(-1,-1),6),
    ]
    if header_cmds: cmds.extend(header_cmds)
    t.setStyle(TableStyle(cmds))
    return t

def build_pdf_report(report_text, data, charts, center_name="", logo_bytes=None) -> io.BytesIO:
    buf = io.BytesIO()
    W_page, H_page = A4
    doc = SimpleDocTemplate(buf, pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm, topMargin=2*cm, bottomMargin=2*cm)
    W = W_page - 4*cm
    S = _pdf_styles()
    story = []

    # ── Logo + Center Name + Title ──
    # Uploaded logo takes priority, fallback to file logo
    logo_shown = False
    if logo_bytes:
        try:
            logo = RLImage(io.BytesIO(logo_bytes), width=4.5*cm, height=2*cm)
            logo.hAlign = 'CENTER'
            story.append(logo)
            story.append(Spacer(1, 4))
            logo_shown = True
        except: pass
    if not logo_shown and os.path.exists(LOGO_PATH):
        try:
            logo = RLImage(LOGO_PATH, width=4.5*cm, height=2*cm)
            logo.hAlign = 'CENTER'
            story.append(logo)
            story.append(Spacer(1, 4))
        except: pass

    if center_name:
        story.append(Paragraph(center_name, ParagraphStyle('center_name',
            fontName='Helvetica-Bold', fontSize=13, textColor=PDF_DEEP,
            spaceAfter=4, alignment=TA_CENTER)))

    story.append(Paragraph("Stanford-Binet Intelligence Scales, Fifth Edition", S['title']))
    story.append(Paragraph("Psychological Assessment Report  ·  SB5 · Roid (2003)", S['subtitle']))
    story.append(HRFlowable(width=W, thickness=2, color=PDF_GOLD, spaceAfter=10))

    # ── Demographics table ──
    name    = data.get("name","—") or "—"
    age     = data.get("age","—") or "—"
    gender  = data.get("gender","—") or "—"
    examiner= data.get("examiner","—") or "—"
    dob     = data.get("dob","—") or "—"
    tdate   = data.get("test_date","—") or "—"
    ref     = data.get("referral","—") or "—"

    demo_rows = [
        [Paragraph('<b>Name</b>',S['small']),    Paragraph(name,S['body']),
         Paragraph('<b>Date of Birth</b>',S['small']), Paragraph(dob,S['body'])],
        [Paragraph('<b>Age</b>',S['small']),     Paragraph(age,S['body']),
         Paragraph('<b>Gender</b>',S['small']),  Paragraph(gender,S['body'])],
        [Paragraph('<b>Examiner</b>',S['small']),Paragraph(examiner,S['body']),
         Paragraph('<b>Test Date</b>',S['small']),Paragraph(tdate,S['body'])],
        [Paragraph('<b>Referral</b>',S['small']),Paragraph(ref,S['body']),
         Paragraph('<b>Report Date</b>',S['small']),Paragraph(date.today().strftime('%B %d, %Y'),S['body'])],
    ]
    demo_tbl = Table(demo_rows, colWidths=[2.5*cm,6.2*cm,3*cm,5.8*cm])
    demo_tbl.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,-1),_shade1()),
        ('BOX',(0,0),(-1,-1),0.5,PDF_BORDER),
        ('INNERGRID',(0,0),(-1,-1),0.3,PDF_BORDER),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('TOPPADDING',(0,0),(-1,-1),5),
        ('BOTTOMPADDING',(0,0),(-1,-1),5),
        ('LEFTPADDING',(0,0),(-1,-1),6),
    ]))
    story.append(KeepTogether([demo_tbl, Spacer(1,10)]))

    # ── IQ + Factor Score Summary Table ──
    story.append(KeepTogether([
        Paragraph("IQ AND FACTOR INDEX SCORE SUMMARY", S['section']),
        _build_score_table(data, S, W),
        Spacer(1,8)
    ]))

    # ── Subtest Table ──
    sub_tbl = _build_subtest_table(data, S, W)
    if sub_tbl:
        story.append(KeepTogether([
            Paragraph("SUBTEST SCALED SCORES", S['section']),
            sub_tbl,
            Spacer(1,8)
        ]))

    # ── Classification Colour Legend ──
    story.append(KeepTogether([
        Paragraph("SCORE CLASSIFICATION GUIDE", S['section']),
        _build_legend_table(S, W),
        Spacer(1,8)
    ]))

    # ── Charts ──
    for key, title, h_ratio in [
        ("profile",     "IQ AND FACTOR SCORE PROFILE CHART", 0.52),
        ("gauge",       "FULL SCALE IQ CLASSIFICATION GAUGE", 0.38),
        ("discrepancy", "NV vs. V IQ DISCREPANCY ANALYSIS",  0.42),
        ("radar",       "FACTOR INDEX RADAR PROFILE",         0.55),
        ("subtest",     "SUBTEST COMPARISON CHART",           0.46),
    ]:
        if charts.get(key):
            img = RLImage(io.BytesIO(charts[key]), width=W, height=W*h_ratio)
            img.hAlign = 'CENTER'
            story.append(KeepTogether([
                Paragraph(title, S['section']),
                img,
                Spacer(1,8)
            ]))

    # ── Clinical Narrative ──
    story.append(HRFlowable(width=W, thickness=1, color=PDF_GOLD, spaceAfter=6))
    story.append(Paragraph("CLINICAL NARRATIVE REPORT", S['section']))

    sec_pat = re.compile(r'^\d+\.\s+[A-Z][A-Z\s,&/\(\):\']+$')
    header_words = {
        "ASSESSMENT RESULTS AND INTERPRETATION",
        "STRENGTHS AND WEAKNESSES PROFILE","DIAGNOSTIC IMPRESSIONS",
        "RECOMMENDATIONS","SUMMARY","NOTE TO FAMILY",
        "PARENT-FRIENDLY SUMMARY",  # fallback in case LLM ignores instruction
        "STANFORD-BINET INTELLIGENCE SCALES, FIFTH EDITION — PSYCHOLOGICAL REPORT",
        "CLINICAL NARRATIVE REPORT",
    }
    in_parent = False

    for line in report_text.split('\n'):
        ls = line.strip()
        if not ls:
            story.append(Spacer(1,4)); continue
        if ls.startswith('━') or ls.startswith('═') or ls.startswith('---'):
            story.append(HRFlowable(width=W,thickness=0.4,color=PDF_BORDER,spaceAfter=4)); continue

        upper = ls.upper()
        is_sec = (sec_pat.match(ls) or ls in header_words or
                  any(upper.startswith(h) for h in header_words))

        if is_sec:
            if "NOTE TO FAMILY" in upper or "PARENT" in upper or "FAMILY" in upper:
                in_parent = True
                story.append(HRFlowable(width=W,thickness=2,color=PDF_GOLD,spaceAfter=4))
                story.append(Paragraph("NOTE TO FAMILY", ParagraphStyle('psec',
                    fontName='Helvetica-Bold',fontSize=13,textColor=PDF_GOLD,
                    spaceBefore=12,spaceAfter=8,alignment=TA_CENTER)))
            else:
                story.append(Paragraph(ls, S['section']))
            continue

        if '|' in ls:
            parts = [p.strip() for p in ls.split('|') if p.strip()]
            if len(parts) >= 2:
                # Skip only table-header rows from score tables
                skip_pairs = [("field","value"),("subscale","raw"),("scale","ss")]
                key_pair = (parts[0].lower().strip(), parts[1].lower().strip())
                if key_pair in skip_pairs:
                    continue
                col_w = W / len(parts)
                mini = Table([[Paragraph(p, S['body']) for p in parts]],
                             colWidths=[col_w]*len(parts))
                mini.setStyle(TableStyle([
                    ('BOX',(0,0),(-1,-1),0.3,PDF_BORDER),
                    ('BACKGROUND',(0,0),(-1,-1),_shade1()),
                    ('TOPPADDING',(0,0),(-1,-1),3),
                    ('BOTTOMPADDING',(0,0),(-1,-1),3),
                    ('LEFTPADDING',(0,0),(-1,-1),5),
                ]))
                story.append(KeepTogether([mini])); continue

        style = S['parent'] if in_parent else S['body']
        story.append(Paragraph(ls, style))

    # ── Footer note ──
    story.append(Spacer(1,12))
    story.append(HRFlowable(width=W,thickness=0.5,color=PDF_BORDER))
    story.append(Spacer(1,4))
    footer_center = f" | {center_name}" if center_name else ""
    story.append(Paragraph(
        f"This report is strictly confidential{footer_center}. Results reflect performance at the time of testing only. "
        "Interpretation should be made in conjunction with clinical observation, history, and other assessment data. "
        "No formal diagnosis is provided herein.",
        S['small']))

    doc.build(story)
    buf.seek(0)
    return buf

def _build_score_table(data, S, W):
    rows = [[
        Paragraph('<b>Scale</b>',S['small']),
        Paragraph('<b>SS</b>',S['small']),
        Paragraph('<b>%ile</b>',S['small']),
        Paragraph('<b>90% CI</b>',S['small']),
        Paragraph('<b>Classification</b>',S['small']),
        Paragraph('<b>Band</b>',S['small']),
    ]]
    extra_cmds = []
    row_i = 1
    for k in IQ_SCORES:
        v = data.get(k)
        if not v: continue
        en_c,_,col = classify(v); pct=percentile_from_ss(v); ci=data.get(f"{k}_ci","—") or "—"
        rows.append([
            Paragraph(f"<b>{IQ_LABELS[k]['en']}</b>",S['body']),
            Paragraph(f"<b>{v}</b>",S['body']),
            Paragraph(f"{pct}th",S['body']),
            Paragraph(ci,S['body']),
            Paragraph(en_c,S['body']),
            Paragraph("",S['body']),
        ])
        extra_cmds.append(('BACKGROUND',(5,row_i),(5,row_i),colors.HexColor(col)))
        extra_cmds.append(('TEXTCOLOR', (5,row_i),(5,row_i),colors.black))
        extra_cmds.append(('FONTNAME',  (5,row_i),(5,row_i),'Helvetica-Bold'))
        row_i += 1
    # sub-header
    rows.append([Paragraph('<b>FACTOR INDEX SCORES</b>',S['small']),
                 Paragraph('',S['small']),Paragraph('',S['small']),
                 Paragraph('',S['small']),Paragraph('',S['small']),Paragraph('',S['small'])])
    extra_cmds.append(('BACKGROUND',(0,row_i),(-1,row_i),colors.HexColor('#2E6DB4')))
    extra_cmds.append(('TEXTCOLOR', (0,row_i),(-1,row_i),colors.white))
    extra_cmds.append(('FONTNAME',  (0,row_i),(-1,row_i),'Helvetica-Bold'))
    row_i += 1
    for k in FACTOR_SCORES:
        v = data.get(k)
        if not v: continue
        en_c,_,col = classify(v); pct=percentile_from_ss(v); ci=data.get(f"{k}_ci","—") or "—"
        bg = _shade2() if row_i%2==0 else colors.white
        rows.append([
            Paragraph(FACTOR_LABELS[k]['en'],S['body']),
            Paragraph(f"<b>{v}</b>",S['body']),
            Paragraph(f"{pct}th",S['body']),
            Paragraph(ci,S['body']),
            Paragraph(en_c,S['body']),
            Paragraph("",S['body']),
        ])
        extra_cmds.append(('BACKGROUND',(0,row_i),(4,row_i),bg))
        extra_cmds.append(('BACKGROUND',(5,row_i),(5,row_i),colors.HexColor(col)))
        extra_cmds.append(('TEXTCOLOR', (5,row_i),(5,row_i),colors.black))
        extra_cmds.append(('FONTNAME',  (5,row_i),(5,row_i),'Helvetica-Bold'))
        row_i += 1
    tbl = _score_tbl(rows,[6*cm,1.5*cm,1.8*cm,2.5*cm,4*cm,1.7*cm],extra_cmds)
    return tbl

def _build_subtest_table(data, S, W):
    has_any = any(data.get(f"nv_{k.lower()}") or data.get(f"v_{k.lower()}") for k in FACTOR_SCORES)
    if not has_any: return None
    rows = [[
        Paragraph('<b>Subtest</b>',S['small']),
        Paragraph('<b>Nonverbal SS</b>',S['small']),
        Paragraph('<b>Verbal SS</b>',S['small']),
        Paragraph('<b>NV Classification</b>',S['small']),
        Paragraph('<b>V Classification</b>',S['small']),
    ]]
    extra = []
    for i,k in enumerate(FACTOR_SCORES):
        nv=data.get(f"nv_{k.lower()}"); v=data.get(f"v_{k.lower()}")
        bg = _shade1() if i%2==0 else colors.white
        nv_c = classify(nv)[0] if nv else "—"
        v_c  = classify(v)[0]  if v  else "—"
        rows.append([
            Paragraph(FACTOR_LABELS[k]['en'],S['body']),
            Paragraph(f"<b>{nv}</b>" if nv else "—",S['body']),
            Paragraph(f"<b>{v}</b>"  if v  else "—",S['body']),
            Paragraph(nv_c,S['body']),
            Paragraph(v_c, S['body']),
        ])
        extra.append(('BACKGROUND',(0,i+1),(-1,i+1),bg))
    tbl = _score_tbl(rows,[5.5*cm,2.5*cm,2.5*cm,3.5*cm,3.5*cm],extra)
    return tbl

def _build_legend_table(S, W):
    entries = [
        ("#1565C0","T ≥ 130","Gifted / Very Gifted","Exceptional cognitive ability; highly enriched programming recommended."),
        ("#0288D1","120–129","Superior","Well above average; advanced academic enrichment beneficial."),
        ("#00897B","110–119","High Average","Above average; likely succeeds with standard instruction."),
        ("#388E3C","90–109","Average","Typical cognitive development for age."),
        ("#F9A825","80–89", "Low Average","Mild difficulties; monitoring and support may be needed."),
        ("#EF6C00","70–79", "Borderline","Significant difficulties; intervention recommended."),
        ("#E53935","55–69", "Mild Impairment","Substantial support and adapted instruction needed."),
        ("#B71C1C","40–54", "Moderate Impairment","Intensive support and specialized programming required."),
    ]
    rows = [[Paragraph('<b>Colour</b>',S['small']),Paragraph('<b>SS Range</b>',S['small']),
             Paragraph('<b>Classification</b>',S['small']),Paragraph('<b>Meaning</b>',S['small'])]]
    cmds = [
        ('BACKGROUND',(0,0),(-1,0),_hdr_style()),
        ('TEXTCOLOR', (0,0),(-1,0),colors.white),
        ('FONTNAME',  (0,0),(-1,0),'Helvetica-Bold'),
        ('FONTSIZE',  (0,0),(-1,-1),8.5),
        ('BOX',       (0,0),(-1,-1),0.5,PDF_BORDER),
        ('INNERGRID', (0,0),(-1,-1),0.3,PDF_BORDER),
        ('VALIGN',    (0,0),(-1,-1),'MIDDLE'),
        ('TOPPADDING',(0,0),(-1,-1),4),
        ('BOTTOMPADDING',(0,0),(-1,-1),4),
        ('LEFTPADDING',(0,0),(-1,-1),5),
    ]
    for i,(col,rng,classif,meaning) in enumerate(entries,start=1):
        rows.append([Paragraph("",S['small']),Paragraph(rng,S['body']),
                     Paragraph(f"<b>{classif}</b>",S['body']),Paragraph(meaning,S['body'])])
        cmds.append(('BACKGROUND',(0,i),(0,i),colors.HexColor(col)))
        cmds.append(('TEXTCOLOR', (0,i),(0,i),colors.black))
        cmds.append(('ROWBACKGROUNDS',(1,i),(-1,i),
                     [colors.white if i%2==0 else colors.HexColor('#F5F7FA')]))
    tbl = Table(rows, colWidths=[1.2*cm,2.5*cm,3.8*cm,W-7.5*cm], repeatRows=1)
    tbl.setStyle(TableStyle(cmds))
    return tbl

# ══════════════════════════════════════════════════════════════
#  EMAIL  (PDF only)
# ══════════════════════════════════════════════════════════════
def send_email(data, buf_pdf, fn_pdf):
    name   = data.get("name","—") or "—"
    fsiq   = data.get("FSIQ")
    center = data.get("_center_name","") or ""
    en_c, ar_c, _ = classify(fsiq) if fsiq else ("—","—","#888")
    date_str = date.today().strftime('%B %d, %Y')
    center_line = (f"<tr><td style='padding:5px 0;color:#555;width:40%;'>Center</td>"
                   f"<td><strong>{center}</strong></td></tr>") if center else ""
    subject_center = f" | {center}" if center else ""

    msg = MIMEMultipart('mixed')
    msg['From']    = GMAIL_USER
    msg['To']      = RECIPIENT_EMAIL
    msg['Subject'] = f"[SB5 Report] {name}{subject_center} — {date_str}"

    body = f"""<html><body style="font-family:Georgia,serif;color:#1A1A2E;background:#F5F7FA;padding:20px;">
  <div style="max-width:560px;margin:0 auto;background:white;border:1px solid #C8922A;border-radius:8px;padding:28px;">
    <h2 style="font-weight:600;font-size:18px;color:#1B3A6B;margin-bottom:2px;">Stanford-Binet 5 — Assessment Report</h2>
    <p style="color:#888;font-size:11px;margin-top:0;">SB5 Psychological Assessment — Auto-generated from uploaded report</p>
    <hr style="border:none;border-top:2px solid #C8922A;margin:16px 0;">
    <table style="width:100%;font-size:13px;border-collapse:collapse;">
      {center_line}
      <tr><td style="padding:5px 0;color:#555;width:40%;">Examinee</td><td><strong>{name}</strong></td></tr>
      <tr><td style="padding:5px 0;color:#555;">Age</td><td>{data.get("age","—") or "—"}</td></tr>
      <tr><td style="padding:5px 0;color:#555;">FSIQ</td><td><strong style="color:#1B3A6B;">{fsiq or "—"}</strong></td></tr>
      <tr><td style="padding:5px 0;color:#555;">Classification</td><td>{en_c}</td></tr>
      <tr><td style="padding:5px 0;color:#555;">Examiner</td><td>{data.get("examiner","—") or "—"}</td></tr>
      <tr><td style="padding:5px 0;color:#555;">Test Date</td><td>{data.get("test_date","—") or "—"}</td></tr>
      <tr><td style="padding:5px 0;color:#555;">Report Date</td><td>{date_str}</td></tr>
    </table>
    <hr style="border:none;border-top:1px solid #DDE5F8;margin:16px 0;">
    <p style="font-size:12px;line-height:1.7;">
    📄 <strong>English Report (PDF)</strong> attached — premium clinical report with charts and tables.</p>
    <p style="font-size:10px;color:#888;font-style:italic;">Confidential — for the evaluating clinician only.</p>
  </div></body></html>"""

    msg.attach(MIMEText(body, 'html'))
    buf_pdf.seek(0)
    part = MIMEBase('application', 'pdf')
    part.set_payload(buf_pdf.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment', filename=fn_pdf)
    msg.attach(part)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as srv:
        srv.login(GMAIL_USER, GMAIL_PASS)
        srv.sendmail(GMAIL_USER, RECIPIENT_EMAIL, msg.as_string())

# ══════════════════════════════════════════════════════════════
#  PAGE CONFIG & CSS
# ══════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Stanford-Binet 5 Report",
    page_icon="🎓",
    layout="centered",
    initial_sidebar_state="collapsed",
)
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;700&family=Inter:wght@300;400;500;600&display=swap');
html,body,[class*="css"]{{font-family:'Inter',sans-serif;background:{LIGHT_BG};}}
.stApp{{background:{LIGHT_BG};}}
#MainMenu{{visibility:hidden!important;}}
header[data-testid="stHeader"]{{visibility:hidden!important;}}
footer{{visibility:hidden!important;}}
[data-testid="stToolbar"]{{display:none!important;}}

.sb5-header{{
    background:linear-gradient(135deg,{DEEP_BLUE} 0%,{MID_BLUE} 60%,{DEEP_BLUE} 100%);
    border-radius:16px;padding:28px 36px;margin-bottom:24px;
    box-shadow:0 6px 24px rgba(27,58,107,0.3);border-bottom:4px solid {GOLD};
    text-align:center;
}}
.sb5-header h1{{color:white;font-family:'Playfair Display',serif;font-size:1.9rem;font-weight:700;margin:0 0 6px;}}
.sb5-header p{{color:#A8BFDD;font-size:0.87rem;margin:0;}}

.upload-card{{
    background:white;border-radius:12px;padding:32px 36px;
    box-shadow:0 2px 16px rgba(27,58,107,0.1);
    border:2px dashed {MID_BLUE};text-align:center;margin:16px 0;
}}
.upload-card h3{{color:{DEEP_BLUE};font-size:1.2rem;margin-bottom:8px;}}
.upload-card p{{color:#666;font-size:0.88rem;line-height:1.7;}}

.thank-you{{text-align:center;padding:4rem 2rem;}}
.thank-you h2{{font-family:'Playfair Display',serif;font-size:2rem;font-weight:400;color:{DEEP_BLUE};margin-bottom:1rem;}}
.thank-you p{{color:{MID_BLUE};font-size:.95rem;line-height:1.9;}}

.stButton>button{{
    background:{DEEP_BLUE}!important;color:white!important;border:none!important;
    border-radius:10px!important;padding:12px 32px!important;font-size:14px!important;
    font-weight:600!important;box-shadow:0 3px 12px rgba(27,58,107,0.3)!important;
    transition:all 0.2s!important;
}}
.stButton>button:hover{{background:{MID_BLUE}!important;}}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
#  HEADER
# ══════════════════════════════════════════════════════════════
col_logo, col_hdr = st.columns([1,5])
with col_logo:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=110)
with col_hdr:
    st.markdown("""
    <div class="sb5-header">
        <h1>🎓 Stanford-Binet Intelligence Scales, 5th Ed.</h1>
        <p>مقياس ستانفورد-بينيه للذكاء — الصورة الخامسة&nbsp;&nbsp;·&nbsp;&nbsp;
        Upload your Arabic SB5 report → receive a premium English PDF report by email</p>
    </div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
#  SESSION STATE
# ══════════════════════════════════════════════════════════════
if "done" not in st.session_state: st.session_state.done = False

# ══════════════════════════════════════════════════════════════
#  THANK-YOU SCREEN
# ══════════════════════════════════════════════════════════════
if st.session_state.done:
    data        = st.session_state.get("last_data", {})
    name        = data.get("name","—") or "—"
    fsiq        = data.get("FSIQ")
    center_done = data.get("_center_name","")
    en_c, _, _  = classify(fsiq) if fsiq else ("—","—","#888")
    logo_b = st.session_state.get("center_logo_bytes")
    if logo_b:
        c1,c2,c3 = st.columns([1,2,1])
        with c2: st.image(logo_b, use_container_width=True)
    elif os.path.exists(LOGO_PATH):
        c1,c2,c3 = st.columns([1,2,1])
        with c2: st.image(LOGO_PATH, use_container_width=True)
    center_line = (f"<p style='color:{MID_BLUE};font-size:.9rem;margin-bottom:.5rem;'>"
                   f"<strong>{center_done}</strong></p>") if center_done else ""
    st.markdown(f"""<div class="thank-you">
        {center_line}
        <h2>Report Sent Successfully</h2>
        <p><strong>{name}</strong></p>
        <p>FSIQ: <strong>{fsiq or "—"}</strong> &nbsp;·&nbsp; {en_c}</p>
        <p style="margin-top:1.2rem;font-size:.87rem;">
            The English PDF report has been sent to the clinic email.<br>
            تم إرسال التقرير الإنجليزي (PDF) إلى البريد الإلكتروني للعيادة.
        </p>
    </div>""", unsafe_allow_html=True)
    _, btn_col, _ = st.columns([2,2,2])
    with btn_col:
        if st.button("↺ Upload New Report", use_container_width=True):
            for k in list(st.session_state.keys()): del st.session_state[k]
            st.rerun()
    st.stop()

# ══════════════════════════════════════════════════════════════
#  UPLOAD SECTION
# ══════════════════════════════════════════════════════════════
st.markdown("""<div class="upload-card">
    <h3>Upload Arabic SB5 Report</h3>
    <p>Upload the Arabic <strong>Word report (.docx)</strong> — copy-paste the text from your offline
    software into a new Word document and save as .docx.<br>
    The app extracts all data automatically and emails you a premium English PDF report.<br><br>
    <strong>الرجاء رفع التقرير العربي بصيغة Word (.docx)</strong> — سيتم استخراج البيانات تلقائياً
    وإرسال التقرير الإنجليزي (PDF) إلى البريد الإلكتروني.</p>
</div>""", unsafe_allow_html=True)

# ── Center info ──
st.markdown("<div style='background:white;border-radius:10px;padding:20px 24px;"
            "box-shadow:0 2px 12px rgba(27,58,107,0.08);border-left:4px solid #C8922A;"
            "margin-bottom:16px;'>", unsafe_allow_html=True)
col_cn, col_logo_up = st.columns([3, 2])
with col_cn:
    st.markdown("<div style='font-size:.75rem;font-weight:700;color:#1B3A6B;"
                "text-transform:uppercase;letter-spacing:.08em;margin-bottom:6px;'>"
                "Center / Clinic Name &nbsp;·&nbsp; اسم المركز / العيادة</div>",
                unsafe_allow_html=True)
    center_name = st.text_input("center_name", value="",
                                placeholder="e.g. Wijdan Therapy Center / مركز وجدان للعلاج النفسي",
                                label_visibility="collapsed", key="center_name_inp")
with col_logo_up:
    st.markdown("<div style='font-size:.75rem;font-weight:700;color:#1B3A6B;"
                "text-transform:uppercase;letter-spacing:.08em;margin-bottom:6px;'>"
                "Center Logo (optional)</div>", unsafe_allow_html=True)
    logo_upload = st.file_uploader("Center logo", type=["png","jpg","jpeg"],
                                   label_visibility="collapsed", key="logo_upload_inp")
    if logo_upload:
        st.session_state["center_logo_bytes"] = logo_upload.read()
        st.success("✓ Logo uploaded")
st.markdown("</div>", unsafe_allow_html=True)

st.markdown("<div style='font-size:.75rem;font-weight:700;color:#1B3A6B;"
            "text-transform:uppercase;letter-spacing:.08em;margin-bottom:6px;'>"
            "Arabic SB5 Report (.docx)</div>", unsafe_allow_html=True)

uploaded = st.file_uploader(
    "Choose Arabic SB5 Word report",
    type=["docx"],
    label_visibility="collapsed",
)

if uploaded:
    safe_name = re.sub(r'[^\w\-]', '_',
                       re.sub(r'\.docx?$', '', uploaded.name, flags=re.IGNORECASE))
    fn_pdf = f"{safe_name}_EN_Report.pdf"

    with st.spinner("⏳ Reading report, extracting data, generating charts and report — please wait 30–60 seconds..."):

        # 1. Extract raw text
        raw_text = extract_word_text(uploaded)
        if not raw_text.strip():
            st.error("Could not extract text from this file. "
                     "Please make sure you copied the report text into a .docx file.")
            st.stop()

        # 2. Parse structured data
        data = extract_data_from_text(raw_text)
        if not data:
            st.error("Could not parse the report data. Please check the uploaded file.")
            st.stop()

        # 2b. Translate/transliterate any Arabic demographic fields to English
        data = translate_demographics(data)

        # 3. Generate charts
        charts = {}
        pb = make_profile_chart(data)
        if pb: charts["profile"] = pb
        fsiq_val = data.get("FSIQ")
        if fsiq_val: charts["gauge"] = make_classification_gauge(fsiq_val)
        disc_b = make_discrepancy_chart(data)
        if disc_b: charts["discrepancy"] = disc_b
        radar_b = make_factor_radar(data)
        if radar_b: charts["radar"] = radar_b
        sub_b = make_subtest_chart(data)
        if sub_b: charts["subtest"] = sub_b

        # 4. Generate English narrative report only
        report_en = generate_en_report(data)

        # 5. Build English PDF
        center_name_v = center_name.strip() if center_name else ""
        logo_bytes_v  = st.session_state.get("center_logo_bytes", None)
        buf_pdf = build_pdf_report(report_en, data, charts, center_name_v, logo_bytes_v)

        # 6. Send email
        try:
            send_email(data, buf_pdf, fn_pdf)
        except Exception as e:
            st.warning(f"Report generated but email failed: {e}")

        # 7. Save state and redirect
        data["_center_name"] = center_name_v
        st.session_state["last_data"] = data
        st.session_state.done = True
        st.rerun()
