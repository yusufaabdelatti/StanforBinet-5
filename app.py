import streamlit as st
from groq import Groq
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io, os, smtplib, re, json
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import date
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

# ══════════════════════════════════════════════════════════════
#  CONFIG
# ══════════════════════════════════════════════════════════════
GMAIL_USER      = "Wijdan.psyc@gmail.com"
GMAIL_PASS      = "rias eeul lyuu stce"
RECIPIENT_EMAIL = "Wijdan.psyc@gmail.com"
LOGO_PATH       = os.path.join(os.path.dirname(__file__), "logo.png")

DEEP_BLUE  = "#1B3A6B"
MID_BLUE   = "#2E6DB4"
GOLD       = "#C8922A"
LIGHT_BG   = "#F5F7FA"
DARK       = "#1A1A2E"

DEEP_BLUE_RGB = RGBColor(0x1B, 0x3A, 0x6B)
MID_BLUE_RGB  = RGBColor(0x2E, 0x6D, 0xB4)
GOLD_RGB      = RGBColor(0xC8, 0x92, 0x2A)
DARK_RGB      = RGBColor(0x1A, 0x1A, 0x2E)

# ══════════════════════════════════════════════════════════════
#  SB-5 STRUCTURE
# ══════════════════════════════════════════════════════════════
IQ_SCORES = ["FSIQ", "NVIQ", "VIQ"]
FACTOR_SCORES = ["FR", "KN", "QR", "VS", "WM"]

IQ_LABELS = {
    "FSIQ": {"en": "Full Scale IQ", "ar": "درجة الذكاء الكلية"},
    "NVIQ": {"en": "Nonverbal IQ",  "ar": "درجة المجال غير اللفظي"},
    "VIQ":  {"en": "Verbal IQ",     "ar": "درجة المجال اللفظي"},
}
FACTOR_LABELS = {
    "FR": {"en": "Fluid Reasoning",       "ar": "الاستدلال التحليلي"},
    "KN": {"en": "Knowledge",             "ar": "المعلومات"},
    "QR": {"en": "Quantitative Reasoning","ar": "الاستدلال الكمي"},
    "VS": {"en": "Visual-Spatial",        "ar": "المعالجة البصرية المكانية"},
    "WM": {"en": "Working Memory",        "ar": "الذاكرة العاملة"},
}

CLASSIFICATIONS = [
    (145, 160, "Very Gifted",          "موهوب بشدة",        "#1565C0"),
    (130, 144, "Gifted",               "موهوب",             "#1976D2"),
    (120, 129, "Superior",             "متفوق",             "#0288D1"),
    (110, 119, "High Average",         "متوسط مرتفع",       "#00897B"),
    (90,  109, "Average",              "متوسط",             "#388E3C"),
    (80,   89, "Low Average",          "متوسط منخفض",       "#F9A825"),
    (70,   79, "Borderline",           "الفئة البينية",     "#EF6C00"),
    (55,   69, "Mild Impairment",      "إعاقة بسيطة",      "#E53935"),
    (40,   54, "Moderate Impairment",  "إعاقة متوسطة",     "#B71C1C"),
    (0,    39, "Severe Impairment",    "إعاقة شديدة",      "#880E4F"),
]

def classify(ss: int):
    for lo, hi, en, ar, color in CLASSIFICATIONS:
        if lo <= ss <= hi:
            return en, ar, color
    return "Unknown", "غير محدد", "#888888"

def percentile_from_ss(ss: int) -> int:
    """Approximate percentile from standard score (mean=100, sd=15)."""
    from math import erf, sqrt
    z = (ss - 100) / 15.0
    p = 0.5 * (1 + erf(z / sqrt(2)))
    return max(1, min(99, round(p * 100)))

# ══════════════════════════════════════════════════════════════
#  CHART GENERATORS
# ══════════════════════════════════════════════════════════════
def make_profile_chart(scores: dict, lang: str = "en") -> bytes:
    """Horizontal bar chart of all IQ + Factor scores with classification bands."""
    fig, ax = plt.subplots(figsize=(11, 6.5))
    fig.patch.set_facecolor('#FAFBFF')
    ax.set_facecolor('#FAFBFF')

    all_keys   = ["FSIQ", "NVIQ", "VIQ", "FR", "KN", "QR", "VS", "WM"]
    all_labels = []
    all_vals   = []
    all_colors = []

    for k in all_keys:
        if k not in scores or scores[k] is None:
            continue
        v = scores[k]
        lbl = (FACTOR_LABELS if k in FACTOR_LABELS else IQ_LABELS)[k][lang]
        _, _, color = classify(v)
        all_labels.append(lbl)
        all_vals.append(v)
        all_colors.append(color)

    y_pos = np.arange(len(all_labels))

    # Classification band background
    bands = [(40,54,"#FCE4EC"),(55,69,"#FFEBEE"),(70,79,"#FFF3E0"),
             (80,89,"#FFFDE7"),(90,109,"#F1F8E9"),(110,119,"#E3F2FD"),
             (120,129,"#E1F5FE"),(130,145,"#E8EAF6")]
    for lo, hi, col in bands:
        ax.axvspan(lo, hi, alpha=0.35, color=col, zorder=0)

    bars = ax.barh(y_pos, all_vals, color=all_colors, height=0.6,
                   edgecolor='white', linewidth=1.2, zorder=3)

    # Divider between IQ and Factor sections
    n_iq = sum(1 for k in all_keys if k in IQ_LABELS and k in scores and scores[k])
    if n_iq > 0 and n_iq < len(all_labels):
        ax.axhline(y=n_iq - 0.5, color='#BDBDBD', linestyle='--', linewidth=1, zorder=4)

    # Mean line
    ax.axvline(x=100, color='#1B3A6B', linestyle='-', linewidth=1.8, alpha=0.6,
               label='Mean (100)' if lang=='en' else 'المتوسط (100)')
    ax.axvline(x=85,  color='#EF6C00', linestyle=':', linewidth=1.2, alpha=0.5)
    ax.axvline(x=115, color='#EF6C00', linestyle=':', linewidth=1.2, alpha=0.5)

    # Value labels
    for bar_, val in zip(bars, all_vals):
        en_c, ar_c, _ = classify(val)
        cat = en_c if lang == 'en' else ar_c
        pct = percentile_from_ss(val)
        label = f"{val}  ({cat}, {pct}th)" if lang=='en' else f"{val}  ({cat})"
        ax.text(bar_.get_width() + 1, bar_.get_y() + bar_.get_height()/2,
                label, va='center', ha='left', fontsize=8.5, color='#1A1A2E',
                fontfamily='DejaVu Sans')

    ax.set_yticks(y_pos)
    ax.set_yticklabels(all_labels, fontsize=10, fontfamily='DejaVu Sans')
    ax.set_xlim(40, 165)
    ax.set_xlabel("Standard Score" if lang=='en' else "الدرجة المعيارية",
                  fontsize=11, color='#1A1A2E')
    title = "Stanford-Binet 5 — Score Profile" if lang=='en' else "ستانفورد-بينيه 5 — ملف الدرجات"
    ax.set_title(title, fontsize=13, fontweight='bold', color='#1B3A6B', pad=12)
    ax.legend(fontsize=9, framealpha=0.7)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.grid(axis='x', linestyle=':', alpha=0.4, zorder=1)

    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    plt.close(fig)
    buf.seek(0)
    return buf.read()

def make_factor_radar(scores: dict, lang: str = "en") -> bytes:
    """Radar chart of the 5 factor scores."""
    labels = [FACTOR_LABELS[k][lang] for k in FACTOR_SCORES if k in scores and scores[k]]
    vals   = [scores[k] for k in FACTOR_SCORES if k in scores and scores[k]]
    if len(vals) < 3:
        return None

    N = len(labels)
    angles = np.linspace(0, 2 * np.pi, N, endpoint=False).tolist()
    vals_   = vals + [vals[0]]
    angles += [angles[0]]

    fig, ax = plt.subplots(figsize=(6, 6), subplot_kw=dict(polar=True))
    fig.patch.set_facecolor('#FAFBFF')
    ax.set_facecolor('#EEF2FF')

    # Reference circles
    for ref, col in [(70,'#EF6C00'),(100,'#1B3A6B'),(130,'#0288D1')]:
        ref_norm = ref / 160.0
        ax.plot([a for a in angles[:-1]] + [angles[0]],
                [ref_norm] * N + [ref_norm],
                '--', color=col, linewidth=0.8, alpha=0.5)

    vals_norm = [v/160.0 for v in vals_]
    ax.plot(angles, vals_norm, 'o-', linewidth=2, color='#2E6DB4', markersize=6)
    ax.fill(angles, vals_norm, alpha=0.28, color='#2E6DB4')

    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(labels, size=9.5, fontfamily='DejaVu Sans')
    ax.set_ylim(0, 1)
    ax.set_yticks([70/160, 85/160, 100/160, 115/160, 130/160])
    ax.set_yticklabels(['70','85','100','115','130'], size=7.5, color='#666')
    title = "Factor Index Profile" if lang=='en' else "ملف مؤشرات العوامل"
    ax.set_title(title, size=12, fontweight='bold', color='#1B3A6B', pad=18)

    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    plt.close(fig)
    buf.seek(0)
    return buf.read()

def make_classification_gauge(fsiq: int) -> bytes:
    """A visual gauge/meter showing where FSIQ falls."""
    fig, ax = plt.subplots(figsize=(8, 3.5))
    fig.patch.set_facecolor('#FAFBFF')
    ax.set_facecolor('#FAFBFF')

    # Draw classification bands as colored segments
    bands_ordered = [
        (40,  54,  "#B71C1C", "Moderate\nImpairment"),
        (55,  69,  "#E53935", "Mild\nImpairment"),
        (70,  79,  "#EF6C00", "Borderline"),
        (80,  89,  "#F9A825", "Low Average"),
        (90,  109, "#388E3C", "Average"),
        (110, 119, "#00897B", "High Avg"),
        (120, 129, "#0288D1", "Superior"),
        (130, 145, "#1565C0", "Gifted+"),
    ]

    for lo, hi, col, lbl in bands_ordered:
        ax.barh(0, hi - lo, left=lo, height=0.5, color=col, alpha=0.8, edgecolor='white')
        mid = (lo + hi) / 2
        ax.text(mid, 0, lbl, ha='center', va='center', fontsize=7,
                color='white', fontweight='bold', fontfamily='DejaVu Sans')

    # Arrow for FSIQ
    ax.annotate('', xy=(fsiq, 0.32), xytext=(fsiq, 0.65),
                arrowprops=dict(arrowstyle='->', color='#1A1A2E', lw=2.5))
    en_c, _, _ = classify(fsiq)
    ax.text(fsiq, 0.8, f"FSIQ = {fsiq}\n{en_c}", ha='center', va='bottom',
            fontsize=11, fontweight='bold', color='#1A1A2E')

    ax.set_xlim(40, 145)
    ax.set_ylim(-0.4, 1.1)
    ax.set_xlabel("Standard Score", fontsize=10)
    ax.set_yticks([])
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.set_title("Full Scale IQ Classification", fontsize=12, fontweight='bold',
                 color='#1B3A6B', pad=8)

    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    plt.close(fig)
    buf.seek(0)
    return buf.read()

def make_subtest_chart(nv_subtests: dict, v_subtests: dict) -> bytes:
    """Side-by-side NV and V subtest scaled scores."""
    keys = ["FR", "KN", "QR", "VS", "WM"]
    nv_vals = [nv_subtests.get(k) for k in keys]
    v_vals  = [v_subtests.get(k)  for k in keys]

    if all(v is None for v in nv_vals + v_vals):
        return None

    labels = [FACTOR_LABELS[k]["en"] for k in keys]
    x = np.arange(len(labels))
    w = 0.35

    fig, ax = plt.subplots(figsize=(10, 5))
    fig.patch.set_facecolor('#FAFBFF')
    ax.set_facecolor('#FAFBFF')

    nv_clean = [v if v is not None else 0 for v in nv_vals]
    v_clean  = [v if v is not None else 0 for v in v_vals]

    b1 = ax.bar(x - w/2, nv_clean, w, label='Nonverbal', color='#2E6DB4',
                edgecolor='white', linewidth=0.8)
    b2 = ax.bar(x + w/2, v_clean,  w, label='Verbal',    color='#C8922A',
                edgecolor='white', linewidth=0.8)

    ax.axhline(y=10, color='#1B3A6B', linestyle='--', linewidth=1.5,
               alpha=0.6, label='Mean (10)')
    ax.axhline(y=7,  color='#EF6C00', linestyle=':', linewidth=1.0, alpha=0.5)
    ax.axhline(y=13, color='#EF6C00', linestyle=':', linewidth=1.0, alpha=0.5)

    for bar_ in list(b1) + list(b2):
        h = bar_.get_height()
        if h > 0:
            ax.text(bar_.get_x() + bar_.get_width()/2., h + 0.1,
                    str(int(h)), ha='center', va='bottom', fontsize=9, fontweight='bold')

    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=15, ha='right', fontsize=9.5)
    ax.set_ylabel("Scaled Score (Mean=10, SD=3)", fontsize=10)
    ax.set_title("Subtest Scaled Scores — Nonverbal vs. Verbal", fontsize=12,
                 fontweight='bold', color='#1B3A6B')
    ax.set_ylim(0, 20)
    ax.legend(fontsize=10)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.grid(axis='y', linestyle=':', alpha=0.4)

    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    plt.close(fig)
    buf.seek(0)
    return buf.read()

# ══════════════════════════════════════════════════════════════
#  GROQ REPORT GENERATION
# ══════════════════════════════════════════════════════════════
def build_score_summary(data: dict) -> str:
    lines = ["IQ SCORES:"]
    for k in IQ_SCORES:
        if data.get(k):
            en_c, _, _ = classify(data[k])
            pct = percentile_from_ss(data[k])
            ci = data.get(f"{k}_ci", "—")
            lines.append(f"  {IQ_LABELS[k]['en']}: SS={data[k]}, %ile={pct}, CI={ci}, Classification={en_c}")
    lines.append("\nFACTOR INDEX SCORES:")
    for k in FACTOR_SCORES:
        if data.get(k):
            en_c, _, _ = classify(data[k])
            pct = percentile_from_ss(data[k])
            ci = data.get(f"{k}_ci", "—")
            lines.append(f"  {FACTOR_LABELS[k]['en']}: SS={data[k]}, %ile={pct}, CI={ci}, Classification={en_c}")
    lines.append("\nSUBTEST SCALED SCORES (NV | V):")
    for k in FACTOR_SCORES:
        nv = data.get(f"nv_{k.lower()}")
        v  = data.get(f"v_{k.lower()}")
        if nv or v:
            lines.append(f"  {FACTOR_LABELS[k]['en']}: NV={nv or '—'} | V={v or '—'}")
    return "\n".join(lines)

def generate_en_report(data: dict) -> str:
    score_summary = build_score_summary(data)
    name     = data.get("name","the examinee")
    age      = data.get("age","—")
    gender   = data.get("gender","they/them")
    pronoun  = "he" if "male" in gender.lower() or "ذكر" in gender else "she"
    examiner = data.get("examiner","—")
    referral = data.get("referral","—")
    complaints = data.get("complaints","—")

    prompt = f"""You are a senior licensed psychologist writing a world-class Stanford-Binet 5 (SB5) assessment report.
Use the most current research on the SB5 (Roid, 2003; Roid et al., 2016) and CHC theory of intelligence.

EXAMINEE: {name} | AGE: {age} | GENDER: {gender} | EXAMINER: {examiner}
REFERRAL SOURCE: {referral}
REASON FOR REFERRAL: {complaints}
TEST DATE: {data.get("test_date","—")}
REPORT DATE: {date.today().strftime('%B %d, %Y')}

SCORE SUMMARY:
{score_summary}

BEHAVIORAL OBSERVATIONS:
{data.get("behavioral_obs","Not provided")}

BACKGROUND INFORMATION:
{data.get("background","Not provided")}

WRITE A COMPREHENSIVE PROFESSIONAL SB5 REPORT with these sections.
Use formal clinical/psychoeducational language. Be specific to the scores. No markdown symbols.
Section titles: ALL CAPS on their own line.

STANFORD-BINET INTELLIGENCE SCALES, FIFTH EDITION — PSYCHOLOGICAL REPORT
Name | {name}
Date of Birth | {data.get("dob","—")}
Age | {age}
Gender | {gender}
Examiner | {examiner}
Test Date | {data.get("test_date","—")}
Report Date | {date.today().strftime('%B %d, %Y')}
Referral Source | {referral}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

REASON FOR REFERRAL
3–5 sentences: who referred, what was the clinical question, what this evaluation addresses.

BACKGROUND INFORMATION
Relevant developmental, educational, medical, and family history from the provided background.

TESTS ADMINISTERED
List SB5 and any other assessments mentioned.

BEHAVIORAL OBSERVATIONS
Describe behavior during testing: cooperation, attention, affect, language, motor, approach to tasks.
Note any factors that may affect validity.

ASSESSMENT RESULTS AND INTERPRETATION

1. FULL SCALE IQ (FSIQ)
Interpret FSIQ score deeply: what it represents, the confidence interval, percentile rank,
classification, and what it means for this individual's overall cognitive functioning.
Reference CHC theory (Gf-Gc framework).

2. NONVERBAL IQ (NVIQ) AND VERBAL IQ (VIQ)
Compare NV and V domains. Discuss the difference (statistically significant or not).
Clinical implications of any discrepancy.

3. FLUID REASONING (FR)
Interpret the factor score, NV and V subtest scores, and what this means for the examinee's
inductive and deductive reasoning abilities.

4. KNOWLEDGE (KN)
Interpret crystallized intelligence, vocabulary, and acquired knowledge.

5. QUANTITATIVE REASONING (QR)
Interpret numerical and mathematical reasoning abilities.

6. VISUAL-SPATIAL PROCESSING (VS)
Interpret spatial visualization, pattern recognition, and visuoconstructive abilities.

7. WORKING MEMORY (WM)
Interpret short-term memory, attention, and cognitive control.

8. STRENGTHS AND WEAKNESSES PROFILE
Identify relative strengths (highest scores) and weaknesses (lowest scores).
Discuss intra-cognitive variability. Is the profile consistent or scattered?

9. DIAGNOSTIC IMPRESSIONS
Based on the full profile, what patterns emerge? Reference relevant diagnostic considerations
(e.g., learning disability, intellectual disability, giftedness, ADHD, etc.) as hypotheses only.
This section does NOT provide a formal diagnosis.

10. RECOMMENDATIONS
Provide 8–12 specific, evidence-based recommendations across these domains:
a) Educational accommodations and classroom strategies
b) Intervention priorities (cognitive, academic, behavioral)
c) Further evaluation needs
d) Family and home support strategies
e) Therapeutic or clinical referrals if indicated

11. SUMMARY
A 2-paragraph executive summary suitable for school teams and parents.
Paragraph 1: Score profile and what it means.
Paragraph 2: Key strengths, challenges, and priority recommendations.

PARENT-FRIENDLY SUMMARY (clearly labeled as a separate simplified section)
Write a plain-language 2-paragraph explanation for parents/caregivers with NO jargon.
Explain what the test found, what it means for their child day-to-day, and the 3 most important things to do.

Use {pronoun}/{pronoun}s consistently. Reference specific T-scores and percentiles throughout.
Write with depth — this report should stand alone as a complete clinical document."""

    client = Groq(api_key=st.secrets["GROQ_API_KEY"])
    r = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role":"user","content":prompt}],
        max_tokens=4000
    )
    return r.choices[0].message.content.strip()

def generate_ar_report(data: dict) -> str:
    name     = data.get("name","المفحوص")
    age      = data.get("age","—")
    gender   = data.get("gender","—")
    examiner = data.get("examiner","—")
    referral = data.get("referral","—")

    ar_scores = []
    for k in IQ_SCORES:
        if data.get(k):
            _, ar_c, _ = classify(data[k])
            pct = percentile_from_ss(data[k])
            ci = data.get(f"{k}_ci","—")
            ar_scores.append(f"  {IQ_LABELS[k]['ar']}: {data[k]} — {ar_c} (رتبة مئينية: {pct}، مدى الثقة 90%: {ci})")
    for k in FACTOR_SCORES:
        if data.get(k):
            _, ar_c, _ = classify(data[k])
            pct = percentile_from_ss(data[k])
            ci = data.get(f"{k}_ci","—")
            ar_scores.append(f"  {FACTOR_LABELS[k]['ar']}: {data[k]} — {ar_c} (رتبة مئينية: {pct}، مدى الثقة: {ci})")

    prompt = f"""أنت طبيب نفسي متخصص تكتب تقريراً سريرياً شاملاً ومتقدماً لمقياس ستانفورد-بينيه الصورة الخامسة (SB5).
استخدم نظرية CHC في الذكاء (Cattell-Horn-Carroll) وأحدث الأبحاث المتعلقة بالمقياس.

المفحوص: {name} | العمر: {age} | النوع: {gender} | الفاحص: {examiner}
جهة الإحالة: {referral}
سبب الإحالة: {data.get("complaints","—")}
تاريخ التطبيق: {data.get("test_date","—")}
تاريخ التقرير: {date.today().strftime('%Y/%m/%d')}

ملخص الدرجات:
{chr(10).join(ar_scores)}

الملاحظات السلوكية:
{data.get("behavioral_obs","لم تُذكر")}

المعلومات الأساسية:
{data.get("background","لم تُذكر")}

اكتب تقريراً سريرياً نفسياً احترافياً شاملاً بالعربية الفصحى.
لا تستخدم رموز markdown. عناوين الأقسام: أرقام + عناوين واضحة.
استخدم الاتجاه من اليمين إلى اليسار. لا إنجليزية إلا للاختصارات المقبولة (SB5, IQ, CHC).

تقرير مقياس ستانفورد-بينيه للذكاء — الصورة الخامسة
الاسم | {name}
تاريخ الميلاد | {data.get("dob","—")}
العمر | {age}
النوع | {gender}
الفاحص | {examiner}
تاريخ التطبيق | {data.get("test_date","—")}
تاريخ التقرير | {date.today().strftime('%Y/%m/%d')}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

١. سبب الإحالة
فقرة من ٣-٥ جمل: من أحال وما هو التساؤل الإكلينيكي.

٢. المعلومات الأساسية
التاريخ التطوري والتعليمي والطبي والأسري ذات الصلة.

٣. الاختبارات المطبقة
اذكر المقياس وأي أدوات أخرى ذُكرت.

٤. الملاحظات السلوكية
وصف سلوك المفحوص أثناء التطبيق. اذكر العوامل المؤثرة في صدق النتائج.

٥. نتائج التقييم وتفسيرها

أ. درجة الذكاء الكلية
تفسير معمق للدرجة: ماذا تعني، فترة الثقة، الرتبة المئينية، الفئة، الدلالة الإكلينيكية.

ب. الذكاء غير اللفظي واللفظي
مقارنة المجالين. الدلالة الإكلينيكية لأي فرق.

ج. الاستدلال التحليلي
تفسير عامل الاستدلال السائل — درجات المقياسين الفرعيين وما تعنيه.

د. المعلومات
تفسير الذكاء المتبلور — المعرفة العامة، الرصيد المعرفي.

هـ. الاستدلال الكمي
القدرة الرياضية والعددية.

و. المعالجة البصرية المكانية
القدرة على التصور والمعالجة البصرية.

ز. الذاكرة العاملة
الذاكرة قصيرة المدى، الانتباه، الكفاءة المعرفية.

٦. نقاط القوة والقصور
تحديد نقاط القوة النسبية والقصور النسبي بناءً على الملف الكامل.

٧. الانطباعات التشخيصية
الأنماط الإكلينيكية الظاهرة. لا تقديم تشخيص رسمي — فرضيات فقط.

٨. التوصيات
قدم ١٠-١٢ توصية محددة ومبنية على الأدلة في المجالات:
أ) التعليمية والأكاديمية
ب) التدخل العلاجي
ج) تقييمات إضافية
د) دعم الأسرة
هـ) الإحالات العلاجية

٩. الملخص
فقرتان موجزتان للفريق المتخصص.

ملخص للوالدين (قسم مبسط واضح التسمية)
فقرتان بلغة مبسطة للأسرة، بدون مصطلحات تخصصية.
اشرح ما وجده الاختبار، ماذا يعني لطفلهم في حياته اليومية، وأهم ٣ أشياء يجب فعلها.
"""

    client = Groq(api_key=st.secrets["GROQ_API_KEY"])
    r = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role":"user","content":prompt}],
        max_tokens=4000
    )
    return r.choices[0].message.content.strip()

# ══════════════════════════════════════════════════════════════
#  WORD DOC BUILDER
# ══════════════════════════════════════════════════════════════
def build_word_doc(report_text: str, data: dict, charts: dict, lang: str) -> io.BytesIO:
    is_rtl = (lang == "ar")
    doc = Document()

    for sec_ in doc.sections:
        sec_.top_margin    = Cm(2.0)
        sec_.bottom_margin = Cm(2.0)
        sec_.left_margin   = Cm(2.5)
        sec_.right_margin  = Cm(2.5)

    # Page border
    for sec_ in doc.sections:
        sp = sec_._sectPr
        pb = OxmlElement('w:pgBorders')
        pb.set(qn('w:offsetFrom'), 'page')
        for side in ('top','left','bottom','right'):
            b = OxmlElement(f'w:{side}')
            b.set(qn('w:val'), 'single')
            b.set(qn('w:sz'),  '8')
            b.set(qn('w:space'), '24')
            b.set(qn('w:color'), '1B3A6B')
            pb.append(b)
        sp.append(pb)

    # Footer
    for sec_ in doc.sections:
        ft = sec_.footer
        fp = ft.paragraphs[0] if ft.paragraphs else ft.add_paragraph()
        fp.clear(); fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_ = fp.add_run()
        r_.font.size = Pt(9); r_.font.color.rgb = MID_BLUE_RGB
        for tag, text in [('begin',None),(None,' PAGE '),('end',None)]:
            if tag:
                el = OxmlElement('w:fldChar'); el.set(qn('w:fldCharType'), tag); r_._r.append(el)
            else:
                it = OxmlElement('w:instrText'); it.text = text; r_._r.append(it)

    def set_rtl(p):
        if is_rtl:
            pPr = p._p.get_or_add_pPr()
            pPr.append(OxmlElement("w:bidi"))
            jc = OxmlElement("w:jc"); jc.set(qn("w:val"), "right"); pPr.append(jc)

    def add_para(text, bold=False, size=11, color=None, space_before=0,
                 space_after=4, alignment=None, keep_next=False, italic=False):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after  = Pt(space_after)
        if keep_next: p.paragraph_format.keep_with_next = True
        set_rtl(p)
        if alignment: p.alignment = alignment
        r_ = p.add_run(text)
        r_.font.size = Pt(size); r_.font.name = "Arial"
        r_.font.bold = bold; r_.font.italic = italic
        if color: r_.font.color.rgb = color
        return p

    def add_section_title(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after  = Pt(4)
        p.paragraph_format.keep_with_next = True
        set_rtl(p)
        r_ = p.add_run(text.strip())
        r_.font.size = Pt(13); r_.font.name = "Arial"
        r_.font.bold = True; r_.font.color.rgb = DEEP_BLUE_RGB
        pPr  = p._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bot  = OxmlElement('w:bottom')
        bot.set(qn('w:val'),'single'); bot.set(qn('w:sz'),'8')
        bot.set(qn('w:space'),'2');    bot.set(qn('w:color'),'1B3A6B')
        pBdr.append(bot); pPr.append(pBdr)

    def make_table(col_widths, header_color='1B3A6B'):
        t = doc.add_table(rows=0, cols=len(col_widths))
        t.style = 'Table Grid'
        try:
            tPr = t._tbl.tblPr
            if is_rtl:
                bv = OxmlElement('w:bidiVisual'); tPr.append(bv)
            tW = OxmlElement('w:tblW')
            tW.set(qn('w:w'),'9026'); tW.set(qn('w:type'),'dxa'); tPr.append(tW)
            tg = OxmlElement('w:tblGrid')
            for w in col_widths:
                gc = OxmlElement('w:gridCol'); gc.set(qn('w:w'), str(w)); tg.append(gc)
            t._tbl.insert(0, tg)
        except: pass
        return t

    def add_table_row(table, cells_data, is_header=False, shade=None):
        row  = table.add_row()
        trPr = row._tr.get_or_add_trPr()
        cs   = OxmlElement('w:cantSplit'); cs.set(qn('w:val'),'1'); trPr.append(cs)
        if is_rtl:
            bidi_ = OxmlElement('w:bidi'); trPr.append(bidi_)
        for cell, (txt, bold_, right_align) in zip(row.cells, cells_data):
            cell.text = ""
            p = cell.paragraphs[0]
            if is_rtl or right_align:
                pPr = p._p.get_or_add_pPr()
                if is_rtl: pPr.append(OxmlElement("w:bidi"))
                jc = OxmlElement("w:jc")
                jc.set(qn("w:val"), "right" if (is_rtl or right_align) else "left")
                pPr.append(jc)
            vr = p.add_run(str(txt))
            vr.font.size = Pt(9.5); vr.font.name = "Arial"; vr.font.bold = bold_
            if is_header: vr.font.color.rgb = RGBColor(0xFF,0xFF,0xFF)
            tc = cell._tc; tcP = tc.get_or_add_tcPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto')
            if is_header:
                shd.set(qn('w:fill'), '1B3A6B')
            elif shade:
                shd.set(qn('w:fill'), shade)
            else:
                shd.set(qn('w:fill'), 'FFFFFF')
            tcP.append(shd)
            mg = OxmlElement('w:tcMar')
            for side in ['top','bottom','left','right']:
                m = OxmlElement(f'w:{side}'); m.set(qn('w:w'),'70'); m.set(qn('w:type'),'dxa'); mg.append(m)
            tcP.append(mg)

    # ── HEADER ──
    p_hdr = doc.add_paragraph()
    p_hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_hdr.paragraph_format.space_after = Pt(6)
    if os.path.exists(LOGO_PATH):
        p_hdr.add_run().add_picture(LOGO_PATH, width=Inches(2.8))

    title_text = ("Stanford-Binet Intelligence Scales, Fifth Edition\nPsychological Assessment Report"
                  if lang=="en"
                  else "مقياس ستانفورد-بينيه للذكاء — الصورة الخامسة\nتقرير التقييم النفسي")
    r_t = p_hdr.add_run(f"\n{title_text}")
    r_t.font.name = "Arial"; r_t.font.size = Pt(16)
    r_t.font.bold = True; r_t.font.color.rgb = DEEP_BLUE_RGB

    sub_text = ("SB5 · C. Keith Roid (2003) · Psychological Corporation"
                if lang=="en" else "SB5 · ترجمة وتقنين أ.د/ صفوت فرج")
    add_para(sub_text, size=9, color=GOLD_RGB,
             alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)

    # Separator
    p_sep = doc.add_paragraph()
    p_sep.paragraph_format.space_before = Pt(2); p_sep.paragraph_format.space_after = Pt(10)
    pPr = p_sep._p.get_or_add_pPr()
    pBdr2 = OxmlElement('w:pBdr')
    bot2  = OxmlElement('w:bottom')
    bot2.set(qn('w:val'),'single'); bot2.set(qn('w:sz'),'12')
    bot2.set(qn('w:space'),'2');    bot2.set(qn('w:color'),'C8922A')
    pBdr2.append(bot2); pPr.append(pBdr2)

    # Client info table
    def row2(f1, v1, f2, v2):
        return [(f1, True, False), (v1, False, False),
                (f2, True, False), (v2, False, False)]

    info_tbl = make_table([2400, 3000, 2400, 3000])
    if lang == "en":
        add_table_row(info_tbl, [("Field",True,False),("",False,False),("Field",True,False),("",False,False)], is_header=True)
        add_table_row(info_tbl, row2("Name:", data.get("name","—"), "Date of Birth:", data.get("dob","—")))
        add_table_row(info_tbl, row2("Age:", data.get("age","—"), "Gender:", data.get("gender","—")), shade="EEF2FF")
        add_table_row(info_tbl, row2("Examiner:", data.get("examiner","—"), "Test Date:", data.get("test_date","—")))
        add_table_row(info_tbl, row2("Referral:", data.get("referral","—"), "Report Date:", date.today().strftime('%B %d, %Y')), shade="EEF2FF")
    else:
        add_table_row(info_tbl, [("الحقل",True,True),("",False,True),("الحقل",True,True),("",False,True)], is_header=True)
        add_table_row(info_tbl, [(data.get("dob","—"),False,True),("تاريخ الميلاد",True,True),(data.get("name","—"),False,True),("الاسم",True,True)])
        add_table_row(info_tbl, [(data.get("gender","—"),False,True),("النوع",True,True),(data.get("age","—"),False,True),("العمر",True,True)], shade="EEF2FF")
        add_table_row(info_tbl, [(data.get("test_date","—"),False,True),("تاريخ التطبيق",True,True),(data.get("examiner","—"),False,True),("الفاحص",True,True)])
        add_table_row(info_tbl, [(date.today().strftime('%Y/%m/%d'),False,True),("تاريخ التقرير",True,True),(data.get("referral","—"),False,True),("جهة الإحالة",True,True)], shade="EEF2FF")
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # IQ Score Summary Table
    sec_title = "IQ AND FACTOR INDEX SCORE SUMMARY" if lang=="en" else "ملخص درجات الذكاء ومؤشرات العوامل"
    add_section_title(sec_title)
    score_tbl = make_table([3500, 1200, 1100, 2200, 2000])
    if lang == "en":
        add_table_row(score_tbl, [("Scale",True,False),("SS",True,False),("%ile",True,False),("90% CI",True,False),("Classification",True,False)], is_header=True)
        for k in IQ_SCORES:
            if data.get(k):
                en_c,_,_ = classify(data[k]); pct=percentile_from_ss(data[k])
                shade_ = "F5F7FF" if IQ_SCORES.index(k)%2==0 else "FFFFFF"
                add_table_row(score_tbl, [(IQ_LABELS[k]['en'],True,False),(str(data[k]),False,False),(str(pct),False,False),(data.get(f"{k}_ci","—"),False,False),(en_c,False,False)], shade=shade_)
        add_table_row(score_tbl, [("FACTOR INDEX SCORES",True,False),("",False,False),("",False,False),("",False,False),("",False,False)], is_header=True)
        for k in FACTOR_SCORES:
            if data.get(k):
                en_c,_,_ = classify(data[k]); pct=percentile_from_ss(data[k])
                shade_ = "FFF8EE" if FACTOR_SCORES.index(k)%2==0 else "FFFFFF"
                add_table_row(score_tbl, [(FACTOR_LABELS[k]['en'],True,False),(str(data[k]),False,False),(str(pct),False,False),(data.get(f"{k}_ci","—"),False,False),(en_c,False,False)], shade=shade_)
    else:
        add_table_row(score_tbl, [("المقياس",True,True),("الدرجة",True,True),("المئيني",True,True),("مدى الثقة",True,True),("الفئة",True,True)], is_header=True)
        for k in IQ_SCORES:
            if data.get(k):
                _,ar_c,_ = classify(data[k]); pct=percentile_from_ss(data[k])
                shade_ = "F5F7FF" if IQ_SCORES.index(k)%2==0 else "FFFFFF"
                add_table_row(score_tbl, [(IQ_LABELS[k]['ar'],True,True),(str(data[k]),False,True),(str(pct),False,True),(data.get(f"{k}_ci","—"),False,True),(ar_c,False,True)], shade=shade_)
        add_table_row(score_tbl, [("درجات مؤشرات العوامل",True,True),("","",""),("","",""),("","",""),("","","")], is_header=True)
        for k in FACTOR_SCORES:
            if data.get(k):
                _,ar_c,_ = classify(data[k]); pct=percentile_from_ss(data[k])
                shade_ = "FFF8EE" if FACTOR_SCORES.index(k)%2==0 else "FFFFFF"
                add_table_row(score_tbl, [(FACTOR_LABELS[k]['ar'],True,True),(str(data[k]),False,True),(str(pct),False,True),(data.get(f"{k}_ci","—"),False,True),(ar_c,False,True)], shade=shade_)
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # Subtest table
    if any(data.get(f"nv_{k.lower()}") or data.get(f"v_{k.lower()}") for k in FACTOR_SCORES):
        sub_title = "SUBTEST SCALED SCORES" if lang=="en" else "الدرجات المعيارية للاختبارات الفرعية"
        add_section_title(sub_title)
        sub_tbl = make_table([4000, 2000, 2000])
        if lang == "en":
            add_table_row(sub_tbl, [("Subtest",True,False),("Nonverbal",True,False),("Verbal",True,False)], is_header=True)
        else:
            add_table_row(sub_tbl, [("الاختبار الفرعي",True,True),("غير لفظي",True,True),("لفظي",True,True)], is_header=True)
        for i, k in enumerate(FACTOR_SCORES):
            nv = data.get(f"nv_{k.lower()}","—"); v = data.get(f"v_{k.lower()}","—")
            shade_ = "F5F7FF" if i%2==0 else "FFFFFF"
            lbl = FACTOR_LABELS[k][lang]
            if lang == "en":
                add_table_row(sub_tbl, [(lbl,True,False),(str(nv),False,False),(str(v),False,False)], shade=shade_)
            else:
                add_table_row(sub_tbl, [(lbl,True,True),(str(nv),False,True),(str(v),False,True)], shade=shade_)
        doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # Charts
    for chart_key, chart_title_en, chart_title_ar in [
        ("profile",    "IQ AND FACTOR SCORE PROFILE", "ملف الدرجات المعيارية"),
        ("gauge",      "FSIQ CLASSIFICATION GAUGE",   "مقياس تصنيف الذكاء الكلي"),
        ("radar",      "FACTOR INDEX RADAR CHART",    "مخطط رادار مؤشرات العوامل"),
        ("subtest",    "SUBTEST COMPARISON CHART",    "مخطط مقارنة الاختبارات الفرعية"),
    ]:
        if charts.get(chart_key):
            title = chart_title_en if lang=="en" else chart_title_ar
            add_section_title(title)
            p_c = doc.add_paragraph()
            p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_c.paragraph_format.space_after = Pt(8)
            w = Inches(5.8) if chart_key in ("profile","subtest") else Inches(4.0)
            p_c.add_run().add_picture(io.BytesIO(charts[chart_key]), width=w)

    # Narrative
    narrative_title = "CLINICAL NARRATIVE REPORT" if lang=="en" else "التقرير السريري التفصيلي"
    add_section_title(narrative_title)

    sec_en_pat = re.compile(r'^\d+\.\s+[A-Z][A-Z\s,&/\(\):\']+$')
    sec_ar_pat = re.compile(r'^[١٢٣٤٥٦٧٨٩أبجدهوزحطيكلمنسعفصقرشت\d]+[\.،:]\s+[\u0600-\u06FF]')
    header_words = {
        "STANFORD-BINET INTELLIGENCE SCALES", "REASON FOR REFERRAL",
        "BACKGROUND INFORMATION", "TESTS ADMINISTERED",
        "BEHAVIORAL OBSERVATIONS", "ASSESSMENT RESULTS AND INTERPRETATION",
        "STRENGTHS AND WEAKNESSES PROFILE", "DIAGNOSTIC IMPRESSIONS",
        "RECOMMENDATIONS", "SUMMARY", "PARENT-FRIENDLY SUMMARY",
        "ملخص للوالدين",
    }

    in_parent_section = False
    in_table = False; current_table = None

    for line in report_text.split('\n'):
        ls = line.strip()
        if not ls:
            if in_table: in_table = False; current_table = None
            doc.add_paragraph().paragraph_format.space_after = Pt(2)
            continue

        upper = ls.upper()
        is_section = (sec_en_pat.match(ls) or sec_ar_pat.match(ls) or
                      any(ls.startswith(h) or upper.startswith(h) for h in header_words))

        if is_section:
            in_table = False; current_table = None
            if "PARENT" in upper or "والدين" in ls or "SIMPLIFIED" in upper:
                in_parent_section = True
                # Gold separator before parent section
                p_sep2 = doc.add_paragraph()
                pPr2   = p_sep2._p.get_or_add_pPr()
                pBdr3  = OxmlElement('w:pBdr')
                t2     = OxmlElement('w:top')
                t2.set(qn('w:val'),'single'); t2.set(qn('w:sz'),'8')
                t2.set(qn('w:space'),'4');    t2.set(qn('w:color'),'C8922A')
                pBdr3.append(t2); pPr2.append(pBdr3)
                p2 = doc.add_paragraph()
                set_rtl(p2)
                r2_ = p2.add_run("⭐ " + ls + " ⭐")
                r2_.font.size = Pt(13); r2_.font.name = "Arial"
                r2_.font.bold = True; r2_.font.color.rgb = GOLD_RGB
            else:
                add_section_title(ls)
            continue

        if ls.startswith('━') or ls.startswith('═') or ls.startswith('---'):
            in_table = False; current_table = None; continue

        if '|' in ls:
            parts = [p.strip() for p in ls.split('|') if p.strip()]
            if not parts: continue
            if all(set(p) <= set('-: ') for p in parts): continue
            if not in_table or current_table is None:
                in_table = True
                n_cols = len(parts)
                w_each = 9026 // n_cols
                current_table = make_table([w_each]*n_cols)
            cells = [(p, False, is_rtl) for p in parts]
            add_table_row(current_table, cells, shade="F5F7FF")
            continue

        in_table = False; current_table = None
        size_ = 10.5
        color_ = None
        if in_parent_section:
            size_  = 11
            color_ = DARK_RGB
        add_para(ls, size=size_, space_before=0, space_after=3, color=color_)

    # Disclaimer
    disc = ("This report is confidential. Results reflect performance on this date only. "
            "Interpretation should be made in context with clinical observation and other data.") if lang=="en" \
           else ("هذا التقرير سري. النتائج تعكس الأداء في تاريخ التطبيق فقط. "
                 "يجب تفسير النتائج في سياق الملاحظة الإكلينيكية والبيانات الأخرى.")
    doc.add_paragraph().paragraph_format.space_after = Pt(12)
    add_para(disc, size=8, color=MID_BLUE_RGB, italic=True)

    buf = io.BytesIO()
    doc.save(buf); buf.seek(0)
    return buf

# ══════════════════════════════════════════════════════════════
#  EMAIL
# ══════════════════════════════════════════════════════════════
def send_email(data, buf_en, buf_ar, fn_en, fn_ar):
    name = data.get("name","—")
    fsiq = data.get("FSIQ")
    en_c, ar_c, _ = classify(fsiq) if fsiq else ("—","—","#888")

    msg = MIMEMultipart('mixed')
    msg['From']    = GMAIL_USER
    msg['To']      = RECIPIENT_EMAIL
    msg['Subject'] = f"[SB5 Report] {name} — {date.today().strftime('%B %d, %Y')}"

    body = f"""<html><body style="font-family:Georgia,serif;color:#1A1A2E;background:#F5F7FA;padding:20px;">
  <div style="max-width:560px;margin:0 auto;background:white;border:1px solid #C8922A;border-radius:8px;padding:28px;">
    <h2 style="font-weight:600;font-size:18px;color:#1B3A6B;margin-bottom:2px;">
      Stanford-Binet 5 — Assessment Report</h2>
    <p style="color:#888;font-size:11px;margin-top:0;">SB5 Psychological Assessment</p>
    <hr style="border:none;border-top:2px solid #C8922A;margin:16px 0;">
    <table style="width:100%;font-size:13px;border-collapse:collapse;">
      <tr><td style="padding:5px 0;color:#555;width:40%;">Examinee</td><td><strong>{name}</strong></td></tr>
      <tr><td style="padding:5px 0;color:#555;">Age</td><td>{data.get("age","—")}</td></tr>
      <tr><td style="padding:5px 0;color:#555;">FSIQ</td><td><strong style="color:#1B3A6B;">{fsiq or "—"}</strong></td></tr>
      <tr><td style="padding:5px 0;color:#555;">Classification</td><td>{en_c} / {ar_c}</td></tr>
      <tr><td style="padding:5px 0;color:#555;">Examiner</td><td>{data.get("examiner","—")}</td></tr>
    </table>
    <hr style="border:none;border-top:1px solid #DDE5F8;margin:16px 0;">
    <p style="font-size:12px;">English and Arabic reports are attached as Word documents.</p>
    <p style="font-size:10px;color:#888;font-style:italic;">Confidential — for the evaluating clinician only.</p>
  </div></body></html>"""

    msg.attach(MIMEText(body, 'html'))
    for buf_, fname_ in [(buf_en, fn_en), (buf_ar, fn_ar)]:
        buf_.seek(0)
        part = MIMEBase('application','vnd.openxmlformats-officedocument.wordprocessingml.document')
        part.set_payload(buf_.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename=fname_)
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
    layout="wide"
)

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;700&family=Inter:wght@300;400;500;600&family=Cairo:wght@400;600;700&display=swap');

html, body, [class*="css"] {{ font-family: 'Inter','Cairo',sans-serif; background:{LIGHT_BG}; }}
.stApp {{ background:{LIGHT_BG}; }}

.sb5-header {{
    background: linear-gradient(135deg, {DEEP_BLUE} 0%, {MID_BLUE} 60%, {DEEP_BLUE} 100%);
    border-radius: 16px; padding: 28px 36px; margin-bottom: 24px;
    box-shadow: 0 6px 24px rgba(27,58,107,0.3);
    border-bottom: 4px solid {GOLD};
}}
.sb5-header h1 {{
    color: white; font-family:'Playfair Display',serif;
    font-size: 1.8rem; font-weight: 700; margin: 0 0 6px 0;
}}
.sb5-header p {{ color: #A8BFDD; font-size: 0.85rem; margin:0; }}

.section-card {{
    background: white; border-radius: 12px; padding: 20px 24px;
    margin-bottom: 16px; box-shadow: 0 2px 12px rgba(27,58,107,0.08);
    border-left: 4px solid {GOLD};
}}
.section-card.blue {{ border-left-color: {DEEP_BLUE}; }}
.section-title {{
    font-size: 0.78rem; font-weight: 700; color: {DEEP_BLUE};
    text-transform: uppercase; letter-spacing: 0.1em; margin-bottom: 14px;
}}
.score-badge {{
    display: inline-block; padding: 4px 14px; border-radius: 20px;
    font-size: 0.8rem; font-weight: 600; color: white; margin: 2px;
}}
.field-label {{
    font-size: 12px; font-weight: 600; color: {DEEP_BLUE}; margin-bottom: 4px;
}}

div[data-testid="stTextInput"] input,
div[data-testid="stTextArea"] textarea {{
    background: white !important; border: 1.5px solid #C5D3F5 !important;
    border-radius: 8px !important; font-size: 13.5px !important;
}}
div[data-testid="stNumberInput"] input {{
    background: white !important; border: 1.5px solid #C5D3F5 !important;
    border-radius: 8px !important;
}}

.stButton > button {{
    background: {DEEP_BLUE} !important; color: white !important;
    border: none !important; border-radius: 10px !important;
    padding: 10px 26px !important; font-size: 14px !important;
    font-weight: 600 !important; transition: all 0.2s !important;
    box-shadow: 0 3px 12px rgba(27,58,107,0.3) !important;
}}
.stButton > button:hover {{ background: {MID_BLUE} !important; }}
.stButton > button[kind="primary"] {{
    background: linear-gradient(135deg, {DEEP_BLUE}, {MID_BLUE}) !important;
    font-size: 16px !important; padding: 14px 40px !important;
}}

div[data-testid="stDivider"] {{ margin: 16px 0; }}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
#  HEADER
# ══════════════════════════════════════════════════════════════
col_logo, col_hdr = st.columns([1, 5])
with col_logo:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=110)
with col_hdr:
    st.markdown("""
    <div class="sb5-header">
        <h1>🎓 Stanford-Binet Intelligence Scales, 5th Ed.</h1>
        <p>مقياس ستانفورد-بينيه للذكاء — الصورة الخامسة · Professional Assessment Report Generator</p>
    </div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
#  SESSION STATE
# ══════════════════════════════════════════════════════════════
if "report_done" not in st.session_state: st.session_state.report_done = False

# ══════════════════════════════════════════════════════════════
#  REPORT DISPLAY
# ══════════════════════════════════════════════════════════════
if st.session_state.report_done:
    data   = st.session_state["sb5_data"]
    rt_en  = st.session_state["report_en"]
    rt_ar  = st.session_state["report_ar"]
    charts = st.session_state["charts"]
    name   = data.get("name","—")
    fsiq   = data.get("FSIQ")
    en_c, ar_c, badge_color = classify(fsiq) if fsiq else ("—","—","#888")

    st.markdown(f"""
    <div style="background:linear-gradient(135deg,{DEEP_BLUE},{MID_BLUE});border-radius:12px;
                padding:18px 28px;margin-bottom:20px;color:white;font-size:15px;">
        ✅ Report Generated — <strong>{name}</strong>
        &nbsp;|&nbsp; FSIQ: <strong>{fsiq}</strong>
        &nbsp;|&nbsp; <span style="background:{badge_color};padding:2px 12px;border-radius:12px;">
        {en_c} / {ar_c}</span>
    </div>""", unsafe_allow_html=True)

    # Metric row
    mc = st.columns(len(IQ_SCORES) + len(FACTOR_SCORES))
    for i, k in enumerate(IQ_SCORES + FACTOR_SCORES):
        if data.get(k):
            with mc[i]:
                lbl = (IQ_LABELS if k in IQ_LABELS else FACTOR_LABELS)[k]["en"]
                pct = percentile_from_ss(data[k])
                st.metric(lbl, data[k], f"{pct}th %ile")

    # Charts
    st.subheader("📊 Score Visualizations")
    if charts.get("profile"):
        st.image(charts["profile"], use_container_width=True)
    
    c1, c2 = st.columns(2)
    with c1:
        if charts.get("gauge"):
            st.image(charts["gauge"], use_container_width=True)
    with c2:
        if charts.get("radar"):
            st.image(charts["radar"], use_container_width=True)

    if charts.get("subtest"):
        st.image(charts["subtest"], use_container_width=True)

    # Report tabs
    tab_en, tab_ar = st.tabs(["🇬🇧 English Report", "🇸🇦 Arabic Report"])
    with tab_en:
        st.text_area("", value=rt_en, height=500, label_visibility="collapsed")
    with tab_ar:
        st.text_area("", value=rt_ar, height=500, label_visibility="collapsed")

    # Downloads
    st.divider()
    fn_en = f"SB5_{name.replace(' ','_')}_EN.docx"
    fn_ar = f"SB5_{name.replace(' ','_')}_AR.docx"

    dl1, dl2, dl3, dl4 = st.columns(4)
    with dl1:
        buf_en = build_word_doc(rt_en, data, {k: v for k, v in charts.items()}, "en")
        st.download_button("📄 English Report (.docx)", data=buf_en, file_name=fn_en,
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           use_container_width=True)
    with dl2:
        buf_ar = build_word_doc(rt_ar, data, {k: v for k, v in charts.items()}, "ar")
        st.download_button("📄 التقرير العربي (.docx)", data=buf_ar, file_name=fn_ar,
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           use_container_width=True)
    with dl3:
        if st.button("📧 Send via Email", use_container_width=True):
            try:
                buf_en2 = build_word_doc(rt_en, data, charts, "en")
                buf_ar2 = build_word_doc(rt_ar, data, charts, "ar")
                send_email(data, buf_en2, buf_ar2, fn_en, fn_ar)
                st.success(f"✅ Sent to {RECIPIENT_EMAIL}")
            except Exception as e:
                st.error(f"Email error: {e}")
    with dl4:
        if st.button("↺ New Report", use_container_width=True):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.rerun()
    st.stop()

# ══════════════════════════════════════════════════════════════
#  DATA ENTRY FORM
# ══════════════════════════════════════════════════════════════
st.markdown('<div class="section-card blue"><div class="section-title">👤 Examinee Information / بيانات المفحوص</div>', unsafe_allow_html=True)
c1, c2, c3 = st.columns(3)
with c1:
    name_v    = st.text_input("Name / الاسم", placeholder="Full name / الاسم الكامل")
    dob_v     = st.text_input("Date of Birth / تاريخ الميلاد", placeholder="DD/MM/YYYY")
    age_v     = st.text_input("Age / العمر", placeholder="e.g. 14 years 6 months")
with c2:
    gender_v  = st.selectbox("Gender / النوع", ["Male / ذكر","Female / أنثى"])
    grade_v   = st.text_input("Grade / Occupation / الصف", placeholder="e.g. Grade 5")
    referral_v= st.text_input("Referral Source / جهة الإحالة", placeholder="School / Hospital / Clinic")
with c3:
    examiner_v= st.text_input("Examiner / الفاحص", placeholder="Clinician name")
    test_date_v= st.text_input("Test Date / تاريخ التطبيق", placeholder="DD/MM/YYYY")
    school_v  = st.text_input("School / Agency / المدرسة", placeholder="")
st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="section-card"><div class="section-title">🧠 IQ Scores / درجات الذكاء</div>', unsafe_allow_html=True)
st.caption("Standard scores: Mean=100, SD=15. Enter 90% Confidence Intervals as 'low–high' (e.g. 96–108)")
c1, c2, c3 = st.columns(3)
iq_vals = {}
with c1:
    st.markdown("**Full Scale IQ (FSIQ) / درجة الذكاء الكلية**")
    iq_vals["FSIQ"] = st.number_input("FSIQ", 40, 160, value=None, placeholder="e.g. 95", key="fsiq", label_visibility="collapsed")
    iq_vals["FSIQ_ci"] = st.text_input("CI FSIQ", placeholder="e.g. 91–99", key="fsiq_ci", label_visibility="collapsed")
with c2:
    st.markdown("**Nonverbal IQ (NVIQ) / غير لفظي**")
    iq_vals["NVIQ"] = st.number_input("NVIQ", 40, 160, value=None, placeholder="e.g. 92", key="nviq", label_visibility="collapsed")
    iq_vals["NVIQ_ci"] = st.text_input("CI NVIQ", placeholder="e.g. 87–97", key="nviq_ci", label_visibility="collapsed")
with c3:
    st.markdown("**Verbal IQ (VIQ) / لفظي**")
    iq_vals["VIQ"] = st.number_input("VIQ", 40, 160, value=None, placeholder="e.g. 98", key="viq", label_visibility="collapsed")
    iq_vals["VIQ_ci"] = st.text_input("CI VIQ", placeholder="e.g. 93–103", key="viq_ci", label_visibility="collapsed")
st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="section-card"><div class="section-title">📊 Factor Index Scores / مؤشرات العوامل</div>', unsafe_allow_html=True)
factor_vals = {}
cols_f = st.columns(5)
factor_info = [
    ("FR","Fluid Reasoning\nالاستدلال التحليلي"),
    ("KN","Knowledge\nالمعلومات"),
    ("QR","Quantitative Reasoning\nالاستدلال الكمي"),
    ("VS","Visual-Spatial\nالمعالجة البصرية"),
    ("WM","Working Memory\nالذاكرة العاملة"),
]
for i, (key, label) in enumerate(factor_info):
    with cols_f[i]:
        st.markdown(f"**{label}**")
        factor_vals[key]       = st.number_input(key, 40, 160, value=None, placeholder="SS", key=f"f_{key}", label_visibility="collapsed")
        factor_vals[f"{key}_ci"] = st.text_input(f"CI {key}", placeholder="CI", key=f"fci_{key}", label_visibility="collapsed")
st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="section-card"><div class="section-title">🔢 Subtest Scaled Scores / الدرجات المعيارية للاختبارات الفرعية</div>', unsafe_allow_html=True)
st.caption("Scaled scores: Mean=10, SD=3")
nv_vals = {}; v_vals = {}
sub_cols = st.columns(5)
for i, (key, label) in enumerate(factor_info):
    with sub_cols[i]:
        st.markdown(f"**{label.split(chr(10))[0]}**")
        nv_vals[key] = st.number_input(f"NV {key}", 1, 19, value=None, placeholder="NV", key=f"nv_{key}", label_visibility="collapsed")
        v_vals[key]  = st.number_input(f"V {key}",  1, 19, value=None, placeholder="V",  key=f"v_{key}",  label_visibility="collapsed")
st.caption("First row = Nonverbal · Second row = Verbal")
st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="section-card"><div class="section-title">📝 Clinical Information / المعلومات الإكلينيكية</div>', unsafe_allow_html=True)
c1, c2 = st.columns(2)
with c1:
    complaints_v   = st.text_area("Reason for Referral / سبب الإحالة", height=90, placeholder="Describe the referral question...")
    behavioral_v   = st.text_area("Behavioral Observations / الملاحظات السلوكية", height=110,
                                   placeholder="Describe behavior during testing: cooperation, attention, affect, language...")
with c2:
    background_v   = st.text_area("Background Information / المعلومات الأساسية", height=90,
                                   placeholder="Relevant developmental, educational, medical, family history...")
    recommendations_v = st.text_area("Additional Recommendations / توصيات إضافية", height=110,
                                      placeholder="Any specific recommendations to include...")
st.markdown('</div>', unsafe_allow_html=True)

# ── GENERATE ──
st.markdown("<br>", unsafe_allow_html=True)
gen_col, _ = st.columns([2, 3])
with gen_col:
    generate = st.button("✦ Generate Full Report / توليد التقرير الكامل",
                         type="primary", use_container_width=True)

if generate:
    if not iq_vals.get("FSIQ"):
        st.error("Please enter at least the Full Scale IQ (FSIQ) score.")
        st.stop()

    data = {
        "name":     name_v or "—",
        "dob":      dob_v,
        "age":      age_v,
        "gender":   gender_v,
        "grade":    grade_v,
        "referral": referral_v,
        "examiner": examiner_v,
        "test_date":test_date_v,
        "school":   school_v,
        "complaints": complaints_v,
        "behavioral_obs": behavioral_v,
        "background": background_v,
        "extra_recs": recommendations_v,
    }
    # Merge scores
    for k, v in iq_vals.items():
        data[k] = v
    for k, v in factor_vals.items():
        data[k] = v
    for k, v in nv_vals.items():
        data[f"nv_{k.lower()}"] = v
    for k, v in v_vals.items():
        data[f"v_{k.lower()}"] = v

    with st.spinner("⏳ Generating charts and reports — this may take 30–60 seconds..."):
        # Charts
        charts = {}
        profile_bytes = make_profile_chart(data, "en")
        charts["profile"] = profile_bytes

        fsiq_val = data.get("FSIQ")
        if fsiq_val:
            charts["gauge"] = make_classification_gauge(fsiq_val)

        nv_sub = {k: data.get(f"nv_{k.lower()}") for k in FACTOR_SCORES}
        v_sub  = {k: data.get(f"v_{k.lower()}")  for k in FACTOR_SCORES}
        factor_ss = {k: data.get(k) for k in FACTOR_SCORES}

        radar_bytes = make_factor_radar(factor_ss, "en")
        if radar_bytes:
            charts["radar"] = radar_bytes

        subtest_bytes = make_subtest_chart(nv_sub, v_sub)
        if subtest_bytes:
            charts["subtest"] = subtest_bytes

        # Reports
        report_en = generate_en_report(data)
        report_ar = generate_ar_report(data)

        st.session_state["sb5_data"]  = data
        st.session_state["report_en"] = report_en
        st.session_state["report_ar"] = report_ar
        st.session_state["charts"]    = charts
        st.session_state.report_done  = True
        st.rerun()
