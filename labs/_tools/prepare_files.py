# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl>=3.1"]
# ///
"""Generate the Al-Suha University file pack used by every lab (labs/files/).

    uv run labs/_tools/prepare_files.py

Deterministic: a fixed seed and fixed file times, so the same bytes on every run. Everything is
synthetic: the university, its people, its documents, its numbers. The answer keys go to
labs/solutions/reference/ for the checker (labs/solutions/check.py), which recomputes what it can
from the files.
"""
from __future__ import annotations

import csv
import io
import json
import random
import re
import sys
import zipfile
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, str(Path(__file__).resolve().parent))
import case_data as C  # noqa: E402

LABS = Path(__file__).resolve().parents[1]
OUT = LABS / "files"
REF = LABS / "solutions" / "reference"
SEED = 207
FIXED = datetime(2027, 7, 4, 9, 0, 0)
NOTE = "ملف تدريبي اصطناعي لمقرر أدلة الامتثال وتقييم نضج البيانات. كل الأسماء والوثائق والأرقام وهمية."
RESTRICTED = "تصنيف هذا الملف داخل الحالة: مقيد. لا يُلصق ولا يُرفع في أي أداة ذكاء اصطناعي توليدي."
PUBLIC = "تصنيف هذا الملف داخل الحالة: عام. يجوز إدخاله في المساعد المعتمد في مختبر 3.1."
HEAD_FILL = PatternFill("solid", fgColor="0F766E")
FILL_FILL = PatternFill("solid", fgColor="FFF7E0")
rng = random.Random(SEED)
W0, W1 = (date.fromisoformat(d) for d in C.WINDOW)
SPEC = {s[0]: s for s in C.SPECS}


def head(cls=RESTRICTED):
    return f"{C.ORG}\n{NOTE}\n{cls}\n"


# ------------------------------------------------------------------ writing helpers
def write_text(rel: str, text: str) -> None:
    path = OUT / rel
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes((text.strip() + "\n").encode("utf-8"))


def write_csv(rel: str, header, rows) -> None:
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\n")
    w.writerow(header)
    w.writerows(rows)
    (OUT / rel).write_bytes(buf.getvalue().encode("utf-8"))


def sheet(wb, title, header, rows, fill_cols=(), widths=None, validations=None):
    ws = wb.create_sheet(title)
    ws.sheet_view.rightToLeft = True
    ws.append(header)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HEAD_FILL
        c.alignment = Alignment(wrap_text=True, vertical="top")
    for r in rows:
        ws.append(list(r))
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
            if c.column - 1 in fill_cols:
                c.fill = FILL_FILL
    for i, w in enumerate(widths or [18] * len(header)):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = w
    ws.freeze_panes = "A2"
    for col, values in (validations or {}).items():
        dv = DataValidation(type="list", formula1='"' + ",".join(values) + '"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"{col}2:{col}{max(2, len(rows) + 1)}")
    return ws


def note_sheet(wb, lines, cls=RESTRICTED):
    ws = wb.create_sheet("اقرأ أولًا")
    ws.sheet_view.rightToLeft = True
    for line in [C.ORG, NOTE, cls, ""] + lines:
        ws.append([line])
    ws.column_dimensions["A"].width = 110
    return ws


def save(wb, rel: str) -> None:
    del wb[wb.sheetnames[0]]
    wb.properties.creator = C.ORG
    wb.properties.lastModifiedBy = C.ORG
    wb.properties.created = wb.properties.modified = FIXED
    path = OUT / rel
    tmp = path.with_suffix(".tmp")
    wb.save(tmp)
    with zipfile.ZipFile(tmp) as z:
        items = [(i.filename, z.read(i.filename)) for i in z.infolist()]
    stamp = FIXED.strftime("%Y-%m-%dT%H:%M:%SZ").encode()
    items = [(n, re.sub(rb"(<dcterms:(?:created|modified)[^>]*>)[^<]*(</dcterms)", rb"\g<1>" + stamp + rb"\g<2>", d))
             if n == "docProps/core.xml" else (n, d) for n, d in items]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in items:
            info = zipfile.ZipInfo(name, date_time=FIXED.timetuple()[:6])
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data)
    tmp.unlink()


def put_ref(name: str, data) -> None:
    REF.mkdir(parents=True, exist_ok=True)
    (REF / name).write_bytes((json.dumps(data, ensure_ascii=False, indent=1) + "\n").encode("utf-8"))


# ------------------------------------------------------------------ the case brief
CASE_BRIEF = f"""
{head()}
موجز الحالة (للعمل التحضيري)

من نحن
جامعة السها جامعة حكومية افتراضية فيها 9 كليات ونحو 21000 طالب و2600 موظف. تدير بيانات الطلاب في نظام القبول
والتسجيل، وبيانات الموظفين في نظام الموارد البشرية، والميزانية في النظام المالي، وتنشر مجموعات بيانات مفتوحة في
صفحة على موقعها.

لماذا الآن
الجامعة تستعد لتقرير الامتثال السنوي ضمن التقييم الوطني للبيانات. فترة الأدلة لهذه الدورة من {C.WINDOW[0]} إلى
{C.WINDOW[1]}، والتقديم بعد 14 أسبوعًا من بداية العمل على الملف (الأسبوع 0). في الدورة الماضية سُجلت مواصفات ممتثلة
بلا دليل، ونُقلت إلى الدورة الحالية ملاحظات لم تُغلق. ولذلك كلّف مكتب إدارة البيانات أخصائي امتثال (أنت) بإعداد حزمة
أدلة جاهزة للتدقيق لثلاثة مجالات: حوكمة البيانات، وجودة البيانات، والبيانات المفتوحة.

الأشخاص
- د. وداد الحسيّن: الرئيسة التنفيذية للبيانات ورئيسة مكتب إدارة البيانات.
- عبدالإله البكيري: مسؤول الامتثال، ومديرك المباشر.
- د. أسماء العواد: مديرة المراجعة الداخلية. تؤدي دور المقيِّم في التدقيق التجريبي.
- د. ثامر الربدي: عميد القبول والتسجيل، ومالك بيانات الطلاب.
- شذى المنيع: مشرفة بيانات الأعمال لبيانات الطلاب.
- قتيبة الصقعبي: مشرف البيانات الفني، ويدير أداة جودة البيانات والمستخرجات.
- دانة السكيت: مسؤولة البيانات المفتوحة وإتاحة المعلومات.

ثلاث قواعد للعمل
1. الحزمة تثبت ما حدث، ولا تصنعه. وثيقة تُكتب الآن لتصف نشاطًا مضى لا تدخل الحزمة.
2. عدم الامتثال يُسجل كما هو. المواصفة التي طُبقت جزئيًا تُسجل 0%، وتدخل سجل الملاحظات.
3. ملفات الحالة مصنفة «مقيد» إلا ما كُتب عليه «عام». لا يدخل أي أداة ذكاء اصطناعي توليدي إلا ما صُنف «عام»، أو ما
   كتبته أنت بكلماتك دون محتوى الملفات. هذه قاعدة مبادئ الذكاء الاصطناعي التوليدي للجهات الحكومية.
"""

METHOD_BRIEF = f"""
{head()}
منهجية التقييم كما يستعملها المقرر

أولًا: الأجزاء الثلاثة
يقيس التقييم الوطني للبيانات ثلاثة أجزاء: الامتثال، والنضج، والتميز التشغيلي. لم نجد نصًا رسميًا منشورًا لمنهجية
المؤشر الوطني للبيانات، فهذا الملف يجمع ما نُشر رسميًا عن جزء الامتثال، وأداتين خاصتين بالمقرر للجزأين الآخرين.

ثانيًا: الامتثال (منشور رسميًا في ضوابط ومواصفات إدارة البيانات الوطنية، القسم 4، بصياغتنا المختصرة)
1. تقيس الجهة الامتثال على مستوى كل مواصفة بقيمة ثنائية: 100% للمواصفة المطبقة كاملة، و0% للمطبقة جزئيًا أو غير
   المطبقة.
2. تُرفع الدرجة من المواصفة إلى الضابط، ثم إلى المجال، ثم إلى الجهة.
3. يُدعم التقرير السنوي بأدلة على تطبيق كل مواصفة، ويُقدَّم في الربع الثالث من السنة.
4. يقود العمل الرئيس التنفيذي للبيانات، ويجوز لمكتب إدارة البيانات الوطنية أن يجري تدقيقًا لاحقًا على جهات مختارة.
النص الرسمي: https://sdaia.gov.sa/ar/SDAIA/about/Documents/DataManagementPersonalDataProtectionStandardsAr.pdf

افتراض المقرر في الرفع (لم يحدده النص): درجة الضابط متوسط مواصفاته في المستخرج، ودرجة المجال متوسط ضوابطه، ودرجة الجهة
متوسط مجالاتها.

ثالثًا: النضج (أداة المقرر، لا أسئلة المؤشر الرسمية)
استبانة لكل مجال، ولكل سؤال مستوى من 0 إلى 5. المستويات تراكمية: لا يُعطى مستوى إلا إذا ثبت كل ما تحته بدليل كاف.
درجة المجال متوسط أسئلته. التفاصيل في questionnaire.xlsx.

رابعًا: التميز التشغيلي (أداة المقرر)
مقاييس تُحسب من مستخرجات أنظمة الجامعة بتعريف مكتوب: البسط، والمقام، والفترة، والاستثناءات. التفاصيل في
metric_definitions.txt.
"""


# ------------------------------------------------------------------ lab 1.1: the sample result and the map
def cascade(results: dict[str, int]):
    """spec -> control mean -> domain mean -> entity mean, the course's assumption."""
    controls: dict[tuple, list] = {}
    for code, val in results.items():
        s = SPEC[code]
        controls.setdefault((s[2], s[3]), []).append(val)
    cmean = {k: sum(v) / len(v) for k, v in controls.items()}
    domains: dict[str, list] = {}
    for (dom, _), v in cmean.items():
        domains.setdefault(dom, []).append(v)
    dmean = {k: sum(v) / len(v) for k, v in domains.items()}
    entity = sum(dmean.values()) / len(dmean)
    return cmean, dmean, entity


def lab_1_1():
    write_text("method_brief.txt", METHOD_BRIEF)
    wb = openpyxl.Workbook()
    note_sheet(wb, [
        "نتيجة الدورة الماضية كما سُجلت في ملف الجامعة، للمواصفات العشرين في مجالات الحالة الثلاثة.",
        "ورقة «الامتثال»: النتيجة المسجلة لكل مواصفة ورمز الدليل الذي أُحيل إليه.",
        "ورقة «النضج» و«المقاييس»: ما سُجل للجزأين الآخرين.",
        "ورقة «الخريطة»: املأ الخانات الملونة: ما يُقاس في كل جزء، وكيف تُحسب الدرجة، وأنواع الأدلة، ومن يقدمها.",
    ])
    rows = [(code, s[1], s[2], s[3], s[4], C.SAMPLE[code][0], C.SAMPLE[code][1], "") for code, s in SPEC.items()]
    sheet(wb, "الامتثال", ["الرمز", "المواصفة", "المجال", "الضابط", "الأولوية", "النتيجة المسجلة %", "مرجع الدليل", "ملاحظتك"],
          rows, fill_cols=(7,), widths=[8, 10, 16, 24, 9, 14, 14, 36])
    sheet(wb, "النضج", ["المجال", "الدرجة المسجلة", "ملاحظة"], C.SAMPLE_MATURITY, widths=[18, 14, 44])
    sheet(wb, "المقاييس", ["المقياس", "القيمة المسجلة", "المصدر"], C.SAMPLE_METRICS, widths=[44, 14, 30])
    sheet(wb, "الخريطة", ["الجزء", "ما يُقاس", "كيف تُحسب الدرجة", "أنواع الأدلة", "من يقدمها"],
          [(p, "", "", "", "") for p in ("الامتثال", "النضج", "التميز التشغيلي")], fill_cols=(1, 2, 3, 4),
          widths=[16, 30, 36, 36, 26])
    save(wb, "sample_result.xlsx")
    recorded = {c: v for c, (v, _) in C.SAMPLE.items()}
    valid = {c: (100 if v == 100 and ref else 0) for c, (v, ref) in C.SAMPLE.items()}
    _, d_rec, e_rec = cascade({c: (v if v in (0, 100) else v) for c, v in recorded.items()})
    _, d_val, e_val = cascade(valid)
    put_ref("lab_1_1.json", {
        "no_evidence": sorted(c for c, (v, ref) in C.SAMPLE.items() if v == 100 and not ref),
        "not_binary": sorted(c for c, (v, _) in C.SAMPLE.items() if v not in (0, 100)),
        "domain_recorded": {k: round(v, 1) for k, v in d_rec.items()},
        "entity_recorded": round(e_rec, 1),
        "domain_corrected": {k: round(v, 1) for k, v in d_val.items()},
        "entity_corrected": round(e_val, 1),
        "map": C.MAP_REFERENCE,
    })


# ------------------------------------------------------------------ lab 1.2: instruments
def lab_1_2():
    wb = openpyxl.Workbook()
    note_sheet(wb, [
        "ستة عشر صكًا والتزامًا مرشحًا لملف الامتثال في هذه الدورة.",
        "لكل صف: اختر الحكم من القائمة، واكتب السبب في سطر.",
        "الأحكام: ينطبق، ينطبق ويُقيَّم عبر جهة أخرى، غير ملزم، لا ينطبق.",
        "الحالة في الجامعة: لا مستشفى جامعي؛ سياسة إدارة البيانات 1.2 ألغت 1.0؛ دليل الجودة مسودة 0.4.",
    ])
    sheet(wb, "الصكوك", ["الرمز", "الصك أو الالتزام", "النوع", "الحكم", "السبب"],
          [(i, n, k, "", "") for i, n, k, _, _ in C.INSTRUMENTS], fill_cols=(3, 4), widths=[7, 60, 16, 24, 44],
          validations={"D": ["ينطبق", "ينطبق ويُقيَّم عبر جهة أخرى", "غير ملزم", "لا ينطبق"]})
    save(wb, "instruments.xlsx")
    put_ref("lab_1_2.json", {i: {"key": k, "reason": r} for i, _, _, k, r in C.INSTRUMENTS})


# ------------------------------------------------------------------ labs 1.3 and 1.4: criteria
def specs_text():
    lines = [head(), "عشرون مواصفة من ثلاثة مجالات: صياغة مختصرة للتدريب. النص الرسمي في:",
             "https://sdaia.gov.sa/ar/SDAIA/about/Documents/DataManagementPersonalDataProtectionStandardsAr.pdf", ""]
    for code, sid, dom, ctl, pr, text, _ in C.SPECS:
        lines.append(f"{code} · {sid} · {dom} · {ctl} · {pr}")
        lines.append(text)
        lines.append("")
    return "\n".join(lines)


def lab_1_3_4():
    write_text("spec_extract.txt", specs_text())
    wb = openpyxl.Workbook()
    note_sheet(wb, [
        "ورقة «المعايير»: لكل مواصفة اكتب معيار قبول في خمس خانات: ما يجب أن يوجد، ومن يعتمده، وما يثبت أنه مطبق، وأنواع الأدلة، والنافذة الزمنية.",
        "المختبر 1.3: المواصفات S01 إلى S10. المختبر 1.4: المواصفات S11 إلى S20، ثم ورقة «مسودة زميل».",
        "ورقة «جدول الاعتماد»: من يعتمد ماذا في الجامعة. ورقة «العيوب»: رموز العيوب الستة.",
        f"فترة الأدلة لهذه الدورة: من {C.WINDOW[0]} إلى {C.WINDOW[1]}.",
    ])
    rows = [(code, s[1], s[2], s[4], s[5], "", "", "", "", "") for code, s in SPEC.items()]
    sheet(wb, "المعايير", ["الرمز", "المواصفة", "المجال", "الأولوية", "نص المواصفة (مختصر)", "ما يجب أن يوجد", "من يعتمده",
                           "ما يثبت أنه مطبق", "أنواع الأدلة", "النافذة الزمنية"],
          rows, fill_cols=(5, 6, 7, 8, 9), widths=[7, 9, 14, 8, 50, 30, 20, 30, 24, 20])
    sheet(wb, "جدول الاعتماد", ["الوثيقة", "يعتمدها"], C.APPROVAL_MATRIX, widths=[50, 40])
    sheet(wb, "مسودة زميل", ["الرمز", "ما يجب أن يوجد", "من يعتمده", "ما يثبت أنه مطبق", "رموز العيوب", "تصحيحك"],
          [(c, a, b, d, "", "") for c, a, b, d, _ in C.COLLEAGUE_DRAFT], fill_cols=(4, 5), widths=[7, 26, 22, 30, 16, 40])
    sheet(wb, "العيوب", ["الرمز", "العيب", "علامته"], C.DEFECTS, widths=[7, 26, 60])
    save(wb, "criteria_template.xlsx")
    put_ref("criteria.json", {c: {"exists": v[0], "approver": v[1], "operates": v[2], "types": v[3], "window": v[4],
                                  "layers": SPEC[c][6]} for c, v in C.CRITERIA.items()})
    put_ref("colleague_draft.json", {c: k for c, _, _, _, k in C.COLLEAGUE_DRAFT})


# ------------------------------------------------------------------ labs 2.1 and 2.2: sixty items
RULES = f"""
{head()}
قواعد الحكم على كفاية الدليل (قواعد الجامعة لهذه الدورة)

فترة الأدلة: من {C.WINDOW[0]} إلى {C.WINDOW[1]}.
يُحكم على كل بند مقابل المواصفة والطبقة اللتين قُدِّم لهما فقط: الوجود، أو الاعتماد، أو التطبيق، أو الفاعلية.
وثائق الإنشاء والاعتماد: استراتيجية، خطة، سياسة، إجراء، قرار، ميثاق، تعيين، وصف وظيفي، سجل أُنشئ، تعريف مؤشر.
وثائق التشغيل: محضر، مستخرج سجل، تقرير دوري، لقطة شاشة، قائمة حضور، بريد، عينة، تقييم.

طبّق القواعد بهذا الترتيب. أول قاعدة تنطبق تحسم الحكم ورمز السبب.
1. R6 غير كاف: البند لا يتناول المواصفة، أو لا يثبت الطبقة التي قُدِّم لها (دليل إرشادي قُدِّم دليلًا على التطبيق مثلًا).
2. R5 غير كاف: أُعد بعد نهاية الفترة ليصف نشاطًا داخلها أو جوابًا عن سؤال الامتثال، أو أُرسل لأول مرة أثناء إعداد الملف. ولا يدخل في ذلك تقرير دوري عن ربع داخل الفترة صدر في موعده المعتاد بعد نهاية ربعه.
3. R8 غير كاف: إصدار استُبدل بإصدار أحدث.
4. R1 غير كاف: وثيقة تشغيل تاريخها قبل بداية الفترة. (وثيقة الإنشاء أو الاعتماد السارية تبقى صالحة مهما كان تاريخها.)
5. R2 غير كاف: في أي طبقة: مسودة، أو محضر بلا توقيع رئيس الاجتماع. وفي طبقة الاعتماد: غير موقع، أو موقع من غير الدور المحدد في جدول الاعتماد.
6. R7 غير كاف: لقطة شاشة بلا تاريخ أو بلا اسم النظام.
7. R3 جزئي: يغطي جزءًا من النطاق (كلية، نظام من ثلاثة، موظف) أو جزءًا مما تطلبه الطبقة.
8. R4 جزئي: دليل تشغيل يغطي جزءًا من الفترة (اجتماعان من أربعة، خمسة أشهر من اثني عشر).
9. إن لم تنطبق قاعدة: كاف.


اكتب لكل بند حكمًا ورمز السبب، وفي «الفجوة» ما يلزم ليصبح الدليل كافيًا أو ما يُستبعد من الحزمة.
مواصفة امتثالها 100% فقط إذا كان لكل طبقة تطلبها معاييرُ القبول بندٌ كاف واحد على الأقل.
"""


def lab_2_1_2():
    write_text("sufficiency_rules.txt", RULES)
    wb = openpyxl.Workbook()
    note_sheet(wb, [
        "ستون بندًا من مستودع أدلة الجامعة للمجالات الثلاثة.",
        "المختبر 2.1: البنود E01 إلى E30، دون أي مساعد ذكاء اصطناعي. المختبر 2.2: البنود E31 إلى E60.",
        "لكل بند: الحكم (كاف، جزئي، غير كاف)، ورمز السبب من sufficiency_rules.txt، والفجوة.",
        "الحقل الفارغ في «موقع من» يعني أن البند لا يحمل توقيعًا. الحقل الفارغ في «التاريخ» يعني أن البند بلا تاريخ.",
    ])
    rows = [(i, SPEC[s][1], layer, typ, title, d, signed, scope, desc, "", "", "")
            for i, s, layer, typ, title, d, signed, scope, desc, *_ in C.ITEMS]
    sheet(wb, "البنود", ["البند", "المواصفة", "الطبقة", "النوع", "العنوان", "التاريخ", "موقع من", "النطاق", "الوصف",
                         "الحكم", "رمز السبب", "الفجوة"],
          rows, fill_cols=(9, 10, 11), widths=[6, 9, 9, 11, 40, 11, 18, 16, 50, 10, 9, 30],
          validations={"J": ["كاف", "جزئي", "غير كاف"], "K": ["", "R1", "R2", "R3", "R4", "R5", "R6", "R7", "R8"]})
    sheet(wb, "جدول الاعتماد", ["الوثيقة", "يعتمدها"], C.APPROVAL_MATRIX, widths=[50, 40])
    sheet(wb, "رموز الأسباب", ["الرمز", "المعنى"], [(k, v) for k, v in C.R.items() if k], widths=[7, 70])
    save(wb, "evidence_register.xlsx")
    key = {i: {"spec": s, "layer": layer, "key": k, "reason": r, "gap": g}
           for i, s, layer, _, _, _, _, _, _, k, r, g in C.ITEMS}
    put_ref("evidence_key.json", key)


# ------------------------------------------------------------------ lab 2.3: questionnaire
def supported_level(cited, items):
    have = {}
    for m in cited:
        lvl, status = items[m]
        if status == "كاف":
            have.setdefault(lvl, True)
    level = 0
    while have.get(level + 1):
        level += 1
    return level


def lab_2_3():
    items = {m: (lvl, st) for m, lvl, _, st, _ in C.M_ITEMS}
    wb = openpyxl.Workbook()
    note_sheet(wb, [
        "استبانة النضج للمجالات الثلاثة، كما أجاب عنها ملاك المجالات، مع الأدلة التي أحالوا إليها.",
        "حالة كل دليل حكم عليها مسؤول الامتثال سلفًا (ورقة «الأدلة»).",
        "لكل سؤال: اكتب المستوى الذي تدعمه الأدلة، وأول مستوى ناقص، والفرق عن المستوى المدعى.",
        "المستويات تراكمية: لا يُعطى مستوى إلا إذا كان لكل مستوى تحته دليل كاف.",
        "مقياس المستويات أداة خاصة بالمقرر، وليس مستويات المؤشر الوطني.",
    ])
    sheet(wb, "المستويات", ["المستوى", "الاسم", "ما يثبته"], C.LEVELS, widths=[9, 12, 60])
    rows = [(q, d, t, lvl, "، ".join(cited), "", "", "") for q, d, t, lvl, cited in C.QUESTIONS]
    sheet(wb, "الأسئلة", ["السؤال", "المجال", "الموضوع", "المستوى المدعى", "الأدلة المحال إليها", "المستوى المدعوم",
                          "أول مستوى ناقص", "الفرق"], rows, fill_cols=(5, 6, 7), widths=[7, 16, 28, 12, 30, 14, 14, 9])
    sheet(wb, "الأدلة", ["الدليل", "يُقدَّم للمستوى", "الوصف", "الحالة", "السبب"], C.M_ITEMS, widths=[7, 12, 50, 10, 24])
    save(wb, "questionnaire.xlsx")
    ref = {}
    for q, d, t, claimed, cited in C.QUESTIONS:
        sup = supported_level(cited, items)
        ref[q] = {"domain": d, "claimed": claimed, "supported": sup, "first_missing": sup + 1 if sup < 5 else None}
    doms = {}
    for q, r in ref.items():
        doms.setdefault(r["domain"], []).append(r)
    put_ref("lab_2_3.json", {
        "questions": ref,
        "domain_claimed": {k: sum(r["claimed"] for r in v) / len(v) for k, v in doms.items()},
        "domain_supported": {k: sum(r["supported"] for r in v) / len(v) for k, v in doms.items()},
    })


# ------------------------------------------------------------------ lab 2.4: metrics
DEFINITIONS = f"""
{head()}
تعريفات المقاييس التشغيلية الأربعة (أداة المقرر)

نهاية الفترة: {C.WINDOW[1]}. بداية الفترة: {C.WINDOW[0]}.

M1 تغطية الكتالوج للأصول الحرجة (catalogue_extract.csv)
البسط: الأصول الحرجة النشطة التي حقولها الأربعة غير فارغة: المالك، والمشرف، والوصف، والتصنيف.
المقام: الأصول الحرجة النشطة.
الاستثناءات: الأصل المكرر يُحسب مرة واحدة بأحدث صف له (أحدث updated). الأصل الملغى (retired) خارج المقام.
الأصل الذي يبدأ اسمه بـ TEST_ أصل تجريبي خارج المقام.

M2 نسبة نجاح قواعد الجودة على بيانات الأولوية الأولى (quality_runs.csv)
البسط: القواعد التي نجح آخر تشغيل لها داخل الفترة.
المقام: القواعد المفعلة (enabled) على بيانات الأولوية 1 التي لها تشغيل داخل الفترة.
الاستثناءات: القاعدة المعطلة خارج المقام ولو كان لها تشغيل. التشغيل بعد نهاية الفترة لا يُحسب.

M3 وسيط مدة البت في طلبات مشاركة البيانات (sharing_requests.csv)
القيمة: وسيط الأيام بين received وdecided للطلبات المقبولة والمرفوضة التي decided لها داخل الفترة.
الاستثناءات: الطلب المسحوب (withdrawn) خارج الحساب. ويُذكر معه عدد الطلبات المفتوحة (open) التي مضى على
استلامها أكثر من 30 يومًا في نهاية الفترة.

M4 نسبة المجموعات المفتوحة المحدثة في موعدها (opendata_register.csv)
البسط: المجموعات المنشورة التي آخر تحديث لها في نهاية الفترة لا يتجاوز دوريتها: شهري 31 يومًا، ربعي 92 يومًا،
سنوي 366 يومًا.
المقام: المجموعات المنشورة (published). المحذوفة (removed) والمسودة (draft) خارج المقام.
"""


def d(x: date) -> str:
    return x.isoformat()


def lab_2_4():
    write_text("metric_definitions.txt", DEFINITIONS)
    # M1 catalogue: 70 critical assets (6 retired, 3 test), 40 non-critical
    cat = []
    n = 0
    base = date(2026, 9, 1)
    complete_targets = set(rng.sample(range(61), 46))  # of the 61 critical active non-test assets
    for i in range(70):
        n += 1
        aid = f"A{n:03d}"
        name = f"TEST_asset_{i}" if i in (10, 33, 57) else f"asset_{n:03d}"
        status = "retired" if i in (5, 21, 40, 52, 63, 68) else "active"
        eligible_idx = len([1 for r in cat if r[2] == "yes" and r[3] == "active" and not r[1].startswith("TEST_")
                            and r[0] not in {x[0] for x in cat[:0]}])
        full = (eligible_idx in complete_targets) if (status == "active" and not name.startswith("TEST_")) else rng.random() < 0.5
        owner = "مالك البيانات" if full or rng.random() < 0.5 else ""
        steward = "مشرف البيانات" if full or rng.random() < 0.3 else ""
        desc = "موصوف" if full else ""
        cls = rng.choice(["عام", "مقيد"]) if full else (rng.choice(["", "مقيد"]))
        cat.append([aid, name, "yes", status, owner, steward, desc, cls, d(base + timedelta(days=rng.randint(0, 200)))])
    for i in range(40):
        n += 1
        cat.append([f"A{n:03d}", f"asset_{n:03d}", "no", "active", "مالك البيانات", "", rng.choice(["موصوف", ""]),
                    "مقيد", d(base + timedelta(days=rng.randint(0, 200)))])
    # duplicates: older incomplete rows for 8 eligible complete assets, and a newer incomplete row for 2
    eligible = [r for r in cat if r[2] == "yes" and r[3] == "active" and not r[1].startswith("TEST_")]
    complete = [r for r in eligible if all(r[4:8])]
    incomplete = [r for r in eligible if not all(r[4:8])]
    dups = []
    for r in complete[:8]:
        old = list(r)
        old[6] = ""
        old[8] = d(date.fromisoformat(r[8]) - timedelta(days=60))
        dups.append(old)
    for r in incomplete[:2]:
        newer = list(r)
        newer[4:8] = ["مالك البيانات", "مشرف البيانات", "موصوف", "مقيد"]
        newer[8] = d(date.fromisoformat(r[8]) + timedelta(days=30))
        dups.append(newer)
    rows = cat + dups
    rng.shuffle(rows)
    write_csv("catalogue_extract.csv", ["asset_id", "asset_name", "critical", "status", "owner", "steward", "description",
                                         "classification", "updated"], rows)
    latest = {}
    for r in rows:
        if r[0] not in latest or r[8] > latest[r[0]][8]:
            latest[r[0]] = r
    elig = [r for r in latest.values() if r[2] == "yes" and r[3] == "active" and not r[1].startswith("TEST_")]
    m1 = sum(all(r[4:8]) for r in elig) / len(elig)
    naive = sum(all(r[4:8]) for r in rows if r[2] == "yes") / sum(1 for r in rows if r[2] == "yes")

    # M2 quality runs
    runs = []
    for k in range(1, 61):
        rid = f"R{k:02d}"
        prio = 1 if k <= 40 else 2
        status = "disabled" if k in (7, 18, 29, 44) else "enabled"
        dates = sorted(rng.sample(range(0, 360), 3))
        for j, off in enumerate(dates):
            runs.append([rid, prio, status, d(W0 + timedelta(days=off)), "pass" if rng.random() < 0.82 else "fail"])
        if k in (3, 12, 25, 31):  # a run after the window that must not count
            runs.append([rid, prio, status, d(W1 + timedelta(days=5)), "fail" if k != 25 else "pass"])
    runs.sort(key=lambda r: (r[3], r[0]))
    write_csv("quality_runs.csv", ["rule_id", "data_priority", "status", "run_date", "result"], runs)
    last = {}
    for r in runs:
        if r[1] == 1 and r[2] == "enabled" and W0 <= date.fromisoformat(r[3]) <= W1:
            if r[0] not in last or r[3] > last[r[0]][3]:
                last[r[0]] = r
    m2 = sum(r[4] == "pass" for r in last.values()) / len(last)

    # M3 sharing requests
    req = []
    for k in range(1, 49):
        rec = W0 + timedelta(days=rng.randint(-20, 340))
        kind = rng.choices(["approved", "rejected", "withdrawn", "open"], [30, 8, 4, 6])[0]
        if kind == "open":
            req.append([f"SR{k:02d}", "جهة حكومية", d(rec), "", "open"])
            continue
        dec = rec + timedelta(days=rng.choice([4, 6, 9, 12, 14, 17, 21, 25, 33, 41, 58]))
        req.append([f"SR{k:02d}", "جهة حكومية", d(rec), d(dec), kind])
    write_csv("sharing_requests.csv", ["request_id", "requester", "received", "decided", "status"], req)
    durs = sorted((date.fromisoformat(r[3]) - date.fromisoformat(r[2])).days for r in req
                  if r[4] in ("approved", "rejected") and W0 <= date.fromisoformat(r[3]) <= W1)
    mid = len(durs) // 2
    m3 = durs[mid] if len(durs) % 2 else (durs[mid - 1] + durs[mid]) / 2
    open_old = sum(1 for r in req if r[4] == "open" and (W1 - date.fromisoformat(r[2])).days > 30)

    # M4 open data register: 9 published (5 on time), 2 removed, 1 draft
    od = [
        ("OD01", "أعداد الطلاب حسب الكلية", "published", "quarterly", 40),
        ("OD02", "أعداد الخريجين حسب التخصص", "published", "annual", 200),
        ("OD03", "الطاقة الاستيعابية للسكن الجامعي", "published", "annual", 410),
        ("OD04", "أعداد أعضاء هيئة التدريس", "published", "annual", 120),
        ("OD05", "البرامج الأكاديمية المعتمدة", "published", "quarterly", 150),
        ("OD06", "مواعيد التقويم الأكاديمي", "published", "quarterly", 60),
        ("OD07", "إحصاءات المكتبة الشهرية", "published", "monthly", 75),
        ("OD08", "المنح البحثية المنشورة", "published", "annual", 380),
        ("OD09", "زيارات الموقع الإلكتروني", "published", "monthly", 20),
        ("OD10", "أعداد المقبولين في الفصل الصيفي", "removed", "annual", 500),
        ("OD11", "أسعار خدمات الإسكان", "removed", "annual", 450),
        ("OD12", "رضا الطلاب عن الخدمات", "draft", "annual", 0),
    ]
    odrows = [[i, t, s, f, d(W1 - timedelta(days=age)) if s != "draft" else ""] for i, t, s, f, age in od]
    write_csv("opendata_register.csv", ["dataset_id", "title", "status", "frequency", "last_update"], odrows)
    lim = {"monthly": 31, "quarterly": 92, "annual": 366}
    pub = [r for r in odrows if r[2] == "published"]
    on_time = [r[0] for r in pub if (W1 - date.fromisoformat(r[4])).days <= lim[r[3]]]
    m4 = len(on_time) / len(pub)

    wb = openpyxl.Workbook()
    note_sheet(wb, ["احسب المقاييس الأربعة من المستخرجات بتعريفاتها في metric_definitions.txt، واكتب البسط والمقام والقيمة،"
                    " وما استبعدته ولماذا."])
    sheet(wb, "المقاييس", ["المقياس", "البسط", "المقام", "القيمة", "ما استُبعد ولماذا"],
          [("M1 تغطية الكتالوج", "", "", "", ""), ("M2 نجاح قواعد الأولوية الأولى", "", "", "", ""),
           ("M3 وسيط مدة البت (أيام)", "", "", "", ""), ("M4 المجموعات المحدثة في موعدها", "", "", "", "")],
          fill_cols=(1, 2, 3, 4), widths=[30, 10, 10, 10, 50])
    save(wb, "metrics_sheet.xlsx")
    put_ref("lab_2_4.json", {
        "M1": {"num": sum(all(r[4:8]) for r in elig), "den": len(elig), "value": round(m1 * 100, 1),
               "naive_all_rows": round(naive * 100, 1), "rows": len(rows), "duplicates": len(dups)},
        "M2": {"num": sum(r[4] == "pass" for r in last.values()), "den": len(last), "value": round(m2 * 100, 1)},
        "M3": {"n": len(durs), "median_days": m3, "open_over_30": open_old},
        "M4": {"num": len(on_time), "den": len(pub), "value": round(m4 * 100, 1),
               "late": sorted(r[0] for r in pub if r[0] not in on_time)},
    })


# ------------------------------------------------------------------ lab 3.1: the assistant's mapping
def lab_3_1():
    wb = openpyxl.Workbook()
    note_sheet(wb, ["أربعون وثيقة منشورة في موقع الجامعة، مصنفة «عام». هذه وحدها يجوز أن تدخل المساعد المعتمد.",
                    "العمود «المواصفة» فارغ: هو ما يربطه المساعد أو المقيِّم."], cls=PUBLIC)
    sheet(wb, "الوثائق", ["الوثيقة", "النوع", "العنوان", "التصنيف"],
          [(i, t, title, "عام") for i, t, title, *_ in C.PUBLIC_DOCS], widths=[8, 9, 60, 9])
    save(wb, "public_docs.xlsx")

    wb = openpyxl.Workbook()
    note_sheet(wb, [
        "سجل الدفعة الأولى من تجربة المساعد: ربط كل وثيقة بمواصفة، كما أخرجه المساعد. المخرجات في هذا الملف مكتوبة سلفًا"
        " لتطابق ما يقيسه المختبر، فلا يحتاج المختبر نموذجًا.",
        "ورقة «مخرجات المساعد»: المواصفة التي اقترحها المساعد وثقته وتصنيف الوثيقة كما أُدخلت.",
        "ورقة «مفتاح المقيم»: المواصفة الصحيحة لكل وثيقة، أعدتها د. أسماء العواد.",
        "ورقة «فهرس المواصفات»: كل رقم مواصفة في النص المنشور من المعيار (190 رقمًا).",
    ])
    out_rows = [(i, title, a, conf, "عام", "", "") for i, _, title, _, a, conf, _ in C.PUBLIC_DOCS]
    r = C.RESTRICTED_ROW
    out_rows.append((r[0], r[2], r[4], r[5], "مقيد", "", ""))
    sheet(wb, "مخرجات المساعد", ["الوثيقة", "العنوان", "مواصفة المساعد", "ثقة المساعد", "تصنيف المدخل", "صحيح؟",
                                  "نوع الخطأ"], out_rows, fill_cols=(5, 6), widths=[8, 50, 14, 11, 11, 9, 26],
          validations={"F": ["نعم", "لا", "مستبعد"], "G": C.ERROR_TYPES + ["خرق قاعدة التصنيف"]})
    sheet(wb, "مفتاح المقيم", ["الوثيقة", "النوع", "المواصفة الصحيحة"],
          [(i, t, k) for i, t, _, k, *_ in C.PUBLIC_DOCS] + [(r[0], r[1], r[3])], widths=[8, 9, 16])
    sheet(wb, "فهرس المواصفات", ["المواصفة", "المجال"], [(s, s.split(".")[0]) for s in C.SPEC_IDS], widths=[12, 9])
    sheet(wb, "أنواع الأخطاء", ["النوع"], [(t,) for t in C.ERROR_TYPES], widths=[40])
    save(wb, "mapping_pilot.xlsx")

    by_type = {}
    errors = {}
    for i, t, _, k, a, conf, e in C.PUBLIC_DOCS:
        by_type.setdefault(t, [0, 0])
        by_type[t][0] += 1
        if a != k:
            by_type[t][1] += 1
            errors[i] = {"type": e, "confidence": conf}
    put_ref("lab_3_1.json", {"n": len(C.PUBLIC_DOCS), "wrong": len(errors), "by_type": by_type, "errors": errors,
                             "excluded": {r[0]: "مصنف «مقيد»"}})


# ------------------------------------------------------------------ lab 3.2: usage rule
USAGE_TEMPLATE = f"""
{head()}
قالب قاعدة استخدام المساعد في عملية الامتثال (املأ كل بند في سطر أو سطرين)

1. النطاق: في أي خطوة من عملية الامتثال يُستعمل المساعد، وفي أيها لا يُستعمل.
2. المدخلات المسموحة: ما يجوز أن يدخل المساعد، وما لا يجوز، ومن يتحقق من التصنيف قبل الإدخال.
3. أنواع الوثائق: الأنواع التي يقترح لها المساعد ربطًا أوليًا، والأنواع التي يربطها إنسان من البداية، ولماذا.
4. التوقيع البشري: من يراجع كل ربط ويعتمده، وماذا يعني الاعتماد.
5. سجل التدقيق: الحقول التي تُسجل لكل ربط.
6. إعادة القياس: حجم العينة في كل دورة، والمفتاح الذي تُقاس عليه، ومن يعده.
7. شرط التوقف: متى يتوقف استعمال المساعد لنوع من الوثائق.
8. ما يُبلغ للجنة إدارة البيانات.
"""


def lab_3_2():
    write_text("usage_rule_template.txt", USAGE_TEMPLATE)
    wb = openpyxl.Workbook()
    note_sheet(wb, ["الدفعة الثانية من تجربة المساعد: ستون وثيقة «عام» أخرى، قاسها مكتب الامتثال على مفتاح أعدته المراجعة"
                    " الداخلية. لم تدخل الدفعة أي وثيقة «مقيد»."])
    sheet(wb, "الدفعة الثانية", ["النوع", "الوثائق", "الربط الخاطئ", "منها بثقة عالية"], C.BATCH2, widths=[12, 10, 14, 16])
    save(wb, "error_summary_batch2.xlsx")


# ------------------------------------------------------------------ lab 3.3: findings register
RANKING = f"""
{head()}
قاعدة ترتيب الملاحظات بحسب المخاطر (قاعدة الجامعة)

الأثر من أولوية المواصفة: P1 = 3، P2 = 2، P3 = 1.
الاحتمال من امتداد الفجوة: «كل الجامعة» = 3، «جزئي» = 2، «حالة واحدة» = 1.
درجة الخطر = الأثر × الاحتمال. الترتيب تنازلي بالدرجة؛ عند التساوي الأولوية الأعلى أولًا، ثم رقم الملاحظة الأصغر.
ملاحظتان لهما سبب جذري واحد تُدمجان في ملاحظة واحدة، وتأخذ الملاحظة المدمجة أعلى أثر وأعلى احتمال بين أصليها.
الممارسة السليمة لا تُسجل ملاحظة. والمخالفة لا تُحذف لأن مالكها يطلب ذلك؛ يُسجل طلبه وردك في عمود الملاحظات.
كل ملاحظة تُكتب بأربعة أجزاء: الحالة (ما وُجد)، والمعيار (المواصفة)، والسبب الجذري، والأثر.
"""


def lab_3_3():
    write_text("ranking_rule.txt", RANKING)
    wb = openpyxl.Workbook()
    note_sheet(wb, ["اثنتا عشرة مشاهدة من مراجعة مجالي جودة البيانات والبيانات المفتوحة.",
                    "اكتب السجل في ورقة «السجل»: ملاحظة لكل سبب جذري، بأجزائها الأربعة ودرجة خطرها وترتيبها."])
    sheet(wb, "المشاهدات", ["المشاهدة", "المواصفة", "الأولوية", "ما وُجد", "الدليل", "الامتداد", "ملاحظة"],
          C.OBSERVATIONS, widths=[9, 9, 8, 50, 10, 11, 40])
    sheet(wb, "السجل", ["الملاحظة", "المشاهدات", "المواصفة", "الحالة", "السبب الجذري", "الأثر", "الأثر (1 إلى 3)",
                        "الاحتمال (1 إلى 3)", "الدرجة", "الترتيب", "ملاحظات"],
          [("F%02d" % k, "", "", "", "", "", "", "", "", "", "") for k in range(1, 11)],
          fill_cols=tuple(range(1, 11)), widths=[9, 12, 12, 34, 30, 26, 9, 9, 8, 8, 30])
    save(wb, "observations.xlsx")
    obs = {o[0]: o for o in C.OBSERVATIONS}
    reg = []
    for fid, olist, specs, cause in C.FINDINGS_33:
        impact = max(C.IMPACT[obs[o][2]] for o in olist)
        like = max(C.EXTENT_LIKELIHOOD[obs[o][5]] for o in olist)
        prio = min(obs[o][2] for o in olist)
        reg.append({"id": fid, "obs": olist, "specs": specs, "cause": cause, "impact": impact, "likelihood": like,
                    "score": impact * like, "priority": prio, "first": min(olist)})
    reg.sort(key=lambda f: (-f["score"], f["priority"], f["first"]))
    for k, f in enumerate(reg, 1):
        f["rank"] = k
    put_ref("lab_3_3.json", {"register": reg, "not_findings": C.NOT_FINDINGS})


# ------------------------------------------------------------------ lab 3.4: corrective plan
PLAN_RULES = f"""
{head()}
قواعد الخطة التصحيحية وتقويم الجاهزية

الأسبوع 0 بداية العمل على الملف، والتقديم في الأسبوع {C.CALENDAR['submission_week']}.
تجتمع لجنة إدارة البيانات في الأسابيع {"، ".join(map(str, C.CALENDAR['committee_weeks']))}.
آخر موعد لدليل يدخل ملف هذه الدورة: الأسبوع {C.CALENDAR['evidence_by_week']}. ويتحقق منه طرف غير المالك قبل الأسبوع
{C.CALENDAR['verify_by_week']}.

1. لكل ملاحظة إجراء واحد على الأقل يعالج سببها الجذري، لا عرضها.
2. المالك دور مسمى مسؤول عن الإجراء، لا «مكتب إدارة البيانات» ولا «الجميع».
3. المتحقق غير المالك: مسؤول الامتثال، أو المراجعة الداخلية حين يكون مسؤول الامتثال هو المالك.
4. دليل الإغلاق سجل يثبت أن الإجراء نُفذ وأثره ظهر، لا بريد يقول «تم».
5. وثيقة إنشاء أو اعتماد تصدر قبل الأسبوع {C.CALENDAR['evidence_by_week']} يجوز أن تغلق الملاحظة في هذه الدورة.
   دليل تشغيل عن فترة مضت لا يُصنع الآن: الملاحظة تُسجل غير ممتثلة في هذه الدورة، ويُخطط إغلاقها للدورة القادمة.
6. كل ما يحتاج قرار اللجنة يُربط بأحد اجتماعاتها.
"""


def lab_3_4():
    write_text("plan_rules.txt", PLAN_RULES)
    wb = openpyxl.Workbook()
    note_sheet(wb, ["سجل ملاحظات مجال حوكمة البيانات كما اعتمده مسؤول الامتثال، مرتبًا بدرجة الخطر.",
                    "اكتب الخطة في ورقة «الخطة»، ثم ضع كل إجراء في أسبوعه في ورقة «التقويم»."])
    sheet(wb, "السجل", ["الملاحظة", "المواصفة", "الأولوية", "الحالة", "السبب الجذري", "الدرجة"],
          [g[:6] for g in C.DG_FINDINGS], widths=[9, 14, 8, 50, 40, 8])
    sheet(wb, "الخطة", ["الملاحظة", "الإجراء", "المالك", "أسبوع الإنجاز", "دليل الإغلاق", "المتحقق", "تُغلق في هذه الدورة؟",
                        "اجتماع اللجنة"],
          [(g[0], "", "", "", "", "", "", "") for g in C.DG_FINDINGS], fill_cols=tuple(range(1, 8)),
          widths=[9, 40, 22, 10, 34, 18, 12, 10], validations={"G": ["نعم", "لا"]})
    sheet(wb, "التقويم", ["الأسبوع", "ما يحدث"], [(w, "") for w in range(0, C.CALENDAR["submission_week"] + 1)],
          fill_cols=(1,), widths=[9, 70])
    save(wb, "dg_findings.xlsx")
    put_ref("lab_3_4.json", {"calendar": C.CALENDAR, "owners": C.OWNERS,
                             "closable": {g[0]: g[6] for g in C.DG_FINDINGS}})


# ------------------------------------------------------------------ capstone
CAPSTONE = f"""
{head()}
قالب حزمة التدقيق التجريبي (المشروع الختامي)

اختر مجالين من الثلاثة (حوكمة البيانات، وجودة البيانات، والبيانات المفتوحة). مجلد مشترك واحد في Google Drive فيه:

1. فهرس الأدلة للمجالين: لكل مواصفة معيار قبولها، والبنود المقدمة لكل طبقة، وحكم كل بند وسببه.
2. ورقة الامتثال: 100% أو 0% لكل مواصفة، ثم الضوابط والمجال. لا قيمة بين الاثنتين.
3. ورقة النضج والمقاييس للمجالين: المستوى المدعوم لكل سؤال بإحالاته، وقيمة كل مقياس بحسابه.
4. سجل الملاحظات للمجالين، مرتبًا بقاعدة الجامعة.
5. الخطة التصحيحية وتقويم الجاهزية: المالك، والأسبوع، ودليل الإغلاق، والمتحقق.
6. سجل المساعد: كل ربط اقترحه مساعد، ومن اعتمده، وما غيّره. لا يدخل المساعد إلا ما صُنف «عام».

في التدقيق: ست دقائق تعرض فيها الحزمة، ثم ثماني دقائق تختار فيها المقيِّمة بنودًا من فهرسك وتسألك عن حكم كل منها،
وعن ترتيب ملاحظتين، وعن مالك إجراء وموعده.

مهمة المتابعة (خلال 30 يومًا): أغلق ملاحظتين من سجلك بدليل منقح، واكتب لكل منهما مذكرة تحقق قصيرة: ما الدليل
الجديد، ولماذا هو كاف بقواعد الجامعة، ومن تحقق منه. في جهتك أنت أو في الحالة، دون أي محتوى مصنف «مقيد» فأعلى.
"""


def main() -> None:
    if OUT.exists():
        for p in sorted(OUT.rglob("*"), reverse=True):
            p.unlink() if p.is_file() else p.rmdir()
    OUT.mkdir(parents=True, exist_ok=True)
    write_text("case_brief.txt", CASE_BRIEF)
    write_text("capstone_pack_template.txt", CAPSTONE)
    lab_1_1()
    lab_1_2()
    lab_1_3_4()
    lab_2_1_2()
    lab_2_3()
    lab_2_4()
    lab_3_1()
    lab_3_2()
    lab_3_3()
    lab_3_4()
    print("files:", sum(1 for p in OUT.rglob("*") if p.is_file()))


if __name__ == "__main__":
    main()
