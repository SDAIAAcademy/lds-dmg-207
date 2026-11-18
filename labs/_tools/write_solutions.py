# /// script
# requires-python = ">=3.10"
# ///
"""Write the reference solutions labs/solutions/D*.md from the answer keys and the case data.

    uv run labs/_tools/write_solutions.py

Run after prepare_files.py. The numbers come from solutions/reference/, which check.py verifies
against the files.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import case_data as C  # noqa: E402

LABS = Path(__file__).resolve().parents[1]
SOL = LABS / "solutions"
REF = SOL / "reference"
SPEC = {s[0]: s for s in C.SPECS}


def ref(name):
    return json.loads((REF / name).read_text(encoding="utf-8"))


def write(name, title, start, lines):
    head = [f"# الحل المرجعي: المختبر {title}", "",
            f"نقطة البداية: {start}. التحقق: `uv run labs/solutions/check.py {title.split(' ')[0]}`.", ""]
    (SOL / name).write_bytes(("\n".join(head + lines).strip() + "\n").encode("utf-8"))


def table(header, rows):
    out = ["| " + " | ".join(header) + " |", "|" + "---|" * len(header)]
    out += ["| " + " | ".join(str(c) for c in r) + " |" for r in rows]
    return out


def criteria_rows(lo, hi):
    crit = ref("criteria.json")
    return table(["الرمز", "المواصفة", "الطبقات", "ما يجب أن يوجد", "من يعتمده", "ما يثبت أنه مطبق", "أنواع الأدلة", "النافذة"],
                 [(c, SPEC[c][1], "، ".join(v["layers"]), v["exists"], v["approver"], v["operates"], v["types"], v["window"])
                  for c, v in crit.items() if lo <= int(c[1:]) <= hi])


def main():
    r = ref("lab_1_1.json")
    write("D1_1_three_parts.md", "1.1 · خريطة الأجزاء الثلاثة", "`files/method_brief.txt` و`files/sample_result.xlsx`", [
        f"نتائج لا تصح: {'، '.join(r['no_evidence'])} سُجلت 100 بلا مرجع دليل، و{'، '.join(r['not_binary'])} سُجلت 50 والمنهجية ثنائية.", "",
        *table(["المجال", "كما سُجل %", "بعد التصحيح %"], [(k, r["domain_recorded"][k], r["domain_corrected"][k]) for k in r["domain_recorded"]]),
        "", f"الجامعة: {r['entity_recorded']}% كما سُجلت، و{r['entity_corrected']}% بعد التصحيح (الرفع: الضابط ثم المجال ثم الجهة، افتراض المقرر).", "",
        "## الخريطة", "", *table(["الجزء", "ما يُقاس", "كيف تُحسب الدرجة", "أنواع الأدلة", "من يقدمها"], r["map"]),
    ])
    r = ref("lab_1_2.json")
    write("D1_2_instruments.md", "1.2 · جرد الالتزامات ونطاق الدورة", "`files/instruments.xlsx`",
          table(["الرمز", "الصك", "الحكم", "السبب"], [(i, n, r[i]["key"], r[i]["reason"]) for i, n, *_ in C.INSTRUMENTS])
          + ["", "المجموع: ينطبق 11، وينطبق ويُقيَّم عبر جهة أخرى 1، وغير ملزم 2، ولا ينطبق 2."])
    write("D1_3_criteria.md", "1.3 · معايير لعشر مواصفات", "`files/spec_extract.txt` و`files/criteria_template.xlsx`",
          ["الإجابة النموذجية. صياغة أخرى صحيحة إذا حملت العناصر نفسها.", "", *criteria_rows(1, 10), "",
           "تطلب دليل تشغيل: S02 وS03 وS04 وS05 وS08. الفخ: S06 لا تطلب تقارير، فطلبها عيب D4."])
    draft = ref("colleague_draft.json")
    write("D1_4_review.md", "1.4 · عشر مواصفات أخرى ومسودة زميل", "`files/spec_extract.txt` و`files/criteria_template.xlsx`",
          ["توزع هذه الإجابة في الخطوة الرابعة فقط.", "", *criteria_rows(11, 20), "", "## عيوب مسودة الزميل", "",
           *table(["الرمز", "العيوب"], list(draft.items())), "",
           "التشغيلية البحتة: S17 وS18 وS19؛ وS12 تطلب تطبيقًا وفاعلية دون وثيقة إنشاء."])
    key = ref("evidence_key.json")
    for name, title, lo, hi in (("D2_1_items_first.md", "2.1 · ثلاثون بندًا دون مساعد", 1, 30),
                                ("D2_2_items_second.md", "2.2 · ثلاثون بندًا آخر وقياس الاتفاق", 31, 60)):
        rows = [(i, SPEC[v["spec"]][1], v["layer"], v["key"], v["reason"], v["gap"]) for i, v in key.items() if lo <= int(i[1:]) <= hi]
        counts = {}
        for row in rows:
            counts[row[3]] = counts.get(row[3], 0) + 1
        extra = []
        if hi == 60:
            crit = ref("criteria.json")
            comp = {s: all(any(v["spec"] == s and v["layer"] == L and v["key"] == "كاف" for v in key.values()) for L in c["layers"])
                    for s, c in crit.items()}
            extra = ["", "## امتثال المواصفات العشرين من المفتاح", "",
                     *table(["الرمز", "المواصفة", "الامتثال"], [(s, SPEC[s][1], "100%" if ok else "0%") for s, ok in comp.items()]),
                     "", "حوكمة البيانات 4 من 8، وجودة البيانات 3 من 6، والبيانات المفتوحة 4 من 6: 11 من 20.",
                     "السجل كاملًا مقيّم: الهدف اتفاق 85% على الأقل مع المفتاح، أي 51 بندًا من 60."]
        write(name, title, "`files/evidence_register.xlsx` و`files/sufficiency_rules.txt`",
              table(["البند", "المواصفة", "الطبقة", "الحكم", "السبب", "الفجوة"], rows)
              + ["", "المجموع: " + "، ".join(f"{k} {v}" for k, v in counts.items()) + "."] + extra)
    r = ref("lab_2_3.json")
    write("D2_3_questionnaire.md", "2.3 · استبانة بإحالات كاملة", "`files/questionnaire.xlsx`",
          table(["السؤال", "المجال", "المدعى", "المدعوم", "أول مستوى ناقص"],
                [(q, v["domain"], v["claimed"], v["supported"], v["first_missing"] or "") for q, v in r["questions"].items()])
          + ["", *table(["المجال", "المدعى", "المدعوم"], [(k, r["domain_claimed"][k], r["domain_supported"][k]) for k in r["domain_claimed"]]),
             "", "Q04 ادعى أقل من أدلته (2 والمدعوم 4). Q11 بلا دليل للمستوى 2، فيتوقف عند 1."])
    r = ref("lab_2_4.json")
    write("D2_4_metrics.md", "2.4 · أربعة مقاييس من المستخرجات", "`files/metric_definitions.txt` والمستخرجات الأربعة", [
        *table(["المقياس", "البسط", "المقام", "القيمة"], [
            ("M1", r["M1"]["num"], r["M1"]["den"], f"{r['M1']['value']}%"),
            ("M2", r["M2"]["num"], r["M2"]["den"], f"{r['M2']['value']}%"),
            ("M3", "", r["M3"]["n"], f"وسيط {r['M3']['median_days']} يومًا؛ {r['M3']['open_over_30']} طلبات مفتوحة أكثر من 30 يومًا"),
            ("M4", r["M4"]["num"], r["M4"]["den"], f"{r['M4']['value']}%")]),
        "", f"M1: {r['M1']['rows']} صفًا، منها {r['M1']['duplicates']} مكررة؛ دون استثناء تصير القيمة {r['M1']['naive_all_rows']}%.",
        "M1 يستبعد ستة أصول ملغاة وثلاثة تجريبية. M2 يستبعد أربع قواعد معطلة وتشغيلات بعد نهاية الفترة. M3 يستبعد المسحوبة.",
        f"M4: المتأخرة {'، '.join(r['M4']['late'])}؛ المحذوفة والمسودة خارج المقام."])
    r = ref("lab_3_1.json")
    write("D3_1_mapping_errors.md", "3.1 · أخطاء المساعد مقابل المفتاح", "`files/mapping_pilot.xlsx`", [
        f"الخطأ الكلي: {r['wrong']} من {r['n']}. لكل نوع: " + "، ".join(f"{t} {w} من {n}" for t, (n, w) in r["by_type"].items()) + ".", "",
        *table(["الوثيقة", "نوع الخطأ", "ثقة المساعد"], [(k, v["type"], v["confidence"]) for k, v in r["errors"].items()]),
        "", "خمسة أخطاء من 19 ربطًا بثقة عالية: الثقة ليست دليلًا على الصحة.",
        "P41 محضر مصنف «مقيد» دخل الدفعة: يُستبعد من القياس، ويُسجل خرقًا لمبادئ الذكاء الاصطناعي التوليدي للجهات الحكومية.",
        "المخرجات مكتوبة سلفًا في الملف لتطابق ما يقيسه المختبر؛ لم تُستدع أي نماذج."])
    write("D3_2_usage_rule.md", "3.2 · قاعدة استخدام مكتوبة", "`files/error_summary_batch2.xlsx` و`files/usage_rule_template.txt`", [
        "1. النطاق: ربط أولي للوثائق بالمواصفات فقط. لا يحكم المساعد على الكفاية ولا يكتب الملاحظات.",
        "2. المدخلات: ما صُنف «عام» فقط، ويتحقق أخصائي الامتثال من التصنيف قبل الإدخال. لا يدخل «مقيد» فأعلى.",
        "3. الأنواع: السياسات (خطأ 1 من 20 في الدفعة الثانية، و1 من 14 في الأولى) ربط أولي يراجعه إنسان. التقارير (7 من 20) والصفحات (3 من 20) يربطها إنسان من البداية.",
        "4. التوقيع: أخصائي الامتثال يراجع كل ربط ويوقعه؛ التوقيع يعني أنه قرأ الوثيقة والمواصفة.",
        "5. السجل: الوثيقة، وتصنيفها، والأداة وإصدارها، والمواصفة المقترحة، والمعتمدة، ومن وقع، ومتى، وما غُيّر.",
        "6. إعادة القياس: 20 ربطًا موقعًا من السياسات كل دورة، على مفتاح تعده المراجعة الداخلية.",
        "7. التوقف: 3 أخطاء أو أكثر في العينة توقف الربط الأولي للسياسات حتى قياس جديد. إن كان الخطأ الحقيقي 5% يتوقف خطأً بنسبة 7.5% تقريبًا، وإن كان 25% يتوقف بنسبة 91% تقريبًا.",
        "8. للجنة: نسبة الخطأ في كل دورة، وأي خرق للتصنيف، وقرار الاستمرار أو التوقف."])
    r = ref("lab_3_3.json")
    write("D3_3_register.md", "3.3 · سجل ملاحظات مرتب", "`files/observations.xlsx` و`files/ranking_rule.txt`", [
        *table(["الترتيب", "الملاحظة", "المشاهدات", "المواصفة", "السبب الجذري", "الأثر", "الاحتمال", "الدرجة"],
               [(f["rank"], f["id"], "، ".join(f["obs"]), f["specs"], f["cause"], f["impact"], f["likelihood"], f["score"]) for f in r["register"]]),
        "", "ليست ملاحظات:", *[f"- {k}: {v}" for k, v in r["not_findings"].items()],
        "", "طلب العميد حذف O07 (F05): تبقى الملاحظة، ويُكتب الطلب والرد في عمود الملاحظات."])
    r = ref("lab_3_4.json")
    rows = []
    actions = {
        "G01": ("عرض الخطة على اللجنة واعتمادها", "2", "محضر اللجنة الموقع بالاعتماد", "مسؤول الامتثال"),
        "G02": ("تقرير متابعة ربعي بموعد ثابت في تقويم المكتب", "4", "تقارير الربعين التاليين في وقتها", "مسؤول الامتثال"),
        "G03": ("تعميم السياسة مع إقرار إلكتروني يُسجل في كل ربع", "6", "سجل الإقرارات داخل الفترة القادمة", "مسؤول الامتثال"),
        "G04": ("تقويم سنوي لاجتماعات اللجنة، وتوقيع المحضر في أسبوع الاجتماع", "2", "محاضر موقعة في الدورة القادمة", "مسؤول الامتثال"),
        "G05": ("اعتماد الوصف الوظيفي من الرئيس التنفيذي للبيانات", "4", "الوصف الموقع", "المراجعة الداخلية"),
        "G06": ("فحص التصنيف قبل إدخال أي وثيقة في المساعد، وإبلاغ مسؤول حماية البيانات الشخصية", "1", "قائمة فحص موقعة لكل دفعة", "المراجعة الداخلية"),
    }
    for g in C.DG_FINDINGS:
        a = actions[g[0]]
        rows.append((g[0], a[0], r["owners"][g[0]], a[1], a[2], a[3], r["closable"][g[0]]))
    write("D3_4_plan.md", "3.4 · خطة تصحيحية وتقويم الجاهزية", "`files/dg_findings.xlsx` و`files/plan_rules.txt`", [
        *table(["الملاحظة", "الإجراء", "المالك", "الأسبوع", "دليل الإغلاق", "المتحقق", "تُغلق في الدورة؟"], rows),
        "", "G02 وG03 وG04 تُسجل غير ممتثلة في هذه الدورة: دليل التشغيل عن فترة مضت لا يُصنع الآن.",
        "اللجنة: الأسبوع 2 اعتماد الخطة (G01) والتقويم السنوي (G04)؛ الأسبوع 6 تقرير التقدم؛ الأسبوع 10 مراجعة الحزمة قبل التقديم في الأسبوع 14."])
    print("solutions written")


if __name__ == "__main__":
    main()
