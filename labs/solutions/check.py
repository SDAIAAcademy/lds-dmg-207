# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl>=3.1"]
# ///
"""Run the reference solution of every lab against the course files and check it.

    uv run labs/solutions/check.py            # all labs
    uv run labs/solutions/check.py 2.1 3.1    # some labs

Every lab of this course is done in a spreadsheet or a written document; none calls a model. The
assistant's mapping in lab 3.1 is a scripted stand-in written into mapping_pilot.xlsx. This script
recomputes from the files what the reference solution computes (the cascade, the levels, the
metrics, the error rates, the ranking), checks the answer keys in solutions/reference/ against the
files, and prints the numbers the debrief slides quote and PASS per lab. Nothing is written to disk.
"""
from __future__ import annotations

import csv
import json
import statistics
import sys
from datetime import date
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
FILES = HERE.parent / "files"
REF = HERE / "reference"
W0, W1 = date(2026, 7, 1), date(2027, 6, 30)
LABS: dict = {}


def lab(lab_id):
    def wrap(fn):
        LABS[lab_id] = fn
        return fn
    return wrap


def ref(name):
    return json.loads((REF / name).read_text(encoding="utf-8"))


def rows(rel, sheet):
    ws = openpyxl.load_workbook(FILES / rel, read_only=True)[sheet]
    it = ws.iter_rows(values_only=True)
    head = next(it)
    return [dict(zip(head, r)) for r in it if any(v is not None for v in r)]


def read_csv(rel):
    with open(FILES / rel, encoding="utf-8", newline="") as f:
        return list(csv.DictReader(f))


def pct(a, b):
    return round(100 * a / b, 1)


@lab("1.1")
def lab_1_1():
    r = ref("lab_1_1.json")
    res = rows("sample_result.xlsx", "الامتثال")
    no_ev = sorted(x["الرمز"] for x in res if x["النتيجة المسجلة %"] == 100 and not x["مرجع الدليل"])
    not_bin = sorted(x["الرمز"] for x in res if x["النتيجة المسجلة %"] not in (0, 100))
    assert no_ev == r["no_evidence"] and not_bin == r["not_binary"], (no_ev, not_bin)

    def cascade(val):
        ctl = {}
        for x in res:
            ctl.setdefault((x["المجال"], x["الضابط"]), []).append(val(x))
        dom = {}
        for (d, _), v in ctl.items():
            dom.setdefault(d, []).append(sum(v) / len(v))
        dm = {d: round(sum(v) / len(v), 1) for d, v in dom.items()}
        return dm, round(sum(sum(v) / len(v) for v in dom.values()) / len(dom), 1)

    d_rec, e_rec = cascade(lambda x: x["النتيجة المسجلة %"])
    d_cor, e_cor = cascade(lambda x: 100 if x["النتيجة المسجلة %"] == 100 and x["مرجع الدليل"] else 0)
    assert d_rec == r["domain_recorded"] and e_rec == r["entity_recorded"]
    assert d_cor == r["domain_corrected"] and e_cor == r["entity_corrected"]
    print(f"  recorded 100 without evidence: {no_ev}; not binary: {not_bin}")
    print(f"  entity score as recorded {e_rec}%, after correction {e_cor}%; domains {d_cor}")


@lab("1.2")
def lab_1_2():
    r = ref("lab_1_2.json")
    items = rows("instruments.xlsx", "الصكوك")
    assert len(items) == 16 and {x["الرمز"] for x in items} == set(r)
    counts = {}
    for v in r.values():
        counts[v["key"]] = counts.get(v["key"], 0) + 1
    assert counts == {"ينطبق": 11, "ينطبق ويُقيَّم عبر جهة أخرى": 1, "غير ملزم": 2, "لا ينطبق": 2}, counts
    print(f"  {counts}")


@lab("1.3")
def lab_1_3():
    crit = ref("criteria.json")
    sheet = rows("criteria_template.xlsx", "المعايير")
    assert [x["الرمز"] for x in sheet] == [f"S{k:02d}" for k in range(1, 21)]
    approvers = {x["الوثيقة"]: x["يعتمدها"] for x in rows("criteria_template.xlsx", "جدول الاعتماد")}
    assert len(approvers) == 11
    for k in range(1, 11):
        c = crit[f"S{k:02d}"]
        assert c["layers"] and all(c[f] for f in ("exists", "approver", "operates", "types", "window"))
    layers = {f"S{k:02d}": crit[f"S{k:02d}"]["layers"] for k in range(1, 11)}
    n_op = sum("التطبيق" in v for v in layers.values())
    print(f"  S01 to S10: {n_op} need proof of operation; layers per spec {[len(v) for v in layers.values()]}")


@lab("1.4")
def lab_1_4():
    crit = ref("criteria.json")
    draft = ref("colleague_draft.json")
    codes = {x["الرمز"] for x in rows("criteria_template.xlsx", "العيوب")}
    for spec, key in draft.items():
        found = {w for w in key.replace("؛", " ").split() if w.startswith("D") and w[1:].isdigit()}
        assert found and found <= codes, (spec, found)
    op_only = [s for s in (f"S{k:02d}" for k in range(11, 21)) if crit[s]["layers"] == ["التطبيق"]]
    print(f"  colleague draft: {len(draft)} rows with defects; operation-only specs {op_only}")


def compliance(key, crit):
    out = {}
    for spec, c in crit.items():
        ok = all(any(v["spec"] == spec and v["layer"] == layer and v["key"] == "كاف" for v in key.values())
                 for layer in c["layers"])
        out[spec] = 100 if ok else 0
    return out


def check_items(lo, hi):
    key = ref("evidence_key.json")
    reg = {x["البند"]: x for x in rows("evidence_register.xlsx", "البنود")}
    assert len(reg) == 60
    for i in (f"E{k:02d}" for k in range(lo, hi + 1)):
        x, k = reg[i], key[i]
        dt = x["التاريخ"]
        if x["النوع"] == "لقطة شاشة" and not dt:
            assert k["reason"] in ("R7", "R6"), i
        if dt and date.fromisoformat(dt) > W1 and k["reason"] != "R5":
            assert "ربع" in x["العنوان"] or "تقارير" in x["النوع"], i  # a quarterly report in its usual time
        if k["key"] == "كاف":
            assert not k["reason"], i
        else:
            assert k["reason"], i
    sub = {i: key[i] for i in key if lo <= int(i[1:]) <= hi}
    counts = {}
    for v in sub.values():
        counts[v["key"]] = counts.get(v["key"], 0) + 1
    return key, counts


@lab("2.1")
def lab_2_1():
    _, counts = check_items(1, 30)
    assert counts == {"كاف": 16, "جزئي": 5, "غير كاف": 9}, counts
    print(f"  E01 to E30: {counts}")


@lab("2.2")
def lab_2_2():
    key, counts = check_items(31, 60)
    assert counts == {"كاف": 13, "جزئي": 6, "غير كاف": 11}, counts
    comp = compliance(key, ref("criteria.json"))
    by_dom = {"DG": [], "DQ": [], "OD": []}
    for k in range(1, 21):
        by_dom[["DG", "DQ", "OD"][0 if k <= 8 else 1 if k <= 14 else 2]].append(comp[f"S{k:02d}"])
    lab_specs = [f"S{k:02d}" for k in range(11, 21)]
    lab_ok = sum(comp[s] == 100 for s in lab_specs)
    assert lab_ok == 6, lab_ok
    print(f"  E31 to E60: {counts}; agreement target 85% of 60 = {-(-60 * 85 // 100)} items")
    print(f"  main step: S11 to S20 compliant {lab_ok} of 10; non-compliant "
          f"{[s for s in lab_specs if comp[s] == 0]}")
    print(f"  compliant specs {sum(v == 100 for v in comp.values())} of 20: "
          + ", ".join(f"{d} {sum(v == 100 for v in vs)}/{len(vs)}" for d, vs in by_dom.items()))
    print(f"  non-compliant: {[s for s, v in comp.items() if v == 0]}")


@lab("2.3")
def lab_2_3():
    r = ref("lab_2_3.json")
    ev = {x["الدليل"]: (x["يُقدَّم للمستوى"], x["الحالة"]) for x in rows("questionnaire.xlsx", "الأدلة")}
    for q in rows("questionnaire.xlsx", "الأسئلة"):
        cited = [c.strip() for c in q["الأدلة المحال إليها"].split("،")]
        good = {ev[c][0] for c in cited if ev[c][1] == "كاف"}
        level = 0
        while level + 1 in good:
            level += 1
        assert level == r["questions"][q["السؤال"]]["supported"], q["السؤال"]
    over = [q for q, v in r["questions"].items() if v["supported"] < v["claimed"]]
    under = [q for q, v in r["questions"].items() if v["supported"] > v["claimed"]]
    print(f"  over-claimed {over}; under-claimed {under}")
    print(f"  domain claimed {r['domain_claimed']}; supported {r['domain_supported']}")


@lab("2.4")
def lab_2_4():
    r = ref("lab_2_4.json")
    cat = read_csv("catalogue_extract.csv")
    latest = {}
    for x in cat:
        if x["asset_id"] not in latest or x["updated"] > latest[x["asset_id"]]["updated"]:
            latest[x["asset_id"]] = x
    el = [x for x in latest.values() if x["critical"] == "yes" and x["status"] == "active"
          and not x["asset_name"].startswith("TEST_")]
    num = sum(all(x[f] for f in ("owner", "steward", "description", "classification")) for x in el)
    assert (num, len(el)) == (r["M1"]["num"], r["M1"]["den"])
    runs = read_csv("quality_runs.csv")
    last = {}
    for x in runs:
        dt = date.fromisoformat(x["run_date"])
        if x["data_priority"] == "1" and x["status"] == "enabled" and W0 <= dt <= W1:
            if x["rule_id"] not in last or x["run_date"] > last[x["rule_id"]]["run_date"]:
                last[x["rule_id"]] = x
    n2 = sum(x["result"] == "pass" for x in last.values())
    assert (n2, len(last)) == (r["M2"]["num"], r["M2"]["den"])
    req = read_csv("sharing_requests.csv")
    durs = [(date.fromisoformat(x["decided"]) - date.fromisoformat(x["received"])).days for x in req
            if x["status"] in ("approved", "rejected") and W0 <= date.fromisoformat(x["decided"]) <= W1]
    assert statistics.median(durs) == r["M3"]["median_days"]
    od = [x for x in read_csv("opendata_register.csv") if x["status"] == "published"]
    lim = {"monthly": 31, "quarterly": 92, "annual": 366}
    on = sum((W1 - date.fromisoformat(x["last_update"])).days <= lim[x["frequency"]] for x in od)
    assert (on, len(od)) == (r["M4"]["num"], r["M4"]["den"])
    print(f"  M1 {num}/{len(el)} = {pct(num, len(el))}% (all critical rows, no exclusions: {r['M1']['naive_all_rows']}%)")
    print(f"  M2 {n2}/{len(last)} = {pct(n2, len(last))}%; M3 median {statistics.median(durs)} days over {len(durs)}"
          f" requests, {r['M3']['open_over_30']} open over 30 days; M4 {on}/{len(od)} = {pct(on, len(od))}%")


@lab("3.1")
def lab_3_1():
    r = ref("lab_3_1.json")
    out = rows("mapping_pilot.xlsx", "مخرجات المساعد")
    key = {x["الوثيقة"]: x for x in rows("mapping_pilot.xlsx", "مفتاح المقيم")}
    ids = {x["المواصفة"] for x in rows("mapping_pilot.xlsx", "فهرس المواصفات")}
    assert len(ids) == 190
    measured = [x for x in out if x["تصنيف المدخل"] == "عام"]
    excluded = [x["الوثيقة"] for x in out if x["تصنيف المدخل"] != "عام"]
    assert excluded == list(r["excluded"])
    wrong = [x for x in measured if x["مواصفة المساعد"] != key[x["الوثيقة"]]["المواصفة الصحيحة"]]
    assert len(measured) == r["n"] and len(wrong) == r["wrong"]
    for x in wrong:
        a, k = x["مواصفة المساعد"], key[x["الوثيقة"]]["المواصفة الصحيحة"]
        t = r["errors"][x["الوثيقة"]]["type"]
        if not a:
            assert t == "لم يُربط"
        elif a not in ids:
            assert t == "رقم مواصفة غير موجود"
        elif a.split(".")[0] != k.split(".")[0]:
            assert t == "مجال آخر"
        elif a.split(".")[:2] != k.split(".")[:2]:
            assert t == "ضابط آخر في المجال نفسه"
        else:
            assert t == "مواصفة أخرى في الضابط نفسه"
    assert all(k in ids for k in (v["المواصفة الصحيحة"] for v in key.values()))
    by_type = {}
    for x in measured:
        t = key[x["الوثيقة"]]["النوع"]
        by_type.setdefault(t, [0, 0])
        by_type[t][0] += 1
        by_type[t][1] += x in wrong
    hi = sum(x["ثقة المساعد"] == "عالية" for x in wrong)
    hi_all = sum(x["ثقة المساعد"] == "عالية" for x in measured)
    types = {}
    for e in r["errors"].values():
        types[e["type"]] = types.get(e["type"], 0) + 1
    print(f"  wrong {len(wrong)}/{len(measured)} = {pct(len(wrong), len(measured))}%; excluded {excluded}")
    print(f"  by type {by_type}; high-confidence wrong {hi} of {hi_all} high-confidence mappings")
    print(f"  error types {types}")


@lab("3.2")
def lab_3_2():
    b = rows("error_summary_batch2.xlsx", "الدفعة الثانية")
    tot = sum(x["الوثائق"] for x in b)
    wrong = sum(x["الربط الخاطئ"] for x in b)
    assert (tot, wrong) == (60, 11)
    print("  batch 2: " + ", ".join(f"{x['النوع']} {x['الربط الخاطئ']}/{x['الوثائق']}" for x in b)
          + f"; total {wrong}/{tot} = {pct(wrong, tot)}%")


@lab("3.3")
def lab_3_3():
    r = ref("lab_3_3.json")
    obs = {x["المشاهدة"]: x for x in rows("observations.xlsx", "المشاهدات")}
    imp = {"P1": 3, "P2": 2, "P3": 1}
    lik = {"كل الجامعة": 3, "جزئي": 2, "حالة واحدة": 1}
    used = set()
    for f in r["register"]:
        i = max(imp[obs[o]["الأولوية"]] for o in f["obs"])
        li = max(lik[obs[o]["الامتداد"]] for o in f["obs"])
        assert (i, li, i * li) == (f["impact"], f["likelihood"], f["score"]), f["id"]
        used |= set(f["obs"])
    assert used | set(r["not_findings"]) == set(obs) and not used & set(r["not_findings"])
    order = [f["id"] for f in r["register"]]
    print(f"  {len(order)} findings from {len(obs)} observations; order {order}; scores "
          f"{[f['score'] for f in r['register']]}; not findings {sorted(r['not_findings'])}")


@lab("3.4")
def lab_3_4():
    r = ref("lab_3_4.json")
    reg = rows("dg_findings.xlsx", "السجل")
    assert [x["الملاحظة"] for x in reg] == list(r["owners"])
    assert all(o != "مكتب إدارة البيانات" for o in r["owners"].values())
    yes = sorted(g for g, v in r["closable"].items() if v == "نعم")
    print(f"  closable this cycle {yes}; reported non-compliant and planned for next cycle "
          f"{sorted(g for g, v in r['closable'].items() if v == 'لا')}; submission week {r['calendar']['submission_week']}")


def main(argv):
    wanted = argv or list(LABS)
    failed = 0
    for lab_id in wanted:
        print(f"lab {lab_id}")
        try:
            LABS[lab_id]()
            print("  PASS")
        except AssertionError as e:
            failed += 1
            print(f"  FAIL {e}")
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main(sys.argv[1:])
