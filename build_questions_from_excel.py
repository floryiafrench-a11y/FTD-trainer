#!/usr/bin/env python3
import json
import re
import sys
from pathlib import Path
from datetime import datetime
import openpyxl

def clean_text(s):
    if s is None:
        return ""
    s = str(s).strip().replace("\xa0", " ")
    s = re.sub(r"\s+", " ", s)
    return s

def parse_labeled_options(text):
    text = clean_text(text)
    parts = re.split(r'\s*;\s*(?=(?:[A-ZА-Я]|\d+)[\)\.]\s*)', text)
    out = []
    for part in parts:
        part = part.strip().strip(';')
        if not part:
            continue
        m = re.match(r'((?:[A-ZА-Я]|\d+))[\)\.]\s*(.*)', part)
        if m:
            out.append({"id": m.group(1), "text": m.group(2).strip()})
        else:
            out.append({"id": "", "text": part})
    return out

def detect_type(qopt, a):
    a_str = clean_text(a)
    qopt_str = clean_text(qopt)
    if qopt_str and re.fullmatch(r'(?:[A-ZА-Я]\d+,)*[A-ZА-Я]\d+', a_str):
        return 'matching'
    if qopt_str and re.fullmatch(r'\d+', a_str):
        return 'sequence'
    if qopt_str and re.fullmatch(r'[A-ZА-Я]', a_str):
        return 'single'
    if qopt_str and re.fullmatch(r'(?:[A-ZА-Я],)+[A-ZА-Я]', a_str):
        return 'multiple'
    return 'text'

def split_matching_options(qopt):
    qopt_str = clean_text(qopt)
    if '||' in qopt_str:
        left, right = [part.strip() for part in qopt_str.split('||', 1)]
        return parse_labeled_options(left), parse_labeled_options(right)
    items = parse_labeled_options(qopt_str)
    left = [it for it in items if re.fullmatch(r'[A-ZА-Я]', it["id"])]
    right = [it for it in items if re.fullmatch(r'\d+', it["id"])]
    return left, right

def build_questions(excel_path, sheet_name='СВОД'):
    wb = openpyxl.load_workbook(excel_path, data_only=False)
    ws = wb[sheet_name]
    questions = []
    idx = 1
    for row in ws.iter_rows(min_row=2, values_only=True):
        q, qopt, a, _id, *rest = row
        if not q:
            continue
        if isinstance(a, float) and a.is_integer():
            a = int(a)
        q = clean_text(q)
        qopt_str = clean_text(qopt) if qopt is not None else ""
        a_str = clean_text(a)
        qtype = detect_type(qopt_str, a_str)

        item = {"n": idx, "type": qtype, "question": q}

        if qtype in ('single', 'multiple'):
            item["options"] = parse_labeled_options(qopt_str)
            item["answer"] = a_str.split(',')

        elif qtype == 'matching':
            left, right = split_matching_options(qopt_str)
            mapping = []
            for pair in a_str.split(','):
                pair = pair.strip()
                if not pair:
                    continue
                m = re.match(r'([A-ZА-Я])(\d+)', pair)
                if m:
                    mapping.append({"left": m.group(1), "right": m.group(2)})
            item["left"] = left
            item["right"] = right
            item["answer"] = mapping

        elif qtype == 'sequence':
            item["options"] = parse_labeled_options(qopt_str)
            item["answer"] = list(a_str)

        else:
            item["answer"] = a_str

        questions.append(item)
        idx += 1

    return questions

def main():
    excel_path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("Вопросы FTD_Апрель.xlsx")
    if not excel_path.exists():
        raise SystemExit(f"Файл не найден: {excel_path}")

    questions = build_questions(excel_path)
    meta = {
        "sourceFile": excel_path.name,
        "sheet": "СВОД",
        "generatedAt": datetime.now().isoformat(timespec='seconds'),
        "total": len(questions),
    }

    output_path = Path("questions.js")
    content = (
        "window.QUESTION_BANK_META = " + json.dumps(meta, ensure_ascii=False, indent=2) + ";\n\n"
        + "window.QUESTIONS = " + json.dumps(questions, ensure_ascii=False, indent=2) + ";\n"
    )
    output_path.write_text(content, encoding="utf-8")
    print(f"Готово: {output_path} | вопросов: {len(questions)}")

if __name__ == "__main__":
    main()
