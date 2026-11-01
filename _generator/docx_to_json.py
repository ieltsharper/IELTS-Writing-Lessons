"""Reads the IELTS Writing Task 2 Sentence Writing Workbook (.docx) and writes its content as JSON.

Usage: python -I docx_to_json.py <workbook.docx> <out.json>

Layout this expects (IELTS_WT2_Sentence_Writing_Workbook.docx):
  cover lines, "Hướng dẫn sử dụng sách" and "Bảng màu đánh dấu lỗi" (Heading 1, guide);
  "Unit 0  Diagnostic Test" and "Final Review" (Heading 1) with "Phần …" exercises and
  "Bảng …" tables; "Part A/B/…" (Heading 1) grouping "Unit N  Name" (Heading 2), each with
  "A. Lý thuyết" (theory: sub-headings, paragraphs, tables, "Câu mẫu", 1x1 boxes "Lỗi cần tránh"
  and "Chọn cho đúng") and "Bài tập N — …" (Heading 3) exercises;
  "Đáp án" (Heading 1) with the same units/exercises in the same order.
Exercise items are List Paragraphs starting with ●. Item kinds:
  choice    A/B/C option lines, or "A   B   C   D" after the sentence (legend box in the exercise)
  pick      "( x / y )" inside the sentence — choose the right phrase
  fix       "☐ đúng" after the sentence — find and correct errors (or say it is correct)
  translate "Bản dịch của bạn:" + "Viết lại với: …" — translate, then rewrite
  write     everything else (optional exam prompt and hints)
Bold, italic and struck-through runs are kept as <b>, <i>, <s>; criterion tags (TR/CC/LR/GRA)
become <span class="tag X">. Fails loudly on anything it cannot place.
"""
import html
import json
import re
import sys

import docx
from docx.table import Table
from docx.text.paragraph import Paragraph

CRIT = ('TR', 'CC', 'LR', 'GRA')


def runs_of(p):
    """[(text, fmt)] with neighbouring runs of the same format merged; fmt = (b, i, s)."""
    out = []
    for r in p.runs:
        if not r.text:
            continue
        k = (bool(r.bold), bool(r.italic), bool(r.font.strike))
        if out and out[-1][1] == k:
            out[-1] = (out[-1][0] + r.text, k)
        else:
            out.append((r.text, k))
    return out


def plain(runs):
    return re.sub(r'\s+', ' ', ''.join(t for t, _ in runs)).strip()


def strip_prefix(runs, prefix_re):
    m = re.match(prefix_re, ''.join(t for t, _ in runs))
    if not m:
        return runs
    n, out = m.end(), []
    for t, f in runs:
        if n >= len(t):
            n -= len(t)
            continue
        out.append((t[n:], f))
        n = 0
    return out


def strip_suffix(runs, suffix_re):
    m = re.search(suffix_re, ''.join(t for t, _ in runs))
    if not m:
        return runs
    keep, out = m.start(), []
    for t, f in runs:
        if keep <= 0:
            break
        out.append((t[:keep], f))
        keep -= len(t)
    return out


def to_html(runs):
    s = ''
    for t, (b, i, st) in runs:
        if b and t.strip() in CRIT:
            s += f' <span class="tag {t.strip()}">{t.strip()}</span> '
            continue
        e = html.escape(t, quote=False)
        if i:
            e = f'<i>{e}</i>'
        if st:
            e = f'<s>{e}</s>'
        if b:
            e = f'<b>{e}</b>'
        s += e
    s = re.sub(r'</b><b>|</i><i>|</s><s>', '', s)
    return re.sub(r'\s+', ' ', s).strip()


def level_of(text):
    m = re.match(r'\s*(●+)', text)
    return len(m.group(1)) if m else 0


DOT = r'\s*●+\s*'
FIXBOX = r'\s*☐\s*đúng\s*$'
ABCD = r'\s{2,}A\s+B\s+C\s+D\s*$'
OPTLINE = re.compile(r'([A-D])\s{2,}(.*)')
PICK = re.compile(r'\(\s*([^()]*?/[^()]*?)\s*\)')
UNITREF = r'\s*\[(Unit \d+)\]\s*'


def cell_paras(cell):
    return [runs_of(p) for p in cell.paragraphs if p.text.strip()]


def cell_html(cell):
    return '<br>'.join(to_html(r) for r in cell_paras(cell))


def table_block(t):
    head = [plain(cell_paras(c)[0]) if cell_paras(c) else '' for c in t.rows[0].cells]
    rows = []
    for r in t.rows[1:]:
        cells = [cell_html(c) for c in r.cells]
        lv = level_of(cells[0])
        if lv:
            cells[0] = re.sub(r'^●+\s*', '', cells[0])
        rows.append({'level': lv, 'cells': cells})
    return {'type': 'table', 'head': head, 'rows': rows}


def box_block(t):
    paras = cell_paras(t.cell(0, 0))
    head, body = plain(paras[0]), paras[1:]
    if head.startswith('Lỗi'):
        if len(body) % 3:
            raise SystemExit(f'Mistakes box not in triples: {head}')
        return {'type': 'mistakes', 'items': [
            {'bad': to_html(strip_prefix(body[j], r'\s*✗\s*')),
             'good': to_html(strip_prefix(body[j + 1], r'\s*✓\s*')),
             'note': to_html(body[j + 2])} for j in range(0, len(body), 3)]}
    if head.startswith('Chọn'):
        if len(body) % 2:
            raise SystemExit(f'Choose box not in pairs: {head}')
        return {'type': 'choose', 'items': [
            {'h': to_html(body[j]), 'html': to_html(body[j + 1])} for j in range(0, len(body), 2)]}
    raw = [''.join(t for t, _ in r).strip() for r in paras]
    if re.match(r'A\s{2,}', raw[0]):
        return {'type': 'legend', 'options': [
            {'letter': m.group(1), 'text': m.group(2).strip()}
            for m in (OPTLINE.match(r) for r in raw)]}
    raise SystemExit(f'Unknown box: {head}')


def main(src, dst):
    d = docx.Document(src)
    blocks = []
    for el in d.element.body.iterchildren():
        tag = el.tag.split('}')[1]
        if tag == 'p':
            p = Paragraph(el, d)
            if p.text.strip():
                blocks.append(('p', p.style.name if p.style is not None else '', p))
        elif tag == 'tbl':
            blocks.append(('t', '', Table(el, d)))

    data = {'cover': [], 'guide': [], 'parts': [], 'units': []}
    section = 'cover'
    unit = ex = None
    part = None          # 'intro' | 'theory' | 'exhead' | 'ex' | 'after'
    key_unit = key_ex = item = None
    key_idx = 0

    def new_unit(num, name, partless=False):
        u = {'num': num, 'name': name, 'intro': [], 'theory': [], 'exercises': [], 'after': []}
        data['units'].append(u)
        if partless:
            data['parts'].append({'title': '', 'units': [num]})
        else:
            data['parts'][-1]['units'].append(num)
        return u

    def unit_by_title(text):
        if text.startswith('Final Review'):
            return next(u for u in data['units'] if u['num'] == 'review')
        n = int(re.match(r'Unit\s+(\d+)', text).group(1))
        return next(u for u in data['units'] if u['num'] == n)

    def prose(dst_list, runs, text, style):
        if style == 'List Paragraph':
            if not dst_list or dst_list[-1]['type'] != 'samples':
                dst_list.append({'type': 'samples', 'items': []})
            dst_list[-1]['items'].append(to_html(runs))
        elif all(f[0] for _, f in runs) or (len(text) < 70 and not text.endswith(('.', ':'))):
            dst_list.append({'type': 'sub', 'text': text})
        else:
            dst_list.append({'type': 'p', 'html': to_html(runs)})

    for kind, style, obj in blocks:
        text = obj.text.strip() if kind == 'p' else ''
        runs = runs_of(obj) if kind == 'p' else []

        if kind == 'p' and style == 'Heading 1':
            if text.startswith(('Hướng dẫn', 'Bảng màu')):
                section = 'guide'
                data['guide'].append({'type': 'h1', 'text': text})
            elif text.startswith('Unit 0'):
                section = 'units'
                m = re.match(r'Unit\s+0\s+(.*)', text)
                unit, part = new_unit(0, m.group(1).strip(), True), 'intro'
            elif text.startswith('Final Review'):
                section = 'units'
                unit, part = new_unit('review', 'Final Review', True), 'intro'
            elif text.startswith('Part'):
                section = 'units'
                data['parts'].append({'title': re.sub(r'\s+', ' ', text), 'units': []})
            elif text.startswith('Đáp án'):
                section = 'key'
            else:
                raise SystemExit(f'Unknown Heading 1: {text}')
            continue

        if section == 'cover':
            if text != 'Mục lục':
                data['cover'].append(text)
            continue

        if section == 'guide':
            if kind == 't':
                data['guide'].append(table_block(obj))
            elif style == 'List Paragraph':
                data['guide'].append({'type': 'li', 'html': to_html(runs)})
            elif all(f[0] for _, f in runs) or len(text) < 60 and not text.endswith('.'):
                data['guide'].append({'type': 'h', 'html': to_html(runs)})
            else:
                data['guide'].append({'type': 'p', 'html': to_html(runs)})
            continue

        if section == 'units':
            if kind == 'p' and style == 'Heading 2':
                m = re.match(r'Unit\s+(\d+)\s+(.*)', text)
                unit, part = new_unit(int(m.group(1)), m.group(2).strip()), None
                continue
            if kind == 'p' and style == 'Heading 3':
                if re.match(r'(Bài tập \d|Phần \d|Phần [A-E] )', text):
                    part = 'ex'
                    ex = {'title': text, 'intro': [], 'items': []}
                    unit['exercises'].append(ex)
                elif text.startswith('A.'):
                    part = 'theory'
                elif text.startswith('B.'):
                    part = 'exhead'
                elif text.startswith('Bảng'):
                    part = 'after'
                    unit['after'].append({'type': 'sub', 'text': text})
                else:
                    raise SystemExit(f'Unknown Heading 3: {text}')
                continue
            if part in ('intro', 'theory', 'after'):
                dst_list = unit[part]
                if kind == 't':
                    one = len(obj.rows) == 1 and len(obj.columns) == 1
                    dst_list.append(box_block(obj) if one else table_block(obj))
                else:
                    prose(dst_list, runs, text, style)
                continue
            if part == 'ex':
                if kind == 't':
                    if not (len(obj.rows) == 1 and len(obj.columns) == 1):
                        raise SystemExit(f'Table inside exercise: {ex["title"]}')
                    b = box_block(obj)
                    if b['type'] != 'legend':
                        raise SystemExit(f'Box inside exercise: {ex["title"]}')
                    ex['legend'] = b['options']
                    continue
                if style == 'List Paragraph' and level_of(text):
                    r2 = strip_prefix(runs, DOT)
                    item = {'level': level_of(text), 'n': len(ex['items'])}
                    t2 = ''.join(t for t, _ in r2)
                    if re.search(FIXBOX, t2):
                        item['kind'] = 'fix'
                        r2 = strip_suffix(r2, FIXBOX)
                    elif re.search(ABCD, t2):
                        if 'legend' not in ex:
                            raise SystemExit(f'A B C D item without legend: {text}')
                        item['kind'] = 'choice'
                        item['options'] = ex['legend']
                        r2 = strip_suffix(r2, ABCD)
                    elif PICK.search(t2):
                        item['kind'] = 'pick'
                        item['picks'] = [[o.strip() for o in m.group(1).split('/')]
                                         for m in PICK.finditer(t2)]
                    item['cue'] = to_html(r2)
                    ex['items'].append(item)
                    continue
                if not ex['items']:
                    ex['intro'].append(to_html(runs))
                    continue
                it = ex['items'][-1]
                m = OPTLINE.match(text)
                if m and it.get('kind') in (None, 'choice'):
                    it['kind'] = 'choice'
                    it.setdefault('options', []).append(
                        {'letter': m.group(1), 'text': to_html(strip_prefix(runs, r'\s*[A-D]\s{2,}'))})
                elif text.startswith('Gợi ý:'):
                    h = to_html(strip_prefix(runs, r'\s*Gợi ý:\s*'))
                    it['hint'] = it['hint'] + '<br>' + h if 'hint' in it else h
                elif text.startswith('Bản dịch của bạn'):
                    it['kind'] = 'translate'
                elif text.startswith('Viết lại với:'):
                    it['rewriteWith'] = to_html(strip_prefix(runs, r'\s*Viết lại với:\s*'))
                elif style != 'List Paragraph' and 'hint' not in it and 'options' not in it:
                    it['prompt'] = (it['prompt'] + '<br>' if 'prompt' in it else '') + to_html(runs)
                else:
                    raise SystemExit(f'Unknown exercise line: {text}')
                continue
            raise SystemExit(f'Unplaced unit text ({part}): {text}')

        if section == 'key':
            if style == 'Normal' and re.match(r'(Unit\s+\d+\s|Final Review)', text) and not runs[0][0].startswith(' '):
                if re.match(r'(Unit\s+\d+\s{2}|Final Review$)', text):
                    key_unit, key_ex = unit_by_title(text), None
                    continue
            if style == 'Heading 3':
                key_ex = next((e for e in key_unit['exercises'] if e['title'] == text), None)
                if key_ex is None:
                    raise SystemExit(f'Key exercise not found: {text}')
                key_idx = 0
                continue
            if key_ex is None:
                raise SystemExit(f'Key text outside an exercise: {text}')
            if style == 'List Paragraph' and level_of(text):
                if key_idx >= len(key_ex['items']):
                    raise SystemExit(f'Too many keys in {key_ex["title"]}: {text}')
                item = key_ex['items'][key_idx]
                key_idx += 1
                if level_of(text) != item['level']:
                    raise SystemExit(f'Level mismatch in key ({key_ex["title"]}): {text}')
                r2 = strip_prefix(runs, DOT)
                m = re.match(UNITREF, plain(r2))
                if m:
                    item['ref'] = m.group(1)
                    r2 = strip_prefix(r2, UNITREF)
                k = item.get('kind')
                if k in ('choice', 'pick'):
                    m = re.match(r'\s*(.*?)\s+—\s+', ''.join(t for t, _ in r2))
                    if not m:
                        raise SystemExit(f'Choice key expected: {text}')
                    if k == 'choice':
                        item['correct'] = m.group(1).strip()
                        if item['correct'] not in [o['letter'] for o in item['options']]:
                            raise SystemExit(f'Bad letter: {text}')
                    else:
                        item['answer'] = m.group(1).strip()
                    item['note'] = to_html(strip_prefix(r2, r'\s*.*?\s+—\s+'))
                elif k == 'translate':
                    if not plain(r2).startswith('Bản dịch:'):
                        raise SystemExit(f'Translation key expected: {text}')
                    item['translation'] = to_html(strip_prefix(r2, r'\s*Bản dịch:\s*'))
                else:
                    item['answer'] = to_html(r2)
                continue
            if item is None:
                raise SystemExit(f'Key line before any item: {text}')
            if text.startswith('Viết lại:'):
                item['answer'] = to_html(strip_prefix(runs, r'\s*Viết lại:\s*'))
            elif text.startswith('Cũng chấp nhận:'):
                item['alt'] = to_html(strip_prefix(runs, r'\s*Cũng chấp nhận:\s*'))
            else:
                item.setdefault('notes', []).append(to_html(runs))
            continue

    for u in data['units']:
        for e in u['exercises']:
            for it in e['items']:
                it.setdefault('kind', 'write')
                where = f'Unit {u["num"]} {e["title"]} #{it["n"] + 1}'
                if 'answer' not in it and 'correct' not in it:
                    raise SystemExit(f'No answer: {where}')
                if it['kind'] == 'choice' and len(it['options']) < 2:
                    raise SystemExit(f'Options: {where}')
                if it['kind'] == 'translate' and 'translation' not in it:
                    raise SystemExit(f'No translation key: {where}')
                if it['kind'] == 'pick' and len(it['picks']) != 1:
                    raise SystemExit(f'Several pick groups: {where}')
    with open(dst, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    for u in data['units']:
        print(u['num'], [(len(e['items']), sorted({i['kind'] for i in e['items']})) for e in u['exercises']])


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
