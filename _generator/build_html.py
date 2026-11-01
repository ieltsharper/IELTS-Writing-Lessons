"""Builds the Writing workbook site: index.html plus one page per unit, from the parsed JSON.

Usage: python -I build_html.py <lesson.json> <out dir> <localStorage key prefix> [portal .env.pages]

<lesson>.edits.json next to the JSON holds the teacher's changes on top of the .docx
(dropGuide, replace, en — same format as the Speaking site). Pages: index.html (guide + list of
units) and unit-<n>.html / final-review.html. Each page saves answers under
'<prefix>-<unit>'; with the portal's .env.pages (VITE_API_URL) finished exercises are sent to
the portal (workbook.submit) and Submit asks for AI feedback (workbook.assess).
"""
import json
import pathlib
import re
import sys

HERE = pathlib.Path(__file__).parent
VI_CHARS = re.compile('[ăâđêôơưạảấầẩẫậắằẳẵặẹẻẽếềểễệỉịọỏốồổỗộớờởỡợụủứừửữựỳỵỷỹ]', re.I)


def api_url(env_path):
    if not env_path:
        return ''
    for line in pathlib.Path(env_path).read_text(encoding='utf-8').splitlines():
        if line.startswith('VITE_API_URL='):
            return line.split('=', 1)[1].strip()
    raise SystemExit('VITE_API_URL not found in ' + env_path)


def apply_english(data, en):
    """Instructions students act on in English (teacher's rule); theory, notes and the guide stay
    Vietnamese. Fails on a Vietnamese title, instruction, hint or option without an English version."""
    missing = []

    def tr(s):
        if s in en:
            return en[s]
        if VI_CHARS.search(s):
            missing.append(s)
        return s

    data['cover'] = [tr(c) for c in data['cover']]
    for u in data['units']:
        u['name'] = tr(u['name'])
        for ex in u['exercises']:
            ex['title'] = tr(ex['title'])
            ex['intro'] = [tr(x) for x in ex['intro']]
            for o in ex.get('legend', []):
                o['text'] = tr(o['text'])
            for it in ex['items']:
                for f in ('hint', 'rewriteWith'):
                    if it.get(f):
                        it[f] = tr(it[f])
                if it['kind'] == 'choice' and it['options'] is not ex.get('legend'):
                    for o in it['options']:
                        o['text'] = tr(o['text'])
    if missing:
        raise SystemExit(f'No English for {len(missing)}: ' + ' | '.join(missing[:20]))


def page_name(num):
    return 'final-review.html' if num == 'review' else f'unit-{num}.html'


def main(src, out_dir, store, env_path=None):
    data = json.loads(pathlib.Path(src).read_text(encoding='utf-8'))
    edits_path = pathlib.Path(src).with_suffix('.edits.json')
    edits = json.loads(edits_path.read_text(encoding='utf-8')) if edits_path.exists() else {}
    drop = edits.get('dropGuide', [])
    if not all(any(b.get('html') == h for b in data['guide']) for h in drop):
        raise SystemExit('dropGuide: a block was not found')
    data['guide'] = [b for b in data['guide'] if b.get('html') not in drop]
    raw = json.dumps(data, ensure_ascii=False)
    for old, new in edits.get('replace', []):
        o, n = json.dumps(old, ensure_ascii=False)[1:-1], json.dumps(new, ensure_ascii=False)[1:-1]
        if o not in raw:
            raise SystemExit(f'replace: not found: {old[:60]}')
        raw = raw.replace(o, n)
    data = json.loads(raw)
    apply_english(data, edits.get('en', {}))

    def label(u):
        return 'Final Review' if u['num'] == 'review' else f"Unit {u['num']}"

    pages = [('index.html', 'All units', None)] + [(page_name(u['num']), label(u), u) for u in data['units']]
    index = [{'num': u['num'], 'name': u['name'], 'href': page_name(u['num']),
              'store': f"{store}-{u['num']}", 'count': sum(len(e['items']) for e in u['exercises'])}
             for u in data['units']]
    template = (HERE / 'template.html').read_text(encoding='utf-8')
    out = pathlib.Path(out_dir)

    def js(v):
        return json.dumps(v, ensure_ascii=False).replace('</', '<\\/')

    for name, _, u in pages:
        nav = [{'href': n, 'label': lab, 'current': n == name} for n, lab, _ in pages]
        if u is None:
            page_data = {'cover': data['cover'], 'nav': nav, 'units': [], 'parts': data['parts'],
                         'guide': data['guide'], 'index': index}
            title, key = data['cover'][0] + ' · ' + data['cover'][1], f'{store}-index'
        else:
            page_data = {'cover': [data['cover'][1], f"{label(u)} · {u['name']}"], 'nav': nav,
                         'units': [u], 'parts': [], 'guide': [], 'index': []}
            title, key = f"{label(u)} · {u['name']}", f"{store}-{u['num']}"
        page = (template.replace('__TITLE__', title)
                .replace('__STORE__', key)
                .replace('__AUDIO_DIR__', 'audio/')
                .replace('__API_URL__', api_url(env_path))
                .replace('__AUDIO__', '{}')
                .replace('__DATA__', js(page_data)))
        (out / name).write_text(page, encoding='utf-8', newline='\n')
    print('wrote', len(pages), 'pages to', out)


if __name__ == '__main__':
    main(*sys.argv[1:5])
