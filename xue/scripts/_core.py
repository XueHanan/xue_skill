"""Shared, standard-library-only implementation. Canonical copy: xue/scripts.

xue-plus receives a verified copy through tools/sync_shared.py. Semantic review
records are evidence-bound attestations, not an automated proof of equivalence.
"""
from __future__ import annotations

import argparse
from collections import Counter
import hashlib
import json
from pathlib import Path
import re
import sys
import xml.etree.ElementTree as ET
import zipfile

VERSION = 2
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
MARKER = re.compile(r'\[(?:EQ|OBJ)_p\d+_\d+\]')
CITE = re.compile(r'\[\d+(?:\s*[,，\-–]\s*\d+)*\]')
NUM = re.compile(r'(?<![A-Za-z])[-+−]?\d+(?:\.\d+)?(?:[eE][+\-]?\d+)?[%％]?')
LATIN = re.compile(r'(?<![A-Za-z0-9_])[A-Za-z][A-Za-z0-9_+./\-]*')
QUANTITY = re.compile(r'([-+−]?\d+(?:\.\d+)?(?:[eE][+\-]?\d+)?)\s*'
                      r'(个百分点|毫秒|秒|分钟|小时|千米|厘米|毫米|米|赫兹|'
                      r'GHz|MHz|kHz|Hz|fps|FPS|ms|μs|us|ns|min|px|'
                      r'GiB|MiB|GB|MB|KB|kg|mg|km|cm|mm|mA|kW|dB|'
                      r's|h|m|g|W|J|V|A|K|℃|°C|%|％)(?![A-Za-z])')
REF = re.compile(r'(?:图|表)\s*\d+(?:\s*[-–—.]\s*\d+)?|'
                 r'(?:公式|式)\s*[（(]\s*\d+(?:\s*[-–—.]\s*\d+)?\s*[）)]')
HEDGES = re.compile(r'尚待验证|尚未|有待|初步|可能|或许|大概|或有|仅在|仅适用于|'
                    r'不支持|不适用|不得|不能|无法|未能|没有|并非|不再|不|未|无|至少|至多')
ASPECTS = ('claims', 'bindings', 'conditions', 'steps', 'terminology')


class Invalid(ValueError):
    pass


def digest(data):
    return hashlib.sha256(data).hexdigest()


def file_hash(path):
    return digest(Path(path).read_bytes())


def save_json(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + '.tmp')
    tmp.write_text(json.dumps(value, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    tmp.replace(path)


def load_json(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


def strip_markers(text):
    return MARKER.sub('', text)


def outline_of(p, styles):
    props = p.find(W + 'pPr')
    sid = None
    if props is not None:
        direct = props.find(W + 'outlineLvl')
        if direct is not None:
            return int(direct.get(W + 'val', '9'))
        ps = props.find(W + 'pStyle')
        sid = ps.get(W + 'val') if ps is not None else None
    seen = set()
    while sid and sid not in seen:
        seen.add(sid)
        st = styles.get(sid)
        if st is None:
            break
        level = st.find('./' + W + 'pPr/' + W + 'outlineLvl')
        if level is not None:
            return int(level.get(W + 'val', '9'))
        base = st.find(W + 'basedOn')
        sid = base.get(W + 'val') if base is not None else None
    return None


def extract_docx(path):
    rows, objects, tables = [], [], []
    with zipfile.ZipFile(path) as z:
        body = ET.fromstring(z.read('word/document.xml')).find(W + 'body')
        if body is None:
            raise Invalid('DOCX 没有正文 body')
        styles = {}
        if 'word/styles.xml' in z.namelist():
            styles = {s.get(W + 'styleId'): s for s in ET.fromstring(z.read('word/styles.xml'))}
        for child in body:
            if child.tag == W + 'tbl':
                tables.append({'after_paragraph': len(rows) - 1,
                               'text': '\n'.join(''.join(c.itertext()) for c in child.findall('.//' + W + 't')),
                               'xml_sha256': digest(ET.tostring(child))})
                continue
            if child.tag != W + 'p':
                continue
            pid, chunks, ids, frozen = len(rows), [], [], []
            def visit(el):
                if el.tag in (M + 'oMath', M + 'oMathPara', W + 'object', W + 'drawing', W + 'pict'):
                    is_eq = el.tag.startswith(M) or any('Equation' in v or 'MathType' in v
                                  for node in el.iter() for v in node.attrib.values())
                    oid = f"{'EQ' if is_eq else 'OBJ'}_p{pid:06d}_{len(ids)+1:02d}"
                    ids.append(oid)
                    chunks.append('[' + oid + ']')
                    objects.append({'id': oid, 'paragraph': pid, 'kind': 'equation' if is_eq else 'unsupported_object',
                                    'xml_sha256': digest(ET.tostring(el)),
                                    'display_text': ''.join(n.text or '' for n in el.iter() if n.tag == M + 't')})
                    if not is_eq:
                        frozen.append('embedded_object')
                    return
                if el.tag in (W + 'ins', W + 'del', W + 'moveFrom', W + 'moveTo'):
                    frozen.append('tracked_changes')
                if el.tag in (W + 'fldChar', W + 'instrText', W + 'fldSimple'):
                    frozen.append('word_field')
                if el.tag in (W + 'footnoteReference', W + 'endnoteReference'):
                    frozen.append('note_reference')
                if el.tag == W + 't':
                    chunks.append(el.text or '')
                elif el.tag == W + 'tab':
                    chunks.append('\t')
                elif el.tag in (W + 'br', W + 'cr'):
                    chunks.append('\n')
                elif el.tag != W + 'pPr':
                    for node in el:
                        visit(node)
            visit(child)
            text = ''.join(chunks).strip()
            level = outline_of(child, styles)
            prop = child.find('./' + W + 'pPr/' + W + 'pStyle')
            style = prop.get(W + 'val', '') if prop is not None else ''
            kind = 'heading' if level is not None and level < 9 else 'prose'
            if frozen:
                kind = 'frozen'
            if re.match(r'^(?:图|表)\s*\d+[-.\d]*\s+', text) or style.lower().startswith(('caption', 'toc', 'bibliography')):
                kind = 'frozen'
                frozen.append('caption_or_reference_style')
            rows.append({'idx': pid, 'style': style, 'outline': level, 'kind': kind,
                         'text': text, 'objects': ids, 'frozen_reasons': sorted(set(frozen))})
    return rows, objects, tables


def prepare(source, workdir, selected_range=None, terms=(), mode='auto', whole_document=False):
    source, workdir = Path(source).resolve(), Path(workdir).resolve()
    if source.is_relative_to(workdir):
        raise Invalid('输入文件不能位于任务输出目录内，请使用独立工作目录以免覆盖原文')
    if (workdir / 'task.json').exists():
        raise Invalid('该目录已有任务基线；请使用新目录，不要覆盖来源映射')
    if source.suffix.lower() == '.docx':
        if selected_range is None and not whole_document:
            raise Invalid('DOCX 必须给 --range，或在已获全文处理授权时给 --whole-document；先用 inspect_docx.py 查看段号')
        rows, objects, tables = extract_docx(source)
    elif source.suffix.lower() in ('.txt', '.md'):
        text = source.read_text(encoding='utf-8-sig').replace('\r\n', '\n')
        rows = [{'idx': i, 'style': '', 'outline': None, 'kind': 'heading' if re.match(r'^#{1,6}\s', t.strip()) else 'prose', 'text': t.strip(),
                 'objects': [], 'frozen_reasons': []}
                for i, t in enumerate(re.split(r'\n[ \t]*\n+', text.strip())) if t.strip()]
        objects, tables = [], []
    else:
        raise Invalid('输入仅支持 .docx、UTF-8 .txt 或 .md')
    if any('[M]' in r['text'] for r in rows):
        raise Invalid('发现不具唯一身份的旧 [M] 标记，请从原始文件重新提取')
    if selected_range:
        lo, hi = selected_range
        if lo < 0 or hi < lo or hi >= len(rows):
            raise Invalid('段落范围越界；编号从 0 开始')
        ids = [r['idx'] for r in rows if lo <= r['idx'] <= hi and r['text']]
    else:
        ids = [r['idx'] for r in rows if r['kind'] == 'prose' and r['text']]
    if not ids:
        raise Invalid('没有可处理段落；请核对范围或使用原始正文文件')
    workdir.mkdir(parents=True, exist_ok=True)
    save_json(workdir / 'paragraphs.json', rows)
    save_json(workdir / 'objects.json', {'objects': objects, 'read_only_tables': tables})
    manifest = {'schema_version': VERSION, 'source': str(source),
                'paragraphs_sha256': file_hash(workdir / 'paragraphs.json'),
                'objects_sha256': file_hash(workdir / 'objects.json'), 'selected_ids': ids,
                'protected_terms': sorted(set(t for t in terms if t)), 'mode': mode,
                'indexing': 'zero-based top-level body paragraphs; tables separate',
                'scope_note': 'Only selected_ids may change. Non-prose rows are frozen.'}
    save_json(workdir / 'task.json', manifest)
    (workdir / 'marked.txt').write_text('\n\n'.join(f"<{r['idx']}> {r['text']}" for r in rows), encoding='utf-8')
    (workdir / 'structure.tsv').write_text('idx\tkind\tstyle\ttext\n' + '\n'.join(
        f"{r['idx']}\t{r['kind']}\t{r['style']}\t{r['text'].replace(chr(10), ' ').replace(chr(9), ' ')}" for r in rows), encoding='utf-8')
    draft = ['# 润色对照稿', '', '范围与段号以 task.json 为准。未修改段落保留原文；附录由验证器生成。', '']
    for r in rows:
        if r['idx'] in ids:
            quote = '\n'.join('> ' + line for line in r['text'].split('\n'))
            draft.extend([f"#### `<{r['idx']}>`", '', '**原文**', quote, '', '**改后**', quote, '', '**改动**', '- 保留原文。', ''])
    (workdir / 'comparison.md').write_text('\n'.join(draft), encoding='utf-8')
    return manifest


def load_task(workdir):
    workdir = Path(workdir)
    task = load_json(workdir / 'task.json')
    if not isinstance(task, dict) or task.get('schema_version') != VERSION:
        raise Invalid('任务基线版本不支持，请用 prepare.py 从原文件重新生成')
    for name in ('paragraphs', 'objects'):
        if file_hash(workdir / (name + '.json')) != task.get(name + '_sha256'):
            raise Invalid(name + '.json 与提取基线不一致；禁止修改基线使检查通过')
    rows = load_json(workdir / 'paragraphs.json')
    if (not isinstance(rows, list) or not rows or
        any(not isinstance(r, dict) or type(r.get('idx')) is not int or not isinstance(r.get('text'), str)
            or r.get('kind') not in ('heading', 'prose', 'frozen') or not isinstance(r.get('objects'), list) for r in rows)
        or len({r['idx'] for r in rows}) != len(rows)):
        raise Invalid('段落基线结构错误')
    ids = task.get('selected_ids')
    available = {r['idx'] for r in rows if r['text']}
    if not isinstance(ids, list) or not ids or len(ids) != len(set(ids)) or any(type(i) is not int or i not in available for i in ids):
        raise Invalid('selected_ids 为空、重复或含未知/空段落')
    if not isinstance(task.get('protected_terms'), list) or any(not isinstance(t, str) or not t for t in task['protected_terms']):
        raise Invalid('protected_terms 必须是非空术语字符串列表')
    return task, rows


def parse_comparison(md):
    lines, entries, current, i, fence = md.splitlines(), [], None, 0, None
    heading = re.compile(r'^####\s+(?:`<(\d+)>`|<(\d+)>)(?:\s.*)?$')
    while i < len(lines):
        line = lines[i]
        fm = re.match(r'^\s*(`{3,}|~{3,})', line)
        if fm:
            mark = fm.group(1)
            if fence is None:
                fence = (mark[0], len(mark))
            elif mark[0] == fence[0] and len(mark) >= fence[1]:
                fence = None
            i += 1
            continue
        if fence:
            i += 1
            continue
        match = heading.match(line.strip())
        if match:
            current = {'id': int(match.group(1) or match.group(2)), 'blocks': []}
            entries.append(current)
            i += 1
            continue
        if re.match(r'^####\s', line):
            raise Invalid(f'无法识别段号：第 {i+1} 行')
        if re.match(r'^#{1,3}\s', line):
            current = None
            if re.match(r'^##\s*附录', line):
                raise Invalid('附录由验证器生成到 appendix.md；请移除手写附录，避免正文与附录错配')
        label = re.fullmatch(r'\*\*(原文|原句|改后|改句)\*\*[:：]?', line.strip())
        if label:
            if current is None:
                raise Invalid(f'文本标签不属于任何段落：第 {i+1} 行')
            i += 1
            while i < len(lines) and not lines[i].strip():
                i += 1
            buf = []
            while i < len(lines) and re.match(r'^\s*>', lines[i]):
                payload = re.sub(r'^\s*> ?', '', lines[i], count=1)
                buf.append(payload)
                i += 1
            text = '\n'.join(buf)
            if not text.strip():
                raise Invalid(f"<{current['id']}> {label.group(1)}为空或缺少引用块")
            current['blocks'].append((label.group(1), text))
            continue
        i += 1
    if fence:
        raise Invalid('未闭合代码块')
    if not entries:
        raise Invalid('未解析到有效段落；不能把零条目视为通过')
    if len({e['id'] for e in entries}) != len(entries):
        raise Invalid('段号重复')
    return entries


def compose(entries, task, rows):
    source = {r['idx']: r['text'] for r in rows}
    selected, seen = set(task['selected_ids']), {e['id'] for e in entries}
    if selected != seen:
        raise Invalid(f'范围不完整或越界：缺少 {sorted(selected-seen)}；多出 {sorted(seen-selected)}')
    revised = dict(source)
    for entry in entries:
        pid, blocks = entry['id'], entry['blocks']
        labels = [x[0] for x in blocks]
        if labels == ['原文', '改后']:
            if blocks[0][1] != source[pid]:
                raise Invalid(f'<{pid}> 原文与来源逐字不一致（不忽略空白或公式标记）')
            revised[pid] = blocks[1][1]
        elif labels and len(labels) % 2 == 0 and labels == ['原句', '改句'] * (len(labels)//2):
            spans = []
            for j in range(0, len(blocks), 2):
                old, new = blocks[j][1], blocks[j+1][1]
                if source[pid].count(old) != 1:
                    raise Invalid(f'<{pid}> 原句缺失或出现多次；请改用全段对照以消除歧义')
                at = source[pid].index(old)
                spans.append((at, at + len(old), new))
            spans.sort()
            if any(a[1] > b[0] for a, b in zip(spans, spans[1:])):
                raise Invalid(f'<{pid}> 句级修改范围重叠')
            for start, stop, new in reversed(spans):
                revised[pid] = revised[pid][:start] + new + revised[pid][stop:]
        else:
            raise Invalid(f'<{pid}> 标签不成对、顺序错误，或全段与句级模式混用')
    return revised


def invariants(a, b, terms):
    issues, flags = [], []
    if '[M]' in a or '[M]' in b:
        issues.append('不支持有歧义的旧 [M] 标记')
    if MARKER.findall(a) != MARKER.findall(b):
        issues.append('公式/对象 ID 的内容、数量或顺序改变')
    clean_a, clean_b = strip_markers(a), strip_markers(b)
    if Counter(CITE.findall(clean_a)) != Counter(CITE.findall(clean_b)):
        issues.append('引用编号及出现次数改变')
    if Counter(NUM.findall(CITE.sub('', clean_a))) != Counter(NUM.findall(CITE.sub('', clean_b))):
        issues.append('数字、百分比或出现次数改变')
    if Counter(QUANTITY.findall(clean_a)) != Counter(QUANTITY.findall(clean_b)):
        issues.append('数值与物理单位的对应关系改变')
    refs = lambda t: Counter(re.sub(r'\s+', '', r).replace('–', '-').replace('—', '-') for r in REF.findall(t))
    if refs(clean_a) != refs(clean_b):
        issues.append('图表/公式引用改变')
    # Latin tokens in Chinese prose include acronyms, units and single-letter variables.
    if re.search(r'[\u3400-\u9fff]', clean_a) and Counter(LATIN.findall(clean_a)) != Counter(LATIN.findall(clean_b)):
        issues.append('英文技术词、单位或单字母变量改变')
    for term in terms:
        if a.count(term) != b.count(term):
            issues.append(f'受保护术语改变：{term}')
    if HEDGES.findall(a) != HEDGES.findall(b):
        flags.append('否定、限定词、适用条件或其顺序发生变化，必须核对对应命题')
    if CITE.findall(a) != CITE.findall(b):
        flags.append('引用顺序变化，必须核对论断—引文关系')
    if NUM.findall(CITE.sub('', clean_a)) != NUM.findall(CITE.sub('', clean_b)):
        flags.append('数值顺序变化，必须核对指标—数值关系')
    if a != b:
        flags.append('需复核含义、指标/引用归属、适用条件、步骤和术语；集合守恒不是语义证明')
    return issues, flags


def content_units(text):
    text = strip_markers(text)
    return len(re.findall(r'[\u3400-\u9fff]', text)) + len(re.findall(r'[A-Za-z][A-Za-z0-9_\-]*', text))


def metrics(texts):
    sents = [s for text in texts for s in re.split(r'[。！？]|(?<=[.!?])\s+', text) if s.strip()]
    lengths = [len(re.sub(r'\s+', '', strip_markers(s))) for s in sents]
    joined = '\n'.join(texts)
    return {'paragraphs': len(texts), 'characters': sum(len(re.sub(r'\s+', '', strip_markers(t))) for t in texts),
            'sentences': len(lengths), 'average_sentence_length': round(sum(lengths)/len(lengths), 2) if lengths else 0,
            'over_60': sum(n > 60 for n in lengths), 'over_80': sum(n > 80 for n in lengths),
            'steps': len(re.findall(r'首先|其次|然后|最后|步骤\s*\d+', joined)),
            'references': len(REF.findall(joined)), 'equation_objects': len(re.findall(r'\[EQ_p\d+_\d+\]', joined))}


def review_draft(mdpath, workdir):
    task, rows = load_task(workdir)
    md = Path(mdpath).read_text(encoding='utf-8-sig')
    revised = compose(parse_comparison(md), task, rows)
    result = {'schema_version': VERSION, 'comparison_sha256': file_hash(mdpath),
              'task_sha256': file_hash(Path(workdir)/'task.json'),
              'reviewer': 'agent', 'review_type': 'semantic_attestation_not_proof', 'entries': {}}
    for r in rows:
        pid = r['idx']
        if pid in task['selected_ids'] and r['text'] != revised[pid]:
            issues, flags = invariants(r['text'], revised[pid], task['protected_terms'])
            result['entries'][str(pid)] = {'original_sha256': digest(r['text'].encode()),
                'revised_sha256': digest(revised[pid].encode()), 'automatic_issues': issues, 'review_flags': flags,
                'checks': {k: {'result': 'pending', 'evidence': ''} for k in ASPECTS}}
    return result


def validate(mdpath, workdir, reviewpath=None, before_final=None, ratio=0.15):
    mdpath, workdir = Path(mdpath), Path(workdir)
    report = {'schema_version': VERSION, 'status': 'failed', 'errors': [], 'pending_review': [],
              'review_flags': {}, 'outputs': None,
              'meaning_of_pass': '自动检查通过且修改段具有当前版本的语义复核记录；不是技术正确性证明'}
    try:
        task, rows = load_task(workdir)
        report.update({'comparison_sha256': file_hash(mdpath), 'task_sha256': file_hash(workdir/'task.json'),
                       'selected_ids': task['selected_ids']})
        revised = compose(parse_comparison(mdpath.read_text(encoding='utf-8-sig')), task, rows)
        reviewpath = Path(reviewpath) if reviewpath else workdir/'review.json'
        review = load_json(reviewpath) if reviewpath.exists() else None
        if review is not None and (not isinstance(review, dict) or not isinstance(review.get('entries'), dict)):
            raise Invalid('语义复核记录结构错误')
        if review is not None and (review.get('schema_version') != VERSION or
               review.get('comparison_sha256') != report['comparison_sha256'] or
               review.get('task_sha256') != report['task_sha256']):
            report['pending_review'].append('复核记录已过期或版本不匹配；请重新核对，不要只替换摘要')
            review = None
        if review is not None:
            report['review_sha256'] = file_hash(reviewpath)
        changed = []
        for r in rows:
            pid, old = r['idx'], r['text']
            if pid not in task['selected_ids']:
                continue
            new = revised[pid]
            issues, flags = invariants(old, new, task['protected_terms'])
            report['errors'].extend(f'<{pid}> {e}' for e in issues)
            if r['kind'] != 'prose' and old != new:
                report['errors'].append(f"<{pid}> {r['kind']} 段落只读，不允许改写")
            if old == new:
                continue
            changed.append(pid)
            report['review_flags'][str(pid)] = flags
            record = (review or {}).get('entries', {}).get(str(pid), {})
            if not isinstance(record, dict) or not isinstance(record.get('checks', {}), dict):
                raise Invalid(f'<{pid}> 语义复核记录结构错误')
            if record.get('original_sha256') != digest(old.encode()) or record.get('revised_sha256') != digest(new.encode()):
                report['pending_review'].append(f'<{pid}> 缺少当前文本的语义复核记录')
                continue
            for aspect in ASPECTS:
                check = record.get('checks', {}).get(aspect, {})
                if not isinstance(check, dict) or check.get('result') != 'preserved' or not isinstance(check.get('evidence'), str) or not check['evidence'].strip():
                    report['pending_review'].append(f'<{pid}> {aspect} 尚未复核或缺少具体依据')
        if before_final:
            if not 0 <= ratio <= 1:
                raise Invalid('ratio 必须在 0–1 之间')
            before = compose(parse_comparison(Path(before_final).read_text(encoding='utf-8-sig')), task, rows)
            for pid in task['selected_ids']:
                ua, ub = content_units(before[pid]), content_units(revised[pid])
                if before[pid] != revised[pid] and (ua == 0 or abs(ub/ua - 1) > ratio):
                    report['errors'].append(f'<{pid}> 本次终检前后文本单位数变化超过 ±{ratio:.0%}；回退该段')
            report['before_final_sha256'] = file_hash(before_final)
            report['final_polish_ratio_limit'] = ratio
        report['changed_ids'] = changed
        report['kept_ids'] = [i for i in task['selected_ids'] if i not in changed]
        selected_prose = [r for r in rows if r['idx'] in task['selected_ids'] and r['kind'] == 'prose']
        report['stats'] = {'scope': [r['idx'] for r in selected_prose],
                           'original': metrics([r['text'] for r in selected_prose]),
                           'revised': metrics([revised[r['idx']] for r in selected_prose])}
        report['status'] = 'failed' if report['errors'] else 'needs_review' if report['pending_review'] else 'passed'
        if report['status'] == 'passed':
            version = digest((report['comparison_sha256'] + report['task_sha256'] + report.get('review_sha256', '')).encode())
            dest = workdir/'validated'/version
            dest.mkdir(parents=True, exist_ok=True)
            full = dest/'full_revised.txt'
            full.write_text('\n\n'.join(f"<{r['idx']}> {revised[r['idx']]}" for r in rows), encoding='utf-8')
            save_json(dest/'stats.json', report['stats'])
            appendix = ['# 自动核对附录', '', '| 段号 | 状态 | 原/改引用 | 原/改数字 | 对象 |', '| --- | --- | --- | --- | --- |']
            for r in rows:
                pid, a = r['idx'], r['text']
                if pid not in task['selected_ids']:
                    continue
                b = revised[pid]
                appendix.append(f"| {pid} | {'修改并复核' if pid in changed else '原样保留'} | {len(CITE.findall(a))}/{len(CITE.findall(b))} | {len(NUM.findall(CITE.sub('',strip_markers(a))))}/{len(NUM.findall(CITE.sub('',strip_markers(b))))} | {', '.join(r['objects']) or '无'} |")
            appendix.extend(['', '数字与引用计数相等不代表归属正确。修改段另有绑定当前版本的语义复核记录。',
                             '只统计 task.json 所选正文段；标题、冻结段及范围外内容不计入语体指标。'])
            (dest/'appendix.md').write_text('\n'.join(appendix)+'\n', encoding='utf-8')
            report['outputs'] = {n: str((dest/n).resolve()) for n in ('full_revised.txt','stats.json','appendix.md')}
            if review is not None:
                save_json(dest/'review.json', review)
            save_json(dest/'validation.json', report)
    except (Invalid, OSError, ValueError, KeyError, TypeError, zipfile.BadZipFile, ET.ParseError) as exc:
        report['status'] = 'failed'
        report['outputs'] = None
        report['errors'].append(str(exc))
    save_json(workdir/'validation.json', report)
    return report


def prepare_cli():
    ap = argparse.ArgumentParser(description='准备唯一公式标记、处理清单与对照稿；仅需标准库')
    ap.add_argument('source')
    ap.add_argument('workdir')
    ap.add_argument('--range', nargs=2, type=int, metavar=('START','END'))
    ap.add_argument('--whole-document', action='store_true', help='仅当用户已授权处理整个 DOCX 时使用')
    ap.add_argument('--term', action='append', default=[])
    ap.add_argument('--terms-file', help='UTF-8 文件，每行一个受保护术语')
    ap.add_argument('--mode', choices=('auto','intro','theory','method','experiment','abstract','conclusion'), default='auto')
    args = ap.parse_args()
    try:
        terms = args.term + (Path(args.terms_file).read_text(encoding='utf-8-sig').splitlines() if args.terms_file else [])
        task = prepare(args.source, args.workdir, args.range, [t.strip() for t in terms if t.strip()], args.mode, args.whole_document)
        print(f"prepared={len(task['selected_ids'])} workdir={Path(args.workdir).resolve()}")
    except (Invalid, OSError, ValueError, KeyError, zipfile.BadZipFile, ET.ParseError) as exc:
        print('ERROR:', exc, file=sys.stderr)
        raise SystemExit(1)


def validate_cli():
    ap = argparse.ArgumentParser(description='统一验证；0=通过，1=失败，2=尚待语义复核')
    ap.add_argument('md')
    ap.add_argument('workdir')
    ap.add_argument('--review')
    ap.add_argument('--before-final', help='可选：终检前对照稿快照，用于检查这一次微调')
    ap.add_argument('--ratio', type=float, default=0.15, help='仅用于 --before-final，默认 0.15')
    args = ap.parse_args()
    report = validate(args.md, args.workdir, args.review, args.before_final, args.ratio)
    print(f"status={report['status']} errors={len(report['errors'])} pending_review={len(report['pending_review'])}")
    for e in report['errors']:
        print('ERROR:', e)
    for e in report['pending_review']:
        print('REVIEW:', e)
    print('report=' + str((Path(args.workdir)/'validation.json').resolve()))
    raise SystemExit({'passed':0,'failed':1,'needs_review':2}[report['status']])


def review_cli():
    ap = argparse.ArgumentParser(description='生成尚未签署的语义复核模板；不得把 pending 批量替换为 preserved')
    ap.add_argument('md')
    ap.add_argument('workdir')
    ap.add_argument('--out')
    args = ap.parse_args()
    try:
        path = Path(args.out) if args.out else Path(args.workdir)/'review.json'
        if path.exists():
            raise Invalid('复核文件已存在；请指定新的 --out，保留旧版本并重新复核')
        save_json(path, review_draft(args.md, args.workdir))
        print('review_template=' + str(path.resolve()))
    except (Invalid, OSError, ValueError, KeyError) as exc:
        print('ERROR:', exc, file=sys.stderr)
        raise SystemExit(1)


def inspect_cli():
    ap = argparse.ArgumentParser(description='只读查看 DOCX 顶层正文段号、类型与对象；不建立全篇修改授权')
    ap.add_argument('docx')
    args = ap.parse_args()
    try:
        rows, objects, tables = extract_docx(args.docx)
        print('idx\tkind\ttext')
        for r in rows:
            print(f"{r['idx']}\t{r['kind']}\t{r['text'].replace(chr(10), ' ').replace(chr(9), ' ')[:160]}")
        print(f'objects={len(objects)} read_only_tables={len(tables)}')
    except (Invalid, OSError, KeyError, ValueError, zipfile.BadZipFile, ET.ParseError) as exc:
        print('ERROR:', exc, file=sys.stderr)
        raise SystemExit(1)
