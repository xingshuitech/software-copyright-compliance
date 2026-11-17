# -*- coding: utf-8 -*-
"""扫描 DOCX/PDF 的 AIGC、生成器和格式痕迹。

用法:
  python scan_aigc.py 文件/目录 [更多路径...] \
      --header "软件名 V1.0" --min-lines 30 [--json]

退出码：发现硬伤返回 1，全部干净返回 0。
PDF 元数据优先用 pypdf 读取；未安装时退化为字节级检查。
"""
import sys, io, re, json, zipfile, xml.etree.ElementTree as ET
from pathlib import Path

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

EXPECTED_PARTS = {
    '[Content_Types].xml',
    '_rels/.rels',
    'docProps/core.xml',
    'word/document.xml',
    'word/_rels/document.xml.rels',
    'word/styles.xml',
    'word/settings.xml',
    'word/header1.xml',
    'word/footer1.xml',
}

GENERATOR_PATTERNS = [
    (r'AIGC', 'AIGC 标识'),
    (r'ContentProducer', 'ContentProducer'),
    (r'ProduceID', 'ProduceID'),
    (r'Apache\s*POI', 'Apache POI'),
    (r'\bWPS\b', 'WPS'),
    (r'\bKSO\b', 'KSO'),
    (r'\bICV\b', 'ICV'),
    (r'w:rsid', 'rsid 修订标识'),
    (r'cp:revision', '修订号'),
    (r'customXml', 'customXml 部件'),
    (r'docProps/app\.xml', 'app.xml 生成器部件'),
    (r'LibreOffice', 'LibreOffice'),
    (r'OpenOffice', 'OpenOffice'),
    (r'\bWriter\b', 'Writer'),
    (r'docx4j', 'docx4j'),
    (r'python-docx', 'python-docx'),
    (r'Aspose', 'Aspose'),
]

COLOR_PATTERNS = [
    (r'w:color', '字体颜色 w:color'),
    (r'w:shd', '底纹 w:shd'),
    (r'w:highlight', '高亮 w:highlight'),
    (r'w:fill', '填充 w:fill'),
    (r'w:background', '背景 w:background'),
    (r'w:themeColor', '主题色 w:themeColor'),
    (r'FFFFFF', '白色值 FFFFFF'),
]

DRAWING_PATTERNS = [
    (r'word/media/', '内嵌图片'),
    (r'w:drawing', '绘图对象'),
    (r'w:pict', '图片对象'),
    (r'v:shape', 'VML 图形'),
]

PDF_GENERATOR_BYTES = GENERATOR_PATTERNS + [
    (r'pdf:Producer', 'XMP pdf:Producer'),
    (r'xmp:CreatorTool', 'XMP xmp:CreatorTool'),
    (r'<x:xmpmeta', 'XMP 元数据块'),
]

AI_WORD = re.compile(r'\bAI\b')


def lname(tag):
    return tag.rsplit('}', 1)[-1]


def node_text(node, wanted=('t',)):
    return ''.join((el.text or '') for el in node.iter()
                   if lname(el.tag) in wanted)


def has_pagebreak(el):
    for p in el.iter():
        if lname(p.tag) != 'p':
            continue
        for child in p.iter():
            if lname(child.tag) == 'pageBreakBefore':
                return True
    return False


def match_patterns(text, patterns):
    hits = []
    for pat, label in patterns:
        if re.search(pat, text):
            hits.append(label)
    return hits


def iter_files(paths):
    seen = set()
    for raw in paths:
        p = Path(raw)
        if p.is_file():
            candidates = [p]
        elif p.is_dir():
            candidates = sorted(list(p.rglob('*.docx')) + list(p.rglob('*.DOCX')) +
                                list(p.rglob('*.pdf')) + list(p.rglob('*.PDF')))
        else:
            candidates = [p]
        for c in candidates:
            key = str(c.resolve())
            if key in seen:
                continue
            seen.add(key)
            yield c


def scan_docx(path, header=None, min_lines=None):
    res = dict(file=str(path), issues=[], warnings=[], info={})
    if min_lines is not None:
        min_lines = int(min_lines)
    try:
        with zipfile.ZipFile(str(path)) as z:
            names = z.namelist()
            data = {n: z.read(n) for n in names}
    except Exception as exc:
        res['issues'].append('无法打开 DOCX：%s' % exc)
        return res

    extra = sorted(set(names) - EXPECTED_PARTS)
    missing = sorted(EXPECTED_PARTS - set(names))
    res['info']['部件数'] = len(names)
    if extra:
        res['issues'].append('多余部件：%s' % ', '.join(extra))
    if missing:
        res['issues'].append('缺少部件：%s' % ', '.join(missing))

    for name in names:
        if not (name.endswith('.xml') or name.endswith('.rels')):
            continue
        try:
            ET.fromstring(data[name])
        except Exception as exc:
            res['issues'].append('%s XML 解析失败：%s' % (name, exc))

    for name in names:
        raw = data[name].decode('latin-1', 'replace')
        for label in match_patterns(raw, GENERATOR_PATTERNS):
            res['issues'].append('%s：%s' % (name, label))
        for label in match_patterns(raw, COLOR_PATTERNS):
            res['issues'].append('%s：%s' % (name, label))
        for label in match_patterns(raw, DRAWING_PATTERNS):
            res['issues'].append('%s：%s' % (name, label))

    # 只在可见文本里找独立的 AI，避免把哈希串或字体流字节误报。
    for name in names:
        if not (name.endswith('.xml') or name.endswith('.rels')):
            continue
        try:
            root = ET.fromstring(data[name])
        except Exception:
            continue
        text = node_text(root, ('t', 'instrText'))
        for m in AI_WORD.finditer(text):
            res['issues'].append('%s：可见文本出现独立 AI（%s）' % (
                name, text[max(0, m.start() - 12):m.end() + 12].strip()))

    if 'word/header1.xml' in data:
        try:
            hroot = ET.fromstring(data['word/header1.xml'])
            visible = node_text(hroot, ('t',))
            fields = node_text(hroot, ('instrText',))
            if 'PAGE' not in fields:
                res['issues'].append('页眉缺少 PAGE 域页码')
            if header and header.strip() not in visible:
                res['issues'].append('页眉文本不含「%s」，实际为「%s」' % (
                    header.strip(), visible.strip()))
            res['info']['页眉'] = visible.strip()
        except Exception:
            pass

    if 'word/document.xml' in data:
        try:
            doc = ET.fromstring(data['word/document.xml'])
            body = next((el for el in doc.iter() if lname(el.tag) == 'body'), None)
            if body is None:
                res['issues'].append('document.xml 缺少 w:body')
            else:
                counts = [0]
                for child in body:
                    tag = lname(child.tag)
                    if tag == 'p':
                        if has_pagebreak(child):
                            counts.append(0)
                        # 空白源码行也占版面条数，按国家要求计入每页行数。
                        counts[-1] += 1
                    elif tag == 'tbl':
                        if has_pagebreak(child):
                            counts.append(0)
                        counts[-1] += sum(1 for r in child.iter()
                                          if lname(r.tag) == 'tr')
                counts = [c for c in counts if c > 0] or [0]
                res['info']['页行数'] = '%d 页，%s，最少 %d' % (
                    len(counts), counts[:12], min(counts))
                if min_lines and min(counts) < int(min_lines):
                    res['issues'].append('存在少于 %d 行的页：%s' % (min_lines, counts))
        except Exception as exc:
            res['issues'].append('document.xml 结构检查失败：%s' % exc)

    if 'docProps/core.xml' in data:
        try:
            core = ET.fromstring(data['docProps/core.xml'])
            vals = {}
            for el in core.iter():
                if lname(el.tag) in ('creator', 'lastModifiedBy', 'title'):
                    vals[lname(el.tag)] = (el.text or '').strip()
            res['info']['文档属性'] = 'creator=%s；lastModifiedBy=%s' % (
                vals.get('creator', ''), vals.get('lastModifiedBy', ''))
        except Exception:
            pass
    return res


def scan_pdf(path):
    res = dict(file=str(path), issues=[], warnings=[], info={})
    try:
        raw = Path(str(path)).read_bytes()
    except Exception as exc:
        res['issues'].append('无法读取 PDF：%s' % exc)
        return res
    if not raw.startswith(b'%PDF'):
        res['issues'].append('不是有效的 PDF 文件')
        return res

    text = raw.decode('latin-1', 'replace')
    for label in match_patterns(text, PDF_GENERATOR_BYTES):
        res['issues'].append('字节扫描：%s' % label)

    try:
        from pypdf import PdfReader
        reader = PdfReader(str(path))
        res['info']['页数'] = len(reader.pages)
        meta = reader.metadata or {}
        for key in ('/Producer', '/Creator'):
            value = str(meta.get(key) or '').strip()
            res['info'][key] = value or '(空)'
            if value:
                res['issues'].append('PDF %s 非空：%s' % (key, value))
        for key in ('/Title', '/Author', '/Subject'):
            value = str(meta.get(key) or '').strip()
            if value:
                res['info'][key] = value
            if AI_WORD.search(value):
                res['issues'].append('PDF %s 含独立 AI' % key)
    except ImportError:
        res['warnings'].append('pypdf 未安装，PDF 元数据仅做字节级检查')
    except Exception as exc:
        res['warnings'].append('pypdf 读取失败：%s' % exc)
    return res


def print_report(res):
    status = 'FAIL' if res['issues'] else ('WARN' if res['warnings'] else 'OK')
    print('[%s] %s' % (status, res['file']))
    for k, v in res['info'].items():
        print('  . %s: %s' % (k, v))
    for item in res['issues']:
        print('  x %s' % item)
    for item in res['warnings']:
        print('  ! %s' % item)


def main(argv):
    paths = []
    header = None
    min_lines = None
    as_json = False
    i = 0
    while i < len(argv):
        a = argv[i]
        if a == '--header':
            header = argv[i + 1]; i += 2
        elif a == '--min-lines':
            min_lines = argv[i + 1]; i += 2
        elif a == '--json':
            as_json = True; i += 1
        else:
            paths.append(a); i += 1
    if not paths:
        raise SystemExit(__doc__)
    results = []
    for p in iter_files(paths):
        suffix = p.suffix.lower()
        if suffix == '.docx':
            results.append(scan_docx(p, header, min_lines))
        elif suffix == '.pdf':
            results.append(scan_pdf(p))
        else:
            results.append(dict(file=str(p), issues=['不支持的文件类型'],
                                warnings=[], info={}))
    if as_json:
        print(json.dumps(results, ensure_ascii=False, indent=2))
    else:
        for res in results:
            print_report(res)
        failed = sum(1 for r in results if r['issues'])
        print('\n共 %d 个文件，失败 %d 个' % (len(results), failed))
    return 1 if any(r['issues'] for r in results) else 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
