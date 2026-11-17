# -*- coding: utf-8 -*-
"""从纯文本内容生成最小 OOXML DOCX（9 部件，无生成器痕迹）。

用法:
  python build_minimal_docx.py content.txt output.docx "软件名 V1.0" \
      --title "用户手册" --creator "记录人员" \
      --min-lines 30 --max-lines 34 --font-size 24

内容文件格式:
  - 用单独一行 `---` 分页
  - `# 标题` 一级标题，`## 标题` 二级标题
  - `> 代码` 代码行
  - 以 `| ` 开头且连续排列的行为表格，`|---|` 分隔行会被跳过
  - 其余为正文行

源程序材料模式追加 --source：8pt 等宽字体、210twip 固定行距，
默认按 86 个半角显示宽度软换行，每页建议 50-56 行。
"""
import sys, io, re, html, zipfile, xml.etree.ElementTree as ET
from datetime import datetime, timezone

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'

SECTPR = ('<w:sectPr><w:headerReference r:id="rIdHeader" w:type="default"/>'
          '<w:footerReference r:id="rIdFooter" w:type="default"/>'
          '<w:pgSz w:w="11905" w:h="16840"/>'
          '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" '
          'w:header="720" w:footer="720" w:gutter="0"/>'
          '<w:pgNumType w:start="1"/><w:cols w:space="720" w:num="1"/></w:sectPr>')


def parse_args(argv):
    now = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    cfg = dict(content=None, output=None, name=None, title='文档', subject='',
               min_lines=30, max_lines=34, font_size=24, source=False,
               wrap_width=86, creator='', created=now, modified=now)
    pos = []
    i = 0
    keys = ('--title', '--subject', '--min-lines', '--max-lines', '--font-size',
            '--creator', '--created', '--modified', '--wrap-width')
    while i < len(argv):
        a = argv[i]
        if a in keys:
            if i + 1 >= len(argv):
                raise SystemExit('缺少参数值：%s' % a)
            cfg[a[2:].replace('-', '_')] = argv[i + 1]
            i += 2
        elif a == '--source':
            cfg['source'] = True
            i += 1
        else:
            pos.append(a)
            i += 1
    if len(pos) < 3:
        raise SystemExit(__doc__)
    cfg['content'], cfg['output'], cfg['name'] = pos[:3]
    for k in ('min_lines', 'max_lines', 'font_size', 'wrap_width'):
        cfg[k] = int(cfg[k])
    if not cfg['subject']:
        cfg['subject'] = cfg['title']
    if not cfg['creator']:
        cfg['creator'] = cfg['name'].split()[0] if cfg['name'].split() else ''
    return cfg


def build_parts(cfg):
    name = html.escape(cfg['name'], quote=False)
    sz = cfg['font_size']
    line_h = 210 if cfg['source'] else 400
    font = 'Courier New' if cfg['source'] else 'Times New Roman'
    creator = html.escape(cfg['creator'], quote=False)

    styles = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="%s">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="宋体" w:cs="%s"/><w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="%d" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:before="0" w:after="0" w:line="%d" w:lineRule="exact"/><w:jc w:val="left"/></w:pPr><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="宋体"/><w:sz w:val="%d"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="0" w:after="0" w:line="%d" w:lineRule="exact"/><w:jc w:val="center"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="黑体"/><w:b/><w:sz w:val="28"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="0" w:after="0" w:line="%d" w:lineRule="exact"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="黑体"/><w:b/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="0" w:after="0" w:line="%d" w:lineRule="exact"/><w:ind w:left="240"/></w:pPr><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:eastAsia="宋体"/><w:sz w:val="21"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/><w:pPr><w:tabs><w:tab w:val="right" w:pos="8300"/></w:tabs><w:spacing w:before="0" w:after="0" w:line="280" w:lineRule="exact"/></w:pPr><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="18"/></w:rPr></w:style>
</w:styles>''' % (W, font, font, font, sz, sz, line_h, line_h, font, font, sz,
                  line_h, line_h, line_h)

    header = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="%s"><w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:t xml:space="preserve">%s </w:t></w:r><w:r><w:tab/></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:hdr>''' % (W, name)

    footer = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<w:ftr xmlns:w="%s"><w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr></w:p></w:ftr>' % W)

    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:settings xmlns:w="%s"><w:zoom w:percent="100"/>'
                '<w:defaultTabStop w:val="420"/><w:updateFields w:val="true"/>'
                '<w:compat><w:compatSetting w:name="compatibilityMode" '
                'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
                '</w:compat></w:settings>' % W)

    core = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<dc:title>%s</dc:title><dc:subject>%s</dc:subject>'
            '<dc:creator>%s</dc:creator><cp:lastModifiedBy>%s</cp:lastModifiedBy>'
            '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>'
            '<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>'
            '</cp:coreProperties>' % (
                html.escape(cfg['title'], quote=False),
                html.escape(cfg['subject'], quote=False),
                creator, creator,
                html.escape(cfg['created'], quote=False),
                html.escape(cfg['modified'], quote=False)))

    types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="%s">'
             '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
             '<Default Extension="xml" ContentType="application/xml"/>'
             '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
             '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
             '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
             '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
             '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
             '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
             '</Types>' % CT)

    root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="%s">'
                 '<Relationship Id="rId1" Type="%s/officeDocument" Target="word/document.xml"/>'
                 '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
                 '</Relationships>' % (PKG_REL, R))

    doc_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="%s">'
                '<Relationship Id="rIdStyles" Type="%s/styles" Target="styles.xml"/>'
                '<Relationship Id="rIdSettings" Type="%s/settings" Target="settings.xml"/>'
                '<Relationship Id="rIdHeader" Type="%s/header" Target="header1.xml"/>'
                '<Relationship Id="rIdFooter" Type="%s/footer" Target="footer1.xml"/>'
                '</Relationships>' % (PKG_REL, R, R, R, R))
    return dict(styles=styles, header=header, footer=footer, settings=settings,
                core=core, types=types, root_rels=root_rels, doc_rels=doc_rels)


def make_para(text, style='Normal', pagebreak=False):
    ppr = '<w:pPr>'
    if style != 'Normal':
        ppr += '<w:pStyle w:val="%s"/>' % style
    if pagebreak:
        ppr += '<w:pageBreakBefore/>'
    ppr += '</w:pPr>'
    return '<w:p>%s<w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (
        ppr, html.escape(text, quote=False))


def parse_table_row(line):
    s = line.strip()
    if s.startswith('|'):
        s = s[1:]
    if s.endswith('|'):
        s = s[:-1]
    cells = [c.strip() for c in s.split('|')]
    while cells and cells[-1] == '':
        cells.pop()
    return cells


def is_table_separator(row):
    return bool(row) and all(re.fullmatch(r':?-{3,}:?', c) for c in row if c)


def make_table(rows, pagebreak=False, line_h=400):
    if not rows:
        return ''
    cols = max(len(r) for r in rows)
    width = 8300
    col_w = width // cols
    borders = ''.join('<w:%s w:val="single" w:sz="4" w:space="0"/>' % side
                      for side in ('top', 'left', 'bottom', 'right',
                                   'insideH', 'insideV'))
    tbl = ['<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/>'
           '<w:tblLayout w:type="fixed"/><w:tblBorders>%s</w:tblBorders></w:tblPr><w:tblGrid>'
           % (width, borders)]
    tbl.extend('<w:gridCol w:w="%d"/>' % col_w for _ in range(cols))
    tbl.append('</w:tblGrid>')
    for ri, row in enumerate(rows):
        tbl.append('<w:tr>')
        for ci in range(cols):
            text = row[ci] if ci < len(row) else ''
            ppr = '<w:pPr><w:spacing w:before="0" w:after="0" w:line="%d" w:lineRule="exact"/>' % line_h
            if pagebreak and ri == 0 and ci == 0:
                ppr += '<w:pageBreakBefore/>'
            ppr += '</w:pPr>'
            tbl.append('<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>'
                       '<w:p>%s<w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p></w:tc>'
                       % (col_w, ppr, html.escape(text, quote=False)))
        tbl.append('</w:tr>')
    tbl.append('</w:tbl>')
    # 表格后保留一个空段落，避免 Word 在相邻表格间自动合并。
    tbl.append(make_para(''))
    return ''.join(tbl)


def disp_width(s):
    return sum(2 if ord(ch) > 127 else 1 for ch in s)


def wrap_line(line, limit):
    s = line.replace('\t', '    ')
    if not limit or disp_width(s) <= limit:
        return [s]
    out, cur, w = [], '', 0
    for ch in s:
        cw = 2 if ord(ch) > 127 else 1
        if w + cw > limit and cur:
            out.append(cur)
            cur, w = '', 0
        cur += ch
        w += cw
    if cur:
        out.append(cur)
    return out


def build():
    cfg = parse_args(sys.argv[1:])
    content = io.open(cfg['content'], encoding='utf-8').read()
    lines = [l.rstrip() for l in content.strip('\n').split('\n')]
    if cfg['source']:
        expanded = []
        for ln in lines:
            if ln.strip() == '---':
                expanded.append(ln)
            else:
                expanded.extend(wrap_line(ln, cfg['wrap_width']))
        lines = expanded
    pages, cur = [], []
    for ln in lines:
        if ln.strip() == '---':
            pages.append(cur)
            cur = []
        else:
            cur.append(ln)
    if cur:
        pages.append(cur)
    problems = []
    for i, page in enumerate(pages):
        if not (cfg['min_lines'] <= len(page) <= cfg['max_lines']):
            problems.append('第 %d 页 %d 行（要求 %d-%d）' % (
                i + 1, len(page), cfg['min_lines'], cfg['max_lines']))
    if problems:
        print('\n'.join(problems))
        raise SystemExit('分页校验未通过，未写出文件')

    body = []
    line_h = 210 if cfg['source'] else 400
    for pi, page in enumerate(pages):
        j = 0
        first = True
        while j < len(page):
            ln = page[j]
            pb = pi > 0 and first
            if ln.startswith('|') and j + 1 < len(page) and page[j + 1].startswith('|'):
                rows = []
                while j < len(page) and page[j].startswith('|'):
                    row = parse_table_row(page[j])
                    if not is_table_separator(row):
                        rows.append(row)
                    j += 1
                body.append(make_table(rows, pb, line_h))
            else:
                if ln.startswith('# '):
                    body.append(make_para(ln[2:], 'Heading1', pb))
                elif ln.startswith('## '):
                    body.append(make_para(ln[3:], 'Heading2', pb))
                elif ln.startswith('> '):
                    body.append(make_para(ln[2:], 'Code', pb))
                else:
                    body.append(make_para(ln, 'Normal', pb))
                j += 1
            first = False

    parts = build_parts(cfg)
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document xmlns:w="%s" xmlns:r="%s"><w:body>%s%s</w:body></w:document>'
           % (W, R, ''.join(body), SECTPR))
    out = {
        '[Content_Types].xml': parts['types'],
        '_rels/.rels': parts['root_rels'],
        'docProps/core.xml': parts['core'],
        'word/document.xml': doc,
        'word/_rels/document.xml.rels': parts['doc_rels'],
        'word/styles.xml': parts['styles'],
        'word/settings.xml': parts['settings'],
        'word/header1.xml': parts['header'],
        'word/footer1.xml': parts['footer'],
    }
    for k, v in out.items():
        ET.fromstring(v.encode('utf-8'))
    with zipfile.ZipFile(cfg['output'], 'w', zipfile.ZIP_DEFLATED) as z:
        for k, v in out.items():
            z.writestr(k, v.encode('utf-8'))
    print('页数 %d，正文行 %d，每页 %s' % (
        len(pages), sum(len(p) for p in pages), [len(p) for p in pages]))
    print('已生成：%s' % cfg['output'])


if __name__ == '__main__':
    build()
