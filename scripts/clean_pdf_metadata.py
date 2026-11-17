# -*- coding: utf-8 -*-
"""清空 PDF 的生成器元数据，保留页面内容。

用法:
  python clean_pdf_metadata.py input.pdf [output.pdf] \
      --title "软件名 文档名" --author "记录人员" --subject "文档"

省略 output.pdf 时原地重写。依赖 pypdf：
  python -m pip install pypdf
"""
import sys, os, tempfile
from pathlib import Path


def parse_args(argv):
    cfg = dict(src=None, dst=None, title='', author='', subject='')
    pos = []
    i = 0
    while i < len(argv):
        a = argv[i]
        if a in ('--title', '--author', '--subject'):
            if i + 1 >= len(argv):
                raise SystemExit('缺少参数值：%s' % a)
            cfg[a[2:]] = argv[i + 1]
            i += 2
        else:
            pos.append(a)
            i += 1
    if not pos:
        raise SystemExit(__doc__)
    cfg['src'] = pos[0]
    cfg['dst'] = pos[1] if len(pos) > 1 else None
    if not cfg['title']:
        cfg['title'] = Path(cfg['src']).stem
    return cfg


def main(argv):
    cfg = parse_args(argv)
    try:
        from pypdf import PdfReader, PdfWriter
    except ImportError:
        raise SystemExit('缺少 pypdf，请先执行：python -m pip install pypdf')

    src = Path(cfg['src'])
    dst = Path(cfg['dst']) if cfg['dst'] else src
    reader = PdfReader(str(src))
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    # PdfWriter 默认写入 /Producer=pypdf，必须显式清空。
    writer.add_metadata({
        '/Title': cfg['title'],
        '/Author': cfg['author'],
        '/Subject': cfg['subject'],
        '/Creator': '',
        '/Producer': '',
    })
    try:
        writer.xmp_metadata = None
    except Exception:
        pass

    if dst.resolve() == src.resolve():
        fd, tmp = tempfile.mkstemp(prefix='pdfclean_', suffix='.pdf',
                                   dir=str(src.parent))
        os.close(fd)
    else:
        tmp = str(dst)
    try:
        with open(tmp, 'wb') as f:
            writer.write(f)
        if dst.resolve() == src.resolve():
            os.replace(tmp, str(dst))
    finally:
        if os.path.exists(tmp) and dst.resolve() == src.resolve():
            os.remove(tmp)

    check = PdfReader(str(dst))
    meta = check.metadata or {}
    print('已清理：%s' % dst)
    print('页数：%d；/Producer=%r；/Creator=%r；XMP=%s' % (
        len(check.pages),
        str(meta.get('/Producer') or ''),
        str(meta.get('/Creator') or ''),
        '有' if getattr(check, 'xmp_metadata', None) else '无'))
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
