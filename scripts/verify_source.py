# -*- coding: utf-8 -*-
"""核验源程序鉴别材料与真实源码是否逐字一致。

用法:
  python verify_source.py --docx "源程序鉴别材料.docx" --source-dir "源码目录"

原理:
  鉴别材料用 `===== 相对路径 =====` / `===== 相对路径 END =====` 标记文件边界。
  脚本从 DOCX 读取段落，按标记切分文件，再与源码逐文件比对。
  超长行软换行会产生额外段落，因此比较时去掉换行但保留所有其他字符。
  不能对段落做 rstrip：软换行点上的空格属于源码原字符，去掉会造成误报。
"""
import sys, re, argparse, zipfile, xml.etree.ElementTree as ET
from pathlib import Path

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
MARKER = re.compile(r'^=====\s*(.*?)\s*=====$')


def lname(tag):
    return tag.rsplit('}', 1)[-1]


def docx_paragraphs(path):
    with zipfile.ZipFile(str(path)) as z:
        root = ET.fromstring(z.read('word/document.xml'))
    body = next((el for el in root.iter() if lname(el.tag) == 'body'), None)
    if body is None:
        raise SystemExit('DOCX 缺少 w:body')
    out = []
    for p in body.iter():
        if lname(p.tag) != 'p':
            continue
        text = ''.join((el.text or '') for el in p.iter()
                       if lname(el.tag) == 't')
        out.append(text)
    return out


def split_by_markers(lines):
    files = []
    current = None
    for idx, line in enumerate(lines):
        m = MARKER.match(line.strip())
        if not m:
            if current is not None:
                current['lines'].append(line)
            continue
        inner = m.group(1).strip()
        if inner.endswith(' END'):
            rel = inner[:-4].rstrip()
            if current is None:
                raise SystemExit('第 %d 行出现 END 标记但没有起始标记' % (idx + 1))
            current['end'] = idx
            files.append(current)
            current = None
        else:
            if current is not None:
                raise SystemExit('第 %d 行在文件未结束时出现新起始标记' % (idx + 1))
            current = dict(path=inner, start=idx, end=None, lines=[])
    if current is not None:
        raise SystemExit('文件 %s 缺少 END 标记' % current['path'])
    return files


def normalize(lines):
    # 只去掉换行差异，保留包括行尾空格在内的全部字符。
    return ''.join(lines)


def first_diff(a, b):
    n = min(len(a), len(b))
    for i in range(n):
        if a[i] != b[i]:
            return i
    return n if len(a) != len(b) else -1


def main(argv):
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--docx', required=True, help='源程序鉴别材料 DOCX')
    ap.add_argument('--source-dir', required=True, help='真实源码目录')
    ap.add_argument('--encoding', default='utf-8', help='源码文本编码，默认 utf-8')
    args = ap.parse_args(argv)

    docx_path = Path(args.docx)
    source_dir = Path(args.source_dir).resolve()
    if not docx_path.is_file():
        raise SystemExit('找不到 DOCX：%s' % docx_path)
    if not source_dir.is_dir():
        raise SystemExit('找不到源码目录：%s' % source_dir)

    files = split_by_markers(docx_paragraphs(docx_path))
    if not files:
        raise SystemExit('DOCX 中没有找到文件边界标记')

    failures = 0
    print('共 %d 个文件' % len(files))
    for item in files:
        rel = item['path']
        src = (source_dir / rel).resolve()
        if source_dir not in src.parents and src != source_dir:
            print('[FAIL] %s：路径越界 %s' % (rel, src))
            failures += 1
            continue
        if not src.is_file():
            print('[FAIL] %s：源码文件不存在' % rel)
            failures += 1
            continue
        try:
            raw = src.read_text(encoding=args.encoding)
        except Exception as exc:
            print('[FAIL] %s：读取失败 %s' % (rel, exc))
            failures += 1
            continue

        a = normalize(raw.splitlines())
        b = normalize(item['lines'])
        if a == b:
            print('[OK]   %s（%d 字符）' % (rel, len(a)))
            continue
        failures += 1
        pos = first_diff(a, b)
        lo = max(0, pos - 30)
        print('[FAIL] %s：源码 %d 字符，材料 %d 字符，首个差异位于 %d' % (
            rel, len(a), len(b), pos))
        print('       源码片段：%r' % a[lo:pos + 30])
        print('       材料片段：%r' % b[lo:pos + 30])

    print('\n%d/%d 个文件一致' % (len(files) - failures, len(files)))
    return 1 if failures else 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
