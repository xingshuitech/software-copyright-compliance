# 软著送审材料格式硬指标

依据《计算机软件著作权登记办法》第九、十、十二、十七条与中国版权保护中心《计算机软件著作权登记指南》。原文出处见 `legal-basis.md`。政策以登记系统当前显示为准。

## 交存方式

一般交存：

- 源程序：连续前 30 页 + 连续后 30 页
- 文档：任何一种文档，连续前 30 页 + 连续后 30 页
- 整个程序和文档不足 60 页的，提交全部

例外交存，选择《办法》第十二条规定的方式之一：

1. 源程序前、后各连续 30 页，机密部分用黑色宽斜线覆盖，覆盖部分不得超过交存源程序的 50%
2. 源程序连续前 10 页 + 任何部分连续 50 页
3. 目标程序前、后各连续 30 页 + 源程序任何部分连续 20 页

文档作例外交存的，参照上述方式处理。选择例外交存会产生例外交存费，具体以登记机构公布标准为准。

## 页面

- 申请文件应纵向排版，文字从左向右排列；用国际标准 A4 纸 297mm × 210mm。
- 页面尺寸 `11905 × 16840` twip。
- 上下边距 `1440` twip，左右边距 `1800` twip，页眉页脚各 `720` twip。
- 普通文档每页不少于 30 行，建议控制在 30–34 行。
- 源程序鉴别材料每页不少于 50 行，建议控制在 50–56 行。
- 行距用固定值：普通文档 `400` twip，源程序材料 `210` twip。
- 不使用首行缩进、段前段后间距、自动分页留白，否则实际行数会掉到 30 行以下。

## 页眉与页码

- 页眉固定文本为「软件全称 + 空格 + 版本」，例如 `某加密平台 V1.0 `。
- 文本末尾保留一个空格，避免 Word 把版本号和页码显示成 `V1.01`。
- 右对齐制表位放 `PAGE` 域，页码在右上角。
- 页眉不要写公司名称、人名、文档类型、日期；著作权人信息只在申请表里填。
- 页脚留空，不写版权声明、页码或生成器字样。
- 证书版本号以申请表为准；鉴别材料页眉版本号必须与申请表一致，有无 `V` 也以申请表为准。改版本号时要同步改申请表、四类材料页眉和文档属性。

页眉关键结构：

```xml
<w:t xml:space="preserve">某加密平台 V1.0 </w:t>
<w:tab/>
<w:instrText xml:space="preserve"> PAGE </w:instrText>
```

## DOCX 包结构

只保留 9 个部件：

```text
[Content_Types].xml
_rels/.rels
docProps/core.xml
word/document.xml
word/_rels/document.xml.rels
word/styles.xml
word/settings.xml
word/header1.xml
word/footer1.xml
```

禁止出现：

- `docProps/app.xml`、`customXml/`、`word/media/`
- `w:rsid*`、`cp:revision`、`w:color`、`w:shd`、`w:highlight`
- `w:fill`、`w:background`、`w:themeColor`、`FFFFFF`
- `w:drawing`、`w:pict`、`v:shape`
- 任何 `AIGC`、`ContentProducer`、`ProduceID`、Apache POI、WPS、KSO、ICV 字样

颜色指令必须彻底删除，而不是改成白色。暗色预览下出现文字白块，通常是表格边框 `w:color="808080"` 或底纹 `w:shd` 导致的。

## 文档属性

`docProps/core.xml` 只写必要字段：

```xml
<dc:title>文档标题</dc:title>
<dc:subject>文档标题</dc:subject>
<dc:creator>记录人员</dc:creator>
<cp:lastModifiedBy>记录人员</cp:lastModifiedBy>
<dcterms:created ...>...</dcterms:created>
<dcterms:modified ...>...</dcterms:modified>
```

不要写 `cp:revision`、公司模板名、软件安装路径、用户名。开发过程记录可以写真实记录人员和单位；用户手册、源程序鉴别材料、算法特色说明的页眉仍只写软件名和版本。

## 源程序鉴别材料

- 只收录登记范围内的自研源码；排除 `node_modules`、bundle、内联页面、构建产物。
- 逐文件按原顺序全文收录，不删空行，不改标识符，不改功能代码。
- 每个文件前加 `===== 相对路径 =====`，结束后加 `===== 相对路径 END =====`。
- 超长行按可视宽度软换行。中文、全角字符按 2 个半角宽度计。
- 段落数可以多于物理行数，但字符序列必须与源码逐字一致。
- 生成后用构建脚本的段落序列与源码序列逐段比较，结果必须是 `IDENTICAL`。
- 最后一个文件的结束标记位于末页，末页不得只有孤立的分隔符。
- 正文字体 8pt `Courier New`，页内不做语法高亮，不着色。

## PDF

- 上传前逐页核对页数、页眉、页码，确认没有空白页。
- LibreOffice 导出后必须清空元数据。DOCX 干净不代表 PDF 干净。
- 清理后 `/Producer`、`/Creator` 必须为空，XMP 元数据块必须不存在。
- 不要直接用 Word 或 WPS 的「另存为 PDF」交付，除非随后执行元数据清理。
