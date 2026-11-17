# 软件著作权登记材料合规检查工具

一组用于中国软件著作权登记材料的脚本与规则：生成符合格式要求的 DOCX、扫描生成痕迹与元数据、核验鉴别材料与真实源码是否逐字一致。

**定位说明**：这是真实性与格式合规检查工具，不是规避审查的工具。2026 年 3 月 15 日起，中国版权保护中心新版申请表签章页要求申请人声明「未使用 AI 开发编写代码、撰写文档或生成登记申请材料」。使用本工具不代表可以作出虚假承诺；所有材料应由申请人基于真实开发事实核实后签署。

## 功能

- `build_minimal_docx.py`：从纯文本生成 9 部件的最小 OOXML DOCX，支持页眉 `PAGE` 域、分页、表格、源码软换行、行数校验。
- `scan_aigc.py`：扫描 DOCX 包结构、XML 可解析性、生成器关键词、颜色/图片痕迹、独立 `AI` 字样、页眉页码、每页行数，以及 PDF 元数据和 XMP。
- `clean_pdf_metadata.py`：用 pypdf 逐页重写 PDF，清空 `/Producer`、`/Creator` 与 XMP。
- `verify_source.py`：按 `===== 路径 =====` 标记切分鉴别材料，与真实源码逐文件比对字符内容。
- `references/`：法规依据、格式硬指标、写作规则、申请表口径、送审前检查清单。

## 安装

作为 WorkBuddy skill 使用：把本仓库复制到用户级 skill 目录，目录名保持 `software-copyright-aigc`：

```bash
cp -r . ~/.workbuddy/skills/software-copyright-aigc
```

也可以直接在命令行使用脚本，不依赖 WorkBuddy：

```bash
python -m pip install pypdf
```

Python 3.9+ 即可。PDF 清理依赖 pypdf，其余脚本只用标准库。

## 快速开始

生成普通文档：

```bash
python scripts/build_minimal_docx.py examples/content-sample.txt manual.docx \
  "软件全称 V1.0" --title "用户手册" --creator "记录人员" \
  --min-lines 30 --max-lines 34
```

扫描送审目录：

```bash
python scripts/scan_aigc.py ./materials --header "软件全称 V1.0" --min-lines 30
```

清理 PDF 元数据：

```bash
python scripts/clean_pdf_metadata.py input.pdf output.pdf \
  --title "软件全称 用户手册" --author "记录人员" --subject "用户手册"
```

核验源程序鉴别材料：

```bash
python scripts/verify_source.py --docx code.docx --source-dir ./src
```

## 目录结构

```text
SKILL.md                      技能入口与工作流
scripts/
  build_minimal_docx.py      最小 DOCX 生成器
  scan_aigc.py               生成痕迹与格式扫描
  clean_pdf_metadata.py      PDF 元数据清理
  verify_source.py           源码逐字核验
references/
  legal-basis.md             法规依据与官方出处
  format-rules.md            格式硬指标
  writing-rules.md           工程记录式写作规则
  application-form.md        申请表逐字段口径
  checklist.md               送审前检查清单
examples/
  content-sample.txt         内容文件示例
```

## 关键规则摘要

- 申请文件纵向排版，A4 297mm × 210mm。
- 鉴别材料一般交存：源程序和文档各连续前 30 页 + 连续后 30 页；不足 60 页全部提交。
- 程序每页不少于 50 行，文档每页不少于 30 行。
- 页眉写软件全称与版本，右上角 `PAGE` 域；版本号须与申请表一致，有无 `V` 也以申请表为准。
- 主要功能和技术特点 500–1300 字。
- 签章页由申请人本人按系统文本手抄，不代抄、不预填。

完整依据见 `references/legal-basis.md`。

## 使用边界

本工具不会协助：

- 虚构开发过程、伪造提交历史、编造技术指标
- 把 AI 起草的内容伪装成自然人手写后签署承诺
- 规避登记机构的真实性核查

助手起草的材料必须由申请人逐字核实、按真实情况修改确认。是否签署承诺以及由此产生的法律责任，由申请人承担。本工具不构成法律意见；政策以中国版权保护中心登记系统当前显示为准。

## License

MIT。见 `LICENSE`。
