---
name: "GB/T 9704—2012 公文格式 Skill"
version: 1.0.0
description: 按 GB/T 9704—2012 国标生成规范党政机关公文 (.docx)，覆盖普通文件/信函/命令(令)/纪要四种格式
author: WorkBuddy
agent_created: true
read_when:
  - 生成公文 / 红头文件 / 党政机关公文
  - 生成信函 / 命令(令) / 纪要
  - GB/T 9704 / GB9704 / 公文格式
disable: false
---

# GB/T 9704—2012 公文格式 Skill

本 skill 依据国标 **GB/T 9704—2012《党政机关公文格式》** 全文蒸馏构建，参数全部来自标准正文，可重复、可校验地输出规范公文 .docx。

## 触发词
生成公文 / 红头文件 / 党政机关公文 / 公文格式 / GB9704 / 生成信函 / 发命令 / 写纪要 / 公文排版

## 工作流（必须按序执行）

1. **读已知问题**：先读 `references/known_issues.md`，避免重蹈历史偏差。
2. **检测字体**：运行 `python scripts/font_check.py`，得到本机可用字体映射表（方正小标宋是否可用）。若提示"方正小标宋未安装"，提示用户运行 `python scripts/install_font.py <FZXBSJW.TTF 路径>`（见字体说明），未装则用宋体回退。
3. **收集内容**：向用户索取公文要素（发文机关/发文字号/标题/主送/正文/成文日期/抄送/印发机关等）。可自由文本，也可直接给 JSON（见 `templates/content_schema.json`）。
4. **生成 + 自动审查修复**：运行 `python scripts/review_fix.py <content.json> <输出.docx>`。脚本自动执行：
   - 生成 .docx
   - 运行自检（对照国标 16 项检查）
   - 如有失败项，自动分析并修复（最多 3 轮循环）
   - 可自动修复：发文字号括号（`[]`→`〔〕`）、成文日期格式（`2026-7-16`→`2026年7月16日`）、页面参数偏差等
   - 无法自动修复的偏差（如缺少标题/正文）会提示用户介入
5. **交付**：将 .docx 复制到用户工作目录，用 present_files 呈现。
6. **学习闭环**：偏差自动追加到 `references/known_issues.md`，下次触发时优先回读。

> 💡 也可单独运行自检：`python scripts/verify.py <输出.docx>`（仅检查不修复）。

## 输入约定

`content.json` 字段（缺省项自动省略，不影响其余要素）：

| 字段 | 说明 | 示例 |
|---|---|---|
| format | 格式：standard/letter/order/minutes | "standard" |
| fenhao | 份号(6位) | "000001" |
| secret_level | 密级 | "秘密" |
| secret_period | 保密期限 | "1年" |
| urgency | 紧急程度 | "加急" |
| issuer | 发文机关 | "中共云岭市委办公室" |
| doc_number | 发文字号(含〔〕六角括号) | "云办发〔2026〕12号" |
| signer | 签发人(上行文) | "张三" |
| title | 标题 | "关于…的通知" |
| recipient | 主送机关(含末尾全角冒号) | "各县（市、区）委，市直各单位：" |
| body | 正文条目数组 | 见下 |
| attachments | 附件说明 | [{"name":"附件1…"}] |
| issuer_name | 发文机关署名 | "中共云岭市委办公室" |
| date | 成文日期 | "2026年7月16日" |
| annotation | 附注(自动加圆括号) | "此文公开发布" |
| copy_to | 抄送(数组) | ["XX局","XX办"] |
| printer | 印发机关 | "中共云岭市委办公室" |
| print_date | 印发日期 | "2026年7月16日" |
| seal | 印章模式：placeholder(默认)/none | "placeholder" |

`body` 每个条目：
```json
{"type":"para","text":"普通段落…"}
{"type":"h1","text":"一、第一层标题（黑体）"}
{"type":"h2","text":"（一）第二层标题（楷体）"}
{"type":"h3","text":"1.第三层（仿宋）"}
{"type":"h4","text":"（1）第四层（仿宋）"}
```

## 字体说明
- 国标推荐**方正小标宋简体**（发文机关标志/标题红色）。本机未装时用宋体回退，排版合规、红头视觉略逊。
- 官方免费个人版下载：foundertype.com（注册后0元购，个人非商用）。
- 拿到 `FZXBSJW.TTF` 后运行：`python scripts/install_font.py FZXBSJW.TTF`，自动安装到用户级字体并注册，重跑 `font_check.py` 即可接管。
- 其余三款公文标配字体（仿宋/楷体/黑体）Windows 自带，无需处理。

## 准确性边界
- 可控要素（页面/版心/字体/行距/分隔线/页码/版记/各段编排）：~95% 合规。
- 印章：仅留红色占位说明（国标要求红色实盖，属图像）；如需真实印章，用户在 Word 内手动盖电子章。
- 标题梯形/菱形排列：脚本自动居中，多行回行需用户校对断行点。
- 字体最终渲染取决于本机已装字体；`font_check.py` 负责解析实际可用族名。

## 参考文件
- `references/gbt9704-2012-spec.md`：国标参数蒸馏（权威源）
- `references/element-layout.md`：各要素编排速查表
- `references/font-strategy.md`：字体检测与回退表
- `references/known_issues.md`：自我学习闭环（偏差记录，先读后写）
- `templates/content_schema.json`：四种格式完整输入样例
