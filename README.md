# GB/T 9704—2012 公文格式 Skill

> **仓库**：`GB/T 9704—2012-gongwen-format` | **描述**：GB/T 9704—2012 公文格式 Skill

[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

> 依据国标 **GB/T 9704—2012《党政机关公文格式》** 全文蒸馏构建，可重复、可校验地输出规范公文 `.docx`。

## 特性

- **四种格式**：普通文件 / 信函 / 命令(令) / 纪要
- **字体自动检测**：方正小标宋简体（已打包，一键安装）
- **自检闭环**：生成后自动对照国标逐项检查
- **学习闭环**：偏差自动记录，下次优先规避

## 快速开始

### 1. 安装依赖

```bash
pip install python-docx
```

### 2. 安装字体（首次使用）

```bash
python scripts/install_font.py fonts/方正小标宋简体.ttf
```

### 3. 生成公文

```bash
python scripts/gongwen_gen.py templates/content_schema.json output.docx
```

或自定义内容：

```bash
python scripts/gongwen_gen.py my_content.json output.docx
```

### 4. 自检

```bash
python scripts/verify.py output.docx
```

## 项目结构

```
GB/T 9704—2012 公文格式 Skill/
├── fonts/
│   └── 方正小标宋简体.ttf          # 公文红头/标题字体（已打包）
├── references/
│   ├── gbt9704-2012-spec.md        # 国标参数蒸馏
│   ├── element-layout.md           # 要素编排速查表
│   ├── font-strategy.md            # 字体检测与回退表
│   └── known_issues.md             # 自学习闭环（偏差记录）
├── scripts/
│   ├── font_check.py               # 系统字体检测
│   ├── install_font.py             # 字体安装（TTF头校验）
│   ├── gongwen_gen.py              # 核心生成器（4种格式）
│   └── verify.py                   # 生成后自检（16项）
├── templates/
│   └── content_schema.json         # 4种格式输入样例
├── README.md
├── LICENSE
└── requirements.txt
```

## 输入格式

见 `templates/content_schema.json`，支持字段：

| 字段 | 说明 | 示例 |
|---|---|---|
| `format` | 格式 | `standard`/`letter`/`order`/`minutes` |
| `issuer` | 发文机关 | `"中共云岭市委办公室"` |
| `doc_number` | 发文字号 | `"云办发〔2026〕12号"` |
| `title` | 标题 | `"关于做好…工作的通知"` |
| `body` | 正文数组 | `[{"type":"para","text":"…"}]` |
| `seal` | 印章模式 | `placeholder`/`none` |

## 自检报告

生成后自动输出 16 项国标合规检查：

```
[PASS] A4幅面 210×297
[PASS] 天头37mm
[PASS] 版心156×225
[PASS] 发文字号用六角括号〔〕
[PASS] 版头红色分隔线
[PASS] 版记分隔线(黑色)
…
通过 16/16
```

## 准确性边界

- 可控要素（页面/版心/字体/行距/分隔线/页码/版记）：~95% 合规
- 印章：占位说明（需用户在 Word 内手动盖电子章）
- 标题梯形/菱形：脚本自动居中，多行回行需用户校对

## 许可证

MIT License — 字体文件遵循方正字库个人非商业授权协议。

## 致谢

- GB/T 9704—2012《党政机关公文格式》
- [python-docx](https://github.com/python-openxml/python-docx)
