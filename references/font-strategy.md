# 字体检测与回退表

## 国标公文四款字体 → 本机族名映射
`font_check.py` 按下表优先级解析实际可用族名。第一项命中即用，否则用回退。

| 国标字体 | 优先族名(按命中顺序) | 回退 | 文件线索 |
|---|---|---|---|
| 方正小标宋简体 | `方正小标宋简体` / `FZXiaoBiaoSong-B05S` / `FZXiaoBiaoSong-B05` | 宋体 SimSun | FZXBSJW.TTF |
| 仿宋 | `仿宋` / `FangSong` / `FangSong_GB2312` / `STFangsong`(华文仿宋) | — | simfang.ttf |
| 楷体 | `楷体` / `KaiTi` / `KaiTi_GB2312` / `STKaiti`(华文楷体) | — | simkai.ttf |
| 黑体 | `黑体` / `SimHei` / `Microsoft YaHei`(末选) | — | simhei.ttf |
| 宋体(页码/回退) | `宋体` / `SimSun` / `NSimSun` / `STSong`(华文宋体) | — | simsun.ttc |

## 安装方正小标宋简体（未装时）
1. 官方免费个人版：https://www.foundertype.com → 注册 → 搜"方正小标宋简体" → 0元购 → 下载 `FZXBSJW.TTF`（个人非商用免费）。
2. 一键安装：`python scripts/install_font.py <FZXBSJW.TTF 路径>`
   - 复制到用户级字体目录 `%LOCALAPPDATA%\Microsoft\Windows\Fonts`
   - 写注册表 `HKCU\Software\Microsoft\Windows NT\CurrentVersion\Fonts`（无需管理员）
   - 广播字体变更通知
3. 重跑 `python scripts/font_check.py` 确认命中。

## 安全说明
- 不从第三方软件站（华军/PC下载等）下载——它们捆绑安装器/广告插件，存在风险。
- 仅接受来自官方 foundertype.com 或可信来源的 .ttf；`install_font.py` 会校验文件头（TTF: `00 01 00 00`；OTF: `4F 54 54 4F`）再安装。

## 字体缺失影响
- 小标宋缺失：红头/标题用宋体替代，**排版合规、视觉略逊**，不影响字号/位置/间距。
- 仿宋/楷体/黑体缺失（极少见）：影响对应要素视觉，应补装。
