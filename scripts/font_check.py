# -*- coding: utf-8 -*-
"""检测本机公文四款字体可用性，输出可用族名映射。
用法: python font_check.py   (打印并写 font_map.json 到脚本目录)
可被 gongwen_gen.py 导入: from font_check import resolve_fonts
"""
import os, sys, json, winreg

# 国标字体 -> 优先族名(按命中顺序)，第一项为首选
TARGETS = {
    "xiaobiaosong": ["方正小标宋简体", "FZXiaoBiaoSong-B05S", "FZXiaoBiaoSong-B05",
                     "FZShuTi", "方正小标宋_GBK"],
    "fangsong":     ["仿宋", "FangSong", "FangSong_GB2312", "STFangsong", "华文仿宋"],
    "kaiti":        ["楷体", "KaiTi", "KaiTi_GB2312", "STKaiti", "华文楷体"],
    "heiti":        ["黑体", "SimHei", "Microsoft YaHei"],
    "songti":       ["宋体", "SimSun", "NSimSun", "STSong", "华文宋体"],
}


def _installed_font_names():
    """从注册表读取所有已注册字体族名(显示名)。"""
    names = set()
    for hive, flag in ((winreg.HKEY_LOCAL_MACHINE, winreg.KEY_READ | winreg.KEY_WOW64_64KEY),
                       (winreg.HKEY_CURRENT_USER, winreg.KEY_READ)):
        try:
            key = winreg.OpenKey(hive, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts", 0, flag)
        except FileNotFoundError:
            continue
        i = 0
        while True:
            try:
                name, value, vtype = winreg.EnumValue(key, i)
            except OSError:
                break
            # 去掉 "(TrueType)" 等后缀
            for suf in (" (TrueType)", " (OpenType)", " (TrueType,Italic)"):
                if name.endswith(suf):
                    name = name[: -len(suf)]
            names.add(name.strip())
            i += 1
        winreg.CloseKey(key)
    return names


def resolve_fonts():
    """返回各字体最终使用的族名(含回退)。"""
    installed = _installed_font_names()

    def available(cand):
        if cand in installed:
            return True
        # 注册表名常为 "SimSun & NSimSun (TrueType)" 形式, 需分词匹配
        for nm in installed:
            s = nm
            for sep in (" & ", "&", "、", "(", "（"):
                s = s.replace(sep, "|")
            if cand in [t.strip() for t in s.split("|")]:
                return True
        return False

    resolved = {}
    for key, cands in TARGETS.items():
        chosen = next((c for c in cands if available(c)), None)
        resolved[key] = chosen
    # 回退：小标宋未装 -> 宋体
    if not resolved["xiaobiaosong"]:
        resolved["xiaobiaosong"] = resolved["songti"] or "SimSun"
    # 若仿宋/楷体/黑体/宋体任一缺失，给最终回退
    resolved["fangsong"] = resolved["fangsong"] or resolved["songti"] or "仿宋"
    resolved["kaiti"]    = resolved["kaiti"]    or resolved["songti"] or "楷体"
    resolved["heiti"]   = resolved["heiti"]   or resolved["songti"] or "黑体"
    resolved["songti"]  = resolved["songti"]  or "宋体"
    resolved["_installed_xbs"] = bool(resolved.get("xiaobiaosong") and
                                       resolved["xiaobiaosong"] != resolved["songti"])
    return resolved


def main():
    fm = resolve_fonts()
    print("=== 公文字体检测结果 ===")
    for k in ("xiaobiaosong", "fangsong", "kaiti", "heiti", "songti"):
        print(f"  {k:14s} -> {fm[k]}")
    print(f"  方正小标宋已安装: {'是' if fm['_installed_xbs'] else '否(回退宋体)'}")
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "font_map.json")
    # 去掉内部字段再写
    pub = {k: v for k, v in fm.items() if not k.startswith("_")}
    pub["xiaobiaosong_installed"] = fm["_installed_xbs"]
    with open(out, "w", encoding="utf-8") as f:
        json.dump(pub, f, ensure_ascii=False, indent=2)
    print(f"  映射已写入: {out}")
    if not fm["_installed_xbs"]:
        print("\n  [提示] 方正小标宋简体未安装。")
        print("  官方免费个人版: https://www.foundertype.com (注册后0元购)")
        print("  下载 FZXBSJW.TTF 后运行: python install_font.py <路径>")
        print("  未装则用宋体回退(排版合规,红头视觉略逊)。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
