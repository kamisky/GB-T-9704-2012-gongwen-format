# -*- coding: utf-8 -*-
"""安装方正小标宋简体字体(用户级,无需管理员)。
用法: python install_font.py <FZXBSJW.TTF 路径>
校验文件头后复制到 %LOCALAPPDATA%\\Microsoft\\Windows\\Fonts 并写注册表HKCU,广播字体变更。
"""
import os, sys, shutil, ctypes, winreg

TTF_MAGIC = (b"\x00\x01\x00\x00", b"ttcf", b"true", b"typ1")
OTF_MAGIC = (b"OTTO",)


def validate_font(path):
    with open(path, "rb") as f:
        head = f.read(4)
    return head.startswith(TTF_MAGIC) or head.startswith(OTF_MAGIC)


def font_family_name(path):
    """从TTF/OTF读取中英文族名(简易解析name表)。"""
    try:
        from fontTools.ttLib import TTFont
    except Exception:
        return None, None
    try:
        tt = TTFont(path, fontNumber=0)
        name = tt["name"]
        # nameID 1=family, 4=full name; platform 3 (Windows) langID 0x804(zh-CN) or 0x409(en)
        def get(nid, prefer_lang=(0x804, 0x409, 0)):
            best = None
            for rec in name.names:
                if rec.nameID == nid:
                    try:
                        s = rec.toUnicode()
                    except Exception:
                        continue
                    if rec.langID in prefer_lang:
                        return s
                    best = best or s
            return best
        fam = get(1)
        full = get(4)
        return fam, full
    except Exception as e:
        print(f"[warn] 解析字体名失败: {e}")
        return None, None


def main():
    if len(sys.argv) < 2:
        print("用法: python install_font.py <FZXBSJW.TTF 路径>")
        return 2
    src = os.path.abspath(sys.argv[1])
    if not os.path.isfile(src):
        print(f"[错误] 文件不存在: {src}")
        return 1
    if not validate_font(src):
        print(f"[错误] 文件头非合法 TTF/OTF,拒绝安装(可能不是字体文件): {src}")
        print("       仅接受来自 foundertype.com 官方或可信来源的 .ttf。")
        return 1

    fam, full = font_family_name(src)
    if not fam:
        # 默认按文件名推断方正小标宋
        fam = "方正小标宋简体"
    print(f"字体族名: {fam}" + (f" (全名: {full})" if full else ""))

    local_fonts = os.path.join(os.environ["LOCALAPPDATA"], "Microsoft", "Windows", "Fonts")
    os.makedirs(local_fonts, exist_ok=True)
    dst = os.path.join(local_fonts, os.path.basename(src))
    shutil.copy2(src, dst)
    print(f"已复制: {dst}")

    # 注册到 HKCU
    reg_key = winreg.CreateKeyEx(
        winreg.HKEY_CURRENT_USER,
        r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts", 0, winreg.KEY_WRITE)
    val_name = f"{fam} (TrueType)"
    winreg.SetValueEx(reg_key, val_name, 0, winreg.REG_SZ, dst)
    winreg.CloseKey(reg_key)
    print(f"已注册: HKCU\\...\\Fonts\\{val_name}")

    # 广播字体变更
    HWND_BROADCAST = 0xFFFF
    WM_FONTCHANGE = 0x001D
    ctypes.windll.user32.PostMessageW(HWND_BROADCAST, WM_FONTCHANGE, 0, 0)
    print("已广播字体变更通知。")
    print("\n请在 Word/WPS 中重开,字体栏搜索确认已可用。")
    print("运行 font_check.py 复检。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
