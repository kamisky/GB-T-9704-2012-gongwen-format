# -*- coding: utf-8 -*-
"""生成后自检: 对照 GB/T 9704—2012 检查 .docx, 打印 PASS/FAIL, 偏差追加到 known_issues.md。
用法: python verify.py <公文.docx>
"""
import os, sys, datetime
from docx import Document
from docx.shared import Mm, Pt
from docx.oxml.ns import qn

KI = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                  "references", "known_issues.md")

results = []
def check(name, ok, detail=""):
    results.append((name, bool(ok), detail))

def mm(v):
    # python-docx Length 是 EMU 的子类; 1mm=36000EMU。直接用 .mm 属性更稳。
    if v is None:
        return 0
    if hasattr(v, "mm"):
        return round(v.mm, 1)
    return round(v / 36000.0, 1)

def main():
    path = os.path.abspath(sys.argv[1])
    doc = Document(path)
    s = doc.sections[0]

    # ---- 页面 ----
    check("A4幅面 210×297", abs(mm(s.page_width)-210)<1 and abs(mm(s.page_height)-297)<1,
          f"{mm(s.page_width)}×{mm(s.page_height)}")
    check("天头37mm", abs(mm(s.top_margin)-37)<=1, f"{mm(s.top_margin)}")
    check("下边距35mm", abs(mm(s.bottom_margin)-35)<=1, f"{mm(s.bottom_margin)}")
    check("订口28mm", abs(mm(s.left_margin)-28)<=1, f"{mm(s.left_margin)}")
    check("切口26mm", abs(mm(s.right_margin)-26)<=1, f"{mm(s.right_margin)}")
    check("版心156×225", abs(mm(s.page_width-s.left_margin-s.right_margin)-156)<=2 and
          abs(mm(s.page_height-s.top_margin-s.bottom_margin)-225)<=2,
          f"{mm(s.page_width-s.left_margin-s.right_margin)}×{mm(s.page_height-s.top_margin-s.bottom_margin)}")

    # ---- Normal 样式 ----
    st = doc.styles["Normal"]
    sz = st.font.size.pt if st.font.size else 0
    check("默认3号(16pt)仿宋", abs(sz-16)<0.5, f"{sz}pt")

    # ---- 段落要素扫描 ----
    texts = [(p.text, p) for p in doc.paragraphs]
    joined = "\n".join(t for t,_ in texts)
    # 发文字号: 普通格式用六角括号〔〕; 命令(令)格式令号为"第X号"(无括号)
    dn = next((t for t, _ in texts
               if ("〔" in t and "〕" in t and "印章" not in t and "加盖" not in t)
               or (t.strip().startswith("第") and "号" in t)), "")
    ok_dn = ("〔" in dn and "〕" in dn) or (dn.strip().startswith("第") and "号" in dn)
    check("发文字号合规(〔〕或第X号)", ok_dn, dn or "(未找到发文字号)")
    # 标题
    title_found = any(t for t,_ in texts if t.strip())
    check("存在标题段落", title_found, "")
    # 成文日期阿拉伯数字
    date_ok = any(("年" in t and "月" in t and "日" in t) for t,_ in texts)
    check("成文日期含年月日", date_ok, "")
    # 附件说明(若有)
    # 抄送/印发 4号(仅信息性, 不强制)
    copy_found = any("抄送：" in t for t,_ in texts)
    print_found = any("印发" in t for t,_ in texts)
    check("抄送格式(抄送：...。)", (not copy_found) or any(t.startswith("抄送：") and t.rstrip().endswith("。") for t,_ in texts),
          "" if not copy_found else "(检查末尾句号)")

    # ---- 页码(奇偶/宋体4号/一字线) ----
    odd_f = s.footer.paragraphs[0].text if s.footer.paragraphs else ""
    # PAGE field 文本不可见, 检查 footer 是否有 — 与 字段
    odd_xml = s.footer._element.xml
    even_xml = s.even_page_footer._element.xml
    check("页码含一字线(—)", "—" in odd_xml and "—" in even_xml, "")
    check("页码含PAGE域", "PAGE" in odd_xml.upper() and "PAGE" in even_xml.upper(), "")
    check("开启奇偶页不同", s.footer.is_linked_to_previous is False and s.even_page_footer.is_linked_to_previous is False, "")

    # ---- 红色分隔线 ----
    has_red_border = any("FF0000" in p._element.xml.upper() for p in doc.paragraphs)
    check("版头红色分隔线", has_red_border, "")

    # ---- 版记分隔线(仅当存在版记要素时才校验) ----
    has_banji_text = ("抄送" in joined) or ("印发" in joined)
    has_banji_border = any('w:color="000000"' in p._element.xml for p in doc.paragraphs)
    if has_banji_text:
        check("版记分隔线(黑色)", has_banji_border, "")
    else:
        print("[N/A] 版记分隔线(黑色)  (本文无版记要素, 跳过)")

    # ---- 汇总 ----
    fails = [(n, d) for n, ok, d in results if not ok]
    print("=" * 50)
    print("GB/T 9704—2012 自检报告")
    print("=" * 50)
    for n, ok, d in results:
        print(f"[{'PASS' if ok else 'FAIL'}] {n}" + (f"  ({d})" if d else ""))
    print("-" * 50)
    print(f"通过 {len(results)-len(fails)}/{len(results)}")
    if fails:
        with open(KI, "a", encoding="utf-8") as f:
            f.write(f"\n## [{datetime.date.today()}] 自检偏差 ({os.path.basename(path)})\n")
            for n, d in fails:
                f.write(f"- 现象: {n} 不合规" + (f" / {d}" if d else "") + "\n")
                f.write(f"- 国标依据: 见 references/gbt9704-2012-spec.md\n")
                f.write(f"- 修复: 待核实(可能为渲染/近似或脚本逻辑)\n")
        print(f"偏差已追加至: {KI}")
    return 1 if fails else 0

if __name__ == "__main__":
    sys.exit(main())
