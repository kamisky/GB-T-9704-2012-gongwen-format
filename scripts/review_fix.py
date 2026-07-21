# -*- coding: utf-8 -*-
"""自动审查与修复：生成 .docx → 自检 → 分析失败 → 自动修复 → 重新生成，循环至通过。
用法: python review_fix.py <content.json> [输出.docx] [--max-iter 3]
可被 import: from review_fix import review_and_fix
"""
import os, sys, json, re, copy, datetime

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _HERE)

from gongwen_gen import GongwenBuilder
from verify import verify_docx, print_report

KI = os.path.join(os.path.dirname(_HERE), "references", "known_issues.md")


# ============================================================
# 修复策略
# ============================================================

def _normalize_date(date_str):
    """尝试将各种日期格式规范化为 'YYYY年M月D日'。"""
    if not date_str:
        return None
    # 已经是标准格式
    if re.match(r'^\d{4}年\d{1,2}月\d{1,2}日$', date_str):
        return None  # 无需修改
    # 2026-07-16 / 2026/07/16 / 2026.07.16
    m = re.match(r'(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})', date_str)
    if m:
        y, mo, d = m.groups()
        return f"{y}年{int(mo)}月{int(d)}日"
    # 20260716
    m = re.match(r'(\d{4})(\d{2})(\d{2})', date_str)
    if m:
        y, mo, d = m.groups()
        return f"{y}年{int(mo)}月{int(d)}日"
    # 2026年7月16日 (already correct but with possible leading zeros)
    m = re.match(r'(\d{4})年(\d{1,2})月(\d{1,2})日', date_str)
    if m:
        y, mo, d = m.groups()
        return f"{y}年{int(mo)}月{int(d)}日"
    return None


def _fix_content(content, failures):
    """根据失败项自动修正 content dict。
    
    Returns:
        (fixed_content, fix_log): 修正后的 content 和修复日志列表。
        如果无需修改，返回 (None, [])。
    """
    fixed = json.loads(json.dumps(content))  # deep copy
    log = []

    for f in failures:
        name = f["name"]

        # ---- 发文字号：方括号 → 六角括号 ----
        if "发文字号" in name:
            dn = fixed.get("doc_number", "")
            if dn:
                new_dn = dn.replace("[", u"〔").replace("]", u"〕")
                if new_dn != dn:
                    fixed["doc_number"] = new_dn
                    log.append(f"发文字号: '{dn}' → '{new_dn}'")

        # ---- 成文日期：规范化格式 ----
        if "成文日期" in name:
            for key in ("date", "print_date"):
                date_str = fixed.get(key, "")
                new_date = _normalize_date(date_str)
                if new_date:
                    fixed[key] = new_date
                    log.append(f"{key}: '{date_str}' → '{new_date}'")

    if log:
        return fixed, log
    return None, []


def _is_fixable(failure):
    """判断失败项是否可自动修复。"""
    name = failure["name"]
    # 可自动修复的项
    fixable = [
        "发文字号",           # 括号替换
        "成文日期",           # 日期格式规范化
        "A4幅面",             # 重新生成即可
        "天头", "下边距", "订口", "切口",  # 重新生成
        "版心",               # 重新生成
        "默认3号",            # 重新生成
        "页码含",             # 重新生成
        "开启奇偶页",         # 重新生成
        "版头红色分隔线",     # 重新生成
        "版记分隔线",         # 重新生成
        "抄送格式",           # 重新生成（生成器已正确处理）
    ]
    for keyword in fixable:
        if keyword in name:
            return True
    # 不可自动修复：需要用户补充内容
    return False


def _needs_regeneration(failure):
    """判断是否需要重新生成（而非修改 content）。"""
    name = failure["name"]
    content_fixable = ["发文字号", "成文日期"]
    return not any(kw in name for kw in content_fixable)


# ============================================================
# 主循环
# ============================================================

def review_and_fix(content_path, output_path, max_iter=3):
    """主循环：生成 → 自检 → 修复 → 重新生成。
    
    Args:
        content_path: content.json 路径
        output_path: 输出 .docx 路径
        max_iter: 最大迭代次数
    
    Returns:
        dict: {
            "success": bool,
            "iterations": int,
            "final_result": verify result dict,
            "fix_log": [str, ...],
            "unfixable": [{"name": str, "detail": str}, ...],
        }
    """
    # 加载原始 content
    with open(content_path, "r", encoding="utf-8") as f:
        content = json.load(f)

    # 保存原始 content 以便回写
    original_content = json.loads(json.dumps(content))
    all_fix_log = []
    final_result = None

    for i in range(max_iter):
        iteration = i + 1
        print(f"\n{'='*50}")
        print(f"🔄 第 {iteration}/{max_iter} 轮：生成公文...")
        print(f"{'='*50}")

        # 1. 生成
        doc = GongwenBuilder(content).build()
        doc.save(output_path)
        print(f"   📄 已生成: {output_path}")

        # 2. 自检
        result = verify_docx(output_path)
        final_result = result

        # 3. 全部通过
        if result["all_pass"]:
            print(f"\n✅ 第 {iteration} 轮: 全部通过 ({result['passed']}/{result['total']})")
            # 回写修复后的 content.json（如果有修改）
            if content != original_content:
                with open(content_path, "w", encoding="utf-8") as f:
                    json.dump(content, f, ensure_ascii=False, indent=2)
                print(f"   📝 content.json 已更新（修复了 {len(all_fix_log)} 项）")
            return {
                "success": True,
                "iterations": iteration,
                "final_result": result,
                "fix_log": all_fix_log,
                "unfixable": [],
            }

        # 4. 分析失败项
        fixable = [f for f in result["failures"] if _is_fixable(f)]
        unfixable = [f for f in result["failures"] if not _is_fixable(f)]

        print(f"\n   ⚠️  {len(result['failures'])} 项失败:")
        for f in result["failures"]:
            tag = "🔧" if _is_fixable(f) else "❌"
            detail = f" ({f['detail']})" if f["detail"] else ""
            print(f"      {tag} {f['name']}{detail}")

        if unfixable:
            print(f"\n   ❌ {len(unfixable)} 项无法自动修复，需用户介入:")
            for f in unfixable:
                detail = f" ({f['detail']})" if f["detail"] else ""
                print(f"      • {f['name']}{detail}")
            # 即使有无法修复的项，也尝试修复可修复的
            if not fixable:
                break

        if not fixable:
            print(f"\n   ❌ 无项目可自动修复。")
            break

        # 5. 尝试修复
        new_content, fix_log = _fix_content(content, fixable)
        all_fix_log.extend(fix_log)

        if fix_log:
            print(f"\n   🔧 自动修复了 {len(fix_log)} 项:")
            for entry in fix_log:
                print(f"      • {entry}")
            content = new_content
        else:
            # 可修复但 _fix_content 返回了 None（不需要修改 content），
            # 说明需要重新生成
            print(f"\n   🔄 重新生成以修复结构性偏差...")

    # 达到最大迭代次数
    print(f"\n⚠️  已达最大迭代次数 ({max_iter})，仍有 {len(final_result['failures'])} 项未通过。")

    # 回写修复后的 content.json
    if content != original_content:
        with open(content_path, "w", encoding="utf-8") as f:
            json.dump(content, f, ensure_ascii=False, indent=2)
        print(f"   📝 content.json 已更新（修复了 {len(all_fix_log)} 项）")

    # 写入 known_issues
    if final_result["failures"]:
        with open(KI, "a", encoding="utf-8") as f:
            f.write(f"\n## [{datetime.date.today()}] review_fix 残留偏差 ({os.path.basename(output_path)})\n")
            for c in final_result["failures"]:
                f.write(f"- 现象: {c['name']} 不合规" +
                        (f" / {c['detail']}" if c["detail"] else "") + "\n")
                f.write(f"- 国标依据: 见 references/gbt9704-2012-spec.md\n")
                f.write(f"- 修复: 经 {max_iter} 轮自动修复仍未解决\n")
        print(f"   偏差已追加至: {KI}")

    return {
        "success": False,
        "iterations": max_iter,
        "final_result": final_result,
        "fix_log": all_fix_log,
        "unfixable": unfixable if final_result else [],
    }


# ============================================================
# CLI
# ============================================================

def main():
    import argparse
    ap = argparse.ArgumentParser(
        description="GB/T 9704—2012 公文自动生成 + 审查修复")
    ap.add_argument("content", help="content.json 路径")
    ap.add_argument("output", nargs="?", default=None,
                    help="输出 .docx 路径（默认与 content 同目录同名）")
    ap.add_argument("--max-iter", type=int, default=3,
                    help="最大修复迭代次数（默认 3）")
    ap.add_argument("--report", action="store_true",
                    help="仅运行自检，不生成不修复")
    args = ap.parse_args()

    content_path = os.path.abspath(args.content)
    if args.output:
        output_path = os.path.abspath(args.output)
    else:
        output_path = os.path.splitext(content_path)[0] + ".docx"

    if args.report:
        # 仅自检模式
        result = verify_docx(output_path)
        print_report(result, output_path)
        return 1 if result["failures"] else 0

    # 完整流程
    outcome = review_and_fix(content_path, output_path, args.max_iter)

    # 打印最终报告
    print(f"\n{'='*50}")
    print("📋 最终自检报告")
    print(f"{'='*50}")
    if outcome["final_result"]:
        print_report(outcome["final_result"], output_path)

    if outcome["fix_log"]:
        print(f"\n🔧 本轮共自动修复 {len(outcome['fix_log'])} 项:")
        for entry in outcome["fix_log"]:
            print(f"   • {entry}")

    if outcome["success"]:
        print(f"\n✅ 公文生成成功！共 {outcome['iterations']} 轮，全部通过国标自检。")
        return 0
    else:
        print(f"\n⚠️  公文已生成，但仍有 {len(outcome.get('unfixable', []))} 项无法自动修复。")
        print("   请检查 content.json 并手动修正后重新运行。")
        return 1


if __name__ == "__main__":
    sys.exit(main())