# -*- coding: utf-8 -*-
"""模板解析：把一个 pptx 拆成「可继承的版式参数」。

用法：
    python analyze_pptx.py deck.pptx                 # 人类可读摘要
    python analyze_pptx.py deck.pptx -o out.json     # 另存结构化 JSON
    python analyze_pptx.py deck.pptx --page 3        # 只看某页

输出四块，正好对应 references/04-template-scan.md 的 8 类构件：
    1) canvas      画幅
    2) skeleton    每页重复出现的形状（章节块 / 标题 / 副标 / logo / 页脚线 / 页码）
    3) type_scale  字号四档（按字符数加权的分布）
    4) palette     实际用到的颜色 + 用法
"""
import argparse
import json
import re
from collections import Counter, defaultdict

from pptx import Presentation

EMU_CM = 360000.0


def cm(v):
    return None if v is None else round(v / EMU_CM, 2)


def rgb_of(font=None, shape=None, what="font"):
    try:
        if what == "font":
            return str(font.color.rgb)
        if what == "fill" and shape.fill.type == 1:
            return str(shape.fill.fore_color.rgb)
        if what == "line" and shape.line.color.type is not None:
            return str(shape.line.color.rgb)
    except Exception:
        return None
    return None


def snap(prs):
    pages = []
    for i, s in enumerate(prs.slides, 1):
        items = []
        for sh in s.shapes:
            it = {
                "name": sh.name,
                "kind": "shape",
                "l": cm(sh.left), "t": cm(sh.top),
                "w": cm(sh.width), "h": cm(sh.height),
            }
            if sh.shape_type == 13 or sh.__class__.__name__ == "Picture":
                it["kind"] = "pic"
                try:
                    it["px"] = list(sh.image.size)
                    it["ratio"] = round(sh.image.width / sh.image.height, 3)
                    it["box_ratio"] = round(sh.width / sh.height, 3) if sh.height else None
                    it["crop"] = [round(sh.crop_left, 3), round(sh.crop_right, 3),
                                  round(sh.crop_top, 3), round(sh.crop_bottom, 3)]
                except Exception:
                    pass
            elif sh.has_table:
                tb = sh.table
                it["kind"] = "table"
                it["rows"], it["cols"] = len(tb.rows), len(tb.columns)
                it["colw"] = [cm(c.width) for c in tb.columns]
                it["rowh"] = [cm(r.height) for r in tb.rows]
                it["header_fill"] = rgb_of(shape=sh, what="fill")
                it["text"] = [[c.text for c in r.cells] for r in tb.rows]
                szs = [r.font.size.pt for row in tb.rows for c in row.cells
                       for p in c.text_frame.paragraphs for r in p.runs if r.font.size]
                it["sizes"] = sorted(set(szs))
            elif sh.has_text_frame and sh.text_frame.text.strip():
                it["text"] = sh.text_frame.text
                it["fill"] = rgb_of(shape=sh, what="fill")
                runs = []
                for p in sh.text_frame.paragraphs:
                    for r in p.runs:
                        runs.append({
                            "sz": r.font.size.pt if r.font.size else None,
                            "b": r.font.bold,
                            "c": rgb_of(font=r.font),
                            "fn": r.font.name,
                            "n": len(r.text),
                        })
                it["runs"] = runs
                it["align"] = str(sh.text_frame.paragraphs[0].alignment)
                it["line"] = str(sh.text_frame.paragraphs[0].line_spacing)
            items.append(it)
        pages.append({"page": i, "shapes": items})
    return pages


def find_skeleton(pages):
    """跨页重复出现的「版位」→ 骨架构件。

    两道互补判据：
      A 同位置同尺寸（不管文字是什么）→ 标题框、副标框、页码框、校徽
      B 同尺寸同文本模式（数字归一化）→ 章节块、目录编号
    """
    by_slot = defaultdict(list)   # (w,h,t,l) 容差 0.3cm
    by_text = defaultdict(list)
    pic_slot = defaultdict(list)
    for pg in pages:
        for it in pg["shapes"]:
            if it["kind"] == "pic":
                if it["w"] and it["h"]:
                    pic_slot[(round(it["w"], 1), round(it["h"], 1))].append(pg["page"])
                continue
            if it["l"] is None:
                continue
            slot = (round(it["w"], 1), round(it["h"], 1),
                    round(it["t"] * 2) / 2, round(it["l"] * 2) / 2)
            by_slot[slot].append((pg["page"], it))
            if it["kind"] == "shape" and it.get("text"):
                k = (round(it["w"], 1), round(it["h"], 1),
                     re.sub(r"[0-9]", "#", it["text"].strip())[:10])
                by_text[k].append((pg["page"], it))

    thr = max(3, len(pages) // 4)
    out = []
    for k, hits in by_slot.items():
        if len(hits) < thr:
            continue
        it = hits[0][1]
        sz = it.get("runs", [{}])[0].get("sz") if it.get("runs") else None
        if it["kind"] == "table":
            sz = it.get("sizes")
        out.append({
            "判据": "版位", "w": k[0], "h": k[1], "t": k[2], "l": k[3],
            "出现在": f"{len(hits)}/{len(pages)} 页",
            "字号": sz, "填充": it.get("fill"),
            "样例": (it.get("text") or "").replace("\n", " ")[:18],
        })
    for k, hits in by_text.items():
        if len(hits) < thr or len(hits) >= len(pages) * 0.9:
            continue
        it = hits[0][1]
        out.append({
            "判据": "文本", "w": k[0], "h": k[1], "t": it["t"], "l": it["l"],
            "出现在": f"{len(hits)}/{len(pages)} 页",
            "字号": it["runs"][0]["sz"] if it.get("runs") else None,
            "填充": it.get("fill"),
            "样例": (it.get("text") or "").replace("\n", " ")[:18],
        })
    for k, hits in pic_slot.items():
        if len(hits) >= thr:
            out.append({"判据": "图片", "w": k[0], "h": k[1], "t": "-", "l": "-",
                        "出现在": f"{len(hits)}/{len(pages)} 页", "字号": "-",
                        "填充": "-", "样例": "重复图片（logo/水印）"})
    return out


def type_scale(pages):
    c = Counter()
    for pg in pages:
        for it in pg["shapes"]:
            if it["kind"] == "table":
                for sz in it.get("sizes", []):
                    c[sz] += 10 * it["rows"]
            for r in it.get("runs", []):
                if r["sz"]:
                    c[r["sz"]] += r["n"]
    tot = sum(c.values()) or 1
    return [{"pt": k, "字符数": v, "占比": f"{100 * v / tot:.1f}%"}
            for k, v in sorted(c.items(), reverse=True)]


def palette(pages):
    c = Counter()
    for pg in pages:
        for it in pg["shapes"]:
            for r in it.get("runs", []):
                if r["c"]:
                    c[f"文字 #{r['c']}"] += r["n"]
            if it.get("fill"):
                c[f"填充 #{it['fill']}"] += 20
    return [{"项": k, "量": v} for k, v in c.most_common(12)]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("-o", "--out")
    ap.add_argument("--page", type=int)
    a = ap.parse_args()

    prs = Presentation(a.pptx)
    pages = snap(prs)
    data = {
        "file": a.pptx,
        "canvas": [cm(prs.slide_width), cm(prs.slide_height)],
        "pages": len(pages),
        "skeleton": find_skeleton(pages),
        "type_scale": type_scale(pages),
        "palette": palette(pages),
        "detail": pages,
    }
    if a.out:
        with open(a.out, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=1)
        print("json ->", a.out)

    show = [p for p in pages if a.page in (None, p["page"])]
    print(f"\n画幅 {data['canvas'][0]} × {data['canvas'][1]} cm，共 {data['pages']} 页\n")
    print("=== 骨架构件（跨页重复的版位） ===")
    for k in data["skeleton"]:
        print(f"  {k['判据']} {k['出现在']:>8} | {str(k['w']):>5}×{str(k['h']):<5} "
              f"T{k['t']} L{k['l']} 字号{k['字号']} 填充{k['填充']} | {k['样例']}")
    print("\n=== 字号分布（按字符数加权） ===")
    for r in data["type_scale"]:
        print(f"  {r['pt']:>5}pt  {r['字符数']:>5} 字  {r['占比']:>6}")
    print("\n=== 配色 ===")
    for r in data["palette"]:
        print(f"  {r['项']:<14} {r['量']}")
    print("\n=== 逐页 ===")
    for pg in show:
        ts = [it.get("text", "").replace("\n", " ")[:18] for it in pg["shapes"]
              if it["kind"] != "pic" and it.get("text")]
        pics = sum(1 for it in pg["shapes"] if it["kind"] == "pic")
        tabs = sum(1 for it in pg["shapes"] if it["kind"] == "table")
        print(f"  P{pg['page']:>2} 形状{len(pg['shapes']):>3} 图{pics} 表{tabs} | {' / '.join(ts[:4])}")


if __name__ == "__main__":
    main()
