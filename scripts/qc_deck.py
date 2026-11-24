# -*- coding: utf-8 -*-
"""成品自检：对着 8 条硬指标扫一遍，输出「必修 / 建议」两档。

用法：python qc_deck.py deck.pptx [--min-fs 16]
"""
import argparse
import re
from collections import Counter

from pptx import Presentation

CM = 360000.0
SKELETON_FS = {30, 34, 40, 44, 54}       # 章节块 / 标题 / 封面 / 目录 / 结束页，不参与 16pt 下限
PLACEHOLDER = ("待插入", "待补", "占位", "示意图", "TBD", "todo", "XXX")


def cm(v):
    return None if v is None else round(v / CM, 2)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--min-fs", type=float, default=16)
    a = ap.parse_args()
    prs = Presentation(a.pptx)
    W, H = cm(prs.slide_width), cm(prs.slide_height)
    must, soft = [], []
    fs_hist = Counter()

    for i, s in enumerate(prs.slides, 1):
        pageno = []
        for sh in s.shapes:
            name = sh.name
            if sh.left is None:
                continue
            l, t, w, h = cm(sh.left), cm(sh.top), cm(sh.width) or 0, cm(sh.height) or 0
            r, b = l + w, t + h
            if l < -0.05 or t < -0.05 or r > W + 0.05 or b > H + 0.05:
                must.append(f"P{i} 越界：{name} ({l},{t},{w}×{h}) 右下角=({r},{b})")
            if sh.has_table:
                tb = sh.table
                for row in tb.rows:
                    for c in row.cells:
                        for p in c.text_frame.paragraphs:
                            for run in p.runs:
                                if run.font.size:
                                    fs_hist[run.font.size.pt] += len(run.text)
                szs = [run.font.size.pt for row in tb.rows for c in row.cells
                       for p in c.text_frame.paragraphs for run in p.runs if run.font.size]
                hs = [cm(x.height) for x in tb.rows]
                if szs and hs and min(hs) < max(szs) * 0.0353 * 1.25:
                    soft.append(f"P{i} 表格行高 {min(hs)}cm 对 {max(szs)}pt 偏矮，PowerPoint 会自己撑高")
                continue
            if sh.has_text_frame:
                txt = sh.text_frame.text
                low = txt.lower()
                for kw in PLACEHOLDER:
                    if kw.lower() in low:
                        must.append(f"P{i} 残留占位字样「{kw}」：{txt[:20]}")
                if re.fullmatch(r"\d{1,2}", txt.strip()) and l > 28 and t > 16:
                    pageno.append((int(txt.strip()), l, t))
                for p in sh.text_frame.paragraphs:
                    for run in p.runs:
                        if run.font.size:
                            fs_hist[run.font.size.pt] += len(run.text)
                            if (run.font.size.pt < a.min_fs
                                    and run.font.size.pt not in SKELETON_FS
                                    and len(run.text.strip()) > 0):
                                must.append(f"P{i} 字号 {run.font.size.pt}pt < {a.min_fs}pt：{run.text[:18]}")
        if pageno:
            for n, l, t in pageno:
                if n != i:
                    must.append(f"P{i} 页码写成 {n}（应为 {i}）")
                if not (17.35 <= t <= 18.20):
                    soft.append(f"P{i} 页码 T={t}（基准 17.90）")
        else:
            soft.append(f"P{i} 未找到右下角页码")

    print(f"画幅 {W}×{H}，{len(prs.slides)} 页")
    print("字号分布（按字符数）：", dict(sorted(fs_hist.items(), reverse=True)))
    print(f"\n必修 {len(must)} 条")
    for x in must:
        print("  ✗", x)
    print(f"\n建议 {len(soft)} 条")
    for x in soft:
        print("  ·", x)
    if not must:
        print("\n必修项为 0，可交付（仍要逐页渲染看一眼）。")


if __name__ == "__main__":
    main()
