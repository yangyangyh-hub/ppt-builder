# 05 成稿、自检与交付

## 1. 生成

```bash
python scripts/build_deck.py spec.json -o 成品.pptx
```

- 写 `spec.json`（格式见 `06-spec-schema.md`），一页一个对象，`type` 选原型。
- 生成时若有 `⚠ 版式预警`，说明文本估算高度超框：**先改内容或改框高，不要改字号**。
- 版面微调只改 spec 里的坐标/尺寸，不改 `build_deck.py`（除非要动全篇规格）。

## 2. 自检（每次交付前跑）

```bash
python scripts/qc_deck.py 成品.pptx
```

必修项必须清零，重点看：

| 项 | 判据 |
| --- | --- |
| 字号下限 | 无 < 16pt（章节块 30、标题 34、封面 44、目录 54、结束 40 除外） |
| 越界 | 所有形状 0≤l 且 右下角 ≤33.87/19.05 |
| 页码 | 与页序一致；T 在 17.35–18.20 |
| 占位残留 | 无"待插入/待补/示意图/TODO/XX" |
| 表格行高 | ≥ 字号×0.0353×1.30 |
| 图注 | 每条 ≤14 字，无来源括号 |

## 3. 渲染复核（不能省）

```powershell
$app=New-Object -ComObject PowerPoint.Application
$pres=$app.Presentations.Open("成品.pptx",$true,$false,$false)
$i=0
foreach($sl in $pres.Slides){$i++;$sl.Export("png\s{0:d2}.png" -f $i,"PNG",1600,900)}
$pres.Close(); $app.Quit()
```

逐页看四件事：文字有没有压图/压线；图注是否紧贴图底；通栏页的页码是否可见；空白是不是集中在右下角一大块（说明缺内容）。

**像素核验优先于肉眼**：要判断"两个框是否等宽等高"时，用 PIL/numpy 找非白段量像素，`ViewImage` 的缩放会造成错觉。

## 4. "只改格式"类需求（用户高频）

用户常说"内容不要动，只调格式"。这时：

1. 改前备份原件到 `backup/…_格式调整前_时间戳.pptx`；
2. 脚本只改几何/字号/颜色/裁剪，**不碰 run.text**；
3. 收尾做**文本集合差**验证：

```python
a = [(i, t) for i, s in enumerate(prs_a.slides) for t in all_text(s)]
b = [(i, t) for i, s in enumerate(prs_b.slides) for t in all_text(s)]
print(set(a) ^ set(b))     # 只允许页码笔误、前导空格这类可解释差异
```

本次实测：248 个文本块对 248 个，仅 3 处差异（2 处页码笔误 + 1 处前导空格）。凡是超过 3 处，就是脚本误改了内容。

## 5. 交付清单

| 文件 | 说明 |
| --- | --- |
| `xxx.pptx` | 成品 |
| `xxx.pdf` | 同名导出，供打印/报送 |
| `生成说明.md` 或对话内说明 | 页序表、数据口径表、待本人确认项、偏离模板的 3 处及原因 |

不要在交付里夹带：隐藏素材页、占位框、备注里的内部批注。

## 6. 常见崩点速查

| 现象 | 原因 | 处理 |
| --- | --- | --- |
| 中文标点跑到行首 | run 缺 `lang="zh-CN"` | 每个 run 补上 |
| 改了字号后下方元素被压 | `auto_size=SHAPE_TO_FIT_TEXT` 让框自长 | 设 `auto_size=None` 并显式给框高 |
| 数字与中文之间空一大块 | 用了两端对齐 | 改左对齐 |
| 表格盖住下方文字 | 行高不足，PowerPoint 自动撑高 | 行高 ≥ 字号×0.0353×1.30 |
| 图片被拉扁 | 直接设 width/height | 先按目标比例设 `crop_*`，再设尺寸 |
| 通栏页看不到页码 | 红条盖住右下角页码 | 页码放进红条内反白 |
| 补丁不生效 | `str.replace` 没匹配却静默通过 | 每次替换前 `assert a in s` |
| `paragraph_format` 报错 | python-pptx 的 `_Paragraph` 没这个属性 | 用 `pPr.set('marL','0')`、`set('indent','0')` |
| `pic.image.width` 报错 | 属性名不对 | 用 `pic.image.size` |
| `prs.slides[0:3]` 报错 | 切片返回 list | `enumerate(prs.slides)` 后 break |
