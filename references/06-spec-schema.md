# 06 spec.json 格式（build_deck.py 的输入）

顶层：

```json
{
  "out": "成品.pptx",
  "template": "可选，缺省用 assets/template.pptx",
  "start_page": 1,
  "pages": [ ... ]
}
```

每个页面对象必须有 `type`；除 `cover / toc / end` 外，都用 `ch`（章节号）、`title`（34pt）、`sub`（24pt）画骨架。
`notes` 是**演讲备注**（str），`footnotes` 是**表下注**（list），两者不要混写。

## 通用字段

| 字段 | 含义 |
| --- | --- |
| `type` | `cover` `toc` `tri` `flow` `blocks` `table` `image` `custom` `end` |
| `ch` | 五边形数字 1–6 |
| `title` / `sub` | 页标题 / 副标 |
| `banner` | 底部通栏文字（tri / flow 支持） |
| `notes` | 演讲备注 80–110 字 |

## 原型

```json
{"type":"cover","title":"主标题单行","kicker":"副题","byline":["汇报人：  姓名　职称","单位","2026 年 10 月"]}

{"type":"toc","title":"目录","items":["一　预期目标","二　攻关思路","三　团队建设情况","四　联合团队及承担任务","五　科研条件能力","六　经费概算"]}

{"type":"tri","ch":1,"title":"一 预期目标","sub":"1、拟解决的关键问题",
 "cols":[{"lead":"01　基础科学问题","body":"……"},{"lead":"02　…","body":"…"},{"lead":"03　…","body":"…"}],
 "figs":[{"path":"图1.png","caption":"图1　生猪养殖成本核算"}],
 "colh":5.55,"fig_top":9.30,"fig_bottom":16.45,
 "banner":"节粮降本　·　健康防控　·　提质增效"}

{"type":"flow","ch":1,"title":"一 预期目标","sub":"2、预期成效",
 "claim":"构建……技术体系",
 "chips":["饲料原料","饲料配方","减抗替抗","智能饲喂","猪群健康","提升品质"],
 "blocks":[{"lead":"技术规程","body":"……"},{"lead":"饲料产品","body":"……"}],
 "banner":"…"}

{"type":"blocks","ch":5,"title":"五 科研条件能力",
 "cards":[{"lead":"平台","body":"……"},…4 个],
 "pics":[{"path":"a.png"},{"path":"b.png"}],
 "card_h":3.70,"pic_top":12.10,"pic_h":5.60}

{"type":"table","ch":1,"title":"一 预期目标","sub":"3、指标设置","lead":"可选导语",
 "table":{"l":3.40,"t":3.14,"w":27.10,"colw":[7.66,16.73],"rowh":0.92,"size":20,
          "cols":["指标","目标值"],"rows":[["饲料粗蛋白水平","降低 2%"]],"total_row":false},
 "side":{"l":1.70,"t":10.60,"w":14.30,"h":7.20,"lead":"经济效益","body":["…","…"]},
 "pics":[{"path":"…","l":20.30,"t":4.90,"w":11.40,"h":7.00,"caption":"图5　集成验证现场"}],
 "footnotes":["数据来源：…"]}

{"type":"image","ch":2,"title":"二 攻关思路","sub":"2、技术路线",
 "path":"路线图.png","w":21.98,"t":3.30,"h":15.10,"caption":"图4　技术路线图","center_cap":false}

{"type":"end","title":"敬请各位专家指导！","lines":["姓名　手机号","单位"]}
```

## 排布提醒

- `tri`：三列宽 9.55 固定；`fig_top` 默认 = 列底 + 0.20；`fig_bottom` 默认 16.45（图注在 16.50，通栏 17.45）。
- `flow`：`claim` 在 3.60，链条在 5.00，四块在 6.95 起（行距 4.30 + 1.05）。
- `blocks`：卡片 14.00 宽，起点 2.05 / 17.25；图片居中，间距 1.20。
- `table`：`l` 可用 1.70（满宽）或 3.40（居中偏窄）；`rowh` 会被自动抬到 字号×0.0353×1.30 的下限。
- `custom`：兜底原型，`items` 里给 `{"kind":"text|rect|pic", ...}` 的绝对坐标，用于一次性特殊版式。

## 校验

生成后立刻跑 `qc_deck.py`；`⚠ 版式预警`（文本估算超框）必须处理完再看渲染图。
