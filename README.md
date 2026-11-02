# 整卷二次开发 / Exam Paper Redevelopment

把一套**已经考完、已经讲评过**的高中英语卷，逐篇拆成「可再教、可再练、可再测」的材料：词、块、句、篇四个层级各再发力一遍，一式两版（学生版 / 教师版）。

面向高考英语（新高考 I 卷体例），适配阅读 A–D、七选五、完形填空、语法填空、应用文写作、读后续写六个板块。

## 它解决什么问题

一份卷子讲评完通常就丢掉了。这个 skill 把它拆成四层，由小到大、由形到义：

| 层 | 板块 | 干什么 |
|---|---|---|
| I | 词汇变形与拓展 | 从原文取关键词做词性转换，建立词族意识 |
| II | 短语积累与翻译 | 从原文取完整语块，建立块状语感 |
| III | 句式分析与仿写 | 取原句先析结构，再换场景仿写 |
| IV | 篇级改造 | 按体裁把原篇改成另一种语言任务 |

第 IV 层有五种形态：语法填空改编 / 语篇逻辑排序 / 叙事迁移 / 长难句拆解与翻译 / 微写作迁移。**用哪一种由体裁决定，不由篇位决定**——C/D 篇出现过记叙，A 篇出现过说明，按篇位硬套会选错方向。

## 快速开始

```bash
git clone https://github.com/gaofred/fredgao-exam-vocab-redevelopment.git \
  ~/.workbuddy/skills/fredgao-exam-vocab-redevelopment
```

然后在 agent 会话里说「把这份卷做整卷二次开发」，或直接调用 skill。

手动跑三个脚本：

```bash
# 1. 分段，拿到路由清单
python scripts/exam_sectionizer.py 试卷.docx --output manifest.json

# 2. 人工确认体裁与 IV 形态后，写出内容模型 content.json
#    字段规格见 references/output-contract.md 第 7 节

# 3. 渲染师生两版
python scripts/render_packet.py content.json --outdir out --stem "试卷名_二次开发"

# 4. 校验：板块对齐、题号闭合、答案三件套、答案泄漏、Word 排版禁项
python scripts/validate_packet.py --manifest manifest.json \
  --student "out/试卷名_二次开发（学生版）.docx" \
  --teacher "out/试卷名_二次开发（教师版）.docx" \
  --output QA.json
```

出了 `content.json` 之后，两版必须从它派生。改题改模型再重新渲染，不许直接改 docx——否则两版不同源，下次重生成就丢改动。

## 目录结构

```
SKILL.md                     主流程：六步，含两处人工确认闸
references/
  module-registry.json       板块命名与路由的唯一来源
  genre-routing.md           先判体裁再定 IV 形态
  section-blueprints.md      各板块的做法与数量
  quality-rules.md           内容质量规则与教师版答案三件套
  output-contract.md         交付形态、排版硬规矩、内容模型 schema
  failure-modes.md           九类卡点的兜底表
scripts/
  exam_sectionizer.py        分段 + 路由 + 体裁猜测
  render_packet.py           content.json → 两版 docx
  validate_packet.py         结构、闭合、泄漏、排版四类校验
  check_registry.py          防止板块命名漂移
agents/openai.yaml           显示名与默认提示词
```

## 依赖

- Python 3.9+
- `python-docx`（仅渲染需要）：`pip install python-docx`
- 可选：`pdftotext`（文字版 PDF）、`pdftoppm` + `tesseract`（扫描件 OCR）

分段与校验脚本只用标准库。

## 几个设计上的硬约定

- **单一事实来源**：板块标题只写在 `module-registry.json`，脚本运行时读取，`check_registry.py` 防止有人手写标题导致路由错位。
- **原句必须回流**：第 III 层原句逐字引自本卷并标段落定位，学生拿着原卷能对上位置。
- **改造不是新造**：第 IV 层复用原篇事实链与逻辑顺序，只做考点化手术，不许另写一篇冒充改编。
- **教师版不许只给答案**：封闭题要 `答 / 辨 / 错 / 换`，开放题要 `参 / 忌 / 换`。
- **排版零样式**：不用任何 Word 内置标题样式、大纲级别、目录域、自动编号，层级只靠加粗 + 顶格——学校模板一换也不会崩。

## 适用与不适用

适用：整卷、已考完、要产出师生配对讲义。

不适用：考前命题、单篇词汇讲义、单篇命题质量评析——那些是另外的 skill。

## 作者

Fred 老师 · 公众号：佛瑞德英语教学

## License

双许可：

- `scripts/` 下的代码 —— **MIT**，见 [LICENSE](LICENSE)。可自由使用、修改、再分发。
- 文档与生成的教学讲义 —— **CC BY-NC-SA 4.0**，见 [LICENSE-docs](LICENSE-docs)。可自由用于课堂、校内印刷、教研分享与非营利公众号，需署名并注明是否修改；商业用途请先联系。

课堂里怎么用都行，不用额外问我。要拿去做付费课程或出版物，说一声。
