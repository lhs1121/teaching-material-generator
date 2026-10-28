---
name: teaching-material-generator
display_name: 课堂教学材料生成器
display_name_en: Teaching Material Generator
description: 面向职业教育（高职/本科）的备课技能，从岗位典型工作任务或教材章节出发，一次性生成「教案全文 + 配套课件大纲 + 说课稿」三件套。当用户提到生成教案、做课件、写说课稿、备课、教学设计时使用本技能；非职业教育的 K12 课件、实训任务开发、能力图谱绘制不适用本技能。
description_zh: 面向职业教育（高职/本科）的备课技能，从岗位典型工作任务或教材章节出发，一次生成「教案全文 + 课件大纲 + 说课稿」三件套，理实一体化、统一职教编写铁律。
description_en: A teaching-prep skill for vocational education (higher vocational / applied undergraduate). From a job's typical work tasks or a textbook chapter, it generates a complete lesson plan, courseware outline, and teaching speech in one pass, with theory-practice integration and unified vocational writing rules.
examples_zh:
  - 帮我校《商务数据分析》数据清洗这一章做教案、课件和说课稿，高职大二，4 学时
  - 按岗位任务出一份理实一体教案，岗位是跨境电商运营，任务是店铺流量转化诊断
  - 给我出一份供求均衡理论的教案，高职大一，2 学时，经济学基础课
examples_en:
  - Generate a lesson plan, courseware outline and teaching speech for the Data Cleaning chapter of a Business Data Analysis course
  - Create a theory-practice integrated lesson plan for the job task store traffic conversion diagnosis in cross-border e-commerce
version: 0.4.1
author: 意瞬
---

# 课堂教学材料生成器（职业教育）

当用户需要生成职业教育课程的教案、课件大纲、说课稿，或进行备课时，按以下步骤执行。

## 执行步骤

### 第 1 步：判定生成路径

阅读 `@references/input-spec.md` 的「步骤 0 路由判断」，判定走哪条路径，中途不切换：

- **路径 A 岗位驱动**：专业核心课 / 实训课 / 理实一体课，或用户已给出岗位与典型工作任务 → 教案十章。
- **路径 B 章节驱动**：公共基础课 / 专业基础理论课 → 教案九章。
- 判为 A 但用户只给了章节名：按 `@references/prompts.md` 的 `PROMPT_TASK_INFER` 反推典型工作任务，向用户确认一句后走 A，**不得直接退回 B**。

### 第 2 步：收集课程基本信息

按 `@references/input-spec.md` 的字段定义补齐必填项（专业大类、课程名称、授课章节/主题、授课对象）。缺失时用一次简洁提问补齐，不逐条追问；课时默认 2，地点/教师给默认，参考稿可选。路径 A 须额外收集课程性质、对应岗位群、典型工作任务。

### 第 3 步：确定生成范围

默认三件套一次出（教案 + 课件大纲 + 说课稿）；用户明确只要其一时，只生成对应产物。

### 第 4 步：调用 Prompt 模板生成

打开 `@references/prompts.md`，按「使用约定」选择模板，把课程基本信息填入占位符，连同「通用前置约束」一并注入：

| 路径 | 教案 | 课件大纲 | 说课稿 |
|------|------|---------|--------|
| A 岗位驱动 | `PROMPT_LESSON_PLAN_A`（十章） | `PROMPT_COURSEWARE` + 路径 A 附加要求 | `PROMPT_SPEECH` + 路径 A 附加要求 |
| B 章节驱动 | `PROMPT_LESSON_PLAN`（九章） | `PROMPT_COURSEWARE` | `PROMPT_SPEECH` |

若单次输出过长，按「教案 → 课件 → 说课」分三次调用，避免截断。

### 第 5 步：落盘文件

按 `@references/writing-rules.md` 第 1 条命名，三份产物分别写为独立 `.md` 文件。已存在同名 V1 时升为 V2 并保留旧版，不覆盖。

### 第 6 步：逐条自检

对照 `@references/writing-rules.md` 末尾「自检速查」逐项核对（免责提示、缩写标注、禁用词、参训者称呼、A/B/C 代称、无编号引用、真实数据、结构完整、学情分析章、课程思政融入点表）。走路径 A 还须通过「路径 A 专项检查」六项。不合格就地修改。

## 转 Word（可选）

产物为 Markdown。需要 `.docx` 时运行 `scripts/md2docx.py` 批量转换（自动统一宋体、补表格边框、设重复表头），再用 `scripts/verify_docx.py` 校验。依赖 `pandoc` 与 `python-docx`：

```bash
python scripts/md2docx.py --out dist "主题 教案 V1.md" "主题 课件大纲 V1.md" "主题 说课稿 V1.md"
python scripts/verify_docx.py dist/*.docx
```

教案「板书设计」章为 ASCII 树形图，若转 Word 后对齐错乱，给 `md2docx.py` 加 `--code-font Consolas`。

## 边界

- 生成内容为 AI 辅助产出，须结合本校人才培养方案与真实学情二次修订后使用；每份产物标题下方已内嵌免责提示。
- 案例人物一律用「团队成员A/B/C」代称；授课对象称「参训者」，不写「学生」。
- 真实数据须标注来源或注明（示例数据），不编造虚假行业数据。
