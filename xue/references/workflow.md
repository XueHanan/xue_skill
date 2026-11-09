# 文件流程

## 准备环境与范围

脚本仅使用 Python 3.10+ 标准库，可使用当前能运行的 Python，无需安装第三方包。`PYTHON` 表示实际解释器，`SCRIPTS` 表示所调用技能下 scripts 的绝对路径。命令中的 SOURCE、TASK 和 MD 是用户文件、独立工作目录和对照稿路径；路径含空格时加引号。不要把示例段号照抄到用户论文。

先明确用户授权的内容。DOCX 可只读查看顶层正文段号：

```text
PYTHON SCRIPTS/inspect_docx.py SOURCE.docx
```

编号从 0 开始，只包含正文 body 的顶层段落；表格单独列为只读上下文，不占段号。章节号、Word 页码与此段号不是同一概念。核对实际标题及前后文本后确定范围。

```text
PYTHON SCRIPTS/prepare.py SOURCE.docx TASK --range START END --mode theory --term 溺水人员检测
```

DOCX 不提供范围时会拒绝准备；仅在用户授权全文处理时给 `--whole-document`。纯文本文件的整个内容视为所提供范围，可不传 --range。范围包括的标题和冻结段落须在对照稿原样保留；默认全文模式只选择可编辑正文。

```text
PYTHON SCRIPTS/prepare.py SOURCE.txt TASK --mode intro --terms-file TERMS.txt
```

没有文件的小段文字可以直接在聊天内完成，不为跑脚本而创建一整套任务。

## 基线和修改

prepare 生成：

- task.json：选择的段号、模式和受保护术语。范围以此为准。
- paragraphs.json、objects.json：本次提取基线、公式/对象身份和只读表格上下文。
- marked.txt、structure.tsv：查看用的正文和结构。
- comparison.md：覆盖所选段落的对照稿，初始改后与原文相同。

填写 --term 或 TERMS.txt 前，通读当前范围，列出中文任务名、模块名、模型/数据集名和必须原样保留的缩写。已有术语表优先；未声明术语仍须人工复核，自动检查不会理解所有专业名称。

只修改 comparison.md 的改后块和改动说明。原文逐字保留（以提取的文本表示为准，包括 EQ/OBJ 标记和内部空白）。不直接改 paragraphs.json 或 objects.json 来迎合验证。发现提取错误时从原文件在新目录重新准备。

选定范围需要调整时，先按用户意图检查 task.json 的 selected_ids，再同步对照稿；这会使旧语义复核失效。不将范围调整当作可自行扩写论文的授权。

DOCX 的 OMML、可识别的 MathType/Equation OLE 用独立 EQ 标记。普通嵌入对象、Word 域、脚注/尾注引用及未接受的修订等区域保守冻结。表格不改写；文本框等未支持区域不声称覆盖。公式标记只代表保留源对象的映射，不是可直接粘贴进 Word 的真实公式。

## 语义复核与交付

```text
PYTHON SCRIPTS/review_template.py TASK/comparison.md TASK
```

生成 review.json，初始五项检查均为 pending。由执行润色的代理逐段检查并写具体依据：

- claims：命题、结论及因果/比较强度是否一致。
- bindings：指标—数值—单位、对象—属性、论断—引文及公式两侧语义是否对应。
- conditions：否定、限定条件、适用范围与确定性是否一致。
- steps：步骤、依赖关系及先后顺序是否一致；没有步骤也须说明。
- terminology：中文专业术语、缩写、变量和指代对象是否一致。

实际确认保持后，把该项 result 写为 preserved，并写 evidence。不允许用脚本批量把 pending 改为 preserved，或给每段填相同空泛结论。发现改变则先改回正文；无法判断时保留原句或报告待复核。代理自行完成这些检查，不要求用户逐项签字。记录是可追溯的复核说明，不能证明技术结论本身正确。

```text
PYTHON SCRIPTS/validate.py TASK/comparison.md TASK
```

只在统一报告 passed 后，交付当前 comparison.md 和 validation.json 指向的版本化结果。正文再次改变时，生成新的复核文件：

```text
PYTHON SCRIPTS/review_template.py TASK/comparison.md TASK --out TASK/review-next.json
PYTHON SCRIPTS/validate.py TASK/comparison.md TASK --review TASK/review-next.json
```

不覆盖旧复核记录以掩盖版本差异。输出统一使用当前宿主可用的文件链接，无需专用 present 工具。
