# xue_skill

用于中文学位论文润色的两个 AI Agent 技能：**xue** 和 **xue-plus**。支持小段文字直接润色，以及 DOCX 章节的逐段对照、公式标记和内容核对。默认输出对照稿，不写回原 Word。

当前版本：**v2.1（2026-11-09）**。

**[下载 v2.1 技能包 ZIP](https://github.com/XueHanan/xue_skill/raw/refs/heads/main/dist/xue-skills%282.1%29.zip)** · [历史 v2.0 技能包](https://github.com/XueHanan/xue_skill/raw/refs/heads/main/dist/xue-skills%282.0%29.zip) · [历史 v1.0 技能包](https://github.com/XueHanan/xue_skill/raw/refs/heads/main/dist/xue-skills%281.0%29.zip) · [xue 技能说明](xue/SKILL.md) · [xue-plus 技能说明](xue-plus/SKILL.md)

## v2.1 更新

- 论文正文优先从定义、机制、关系及适用条件展开抽象论述，避免例如等引例表达、教学式具体演示和仅删除引例词的表面改写。
- 区分说明性示例与真实实验结果、实际模型配置、对比数据和推导所需条件，保留必要的技术证据；说明性数值删减需人工复核并如实报告自动校验结果。
- 补充数学表达与指代规范，明确张量向量化、雅可比矩阵和恒等连接的关系，区分数学表示与实际网络操作，核对公式占位符。
- 避免仍可按……理解等面向读者的讲解式表达，直接陈述结构调整、下采样倍数和空间映射关系，保留事实、适用条件及必要限定。
- 两个技能入口与风格参考同步更新；处理脚本、任务结构及自动校验口径与 v2.0 相同。

v1.0、v2.0 保留在 `dist/` 作为历史下载版本；仓库当前两个技能目录对应 v2.1。

## v2.0 更新

- 根据作者已确认的改写例，补充段首判断句、分号对举、场景要求作主语等表达偏好，减少叙述句中提示加解释式的冒号。
- 明确保留合理的把字句，按句法位置判断如何的使用，调整关注如何、对……而言、一并／共同考虑／统筹等在指定语境下的表达。
- 两个入口均指向 `references/styles.md` 的用户偏好部分，两份风格参考保持一致。这些偏好只用于当前授权范围内的同类叙述，不扩大为全篇禁令。
- 本次是风格规则迭代，处理脚本、任务结构与校验口径均与 v1.0 相同，不涉及脚本不兼容或任务格式迁移。安装新版仍建议整体替换技能目录。

## 选择技能

| 技能 | 默认适用内容 |
| --- | --- |
| `xue` | 绪论、研究现状、理论基础与相关技术 |
| `xue-plus` | 摘要、方法、算法、系统设计、实验、结论与展望 |

根据实际内容选择，不强制按章节号划分。同一任务中包含多种章节时，代理可按内容使用对应规则。两个技能均保留显式点名调用的约定。

## 安装

### 下载 ZIP

1. 下载上方技能包并解压，进入 `xue-skills/`。
2. 将其中的 **`xue/` 和 `xue-plus/` 两个完整目录**放入宿主的技能目录。
3. 按宿主的加载方式刷新技能或开始新会话，确认两个技能能够被发现。

常见目录约定：

| 宿主 | 技能目录 |
| --- | --- |
| Codex | 实际 `CODEX_HOME/skills/`；未自定义时通常为 `~/.codex/skills/` |
| DSH 项目级 | `<项目>/.dsh/skills/` |
| DSH 全局 | `~/.dsh/skills/` |

Windows 中的 `~` 通常对应 `%USERPROFILE%`。DSH 的发现、热重载和策略字段支持取决于实际版本；本仓库未验证全部宿主版本。

不要把仓库根目录 `xue_skill/` 或 ZIP 外层的 `xue-skills/` 直接作为单个技能安装。正确的结构是：

```text
skills/
├── xue/
│   ├── SKILL.md
│   ├── agents/
│   ├── references/
│   └── scripts/
└── xue-plus/
    ├── SKILL.md
    ├── agents/
    ├── references/
    └── scripts/
```

更新已有安装时，先备份并移走旧的同名目录，再整体放入新版，避免残留旧脚本。两个目录各自包含所需资源，可以独立使用。

### 获取源码

```bash
git clone https://github.com/XueHanan/xue_skill.git
cd xue_skill
```

仓库根目录的 `xue/` 与 `xue-plus/` 即为当前 v2.1 的可安装技能。`dist/xue-skills(2.1).zip` 为本次发布包，包内技能文件与仓库当前目录一致；`dist/xue-skills(2.0).zip` 和 `dist/xue-skills(1.0).zip` 保留历史版本。

## 使用示例

```text
用 xue 润色下面这段绪论，给出原文、改后和简要改动。

用 xue 润色论文第二章的理论基础，保留术语、变量和引用。

用 xue-plus 润色第三章方法部分，保留公式、步骤和参数。

用 xue-plus 润色实验分析，不增加因果解释或加强结论。
```

小段文字可以直接粘贴给代理。DOCX 章节任务应说明文件、章节范围及需要保留的专业术语。没有执行脚本的小段润色不会声称通过了自动校验。

## 文件流程与依赖

文件处理脚本仅依赖 **Python 3.10+ 标准库**，无需安装 `python-docx`。无需额外的润色插件。具体操作见 [文件流程](xue/references/workflow.md) 和 [验证说明](xue/references/verification.md)。

以下命令在仓库根目录运行；路径包含空格时加引号，`TASK` 使用独立工作目录。`START`、`END` 必须替换为实际段号。

```text
python xue/scripts/inspect_docx.py SOURCE.docx
python xue/scripts/prepare.py SOURCE.docx TASK --range START END --mode theory --term 需要保留的术语
```

段号从 0 开始，只计正文顶层段落。DOCX 必须明确处理范围；仅在用户授权全文处理时使用 `--whole-document`。纯文本文件可直接传给 `prepare.py`。

代理修改 `TASK/comparison.md` 后，生成并完成语义复核记录：

```text
python xue/scripts/review_template.py TASK/comparison.md TASK
python xue/scripts/validate.py TASK/comparison.md TASK
```

`review_template.py` 只生成待复核模板，不会自动认定含义相同。执行润色的代理需要逐段核对命题、数值和引用归属、条件、步骤及术语，并填写具体依据。

| 验证状态 | 退出码 | 含义 |
| --- | --- | --- |
| `passed` | 0 | 自动约束满足，且修改段具有当前版本的语义复核记录 |
| `failed` | 1 | 结构、范围或受保护内容等检查失败 |
| `needs_review` | 2 | 缺少有效的语义复核记录或仍有未完成的复核 |

结果由 `TASK/validation.json` 指向。正文再次修改后应重新复核并验证，不能沿用旧版的通过状态。`xue-plus/scripts/` 提供相同的文件处理入口。

## 内容保护与边界

- 检查空内容、缺段和越界，保护原文对照、数字、单位、显式术语、变量、图表引用和公式身份。
- 支持 OMML 和可识别的 Equation/MathType OLE 公式标记。表格只读；其他复杂对象、Word 域及未接受的修订采取保守处理，不声称覆盖所有 Word 区域。
- 数字或引用集合相同不代表对应关系正确，仍需要语义复核；复核记录也不等于对论文技术结论的正确性证明。
- 不设固定句长、连接词配额、正常用词禁令或自评分数门槛，不承诺 AI 检测或查重结果。
- 不计算或校验原始论文文件的 SHA-256；内部版本关联只用于提取基线与修订、复核记录。

旧版分散检查命令已由 `validate.py` 取代。旧任务和没有独立公式身份的 `[M]` 标记需要从原文件重新准备。历史语料及其局限见 [语料参考](xue/references/corpus-notes.md)。

## 许可证

本仓库使用 [MIT License](LICENSE)。使用、修改或分发时请保留许可证与版权声明。
