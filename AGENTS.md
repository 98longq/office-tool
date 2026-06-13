# AGENTS.md

## 项目定位

OfficeTool 是面向文秘和日常办公场景的本地办公助手。当前核心功能是公文审计与格式处理，优先支持普通党政机关公文和红头文件；Excel 小工具后续作为独立模块扩展。

## 当前代码结构

- `src/office_tool/config.py`：默认配置，包含页面参数、审计选项、格式化选项和各类段落样式。
- `src/office_tool/audit.py`：公文结构识别与审计规则，输出 `AuditReport`。
- `src/office_tool/formatter.py`：基于审计结果执行页面设置、字体字号、标题层级、红头、页码等格式处理。
- `src/office_tool/io.py`：输入加载，当前支持 `.docx/.txt/.md`。
- `src/office_tool/reports.py`：JSON/Markdown 审计报告导出。
- `src/office_tool/ai/deepseek.py`：可选 DeepSeek 文本 AI 审查，使用 OpenAI-compatible chat completions。
- `src/office_tool/gui.py`：最小可用 Tkinter 桌面界面，支持选择文件、审计、格式化、写出报告。
- `src/office_tool/cli.py`：命令行入口。
- `src/office_tool/excel/`：Excel 小工具预留边界，暂未实现具体功能。
- `docs/official_document_rules.md`：本次实现采用的公文格式规则摘要。
- `tests/`：使用 `unittest` 的审计、格式化、配置测试。
- `wfp.py`、`wfp_cli.py`：旧项目兼容入口，转发到新 CLI。

## 常用命令

使用本机或 Codex bundled Python 均可；当前环境可用：

```powershell
& 'C:\Users\10915\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m unittest discover -s tests -v
```

GUI：

```powershell
python wfp.py
```

审计：

```powershell
python wfp_cli.py audit input.docx --json report.json --markdown report.md
```

格式化：

```powershell
python wfp_cli.py format input.docx -o output.docx --audit-json report.json
```

DeepSeek AI 审查：

```powershell
python wfp_cli.py audit input.docx --ai-review --ai-base-url http://deepseek.local:8000/v1 --ai-model deepseek-chat --json report.json
```

AI 审查默认关闭。内网模型如需鉴权，默认读取 `DEEPSEEK_API_KEY`，也可通过 `--ai-api-key-env` 或配置 `ai_review.api_key_env` 修改。

生成配置：

```powershell
python -m office_tool init-config -o office_tool_config.json
```

## 公文规则实现现状

已覆盖：

- A4、页边距、页脚距、固定行距等页面设置。
- 发文机关标志、发文字号、签发人、标题、主送机关、正文、附件说明、署名、日期、抄送/印发等结构识别。
- 红头文件版头红色字体、发文字号居中、红色分隔线。
- 标题层级：`一、`、`（一）`、`1.`、`（1）`。
- 审计报告的错误、警告、提示分级。
- 可选 DeepSeek AI 文本审查，补充语义、措辞、一致性和风险表述检查。

暂未覆盖或只做提示：

- `.doc/.wps` 旧格式自动转换。
- 印章位置、红线精确长度、联合行文特殊版式、命令/纪要/信函格式的全部细项。
- 每页 22 行、每行 28 字的视觉级分页校验。
- GUI。

## 开发注意事项

- 不要把 Excel 工具写进公文处理器；在 `src/office_tool/excel/` 下单独建模块。
- 新增公文规则时，优先在 `audit.py` 增加可解释审计，再在 `formatter.py` 增加可修复行为。
- AI 审查只能作为补充层，不要替代确定性格式规则；DeepSeek 调用必须保持可配置、默认关闭，适配内网离线部署。
- 规则必须可配置，默认值放在 `config.py`。
- 所有核心行为都要补 `unittest`，尤其是红头版头不能被误识别为标题。
- 旧入口 `wfp.py`、`wfp_cli.py` 尽量保持可用，便于迁移旧脚本。

## 本次任务完成情况

- 重建项目代码，迁移到 `src/office_tool` 包结构。
- 实现公文审计、红头识别、结构化报告和基础格式修复。
- 增加 DeepSeek AI 文本审查模块和 CLI 配置入口。
- 补回最小 GUI 入口，并增加源码树 `python -m office_tool` 启动 shim。
- 增加默认配置、CLI、兼容入口、规则文档和测试。
- 删除旧版单文件 GUI/Skill 资产，后续如需要 GUI 应基于新服务层重建。
