# OfficeTool 办公助手

OfficeTool 是面向文秘和日常办公场景的本地办公助手。当前核心能力是公文审计与格式处理，优先支持普通党政机关公文和红头文件；Excel 小工具作为独立模块扩展，避免和公文处理逻辑耦合。

## 当前能力

- 审计 `.docx/.txt/.md`，识别红头、发文字号、签发人、标题、主送机关、正文、附件、署名日期、抄送和印发等要素。
- 按常见公文格式参数修复页面、页边距、正文、标题层级、红头版头、发文字号行、页码和基础段落样式。
- 输出 JSON 和 Markdown 审计报告，适合批处理、归档和后续 GUI 展示。
- 可选接入内网 DeepSeek，补充语义、措辞、一致性和风险表述审查。AI 审查默认关闭。
- 提供 Tkinter 桌面界面、命令行、批量处理服务层和旧入口兼容。
- Excel 模块已提供工作簿概况检查、文本首尾空白清洗、空行清理。

## 安装

```powershell
python -m pip install -r requirements.txt
python -m pip install -e .
```

没有安装为包时，也可以直接使用仓库内兼容入口：

```powershell
python wfp.py
python wfp_cli.py --help
```

## 快速试用

启动桌面界面：

```powershell
python wfp.py
```

审计单个文档：

```powershell
python wfp_cli.py audit .\example.docx --json .\report.json --markdown .\report.md
```

格式化单个文档：

```powershell
python wfp_cli.py format .\example.docx -o .\example_formatted.docx --audit-json .\report.json
```

批量审计目录：

```powershell
python wfp_cli.py batch-audit .\docs -r .\reports --markdown
```

批量格式化目录：

```powershell
python wfp_cli.py batch-format .\docs -o .\out -r .\reports --markdown
```

检查 Excel 工作簿：

```powershell
python wfp_cli.py excel inspect .\data.xlsx --json .\excel_report.json
```

清洗 Excel：

```powershell
python wfp_cli.py excel clean .\data.xlsx -o .\data_clean.xlsx
```

生成默认配置：

```powershell
python -m office_tool init-config -o office_tool_config.json
```

命令行覆盖配置：

```powershell
python -m office_tool format .\example.docx -o .\out.docx --set styles.body.font=仿宋_GB2312 --set page.margin_top_cm=3.7
```

## DeepSeek AI 审查

AI 审查作为确定性规则的补充层，不替代格式规则。内网 OpenAI-compatible DeepSeek 服务可这样接入：

```powershell
python wfp_cli.py audit .\example.docx --ai-review --ai-base-url http://deepseek.local:8000/v1 --ai-model deepseek-chat --json .\report.json
```

如服务需要鉴权，默认读取 `DEEPSEEK_API_KEY`；也可以通过 `--ai-api-key-env` 或配置文件里的 `ai_review.api_key_env` 修改。

## 代码结构

- `src/office_tool/config.py`：默认配置和配置覆盖。
- `src/office_tool/audit.py`：公文结构识别与审计规则。
- `src/office_tool/formatter.py`：页面、段落、标题、红头和页码格式修复。
- `src/office_tool/services.py`：单文件和批量处理服务层。
- `src/office_tool/io.py`：文档加载，支持 `.docx/.txt/.md`。
- `src/office_tool/reports.py`：JSON/Markdown 审计报告导出。
- `src/office_tool/ai/deepseek.py`：DeepSeek 文本 AI 审查。
- `src/office_tool/excel/`：Excel 小工具。
- `src/office_tool/gui.py`：本地 Tkinter 桌面界面。
- `wfp.py`、`wfp_cli.py`：旧项目兼容入口。
- `docs/official_document_rules.md`：当前采用的公文格式规则摘要。
- `tests/`：`unittest` 测试。

## 验证

```powershell
python -m compileall src office_tool wfp.py wfp_cli.py
python -m unittest discover -s tests -v
```

当前未实现 `.doc/.wps` 自动转换、印章位置视觉级校验、每页 22 行每行 28 字的真实分页校验，以及完整 Web/企业权限体系。这些应作为后续迭代单独设计。
