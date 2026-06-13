# OfficeTool 办公助手

OfficeTool 是一个面向常见办公场景的本地工具项目。当前阶段优先完成核心能力：**公文审计与公文格式处理**。后续 Excel 清洗、表格核对、批量生成等小工具会作为独立模块接入，避免把所有能力混在同一个处理器里。

## 当前能力

- 审计 `.docx`、`.txt`、`.md` 文档，识别普通公文和红头文件的版头、发文字号、签发人、标题、主送机关、正文、附件、署名日期、版记等关键要素。
- 按 GB/T 9704-2012 常用参数修复页面与段落格式：A4、页边距、正文三号仿宋、标题层级、28 磅行距、两字首行缩进、页码等。
- 对红头文件单独处理发文机关标志、红色版头、发文字号行和红色分隔线。
- 输出结构化审计报告，支持 JSON 和 Markdown，方便后续接 GUI、批量任务或归档。
- 可选接入内网 DeepSeek，对公文文本做 AI 审查，补充规则审计难覆盖的语义、措辞和一致性问题。
- 保留旧入口 `wfp.py`、`wfp_cli.py`，但主体代码已迁移到 `src/office_tool`。

## 安装依赖

```powershell
python -m pip install -r requirements.txt
python -m pip install -e .
```

## 使用方式

启动桌面界面：

```powershell
python wfp.py
```

直接审计文档：

```powershell
python wfp_cli.py audit .\example.docx --json .\report.json --markdown .\report.md
```

审计并生成格式化后的文档：

```powershell
python wfp_cli.py format .\example.docx -o .\example_formatted.docx --audit-json .\report.json
```

使用内网 DeepSeek 做 AI 文本审查：

```powershell
python wfp_cli.py audit .\example.docx --ai-review --ai-base-url http://deepseek.local:8000/v1 --ai-model deepseek-chat --json .\report.json
```

如果内网服务需要鉴权，默认读取 `DEEPSEEK_API_KEY`；也可以用 `--ai-api-key-env` 指定企业环境变量名。AI 审查默认关闭，未配置时不会访问网络。

生成默认配置：

```powershell
python -m office_tool init-config -o office_tool_config.json
```

按需覆盖配置：

```powershell
python -m office_tool format .\example.docx -o .\out.docx --set styles.body.font=仿宋 --set page.margin_top_cm=3.7
```

旧兼容入口：

```powershell
python wfp_cli.py audit .\example.docx
python wfp.py format .\example.docx -o .\out.docx
```

## 代码结构

- `src/office_tool/config.py`：页面、样式、公文审计配置。
- `src/office_tool/audit.py`：公文结构识别和审计规则。
- `src/office_tool/formatter.py`：页面、段落、标题、红头和页码修复。
- `src/office_tool/ai/deepseek.py`：内网 DeepSeek 文本审查客户端。
- `src/office_tool/io.py`：输入文档加载，当前支持 `.docx/.txt/.md`。
- `src/office_tool/reports.py`：审计报告导出。
- `src/office_tool/cli.py`：命令行入口。
- `src/office_tool/excel/`：Excel 工具预留边界。
- `docs/official_document_rules.md`：当前实现依据的公文格式要点。
- `tests/`：核心审计、格式化和 CLI 配置测试。

## 验证

```powershell
python -m unittest discover -s tests -v
```

## 后续路线

1. 将 `.doc/.wps` 通过 Word/WPS COM 或 LibreOffice 接入为可选转换能力。
2. 增加 GUI 或 Web 本地界面，围绕“审计结果 - 一键修复 - 人工确认”工作流组织。
3. 扩展审计规则：版记细项、印章区、联合行文、信函格式、命令格式、纪要格式。
4. 新增 Excel 小工具模块，优先做批量清洗、重复值检查、跨表核对和模板填报。
