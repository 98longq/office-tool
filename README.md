# OfficeTool

> 开发说明：本项目采用 Vibe Coding 方式开发，大量代码由开发者结合实际公文样例、人工复核结果和 AI 编程工具共同迭代完成。代码和排版规则可能仍有遗漏，使用前请结合本单位制度测试，重要公文必须由人员最终复核。

OfficeTool 是一个面向 Windows 的本地公文校对与格式整理工具。当前版本专注“文档”模块，正式支持普通公文、普通红头文件、红头文件（函）和红头文件（会议纪要），并支持普通红头附带制度正文。

项目以确定性规则处理页面、字体、段落和版记。AI 文本校对是可选能力，默认关闭；不开启 AI 时，文档内容不会因本项目发送到网络。

## 下载

- [GitHub Releases](https://github.com/98longq/office-tool/releases)
- [Gitee 发行版](https://gitee.com/longq98/office-tool/releases)

Windows 用户可下载版本化的 `.exe` 文件直接运行，不需要安装 Python。处理旧版 `.doc` 时，电脑需要安装 Microsoft Word 或支持 COM 自动化的 WPS。

## 功能概览

- 导入 `.doc`、`.docx`、`.txt`、`.md` 文件，或直接在界面输入文稿。
- 旧版 `.doc` 通过电脑已有的 Microsoft Word 离线转换，Word 不可用时自动尝试 WPS；中间文件只保存在临时目录。
- 支持单文件和文件夹批量校对、格式整理与 Word 导出。
- 设置 A4 页面、页边距、每行 28 字、每页 22 行和奇偶页页码。
- 处理主标题、标题日期、主送机关、正文层级、附件、落款、日期和版记。
- 普通正文使用仿宋_GB2312 三号和单倍行距；一级标题使用黑体，二级标题使用楷体，三级、四级标题使用仿宋。
- 标题下单日期或起止日期使用正文仿宋三号居中，不作为正文缩进。
- 附件长名称支持悬挂缩进，使换行文字与附件名称起点对齐。
- 支持普通红头的内部资料提示、华文中宋版头、文号、红线和普通正文。
- 支持红头附带制度正文的制度编号、章标题、条文混排和版记。
- 支持函的正文锚定文本框版头、复合红线、函号、首页隐藏页码和联系人信息。
- 支持会议纪要专用版头、期号、编发信息、出席人员和分送版记。
- 红头方案可按需生成完整缺失的红头或版记；检测到完整或部分既有结构时不会重复生成。
- 自定义配置保存在 `%APPDATA%\OfficeTool\profiles`，可以应用、重命名、删除、导入和导出。
- 可选接入 DeepSeek 或 OpenAI-compatible 服务，补充错漏字、语病、前后不一致和风险表述建议。

完整确定性规则参见 [docs/official_document_rules.md](docs/official_document_rules.md)。

## 使用边界

- 当前只维护“文档”模块；界面中的“表格”和“其他”仅显示开发中提示。
- `.doc` 输入需要目标电脑已安装 Microsoft Word 或支持 COM 自动化的 WPS，最终统一输出 `.docx`；暂不处理 `.wps`、扫描 PDF 和图片。
- 程序可以整理已有内容，也可以按用户填写的业务字段生成红头和版记，但不会猜测发文机关、文号、抄送单位等内容。
- 印章位置、联合行文、涉密文件全部合规项及复杂分页仍需人工复核。
- Word、WPS 和 LibreOffice 的排版引擎存在差异，最终版式应以实际使用的 Office 软件和打印预览为准。
- 仿宋_GB2312、楷体_GB2312、华文中宋等字体需要在运行电脑上可用；项目和 EXE 不分发这些字体。

## Windows 源码运行

需要 Python 3.10 或更高版本。

```powershell
git clone <你的仓库地址>
cd OfficeTool
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python wfp.py
```

也可以安装为可编辑包：

```powershell
python -m pip install -e .
office-tool-gui
```

## 命令行

```powershell
python wfp.py proofread input.docx --json report.json
python wfp.py format input.docx -o output.docx
python wfp.py batch-format docs -o output -r reports --markdown
python wfp.py show-config
```

执行 `python wfp.py --help` 可查看全部参数。命令行中的 `proofread` 对应“校对”，保留英文命令名便于脚本调用。

## AI 校对与隐私

AI 校对默认关闭。启用后，程序会把截取后的文档纯文本发送到用户配置的 AI 服务地址，因此：

1. 只应使用单位允许的内网模型或已获准的第三方服务。
2. 不要向未经授权的服务发送涉密、敏感或个人信息。
3. API Key 优先通过 GUI 保存或环境变量提供，不要写进源码和公开 JSON。

GUI 中保存的 Key 使用当前 Windows 用户的 DPAPI 加密，存放在本机用户目录；导出的通用 JSON 配置会清空 Key。也可临时使用环境变量：

```powershell
$env:DEEPSEEK_API_KEY = "你的密钥"
python wfp.py
```

项目不包含统计、遥测或自动上传功能。AI 服务不可用时，确定性格式整理和文件保存仍会继续。

`.doc` 转换同样完全离线。程序会禁用 Office 宏、只读打开源文件并在临时目录生成 `.docx`，但旧版 Office 文件本身可能包含恶意内容，请只处理可信来源的文件。

## 项目结构

```text
OfficeTool/
├─ src/office_tool/          正式 Python 包
│  ├─ audit.py               结构识别与确定性校对规则
│  ├─ formatter.py           DOCX 页面和文字格式处理
│  ├─ generator.py           可选红头与版记生成
│  ├─ legacy_doc.py          Word/WPS 离线 DOC 转换
│  ├─ docx_utils.py          OOXML 辅助函数
│  ├─ services.py            GUI/CLI 共用服务层
│  ├─ profile_store.py       自定义配置持久化
│  ├─ secret_store.py        Windows DPAPI 密钥保护
│  ├─ ai/                    可选 AI 文本校对
│  └─ gui.py                 Tkinter 桌面界面
├─ tests/                    unittest 自动化测试
├─ docs/                     规则和发布文档
├─ wfp.py                    源码运行入口
├─ pyproject.toml            Python 包元数据
└─ requirements.txt          运行依赖
```

测试生成的 DOCX、PDF、PNG 统一写入 `tests/artifacts/`，该目录内容不会提交到 Git。

## 开发与测试

```powershell
python -m unittest discover -s tests -v
python -m compileall -q src tests wfp.py
```

视觉复核需要本机安装 LibreOffice 或 Microsoft Word。验证文档可运行：

```powershell
python tests/generate_final_verification_docs.py
```

贡献代码前请阅读 [CONTRIBUTING.md](CONTRIBUTING.md)。安全和隐私问题参见 [SECURITY.md](SECURITY.md)。

## Gitee、GitHub 与 EXE

同一套源码可以同时公开在 Gitee 和 GitHub。建议将其中一个作为主仓库，另一个作为镜像，避免两边同时合并不同提交。

EXE 不应直接提交进 Git 源码历史。正确做法是创建版本标签和 Release/发行版，将 `OfficeTool-vX.Y.Z-windows-x64.zip`、校验值和更新说明作为发布附件，用户即可在网页下载。双平台同步和 EXE 发布步骤参见 [docs/releasing.md](docs/releasing.md)。

## 许可证

本项目使用 [MIT License](LICENSE)。公文标准、字体软件、Office/WPS/LibreOffice 及第三方 AI 服务分别受其自身条款约束。
