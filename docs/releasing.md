# 源码与 EXE 发布指南

## 同时使用 Gitee 和 GitHub

OfficeTool 可以同时发布到 Gitee 和 GitHub。建议选择一个主仓库处理 Issue、Pull Request 和版本标签，另一个只做同步镜像。

当前本地仓库若已将 Gitee 配置为 `origin`，可增加 GitHub 远程：

```powershell
git remote add github https://github.com/<账号>/OfficeTool.git
git push -u github master
git push github --tags
```

也可以将远程明确命名为 `gitee` 和 `github`：

```powershell
git remote rename origin gitee
git remote add github https://github.com/<账号>/OfficeTool.git
git push gitee master --tags
git push github master --tags
```

新建 GitHub 仓库时不要再次初始化 README、LICENSE 或 `.gitignore`，否则首次推送容易产生无关历史冲突。

## 公开前检查

1. 确认测试通过，GUI 能正常启动。
2. 检查工作区中没有真实公文、报告、截图、内网地址和用户配置。
3. 搜索 `sk-`、`api_key`、`.env`、个人邮箱和本机绝对路径。
4. 检查全部 Git 历史的作者邮箱和历史文件，而不只是当前目录。
5. 更新版本号、README 和变更说明。
6. 创建版本标签，例如 `v0.1.0`。

如果敏感信息曾经提交过，仅删除当前文件还不够；必须重写 Git 历史、轮换已经泄露的密钥，并强制更新远程仓库。

## 构建 Windows EXE

建议在干净的 Windows 环境或全新克隆目录中构建，避免把本机配置带入产物。开发阶段优先使用 PyInstaller 的文件夹模式，启动更快，也更容易排查依赖：

```powershell
py -3.11 -m venv .venv-build
.\.venv-build\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install . pyinstaller
Push-Location src
python -m PyInstaller --noconfirm --clean --onefile --windowed --name OfficeTool `
  --paths . `
  --hidden-import pythoncom `
  --hidden-import pywintypes `
  --hidden-import win32com.client `
  --distpath ..\dist `
  --workpath ..\build\pyinstaller `
  --specpath ..\build `
  office_tool\gui_entry.py
Pop-Location
Compress-Archive -Path dist\OfficeTool.exe,README.md,CHANGELOG.md,LICENSE -DestinationPath OfficeTool-v0.1.0-windows-x64.zip
Get-FileHash OfficeTool-v0.1.0-windows-x64.zip -Algorithm SHA256
```

正式发布前至少在一台未安装 Python 的 Windows 电脑上验证：

- 双击启动和窗口显示正常。
- 普通公文、红头、函、会议纪要均能导出。
- 在装有 Word 或 WPS 的目标电脑上验证 `.doc` 可离线转换并输出 `.docx`。
- 输出目录、AI 配置、自定义配置和异常提示正常。
- EXE 中不包含 API Key、真实文档、用户配置和测试产物。
- 目标电脑具备所需中文字体；EXE 不应私自捆绑字体文件。

如需单文件 EXE，可将 PyInstaller 参数改为 `--onefile`，但启动较慢，杀毒软件误报概率也通常更高。公开分发时建议进行代码签名，并同时提供 SHA256 校验值。

## 上传下载包

不要把 EXE 或 ZIP 直接提交到源码分支。推荐流程：

1. 推送源码和版本标签。
2. 在 GitHub 创建 Release，选择对应标签，填写变更说明并上传 ZIP 和 SHA256。
3. 在 Gitee 创建发行版并上传相同附件；如果当前账户或仓库不支持附件，可在 Gitee README 中链接到 GitHub Release。
4. README 只保留“最新版本下载”链接，不提交二进制文件本体。

每次发布使用唯一版本号，不覆盖旧附件，便于用户回退和核对来源。
