# 参与贡献

感谢参与 OfficeTool。项目当前只维护文档模块，请不要恢复已移除的 Excel、日志页或历史打包产物。

## 开发原则

- 先使用真实但已脱敏的公文结构复现问题，再修改规则。
- 确定性格式规则优先于 AI；AI 失败不得阻断 Word 文件导出。
- 不猜测发文机关、文号、抄送单位、印发日期等业务字段。
- 不在源码、测试、Issue、截图或配置样例中提交真实公文、API Key、内网地址和个人信息。
- UI 和用户文档统一使用“校对”。
- 测试产物只能写入 `tests/artifacts/`，不得提交生成的 DOCX、PDF、PNG。

## 提交前检查

```powershell
python -m unittest discover -s tests -v
python -m compileall -q src tests wfp.py
git diff --check
```

涉及版式、红头、页码或版记的改动，还应使用 Word、WPS 或 LibreOffice 检查打印预览。

## 提交内容

- 一个提交尽量只解决一个明确问题。
- 新增或修复核心行为时同步增加 `unittest`。
- 行为、配置项或使用方式变化时同步更新 README 和规则文档。
- 提交前检查 Git 作者邮箱；不希望公开真实邮箱时使用代码托管平台提供的 noreply 地址。
