# 搞快点（DocuLint）

搞快点是一个面向 Microsoft Word 的 VSTO 文档处理插件，主要用于需求文档整理、章节编号维护、文档格式检查和重复性编辑操作。

## 主要功能

- 需求提取：从 Word 文档中识别需求标识、名称和章节号，支持常规提取与自定义提取。
- 需求追踪：维护需求之间的正向、反向追踪关系，支持候选推荐、暂存和追踪结果导出。
- 章节号修复：检查标题编号连续性，并按文档标题级别修复异常编号。
- 样式管理：管理常用样式、修复标题样式，并实时显示当前段落样式。
- 文档体检：检查多级列表连续性及其他文档格式问题。
- 批量替换：对多个 Word 文档执行批量文本替换。
- 表格与题注工具：按续表拆分表格、补充表头、管理题注和交叉引用。
- 快速工具：插入页码或编号、格式设置、导航窗格、常用语和域编号等。
- 版本管理：在文档旁保存版本存档、备注和修改记录，并支持查看、比较和恢复版本。
- 在线更新：默认从 GitHub 官方仓库检查更新，也支持配置内网更新文件夹。

## 运行环境

- Windows
- Microsoft Word（支持 VSTO 加载项）
- .NET Framework 4.7.2 或更高版本
- Visual Studio（安装 Office/SharePoint 开发工具）
- Visual Studio Tools for Office Runtime（VSTO Runtime）

插件只面向 Word 文档对象模型设计，未安装 Word 或 VSTO Runtime 时无法正常加载。

## 项目结构

```text
DocuLint/
  DocuLint/                  VSTO 插件主工程和 Word 交互代码
  DocuLint.Core/             与宿主无关的共享业务逻辑
  DocuLint.Host.Abstractions/宿主适配接口和共享模型
docs/                        设计资料和测试文档
scripts/                     开发辅助脚本
```

## 构建

使用 .NET Framework 版 MSBuild 构建 VSTO 工程：

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "DocuLint\DocuLint.slnx" `
  /t:Build /p:Configuration=Debug /p:Platform=AnyCPU /m
```

也可以在 Visual Studio 中打开 `DocuLint/DocuLint.slnx`，选择 `Debug` 或 `Release` 后构建。调试启动会打开 Word 并加载插件；目标电脑需要安装对应的 Office 和 VSTO Runtime。

## 更新插件

在功能区“关于”菜单中点击“检查更新”，可选择两种方式：

- 网络更新：选择“网络更新”后点击“检查更新”，发现新版本时点击“安装更新”。
- 本地更新：选择“本地更新”，点击“浏览”选择独立安装包 `.exe` 或 `.msi` 文件，再点击“检查更新”和“安装更新”。本地安装包文件名需包含版本号，例如 `DocuLint-0.0.2.0-setup.exe`。

安装完成后请重新打开 Word。

## 使用注意

涉及批量替换、章节号修复、样式调整和版本恢复的操作可能会修改文档内容。执行前请保存并备份原始文档，先在副本上验证结果。

## 许可与责任

本项目用于内部文档处理和效率提升。使用插件前请根据组织要求进行测试和备份，因文档修改造成的数据损失由使用者自行承担。
