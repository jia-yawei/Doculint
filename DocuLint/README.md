# DocuLint 工程说明

这个解决方案目前有 3 个项目：

- `DocuLint`
  Word 的 VSTO 插件主工程。只有这个工程直接操作 Word 对象模型。
- `DocuLint.Host.Abstractions`
  Word 插件内部共享的接口和模型。
- `DocuLint.Core`
  宿主无关的业务逻辑，目前主要放可复用的数据整理逻辑。

## 建议阅读顺序

如果你是第一次看这个工程，建议按这个顺序读：

1. `DocuLint/ThisAddIn.cs`
2. `DocuLint/Ribbon/Ribbon1.cs`
3. `DocuLint/Features/*`
4. `DocuLint/HostAdapters/WordDocumentHostAdapter.cs`
5. `DocuLint.Core`
6. `DocuLint.Host.Abstractions`

## 主工程目录

`DocuLint/DocuLint` 现在按职责分成这些区域：

- `Ribbon/`
  VSTO Ribbon 外壳和设计器文件。这里放入口，不放具体业务细节。
- `Features/BatchReplace/`
  批量替换功能。
- `Features/Captions/`
  题注列表窗格和题注相关 Ribbon 操作。
- `Features/Outline/`
  多级列表/章节号重建功能。
- `HostAdapters/`
  Word 专属适配层，把 Word API 转成共享接口。
- `Properties/`
  Visual Studio 自动生成的资源和设置。
- `resource/`
  运行时依赖的文本资源。

## 当前哪些是“在用”的

- Word 插件主流程：在用
- `DocuLint.Core` / `DocuLint.Host.Abstractions`：部分在用，后续会继续承接更多共享逻辑

## 为什么之前会显得乱

原来主工程里混着几种不同层次的代码：

- VSTO 入口代码
- Ribbon UI 入口
- 具体功能实现
- 窗格和对话框
- Word 宿主适配代码

这些代码都在一个平级目录里时，新手很难一下看出边界。现在已经按“入口 / 功能 / 宿主适配”分开了，阅读成本会低很多。

## 构建说明

- `DocuLint.Core`、`DocuLint.Host.Abstractions` 可以用 `dotnet build` 验证。
- `DocuLint` 这个 Word VSTO 工程不能直接用 `dotnet build`，要用 Visual Studio 或 .NET Framework 版 `MSBuild.exe`。

示例：

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "DocuLint\DocuLint.slnx" /t:Build /p:Configuration=Debug
```
