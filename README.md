# RY-wps-right-click-to-pdf

在 Windows 资源管理器里右键把 Word 转成 PDF。选中一个或多个 `.doc`、`.docx` 文件，WPS 文字会自动打开并依次导出。

## 安装

需要 Windows 和已安装的 WPS Office 桌面版。已在 Windows 11 上验证。脚本使用 Windows 自带的 .NET Framework 编译程序，将菜单安装到当前 Windows 用户，无需管理员权限。

1. 从 [发布页](https://github.com/jedliuai/RY-wps-right-click-to-pdf/releases/latest) 下载 **Source code (zip)** 并解压。
2. 双击解压目录里的 `Install.cmd`。

也可以在项目目录通过 PowerShell 安装。

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Setup.ps1
```

## 使用

在资源管理器选中一个或多个 `.doc`、`.docx` 文件，右键选择 **显示更多选项 → 转换为 PDF（WPS）**。也可以选中文件后按 **Shift + F10** 打开经典右键菜单。

- `.doc` 和 `.docx` 可以混选，一次最多 100 个文件。 [Microsoft 的多选菜单说明](https://learn.microsoft.com/en-us/windows/win32/shell/how-to-employ-the-verb-selection-model)
- PDF 与原文件保存在同一个文件夹。`报价单.docx` 会生成 `报价单.pdf`，同名时自动加 ` (2)`、` (3)` 等编号。
- 转换结束后会显示成功数量和失败原因。

如果电脑启用了绿盾等文档加密软件，输出的 PDF 也可能受本机加密策略保护。

## 卸载

关闭转换结果窗口后，双击 `Uninstall.cmd`，或在项目目录运行以下命令。

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Setup.ps1 -Uninstall
```

## 问题反馈

如果菜单没有出现，先用 **Shift + F10** 打开经典右键菜单。如果安装或转换遇到问题，可以提交 Issue，并附上 Windows 版本、WPS 版本、文件格式和错误提示。 [提交 Issue](https://github.com/jedliuai/RY-wps-right-click-to-pdf/issues/new)

## 作者网站

代充网站 [aishopvip.com](https://aishopvip.com)。
