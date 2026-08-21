# Mac Office Tools

macOS 菜单栏工具。点击菜单栏中的文档图标，可按 **Word**、**Excel**、**PowerPoint** 分组查看当前窗口标题；没有打开文件的应用不会显示。超长文件标题会保留开头与结尾，并以“…”连接。将鼠标停留在文件标题上，可在右侧展开的菜单中打开对应窗口。刷新上方的“新建”菜单可创建 Word、Excel 或 PowerPoint 文件，即使对应 Office 尚未运行也会自动启动。

每个已打开的文件标题右侧都会展开操作菜单，可选择在 Finder 中显示、复制或关闭。点击“在 Finder 中显示”可直接定位已保存的文件；点击“复制”即可将其作为文件对象复制到 macOS 剪贴板，切换到聊天框后按 **⌘V** 即可粘贴附件。成功、未保存和失败都会显示短暂提示。

## 当前发布

当前版本为 **v3.31**，适用于 Apple Silicon Mac 和 macOS 13 或更高版本。菜单整体宽度固定为 220pt；超长文件标题最多使用 156pt，并以位于文字区中央的“…”保留标题开头和结尾。文件行、刷新和退出均使用 macOS 原生菜单选中效果。

## 安装

从 [GitHub Releases](https://github.com/Only-xueluo/Mac-Office-Tools/releases/latest) 下载 `Mac_Office_Tools_v3.31_macOS_Apple_Silicon.zip`，解压后将 **Mac Office Tools.app** 移至“应用程序”。首次启动时，请在“系统设置 → 隐私与安全性 → 辅助功能”中允许 **Mac Office Tools**；首次使用“复制”或“在 Finder 中显示”时，请允许其控制对应的 Office 应用。

## 隐私与权限

程序仅通过 macOS“辅助功能”接口读取窗口标题、恢复最小化窗口并将窗口前置；不会读取、修改或保存 Office 文档内容。点击“复制”或“在 Finder 中显示”时，程序会请求 macOS 的“自动化”权限，用于取得已保存 Office 文件的位置。

首次启动时，程序会显示权限说明，并可直接打开“系统设置 → 隐私与安全性 → 辅助功能”。启用“Mac Office Tools”后，菜单栏图标即可列出 Office 窗口；首次选择“复制”或“在 Finder 中显示”时，请允许其控制 Microsoft Excel、Word 或 PowerPoint。
