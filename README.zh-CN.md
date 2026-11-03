# PPT Tools

[English](README.md) | 简体中文

`PPTTools.ppam` 为 PowerPoint 添加分段 Section 进度条、Notes 导入导出和代码着色命令。

Ribbon 标签、按钮、提示和错误消息均使用英文。

## 功能

- 顶部使用一行等宽的 Section 进度条。
- Section 名称直接居中显示在对应的进度条区段内。
- 已完成部分使用清晰的淡蓝色，未完成部分使用浅灰色。
- Section 名称使用深墨灰色，章节边界使用置于顶层的 2pt 纯黑矩形条分隔。
- 当前 Section 名称使用蓝色粗体，并显示当前位置标记。
- 进度条末尾使用 `24 × 24pt` 页码格，并只显示当前页码。
- `Update Section Progress` 可在调整页序或 Sections 后同步所有页面。
- `Remove Section Progress` 会移除本加载项生成的图形。
- 未设置 Section 的演示会显示为一个名为 `Presentation` 的区段。
- 名为 `Backup` 的 Section 不显示进度头，并从进度页码中排除。
- Ribbon 使用原生分组线将 `Section Progress` 与右侧的 `Notes` 功能分开。
- `Export Notes to Markdown` 会先选择目录，再把全部幻灯片 notes 导出为一个 UTF-8 Markdown 文件。
- Markdown 使用 `# Slide N` 标题分隔每一页；同名文件使用递增数字后缀。
- `Load Notes from Markdown` 会选择同格式 Markdown，并把每个 `# Slide N` 内容写回对应页面。
- Markdown 中未出现的页面保持原 notes，空 block 会清空对应页面 notes。
- 导入时会移除成对的 `**` 标记，并把其中内容设置为 PowerPoint 原生粗体。
- `Code Prettier` 打开代码编辑窗口，界面与设置参照 [Code to Slides](https://ppt.anemone.top/)。
- 编辑窗口提供语言选择（含 Auto detect），`Format code` 使用 Prettier 格式化 JavaScript、TypeScript、Java、JSON、HTML、CSS 和 Markdown（Java 由 prettier-plugin-java 格式化）。
- 右侧预览使用与幻灯片相同的配色，可设置行号、行号间距和字号（14–40pt）。
- 点击预览中的行号即可高亮该行（默认浅黄色），再次点击取消；高亮行左侧的色块用于修改颜色。幻灯片中的高亮行等宽对齐，覆盖行号和代码。
- `Insert into slide`（⌘Enter）把代码插入当前幻灯片，生成可编辑的 Consolas 着色文本框。
- 选中已插入的代码块后再次选择 `Code Prettier`，编辑窗口会载入原代码和设置，`Update code on slide` 原位更新该代码块。
- 编辑窗口使用本地的 highlight.js、Prettier 和 prettier-plugin-java，离线可用。

## 项目结构

- `src/PPTTools.bas`：VBA 源码。
- `src/customUI14.xml`：PowerPoint Ribbon 定义。
- `scripts/inject_custom_ui.py`：生成 PPAM 并写入 Ribbon 定义。
- `src/code-prettier/PPTToolsCodePrettier.applescript`：`AppleScriptTask` 调用的入口脚本。
- `src/code-prettier/editor/`：Code Prettier 编辑窗口（JXA 窗口宿主、WebKit 页面和 `vendor/` 中的 highlight.js、Prettier、prettier-plugin-java）。
- `scripts/install_code_prettier.sh`：把编辑窗口安装到 PowerPoint 的脚本目录。
- `build/PPTTools.ppam`：可安装的加载项。

## 使用

1. 在 PowerPoint 中打开演示文稿。
2. 打开 `PPT Tools` Ribbon 标签。
3. 选择 `Update Section Progress`。
4. 调整幻灯片或 Sections 后再次选择 `Update Section Progress`。
5. 选择 `Export Notes to Markdown`，指定目录并生成 notes 文件。
6. 选择 `Load Notes from Markdown`，指定 Markdown 文件并写回 notes。
7. 选择 `Code Prettier`，输入或粘贴代码，按需格式化和调整设置，再选择 `Insert into slide`。
8. 选中已插入的代码块并选择 `Code Prettier`，修改后选择 `Update code on slide`。

## 本机安装

加载项安装在：

```text
~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Add-Ins.localized/PPTTools.ppam
```

PowerPoint 已将该加载项设置为自动加载。

Code Prettier 编辑窗口安装在：

```text
~/Library/Application Scripts/com.microsoft.Powerpoint/PPTToolsCodePrettier.scpt
~/Library/Application Scripts/com.microsoft.Powerpoint/PPTToolsCodePrettier/
```

## 源码构建

1. 在 PowerPoint 中新建一个空白演示文稿。
2. 在 Visual Basic 编辑器中导入 `src/PPTTools.bas`，保存为 `build/PPTTools-source.pptm`。
3. 复制源文件并生成 PPAM：

   ```bash
   cp build/PPTTools-source.pptm build/PPTTools.ppam
   python3 scripts/inject_custom_ui.py \
     build/PPTTools.ppam src/customUI14.xml
   ```

4. 在 PowerPoint 的加载项管理界面中添加并启用 `build/PPTTools.ppam`。
5. 安装 Code Prettier 编辑窗口：

   ```bash
   scripts/install_code_prettier.sh
   ```

## 许可证

本项目以 [MIT](LICENSE) 许可证发布。`src/code-prettier/editor/vendor/` 中的 highlight.js（BSD-3-Clause）、Prettier（MIT）和 prettier-plugin-java 打包文件（Apache-2.0 与 MIT）附带各自的许可证文件。
