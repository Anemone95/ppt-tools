# PPT Tools

English | [简体中文](README.zh-CN.md)

`PPTTools.ppam` adds section progress bars, notes import and export, and syntax-highlighted code to PowerPoint.

The Ribbon tab, buttons, tooltips and error messages are in English.

## Features

- A single row of equal-width section progress segments runs along the top of each slide.
- Each section name is centered in its own progress segment.
- Completed sections use a clear light blue; upcoming sections use light gray.
- Section names use a dark ink gray, and 2pt solid black bars on top separate the sections.
- The current section name is bold blue and shows a current-position marker.
- The progress bar ends with a `24 × 24pt` page number cell that shows the current page number.
- `Update Section Progress` syncs all slides after you reorder slides or change sections.
- `Remove Section Progress` removes the shapes this add-in created.
- A presentation without sections shows as one section named `Presentation`.
- A section named `Backup` shows no progress header and is left out of the progress page numbers.
- A native Ribbon group separator divides `Section Progress` from the `Notes` commands on its right.
- `Export Notes to Markdown` asks for a folder, then exports the notes of all slides to one UTF-8 Markdown file.
- The Markdown separates slides with `# Slide N` headings; a file with the same name gets an increasing number suffix.
- `Load Notes from Markdown` reads a Markdown file in the same format and writes each `# Slide N` block back to its slide.
- Slides missing from the Markdown keep their notes, and an empty block clears that slide's notes.
- On import, paired `**` markers are removed and the text between them becomes native PowerPoint bold.
- `Code Prettier` opens a code editor window whose layout and settings follow [Code to Slides](https://ppt.anemone.top/).
- The editor has a language picker (including Auto detect), and `Format code` formats JavaScript, TypeScript, Java, JSON, HTML, CSS and Markdown with Prettier (Java through prettier-plugin-java).
- The preview on the right uses the slide's colors and has settings for line numbers, line number gap and font size (14–40pt).
- Click a line number in the preview to highlight that line (light yellow by default), and click it again to remove the highlight; the swatch to the left of a highlighted line changes its color. On the slide, highlighted lines share one width and cover both the line number and the code.
- `Insert into slide` (⌘Enter) inserts the code into the current slide as an editable, colored Consolas text box.
- Select an inserted code block and choose `Code Prettier` again: the editor loads the original code and settings, and `Update code on slide` updates the block in place.
- The editor uses local copies of highlight.js, Prettier and prettier-plugin-java, so it works offline.

## Project structure

- `src/PPTTools.bas`: VBA source.
- `src/customUI14.xml`: PowerPoint Ribbon definition.
- `scripts/inject_custom_ui.py`: builds the PPAM and writes the Ribbon definition into it.
- `src/code-prettier/PPTToolsCodePrettier.applescript`: entry script that `AppleScriptTask` calls.
- `src/code-prettier/editor/`: Code Prettier editor window (JXA window host, WebKit page, and highlight.js, Prettier and prettier-plugin-java in `vendor/`).
- `scripts/install_code_prettier.sh`: installs the editor window into PowerPoint's script folder.
- `build/PPTTools.ppam`: the installable add-in.

## Usage

1. Open a presentation in PowerPoint.
2. Open the `PPT Tools` Ribbon tab.
3. Choose `Update Section Progress`.
4. After you change slides or sections, choose `Update Section Progress` again.
5. Choose `Export Notes to Markdown` and pick a folder to create the notes file.
6. Choose `Load Notes from Markdown` and pick a Markdown file to write the notes back.
7. Choose `Code Prettier`, type or paste code, format it and adjust the settings as needed, then choose `Insert into slide`.
8. Select an inserted code block, choose `Code Prettier`, make your changes, then choose `Update code on slide`.

## Local installation

The add-in is installed at:

```text
~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Add-Ins.localized/PPTTools.ppam
```

PowerPoint loads the add-in automatically.

The Code Prettier editor window is installed at:

```text
~/Library/Application Scripts/com.microsoft.Powerpoint/PPTToolsCodePrettier.scpt
~/Library/Application Scripts/com.microsoft.Powerpoint/PPTToolsCodePrettier/
```

## Build from source

1. Create a blank presentation in PowerPoint.
2. In the Visual Basic Editor, import `src/PPTTools.bas` and save the presentation as `build/PPTTools-source.pptm`.
3. Copy the source file and build the PPAM:

   ```bash
   cp build/PPTTools-source.pptm build/PPTTools.ppam
   python3 scripts/inject_custom_ui.py \
     build/PPTTools.ppam src/customUI14.xml
   ```

4. In PowerPoint's add-in manager, add and enable `build/PPTTools.ppam`.
5. Install the Code Prettier editor window:

   ```bash
   scripts/install_code_prettier.sh
   ```

## License

This project is released under the [MIT](LICENSE) license. highlight.js (BSD-3-Clause), Prettier (MIT) and the prettier-plugin-java bundle (Apache-2.0 and MIT) in `src/code-prettier/editor/vendor/` include their own license files.
