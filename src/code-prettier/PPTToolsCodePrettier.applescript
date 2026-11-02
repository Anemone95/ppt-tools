-- AppleScriptTask entry point for the PPT Tools Code Prettier editor.
-- Installed next to the PPTToolsCodePrettier editor folder in
-- ~/Library/Application Scripts/com.microsoft.Powerpoint/.

on openEditor(stateText)
	set scriptsFolder to POSIX path of ((path to me as text) & "::")
	set editorFolder to scriptsFolder & "PPTToolsCodePrettier"
	return do shell script "/usr/bin/osascript -l JavaScript " & ¬
		quoted form of (editorFolder & "/editor-window.js") & " " & ¬
		quoted form of editorFolder & " " & ¬
		quoted form of stateText without altering line endings
end openEditor
