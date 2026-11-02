#!/bin/sh
# Install the Code Prettier editor where PowerPoint's AppleScriptTask can run it.
set -eu

root=$(cd "$(dirname "$0")/.." && pwd)
scripts_dir="$HOME/Library/Application Scripts/com.microsoft.Powerpoint"
editor_dir="$scripts_dir/PPTToolsCodePrettier"

mkdir -p "$scripts_dir"
rm -rf "$editor_dir"
cp -R "$root/src/code-prettier/editor" "$editor_dir"
osacompile -o "$scripts_dir/PPTToolsCodePrettier.scpt" \
  "$root/src/code-prettier/PPTToolsCodePrettier.applescript"

echo "Installed Code Prettier editor to: $scripts_dir"
