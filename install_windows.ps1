$ErrorActionPreference = 'Stop'
$project = Split-Path -LiteralPath $MyInvocation.MyCommand.Path -Parent
$venvPythonw = Join-Path $project '.venv\Scripts\pythonw.exe'
$script = Join-Path $project 'visual_gallery.py'
$icon = Join-Path $project 'assets\desktop-visual-gallery.ico'
$desktop = [Environment]::GetFolderPath('Desktop')
$shortcutPath = Join-Path $desktop 'Desktop Visual Gallery.lnk'

if (-not (Test-Path -LiteralPath $venvPythonw)) {
    throw 'Virtual environment not found. Create .venv and install requirements first.'
}

if (-not (Test-Path -LiteralPath $icon)) {
    throw 'Shortcut icon not found.'
}

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $venvPythonw
$shortcut.Arguments = '"' + $script + '"'
$shortcut.WorkingDirectory = $project
$shortcut.Description = 'Refresh and open Desktop Visual Gallery'
$shortcut.IconLocation = $icon + ',0'
$shortcut.WindowStyle = 7
$shortcut.Save()

Write-Output "Created: $shortcutPath"
