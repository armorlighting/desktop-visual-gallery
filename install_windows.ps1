$ErrorActionPreference = 'Stop'
$project = Split-Path -LiteralPath $MyInvocation.MyCommand.Path -Parent
$venvPythonw = Join-Path $project '.venv\Scripts\pythonw.exe'
$script = Join-Path $project 'visual_gallery.py'
$desktop = [Environment]::GetFolderPath('Desktop')
$shortcutPath = Join-Path $desktop 'Desktop Visual Gallery.lnk'

if (-not (Test-Path -LiteralPath $venvPythonw)) {
    throw 'Virtual environment not found. Create .venv and install requirements first.'
}

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $venvPythonw
$shortcut.Arguments = '"' + $script + '"'
$shortcut.WorkingDirectory = $project
$shortcut.Description = 'Refresh and open Desktop Visual Gallery'
$shortcut.IconLocation = 'C:\Windows\System32\shell32.dll,71'
$shortcut.WindowStyle = 7
$shortcut.Save()

Write-Output "Created: $shortcutPath"

