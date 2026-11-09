# HCE Oracle - replace only the shortcut of this exact installation.
$ErrorActionPreference = 'Stop'
$AppDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Desktop = [Environment]::GetFolderPath('Desktop')
$Gui = Join-Path $AppDir 'gui.py'
$PythonW = Join-Path $AppDir '.venv\Scripts\pythonw.exe'
$StartBat = Join-Path $AppDir '10_START_GUI.bat'
if (-not (Test-Path -LiteralPath $Gui -PathType Leaf)) { throw 'gui.py missing in application folder.' }
if (-not $Desktop) { throw 'Desktop location not available.' }
$Shell = New-Object -ComObject WScript.Shell
$NewShortcut = Join-Path $Desktop 'HCE Oracle.lnk'
$Link = $Shell.CreateShortcut($NewShortcut)
if (Test-Path -LiteralPath $PythonW -PathType Leaf) {
    $Link.TargetPath = $PythonW
    $Link.Arguments = '"' + $Gui + '"'
} else {
    if (-not (Test-Path -LiteralPath $StartBat -PathType Leaf)) { throw '10_START_GUI.bat missing.' }
    $Link.TargetPath = $StartBat
    $Link.Arguments = ''
}
$Link.WorkingDirectory = $AppDir
$Link.Description = 'HCE Oracle - Hangcha Europe business assistant'
$Link.Save()
$OldShortcut = Join-Path $Desktop 'Hangcha Sales Assistant.lnk'
if (Test-Path -LiteralPath $OldShortcut -PathType Leaf) {
    try {
        $OldLink = $Shell.CreateShortcut($OldShortcut)
        # Never delete a shortcut pointing to another copy of the software.
        if ($OldLink.WorkingDirectory.TrimEnd('\') -ieq $AppDir.TrimEnd('\') -and $OldLink.Arguments -like '*gui.py*') {
            Remove-Item -LiteralPath $OldShortcut -Force
            Write-Host 'Old shortcut of this installation was removed.'
        } else {
            Write-Host 'Legacy shortcut belongs to another installation and was preserved.'
        }
    } catch { Write-Host 'Legacy shortcut could not be checked and was preserved.' }
}
Write-Host ('HCE Oracle shortcut ready: ' + $NewShortcut)
