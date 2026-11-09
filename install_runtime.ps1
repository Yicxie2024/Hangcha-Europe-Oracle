# HCE Oracle - per-user one-click environment setup.
# No global PATH changes, no admin rights requested, no Excel/PDF modifications.
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$AppDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvDir = Join-Path $AppDir '.venv'
$VenvPy = Join-Path $VenvDir 'Scripts\python.exe'
$GUI = Join-Path $AppDir 'gui.py'
$Requirements = Join-Path $AppDir 'requirements.txt'
$Log = Join-Path $AppDir 'install_setup.log'
$PythonVersion = '3.14.10'
$SourceBase = 'https://www.python.org/ftp/python/' + $PythonVersion + '/'

function Say([string]$Message) {
    Write-Host ('[' + (Get-Date -Format 'HH:mm:ss') + '] ' + $Message)
}

function Get-ValidPython([string]$Executable, [string[]]$Prefix = @()) {
    try {
        $code = 'import sys, tkinter; assert sys.version_info >= (3, 11) and sys.maxsize > 2**32; print(sys.executable)'
        $output = & $Executable @Prefix -c $code 2>$null
        if ($LASTEXITCODE -ne 0) { return $null }
        $path = [string](@($output) | Select-Object -Last 1)
        $path = $path.Trim()
        if ($path -and (Test-Path -LiteralPath $path -PathType Leaf)) { return $path }
    } catch { }
    return $null
}

function Need([string]$Path, [string]$Description) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw ('Missing ' + $Description + ': ' + $Path)
    }
}

function Install-OfficialPython {
    Say 'No compatible Python was found. Downloading the official Python installer.'
    $arch = [System.Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString()
    switch ($arch) {
        'X64' { $filename = 'python-' + $PythonVersion + '-amd64.exe' }
        'Arm64' { $filename = 'python-' + $PythonVersion + '-arm64.exe' }
        default { throw ('Unsupported Windows architecture: ' + $arch + '. 64-bit Windows is required.') }
    }
    $tempDir = Join-Path $env:TEMP 'HangchaSalesAssistant_Setup'
    New-Item -ItemType Directory -Force -Path $tempDir | Out-Null
    $installer = Join-Path $tempDir $filename
    $url = $SourceBase + $filename
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    # Re-download even if a file exists: never trust a stale installer cache.
    Invoke-WebRequest -Uri $url -OutFile $installer -UseBasicParsing
    $signature = Get-AuthenticodeSignature -LiteralPath $installer
    if ($signature.Status -ne 'Valid' -or $null -eq $signature.SignerCertificate -or
        $signature.SignerCertificate.Subject -notmatch 'Python Software Foundation') {
        throw 'Python download failed the official digital-signature check. Installation stopped.'
    }
    Say ('Verified official Python installer: ' + $PythonVersion)
    # Separate current-user location; never override installed global Python/PATH.
    $targetDir = Join-Path $env:LOCALAPPDATA 'HangchaSalesAssistant\Python314'
    $args = @('/quiet', 'InstallAllUsers=0', ('TargetDir="' + $targetDir + '"'),
              'Include_launcher=0', 'Include_pip=1', 'Include_tcltk=1',
              'Include_test=0', 'Include_doc=0', 'PrependPath=0',
              'AppendPath=0', 'AssociateFiles=0', 'Shortcuts=0')
    Say 'Installing Python for the current Windows user (no admin rights required)...'
    $proc = Start-Process -FilePath $installer -ArgumentList $args -PassThru -Wait
    if ($proc.ExitCode -notin @(0, 3010)) {
        throw ('Python installer failed (exit code ' + $proc.ExitCode + '). Check company software policy or network access.')
    }
    $python = Join-Path $targetDir 'python.exe'
    if (-not (Get-ValidPython $python)) {
        throw 'Python setup completed but the expected interpreter cannot be started.'
    }
    return $python
}

try {
    Need $GUI 'GUI program'
    Need $Requirements 'requirements.txt'
    Say 'Checking the local Python environment...'

    $venvGood = $false
    if (Test-Path -LiteralPath $VenvPy -PathType Leaf) {
        $venvGood = [bool](Get-ValidPython $VenvPy)
        if ($venvGood) { Say 'Existing .venv is usable; it will be reused.' }
    }
    if (-not $venvGood) {
        if (Test-Path -LiteralPath $VenvDir) {
            $backupName = '.venv_backup_' + (Get-Date -Format 'yyyyMMdd_HHmmss')
            $backupPath = Join-Path $AppDir $backupName
            Say ('Previous .venv was incomplete. Preserving it as ' + $backupName)
            Move-Item -LiteralPath $VenvDir -Destination $backupPath
        }
        $basePython = $null
        if (Get-Command py -ErrorAction SilentlyContinue) {
            foreach ($ver in @('-3.14', '-3.15', '-3.13', '-3.12', '-3.11')) {
                $basePython = Get-ValidPython 'py' @($ver)
                if ($basePython) { break }
            }
        }
        if (-not $basePython) {
            $pythonCommand = Get-Command python -ErrorAction SilentlyContinue
            if ($pythonCommand) { $basePython = Get-ValidPython $pythonCommand.Source }
        }
        if (-not $basePython) {
            $preinstalled = Join-Path $env:LOCALAPPDATA 'HangchaSalesAssistant\Python314\python.exe'
            if (Test-Path -LiteralPath $preinstalled -PathType Leaf) {
                $basePython = Get-ValidPython $preinstalled
            }
        }
        if (-not $basePython) { $basePython = Install-OfficialPython }
        Say ('Using Python: ' + $basePython)
        Say 'Creating the app-specific virtual environment...'
        & $basePython -m venv $VenvDir
        if ($LASTEXITCODE -ne 0) { throw 'Could not create .venv in the application folder.' }
    }

    Need $VenvPy 'app virtual-environment python'
    Say 'Installing required packages. Internet access may be needed on the first run...'
    & $VenvPy -m pip install --disable-pip-version-check --only-binary=:all: -r $Requirements
    if ($LASTEXITCODE -ne 0) { throw 'Package installation failed. Check the pip error messages above.' }

    Say 'Checking GUI and PDF/Excel dependencies...'
    & $VenvPy -c 'import tkinter; import pdfplumber; import win32com.client; import reportlab'
    if ($LASTEXITCODE -ne 0) { throw 'A required module could not be imported.' }
    Say 'All required Python modules are ready.'

    # Desktop shortcut (per-user) is convenient for non-technical employees.
    try {
        $desktop = [Environment]::GetFolderPath('Desktop')
        if ($desktop) {
            $shell = New-Object -ComObject WScript.Shell
            $link = $shell.CreateShortcut((Join-Path $desktop 'HCE Oracle.lnk'))
            $pythonw = Join-Path $VenvDir 'Scripts\pythonw.exe'
            $link.TargetPath = $pythonw
            $link.Arguments = '"' + $GUI + '"'
            $link.WorkingDirectory = $AppDir
            $link.Description = 'HCE Oracle - business PDF and stock assistant'
            $link.Save()
            # Remove the legacy shortcut only when it points to THIS exact installation.
            $oldLinkPath = Join-Path $desktop 'Hangcha Sales Assistant.lnk'
            if (Test-Path -LiteralPath $oldLinkPath -PathType Leaf) {
                try {
                    $legacyLink = $shell.CreateShortcut($oldLinkPath)
                    if ($legacyLink.WorkingDirectory.TrimEnd('\') -ieq $AppDir.TrimEnd('\') -and
                        $legacyLink.Arguments -like '*gui.py*') {
                        Remove-Item -LiteralPath $oldLinkPath -Force
                        Say 'Renamed legacy desktop shortcut for this installation.'
                    }
                } catch { Say 'Old shortcut was left untouched.' }
            }
            Say 'Desktop shortcut created: HCE Oracle'
        }
    } catch {
        Say 'Desktop shortcut could not be created; 10_START_GUI.bat can still be used.'
    }

    $pythonw = Join-Path $VenvDir 'Scripts\pythonw.exe'
    Need $pythonw 'GUI runtime (pythonw.exe)'
    Say 'Setup complete. Opening HCE Oracle...'
    Start-Process -FilePath $pythonw -ArgumentList ('"' + $GUI + '"') -WorkingDirectory $AppDir | Out-Null
    exit 0
} catch {
    Say ('ERROR: ' + $_.Exception.Message)
    try {
        ('[' + (Get-Date -Format s) + '] ' + ($_ | Out-String)) | Out-File -LiteralPath $Log -Append -Encoding UTF8
        Say ('Error details saved: ' + $Log)
    } catch { }
    exit 1
}
