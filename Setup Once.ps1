$ErrorActionPreference = 'Continue'

$MinimumPythonVersion = [version]'3.11'
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvPath = Join-Path $ProjectRoot '.venv'
$VenvPython = Join-Path $VenvPath 'Scripts/python.exe'
$script:CandidatePaths = New-Object 'System.Collections.Generic.List[string]'

function Add-PythonCandidate {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return
    }

    $candidate = $Path.Trim().Trim([char]34)
    if (Test-Path -LiteralPath $candidate -PathType Container) {
        $candidate = Join-Path $candidate 'python.exe'
    }

    # Ignore the Windows Store execution alias. It is not an installed interpreter.
    if ($candidate.ToLowerInvariant().Contains('windowsapps') -and $candidate.ToLowerInvariant().EndsWith('python.exe')) {
        return
    }

    if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) {
        return
    }

    try {
        $candidate = [System.IO.Path]::GetFullPath($candidate)
    }
    catch {
        return
    }

    if (-not $script:CandidatePaths.Contains($candidate)) {
        [void]$script:CandidatePaths.Add($candidate)
    }
}

function Get-PythonCandidates {
    $script:CandidatePaths.Clear()

    if (Get-Command 'where.exe' -ErrorAction SilentlyContinue) {
        $pathResults = & where.exe python.exe 2>$null
        foreach ($pathResult in $pathResults) {
            Add-PythonCandidate ([string]$pathResult)
        }
    }

    $registryRoots = @(
        'HKCU:/Software/Python/PythonCore',
        'HKLM:/Software/Python/PythonCore',
        'HKLM:/Software/WOW6432Node/Python/PythonCore'
    )

    foreach ($registryRoot in $registryRoots) {
        if (-not (Test-Path -LiteralPath $registryRoot)) {
            continue
        }

        foreach ($pythonKey in (Get-ChildItem -LiteralPath $registryRoot -ErrorAction SilentlyContinue)) {
            $installKeyPath = Join-Path $pythonKey.PSPath 'InstallPath'
            if (-not (Test-Path -LiteralPath $installKeyPath)) {
                continue
            }

            $installKey = Get-Item -LiteralPath $installKeyPath
            Add-PythonCandidate ([string]$installKey.GetValue('ExecutablePath'))
            Add-PythonCandidate ([string]$installKey.GetValue(''))
        }
    }

    $installRoots = @()
    if ($env:LocalAppData) {
        $installRoots += (Join-Path $env:LocalAppData 'Programs/Python')
    }
    if ($env:ProgramFiles) {
        $installRoots += (Join-Path $env:ProgramFiles 'Python')
    }
    $programFilesX86 = [Environment]::GetEnvironmentVariable('ProgramFiles(x86)')
    if ($programFilesX86) {
        $installRoots += (Join-Path $programFilesX86 'Python')
    }

    foreach ($installRoot in $installRoots) {
        if (-not (Test-Path -LiteralPath $installRoot)) {
            continue
        }

        Add-PythonCandidate (Join-Path $installRoot 'python.exe')
        foreach ($installFolder in (Get-ChildItem -LiteralPath $installRoot -Directory -ErrorAction SilentlyContinue)) {
            if ($installFolder.Name -like 'Python*') {
                Add-PythonCandidate (Join-Path $installFolder.FullName 'python.exe')
            }
        }
    }

    $parentRoots = @()
    if ($env:ProgramFiles) {
        $parentRoots += $env:ProgramFiles
    }
    if ($programFilesX86) {
        $parentRoots += $programFilesX86
    }
    if ($env:LocalAppData) {
        $parentRoots += (Join-Path $env:LocalAppData 'Programs')
    }

    foreach ($parentRoot in $parentRoots) {
        if (-not (Test-Path -LiteralPath $parentRoot)) {
            continue
        }

        foreach ($installFolder in (Get-ChildItem -LiteralPath $parentRoot -Directory -ErrorAction SilentlyContinue)) {
            if ($installFolder.Name -like 'Python*') {
                Add-PythonCandidate (Join-Path $installFolder.FullName 'python.exe')
            }
        }
    }

    foreach ($installFolder in (Get-ChildItem -Path 'C:/Python*' -Directory -ErrorAction SilentlyContinue)) {
        Add-PythonCandidate (Join-Path $installFolder.FullName 'python.exe')
    }

    return $script:CandidatePaths.ToArray()
}

function Get-PythonVersion {
    param([string]$Executable)

    try {
        $output = & $Executable -c "import sys; print('.'.join(str(part) for part in sys.version_info[:3]))" 2>$null
        if ($LASTEXITCODE -ne 0 -or -not $output) {
            return $null
        }

        $versionText = ([string]($output | Select-Object -Last 1)).Trim()
        return [version]$versionText
    }
    catch {
        return $null
    }
}

function Find-UsablePython {
    $found = @()

    foreach ($candidate in (Get-PythonCandidates)) {
        $version = Get-PythonVersion $candidate
        if ($null -ne $version -and $version -ge $MinimumPythonVersion) {
            $found += [pscustomobject]@{
                Path = $candidate
                Version = $version
            }
        }
    }

    if ($found.Count -eq 0) {
        return $null
    }

    return ($found | Sort-Object -Property Version -Descending | Select-Object -First 1)
}

try {
    Set-Location -LiteralPath $ProjectRoot -ErrorAction Stop
    $requirementsPath = Join-Path $ProjectRoot 'requirements.txt'
    if (-not (Test-Path -LiteralPath $requirementsPath -PathType Leaf)) {
        throw 'requirements.txt is missing. Download the complete repository and run Setup Once.bat from its main folder.'
    }

    $pythonForVenv = $null
    if (Test-Path -LiteralPath $VenvPython -PathType Leaf) {
        $venvVersion = Get-PythonVersion $VenvPython
        if ($null -ne $venvVersion -and $venvVersion -ge $MinimumPythonVersion) {
            $pythonForVenv = [pscustomobject]@{
                Path = $VenvPython
                Version = $venvVersion
            }
            Write-Host "Using the existing project environment (Python $venvVersion)."
        }
    }

    if ($null -eq $pythonForVenv) {
        $pythonForVenv = Find-UsablePython

        if ($null -eq $pythonForVenv) {
            $oldVersions = @()
            foreach ($candidate in (Get-PythonCandidates)) {
                $foundVersion = Get-PythonVersion $candidate
                if ($null -ne $foundVersion) {
                    $oldVersions += [string]$foundVersion
                }
            }

            if ($oldVersions.Count -gt 0) {
                Write-Host "Found Python $($oldVersions -join ', '), but this app requires Python 3.11 or newer."
            }
            else {
                Write-Host 'No usable Python was found in PATH or the usual Windows install locations.'
            }

            $wingetCommand = Get-Command 'winget.exe' -ErrorAction SilentlyContinue
            if ($null -eq $wingetCommand) {
                throw 'Windows Package Manager (winget) is unavailable. Install Python 3.11 or newer from https://www.python.org/downloads/windows/, then rerun Setup Once.bat.'
            }

            Write-Host 'Installing Python 3.12 for this setup using winget. This can take a few minutes.'
            $savedErrorActionPreference = $ErrorActionPreference
            $ErrorActionPreference = 'Continue'
            & $wingetCommand.Source install --id Python.Python.3.12 --exact --silent --accept-source-agreements --accept-package-agreements --disable-interactivity
            $wingetExitCode = $LASTEXITCODE
            $ErrorActionPreference = $savedErrorActionPreference

            # Search again even if winget returned a non-zero code; Python may already
            # be installed but absent from PATH, or the installer may have completed.
            $pythonForVenv = Find-UsablePython
            if ($null -eq $pythonForVenv) {
                throw "Python installation did not produce a usable interpreter (winget exit code $wingetExitCode). Check internet access or company install restrictions, then install Python from https://www.python.org/downloads/windows/ and rerun Setup Once.bat."
            }
        }

        Write-Host "Using Python $($pythonForVenv.Version) from $($pythonForVenv.Path)."

        if (Test-Path -LiteralPath $VenvPath) {
            $backupPath = Join-Path $ProjectRoot ('.venv.backup-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
            Move-Item -LiteralPath $VenvPath -Destination $backupPath -ErrorAction Stop
            Write-Host "Moved the unusable old environment to $backupPath."
        }

        Write-Host 'Creating the project virtual environment...'
        & $pythonForVenv.Path -m venv $VenvPath
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $VenvPython -PathType Leaf)) {
            throw 'Python was found, but the project virtual environment could not be created.'
        }
    }

    Write-Host 'Restoring pip if needed, then installing the app packages...'
    & $VenvPython -m ensurepip --upgrade
    if ($LASTEXITCODE -ne 0) {
        throw 'pip could not be restored in the project environment. Rerun Setup Once.bat or reinstall Python 3.11 or newer.'
    }

    & $VenvPython -m pip install --disable-pip-version-check --upgrade pip
    if ($LASTEXITCODE -ne 0) {
        throw 'pip could not be updated. Check internet access or company proxy settings, then rerun Setup Once.bat.'
    }

    & $VenvPython -m pip install --disable-pip-version-check -r (Join-Path $ProjectRoot 'requirements.txt')
    if ($LASTEXITCODE -ne 0) {
        throw 'The app packages could not be installed. Check internet access or company proxy settings, then rerun Setup Once.bat.'
    }

    Write-Host ''
    Write-Host "Setup complete. Python $($pythonForVenv.Version) and the app packages are ready."
    Write-Host 'Double-click Prepare Drafts.bat to start the app.'
    exit 0
}
catch {
    Write-Host ''
    Write-Host 'Setup did not finish.' -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host 'After fixing the issue, rerun Setup Once.bat. If automatic Python installation is blocked, install Python 3.11 or newer from https://www.python.org/downloads/windows/.'
    exit 1
}
