$ErrorActionPreference = 'Stop'

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Install = Join-Path $Root 'INSTALL'
$LogFile = Join-Path $Root 'BUILD_LOG.txt'

function Write-Step([string]$Text) {
    Write-Host "`n============================================================"
    Write-Host $Text
    Write-Host "============================================================"
}

function Run-DotNet([string[]]$Arguments) {
    $printable = 'dotnet ' + ($Arguments -join ' ')
    Write-Host $printable
    Add-Content -Path $LogFile -Value ("`r`n> " + $printable)

    & $script:DotNet @Arguments 2>&1 | Tee-Object -FilePath $LogFile -Append
    if ($LASTEXITCODE -ne 0) {
        throw "dotnet exited with code $LASTEXITCODE."
    }
}

function Clean-Project([string]$ProjectDirectory) {
    foreach ($name in @('bin','obj')) {
        $path = Join-Path $ProjectDirectory $name
        if (Test-Path $path) {
            Remove-Item -LiteralPath $path -Recurse -Force
        }
    }
}

function Find-BuiltXll([string]$ProjectDirectory, [string]$Language) {
    $all = @(Get-ChildItem -LiteralPath $ProjectDirectory -Recurse -File -Filter '*.xll' -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -notmatch '[\\/]INSTALL[\\/]' })

    if ($all.Count -eq 0) {
        throw "The $Language build completed, but no .xll file exists under $ProjectDirectory"
    }

    # Prefer packed XLLs. If the current Excel-DNA version does not make packed
    # files, fall back to the regular XLL output.
    $packed = @($all | Where-Object { $_.Name -match '(?i)packed' })
    if ($packed.Count -gt 0) { return $packed }
    return $all
}

try {
    Set-Content -Path $LogFile -Value "FocusCell2021 v1.0 build log`r`nStarted: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Encoding UTF8

    Write-Step 'FocusCell2021 v1.0 - unified KR + EN build'
    Write-Host "Root    : $Root"

    $projects = @(Get-ChildItem -LiteralPath $Root -Recurse -File -Filter 'FocusCell2021.csproj' -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -notmatch '[\\/](bin|obj|INSTALL)[\\/]' })

    if ($projects.Count -eq 0) {
        throw "FocusCell2021.csproj was not found anywhere under $Root"
    }
    if ($projects.Count -gt 1) {
        Write-Host 'Multiple projects were found:'
        $projects | ForEach-Object { Write-Host ('  ' + $_.FullName) }
        throw 'More than one FocusCell2021.csproj exists. Keep only one project copy under this build folder.'
    }

    $Project = $projects[0].FullName
    $ProjectDir = $projects[0].Directory.FullName
    Write-Host "Project : $Project"

    $dotnetCommand = Get-Command dotnet.exe -ErrorAction SilentlyContinue
    if (-not $dotnetCommand) {
        $candidates = @(
            (Join-Path $env:ProgramFiles 'dotnet\dotnet.exe'),
            (if (${env:ProgramFiles(x86)}) { Join-Path ${env:ProgramFiles(x86)} 'dotnet\dotnet.exe' })
        ) | Where-Object { $_ -and (Test-Path $_) }
        if ($candidates.Count -eq 0) {
            throw 'dotnet.exe was not found. Install the .NET SDK or Visual Studio .NET desktop development workload.'
        }
        $script:DotNet = $candidates[0]
    } else {
        $script:DotNet = $dotnetCommand.Source
    }
    Write-Host "dotnet  : $script:DotNet"

    # Remove any previous install output. Language folders are created only
    # after that language has built an actual XLL successfully.
    if (Test-Path $Install) { Remove-Item -LiteralPath $Install -Recurse -Force }

    foreach ($lang in @('KR','EN')) {
        Write-Step "Building $lang"
        Clean-Project $ProjectDir

        Run-DotNet @('restore', $Project, "/p:FocusLanguage=$lang")
        Run-DotNet @('build', $Project, '-c', 'Release', '--no-restore', '--no-incremental', "/p:FocusLanguage=$lang")

        $xlls = @(Find-BuiltXll $ProjectDir $lang)
        $dest = Join-Path $Install $lang
        New-Item -ItemType Directory -Path $dest -Force | Out-Null

        foreach ($xll in $xlls) {
            $target = Join-Path $dest $xll.Name
            Copy-Item -LiteralPath $xll.FullName -Destination $target -Force
            Write-Host "Copied  : $target"
            Add-Content -Path $LogFile -Value ("Copied: " + $target)
        }

        if (@(Get-ChildItem -LiteralPath $dest -File -Filter '*.xll').Count -eq 0) {
            throw "No XLL reached INSTALL\$lang after the copy step."
        }
    }

    @(
        'FocusCell2021 v1.0 unified build',
        'KR edition: INSTALL\KR',
        'EN edition: INSTALL\EN',
        'Use the file containing 64 for 64-bit Excel.',
        'Use the file without 64 for 32-bit Excel.'
    ) | Set-Content -Path (Join-Path $Install 'README.txt') -Encoding UTF8

    Write-Step 'BUILD COMPLETE'
    Write-Host "KR : $(Join-Path $Install 'KR')"
    Write-Host "EN : $(Join-Path $Install 'EN')"
    Write-Host "Log: $LogFile"
    Add-Content -Path $LogFile -Value ("`r`nSUCCESS: " + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    exit 0
}
catch {
    $message = $_.Exception.Message
    Write-Host "`nBUILD ERROR: $message" -ForegroundColor Red
    Add-Content -Path $LogFile -Value ("`r`nBUILD ERROR: " + $message)
    Add-Content -Path $LogFile -Value $_.ScriptStackTrace
    Write-Host "Log saved to: $LogFile"
    exit 1
}
