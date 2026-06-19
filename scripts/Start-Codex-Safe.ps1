[CmdletBinding()]
param(
    [switch]$CheckOnly,
    [switch]$NoLaunch,
    [int]$InitialDelaySeconds = 0,
    [int]$LaunchWaitSeconds = 25,
    [int]$RepairRetryCount = 2
)

$ErrorActionPreference = "Stop"

$PackagePreference = @("OpenAI.CodexBeta", "OpenAI.Codex")
$CodexLocalDir = Join-Path $env:LOCALAPPDATA "OpenAI\Codex"
$RuntimeRoot = Join-Path $CodexLocalDir "runtimes\cua_node"
$LogPath = Join-Path $CodexLocalDir "codex-startup-guard.log"
$MutexName = "Local\OpenAI-Codex-Startup-Guard"

New-Item -ItemType Directory -Force -Path $CodexLocalDir | Out-Null

function Write-GuardLog {
    param([string]$Message)

    try {
        if ((Test-Path -LiteralPath $LogPath) -and ((Get-Item -LiteralPath $LogPath).Length -gt 1048576)) {
            Move-Item -LiteralPath $LogPath -Destination "$LogPath.1" -Force
        }
    }
    catch {
        # Logging must never prevent the guard from running.
    }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -LiteralPath $LogPath -Value "[$timestamp] $Message"
}

function Test-PathStartsWith {
    param(
        [string]$Path,
        [string]$Prefix
    )

    if ([string]::IsNullOrWhiteSpace($Path) -or [string]::IsNullOrWhiteSpace($Prefix)) {
        return $false
    }

    $normalizedPrefix = $Prefix.TrimEnd("\")
    return $Path.StartsWith($normalizedPrefix, [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-PackageInstallFromRegistry {
    param([string]$PackageName)

    $base = "HKCU:\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\Repository\Packages"
    if (-not (Test-Path -LiteralPath $base)) {
        return $null
    }

    $entries = @(Get-ChildItem -Path $base -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -like "$PackageName`_*" } |
        ForEach-Object { Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction SilentlyContinue } |
        Where-Object { $_.PackageRootFolder -and (Test-Path -LiteralPath $_.PackageRootFolder) } |
        Sort-Object PackageID -Descending)

    if ($entries.Count -eq 0) {
        return $null
    }

    return $entries[0]
}

function Get-InstalledTargets {
    $targets = foreach ($packageName in $PackagePreference) {
        $package = Get-AppxPackage -Name $packageName -ErrorAction SilentlyContinue
        if ($package) {
            [pscustomobject]@{
                Name = $package.Name
                PackageFamilyName = $package.PackageFamilyName
                InstallLocation = $package.InstallLocation
                Aumid = "$($package.PackageFamilyName)!App"
            }
            continue
        }

        $registryPackage = Get-PackageInstallFromRegistry -PackageName $packageName
        if ($registryPackage) {
            $familyName = ($registryPackage.PackageID -replace "_\d+\.\d+\.\d+\.\d+_.*$", "_2p2nqsd0c76g0")
            [pscustomobject]@{
                Name = $packageName
                PackageFamilyName = $familyName
                InstallLocation = $registryPackage.PackageRootFolder
                Aumid = "$familyName!App"
            }
        }
    }

    @($targets | Where-Object { $_.InstallLocation -and (Test-Path -LiteralPath $_.InstallLocation) })
}

function Get-ProcessPathSafe {
    param([System.Diagnostics.Process]$Process)

    try {
        return $Process.Path
    }
    catch {
        return $null
    }
}

function Get-TargetRootProcesses {
    param([object[]]$Targets)

    $processes = foreach ($process in Get-Process -ErrorAction SilentlyContinue) {
        $path = Get-ProcessPathSafe -Process $process
        foreach ($target in $Targets) {
            if (Test-PathStartsWith -Path $path -Prefix $target.InstallLocation) {
                $process
                break
            }
        }
    }

    @($processes)
}

function Get-VisibleTargetProcess {
    param([object[]]$Targets)

    Get-TargetRootProcesses -Targets $Targets |
        Where-Object { $_.MainWindowHandle -ne 0 -and $_.Responding } |
        Select-Object -First 1
}

function Get-RunningTargetProcess {
    param([object[]]$Targets)

    Get-TargetRootProcesses -Targets $Targets |
        Where-Object { $_.HasExited -eq $false } |
        Select-Object -First 1
}

function Stop-ProcessTree {
    param([int[]]$RootIds)

    if (-not $RootIds -or $RootIds.Count -eq 0) {
        return
    }

    $processRows = @(Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
        Select-Object ProcessId, ParentProcessId, Name, ExecutablePath)

    $ids = [System.Collections.Generic.HashSet[int]]::new()
    foreach ($rootId in $RootIds) {
        [void]$ids.Add([int]$rootId)
    }

    $changed = $true
    while ($changed) {
        $changed = $false
        foreach ($row in $processRows) {
            if ($ids.Contains([int]$row.ParentProcessId) -and -not $ids.Contains([int]$row.ProcessId)) {
                [void]$ids.Add([int]$row.ProcessId)
                $changed = $true
            }
        }
    }

    foreach ($id in ($ids | Sort-Object -Descending)) {
        try {
            Stop-Process -Id $id -Force -ErrorAction Stop
            Write-GuardLog "Stopped stale process id=$id"
        }
        catch {
            Write-GuardLog "Could not stop process id=${id}: $($_.Exception.Message)"
        }
    }
}

function Get-CuaRuntimeSource {
    param([object]$Target)

    $source = Join-Path $Target.InstallLocation "app\resources\cua_node"
    $node = Join-Path $source "bin\node.exe"
    $manifest = Join-Path $source "manifest.json"

    if ((Test-Path -LiteralPath $node) -and (Test-Path -LiteralPath $manifest)) {
        return [pscustomobject]@{
            Target = $Target
            Source = $source
            Node = $node
            Manifest = $manifest
            ManifestText = Get-Content -LiteralPath $manifest -Raw
        }
    }

    return $null
}

function Test-CuaRuntimeReady {
    param(
        [string]$RuntimePath,
        [string]$ManifestText
    )

    $node = Join-Path $RuntimePath "bin\node.exe"
    $manifest = Join-Path $RuntimePath "manifest.json"
    if (-not ((Test-Path -LiteralPath $node) -and (Test-Path -LiteralPath $manifest))) {
        return $false
    }

    try {
        $existingManifestText = Get-Content -LiteralPath $manifest -Raw
        if ($existingManifestText.Trim() -ne $ManifestText.Trim()) {
            return $false
        }

        $version = (& $node --version 2>$null).Trim()
        return ($version -match "^v\d+")
    }
    catch {
        Write-GuardLog "Runtime validation failed for ${RuntimePath}: $($_.Exception.Message)"
        return $false
    }
}

function Get-StagingRuntimeIds {
    if (-not (Test-Path -LiteralPath $RuntimeRoot)) {
        return @()
    }

    $seen = [System.Collections.Generic.HashSet[string]]::new()
    $ids = foreach ($dir in (Get-ChildItem -LiteralPath $RuntimeRoot -Force -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match "^\.staging-(?<id>[0-9a-fA-F]{16,64})-" } |
        Sort-Object LastWriteTime -Descending)) {
        $id = $Matches["id"].ToLowerInvariant()
        if ($seen.Add($id)) {
            $id
        }
    }

    @($ids)
}

function Get-ReadyRuntimeIds {
    param([string]$ManifestText)

    if (-not (Test-Path -LiteralPath $RuntimeRoot)) {
        return @()
    }

    $ids = foreach ($dir in (Get-ChildItem -LiteralPath $RuntimeRoot -Force -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notlike ".staging-*" -and $_.Name -match "^[0-9a-fA-F]{16,64}$" })) {
        if (Test-CuaRuntimeReady -RuntimePath $dir.FullName -ManifestText $ManifestText) {
            $dir.Name.ToLowerInvariant()
        }
    }

    @($ids)
}

function Copy-CuaRuntime {
    param(
        [string]$Source,
        [string]$Destination,
        [string]$ManifestText
    )

    if (Test-CuaRuntimeReady -RuntimePath $Destination -ManifestText $ManifestText) {
        Write-GuardLog "Runtime already ready: $Destination"
        return $false
    }

    New-Item -ItemType Directory -Force -Path $RuntimeRoot | Out-Null

    if (Test-Path -LiteralPath $Destination) {
        $backup = "$Destination.broken-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
        Move-Item -LiteralPath $Destination -Destination $backup -Force
        Write-GuardLog "Moved incomplete runtime to $backup"
    }

    $temp = "$Destination.repair-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    if (Test-Path -LiteralPath $temp) {
        Remove-Item -LiteralPath $temp -Recurse -Force
    }

    Write-GuardLog "Copying runtime from $Source to $Destination"
    Copy-Item -LiteralPath $Source -Destination $temp -Recurse -Force

    if (-not (Test-CuaRuntimeReady -RuntimePath $temp -ManifestText $ManifestText)) {
        throw "Copied runtime did not validate: $temp"
    }

    Move-Item -LiteralPath $temp -Destination $Destination -Force
    Write-GuardLog "Runtime repair completed: $Destination"
    return $true
}

function Ensure-CuaRuntimes {
    param([object[]]$Targets)

    $changed = $false
    $sources = @($Targets | ForEach-Object { Get-CuaRuntimeSource -Target $_ } | Where-Object { $_ -ne $null })

    foreach ($source in $sources) {
        $readyIds = @(Get-ReadyRuntimeIds -ManifestText $source.ManifestText)
        if ($readyIds.Count -gt 0) {
            Write-GuardLog "Cua runtime already ready for $($source.Target.Name): $($readyIds -join ',')"
            continue
        }

        $candidateIds = @(Get-StagingRuntimeIds | Select-Object -First 3)
        if ($candidateIds.Count -eq 0) {
            Write-GuardLog "No runtime id candidates yet for $($source.Target.Name); will retry after launch if needed."
            continue
        }

        foreach ($id in $candidateIds) {
            $destination = Join-Path $RuntimeRoot $id
            if (Copy-CuaRuntime -Source $source.Source -Destination $destination -ManifestText $source.ManifestText) {
                $changed = $true
            }
        }
    }

    return $changed
}

function Start-CodexTarget {
    param([object]$Target)

    Write-GuardLog "Launching $($Target.Name) via $($Target.Aumid)"
    Start-Process -FilePath "explorer.exe" -ArgumentList "shell:AppsFolder\$($Target.Aumid)" | Out-Null
}

function Test-CodexHealthy {
    param([object[]]$Targets)

    $visible = Get-VisibleTargetProcess -Targets $Targets
    if ($visible) {
        Write-GuardLog "Codex has a visible window. processId=$($visible.Id)"
        return $true
    }

    $running = Get-RunningTargetProcess -Targets $Targets
    if ($running) {
        Write-GuardLog "Codex package process is running. processId=$($running.Id)"
        return $true
    }

    return $false
}

$mutex = [System.Threading.Mutex]::new($false, $MutexName)
$hasMutex = $false

try {
    $hasMutex = $mutex.WaitOne(0)
    if (-not $hasMutex) {
        Write-GuardLog "Another startup guard instance is already running."
        exit 0
    }

    if ($InitialDelaySeconds -gt 0) {
        Start-Sleep -Seconds $InitialDelaySeconds
    }

    $targets = Get-InstalledTargets
    if ($targets.Count -eq 0) {
        Write-GuardLog "No OpenAI Codex package is installed."
        exit 10
    }

    Write-GuardLog "Targets: $(@($targets | ForEach-Object { $_.Name }) -join ', ')"

    [void](Ensure-CuaRuntimes -Targets $targets)

    if ($CheckOnly) {
        if (Test-CodexHealthy -Targets $targets) {
            exit 0
        }

        Write-GuardLog "CheckOnly: Codex is not currently running."
        exit 1
    }

    if ($NoLaunch) {
        Write-GuardLog "NoLaunch requested; repair phase complete."
        exit 0
    }

    if (Test-CodexHealthy -Targets $targets) {
        exit 0
    }

    foreach ($target in $targets) {
        $attempt = 0
        while ($attempt -le $RepairRetryCount) {
            $attempt += 1
            Start-CodexTarget -Target $target
            Start-Sleep -Seconds $LaunchWaitSeconds

            if (Test-CodexHealthy -Targets $targets) {
                Write-GuardLog "Codex launch confirmed. package=$($target.Name) attempt=$attempt"
                exit 0
            }

            Write-GuardLog "Launch attempt did not leave a running Codex process. package=$($target.Name) attempt=$attempt"
            $repaired = Ensure-CuaRuntimes -Targets $targets
            if (-not $repaired) {
                break
            }
        }
    }

    Write-GuardLog "Codex could not be restored automatically."
    exit 2
}
catch {
    Write-GuardLog "Startup guard failed: $($_.Exception.Message)"
    throw
}
finally {
    if ($hasMutex) {
        [void]$mutex.ReleaseMutex()
    }
    $mutex.Dispose()
}
