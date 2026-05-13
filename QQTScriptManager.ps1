# Script Updater Framework (PowerShell)
# ------------------------------------
#
# This PowerShell script provides a simple way to manage a set of script
# packages hosted on GitHub.  It does not require any external modules; it
# relies on built-in cmdlets such as `Invoke-RestMethod` to query GitHub and
# `Invoke-WebRequest`/`Expand-Archive` to download and extract ZIP files.
#
# The script performs the following tasks:
#
#   * Use GitHub's REST API to obtain the latest commit SHA on a specified
#     branch.
#   * Download a ZIP archive of a repository at a particular commit or branch.
#   * Store metadata about the last installed commit SHA so that update
#     operations can determine whether a new version is available.
#   * Extract the downloaded ZIP, flattening the top-level directory so that
#     files land directly in the package's installation directory.
#
# To add your own packages, edit the `$packages` array below.  Each entry
# requires a `name`, `owner`, `repo` and optionally a `branch` (default is
# `main`).

param()

# Configuration: base directory for installed packages.  Change this if you
# want to install packages somewhere else.  By default it uses a directory
# called "scripts" in the current working directory.  The directory is
# created automatically when needed.
$BaseDir = Join-Path -Path (Get-Location) -ChildPath "scripts"

# Central metadata file used to store the last installed commit for all
# packages (both regular and collections).  Stored inside the scripts
# directory.  Tracks the SHA for each package and each individual script
# within collections.
$MetadataFile = Join-Path -Path $BaseDir -ChildPath 'scripts_metadata.json'

# Define the packages you want to manage.  To add a new package, copy one of
# the entries below and fill in `name`, `owner`, `repo`, and optionally
# `branch` (defaults to "main").  Mark a package with `isCollection=$true`
# if its archive contains multiple scripts in subdirectories.
$packages = @(
    @{ name="SilentRaven"; owner="magoogle"; repo="SilentRaven"; branch="main" }
    @{ name="UniversalRotation"; owner="magoogle"; repo="UniversalRotation"; branch="main" }
    @{ name="jimps-gem-farmer"; owner="EZBOOPS"; repo="jimps-gem-farmer"; branch="master" }
    @{ name="War-Pig-Zewx"; owner="Zewx1776"; repo="War-Pig-Zewx"; branch="main"; isCollection=$true }
    @{ name="GoFish"; owner="magoogle"; repo="GoFish"; branch="main" }
    @{ name="HordeDev"; owner="Zewx1776"; repo="HordeDev"; branch="main" }
    @{ name="LooteerV3"; owner="magoogle"; repo="LooteerV3"; branch="master" }
    @{ name="D4QQT"; owner="OldOnSteroid"; repo="D4QQT"; branch="main"; isCollection=$true }
    @{ name="Reaper"; owner="magoogle"; repo="Reaper"; branch="main" }
    @{ name="AlfredTheButler"; owner="Leoc76101111"; repo="AlfredTheButler"; branch="main" }
    @{ name="Scmurd-Warlock"; owner="Scmurd"; repo="Scmurd-Warlock"; branch="main" }
    @{ name="WonderCity"; owner="Leoc76101111"; repo="WonderCity"; branch="main" }
)

function Get-InstallDir {
    param([hashtable]$pkg)
    return Join-Path -Path $BaseDir -ChildPath $pkg.name
}

function Get-LatestSHA {
    param([hashtable]$pkg)
    $branch = if ($pkg.ContainsKey('branch')) { $pkg.branch } else { 'main' }
    $uri = "https://api.github.com/repos/$($pkg.owner)/$($pkg.repo)/commits/$branch"
    $headers = @{ 'User-Agent' = 'script-updater/1.0' }
    if ($env:GITHUB_TOKEN) {
        $headers['Authorization'] = "token $($env:GITHUB_TOKEN)"
    }
    try {
        $resp = Invoke-RestMethod -Uri $uri -Headers $headers -UseBasicParsing
        return $resp.sha
    } catch {
        $statusCode = $null
        if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
            $statusCode = $_.Exception.Response.StatusCode.value__
        }
        if ($statusCode -eq 403) {
            $gitExe = Get-Command git -ErrorAction SilentlyContinue
            if ($gitExe) {
                try {
                    $lsRemote = & git ls-remote "https://github.com/$($pkg.owner)/$($pkg.repo)" $branch 2>$null
                    if ($lsRemote) {
                        $commitSha = $lsRemote.Split("`n")[0].Split("`t")[0]
                        if ($commitSha) { return $commitSha }
                    }
                } catch {}
            }
        }
        throw "Failed to retrieve commit information from GitHub: $($_.Exception.Message)"
    }
}

function Download-Zip {
    param([hashtable]$pkg, [string]$ref)
    $primaryUri  = "https://github.com/$($pkg.owner)/$($pkg.repo)/archive/$ref.zip"
    $fallbackUri = "https://codeload.github.com/$($pkg.owner)/$($pkg.repo)/zip/$ref"
    $tempBase = [System.IO.Path]::GetTempFileName()
    $zipPath  = "$tempBase.zip"
    Remove-Item -Path $tempBase -Force -ErrorAction SilentlyContinue
    $headers = @{ 'User-Agent' = 'script-updater/1.0' }
    $uris = @($primaryUri, $fallbackUri)
    foreach ($downloadUri in $uris) {
        try {
            Invoke-WebRequest -Uri $downloadUri -OutFile $zipPath -Headers $headers -UseBasicParsing
            return $zipPath
        } catch {
            if (Test-Path $zipPath) { Remove-Item -Path $zipPath -Force }
            if ($downloadUri -eq $uris[-1]) {
                throw "Failed to download archive from GitHub: $($_.Exception.Message)"
            }
        }
    }
}

function Read-InstalledSHA {
    param([hashtable]$pkg)
    $manifestPath = $MetadataFile
    if (Test-Path $manifestPath) {
        try {
            $entries = Get-Content -Path $manifestPath -Raw | ConvertFrom-Json
            $entriesArray = @()
            if ($null -ne $entries) {
                if ($entries -is [System.Collections.IEnumerable]) {
                    foreach ($item in $entries) { $entriesArray += $item }
                } else {
                    $entriesArray += $entries
                }
            }
            if ($pkg.ContainsKey('subPackOf')) {
                $spEntry = $entriesArray | Where-Object { $_.package -eq $pkg.subPackOf -and $_.script -eq $pkg.name } | Select-Object -First 1
                if ($null -ne $spEntry) { return $spEntry.sha }
                return $null
            }
            $general = $entriesArray | Where-Object { $_.package -eq $pkg.name -and ($_.script -eq $null -or $_.script -eq '' -or $_.script -eq 'standalone') } | Select-Object -First 1
            if ($null -ne $general) { return $general.sha }
            $entry = $entriesArray | Where-Object { $_.package -eq $pkg.name } | Select-Object -First 1
            if ($null -ne $entry) { return $entry.sha }
        } catch {}
    }
    return $null
}

function Write-InstalledSHA {
    param([hashtable]$pkg, [string]$sha)
    Update-PackageMetadata -package $pkg.name -script 'standalone' -sha $sha
}

function Clear-InstallDir {
    param([hashtable]$pkg)
    $installDir = Get-InstallDir $pkg
    if (-not (Test-Path $installDir)) { return }
    Get-ChildItem -Path $installDir -Force | ForEach-Object {
        if ($_.PSIsContainer) {
            Remove-Item -Path $_.FullName -Recurse -Force
        } else {
            Remove-Item -Path $_.FullName -Force
        }
    }
}

function Extract-Zip {
    param([hashtable]$pkg, [string]$zipPath)
    $installDir = Get-InstallDir $pkg
    if (-not (Test-Path $installDir)) { New-Item -ItemType Directory -Path $installDir | Out-Null }
    $tempExtractDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ([System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tempExtractDir | Out-Null
    try {
        Expand-Archive -Path $zipPath -DestinationPath $tempExtractDir -Force
    } catch {
        Remove-Item -Path $tempExtractDir -Recurse -Force
        throw "Failed to extract archive: $($_.Exception.Message)"
    }
    $rootDir = Get-ChildItem -Path $tempExtractDir | Where-Object { $_.PSIsContainer } | Select-Object -First 1
    if ($null -eq $rootDir) { $rootDir = $tempExtractDir }
    Get-ChildItem -Path $rootDir.FullName -Recurse -Force | ForEach-Object {
        $relativePath = $_.FullName.Substring($rootDir.FullName.Length).TrimStart('\\')
        $dest = Join-Path -Path $installDir -ChildPath $relativePath
        if ($_.PSIsContainer) {
            if (-not (Test-Path $dest)) { New-Item -ItemType Directory -Path $dest | Out-Null }
        } else {
            $destDir = Split-Path -Path $dest -Parent
            if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir | Out-Null }
            Copy-Item -Path $_.FullName -Destination $dest -Force
        }
    }
    Remove-Item -Path $tempExtractDir -Recurse -Force
}

# Extract a downloaded ZIP archive and update only changed files in the install
# directory.  Files that don't exist in the archive remain untouched to
# preserve user settings.  New or changed files are copied from the archive.
function Extract-And-Update {
    param([hashtable]$pkg, [string]$zipPath)
    $installDir = Get-InstallDir $pkg
    if (-not (Test-Path $installDir)) { New-Item -ItemType Directory -Path $installDir | Out-Null }
    $tempExtractDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ([System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tempExtractDir | Out-Null
    try {
        Expand-Archive -Path $zipPath -DestinationPath $tempExtractDir -Force
    } catch {
        Remove-Item -Path $tempExtractDir -Recurse -Force
        throw "Failed to extract archive: $($_.Exception.Message)"
    }
    $rootDir = Get-ChildItem -Path $tempExtractDir | Where-Object { $_.PSIsContainer } | Select-Object -First 1
    if ($null -eq $rootDir) { $rootDir = $tempExtractDir }
    Get-ChildItem -Path $rootDir.FullName -Recurse -File | ForEach-Object {
        $relative = $_.FullName.Substring($rootDir.FullName.Length).TrimStart('\')
        $destPath = Join-Path -Path $installDir -ChildPath $relative
        $destDir  = Split-Path -Path $destPath -Parent
        if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir | Out-Null }
        if (Test-Path $destPath) {
            try {
                $srcHash = (Get-FileHash -Algorithm SHA256 -Path $_.FullName).Hash
                $dstHash = (Get-FileHash -Algorithm SHA256 -Path $destPath).Hash
            } catch {
                $srcHash = ''; $dstHash = 'different'
            }
            if ($srcHash -ne $dstHash) { Copy-Item -Path $_.FullName -Destination $destPath -Force }
        } else {
            Copy-Item -Path $_.FullName -Destination $destPath -Force
        }
    }
    Remove-Item -Path $tempExtractDir -Recurse -Force
}

# Copy a directory's contents to a destination directory, overwriting files
# only when their content has changed.  Files in the destination but not in
# the source are preserved.  Used to update individual script directories
# inside collection packages.
function Copy-Directory-With-Update {
    param([string]$SourceDir, [string]$DestDir, [string]$RootNameToFlatten)
    if (-not (Test-Path $SourceDir)) { throw "Source directory $SourceDir does not exist." }
    if (-not (Test-Path $DestDir))   { New-Item -ItemType Directory -Path $DestDir | Out-Null }
    Get-ChildItem -Path $SourceDir -Recurse -File | ForEach-Object {
        $relative      = $_.FullName.Substring($SourceDir.Length).TrimStart('\')
        $destFilePath  = Join-Path -Path $DestDir -ChildPath $relative
        $destParentDir = Split-Path -Path $destFilePath -Parent
        if (-not (Test-Path $destParentDir)) { New-Item -ItemType Directory -Path $destParentDir | Out-Null }
        if (Test-Path $destFilePath) {
            try {
                $srcHash = (Get-FileHash -Algorithm SHA256 -Path $_.FullName).Hash
                $dstHash = (Get-FileHash -Algorithm SHA256 -Path $destFilePath).Hash
            } catch {
                $srcHash = ''; $dstHash = 'different'
            }
            if ($srcHash -ne $dstHash) { Copy-Item -Path $_.FullName -Destination $destFilePath -Force }
        } else {
            Copy-Item -Path $_.FullName -Destination $destFilePath -Force
        }
    }
}

# Remove the install directory for a package.  Used for collection packages
# to avoid leaving empty folders behind after extraction.
function Remove-PackageDir {
    param([hashtable]$pkg)
    $dir = Get-InstallDir $pkg
    if (Test-Path $dir) {
        try {
            Remove-Item -Path $dir -Recurse -Force
        } catch {
            $msg = "Failed to remove directory {0}: {1}" -f $dir, $_.Exception.Message
            Write-Host $msg -ForegroundColor Red
        }
    }
}

# Update or add a single entry in the central metadata file.  Ensures at most
# one entry per package/script combination.  Replaces an existing entry if the
# SHA changed; no-ops if the SHA already matches.
function Update-PackageMetadata {
    param(
        [Parameter(Mandatory=$true)] [string]$package,
        [Parameter()] $script,
        [Parameter(Mandatory=$true)] [string]$sha
    )
    $pkgVal       = $package
    $scriptVal    = $script
    $shaVal       = $sha
    $isStandalone = ($null -eq $scriptVal -or $scriptVal -eq 'standalone')
    $manifestPath = $MetadataFile
    $allEntries   = @()
    if (Test-Path $manifestPath) {
        try {
            $raw = Get-Content -Path $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($null -ne $raw) {
                if ($raw -is [System.Collections.IEnumerable]) { foreach ($item in $raw) { $allEntries += $item } }
                else { $allEntries += $raw }
            }
        } catch { $allEntries = @() }
    }

    foreach ($entry in $allEntries) {
        if ($entry.package -ne $pkgVal) { continue }
        $scMatch = if ($isStandalone) {
            ($null -eq $entry.script -or $entry.script -eq '' -or $entry.script -eq 'standalone')
        } else { ($entry.script -eq $scriptVal) }
        if ($scMatch -and $entry.sha -eq $shaVal) { return }
        if ($scMatch) { break }
    }

    $kept = @()
    foreach ($entry in $allEntries) {
        if ($entry.package -ne $pkgVal) { $kept += $entry; continue }
        $scMatch = if ($isStandalone) {
            ($null -eq $entry.script -or $entry.script -eq '' -or $entry.script -eq 'standalone')
        } else { ($entry.script -eq $scriptVal) }
        if (-not $scMatch) { $kept += $entry }
    }

    $kept += @{ package = $pkgVal; sha = $shaVal; script = if ($null -ne $scriptVal) { $scriptVal } else { $null } }

    $manifestDir = Split-Path -Path $manifestPath -Parent
    if (-not (Test-Path $manifestDir)) { New-Item -ItemType Directory -Path $manifestDir | Out-Null }
    $parts   = @($kept | ForEach-Object { ConvertTo-Json -InputObject $_ -Depth 10 })
    $jsonOut = '[' + ($parts -join ',') + ']'
    Set-Content -Path $manifestPath -Value $jsonOut -Encoding UTF8
}

# Extract a downloaded archive for a collection package.  Each top-level
# directory in the archive is a separate script; it gets its own folder
# under $BaseDir (e.g. scripts/ArkhamAsylum/, scripts/Batmobile/).
# Version suffixes in directory names (e.g. -1.0.6, -main) are stripped.
function Extract-Collection {
    param([hashtable]$pkg, [string]$zipPath, [string]$latestSha)
    $tempExtractDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ([System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tempExtractDir | Out-Null
    try {
        Expand-Archive -Path $zipPath -DestinationPath $tempExtractDir -Force
    } catch {
        Remove-Item -Path $tempExtractDir -Recurse -Force
        throw "Failed to extract collection archive: $($_.Exception.Message)"
    }
    $rootDir = Get-ChildItem -Path $tempExtractDir | Where-Object { $_.PSIsContainer } | Select-Object -First 1
    if ($null -eq $rootDir) { $rootDir = [System.IO.DirectoryInfo]$tempExtractDir }

    $insignificant = @('.gitignore', 'readme.md', 'readme.txt', 'license', 'license.md',
                       'license.txt', '.gitattributes', 'changelog.md', 'changelog.txt')
    for ($i = 0; $i -lt 3; $i++) {
        $items    = @(Get-ChildItem -Path $rootDir.FullName)
        $subDirs  = @($items | Where-Object {  $_.PSIsContainer })
        $sigFiles = @($items | Where-Object { -not $_.PSIsContainer -and ($insignificant -notcontains $_.Name.ToLower()) })
        if ($subDirs.Count -eq 1 -and $sigFiles.Count -eq 0) {
            $innerDirs = @(Get-ChildItem -Path $subDirs[0].FullName | Where-Object { $_.PSIsContainer })
            if ($innerDirs.Count -gt 0) {
                Write-Host "  Detected umbrella folder '$($subDirs[0].Name)'; descending." -ForegroundColor DarkGray
                $rootDir = $subDirs[0]
            } else { break }
        } else { break }
    }

    $collName = $pkg.name

    if (Test-Path $MetadataFile) {
        try {
            $preMeta = @()
            $rawPre  = Get-Content -Path $MetadataFile -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($null -ne $rawPre) {
                if ($rawPre -is [System.Collections.IEnumerable]) { foreach ($e in $rawPre) { $preMeta += $e } }
                else { $preMeta += $rawPre }
            }
            $purged = @(); $removed = @()
            foreach ($entry in $preMeta) {
                $isThisColl  = ($entry.package -eq $collName)
                $isSubScript = $isThisColl -and ($null -ne $entry.script) -and ($entry.script -ne '') -and ($entry.script -ne 'standalone')
                if ($isSubScript) { $removed += $entry.script } else { $purged += $entry }
            }
            $parts = @($purged | ForEach-Object { ConvertTo-Json -InputObject $_ -Depth 5 })
            Set-Content -Path $MetadataFile -Value ('[' + ($parts -join ',') + ']') -Encoding UTF8
        } catch {}
    }

    $versionRegex   = [regex]'^(.+?)-(?:\d+(?:\.\d+)*|master|main)$'
    $extractedNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    Get-ChildItem -Path $rootDir.FullName | Where-Object { $_.PSIsContainer } | ForEach-Object {
        $origName = $_.Name
        $baseName = $origName
        $m = $versionRegex.Match($origName)
        if ($m.Success) { $baseName = $m.Groups[1].Value }
        $destPath = Join-Path -Path $BaseDir -ChildPath $baseName
        if (Test-Path $destPath) { Remove-Item -Path $destPath -Recurse -Force }
        Write-Host "  Installing $baseName..." -ForegroundColor DarkGray
        try {
            Copy-Directory-With-Update -SourceDir $_.FullName -DestDir $destPath
            Update-PackageMetadata -package $collName -script $baseName -sha $latestSha
            [void]$extractedNames.Add($baseName)
        } catch {
            Write-Host "  Error installing $baseName : $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    Update-PackageMetadata -package $collName -script 'standalone' -sha $latestSha

    if ($extractedNames.Count -gt 0 -and (Test-Path $MetadataFile)) {
        try {
            $allMeta = @()
            $rawMeta = Get-Content -Path $MetadataFile -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($null -ne $rawMeta) {
                if ($rawMeta -is [System.Collections.IEnumerable]) { foreach ($e in $rawMeta) { $allMeta += $e } }
                else { $allMeta += $rawMeta }
            }
            $cleaned = @(); $hasStale = $false
            foreach ($entry in $allMeta) {
                $isThisColl  = ($entry.package -eq $collName)
                $isSubScript = $isThisColl -and ($null -ne $entry.script) -and ($entry.script -ne '') -and ($entry.script -ne 'standalone')
                if ($isSubScript -and -not $extractedNames.Contains($entry.script)) {
                    $hasStale = $true
                } else {
                    $cleaned += $entry
                }
            }
            if ($hasStale) {
                Write-Host "  Removing stale metadata entries for $collName..." -ForegroundColor DarkGray
                $parts = @($cleaned | ForEach-Object { ConvertTo-Json -InputObject $_ -Depth 5 })
                Set-Content -Path $MetadataFile -Value ('[' + ($parts -join ',') + ']') -Encoding UTF8
            }
        } catch {}
    }

    Remove-Item -Path $tempExtractDir -Recurse -Force
}

function Update-Package {
    param([hashtable]$pkg)
    try {
        $latestSha = Get-LatestSHA $pkg
    } catch {
        Write-Host "Warning: Could not retrieve latest commit for $($pkg.name); using branch name instead." -ForegroundColor Yellow
        $latestSha = if ($pkg.ContainsKey('branch')) { $pkg.branch } else { 'main' }
    }
    $installedSha = Read-InstalledSHA $pkg
    if ($installedSha -and $latestSha -and ($installedSha -eq $latestSha)) {
        Write-Host "$($pkg.name) is already up-to-date (commit $latestSha)." -ForegroundColor Green
        return
    }
    Write-Host "Downloading $($pkg.name) at commit $latestSha..." -ForegroundColor Cyan
    try {
        $zipPath = Download-Zip $pkg $latestSha
    } catch {
        Write-Host "Error downloading archive: $($_.Exception.Message)" -ForegroundColor Red
        return
    }
    try {
        $isCollection = $pkg.ContainsKey('isCollection') -and $pkg.isCollection
        if ($isCollection) {
            Extract-Collection $pkg $zipPath $latestSha
        } else {
            if (-not $installedSha) {
                Clear-InstallDir $pkg
                Extract-Zip $pkg $zipPath
            } else {
                Extract-And-Update $pkg $zipPath
            }
            Write-InstalledSHA $pkg $latestSha
        }
        Write-Host "Installed/updated $($pkg.name) successfully." -ForegroundColor Green
    } catch {
        Write-Host "Error installing $($pkg.name): $($_.Exception.Message)" -ForegroundColor Red
    } finally {
        if (Test-Path $zipPath) { Remove-Item -Path $zipPath -Force }
    }
}

# Force a clean install of a package by clearing the install directory and
# extracting all files from the downloaded archive.  Used by "Download All"
# to install every package regardless of whether it was previously installed.
function Install-Package {
    param([hashtable]$pkg)
    try {
        $latestSha = Get-LatestSHA $pkg
    } catch {
        Write-Host "Warning: Could not retrieve latest commit for $($pkg.name); using branch name instead." -ForegroundColor Yellow
        $latestSha = if ($pkg.ContainsKey('branch')) { $pkg.branch } else { 'main' }
    }
    Write-Host "Downloading $($pkg.name) at commit $latestSha..." -ForegroundColor Cyan
    try {
        $zipPath = Download-Zip $pkg $latestSha
    } catch {
        Write-Host "Error downloading archive: $($_.Exception.Message)" -ForegroundColor Red
        return
    }
    try {
        $isCollection = $pkg.ContainsKey('isCollection') -and $pkg.isCollection
        if ($isCollection) {
            Extract-Collection $pkg $zipPath $latestSha
        } else {
            Clear-InstallDir $pkg
            Extract-Zip $pkg $zipPath
            Write-InstalledSHA $pkg $latestSha
        }
        Write-Host "Installed $($pkg.name) successfully." -ForegroundColor Green
    } catch {
        Write-Host "Error installing $($pkg.name): $($_.Exception.Message)" -ForegroundColor Red
    } finally {
        if (Test-Path $zipPath) { Remove-Item -Path $zipPath -Force }
    }
}

# Update all packages that are currently installed.  Skips packages that have
# never been installed and skips sub-packs (they are updated via their parent
# collection).
function Update-AllInstalled {
    foreach ($pkg in $packages) {
        if ($pkg.ContainsKey('subPackOf')) {
            Write-Host "$($pkg.name) is part of $($pkg.subPackOf); updated with collection." -ForegroundColor DarkGray
            continue
        }
        $isCollection = $pkg.ContainsKey('isCollection') -and $pkg.isCollection
        $manifestPath = $MetadataFile
        $installed    = $false
        if (Test-Path $manifestPath) {
            try {
                $entries = Get-Content -Path $manifestPath -Raw | ConvertFrom-Json
                $entriesArray = @()
                if ($null -ne $entries) {
                    if ($entries -is [System.Collections.IEnumerable]) {
                        foreach ($item in $entries) { $entriesArray += $item }
                    } else { $entriesArray += $entries }
                }
                if ($isCollection) {
                    $match = $entriesArray | Where-Object { $_.package -eq $pkg.name } | Select-Object -First 1
                    if ($match) { $installed = $true }
                } else {
                    $match = $entriesArray | Where-Object { $_.package -eq $pkg.name -and ($_.script -eq $null -or $_.script -eq '' -or $_.script -eq 'standalone') } | Select-Object -First 1
                    if ($match) { $installed = $true }
                }
            } catch {}
        }
        if ($installed) {
            Update-Package $pkg
        } else {
            Write-Host "$($pkg.name) is not installed; skipping update." -ForegroundColor DarkGray
        }
    }
}

# Download (clean install) all packages regardless of their current state.
# Installs each package fresh; ensures all packages are present and up-to-date.
function Download-AllPackages {
    foreach ($pkg in $packages) { Install-Package $pkg }
}

function Uninstall-Package {
    param([hashtable]$pkg)
    $isCollection = $pkg.ContainsKey('isCollection') -and $pkg.isCollection
    if ($isCollection) {
        if (Test-Path $MetadataFile) {
            try {
                $rawArr = Get-Content $MetadataFile -Raw -Encoding UTF8 | ConvertFrom-Json
                $arr = @()
                if ($null -ne $rawArr) {
                    if ($rawArr -is [System.Collections.IEnumerable]) { foreach ($e in $rawArr) { $arr += $e } }
                    else { $arr += $rawArr }
                }
                $scriptNames = @($arr | Where-Object {
                    $_.package -eq $pkg.name -and
                    $null -ne $_.script -and $_.script -ne '' -and $_.script -ne 'standalone'
                } | ForEach-Object { $_.script })
                foreach ($sName in $scriptNames) {
                    $dir = Join-Path $BaseDir $sName
                    if (Test-Path $dir) {
                        Remove-Item $dir -Recurse -Force -ErrorAction SilentlyContinue
                        Write-Host "Removed $sName" -ForegroundColor Yellow
                    }
                }
            } catch {}
        }
    } else {
        $dir = Get-InstallDir $pkg
        if (Test-Path $dir) {
            Remove-Item $dir -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "Removed $($pkg.name)" -ForegroundColor Yellow
        }
    }
    if (Test-Path $MetadataFile) {
        try {
            $rawDel = Get-Content $MetadataFile -Raw -Encoding UTF8 | ConvertFrom-Json
            $allDel = @()
            if ($null -ne $rawDel) {
                if ($rawDel -is [System.Collections.IEnumerable]) { foreach ($e in $rawDel) { $allDel += $e } }
                else { $allDel += $rawDel }
            }
            $filtered = @($allDel | Where-Object { $_.package -ne $pkg.name })
            $parts = @($filtered | ForEach-Object { ConvertTo-Json -InputObject $_ -Depth 10 })
            Set-Content -Path $MetadataFile -Value ('[' + ($parts -join ',') + ']') -Encoding UTF8
        } catch {}
    }
    if ($script:statusCache.ContainsKey($pkg.name)) {
        $script:statusCache.Remove($pkg.name)
        Save-StatusCache $script:statusCache
    }
    Write-Host "Uninstalled $($pkg.name)." -ForegroundColor Green
}

# Discover installed sub-packs for a collection from metadata and add them to
# $script:packages if not already present.  Call after any collection
# install/update so the ListView shows the individual scripts inside the pack.
function Sync-CollectionSubPacks {
    param([hashtable]$pkg)
    if (-not (Test-Path $MetadataFile)) { return }
    try {
        $parsedJson = Get-Content $MetadataFile -Raw -Encoding UTF8 | ConvertFrom-Json
        $arr = @()
        if ($null -ne $parsedJson) {
            if ($parsedJson -is [System.Collections.IEnumerable]) { foreach ($item in $parsedJson) { $arr += $item } }
            else { $arr += $parsedJson }
        }
        $pkgName  = $pkg.name
        $pkgOwner = $pkg.owner
        $pkgRepo  = $pkg.repo
        $branch   = if ($pkg.ContainsKey('branch')) { $pkg.branch } else { 'main' }

        $subNames = @()
        foreach ($entry in $arr) {
            $ePkg = $entry.package; $eSc = $entry.script
            if ($ePkg -eq $pkgName -and $null -ne $eSc -and $eSc -ne '' -and $eSc -ne 'standalone') {
                $subNames += $eSc
            }
        }
        $changed = $false

        $newPkgs = @()
        foreach ($p in $script:packages) {
            $keep = $true
            if ($p.ContainsKey('subPackOf') -and $p.subPackOf -eq $pkgName) {
                $pName  = $p.name
                $inMeta = $false
                foreach ($sn in $subNames) { if ($sn -eq $pName) { $inMeta = $true; break } }
                if (-not $inMeta) { $keep = $false; $changed = $true }
            }
            if ($keep) { $newPkgs += $p }
        }
        if ($changed) { $script:packages = $newPkgs }

        foreach ($sName in $subNames) {
            $found = $false
            foreach ($p in $script:packages) {
                if ($p.ContainsKey('subPackOf') -and $p.subPackOf -eq $pkgName -and $p.name -eq $sName) {
                    $found = $true; break
                }
            }
            if (-not $found) {
                $script:packages += @{ name=$sName; owner=$pkgOwner; repo=$pkgRepo; branch=$branch; subPackOf=$pkgName }
                $changed = $true
            }
        }
        if ($changed) { Rebuild-ListView }
    } catch {}
}

if (-not (Test-Path $BaseDir)) { New-Item -ItemType Directory -Path $BaseDir | Out-Null }

$StatusCacheFile = Join-Path $BaseDir 'scripts_status_cache.json'

function Read-StatusCache {
    if (-not (Test-Path $StatusCacheFile)) { return @{} }
    try {
        $obj = Get-Content $StatusCacheFile -Raw | ConvertFrom-Json
        $ht  = @{}
        if ($obj) { $obj.PSObject.Properties | ForEach-Object { $ht[$_.Name] = $_.Value } }
        return $ht
    } catch { return @{} }
}

function Save-StatusCache {
    param([hashtable]$cache)
    $obj = [PSCustomObject]@{}
    foreach ($k in $cache.Keys) { $obj | Add-Member -NotePropertyName $k -NotePropertyValue ([PSCustomObject]$cache[$k]) }
    ConvertTo-Json -InputObject $obj -Depth 5 | Set-Content $StatusCacheFile
}

function Get-PackageStatus {
    param([hashtable]$pkg, [hashtable]$cache)
    $local    = Read-InstalledSHA $pkg
    $cacheKey = if ($pkg.ContainsKey('subPackOf')) { $pkg.subPackOf } else { $pkg.name }
    $entry    = $cache[$cacheKey]
    $remote   = if ($null -ne $entry) { $entry.remoteSha } else { $null }
    $stale    = $true
    if ($null -ne $entry -and $null -ne $entry.checkedAt) {
        try { $stale = ((Get-Date) - [datetime]::Parse($entry.checkedAt)).TotalMinutes -gt 30 } catch {}
    }
    if (-not $local) {
        return [PSCustomObject]@{ Icon=''; Label='Not installed'; Color=[System.Drawing.Color]::DimGray }
    }
    if (-not $remote) {
        if ($null -eq $entry) {
            return [PSCustomObject]@{ Icon='?'; Label='Installed - not yet checked'; Color=[System.Drawing.Color]::DarkGoldenrod }
        }
        if (-not $stale) {
            return [PSCustomObject]@{ Icon='X'; Label='Version check failed - try again in a few minutes'; Color=[System.Drawing.Color]::Salmon }
        }
        return [PSCustomObject]@{ Icon='?'; Label='Installed - checking...'; Color=[System.Drawing.Color]::DarkGoldenrod }
    }
    if ($local -eq $remote) {
        return [PSCustomObject]@{ Icon=[char]0x2714; Label='Up to date'; Color=[System.Drawing.Color]::LightGreen }
    }
    return [PSCustomObject]@{ Icon='!'; Label='Update available'; Color=[System.Drawing.Color]::Gold }
}

$script:logBox         = $null
$script:logAtLineStart = $true

function Write-Host {
    param([Parameter(Position=0,ValueFromPipeline)]$Object,[switch]$NoNewline,$ForegroundColor,$BackgroundColor,$Separator)
    $fwd = @{}
    if ($PSBoundParameters.ContainsKey('ForegroundColor')) { $fwd['ForegroundColor'] = $ForegroundColor }
    if ($NoNewline) { $fwd['NoNewline'] = $true }
    Microsoft.PowerShell.Utility\Write-Host $Object @fwd

    if ($null -ne $script:logBox) {
        $txt = if ($null -ne $Object) { "$Object" } else { '' }
        if ($script:logAtLineStart -and $txt -ne '') {
            $txt = "[$(Get-Date -Format 'HH:mm:ss')] $txt"
        }
        if ($NoNewline) {
            $script:logAtLineStart = $false
        } else {
            $txt += "`r`n"
            $script:logAtLineStart = $true
        }
        $col = [System.Drawing.Color]::WhiteSmoke
        if ($PSBoundParameters.ContainsKey('ForegroundColor')) {
            switch ($ForegroundColor.ToString()) {
                'Green'    { $col = [System.Drawing.Color]::LightGreen }
                'Red'      { $col = [System.Drawing.Color]::Salmon }
                'Yellow'   { $col = [System.Drawing.Color]::Gold }
                'Cyan'     { $col = [System.Drawing.Color]::Cyan }
                'DarkGray' { $col = [System.Drawing.Color]::DarkGray }
                'Magenta'  { $col = [System.Drawing.Color]::Violet }
                'White'    { $col = [System.Drawing.Color]::White }
            }
        }
        $script:logBox.SelectionStart  = $script:logBox.TextLength
        $script:logBox.SelectionLength = 0
        $script:logBox.SelectionColor  = $col
        $script:logBox.AppendText($txt)
        $script:logBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }
}

# Writes a visible divider line in the log to separate action sections.
function Write-LogSection {
    param([string]$title)
    Write-Host "---- $title ----" -ForegroundColor White
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

Add-Type -Name ConsoleWindow -Namespace Util -MemberDefinition @'
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]   public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@

$script:statusCache = Read-StatusCache

$form = New-Object System.Windows.Forms.Form
$form.Text          = 'QQT Script Manager'
$form.Size          = [System.Drawing.Size]::new(800, 660)
$form.MinimumSize   = [System.Drawing.Size]::new(620, 520)
$form.StartPosition = 'CenterScreen'
$form.BackColor     = [System.Drawing.Color]::FromArgb(30, 30, 30)
$form.ForeColor     = [System.Drawing.Color]::WhiteSmoke
$form.Font          = New-Object System.Drawing.Font('Segoe UI', 9)

$split = New-Object System.Windows.Forms.SplitContainer
$split.Dock             = 'Fill'
$split.Orientation      = 'Horizontal'
$split.SplitterDistance = 340
$split.SplitterWidth    = 4
$split.BackColor        = [System.Drawing.Color]::FromArgb(60, 60, 60)
$split.Panel1.Padding   = [System.Windows.Forms.Padding]::new(8, 8, 8, 0)
$split.Panel2.Padding   = [System.Windows.Forms.Padding]::new(8, 4, 8, 4)
$split.Panel1.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$split.Panel2.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)

$lv = New-Object System.Windows.Forms.ListView
$lv.Dock          = 'Fill'
$lv.View          = 'Details'
$lv.FullRowSelect = $true
$lv.MultiSelect   = $true
$lv.GridLines     = $false
$lv.BackColor     = [System.Drawing.Color]::FromArgb(40, 40, 40)
$lv.ForeColor     = [System.Drawing.Color]::WhiteSmoke
$lv.BorderStyle   = 'FixedSingle'
$lv.HeaderStyle   = 'Nonclickable'
[void]$lv.Columns.Add('',        30)
[void]$lv.Columns.Add('Package', 190)
[void]$lv.Columns.Add('Author',  130)
[void]$lv.Columns.Add('Status',  300)
$split.Panel1.Controls.Add($lv)

$btnBar = New-Object System.Windows.Forms.FlowLayoutPanel
$btnBar.Dock      = 'Bottom'
$btnBar.AutoSize  = $true
$btnBar.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$btnBar.Padding   = [System.Windows.Forms.Padding]::new(0, 6, 0, 0)

function New-Btn($label, [System.Drawing.Color]$bg) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text      = $label
    $b.AutoSize  = $true
    $b.BackColor = $bg
    $b.ForeColor = [System.Drawing.Color]::White
    $b.FlatStyle = 'Flat'
    $b.FlatAppearance.BorderSize = 0
    $b.Margin    = [System.Windows.Forms.Padding]::new(0, 0, 6, 0)
    $b.Padding   = [System.Windows.Forms.Padding]::new(10, 4, 10, 4)
    return $b
}

$blue   = [System.Drawing.Color]::FromArgb(0,  120, 215)
$green  = [System.Drawing.Color]::FromArgb(0,  140,   0)
$orange = [System.Drawing.Color]::FromArgb(180, 90,   0)
$grey   = [System.Drawing.Color]::FromArgb(80,  80,  80)
$red    = [System.Drawing.Color]::FromArgb(160,  0,   0)

$btnUpdateAll   = New-Btn 'Update All Installed' $green
$btnDownloadAll = New-Btn 'Download All'         $orange
$btnRefresh     = New-Btn 'Force Refresh'        $grey
$btnAddRepo     = New-Btn 'Add Repository'       $grey
$btnToken       = New-Btn 'GitHub Token'         $grey
$btnExit        = New-Btn 'Exit'                 $red
[void]$btnBar.Controls.AddRange(@($btnUpdateAll,$btnDownloadAll,$btnRefresh,$btnAddRepo,$btnToken,$btnExit))
$split.Panel1.Controls.Add($btnBar)

$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.Dock        = 'Fill'
$logBox.ReadOnly    = $true
$logBox.BackColor   = [System.Drawing.Color]::FromArgb(20, 20, 20)
$logBox.ForeColor   = [System.Drawing.Color]::WhiteSmoke
$logBox.Font        = New-Object System.Drawing.Font('Consolas', 9)
$logBox.BorderStyle = 'FixedSingle'
$script:logBox      = $logBox
$split.Panel2.Controls.Add($logBox)

$strip = New-Object System.Windows.Forms.StatusStrip
$strip.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
$statusLbl = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLbl.Text      = 'Ready'
$statusLbl.ForeColor = [System.Drawing.Color]::WhiteSmoke
[void]$strip.Items.Add($statusLbl)

$form.Controls.Add($split)
$form.Controls.Add($strip)

function Get-PkgTag([hashtable]$pkg) {
    if ($pkg.ContainsKey('subPackOf')) { return "$($pkg.subPackOf)/$($pkg.name)" }
    return $pkg.name
}

function Find-PkgByTag([string]$tag) {
    if ($tag -match '^(.+)/(.+)$') {
        $pName = $Matches[1]; $cName = $Matches[2]
        return $script:packages | Where-Object { $_.ContainsKey('subPackOf') -and $_.subPackOf -eq $pName -and $_.name -eq $cName } | Select-Object -First 1
    }
    return $script:packages | Where-Object { $_.name -eq $tag -and -not $_.ContainsKey('subPackOf') } | Select-Object -First 1
}

function Get-DisplayName([hashtable]$pkg) {
    if ($pkg.ContainsKey('isCollection') -and $pkg.isCollection) { return "[P] $($pkg.name)" }
    if ($pkg.ContainsKey('subPackOf'))                           { return "  - $($pkg.name)" }
    return $pkg.name
}

# Wraps Get-PackageStatus and adds "another version installed" detection for
# standalone packages that also exist as a sub-pack of an installed collection.
function Get-ItemStatus([hashtable]$pkg, [hashtable]$cache) {
    $s = Get-PackageStatus $pkg $cache
    if (-not $pkg.ContainsKey('subPackOf') -and
        -not ($pkg.ContainsKey('isCollection') -and $pkg.isCollection) -and
        $null -eq (Read-InstalledSHA $pkg)) {
        $alt = $script:packages | Where-Object { $_.ContainsKey('subPackOf') -and $_.name -eq $pkg.name } | Select-Object -First 1
        if ($alt -and $null -ne (Read-InstalledSHA $alt)) {
            return [PSCustomObject]@{ Icon='~'; Label='Another version installed'; Color=[System.Drawing.Color]::Goldenrod }
        }
    }
    return $s
}

function Update-ListItem($item, [hashtable]$pkg) {
    $s = Get-ItemStatus $pkg $script:statusCache
    $item.Text             = $s.Icon
    $item.SubItems[1].Text = Get-DisplayName $pkg
    $item.SubItems[2].Text = $pkg.owner
    $item.SubItems[3].Text = $s.Label
    $item.ForeColor        = $s.Color
    $item.Tag              = Get-PkgTag $pkg
}

function Rebuild-ListView {
    $lv.BeginUpdate()
    $lv.Items.Clear()
    $topLevel     = @($script:packages | Where-Object { -not $_.ContainsKey('subPackOf') -and $_.name -and $_.name -ne 'standalone' })
    $installed    = @($topLevel | Where-Object { $null -ne (Read-InstalledSHA $_) } | Sort-Object @{
        Expression = { $d = Get-InstallDir $_; if (Test-Path $d) { (Get-Item $d).LastWriteTime } else { [datetime]::MinValue } }
    })
    $notInstalled = @($topLevel | Where-Object { $null -eq (Read-InstalledSHA $_) } | Sort-Object { $_.name })
    $topLevel     = $installed + $notInstalled
    foreach ($pkg in $topLevel) {
        $s    = Get-ItemStatus $pkg $script:statusCache
        $item = New-Object System.Windows.Forms.ListViewItem($s.Icon)
        [void]$item.SubItems.Add((Get-DisplayName $pkg))
        [void]$item.SubItems.Add($pkg.owner)
        [void]$item.SubItems.Add($s.Label)
        $item.ForeColor = $s.Color
        $item.Tag       = Get-PkgTag $pkg
        [void]$lv.Items.Add($item)
        if ($pkg.ContainsKey('isCollection') -and $pkg.isCollection) {
            $subPacks = @($script:packages | Where-Object {
                $_.ContainsKey('subPackOf') -and $_.subPackOf -eq $pkg.name -and $_.name -and $_.name -ne 'standalone'
            })
            foreach ($sp in $subPacks) {
                $ss    = Get-PackageStatus $sp $script:statusCache
                $sItem = New-Object System.Windows.Forms.ListViewItem($ss.Icon)
                [void]$sItem.SubItems.Add((Get-DisplayName $sp))
                [void]$sItem.SubItems.Add($sp.owner)
                [void]$sItem.SubItems.Add($ss.Label)
                $sItem.ForeColor = $ss.Color
                $sItem.Tag       = Get-PkgTag $sp
                [void]$lv.Items.Add($sItem)
            }
        }
    }
    $lv.EndUpdate()
}

Rebuild-ListView

foreach ($p in @($script:packages | Where-Object { $_.ContainsKey('isCollection') -and $_.isCollection })) {
    Sync-CollectionSubPacks $p
}

$script:refreshHandles = @()
$script:refreshPool    = $null

$refreshTimer          = New-Object System.Windows.Forms.Timer
$refreshTimer.Interval = 300

$refreshTimer.Add_Tick({
    $allDone = $true
    foreach ($h in $script:refreshHandles) {
        if ($h.Done) { continue }
        if (-not $h.Handle.IsCompleted) { $allDone = $false; continue }
        $raw = $h.PS.EndInvoke($h.Handle)
        $h.PS.Dispose()
        $h.Done = $true
        $sha = if ($raw -is [System.Array]) { $raw[0] } else { $raw }
        $script:statusCache[$h.Pkg.name] = @{ remoteSha=$sha; checkedAt=(Get-Date -Format 'o') }
        foreach ($item in $lv.Items) {
            $itemPkg = Find-PkgByTag $item.Tag
            if ($null -eq $itemPkg) { continue }
            $ownerKey = if ($itemPkg.ContainsKey('subPackOf')) { $itemPkg.subPackOf } else { $itemPkg.name }
            if ($ownerKey -eq $h.Pkg.name) { Update-ListItem $item $itemPkg }
        }
    }
    if ($allDone) {
        $refreshTimer.Stop()
        if ($script:refreshPool) {
            try { $script:refreshPool.Close(); $script:refreshPool.Dispose() } catch {}
            $script:refreshPool = $null
        }
        Save-StatusCache $script:statusCache
        $btnRefresh.Enabled = $true
        $btnRefresh.Text    = 'Force Refresh'
        $updatesAvailable = @($script:packages | Where-Object {
            (Get-PackageStatus $_ $script:statusCache).Label -eq 'Update available'
        }).Count
        if ($updatesAvailable -gt 0) {
            $statusLbl.Text      = "$updatesAvailable update(s) available!  -  Last checked $(Get-Date -Format 'HH:mm')"
            $statusLbl.ForeColor = [System.Drawing.Color]::Gold
        } else {
            $statusLbl.Text      = "All up to date  -  Last checked $(Get-Date -Format 'HH:mm')"
            $statusLbl.ForeColor = [System.Drawing.Color]::LightGreen
        }
    }
})

function Start-StatusRefresh {
    param([bool]$force = $false)
    if ($script:refreshPool) { return }
    if ($force) {
        $toCheck = @($script:packages)
    } else {
        $toCheck = @($script:packages | Where-Object {
            $e = $script:statusCache[$_.name]
            if ($null -eq $e -or $null -eq $e.checkedAt) { return $true }
            try { return ((Get-Date) - [datetime]::Parse($e.checkedAt)).TotalMinutes -gt 30 } catch { return $true }
        })
    }
    if ($toCheck.Count -eq 0) {
        $statusLbl.Text = 'All statuses are fresh (< 30 min). Use Force Refresh to recheck.'
        return
    }
    $btnRefresh.Enabled = $false
    $btnRefresh.Text    = 'Refreshing...'
    $statusLbl.Text     = "Checking $($toCheck.Count) package(s) on GitHub..."
    $script:refreshHandles = @()
    $pool = [runspacefactory]::CreateRunspacePool(1, [Math]::Max(1, $toCheck.Count))
    $pool.Open()
    $script:refreshPool = $pool
    foreach ($pkg in $toCheck) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $pool
        $pkgBranch = if ($pkg.ContainsKey('branch')) { $pkg.branch } else { 'main' }
        [void]$ps.AddScript({
            param($owner,$repo,$branch,$tok)
            $h = @{ 'User-Agent'='script-updater/1.0' }
            if ($tok) { $h['Authorization']="token $tok" }
            try { (Invoke-RestMethod -Uri "https://api.github.com/repos/$owner/$repo/commits/$branch" -Headers $h -UseBasicParsing).sha } catch { $null }
        }).AddArgument($pkg.owner).AddArgument($pkg.repo).AddArgument($pkgBranch).AddArgument($env:GITHUB_TOKEN)
        $script:refreshHandles += [PSCustomObject]@{ PS=$ps; Handle=$ps.BeginInvoke(); Pkg=$pkg; Done=$false }
    }
    $refreshTimer.Start()
}

$autoCheckTimer          = New-Object System.Windows.Forms.Timer
$autoCheckTimer.Interval = 30 * 60 * 1000
$autoCheckTimer.Add_Tick({ if (-not $script:refreshPool) { Start-StatusRefresh -force $false } })
$autoCheckTimer.Start()

function Set-ButtonsEnabled([bool]$on) {
    foreach ($b in @($btnUpdateAll,$btnDownloadAll,$btnAddRepo,$btnToken)) { $b.Enabled = $on }
    if ($on) { $btnRefresh.Enabled = $true }
}

# Shared install/update logic used by the context menu and double-click handler.
function Invoke-InstallSelected {
    param([bool]$forceInstall = $false)
    $sel = @($lv.SelectedItems)
    if ($sel.Count -eq 0) { return }
    Set-ButtonsEnabled $false
    $processed = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($item in $sel) {
        $pkg = Find-PkgByTag $item.Tag
        if (-not $pkg) { continue }
        if ($pkg.ContainsKey('subPackOf')) {
            $pkg = $script:packages | Where-Object { $_.name -eq $pkg.subPackOf -and -not $_.ContainsKey('subPackOf') } | Select-Object -First 1
        }
        if (-not $pkg) { continue }
        if (-not $processed.Add($pkg.name)) { continue }
        $statusLbl.Text = "Processing $($pkg.name)..."
        if ($forceInstall) { Install-Package $pkg } else { Update-Package $pkg }
        if ($pkg.ContainsKey('isCollection') -and $pkg.isCollection) {
            $purgeName = $pkg.name
            $script:packages = @($script:packages | Where-Object { -not ($_.ContainsKey('subPackOf') -and $_.subPackOf -eq $purgeName) })
            Sync-CollectionSubPacks $pkg
        }
    }
    Rebuild-ListView
    Set-ButtonsEnabled $true
    $statusLbl.Text = 'Done.'
}

$btnUpdateAll.Add_Click({
    Set-ButtonsEnabled $false
    Write-LogSection 'Update All Installed'
    $statusLbl.Text = 'Updating all installed packages...'
    Update-AllInstalled
    foreach ($p in @($script:packages | Where-Object { $_.ContainsKey('isCollection') -and $_.isCollection })) { Sync-CollectionSubPacks $p }
    Rebuild-ListView
    Set-ButtonsEnabled $true
    $statusLbl.Text = 'Update complete.'
})

$btnDownloadAll.Add_Click({
    $r = [System.Windows.Forms.MessageBox]::Show('This will download/reinstall ALL packages. Continue?','Confirm','YesNo','Question')
    if ($r -ne 'Yes') { return }
    Set-ButtonsEnabled $false
    Write-LogSection 'Download All'
    $statusLbl.Text = 'Downloading all packages...'
    Download-AllPackages
    foreach ($p in @($script:packages | Where-Object { $_.ContainsKey('isCollection') -and $_.isCollection })) { Sync-CollectionSubPacks $p }
    Rebuild-ListView
    Set-ButtonsEnabled $true
    $statusLbl.Text = 'Download complete.'
})

$btnRefresh.Add_Click({
    $collections = @()
    foreach ($p in $script:packages) {
        if ($p.ContainsKey('isCollection') -and $p.isCollection) { $collections += $p }
    }
    foreach ($coll in $collections) {
        $collN  = $coll.name
        $purged = @()
        foreach ($p in $script:packages) {
            if ($p.ContainsKey('subPackOf') -and $p.subPackOf -eq $collN) { continue }
            $purged += $p
        }
        $script:packages = $purged
        Sync-CollectionSubPacks $coll
    }
    Rebuild-ListView
    Start-StatusRefresh -force $true
})

$btnAddRepo.Add_Click({
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = 'Add GitHub Repository'
    $dlg.Size            = [System.Drawing.Size]::new(440, 230)
    $dlg.StartPosition   = 'CenterParent'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox     = $false
    $dlg.BackColor       = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $dlg.ForeColor       = [System.Drawing.Color]::WhiteSmoke

    function New-DlgLabel($txt, $x, $y) {
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $txt; $l.AutoSize = $true; $l.Location = [System.Drawing.Point]::new($x,$y); return $l
    }
    function New-DlgText($x, $y, $w) {
        $t = New-Object System.Windows.Forms.TextBox
        $t.Location = [System.Drawing.Point]::new($x,$y); $t.Width = $w
        $t.BackColor = [System.Drawing.Color]::FromArgb(50,50,50); $t.ForeColor = [System.Drawing.Color]::White; return $t
    }

    $txtUrl    = New-DlgText 10 36 410
    $txtBranch = New-DlgText 10 88 200
    $chkColl   = New-Object System.Windows.Forms.CheckBox
    $chkColl.Text     = 'Collection (multiple scripts in subdirectories)'
    $chkColl.Location = [System.Drawing.Point]::new(10,118); $chkColl.AutoSize = $true

    $btnOk  = New-Btn 'Add'    $blue; $btnOk.Location  = [System.Drawing.Point]::new(250,158); $btnOk.DialogResult  = 'OK'
    $btnCan = New-Btn 'Cancel' $grey; $btnCan.Location = [System.Drawing.Point]::new(340,158); $btnCan.DialogResult = 'Cancel'
    $dlg.Controls.AddRange(@(
        (New-DlgLabel 'GitHub URL (e.g. https://github.com/owner/repo):' 10 14),
        $txtUrl,
        (New-DlgLabel 'Branch (leave blank for main):' 10 68),
        $txtBranch, $chkColl, $btnOk, $btnCan
    ))
    $dlg.AcceptButton = $btnOk; $dlg.CancelButton = $btnCan

    if ($dlg.ShowDialog($form) -ne 'OK') { return }
    $url = $txtUrl.Text.Trim()
    if (-not $url) { return }
    $m = [regex]::Match($url, '(?:https?://)?github\.com/([^/]+)/([^/]+?)(?:\.git|/?$)')
    if (-not $m.Success) {
        [void][System.Windows.Forms.MessageBox]::Show("Invalid GitHub URL: $url",'Error','OK','Error'); return
    }
    $owner  = $m.Groups[1].Value
    $repo   = $m.Groups[2].Value
    $branch = if ($txtBranch.Text.Trim()) { $txtBranch.Text.Trim() } else { 'main' }
    $newPkg = @{ name=$repo; owner=$owner; repo=$repo; branch=$branch }
    if ($chkColl.Checked) { $newPkg.isCollection = $true }
    $script:packages += $newPkg
    Rebuild-ListView
    Write-Host "Added $repo by $owner (branch: $branch)." -ForegroundColor Green
})

$btnToken.Add_Click({
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = 'GitHub Token'
    $dlg.Size            = [System.Drawing.Size]::new(460, 160)
    $dlg.StartPosition   = 'CenterParent'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox     = $false
    $dlg.BackColor       = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $dlg.ForeColor       = [System.Drawing.Color]::WhiteSmoke

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text     = 'Paste your GitHub personal access token (public_repo scope):'
    $lbl.AutoSize = $true
    $lbl.Location = [System.Drawing.Point]::new(10, 14)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Location    = [System.Drawing.Point]::new(10, 36)
    $txt.Width       = 430
    $txt.BackColor   = [System.Drawing.Color]::FromArgb(50, 50, 50)
    $txt.ForeColor   = [System.Drawing.Color]::White
    $txt.Text        = if ($env:GITHUB_TOKEN) { $env:GITHUB_TOKEN } else { '' }
    $txt.UseSystemPasswordChar = $true

    $chkShow = New-Object System.Windows.Forms.CheckBox
    $chkShow.Text     = 'Show token'
    $chkShow.AutoSize = $true
    $chkShow.Location = [System.Drawing.Point]::new(10, 64)
    $chkShow.Add_CheckedChanged({ $txt.UseSystemPasswordChar = -not $chkShow.Checked })

    $btnOk  = New-Btn 'Save'   $blue; $btnOk.Location  = [System.Drawing.Point]::new(270, 90); $btnOk.DialogResult  = 'OK'
    $btnCan = New-Btn 'Cancel' $grey; $btnCan.Location = [System.Drawing.Point]::new(360, 90); $btnCan.DialogResult = 'Cancel'
    $btnClear = New-Btn 'Clear Token' $grey; $btnClear.Location = [System.Drawing.Point]::new(10, 90)
    $btnClear.Add_Click({ $txt.Text = '' })

    $dlg.Controls.AddRange(@($lbl, $txt, $chkShow, $btnOk, $btnCan, $btnClear))
    $dlg.AcceptButton = $btnOk; $dlg.CancelButton = $btnCan

    if ($dlg.ShowDialog($form) -ne 'OK') { return }
    $token = $txt.Text.Trim()
    if ($token) {
        $env:GITHUB_TOKEN = $token
        Write-Host "GitHub token saved for this session." -ForegroundColor Green
    } else {
        Remove-Item Env:GITHUB_TOKEN -ErrorAction SilentlyContinue
        Write-Host "GitHub token cleared." -ForegroundColor Yellow
    }
})

$lv.Add_DoubleClick({ if ($lv.SelectedItems.Count -gt 0) { Invoke-InstallSelected } })

$ctxMenu      = New-Object System.Windows.Forms.ContextMenuStrip
$ctxInstall   = New-Object System.Windows.Forms.ToolStripMenuItem('Install')
$ctxUpdate    = New-Object System.Windows.Forms.ToolStripMenuItem('Update')
$ctxUninstall = New-Object System.Windows.Forms.ToolStripMenuItem('Uninstall')
[void]$ctxMenu.Items.Add($ctxInstall)
[void]$ctxMenu.Items.Add($ctxUpdate)
[void]$ctxMenu.Items.Add($ctxUninstall)

$ctxMenu.Add_Opening({
    $sel = @($lv.SelectedItems)
    if ($sel.Count -eq 0) {
        $ctxInstall.Visible   = $true; $ctxInstall.Enabled   = $false
        $ctxUpdate.Visible    = $false
        $ctxUninstall.Visible = $false
        return
    }
    $firstPkg  = Find-PkgByTag $sel[0].Tag
    $installed = $false
    if ($firstPkg) {
        $checkPkg = if ($firstPkg.ContainsKey('subPackOf')) {
            $script:packages | Where-Object { $_.name -eq $firstPkg.subPackOf -and -not $_.ContainsKey('subPackOf') } | Select-Object -First 1
        } else { $firstPkg }
        if ($checkPkg) { $installed = $null -ne (Read-InstalledSHA $checkPkg) }
    }
    $ctxInstall.Visible   = -not $installed; $ctxInstall.Enabled   = $true
    $ctxUpdate.Visible    = $installed;       $ctxUpdate.Enabled    = $true
    $ctxUninstall.Visible = $installed;       $ctxUninstall.Enabled = $true
})

$ctxInstall.Add_Click({ Invoke-InstallSelected -forceInstall $true })
$ctxUpdate.Add_Click({  Invoke-InstallSelected })

$ctxUninstall.Add_Click({
    $sel = @($lv.SelectedItems)
    if ($sel.Count -eq 0) { return }
    $names = ($sel | ForEach-Object { $_.SubItems[1].Text }) -join ', '
    $hasCollection = $false
    foreach ($item in $sel) {
        $p = Find-PkgByTag $item.Tag
        if ($p -and $p.ContainsKey('isCollection') -and $p.isCollection) { $hasCollection = $true; break }
    }
    $msg = "Uninstall the following and delete their files?`n`n$names"
    if ($hasCollection) { $msg += "`n`nNote: all associated scripts inside the collection will also be deleted." }
    $r = [System.Windows.Forms.MessageBox]::Show($msg, 'Confirm Uninstall', 'YesNo', 'Warning')
    if ($r -ne 'Yes') { return }
    Set-ButtonsEnabled $false
    Write-LogSection 'Uninstall'
    foreach ($item in $sel) {
        $pkg = Find-PkgByTag $item.Tag
        if (-not $pkg) { continue }
        $statusLbl.Text = "Uninstalling $($pkg.name)..."
        if ($pkg.ContainsKey('subPackOf')) {
            $dir = Get-InstallDir $pkg
            if (Test-Path $dir) {
                Remove-Item $dir -Recurse -Force -ErrorAction SilentlyContinue
                Write-Host "Removed $($pkg.name) directory." -ForegroundColor Yellow
            }
            if (Test-Path $MetadataFile) {
                try {
                    $rawSub = Get-Content $MetadataFile -Raw -Encoding UTF8 | ConvertFrom-Json
                    $allSub = @()
                    if ($null -ne $rawSub) {
                        if ($rawSub -is [System.Collections.IEnumerable]) { foreach ($e in $rawSub) { $allSub += $e } }
                        else { $allSub += $rawSub }
                    }
                    $filtered = @($allSub | Where-Object { -not ($_.package -eq $pkg.subPackOf -and $_.script -eq $pkg.name) })
                    $parts = @($filtered | ForEach-Object { ConvertTo-Json -InputObject $_ -Depth 10 })
                    Set-Content -Path $MetadataFile -Value ('[' + ($parts -join ',') + ']') -Encoding UTF8
                } catch {}
            }
            $tagToRemove = Get-PkgTag $pkg
            $script:packages = @($script:packages | Where-Object { (Get-PkgTag $_) -ne $tagToRemove })
            Write-Host "Uninstalled $($pkg.name)." -ForegroundColor Green
        } else {
            Uninstall-Package $pkg
            if ($pkg.ContainsKey('isCollection') -and $pkg.isCollection) {
                $script:packages = @($script:packages | Where-Object { -not ($_.ContainsKey('subPackOf') -and $_.subPackOf -eq $pkg.name) })
            }
        }
    }
    Rebuild-ListView
    Set-ButtonsEnabled $true
    $statusLbl.Text = 'Uninstall complete.'
})

$lv.ContextMenuStrip = $ctxMenu

$btnExit.Add_Click({ $form.Close() })
$form.Add_FormClosing({
    $refreshTimer.Stop()
    $autoCheckTimer.Stop()
    if ($script:refreshPool) { try { $script:refreshPool.Close(); $script:refreshPool.Dispose() } catch {}; $script:refreshPool = $null }
    $script:logBox = $null
    $hwnd = [Util.ConsoleWindow]::GetConsoleWindow()
    if ($hwnd -ne [IntPtr]::Zero) { [Util.ConsoleWindow]::ShowWindow($hwnd, 5) | Out-Null }
})

$form.Add_Shown({
    $form.Activate()
    $hwnd = [Util.ConsoleWindow]::GetConsoleWindow()
    if ($hwnd -ne [IntPtr]::Zero) { [Util.ConsoleWindow]::ShowWindow($hwnd, 0) | Out-Null }
    Start-StatusRefresh -force $false
})

[System.Windows.Forms.Application]::Run($form)
