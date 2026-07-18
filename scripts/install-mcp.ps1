#Requires -Version 5.1
<#
.SYNOPSIS
Registers EULE Outlook MCP with local AI agents: Claude Code, Claude Desktop, Codex, and Antigravity.

.DESCRIPTION
Finds (or builds) OutlookMcp.Server.exe, then adds an "outlook" MCP server entry to every
detected AI client configuration. Existing configuration files are backed up next to the
original before every change, and re-running the script updates entries in place.

The script works from all three distribution layouts:
- Installed copy: script next to OutlookMcp.Server.exe under %LocalAppData%\Programs\EULE Outlook MCP.
- Portable ZIP: script next to OutlookMcp.Server.exe in any stable folder.
- Source checkout: script under scripts\; builds with the .NET 8 SDK and installs the
  executable to %LocalAppData%\Programs\EULE Outlook MCP so client paths stay stable.

.EXAMPLE
powershell -ExecutionPolicy Bypass -File .\install-mcp.ps1
Registers the server with every detected client.

.EXAMPLE
powershell -ExecutionPolicy Bypass -File .\install-mcp.ps1 -Clients codex,antigravity -DryRun
Shows what would change for Codex and Antigravity without writing anything.

.EXAMPLE
powershell -ExecutionPolicy Bypass -File .\install-mcp.ps1 -Remove
Removes the outlook entry from every detected client configuration.
#>
[CmdletBinding()]
param(
    # Which clients to configure. 'auto' targets every detected client; 'all' forces every
    # supported client, creating configuration files that do not exist yet.
    [ValidateSet('auto', 'all', 'claude-code', 'claude-desktop', 'codex', 'antigravity')]
    [string[]]$Clients = @('auto'),

    # Full path to OutlookMcp.Server.exe. Overrides automatic discovery.
    [string]$ExePath,

    # MCP server name registered in each client.
    [ValidatePattern('^[A-Za-z0-9][A-Za-z0-9_-]*$')]
    [string]$ServerName = 'outlook',

    # Runtime identifier used when building from source.
    [ValidateSet('win-x64', 'win-x86')]
    [string]$Runtime = 'win-x64',

    # Advertised MCP tool set. Compact minimizes model-context and credit usage.
    [ValidateSet('compact', 'mail', 'style', 'full')]
    [string]$ToolProfile = 'compact',

    # Rebuild from source even when an existing executable is found.
    [switch]$Rebuild,

    # Remove the server entry from the selected clients instead of adding it.
    [switch]$Remove,

    # Print planned changes without modifying anything.
    [switch]$DryRun,

    # Wait for Enter before exiting (used by the installer's post-install step).
    [switch]$Pause
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$script:Results = New-Object System.Collections.Generic.List[object]
$script:CanonicalInstallDir = Join-Path $env:LOCALAPPDATA 'Programs\EULE Outlook MCP'

$script:ClientPaths = @{
    ClaudeDesktopConfig     = Join-Path $env:APPDATA 'Claude\claude_desktop_config.json'
    ClaudeDesktopDir        = Join-Path $env:APPDATA 'Claude'
    ClaudeCodeUserConfig    = Join-Path $env:USERPROFILE '.claude.json'
    CodexDir                = Join-Path $env:USERPROFILE '.codex'
    CodexConfig             = Join-Path $env:USERPROFILE '.codex\config.toml'
    AntigravityDir          = Join-Path $env:USERPROFILE '.gemini\antigravity'
    AntigravityConfig       = Join-Path $env:USERPROFILE '.gemini\antigravity\mcp_config.json'
    AntigravitySharedConfig = Join-Path $env:USERPROFILE '.gemini\config\mcp_config.json'
    AntigravityAppDir       = Join-Path $env:LOCALAPPDATA 'Programs\Antigravity'
}

function Write-Step([string]$Message) { Write-Host $Message -ForegroundColor Cyan }
function Write-Note([string]$Message) { Write-Host "  $Message" -ForegroundColor DarkGray }

function Add-Result([string]$Client, [string]$Status, [string]$Detail) {
    $script:Results.Add([pscustomobject]@{ Client = $Client; Status = $Status; Detail = $Detail })
}

function Backup-ConfigFile([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return }
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $backup = "$Path.backup-$stamp"
    if ($DryRun) {
        Write-Note "[dry-run] would back up $Path to $backup"
        return
    }
    Copy-Item -LiteralPath $Path -Destination $backup -Force
    Write-Note "Backed up $Path"
}

function Write-TextFile([string]$Path, [string]$Content) {
    if ($DryRun) {
        Write-Note "[dry-run] would write $Path"
        return
    }
    $directory = Split-Path -Parent $Path
    if ($directory -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }
    [System.IO.File]::WriteAllText($Path, $Content, (New-Object System.Text.UTF8Encoding($false)))
}

function Read-JsonConfig([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    $raw = Get-Content -LiteralPath $Path -Raw
    if (-not $raw -or -not $raw.Trim()) { return $null }
    try {
        return $raw | ConvertFrom-Json
    } catch {
        throw "Could not parse existing JSON in '$Path'. Fix or remove that file, then re-run. ($($_.Exception.Message))"
    }
}

function Test-JsonProperty($Object, [string]$Name) {
    return $Object.PSObject.Properties.Match($Name).Count -gt 0
}

# Adds or replaces mcpServers.<ServerName> in a Claude-Desktop-style JSON config.
function Set-JsonServerEntry([string]$Path, [string]$Exe) {
    $config = Read-JsonConfig $Path
    if ($null -eq $config) { $config = New-Object PSObject }
    $hasServers = Test-JsonProperty $config 'mcpServers'
    if (-not $hasServers -or $null -eq $config.mcpServers) {
        $config | Add-Member -NotePropertyName 'mcpServers' -NotePropertyValue (New-Object PSObject) -Force
    }
    $entry = New-Object PSObject
    $entry | Add-Member -NotePropertyName 'command' -NotePropertyValue $Exe
    $entry | Add-Member -NotePropertyName 'args' -NotePropertyValue @('--tool-profile', $ToolProfile)
    $config.mcpServers | Add-Member -NotePropertyName $ServerName -NotePropertyValue $entry -Force
    Backup-ConfigFile $Path
    Write-TextFile $Path ($config | ConvertTo-Json -Depth 100)
}

# Removes mcpServers.<ServerName>; returns $true when an entry was present.
function Remove-JsonServerEntry([string]$Path) {
    $config = Read-JsonConfig $Path
    if ($null -eq $config) { return $false }
    if (-not (Test-JsonProperty $config 'mcpServers')) { return $false }
    if ($null -eq $config.mcpServers) { return $false }
    if (-not (Test-JsonProperty $config.mcpServers $ServerName)) { return $false }
    $config.mcpServers.PSObject.Properties.Remove($ServerName)
    Backup-ConfigFile $Path
    Write-TextFile $Path ($config | ConvertTo-Json -Depth 100)
    return $true
}

# Rewrites ~/.codex/config.toml, replacing the [mcp_servers.<ServerName>] block (and any of
# its sub-tables) while leaving every other line untouched.
function Set-CodexServerEntry([string]$Path, [string]$Exe, [bool]$RemoveOnly) {
    $lines = @()
    if (Test-Path -LiteralPath $Path) { $lines = @(Get-Content -LiteralPath $Path) }
    $sectionPattern = '^\s*\[mcp_servers\.' + [regex]::Escape($ServerName) + '(\.[^\]]+)?\]\s*(#.*)?$'
    $kept = New-Object System.Collections.Generic.List[string]
    $inBlock = $false
    $found = $false
    foreach ($line in $lines) {
        if ($line -match $sectionPattern) { $inBlock = $true; $found = $true; continue }
        if ($inBlock -and $line -match '^\s*\[') { $inBlock = $false }
        if (-not $inBlock) { $kept.Add($line) }
    }
    if ($RemoveOnly -and -not $found) { return $false }
    if (-not $RemoveOnly) {
        while ($kept.Count -gt 0 -and -not $kept[$kept.Count - 1].Trim()) { $kept.RemoveAt($kept.Count - 1) }
        if ($kept.Count -gt 0) { $kept.Add('') }
        if ($Exe.Contains("'")) {
            $escaped = $Exe.Replace('\', '\\').Replace('"', '\"')
            $commandLine = "command = ""$escaped"""
        } else {
            $commandLine = "command = '$Exe'"
        }
        $kept.Add("[mcp_servers.$ServerName]")
        $kept.Add($commandLine)
        $kept.Add("args = [`"--tool-profile`", `"$ToolProfile`"]")
        $kept.Add('startup_timeout_sec = 30')
        $kept.Add('tool_timeout_sec = 60')
        $kept.Add('enabled = true')
    }
    Backup-ConfigFile $Path
    Write-TextFile $Path (($kept -join "`r`n").TrimEnd("`r", "`n") + "`r`n")
    return $true
}

function Invoke-NativeCommand([string]$File, [string[]]$Arguments) {
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $output = & $File @Arguments 2>&1
        $exitCode = 0
        if (Test-Path variable:global:LASTEXITCODE) { $exitCode = $global:LASTEXITCODE }
        return [pscustomobject]@{
            ExitCode = $exitCode
            Output   = (($output | Out-String)).Trim()
        }
    } catch {
        return [pscustomobject]@{ ExitCode = -1; Output = $_.Exception.Message }
    } finally {
        $ErrorActionPreference = $previous
    }
}

function Test-ClientDetected([string]$Client) {
    switch ($Client) {
        'claude-code' {
            return [bool](Get-Command claude -ErrorAction SilentlyContinue) -or
                (Test-Path -LiteralPath $script:ClientPaths.ClaudeCodeUserConfig)
        }
        'claude-desktop' {
            return Test-Path -LiteralPath $script:ClientPaths.ClaudeDesktopDir
        }
        'codex' {
            return [bool](Get-Command codex -ErrorAction SilentlyContinue) -or
                (Test-Path -LiteralPath $script:ClientPaths.CodexDir)
        }
        'antigravity' {
            return (Test-Path -LiteralPath $script:ClientPaths.AntigravityDir) -or
                (Test-Path -LiteralPath $script:ClientPaths.AntigravitySharedConfig) -or
                (Test-Path -LiteralPath $script:ClientPaths.AntigravityAppDir) -or
                [bool](Get-Command antigravity -ErrorAction SilentlyContinue)
        }
    }
    return $false
}

function Resolve-TargetClients {
    $supported = @('claude-code', 'claude-desktop', 'codex', 'antigravity')
    if ($Clients -contains 'all') { return $supported }
    $explicit = @($Clients | Where-Object { $_ -ne 'auto' } | Select-Object -Unique)
    if (-not ($Clients -contains 'auto')) { return $explicit }
    $detected = @($supported | Where-Object { Test-ClientDetected $_ })
    foreach ($skipped in ($supported | Where-Object { $detected -notcontains $_ -and $explicit -notcontains $_ })) {
        Add-Result $skipped 'not detected' 'No installation found; pass it via -Clients to force setup.'
    }
    return @($detected + $explicit | Select-Object -Unique)
}

function Resolve-ServerExecutable {
    if ($ExePath) {
        $resolved = (Resolve-Path -LiteralPath $ExePath).Path
        if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) {
            throw "-ExePath '$ExePath' is not a file."
        }
        return $resolved
    }

    $portable = Join-Path $PSScriptRoot 'OutlookMcp.Server.exe'
    $canonical = Join-Path $script:CanonicalInstallDir 'OutlookMcp.Server.exe'
    $repoRoot = Split-Path -Parent $PSScriptRoot
    $project = Join-Path $repoRoot 'src\OutlookMcp.Server\OutlookMcp.Server.csproj'
    $inRepo = Test-Path -LiteralPath $project

    if (-not ($Rebuild -and $inRepo)) {
        if (Test-Path -LiteralPath $portable) {
            Write-Note "Using executable next to this script: $portable"
            return $portable
        }
        if (Test-Path -LiteralPath $canonical) {
            Write-Note "Using installed executable: $canonical"
            if ($inRepo) { Write-Note 'Pass -Rebuild to refresh it from this source checkout.' }
            return $canonical
        }
    }

    if (-not $inRepo) {
        throw ("OutlookMcp.Server.exe was not found next to this script or under " +
            "'$script:CanonicalInstallDir', and no source checkout is available. " +
            "Run the installer or extract the release ZIP first, or pass -ExePath.")
    }

    if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
        throw 'Building from source requires the .NET 8 SDK (dotnet was not found on PATH).'
    }

    $publishDir = Join-Path $repoRoot "artifacts\publish\$Runtime"
    Write-Step "Building OutlookMcp.Server ($Runtime) from source..."
    if ($DryRun) {
        Write-Note "[dry-run] would run dotnet publish and install to $script:CanonicalInstallDir"
        return $canonical
    }
    # Pipe build output through Write-Host so it displays live without polluting this
    # function's return value (all uncaptured pipeline output becomes part of it).
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    & dotnet publish $project -c Release -r $Runtime --self-contained true -p:PublishSingleFile=true -o $publishDir 2>&1 |
        ForEach-Object { Write-Host $_ }
    $publishExit = $global:LASTEXITCODE
    $ErrorActionPreference = $previous
    if ($publishExit -ne 0) {
        throw 'dotnet publish failed; see the build output above.'
    }

    New-Item -ItemType Directory -Path $script:CanonicalInstallDir -Force | Out-Null
    Copy-Item (Join-Path $publishDir 'OutlookMcp.Server.exe') $script:CanonicalInstallDir -Force
    Copy-Item (Join-Path $repoRoot 'src\OutlookMcp.Server\config.sample.json') $script:CanonicalInstallDir -Force
    Copy-Item (Join-Path $repoRoot 'README.md') $script:CanonicalInstallDir -Force
    Copy-Item $PSCommandPath $script:CanonicalInstallDir -Force
    Write-Note "Installed to $script:CanonicalInstallDir"
    return $canonical
}

function Set-ClaudeCodeClient([string]$Exe) {
    $cli = Get-Command claude -ErrorAction SilentlyContinue
    if (-not $cli) {
        if ($Remove) {
            Add-Result 'claude-code' 'manual' "claude CLI not found. Run: claude mcp remove --scope user $ServerName"
        } else {
            Add-Result 'claude-code' 'manual' "claude CLI not found. Run: claude mcp add --scope user $ServerName -- `"$Exe`" --tool-profile $ToolProfile"
        }
        return
    }
    if ($DryRun) {
        if ($Remove) {
            Add-Result 'claude-code' 'dry-run' "Would run: claude mcp remove --scope user $ServerName"
        } else {
            Add-Result 'claude-code' 'dry-run' "Would run: claude mcp add --scope user $ServerName -- `"$Exe`" --tool-profile $ToolProfile"
        }
        return
    }
    if ($Remove) {
        $removal = Invoke-NativeCommand $cli.Source @('mcp', 'remove', '--scope', 'user', $ServerName)
        if ($removal.ExitCode -eq 0) {
            Add-Result 'claude-code' 'removed' 'Removed from user scope.'
        } else {
            Add-Result 'claude-code' 'skipped' 'No user-scope entry found.'
        }
        return
    }
    # Remove any stale user-scope entry first so add always succeeds, then re-add.
    Invoke-NativeCommand $cli.Source @('mcp', 'remove', '--scope', 'user', $ServerName) | Out-Null
    $addition = Invoke-NativeCommand $cli.Source @('mcp', 'add', '--scope', 'user', $ServerName, '--', $Exe, '--tool-profile', $ToolProfile)
    if ($addition.ExitCode -eq 0) {
        Add-Result 'claude-code' 'configured' 'User scope; restart Claude Code sessions, then verify with /mcp.'
    } else {
        Add-Result 'claude-code' 'failed' "claude mcp add failed: $($addition.Output)"
    }
}

function Get-DoneStatus([string]$Status) {
    if ($DryRun) { return 'dry-run' }
    return $Status
}

function Set-ClaudeDesktopClient([string]$Exe) {
    $path = $script:ClientPaths.ClaudeDesktopConfig
    if ($Remove) {
        if (Remove-JsonServerEntry $path) {
            Add-Result 'claude-desktop' (Get-DoneStatus 'removed') $path
        } else {
            Add-Result 'claude-desktop' 'skipped' "No '$ServerName' entry in $path"
        }
        return
    }
    Set-JsonServerEntry $path $Exe
    Add-Result 'claude-desktop' (Get-DoneStatus 'configured') "$path; fully quit Claude Desktop (tray icon too) and reopen."
}

function Set-CodexClient([string]$Exe) {
    $path = $script:ClientPaths.CodexConfig
    if ($Remove) {
        if (Test-Path -LiteralPath $path) {
            if (Set-CodexServerEntry $path $Exe $true) {
                Add-Result 'codex' (Get-DoneStatus 'removed') $path
            } else {
                Add-Result 'codex' 'skipped' "No '$ServerName' entry in $path"
            }
        } else {
            Add-Result 'codex' 'skipped' "$path does not exist."
        }
        return
    }
    Set-CodexServerEntry $path $Exe $false | Out-Null
    Add-Result 'codex' (Get-DoneStatus 'configured') "$path; restart Codex, then verify with /mcp or 'codex mcp list'."
}

function Set-AntigravityClient([string]$Exe) {
    $targets = @($script:ClientPaths.AntigravityConfig)
    if (Test-Path -LiteralPath $script:ClientPaths.AntigravitySharedConfig) {
        $targets += $script:ClientPaths.AntigravitySharedConfig
    }
    if ($Remove) {
        $removedFrom = @()
        foreach ($target in $targets) {
            if (Remove-JsonServerEntry $target) { $removedFrom += $target }
        }
        if ($removedFrom.Count -gt 0) {
            Add-Result 'antigravity' (Get-DoneStatus 'removed') ($removedFrom -join '; ')
        } else {
            Add-Result 'antigravity' 'skipped' "No '$ServerName' entry found."
        }
        return
    }
    foreach ($target in $targets) {
        Set-JsonServerEntry $target $Exe
    }
    Add-Result 'antigravity' (Get-DoneStatus 'configured') "$($targets -join '; '); refresh the MCP servers panel in Antigravity."
}

function Show-Summary {
    Write-Host ''
    Write-Step 'Summary'
    foreach ($result in $script:Results) {
        $color = switch ($result.Status) {
            'configured'   { 'Green' }
            'removed'      { 'Green' }
            'dry-run'      { 'Yellow' }
            'manual'       { 'Yellow' }
            'failed'       { 'Red' }
            default        { 'DarkGray' }
        }
        Write-Host ("  {0,-15} {1,-13} {2}" -f $result.Client, $result.Status, $result.Detail) -ForegroundColor $color
    }
}

# --- Main -----------------------------------------------------------------

try {
    $mode = 'Registering EULE Outlook MCP with AI clients'
    if ($Remove) { $mode = 'Removing EULE Outlook MCP from AI clients' }
    if ($DryRun) { $mode = "$mode (dry run)" }
    Write-Step $mode
    Write-Note "Server name: $ServerName"

    $targetClients = @(Resolve-TargetClients)
    if ($targetClients.Count -eq 0) {
        Write-Warning 'No supported AI clients were detected. Pass -Clients (e.g. -Clients codex) to force setup.'
    }

    $exe = ''
    if (-not $Remove -and $targetClients.Count -gt 0) {
        $exe = [string](Resolve-ServerExecutable | Select-Object -Last 1)
        if (-not $DryRun -and -not (Test-Path -LiteralPath $exe -PathType Leaf)) {
            throw "Resolved executable path is not a file: $exe"
        }
        Write-Note "Server executable: $exe"
        if (-not $DryRun) {
            $versionCheck = Invoke-NativeCommand $exe @('--version')
            if ($versionCheck.ExitCode -ne 0) {
                Write-Warning "The executable did not respond to --version; clients may fail to start it. ($($versionCheck.Output))"
            }
        }
    }

    foreach ($client in $targetClients) {
        Write-Host ''
        Write-Step "[$client]"
        try {
            switch ($client) {
                'claude-code'    { Set-ClaudeCodeClient $exe }
                'claude-desktop' { Set-ClaudeDesktopClient $exe }
                'codex'          { Set-CodexClient $exe }
                'antigravity'    { Set-AntigravityClient $exe }
            }
        } catch {
            Add-Result $client 'failed' $_.Exception.Message
        }
    }

    Show-Summary

    if (-not $Remove -and -not $DryRun -and $exe) {
        Write-Host ''
        Write-Note "Verify Outlook connectivity with: & `"$exe`" --diagnose"
        Write-Note 'Restart each client before use; MCP configuration is read at startup.'
    }

    $failures = @($script:Results | Where-Object { $_.Status -eq 'failed' })
    if ($Pause) { Read-Host 'Press Enter to close' | Out-Null }
    if ($failures.Count -gt 0) { exit 1 }
    exit 0
} catch {
    Write-Host ''
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    if ($Pause) { Read-Host 'Press Enter to close' | Out-Null }
    exit 1
}
