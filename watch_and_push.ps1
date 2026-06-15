# watch_and_push.ps1
# Watches for changes to Excel files in the cps_dashboard folder and auto-pushes to GitHub.

$FolderPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogFile    = Join-Path $FolderPath "watch_and_push.log"

function Write-Log {
    param([string]$Message)
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] $Message"
    Write-Host $line
    Add-Content -Path $LogFile -Value $line
}

Write-Log "Watcher started. Monitoring: $FolderPath"
Write-Log "Watching for changes to *.xlsx files..."

# Set up the FileSystemWatcher
$Watcher = New-Object System.IO.FileSystemWatcher
$Watcher.Path   = $FolderPath
$Watcher.Filter = "*.xlsx"
$Watcher.NotifyFilter = [System.IO.NotifyFilters]::LastWrite `
                      -bor [System.IO.NotifyFilters]::FileName

# Debounce: track last push time to avoid duplicate commits on rapid saves
$script:LastPushed = [datetime]::MinValue
$script:PendingPush = $false

function Invoke-GitPush {
    param([string]$ChangedFile)

    # Debounce — ignore if we pushed within the last 10 seconds
    $now = Get-Date
    if (($now - $script:LastPushed).TotalSeconds -lt 10) {
        Write-Log "Debounce: skipping push (too soon after last push)"
        return
    }
    $script:LastPushed = $now

    Write-Log "Change detected: $ChangedFile — waiting 3s for save to complete..."
    Start-Sleep -Seconds 3

    Push-Location $FolderPath
    try {
        # Pull first to avoid conflicts
        $pullOut = & git pull --rebase 2>&1
        Write-Log "git pull: $pullOut"

        $addOut = & git add "*.xlsx" 2>&1
        Write-Log "git add: $addOut"

        $statusOut = & git status --short 2>&1
        if (-not $statusOut) {
            Write-Log "No staged changes — nothing to commit."
            return
        }

        $ts      = Get-Date -Format "yyyy-MM-dd HH:mm"
        $commitMsg = "Auto-push Excel update [$ts]"
        $commitOut = & git commit -m $commitMsg 2>&1
        Write-Log "git commit: $commitOut"

        $pushOut = & git push 2>&1
        Write-Log "git push: $pushOut"
        Write-Log "Done. GitHub Actions will now regenerate the dashboard."
    }
    catch {
        Write-Log "ERROR: $_"
    }
    finally {
        Pop-Location
    }
}

# Register events
$OnChanged = Register-ObjectEvent -InputObject $Watcher -EventName Changed -Action {
    Invoke-GitPush -ChangedFile $Event.SourceEventArgs.FullPath
}
$OnCreated = Register-ObjectEvent -InputObject $Watcher -EventName Created -Action {
    Invoke-GitPush -ChangedFile $Event.SourceEventArgs.FullPath
}
$OnRenamed = Register-ObjectEvent -InputObject $Watcher -EventName Renamed -Action {
    Invoke-GitPush -ChangedFile $Event.SourceEventArgs.FullPath
}

$Watcher.EnableRaisingEvents = $true

Write-Log "Watching... Press Ctrl+C to stop."

try {
    while ($true) { Start-Sleep -Seconds 5 }
}
finally {
    $Watcher.EnableRaisingEvents = $false
    Unregister-Event -SourceIdentifier $OnChanged.Name
    Unregister-Event -SourceIdentifier $OnCreated.Name
    Unregister-Event -SourceIdentifier $OnRenamed.Name
    $Watcher.Dispose()
    Write-Log "Watcher stopped."
}
