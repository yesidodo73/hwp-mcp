$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$brokerScript = Join-Path $repoRoot 'hwp_mcp_broker.py'

$existing = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
    Where-Object {
        $_.Name -match '^python(\.exe)?$' -and
        $_.CommandLine -match 'hwp_mcp_broker\.py'
    } |
    Select-Object -First 1

if ($existing) {
    Write-Output "HWP broker already running (PID: $($existing.ProcessId))"
    exit 0
}

$python = (Get-Command python -ErrorAction SilentlyContinue).Source
if (-not $python) {
    $python = (Get-Command py -ErrorAction SilentlyContinue).Source
}

if (-not $python) {
    throw 'Python executable was not found.'
}

if ($python -like '*\\py.exe') {
    Start-Process -FilePath $python -ArgumentList @('-3', '-X', 'utf8', $brokerScript) -WorkingDirectory $repoRoot -WindowStyle Hidden
} else {
    Start-Process -FilePath $python -ArgumentList @('-X', 'utf8', $brokerScript) -WorkingDirectory $repoRoot -WindowStyle Hidden
}

Start-Sleep -Seconds 2
$started = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
    Where-Object {
        $_.Name -match '^python(\.exe)?$' -and
        $_.CommandLine -match 'hwp_mcp_broker\.py'
    } |
    Select-Object -First 1

if ($started) {
    Write-Output "Started HWP broker (PID: $($started.ProcessId))"
} else {
    throw 'Failed to start HWP broker.'
}
