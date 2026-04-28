param(
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

$BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProxyPath = Join-Path $BaseDir 'naver_category_proxy.py'
$RequiredHelperVersion = 4

function Write-Step([string]$Message) {
    Write-Host "[MarketPlus] $Message"
}

function Test-HelperServer {
    try {
        $response = Invoke-RestMethod -Uri 'http://127.0.0.1:5555/api/map/status' -TimeoutSec 1
        if ($null -eq $response -or -not ($response.PSObject.Properties.Name -contains 'uploaded')) {
            return $false
        }
        if (-not ($response.PSObject.Properties.Name -contains 'helperVersion')) {
            return $false
        }
        return ([int]$response.helperVersion -ge $RequiredHelperVersion)
    } catch {
        return $false
    }
}

function Stop-OldHelperServer {
    $listeners = Get-NetTCPConnection -LocalPort 5555 -State Listen -ErrorAction SilentlyContinue
    foreach ($listener in $listeners) {
        $proc = Get-CimInstance Win32_Process -Filter "ProcessId = $($listener.OwningProcess)" -ErrorAction SilentlyContinue
        if ($proc -and $proc.CommandLine -like "*naver_category_proxy.py*") {
            Write-Step "Restarting old helper server process $($listener.OwningProcess)."
            Stop-Process -Id $listener.OwningProcess -Force -ErrorAction SilentlyContinue
        }
    }
}

function Start-HelperServer {
    if (-not (Test-Path -LiteralPath $ProxyPath)) {
        throw "Cannot find naver_category_proxy.py: $ProxyPath"
    }

    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($py) {
        Start-Process -FilePath $py.Source -ArgumentList @('-3', "`"$ProxyPath`"") -WorkingDirectory $BaseDir -WindowStyle Minimized
        return
    }

    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python) {
        Start-Process -FilePath $python.Source -ArgumentList @("`"$ProxyPath`"") -WorkingDirectory $BaseDir -WindowStyle Minimized
        return
    }

    throw 'Cannot find Python. The py -3 or python command is required.'
}

try {
    if (-not $Force -and (Test-HelperServer)) {
        Write-Step 'MarketPlus helper server is already running.'
    } else {
        Stop-OldHelperServer
        Write-Step 'Starting MarketPlus helper server.'
        Start-HelperServer

        $ready = $false
        for ($i = 0; $i -lt 20; $i++) {
            Start-Sleep -Milliseconds 300
            if (Test-HelperServer) {
                $ready = $true
                break
            }
        }

        if ($ready) {
            Write-Step 'MarketPlus helper server is ready: http://localhost:5555'
        } else {
            Write-Warning 'Server startup check is taking longer than expected. If category-map lookup fails, check the naver_category_proxy.py window.'
        }
    }

    Write-Step 'Now select products in Cafe24 MarketPlus and open the bulk register/send screen.'
    Write-Step 'When the registerall window opens, Tampermonkey will show the Category Helper panel.'
    Start-Sleep -Seconds 2
} catch {
    Write-Host ''
    Write-Host '[MarketPlus] Failed:' -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ''
    Read-Host 'Press Enter to close'
    exit 1
}
