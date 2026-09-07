# ConZoL Auto Download - daily opener
# Opens ConZoL in a named Chrome profile so the userscript can run its daily check.
# Change $ProfileName below if you ever want a different Chrome profile.

$ProfileName = 'setthawut'
$Url         = 'https://edms.gulf.co.th/dms/drawing.asp'
$LogFile     = Join-Path $PSScriptRoot 'ConZoL-Daily.log'

function Write-Log([string]$Message) {
    $line = '{0}  {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Add-Content -Path $LogFile -Value $line
}

try {
    # ---- 1) locate chrome.exe -------------------------------------------------
    $chrome = $null
    $roots = @($env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:LOCALAPPDATA) | Where-Object { $_ }
    foreach ($r in $roots) {
        $p = Join-Path $r 'Google\Chrome\Application\chrome.exe'
        if (Test-Path $p) { $chrome = $p; break }
    }
    if (-not $chrome) {
        $key = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe'
        if (Test-Path $key) { $chrome = (Get-ItemProperty -Path $key).'(default)' }
    }
    if (-not $chrome) { throw 'chrome.exe not found' }

    # ---- 2) turn the profile name shown in Chrome into its folder name --------
    # Chrome stores them as Default, Profile 1, Profile 2 ... and keeps the
    # display names in Local State, so the folder is looked up rather than guessed.
    $dir   = 'Default'
    $state = Join-Path $env:LOCALAPPDATA 'Google\Chrome\User Data\Local State'
    if (Test-Path $state) {
        $json  = Get-Content -Path $state -Raw | ConvertFrom-Json
        $cache = $json.profile.info_cache
        $want  = $ProfileName.ToLower()
        $found = $null
        foreach ($p in $cache.PSObject.Properties) {
            $n = $p.Value.name
            if ($n -and $n.ToLower() -eq $want) { $found = $p.Name; break }
        }
        if (-not $found) {
            foreach ($p in $cache.PSObject.Properties) {
                $n = $p.Value.name
                if ($n -and $n.ToLower().StartsWith($want)) { $found = $p.Name; break }
            }
        }
        if ($found) {
            $dir = $found
        } else {
            Write-Log ("profile '" + $ProfileName + "' not found in Local State - falling back to Default")
        }
    } else {
        Write-Log 'Local State not found - falling back to Default'
    }

    # ---- 3) open ConZoL in that profile --------------------------------------
    # Chrome adds a tab to the running window when that profile is already open.
    $arguments = '--profile-directory="' + $dir + '" ' + $Url
    Start-Process -FilePath $chrome -ArgumentList $arguments
    Write-Log ('opened ' + $Url + " in profile '" + $ProfileName + "' (" + $dir + ')')
}
catch {
    Write-Log ('ERROR: ' + $_.Exception.Message)
    exit 1
}
