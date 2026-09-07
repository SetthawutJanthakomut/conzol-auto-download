# ConZoL Auto Download - daily opener
# Opens ConZoL in a named Chrome profile so the userscript can run its daily check.
# Change $ProfileName below if you ever want a different Chrome profile.

$ProfileName = 'setthawut'
# Chrome keeps profiles in folders called Default, Profile 1, Profile 2 ...
# Set this to pin one and skip the name lookup entirely. Leave it empty to
# look the folder up from the name above.
$ProfileDir  = 'Profile 1'
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

    # ---- 2) work out which profile folder to open ----------------------------
    # $ProfileDir wins if it is set. Otherwise the name is looked up in Local
    # State, which Chrome rewrites while it runs, so a read can catch it empty -
    # hence the retries and the folder remembered next to this script.
    #
    # The lookup checks the profile label first (name, shortcut_name) and only
    # then the Google account fields, because a second profile signed in to the
    # same Google account carries the same gaia name and would match first.
    $cacheFile = Join-Path $PSScriptRoot 'ConZoL-Daily.profile'
    $state     = Join-Path $env:LOCALAPPDATA 'Google\Chrome\User Data\Local State'
    $want      = $ProfileName.ToLower()
    $dir       = $null

    if ($ProfileDir) {
        $dir = $ProfileDir
        Write-Log ("using the folder set in the script: '" + $dir + "'")
    }

    for ($try = 1; $try -le 3 -and -not $dir; $try++) {
        if (-not (Test-Path $state)) { break }
        try {
            $json  = Get-Content -Path $state -Raw -ErrorAction Stop | ConvertFrom-Json
            $cache = $json.profile.info_cache
        } catch {
            $cache = $null
        }
        if ($cache) {
            foreach ($fields in @(@('name', 'shortcut_name'), @('gaia_name', 'gaia_given_name'))) {
                $rows = @()
                foreach ($p in $cache.PSObject.Properties) {
                    foreach ($f in $fields) {
                        $v = $p.Value.$f
                        if ($v) { $rows += [pscustomobject]@{ Dir = $p.Name; Text = $v.ToLower() } }
                    }
                }
                $hit = $rows | Where-Object { $_.Text -eq $want } | Select-Object -First 1
                if (-not $hit) { $hit = $rows | Where-Object { $_.Text.StartsWith($want) } | Select-Object -First 1 }
                if ($hit) { $dir = $hit.Dir; break }
            }
        }
        if (-not $dir -and $try -lt 3) { Start-Sleep -Milliseconds 700 }
    }

    if ($dir) {
        Set-Content -Path $cacheFile -Value $dir
    } elseif (Test-Path $cacheFile) {
        $dir = (Get-Content -Path $cacheFile -Raw).Trim()
        Write-Log ("Local State unreadable - using the remembered folder '" + $dir + "'")
    } else {
        $dir = 'Default'
        Write-Log ("profile '" + $ProfileName + "' not found and nothing remembered - using Default")
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
