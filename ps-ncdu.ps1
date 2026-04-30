# ============================================================
#  ____  ____      _   _  ____ ____  _   _
# |  _ \/ ___|    | \ | |/ ___|  _ \| | | |
# | |_) \___ \    |  \| | |   | | | | | | |
# |  __/ ___) |   | |\  | |___| |_| | |_| |
# |_|   |____/    |_| \_|\____|____/ \___/
#
# ============================================================
#  Script      : PS-NCDU HTML Edition
#  Description : Disk Usage Analyzer - Rapport HTML interactif
#  Version     : 3.6
#  Date        : 2026-05-01
#  Auteur      : Eric Guiffaut (EGUI@NOVONORDISK.COM)
#  Societe     : Novo Nordisk
# ------------------------------------------------------------
#  Changements v3.6 :
#  - Detection automatique des dossiers OneDrive
#  - Signalement fichiers cloud-only (taille reelle vs locale)
#  - Option scan OneDrive dans le menu de demarrage
#  - Avertissement HTML si fichiers cloud-only detectes
#  - Fix validation chemin UNC (v3.5)
#  - JS hors here-string (fix navigation)
# ============================================================

$SCRIPT_VERSION  = "3.6"
$SCRIPT_DATE     = "2026-05-01"
$USER_EMAIL      = "EGUI@NOVONORDISK.COM"
$SCRIPT_AUTHOR   = "Eric Guiffaut"
$DEFAULT_PATH    = "C:\Users"
$DEFAULT_DEPTH   = 3

$NN_LOGO_SVG = @'
<svg version="1.1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"
     viewBox="0 0 600 340" preserveAspectRatio="xMidYMid meet"
     style="height:26px;width:auto;display:block">
<style>.st0{fill:currentColor}</style>
<g><g>
<path class="st0" d="M148.856415,134.270599c4.715683,6.673569,14.127213-4.425323,13.298813-6.485603C161.329163,125.731544,144.183945,127.611404,148.856415,134.270599z"/>
<path class="st0" d="M186.063507,91.384827c24.809967,0,44.919418-20.038872,44.919418-44.700043C230.982925,22.008509,210.873474,2,186.063507,2s-44.926422,20.008509-44.926422,44.684784C141.137085,71.345955,161.253555,91.384827,186.063507,91.384827z M187.289169,12.43213c16.446503,0,29.773361,13.270297,29.773361,29.627888c0,16.364708-13.326859,29.620335-29.773361,29.620335c-16.453506,0-29.818924-13.255627-29.818924-29.620335C157.470245,25.702427,170.835663,12.43213,187.289169,12.43213z"/>
<path class="st0" d="M411.662659,170.312729c-6.047607-28.278595-30.629578-29.15416-43.750763-28.154999c-13.50061,1.02803-39.134796,5.037415-65.386902,5.037415c-35.616913,0-75.451599-9.959641-106.089844-24.113098c-5.303085-2.449768-2.665375-3.471626-0.252365-4.774796c11.484314-6.202248,21.484543-11.961472,27.557968-21.791771c4.649231-7.526573,2.873337-9.947327-2.263077-5.723434c-13.894852,11.426331-33.64212,17.818169-55.751816,9.993805c-22.109894-7.824417-30.670654-29.012489-32.635925-34.355553c-1.94194-5.335983-4.708725-5.729958-4.708725,2.092819c0,19.729858,8.786514,31.683937,11.714584,35.095177c2.939636,3.394241,4.298767,6.753876,2.673233,8.715385c-4.315704,5.207855-9.067581,10.7742-9.247055,13.010391c-0.188141,2.342445,0.034073,3.633438-0.923981,6.401962c-0.958054,2.768555-4.53688,6.701492-9.964096,13.508102c-3.094116,3.902222-1.203613,8.011185,1.404358,10.891022c3.146576,3.441284,5.150253,7.409286,8.98642,9.195786c3.835953,1.786499,7.424271,0.465973,11.894669,1.286789c4.398926,0.807678,9.92981,3.372528,14.272766,11.106094c6.279114,11.218048,13.147018,28.702728,25.377914,43.257706c5.463348,6.466705,5.440002,18.139679,5.463348,22.078415c0.3517,33.957733-17.603165,66.089569-29.012711,86.455414c-2.704819,4.837311-2.155685,8.823761,1.487381,8.854553c4.152542,0.044739,19.717743,0,22.216873,0c2.987961,0,4.628784-3.118835,4.84584-6.305603c1.889893-27.745422,17.708496-78.778809,29.673996-87.496811c20.052719,12.408707,39.953033,39.11058,30.210953,85.675278c-0.525406,2.510742-2.873764,8.141022,3.102371,8.141022c0,0,16.905945,0,19.269806,0c2.444183,0,6.443298-0.956451,6.196899-5.421539c-1.005341-18.226379-16.292145-56.61972-20.337311-79.276566c-2.556534-14.319214,8.924072-39.798752,52.129822-20.235687c16.70871-14.770813,38.423309-5.782944,29.146942,22.721359c-9.182831,28.217743-17.153564,40.83905-48.641846,74.125549c-3.609589,3.815796-3.797119,8.110352,1.876465,8.110352c2.827698,0,17.31073-0.023468,20.225159-0.023468c4.338013,0,6.080658-1.207153,7.42099-4.692444c1.340546-3.485291,15.659576-48.820862,43.966797-69.997437c2.188721-1.637268,4.071991-8.909637,8.459167-25.332977c13.770935,17.790436,15.4487,50.36084-7.886688,92.240692c-2.438416,4.357056-0.914886,7.76828,1.867157,7.76828c1.638153,0,14.379547,0,18.328888,0c3.415253,0,4.400177-1.842682,5.732025-6.923828c1.630096-6.219757,2.233551-9.26886,4.398956-15.334534c1.342804-3.761566,1.963409-4.367188,16.461914-4.214386c4.634766,0.048859,4.417908-2.418274,4.417908-6.282318C415.591125,248.933304,415.94809,190.351379,411.662659,170.312729z M360.095215,151.569916c4.747131,8.644669,5.004242,13.705551,3.644714,18.035294c-1.359741,4.329666-7.410858,6.76593-8.737152,6.305206c-0.160065-9.27565-2.674469-16.315247-7.060181-22.677094C351.987549,152.624542,356.037262,152.021973,360.095215,151.569916z M322.527313,156.060486c2.742218,10.216675,2.401031,23.64502-1.468597,27.554642c-6.105011,6.17009-33.759003,4.923004-41.720428,3.726898c-2.050568-0.3078-4.411957-0.69957-5.438995-6.257812c-1.224457-6.613266-1.50946-18.66687-1.890503-26.691498C288.742889,156.57309,305.686005,157.131958,322.527313,156.060486z M224.150482,175.486908c2.258331-7.456421,4.578201-19.956223,5.782837-30.153183c5.358231,1.625198,10.766861,3.083664,16.218643,4.35936c-3.668472,23.286591-10.998154,29.838699-14.868591,31.38353C228.304688,182.265472,221.892166,182.943329,224.150482,175.486908z M174.684402,187.742676c-4.718002-6.173126-10.846771-21.083054-16.158325-25.253998c-10.198914-8.015503-19.011017-1.736069-23.015289-6.500137c-4.262207-5.101501-6.005463-5.482849-2.022476-10.342636c0,0,4.484848-5.958939,6.360901-8.476593s1.523514-7.380692,2.277115-9.478897c0.745544-2.089783,5.41275-8.221367,9.982468-12.182281c5.049896-4.376823,9.20224-2.453056,25.090851,7.091614c19.301178,11.594513,25.200714,9.085114,18.589096,19.49025C191.61058,148.665222,178.007141,166.213058,174.684402,187.742676z M387.063751,180.647156c-1.06958,17.411148-4.950348,42.018463-10.972137,44.667999c-6.021973,2.649521-15.38797-13.339767-64.320618-14.213028c-34.750946-0.620132-61.450592,3.943466-66.355087,35.476791c-0.357071,2.295654-1.691833,1.859665-2.675293,0.798981c-9.539902-10.289474-18.225403-14.871811-32.290421-23.018311c-14.065231-8.1465-23.464691-19.293625-25.212906-30.193695c-1.748215-10.900116,2.197815-25.56279,22.084702-52.316223c2.76059-3.71376,3.259552-3.373184,14.356003,0.811111c-0.850037,17.873428-5.270859,27.924896-7.12854,37.157394c-1.884933,9.368073,8.593613,11.392853,17.798126,9.234299c9.204498-2.158554,17.959824-11.798889,21.818283-37.617004c3.61145,0.732544,6.478394,1.250443,10.118164,1.830048c0.526611,26.772079,2.100952,32.888885,3.199249,35.394989c2.160431,4.929214,7.15332,6.556671,11.208984,6.875824c21.187775,1.667358,41.672974,0.975189,49.007416-6.14267s4.803101-24.510056,2.789917-33.963684c3.31366-0.30986,6.144653-0.616028,9.444275-1.054169c6.887329,8.067734,7.844757,16.729187,7.844757,22.389984c0,6.143997,6.824951,9.323547,16.41626,2.76532c9.591339-6.558212,8.518646-18.095154,3.388214-28.574188C389.302673,149.000137,388.133545,163.235962,387.063751,180.647156z M403.184052,266.075226c-0.332397,2.54129-2.808838,4.614594-3.405396-0.651794c-0.766815-6.769653-4.288574-22.190384-5.427948-28.6707c-1.067535-6.072327,2.2005-24.604492,3.39212-39.98732c0.102661-1.325974,1.616394-1.282898,2.536926-0.31839C406.869293,203.351227,404.554871,255.595245,403.184052,266.075226z"/>
</g></g>
</svg>
'@

$EXCLUDED_DIRS = @(
    "C:\Windows\WinSxS",
    "C:\Windows\Installer",
    "C:\`$Recycle.Bin",
    "C:\System Volume Information",
    "C:\Recovery",
    "C:\ProgramData\Microsoft\Windows Defender",
    "C:\Windows\SoftwareDistribution"
)

$SCAN_MODES = @{
    1 = @{ Name="Equilibre     - Ignore sous-dossiers < 100 MB [RECOMMANDE]"; Speed="~25-40s"; Accuracy="Tres bonne (>95%)" }
    2 = @{ Name="Precis        - Ignore sous-dossiers < 10 MB";               Speed="~40-80s"; Accuracy="Excellente (>99%)" }
    3 = @{ Name="Complet       - Scan integral de tous les fichiers";          Speed="~80-150s";Accuracy="100% exacte"       }
}

# Attributs fichiers OneDrive cloud-only
$ONEDRIVE_CLOUD_ATTR = [long]0x00400000  # FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS
$ONEDRIVE_PINNED_ATTR = [long]0x00080000 # FILE_ATTRIBUTE_PINNED

$scriptTemp = "$env:TEMP\psncdu"
if (-not (Test-Path $scriptTemp)) { New-Item -ItemType Directory -Path $scriptTemp -Force | Out-Null }
$logFile  = "$scriptTemp\psncdu_debug.log"
$htmlFile = "$scriptTemp\psncdu_report.html"
"============================================================" | Set-Content -Path $logFile -ErrorAction SilentlyContinue

function Write-Log {
    param([string]$Message,[ValidateSet("INFO","WARN","ERROR","DEBUG")][string]$Level="INFO")
    try {
        $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
        Add-Content -Path $logFile -Value "[$ts] [$Level] $Message" -ErrorAction SilentlyContinue
    } catch {}
}
function Write-Step {
    param([string]$Step,[datetime]$StepStart,[string]$Extra="")
    $elapsed=[int]((Get-Date)-$StepStart).TotalSeconds
    $msg="$Step (${elapsed}s)"; if($Extra){$msg+=" - $Extra"}
    Write-Log $msg; Write-Host "  $msg" -ForegroundColor DarkGray
}

$script:LastProgressUpdate = [datetime]::MinValue
$script:ScanGlobalStart    = $null

function Update-Progress {
    param([string]$Activity,[string]$Status,[string]$Operation="",[int]$Pct,[bool]$LogIt=$false,[bool]$Force=$false)
    $now = Get-Date
    if (-not $Force -and ($now - $script:LastProgressUpdate).TotalMilliseconds -lt 500) { return }
    $script:LastProgressUpdate = $now
    $elapsed=""
    if ($null -ne $script:ScanGlobalStart) {
        $secs=[int]($now-$script:ScanGlobalStart).TotalSeconds
        if($secs -lt 60){$elapsed=" [${secs}s]"}else{$mins=[int]($secs/60);$secs=$secs%60;$elapsed=" [${mins}m${secs}s]"}
    }
    Write-Progress -Activity "$Activity$elapsed" -Status $Status -CurrentOperation $Operation -PercentComplete $Pct
    if ($LogIt) { Write-Log "[PROGRESS $Pct%$elapsed] $Status$(if($Operation){" | $Operation"})" }
}

function Get-FullUserName {
    try {
        $wmi = Get-WmiObject -Class Win32_UserAccount -Filter "Name='$($env:USERNAME)'" -ErrorAction SilentlyContinue
        if ($wmi -and -not [string]::IsNullOrWhiteSpace($wmi.FullName)) { return $wmi.FullName }
    } catch {}
    try {
        $adsi = [ADSI]"WinNT://$($env:USERDOMAIN)/$($env:USERNAME),user"
        $fn   = $adsi.FullName
        if (-not [string]::IsNullOrWhiteSpace($fn)) { return $fn.ToString() }
    } catch {}
    return $env:USERNAME
}

# ============================================================
# ✅ v3.6 - Detection automatique des dossiers OneDrive
# ============================================================
function Get-OneDrivePaths {
    $paths = @()

    # 1. Variable d environnement standard
    $envOD = $env:OneDrive
    if (-not [string]::IsNullOrWhiteSpace($envOD) -and (Test-Path $envOD -ErrorAction SilentlyContinue)) {
        $paths += @{ Path=$envOD; Type="OneDrive Personnel"; Icon="☁️" }
        Write-Log "[OD] Detecte via env:OneDrive : '$envOD'"
    }

    # 2. Variable OneDriveCommercial (entreprise)
    $envODC = $env:OneDriveCommercial
    if (-not [string]::IsNullOrWhiteSpace($envODC) -and (Test-Path $envODC -ErrorAction SilentlyContinue)) {
        if (-not ($paths | Where-Object { $_.Path -eq $envODC })) {
            $paths += @{ Path=$envODC; Type="OneDrive Entreprise"; Icon="🏢" }
            Write-Log "[OD] Detecte via env:OneDriveCommercial : '$envODC'"
        }
    }

    # 3. Scan du profil utilisateur pour dossiers OneDrive*
    $userProfile = $env:USERPROFILE
    if (-not [string]::IsNullOrWhiteSpace($userProfile)) {
        try {
            $odDirs = Get-ChildItem -Path $userProfile -Directory -ErrorAction SilentlyContinue |
                      Where-Object { $_.Name -like "OneDrive*" }
            foreach ($d in $odDirs) {
                if (-not ($paths | Where-Object { $_.Path -eq $d.FullName })) {
                    $type = if ($d.Name -match "Novo|Enterprise|Business|Corp") { "OneDrive Entreprise" } else { "OneDrive" }
                    $paths += @{ Path=$d.FullName; Type=$type; Icon="☁️" }
                    Write-Log "[OD] Detecte via scan profil : '$($d.FullName)'"
                }
            }
        } catch {}
    }

    # 4. Registre Windows - cles OneDrive
    $regPaths = @(
        "HKCU:\Software\Microsoft\OneDrive\Accounts",
        "HKCU:\Software\Microsoft\SkyDrive"
    )
    foreach ($regPath in $regPaths) {
        try {
            if (Test-Path $regPath -ErrorAction SilentlyContinue) {
                $accounts = Get-ChildItem $regPath -ErrorAction SilentlyContinue
                foreach ($acc in $accounts) {
                    $localPath = (Get-ItemProperty -Path $acc.PSPath -ErrorAction SilentlyContinue).UserFolder
                    if (-not [string]::IsNullOrWhiteSpace($localPath) -and
                        (Test-Path $localPath -ErrorAction SilentlyContinue) -and
                        -not ($paths | Where-Object { $_.Path -eq $localPath })) {
                        $accName = Split-Path $acc.PSPath -Leaf
                        $type = if ($accName -match "Business|1|2") { "OneDrive Entreprise ($accName)" } else { "OneDrive Personnel" }
                        $paths += @{ Path=$localPath; Type=$type; Icon="☁️" }
                        Write-Log "[OD] Detecte via registre '$regPath\$accName' : '$localPath'"
                    }
                }
            }
        } catch {}
    }

    return $paths
}

# ============================================================
# ✅ v3.6 - Detection fichiers cloud-only OneDrive
# ============================================================
function Test-IsCloudOnly {
    param([object]$FileItem)
    try {
        $attr = [long]$FileItem.Attributes
        # FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS (0x400000) = cloud only
        # FILE_ATTRIBUTE_RECALL_ON_OPEN (0x40000) = aussi cloud
        if (($attr -band 0x400000) -ne 0) { return $true }
        if (($attr -band 0x40000)  -ne 0) { return $true }
        return $false
    } catch { return $false }
}

function Get-OneDriveCloudSize {
    param([object]$FileItem)
    # Pour les fichiers cloud-only, tenter de lire la taille reelle
    # via les streams alternatifs ou la taille sur disque
    try {
        # La propriete Length donne 0 pour cloud-only
        # On peut tenter de lire le flux de donnees alternatif OneDrive
        # qui stocke parfois la taille reelle
        $fi = New-Object System.IO.FileInfo($FileItem.FullName)
        # Si Length > 0 malgre l attribut cloud, c est la vraie taille
        if ($fi.Length -gt 0) { return $fi.Length }
    } catch {}
    return [long]0
}

Write-Log "PS-NCDU v$SCRIPT_VERSION ($SCRIPT_DATE) Demarrage"
Write-Log "PSVersion    : $($PSVersionTable.PSVersion)"
Write-Log "LanguageMode : $($ExecutionContext.SessionState.LanguageMode)"
Write-Log "Auteur       : $SCRIPT_AUTHOR ($USER_EMAIL)"
Write-Log "User         : $($env:USERNAME)"
Write-Log "Machine      : $($env:COMPUTERNAME)"

# ============================================================
# Utilitaires
# ============================================================
function Normalize-Path {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }
    $Path = $Path.Trim() -replace '/','\' 
    if ($Path -match '^\\\\') {
        $Path = '\\' + ($Path.Substring(2) -replace '\\{2,}','\')
        return $Path.TrimEnd('\')
    }
    $Path = $Path.TrimEnd('\')
    if ($Path -match '^[A-Za-z]:$') { $Path += '\' }
    return $Path
}
function Get-ParentPath {
    param([string]$Path)
    if ([string]::IsNullOrEmpty($Path)) { return $null }
    if ($Path -match '^\\\\') {
        $Path = $Path.TrimEnd('\')
        $parts = $Path.Substring(2) -split '\\'
        if ($parts.Count -le 2) { return $null }
        return '\\' + ($parts[0..($parts.Count-2)] -join '\')
    }
    $Path = $Path.TrimEnd('\'); $idx = $Path.LastIndexOf('\')
    if ($idx -le 0) { return $null }
    $parent = $Path.Substring(0,$idx)
    if ($parent -match '^[A-Za-z]:$') { $parent += '\' }
    if ($parent -eq $Path) { return $null }
    return $parent
}
function Get-RootDepth {
    param([string]$RootPath)
    if ($RootPath -match '^\\\\') { return 0 }
    $parts = ($RootPath.TrimEnd('\') -split '\\') | Where-Object { $_ -ne '' }
    return ($parts | Measure-Object).Count
}
function Get-PathDepth {
    param([string]$Path,[string]$RootPath,[int]$RootDepth)
    $pathClean=$Path.TrimEnd('\'); $rootClean=$RootPath.TrimEnd('\')
    if ($pathClean -match '^\\\\') {
        if ($pathClean.ToLower() -eq $rootClean.ToLower()) { return 0 }
        $relative=$pathClean.Substring($rootClean.Length).TrimStart('\')
        if ([string]::IsNullOrEmpty($relative)) { return 0 }
        return ($relative -split '\\' | Where-Object { $_ -ne '' } | Measure-Object).Count
    }
    $parts=($pathClean -split '\\') | Where-Object { $_ -ne '' }
    return ($parts | Measure-Object).Count - $RootDepth
}
function Test-IsExcluded {
    param([string]$Path,[string[]]$ExcludedList)
    $pn=$Path.TrimEnd('\').ToLower()
    foreach ($ex in $ExcludedList) {
        $en=$ex.TrimEnd('\').ToLower()
        if ($pn -eq $en -or $pn.StartsWith($en+'\')) { return $true }
    }
    return $false
}
function Test-IsJunction {
    param([string]$Path)
    try {
        $attr = (Get-Item -Path $Path -Force -ErrorAction Stop).Attributes
        return ($attr -band [System.IO.FileAttributes]::ReparsePoint) -ne 0
    } catch { return $false }
}
function Test-PathAccessible {
    param([string]$Path)
    if ($Path -notmatch '^\\\\') {
        return (Test-Path -Path $Path -ErrorAction SilentlyContinue)
    }
    try { if(Test-Path -Path $Path -ErrorAction Stop){ return $true } } catch {}
    try {
        $di=New-Object System.IO.DirectoryInfo($Path)
        $null=$di.GetDirectories()
        return $true
    } catch {}
    try {
        $prevEnc=[Console]::OutputEncoding
        [Console]::OutputEncoding=[System.Text.Encoding]::UTF8
        $cmdOut=& cmd /c "dir `"$Path`" 2>nul"
        [Console]::OutputEncoding=$prevEnc
        if($null -ne $cmdOut -and $cmdOut.Count -gt 0){ return $true }
    } catch {}
    return $false
}
function Format-Size {
    param([long]$Size)
    if ($Size -lt 0)       { return "?" }
    if ($Size -ge 1TB)     { return "{0:N2} TB" -f ($Size/1TB) }
    elseif ($Size -ge 1GB) { return "{0:N2} GB" -f ($Size/1GB) }
    elseif ($Size -ge 1MB) { return "{0:N2} MB" -f ($Size/1MB) }
    elseif ($Size -ge 1KB) { return "{0:N2} KB" -f ($Size/1KB) }
    else                   { return "{0} B"     -f $Size }
}
function ConvertTo-HtmlEncoded {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return "" }
    $Text=$Text -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;' -replace "'","&#39;"
    return $Text
}
function ConvertTo-JsonSafe {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return "" }
    $Text = $Text -replace '\\','\\\\' -replace '"','\"' -replace "`r",'' -replace "`n",'' -replace "`t",'\t'
    $Text = $Text -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', ''
    return $Text
}
function Invoke-CmdDirSafe {
    param([string]$Path, [string]$DirArgs = "/b /ad")
    $prevEnc = [Console]::OutputEncoding
    try {
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        $cmdOut = & cmd /c "chcp 65001 >nul 2>nul & dir $DirArgs `"$Path`" 2>nul"
        return $cmdOut
    }
    catch { return $null }
    finally { [Console]::OutputEncoding = $prevEnc }
}
function Get-SubDirectories {
    param([string]$Path)
    try {
        $items = Get-ChildItem -Path $Path -Directory -ErrorAction Stop -Force
        $jCount=0; $filtered=@()
        foreach ($item in $items) {
            if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { $jCount++ }
            else { $filtered += $item }
        }
        return @{ Items=$filtered; Method="GCI"; Denied=$false; Junctions=$jCount }
    }
    catch [System.UnauthorizedAccessException] {}
    catch { Write-Log "[DIR-M1] Erreur '$Path' : $_" -Level DEBUG }
    try {
        $di=New-Object System.IO.DirectoryInfo($Path); $subs=$di.GetDirectories(); $jCount=0; $items=@()
        foreach ($sub in $subs) {
            if (($sub.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { $jCount++ }
            else { $items += New-Object PSObject -Property @{ FullName=$sub.FullName; Name=$sub.Name } }
        }
        return @{ Items=$items; Method="NET"; Denied=$false; Junctions=$jCount }
    }
    catch [System.UnauthorizedAccessException] {}
    catch { Write-Log "[DIR-M2] Erreur '$Path' : $_" -Level DEBUG }
    try {
        $cmdOut=Invoke-CmdDirSafe -Path $Path -DirArgs "/b /ad"
        if ($null -ne $cmdOut) {
            $jCount=0; $items=@()
            foreach ($line in ($cmdOut | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
                $fp=Join-Path $Path $line
                try {
                    $attr=(Get-Item $fp -Force -ErrorAction Stop).Attributes
                    if (($attr -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { $jCount++; continue }
                } catch {}
                $items += New-Object PSObject -Property @{ FullName=$fp; Name=$line }
            }
            return @{ Items=$items; Method="CMD"; Denied=$false; Junctions=$jCount }
        }
    }
    catch { Write-Log "[DIR-M3] Erreur '$Path' : $_" -Level DEBUG }
    return @{ Items=@(); Method="NONE"; Denied=$true; Junctions=0 }
}
function Get-DirectFiles {
    param([string]$Path)
    try { return Get-ChildItem -Path $Path -File -ErrorAction Stop -Force }
    catch [System.UnauthorizedAccessException] {}
    catch {}
    try {
        $di=New-Object System.IO.DirectoryInfo($Path)
        return $di.GetFiles() | ForEach-Object {
            New-Object PSObject -Property @{ FullName=$_.FullName; Name=$_.Name; Length=$_.Length; Extension=$_.Extension; DirectoryName=$Path; Attributes=$_.Attributes }
        }
    }
    catch {}
    return @()
}
function Get-RecursiveFiles {
    param([string]$Path)
    try { return Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue -Force }
    catch [System.UnauthorizedAccessException] {}
    catch {}
    try {
        $result=@(); $dirQueue=@($Path)
        while ($dirQueue.Count -gt 0) {
            $current=$dirQueue[0]; $dirQueue=if($dirQueue.Count -gt 1){$dirQueue[1..($dirQueue.Count-1)]}else{@()}
            try {
                $di=New-Object System.IO.DirectoryInfo($current)
                foreach ($f in $di.GetFiles()) {
                    $result += New-Object PSObject -Property @{ FullName=$f.FullName; Name=$f.Name; Length=$f.Length; Extension=$f.Extension; DirectoryName=$current; Attributes=$f.Attributes }
                }
                foreach ($s in $di.GetDirectories()) {
                    if (($s.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -eq 0) { $dirQueue += $s.FullName }
                }
            } catch {}
        }
        if ($result.Count -gt 0) { return $result }
    } catch {}
    try {
        $cmdOut=Invoke-CmdDirSafe -Path $Path -DirArgs "/s /b /a-d"
        if ($null -ne $cmdOut) {
            return $cmdOut | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                   ForEach-Object {
                       try {
                           $fi=New-Object System.IO.FileInfo($_)
                           New-Object PSObject -Property @{ FullName=$_; Name=$fi.Name; Length=$fi.Length; Extension=$fi.Extension; DirectoryName=$fi.DirectoryName; Attributes=$fi.Attributes }
                       } catch { $null }
                   } | Where-Object { $null -ne $_ }
        }
    } catch {}
    return @()
}

# ============================================================
# SCAN v3.6 - avec comptage fichiers cloud-only OneDrive
# ============================================================
function Start-FastScan {
    param(
        [string]$RootPath,
        [int]$MaxDepth=3,
        [hashtable]$Mode=$null,
        [bool]$IsOneDrive=$false
    )
    $ACT="PS-NCDU v$SCRIPT_VERSION"; $scanStart=Get-Date; $script:ScanGlobalStart=$scanStart
    $rootDepth=Get-RootDepth -RootPath $RootPath
    $modeName=if($null -ne $Mode){$Mode["Name"]}else{"Complet"}
    $isUNC=$RootPath -match '^\\\\'; $pathType=if($isUNC){"UNC reseau"}else{"Local"}
    $unlimited=($MaxDepth -eq 0); $depthLabel=if($unlimited){"ILLIMITEE"}else{"$MaxDepth niveaux"}

    # Stats OneDrive
    $cloudOnlyCount  = 0
    $cloudOnlySize   = [long]0
    $localOnlySize   = [long]0

    Write-Log "=========================================="; Write-Log "DEBUT SCAN v$SCRIPT_VERSION : $RootPath"
    Write-Log "Type : $pathType | Mode : $modeName | Profondeur : $depthLabel | OneDrive : $IsOneDrive"
    Write-Log "=========================================="

    $t=Get-Date; Update-Progress $ACT "E1/5 : Liste des dossiers [$depthLabel]..." "" 1 $true $true
    $allLevels=@(@($RootPath)); $excludedFound=@(); $accessDenied=@(); $junctionsTotal=0
    $methodStats=@{GCI=0;NET=0;CMD=0;NONE=0}

    if ($unlimited) {
        Update-Progress $ACT "E1/5 : Enumeration illimitee..." "" 5 $true $true
        $flatList=@($RootPath)
        try {
            $allDirsFound=Get-ChildItem -Path $RootPath -Recurse -Directory -ErrorAction SilentlyContinue -Force
            $dirsDone=0
            foreach ($dir in $allDirsFound) {
                $dirsDone++
                if(($dirsDone%500)-eq 0){Update-Progress $ACT "E1/5 : $dirsDone dossiers..." "" 15 ($dirsDone%5000-eq 0)}
                if(($dir.Attributes -band [System.IO.FileAttributes]::ReparsePoint)-ne 0){$junctionsTotal++;continue}
                if(Test-IsExcluded -Path $dir.FullName -ExcludedList $EXCLUDED_DIRS){
                    if(-not($excludedFound -contains $dir.FullName)){$excludedFound+=$dir.FullName}
                } else {$flatList+=$dir.FullName}
            }
        } catch {Write-Log "[E1] Erreur recurse : $_" -Level WARN}
        $allLevels=@($flatList)
    } else {
        for ($d=0;$d -lt $MaxDepth;$d++) {
            $currentLevel=$allLevels[$d]; if($null -eq $currentLevel -or $currentLevel.Count -eq 0){break}
            $tLevel=Get-Date; $nextLevel=@(); $totalDirs=$currentLevel.Count; $dirsDone=0
            $levelNum=$d+1; $levelBase=1+[int](($d*24)/$MaxDepth); $levelRng=[int](24/$MaxDepth)
            foreach ($dir in $currentLevel) {
                $dirsDone++; $pct=$levelBase+[int](($dirsDone*$levelRng)/($totalDirs+1))
                Update-Progress $ACT "E1/5 : Niveau $levelNum/$MaxDepth | $dirsDone/$totalDirs" "Scan : $dir" $pct ($dirsDone%200-eq 0)
                $subResult=Get-SubDirectories -Path $dir
                $methodStats[$subResult["Method"]]++; $junctionsTotal+=$subResult["Junctions"]
                if($subResult["Denied"]){if(-not($accessDenied -contains $dir)){$accessDenied+=$dir}}
                else{foreach($sub in $subResult["Items"]){if(Test-IsExcluded -Path $sub.FullName -ExcludedList $EXCLUDED_DIRS){if(-not($excludedFound -contains $sub.FullName)){$excludedFound+=$sub.FullName}}else{$nextLevel+=$sub.FullName}}}
            }
            $allLevels+=,$nextLevel
            Write-Log "[E1] Niveau $levelNum/$MaxDepth : $($nextLevel.Count) en $([int]((Get-Date)-$tLevel).TotalSeconds)s"
            if($nextLevel.Count -eq 0){break}
        }
    }

    $dirsInScope=@(); foreach($level in $allLevels){foreach($dir in $level){$dirsInScope+=$dir}}; $nbScope=$dirsInScope.Count
    $dirsUnscanned=@()
    if(-not $unlimited){
        $lastLevel=$allLevels[$allLevels.Count-1]
        if($null -ne $lastLevel){
            foreach($dir in $lastLevel){
                $depth=Get-PathDepth -Path $dir -RootPath $RootPath -RootDepth $rootDepth
                if($depth -eq $MaxDepth){
                    $subResult=Get-SubDirectories -Path $dir
                    if(-not $subResult["Denied"]){foreach($sub in $subResult["Items"]){if(-not(Test-IsExcluded -Path $sub.FullName -ExcludedList $EXCLUDED_DIRS)){$dirsUnscanned+=$sub.FullName}}}
                    else{if(-not($accessDenied -contains $dir)){$accessDenied+=$dir}}
                }
            }
        }
    }
    Write-Step "[E1] TERMINEE" $t "Scope : $nbScope | NonScannes : $($dirsUnscanned.Count) | Proteges : $($accessDenied.Count) | Junctions : $junctionsTotal"

    $t=Get-Date; Update-Progress $ACT "E2/5 : Init ($nbScope dossiers)..." "" 25 $true $true
    $dirOwnSizes=@{}; $dirTotalSizes=@{}; $initDone=0
    foreach($dir in $dirsInScope){
        $initDone++; $dirOwnSizes[$dir]=[long]0; $dirTotalSizes[$dir]=[long]0
        Update-Progress $ACT "E2/5 : Init $initDone / $nbScope" "" (25+[int](($initDone*5)/($nbScope+1))) ($initDone%5000-eq 0)
    }
    Write-Step "[E2] TERMINEE" $t "$nbScope entrees"

    $t=Get-Date; $totalFilesScanned=0; $totalSizeScanned=[long]0
    Update-Progress $ACT "E3/5 : Scan tailles..." "" 30 $true $true
    try {
        $rootFiles=Get-DirectFiles -Path $RootPath
        foreach($f in $rootFiles){
            if($null -eq $f){continue}
            $fl=[long]$f.Length
            if($IsOneDrive -and (Test-IsCloudOnly $f)){
                $cloudOnlyCount++
                # Ne pas compter la taille locale (0) mais signaler
            } else {
                $dirOwnSizes[$RootPath]+=$fl
                $localOnlySize+=$fl
            }
        }
    } catch {}

    $level1Dirs=@(); $l1Result=Get-SubDirectories -Path $RootPath
    if(-not $l1Result["Denied"]){foreach($s in $l1Result["Items"]){$level1Dirs+=$s.FullName}}
    $totalL1=$level1Dirs.Count; $doneL1=0

    foreach($l1dir in $level1Dirs){
        $doneL1++; $l1Name=Split-Path $l1dir -Leaf
        $pctL1=30+[int](($doneL1*45)/($totalL1+1)); if($pctL1 -gt 74){$pctL1=74}
        if(Test-IsExcluded -Path $l1dir -ExcludedList $EXCLUDED_DIRS){Update-Progress $ACT "E3/5 : L1 $doneL1/$totalL1 - $l1Name [EXCLU]" "" $pctL1 $true $true;continue}
        if(Test-IsJunction -Path $l1dir){Write-Log "[E3] L1 JUNCTION : '$l1dir'";continue}
        Update-Progress $ACT "E3/5 : L1 $doneL1/$totalL1 - $l1Name | $(Format-Size $totalSizeScanned)" "Collecte..." $pctL1 $true $true
        try {
            $filesInL1=Get-RecursiveFiles -Path $l1dir; $nbF1=0; $sz1=[long]0; $fdL1=0
            foreach($file in $filesInL1){
                if($null -eq $file){continue}
                $pd=$null; try{$pd=$file.DirectoryName}catch{$pd=Get-ParentPath $file.FullName}
                if($null -eq $pd){continue}
                if(Test-IsExcluded -Path $pd -ExcludedList $EXCLUDED_DIRS){continue}

                # ✅ v3.6 : Gestion fichiers cloud-only OneDrive
                if($IsOneDrive -and (Test-IsCloudOnly $file)){
                    $cloudOnlyCount++
                    # Fichier cloud-only : on ne compte pas sa taille locale (0)
                    # mais on le comptabilise pour l avertissement
                    continue
                }

                $fl=[long]$file.Length; $sz1+=$fl; $nbF1++; $fdL1++; $totalFilesScanned++; $totalSizeScanned+=$fl; $localOnlySize+=$fl
                Update-Progress $ACT "E3/5 : L1 $doneL1/$totalL1 - $l1Name | $(Format-Size $totalSizeScanned)" "$fdL1 fichiers ($(Format-Size $sz1))" $pctL1 ($fdL1%50000-eq 0)
                if($dirOwnSizes.ContainsKey($pd)){$dirOwnSizes[$pd]+=$fl}
                else{
                    $anc=$pd; $hops=0; $found=$false
                    while(-not[string]::IsNullOrEmpty($anc) -and $hops -lt 50){
                        $anc=Get-ParentPath $anc; $hops++
                        if(-not[string]::IsNullOrEmpty($anc) -and $dirOwnSizes.ContainsKey($anc)){$dirOwnSizes[$anc]+=$fl;$found=$true;break}
                    }
                    if(-not $found){$dirOwnSizes[$RootPath]+=$fl}
                }
            }
            Write-Log "[E3] L1 $doneL1/$totalL1 '$l1Name' : $nbF1 = $(Format-Size $sz1)"
        } catch {Write-Log "[E3] Erreur '$l1dir' : $_" -Level ERROR}
    }
    if($IsOneDrive -and $cloudOnlyCount -gt 0){
        Write-Log "[OD] Fichiers cloud-only detectes : $cloudOnlyCount (non comptes dans les tailles locales)"
        Write-Host "  ☁️  OneDrive : $cloudOnlyCount fichiers cloud-only detectes (taille locale = 0)" -ForegroundColor Cyan
    }
    Write-Step "[E3] TERMINEE" $t "$totalFilesScanned fichiers = $(Format-Size $totalSizeScanned)"

    $t=Get-Date; Update-Progress $ACT "E4/5 : Propagation..." "" 75 $true $true
    foreach($dir in $dirsInScope){$dirTotalSizes[$dir]=$dirOwnSizes[$dir]}; $propDone=0
    foreach($dir in $dirsInScope){
        $propDone++; $ownSize=$dirOwnSizes[$dir]; if($ownSize -eq 0){continue}
        $pct=75+[int](($propDone*10)/($nbScope+1)); if($pct -gt 84){$pct=84}
        Update-Progress $ACT "E4/5 : $propDone / $nbScope" "Racine : $(Format-Size $dirTotalSizes[$RootPath])" $pct ($propDone%5000-eq 0)
        $current=Get-ParentPath $dir; $hops=0
        while(-not[string]::IsNullOrEmpty($current) -and $hops -lt 50){
            if($dirTotalSizes.ContainsKey($current)){$dirTotalSizes[$current]+=$ownSize}
            $current=Get-ParentPath $current; $hops++
        }
    }
    Write-Step "[E4] TERMINEE" $t "Racine : $(Format-Size $dirTotalSizes[$RootPath])"

    $t=Get-Date; Update-Progress $ACT "E5/5 : Index + rapport..." "" 85 $true $true
    $childIndex=@{}; $idxDone=0
    foreach($dir in $dirsInScope){
        $idxDone++; $parent=Get-ParentPath $dir
        if($null -ne $parent){if(-not $childIndex.ContainsKey($parent)){$childIndex[$parent]=@()};$childIndex[$parent]+=$dir}
        Update-Progress $ACT "E5/5 : Index $idxDone / $nbScope" "" (85+[int](($idxDone*3)/($nbScope+1))) ($idxDone%5000-eq 0)
    }
    foreach($u in $dirsUnscanned){$p=Get-ParentPath $u;if($null -ne $p){if(-not $childIndex.ContainsKey($p)){$childIndex[$p]=@()};if(-not($childIndex[$p] -contains "UNSCANNED:$u")){$childIndex[$p]+="UNSCANNED:$u"}}}
    foreach($d in $accessDenied){$p=Get-ParentPath $d;if($null -ne $p){if(-not $childIndex.ContainsKey($p)){$childIndex[$p]=@()};if(-not($childIndex[$p] -contains "DENIED:$d")){$childIndex[$p]+="DENIED:$d"}}}
    foreach($e in $excludedFound){$p=Get-ParentPath $e;if($null -ne $p){if(-not $childIndex.ContainsKey($p)){$childIndex[$p]=@()};if(-not($childIndex[$p] -contains "EXCLU:$e")){$childIndex[$p]+="EXCLU:$e"}}}

    Update-Progress $ACT "E5/5 : Construction ($nbScope dossiers)..." "" 90 $true $true
    $allScans=@{}; $reportCount=0
    foreach($dirPath in $dirsInScope){
        if(-not $unlimited){
            $dirDepth=Get-PathDepth -Path $dirPath -RootPath $RootPath -RootDepth $rootDepth
            if($dirDepth -gt $MaxDepth){continue}
        }
        $entries=@()
        if($childIndex.ContainsKey($dirPath)){
            foreach($ce in $childIndex[$dirPath]){
                if($ce -like "EXCLU:*"){$ep=$ce.Substring(6);$entries+=@{Name=Split-Path $ep -Leaf;FullPath=$ep;Size=[long]-1;IsDir=$true;Excluded=$true;Unscanned=$false;Denied=$false;Ext=""}}
                elseif($ce -like "UNSCANNED:*"){$up=$ce.Substring(10);$entries+=@{Name=Split-Path $up -Leaf;FullPath=$up;Size=[long]-2;IsDir=$true;Excluded=$false;Unscanned=$true;Denied=$false;Ext=""}}
                elseif($ce -like "DENIED:*"){$dp=$ce.Substring(7);$entries+=@{Name=Split-Path $dp -Leaf;FullPath=$dp;Size=[long]-3;IsDir=$true;Excluded=$false;Unscanned=$false;Denied=$true;Ext=""}}
                else{$cn=$ce.Substring($dirPath.TrimEnd('\').Length).TrimStart('\');$entries+=@{Name=$cn;FullPath=$ce;Size=$dirTotalSizes[$ce];IsDir=$true;Excluded=$false;Unscanned=$false;Denied=$false;Ext=""}}
            }
        }
        if($dirOwnSizes[$dirPath] -gt 0){
            try{$directFiles=Get-DirectFiles -Path $dirPath;foreach($f in $directFiles){if($null -eq $f){continue};if($IsOneDrive -and (Test-IsCloudOnly $f)){continue};$fl=[long]$f.Length;$entries+=@{Name=$f.Name;FullPath=$f.FullName;Size=$fl;IsDir=$false;Excluded=$false;Unscanned=$false;Denied=$false;Ext=$f.Extension.ToLower()}}}catch{}
        }
        $sN=$entries|Where-Object{-not $_["Excluded"] -and -not $_["Unscanned"] -and -not $_["Denied"]}|Sort-Object -Property{$_["Size"]} -Descending
        $sU=$entries|Where-Object{$_["Unscanned"] -eq $true}
        $sD=$entries|Where-Object{$_["Denied"]    -eq $true}
        $sE=$entries|Where-Object{$_["Excluded"]  -eq $true}
        $sorted=@();foreach($e in $sN){$sorted+=$e};foreach($e in $sU){$sorted+=$e};foreach($e in $sD){$sorted+=$e};foreach($e in $sE){$sorted+=$e}
        $allScans[$dirPath]=@{Items=$sorted;Total=$dirTotalSizes[$dirPath]};$reportCount++
        Update-Progress $ACT "E5/5 : $reportCount / $nbScope" "" (90+[int](($reportCount*9)/($nbScope+1))) ($reportCount%2000-eq 0)
    }
    Write-Step "[E5] TERMINEE" $t "$reportCount construits | $($dirsUnscanned.Count) non scannes | $($accessDenied.Count) proteges"
    Update-Progress $ACT "Termine" "" 100 $true $true; Write-Progress -Activity $ACT -Completed

    $totalElap=[int]((Get-Date)-$scanStart).TotalSeconds
    Write-Log "SCAN TERMINE v$SCRIPT_VERSION en ${totalElap}s | Dossiers : $reportCount | Fichiers : $totalFilesScanned | Taille : $(Format-Size $dirTotalSizes[$RootPath]) | CloudOnly : $cloudOnlyCount"

    return @{
        Scans            = $allScans
        Excluded         = $excludedFound
        Unscanned        = $dirsUnscanned
        AccessDenied     = $accessDenied
        ModeName         = $modeName
        ElapsedSec       = $totalElap
        MaxDepth         = $MaxDepth
        Unlimited        = $unlimited
        PathType         = $pathType
        MethodStats      = $methodStats
        JunctionsSkipped = $junctionsTotal
        IsOneDrive       = $IsOneDrive
        CloudOnlyCount   = $cloudOnlyCount
        CloudOnlySize    = $cloudOnlySize
        LocalSize        = $localOnlySize
    }
}

# ============================================================
# Generation HTML
# ============================================================
function Get-HtmlReport {
    param(
        [string]$RootPath,
        [hashtable]$AllScans,
        [string[]]$UnscannedDirs,
        [string[]]$AccessDeniedDirs,
        [int]$MaxDepth,
        [bool]$Unlimited,
        [string]$ModeName,
        [string]$ModeAccuracy,
        [int]$ElapsedSec,
        [string]$PathType,
        [string]$FullUserName,
        [string]$ScanDateTime,
        [hashtable]$MethodStats,
        [int]$JunctionsSkipped,
        [bool]$IsOneDrive = $false,
        [int]$CloudOnlyCount = 0,
        [long]$LocalSize = 0
    )

    $ACT="PS-NCDU v$SCRIPT_VERSION"; $t=Get-Date
    Write-Log "[HTML] Generation v$SCRIPT_VERSION : $($AllScans.Count) dossiers"
    Update-Progress $ACT "HTML : JSON..." "" 2 $true $true

    $jsonParts=@(); $jsonCount=0
    foreach($scanPath in $AllScans.Keys){
        $jsonCount++
        Update-Progress $ACT "HTML : JSON $jsonCount / $($AllScans.Count)" "" (2+[int](($jsonCount*80)/($AllScans.Count+1))) ($jsonCount%5000-eq 0)
        $scanData=$AllScans[$scanPath]; $items=$scanData["Items"]; $total=$scanData["Total"]
        $ip=@()
        foreach($item in $items){
            $sizeStr=Format-Size $item["Size"]
            $ip+="{"+"""name"":""$(ConvertTo-JsonSafe $item["Name"])"","+"""fullPath"":""$(ConvertTo-JsonSafe $item["FullPath"])"","+"""size"":$($item["Size"]),"+"""sizeStr"":""$(ConvertTo-JsonSafe $sizeStr)"","+"""isDir"":$(if($item["IsDir"]){"true"}else{"false"}),"+"""excluded"":$(if($item["Excluded"]){"true"}else{"false"}),"+"""unscanned"":$(if($item["Unscanned"]){"true"}else{"false"}),"+"""denied"":$(if($item["Denied"]){"true"}else{"false"}),"+"""ext"":""$(ConvertTo-JsonSafe $item["Ext"])"","+"""type"":""$(if($item["IsDir"]){"DIR"}else{"FILE"})"""+"}"
        }
        $jsonParts+="""$(ConvertTo-JsonSafe $scanPath)"":{""items"":["+($ip -join ",")+"],""total"":$total}"
    }
    $jsonScans="{"+($jsonParts -join ",")+"}"
    $exParts=@(); foreach($ex in $EXCLUDED_DIRS){$exParts+="""$(ConvertTo-JsonSafe $ex)"""}
    $jsonExcluded="["+($exParts -join ",")+"]"

    Write-Step "[HTML] JSON serialise" $t "$jsonCount dossiers"
    Update-Progress $ACT "HTML : Generation page..." "" 85 $true $true; $t=Get-Date

    $rootPathSafe  = ConvertTo-JsonSafe $RootPath
    $logEnc        = ConvertTo-HtmlEncoded $logFile
    $rootEnc       = ConvertTo-HtmlEncoded $RootPath
    $nbScans       = $AllScans.Count
    $emailEnc      = ConvertTo-HtmlEncoded $USER_EMAIL
    $authorEnc     = ConvertTo-HtmlEncoded $SCRIPT_AUTHOR
    $modeEncAcc    = ConvertTo-HtmlEncoded $ModeAccuracy
    $logoSvg       = $NN_LOGO_SVG
    $maxDepthJs    = $MaxDepth
    $unlimitedJs   = if($Unlimited){"true"}else{"false"}
    $nbUnscanned   = $UnscannedDirs.Count
    $nbDenied      = $AccessDeniedDirs.Count
    $pathTypeEnc   = ConvertTo-HtmlEncoded $PathType
    $pathIcon      = if($RootPath -match '^\\\\'){"&#127760;"}elseif($IsOneDrive){"&#9729;"}else{"&#128190;"}
    $depthLabel    = if($Unlimited){"Illimitee"}else{"$MaxDepth niveaux"}
    $depthLabelEnc = ConvertTo-HtmlEncoded $depthLabel
    $fullUserEnc   = ConvertTo-HtmlEncoded $FullUserName
    $scanDTEnc     = ConvertTo-HtmlEncoded $ScanDateTime
    $machineEnc    = ConvertTo-HtmlEncoded $env:COMPUTERNAME
    $usernameEnc   = ConvertTo-HtmlEncoded $env:USERNAME
    $footerExtra   = ""
    if ($JunctionsSkipped -gt 0) { $footerExtra += " &middot; $JunctionsSkipped jonctions ignorees" }
    if ($null -ne $MethodStats -and ($MethodStats["NET"] -gt 0 -or $MethodStats["CMD"] -gt 0)) {
        $footerExtra += " &middot; Fallback .NET=$($MethodStats['NET']) CMD=$($MethodStats['CMD'])"
    }
    $exListHtml=""; foreach($ex in $EXCLUDED_DIRS){$exListHtml+="<li><code>$(ConvertTo-HtmlEncoded $ex)</code></li>"}
    $dnListHtml=""; foreach($dn in $AccessDeniedDirs){$dnListHtml+="<li><code>$(ConvertTo-HtmlEncoded $dn)</code></li>"}

    # Barre alertes compacte
    $alertParts = @()
    if ($nbDenied -gt 0)           { $alertParts += "<span class='ac-item ac-dn'>&#128274; <strong>$nbDenied proteges</strong><span class='ac-hint' onclick='toggleDet(""dD"")'>Voir</span></span>" }
    if ($nbUnscanned -gt 0)        { $alertParts += "<span class='ac-item ac-un'>&#128269; <strong>$nbUnscanned non scannes</strong><span class='ac-hint' onclick='toggleDet(""dU"")'>Plus profond</span></span>" }
    if ($EXCLUDED_DIRS.Count -gt 0){ $alertParts += "<span class='ac-item ac-ex'>&#9888; <strong>$($EXCLUDED_DIRS.Count) exclus</strong><span class='ac-hint' onclick='toggleDet(""dE"")'>Voir</span></span>" }
    # ✅ v3.6 : Alerte OneDrive cloud-only
    if ($IsOneDrive -and $CloudOnlyCount -gt 0) {
        $alertParts += "<span class='ac-item ac-od'>&#9729; <strong>$CloudOnlyCount fichiers cloud-only</strong><span class='ac-hint' onclick='toggleDet(""dOD"")'>Details</span></span>"
    }
    $alertBarHtml = ""
    if ($alertParts.Count -gt 0) { $alertBarHtml = "<div class='alert-compact'>" + ($alertParts -join "<span class='ac-sep'>|</span>") + "</div>" }

    $alertDetailsHtml = ""
    if ($nbDenied -gt 0)           { $alertDetailsHtml += "<div class='alert-det' id='dD'><ul>$dnListHtml</ul></div>" }
    if ($nbUnscanned -gt 0)        { $alertDetailsHtml += "<div class='alert-det' id='dU'><p>Relancez avec profondeur <strong>0</strong> (illimitee) ou augmentez le niveau.</p></div>" }
    if ($EXCLUDED_DIRS.Count -gt 0){ $alertDetailsHtml += "<div class='alert-det' id='dE'><ul>$exListHtml</ul></div>" }
    # ✅ v3.6 : Detail OneDrive cloud-only
    if ($IsOneDrive -and $CloudOnlyCount -gt 0) {
        $localSizeStr = Format-Size $LocalSize
        $alertDetailsHtml += @"
<div class='alert-det' id='dOD'>
  <p><strong>&#9729; Fichiers OneDrive cloud-only : $CloudOnlyCount</strong></p>
  <p>Ces fichiers existent uniquement dans le cloud (icone nuage dans l Explorateur).</p>
  <p>Ils ne sont <strong>pas telecharges localement</strong> et ne sont <strong>pas comptes</strong> dans les tailles affichees.</p>
  <p>Taille locale reelle affichee : <strong>$localSizeStr</strong></p>
  <p style='margin-top:8px'>Pour les inclure : faites un clic droit sur le dossier OneDrive &rarr; <strong>Toujours garder sur cet appareil</strong>, puis relancez le scan.</p>
</div>
"@
    }

    # ── JS hors here-string (single-quote = pas d interpolation) ──
    $jsCode = @'
var DATA, EXCLUDED, ROOT_PATH, MAX_DEPTH, UNLIMITED;
var curPath, sortCol='size', sortAsc=false;

function init(data, excluded, rootPath, maxDepth, unlimited){
    DATA=data; EXCLUDED=excluded; ROOT_PATH=rootPath;
    MAX_DEPTH=maxDepth; UNLIMITED=unlimited; curPath=rootPath;
}

(function(){
    var el=document.getElementById('userFullName');
    var fn=el?el.getAttribute('data-name'):'';
    var p=fn.split(' ').filter(function(x){return x.length>0;}),i='';
    if(p.length>=2)i=(p[0][0]+p[p.length-1][0]).toUpperCase();
    else if(p.length===1)i=p[0].substring(0,2).toUpperCase();
    else i='?';
    var av=document.getElementById('userAv');
    if(av)av.textContent=i;
})();

var EI={
    '.pdf':'&#128196;','.doc':'&#128196;','.docx':'&#128196;','.xls':'&#128200;','.xlsx':'&#128200;',
    '.ppt':'&#128202;','.pptx':'&#128202;','.txt':'&#128441;','.md':'&#128441;','.rtf':'&#128441;',
    '.odt':'&#128196;','.ods':'&#128200;','.odp':'&#128202;','.log':'&#128203;',
    '.ps1':'&#9096;','.py':'&#128013;','.js':'&#128252;','.ts':'&#128252;',
    '.html':'&#127760;','.htm':'&#127760;','.css':'&#127912;','.php':'&#128064;',
    '.java':'&#9749;','.cs':'&#128187;','.cpp':'&#128187;','.c':'&#128187;','.h':'&#128187;',
    '.go':'&#128187;','.rs':'&#128187;','.rb':'&#128312;',
    '.sh':'&#128192;','.bat':'&#128192;','.cmd':'&#128192;',
    '.sql':'&#128020;','.xml':'&#127991;','.json':'&#128210;','.yaml':'&#128210;','.yml':'&#128210;',
    '.ini':'&#9881;','.conf':'&#9881;','.config':'&#9881;','.env':'&#128196;',
    '.jpg':'&#128444;','.jpeg':'&#128444;','.png':'&#128444;','.gif':'&#127987;','.bmp':'&#128444;',
    '.svg':'&#128444;','.ico':'&#128444;','.webp':'&#128444;','.tiff':'&#128444;','.raw':'&#128247;',
    '.psd':'&#127912;','.ai':'&#127912;','.heic':'&#128247;',
    '.mp4':'&#127916;','.avi':'&#127916;','.mkv':'&#127916;','.mov':'&#127916;','.wmv':'&#127916;',
    '.flv':'&#127916;','.webm':'&#127916;','.mpg':'&#127916;','.mpeg':'&#127916;',
    '.mp3':'&#127925;','.wav':'&#127925;','.flac':'&#127925;','.aac':'&#127925;',
    '.ogg':'&#127925;','.m4a':'&#127925;','.wma':'&#127925;',
    '.zip':'&#128230;','.rar':'&#128230;','.7z':'&#128230;','.tar':'&#128230;','.gz':'&#128230;',
    '.bz2':'&#128230;','.xz':'&#128230;','.iso':'&#128191;','.dmg':'&#128191;',
    '.exe':'&#9881;','.msi':'&#9881;','.dll':'&#9881;','.so':'&#9881;',
    '.deb':'&#9881;','.rpm':'&#9881;','.apk':'&#128241;','.app':'&#9881;',
    '.db':'&#128020;','.sqlite':'&#128020;','.mdb':'&#128020;','.mdf':'&#128020;','.ldf':'&#128020;',
    '.bak':'&#128226;','.backup':'&#128226;','.old':'&#128226;','.tmp':'&#128226;',
    '.ttf':'&#127381;','.otf':'&#127381;','.woff':'&#127381;','.woff2':'&#127381;',
    '.torrent':'&#128279;','.ics':'&#128197;','.vcf':'&#128100;',
    '.eml':'&#128140;','.msg':'&#128140;','.pst':'&#128188;',
    '.key':'&#128272;','.pem':'&#128272;','.crt':'&#128272;','.cer':'&#128272;',
    'dir':'&#128193;','net':'&#127760;','od':'&#9729;','denied':'&#128274;','unknown':'&#128196;'
};
function gIcon(ext,isDir,isDen,path){
    if(isDen)return EI['denied'];
    if(isDir){
        if(path&&path.indexOf('\\\\')===0)return EI['net'];
        if(path&&(path.toLowerCase().indexOf('onedrive')>=0))return EI['od'];
        return EI['dir'];
    }
    if(!ext)return EI['unknown'];
    return EI[ext.toLowerCase()]||EI['unknown'];
}

function toggleTheme(){
    var h=document.documentElement,isL=h.getAttribute('data-theme')==='light';
    h.setAttribute('data-theme',isL?'dark':'light');
    document.getElementById('themeBtn').innerHTML=isL?'&#9790; Dark':'&#9728; Light';
    try{localStorage.setItem('nn-theme',isL?'dark':'light');}catch(e){}
}
(function(){
    try{
        var t=localStorage.getItem('nn-theme');
        if(t){
            document.documentElement.setAttribute('data-theme',t);
            var b=document.getElementById('themeBtn');
            if(b)b.innerHTML=t==='light'?'&#9728; Light':'&#9790; Dark';
        }
    }catch(e){}
})();

function toggleDet(id){
    var el=document.getElementById(id);
    if(el)el.style.display=(el.style.display==='block')?'none':'block';
}
function escH(s){
    if(!s)return '';
    return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function nP(p){
    if(!p)return '';
    var iu=(p.length>=2&&p.charAt(0)==='\\'&&p.charAt(1)==='\\');
    p=p.replace(/\//g,'\\');
    if(iu){
        var body=p.replace(/^\\+/,'').replace(/\\{2,}/g,'\\');
        return '\\\\'+body;
    }
    p=p.replace(/\\{2,}/g,'\\');
    if(p.length>3&&p.charAt(p.length-1)==='\\')p=p.substring(0,p.length-1);
    if(p.length===2&&((p.charAt(0)>='A'&&p.charAt(0)<='Z')||(p.charAt(0)>='a'&&p.charAt(0)<='z'))&&p.charAt(1)===':')p=p+'\\';
    return p;
}
function pOf(p){
    p=nP(p);
    var iu=(p.length>=2&&p.charAt(0)==='\\'&&p.charAt(1)==='\\');
    if(p.length===3&&p.charAt(1)===':'&&p.charAt(2)==='\\')return null;
    if(iu){
        var parts=p.substring(2).split('\\').filter(function(x){return x.length>0;});
        if(parts.length<=2)return null;
        return '\\\\'+parts.slice(0,parts.length-1).join('\\');
    }
    var i=p.lastIndexOf('\\');
    if(i<0)return null;
    if(i===2)return p.substring(0,3);
    var r=p.substring(0,i);
    if(r.length===2&&r.charAt(1)===':')r=r+'\\';
    return r||null;
}
function fSz(b){
    b=Number(b);if(b<0)return '?';
    if(b>=1099511627776)return (b/1099511627776).toFixed(2)+' TB';
    if(b>=1073741824)return (b/1073741824).toFixed(2)+' GB';
    if(b>=1048576)return (b/1048576).toFixed(2)+' MB';
    if(b>=1024)return (b/1024).toFixed(2)+' KB';
    return b+' B';
}
function bColor(sz){
    if(sz>=1073741824)return 'var(--danger)';
    if(sz>=104857600)return 'var(--warn)';
    if(sz>=10485760)return '#c47d00';
    return 'var(--ok)';
}
function isExcl(path){
    path=nP(path).toLowerCase();
    for(var i=0;i<EXCLUDED.length;i++){
        var ex=nP(EXCLUDED[i]).toLowerCase();
        if(path===ex||path.indexOf(ex+'\\')===0)return true;
    }
    return false;
}
function findS(p){
    p=nP(p).toLowerCase();
    for(var k in DATA){if(nP(k).toLowerCase()===p)return DATA[k];}
    return null;
}
function navigateTo(p){
    if(!p||!p.trim())return;
    p=nP(p.trim());curPath=p;
    document.getElementById('pathInput').value=p;
    renderBC(p);renderContent(p);
}
function goUp(){var p=pOf(curPath);if(p)navigateTo(p);}
function renderBC(path){
    path=nP(path);
    var iu=(path.length>=2&&path.charAt(0)==='\\'&&path.charAt(1)==='\\');
    var html='',built='';
    if(iu){
        var pts=path.substring(2).split('\\').filter(function(x){return x.length>0;});
        if(pts.length>=1){
            built='\\\\'+pts[0];
            html+='<span style="color:rgba(255,255,255,.4);font-size:.78em">&#127760;</span> ';
            if(pts.length===1){html+='<a class="bcp-a c">'+escH('\\\\'+pts[0])+'</a>';}
            else{var b0=built;html+='<a class="bcp-a" onclick="navigateTo(\''+b0.replace(/\\/g,'\\\\')+'\')" >'+escH('\\\\'+pts[0])+'</a>';}
        }
        for(var i=1;i<pts.length;i++){
            built+='\\'+pts[i];
            html+=' <span class="bcp-sep">&#8250;</span> ';
            var bi=built;
            if(i===pts.length-1){html+='<a class="bcp-a c">'+escH(pts[i])+'</a>';}
            else{html+='<a class="bcp-a" onclick="navigateTo(\''+bi.replace(/\\/g,'\\\\')+'\')" >'+escH(pts[i])+'</a>';}
        }
    } else {
        var clean=path;
        if(clean.length>3&&clean.charAt(clean.length-1)==='\\')clean=clean.substring(0,clean.length-1);
        var pts2=clean.split('\\').filter(function(x){return x.length>0;});
        if(pts2.length>0&&pts2[0].charAt(1)===':'){
            built=pts2[0]+'\\';
            if(pts2.length===1){html+='<a class="bcp-a c">'+escH(pts2[0])+'</a>';}
            else{var b1=built;html+='<a class="bcp-a" onclick="navigateTo(\''+b1.replace(/\\/g,'\\\\')+'\')" >'+escH(pts2[0])+'</a>';}
            pts2.shift();
        }
        for(var j=0;j<pts2.length;j++){
            built+=pts2[j]+'\\';
            html+=' <span class="bcp-sep">&#8250;</span> ';
            var bj=built;
            if(j===pts2.length-1){html+='<a class="bcp-a c">'+escH(pts2[j])+'</a>';}
            else{html+='<a class="bcp-a" onclick="navigateTo(\''+bj.replace(/\\/g,'\\\\')+'\')" >'+escH(pts2[j])+'</a>';}
        }
    }
    var bc=document.getElementById('breadcrumb');
    if(bc)bc.innerHTML=html||'<a class="bcp-a c">&#127968; Racine</a>';
}
function renderContent(path){
    if(isExcl(path)){
        document.getElementById('stats').innerHTML='<div class="chip chip-excl"><span class="chip-label">&#9888;</span><span class="chip-value">Repertoire exclu</span></div>';
        document.getElementById('mainContent').innerHTML='<div class="msg"><span class="ic">&#9888;</span><h3>Repertoire exclu du scan</h3><p>Sa taille est inconnue.</p><code>'+escH(path)+'</code></div>';
        return;
    }
    var scan=findS(path);
    var isUDir=false,isDDir=false;
    if(!scan){
        var par=pOf(path);
        if(par){
            var pScan=findS(par);
            if(pScan){
                pScan.items.forEach(function(it){
                    var np=nP(it.fullPath).toLowerCase(),np2=nP(path).toLowerCase();
                    if(np===np2){if(it.unscanned)isUDir=true;if(it.denied)isDDir=true;}
                });
            }
        }
    }
    if(!scan){
        if(isDDir){
            document.getElementById('stats').innerHTML='<div class="chip chip-den"><span class="chip-label">&#128274;</span><span class="chip-value">Protege ACL NTFS</span></div>';
            document.getElementById('mainContent').innerHTML='<div class="msg"><span class="ic">&#128274;</span><h3>Dossier protege par ACL NTFS</h3><p>Inaccessible avec ce compte.</p><p class="hint-d">Demande via <strong>NovoAccess</strong></p><code>'+escH(path)+'</code></div>';
        } else if(isUDir){
            document.getElementById('stats').innerHTML='<div class="chip chip-uns"><span class="chip-label">&#128269;</span><span class="chip-value">Non scanne - profondeur '+MAX_DEPTH+'</span></div>';
            document.getElementById('mainContent').innerHTML='<div class="msg"><span class="ic">&#128194;</span><h3>Dossier non scanne</h3><p>Scan arrete a la profondeur <strong>'+MAX_DEPTH+'</strong>.</p><p class="hint-u">&#128161; Relancez avec profondeur 0 ou ce chemin :</p><code>'+escH(path)+'</code></div>';
        } else {
            var msg=UNLIMITED?'Ce dossier n\'est pas dans le rapport.':'Depasse la profondeur '+MAX_DEPTH+'.';
            document.getElementById('stats').innerHTML='<div class="chip chip-warn"><span class="chip-label">&#9888;</span><span class="chip-value">Non scanne</span></div>';
            document.getElementById('mainContent').innerHTML='<div class="msg"><span class="ic">&#128194;</span><h3>Dossier non scanne</h3><p>'+msg+'</p><p class="hint">&#128161; Relancez avec ce chemin</p><code>'+escH(path)+'</code></div>';
        }
        return;
    }
    var items=scan.items,total=scan.total,nbD=0,nbF=0,nbEx=0,nbUn=0,nbDn=0;
    items.forEach(function(i){if(i.excluded)nbEx++;else if(i.unscanned)nbUn++;else if(i.denied)nbDn++;else if(i.isDir)nbD++;else nbF++;});
    var st='<div class="chip"><span class="chip-label">Taille</span><span class="chip-value">'+fSz(total)+'</span></div>'+
        '<div class="chip"><span class="chip-label">Dossiers</span><span class="chip-value">'+nbD+'</span></div>'+
        '<div class="chip"><span class="chip-label">Fichiers</span><span class="chip-value">'+nbF+'</span></div>'+
        '<div class="chip"><span class="chip-label">Elements</span><span class="chip-value">'+items.length+'</span></div>';
    if(nbDn>0)st+='<div class="chip chip-den"><span class="chip-label">&#128274; Proteges</span><span class="chip-value">'+nbDn+'</span></div>';
    if(nbUn>0)st+='<div class="chip chip-uns"><span class="chip-label">&#128269; Non scannes</span><span class="chip-value">'+nbUn+'</span></div>';
    if(nbEx>0)st+='<div class="chip chip-excl"><span class="chip-label">&#9888; Exclus</span><span class="chip-value">'+nbEx+'</span></div>';
    document.getElementById('stats').innerHTML=st;
    if(items.length===0){document.getElementById('mainContent').innerHTML='<div class="msg"><span class="ic">&#128194;</span><h3>Dossier vide</h3><p>Aucun fichier accessible.</p></div>';return;}
    var rows='';
    items.forEach(function(item){
        var isEx=item.excluded||isExcl(item.fullPath);
        var isUn=item.unscanned,isDn=item.denied;
        var fp=item.fullPath.replace(/\\/g,'\\\\');
        var icon=gIcon(item.ext||'',item.isDir,isDn,item.fullPath);
        var isDeep=!UNLIMITED&&item.isDir&&!isEx&&!isUn&&!isDn&&!findS(item.fullPath);
        var tag=isDeep?'<span style="font-size:.68em;color:var(--td);background:var(--card);padding:1px 5px;border-radius:3px;margin-left:5px;border:1px solid var(--border)">+</span>':'';
        var nC,sC,bC,bg,rC;
        if(isDn){
            rC='row-dn';
            nC='<span class="dir-ld">'+escH(item.name)+'</span><span class="dn-lbl">&#128274; Protege</span>';
            sC='<span class="sz-dn">?</span>';
            bC='<div class="bar-wrap"><div class="bar-bg"><div class="bar-dn"></div></div><span class="bar-dn-t">&#128274;</span></div>';
            bg='<span class="tbadge tb-dn">ACL</span>';
        } else if(isEx){
            rC='row-ex';
            nC='<a class="dir-l" onclick="navigateTo(\''+fp+'\')">'+escH(item.name)+'</a><span class="ex-lbl">&#9888;</span>';
            sC='<span class="sz-unk">?</span>';
            bC='<div class="bar-wrap"><div class="bar-bg"><div class="bar-ex"></div></div><span class="bar-unk">?</span></div>';
            bg='<span class="tbadge tb-ex">EXCLU</span>';
        } else if(isUn){
            rC='row-un';
            nC='<a class="dir-lu" onclick="navigateTo(\''+fp+'\')">'+escH(item.name)+'</a>';
            sC='<span class="sz-uns">?</span>';
            bC='<div class="bar-wrap"><div class="bar-bg"><div class="bar-uns"></div></div><span class="bar-uns-t">?</span></div>';
            bg='<span class="tbadge tb-un">?</span>';
        } else {
            rC='';var pct=total>0?Math.round((item.size/total)*1000)/10:0;
            nC=item.isDir?'<a class="dir-l" onclick="navigateTo(\''+fp+'\')">'+escH(item.name)+'</a>'+tag:'<span class="file-n">'+escH(item.name)+'</span>';
            sC=escH(item.sizeStr);
            bC='<div class="bar-wrap"><div class="bar-bg"><div class="bar-fill" style="width:'+pct+'%;background:'+bColor(item.size)+'"></div></div><span class="bar-pct">'+pct+'%</span></div>';
            bg=item.isDir?'<span class="tbadge tb-dir">DIR</span>':'<span class="tbadge tb-file">FILE</span>';
        }
        rows+='<tr class="'+rC+'"><td class="c-icon">'+icon+'</td><td class="col-name">'+nC+'</td><td class="c-size">'+sC+'</td><td class="c-bar">'+bC+'</td><td class="c-type">'+bg+'</td></tr>';
    });
    document.getElementById('mainContent').innerHTML=
        '<div class="table-wrap"><table><thead><tr>'+
        '<th class="c-icon"></th>'+
        '<th id="th-n" onclick="sortBy(\'name\')">Nom</th>'+
        '<th id="th-s" onclick="sortBy(\'size\')" style="text-align:right">Taille</th>'+
        '<th>Utilisation</th>'+
        '<th class="c-type">Type</th>'+
        '</tr></thead><tbody>'+rows+'</tbody></table></div>';
    ['th-n','th-s'].forEach(function(id){var el=document.getElementById(id);if(el)el.classList.remove('sa','sd');});
    var thEl=document.getElementById(sortCol==='name'?'th-n':'th-s');
    if(thEl)thEl.classList.add(sortAsc?'sa':'sd');
}
function sortBy(col){
    if(sortCol===col){sortAsc=!sortAsc;}else{sortCol=col;sortAsc=(col==='name');}
    var scan=findS(curPath);if(!scan)return;
    var normal=scan.items.filter(function(i){return !i.excluded&&!i.unscanned&&!i.denied;});
    var uns=scan.items.filter(function(i){return i.unscanned;});
    var den=scan.items.filter(function(i){return i.denied;});
    var exc=scan.items.filter(function(i){return i.excluded;});
    normal.sort(function(a,b){
        var va=a[col],vb=b[col];
        if(typeof va==='string'){va=va.toLowerCase();vb=vb.toLowerCase();}
        if(va<vb)return sortAsc?-1:1;if(va>vb)return sortAsc?1:-1;return 0;
    });
    scan.items=normal.concat(uns).concat(den).concat(exc);
    renderContent(curPath);
}
window.onload=function(){navigateTo(ROOT_PATH);};
'@

    $html = @"
<!DOCTYPE html>
<html lang="fr" data-theme="light">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width,initial-scale=1.0">
    <title>PS-NCDU v$SCRIPT_VERSION - $rootEnc</title>
    <style>
    :root{--bg:#f0f4f8;--surface:#fff;--card:#e8edf3;--card-h:#dde4ed;--border:#cdd6e0;--border-l:#dde4ed;--text:#0f1923;--tm:#4a6080;--td:#8da0b8;--accent:#0063be;--al:#004fa3;--ag:rgba(0,99,190,.08);--ok:#007a5e;--warn:#b86b00;--danger:#c0202e;--unscanned:#5e3d8f;--od:#0078d4;--shadow:0 2px 8px rgba(0,0,0,.1);--r:8px;--rs:5px;--rp:20px}
    [data-theme="dark"]{--bg:#0f1923;--surface:#162032;--card:#1e2d42;--card-h:#243350;--border:#2a3f5f;--border-l:#1e3050;--text:#D3B276;--tm:#a08a5a;--td:#6b5a3a;--accent:#0063be;--al:#D3B276;--ag:rgba(211,178,118,.15);--ok:#00c48c;--warn:#f5a623;--danger:#e8394a;--unscanned:#9b7fd4;--od:#4da3ff;--shadow:0 2px 12px rgba(0,0,0,.45)}
    *{box-sizing:border-box;margin:0;padding:0}
    body{font-family:'Segoe UI',Roboto,Arial,sans-serif;background:var(--bg);color:var(--text);min-height:100vh;font-size:14px;transition:background .3s,color .3s}
    .header-outer{background:linear-gradient(160deg,#001965 0%,#0063be 40%,#004fa3 100%);position:sticky;top:0;z-index:200;box-shadow:0 4px 16px rgba(0,0,0,.25)}
    [data-theme="dark"] .header-outer{background:linear-gradient(160deg,#0f1923 0%,#1e2d42 100%)}
    .header-inner{padding:8px 24px 10px;display:flex;flex-direction:column;gap:7px}
    .hdr-r1{display:flex;align-items:center;gap:12px}
    .hdr-r1-left{display:flex;align-items:center;gap:10px;flex-shrink:0}
    .hdr-r1-mid{flex:9;display:flex;align-items:center;gap:6px;min-width:0}
    .hdr-r1-right{display:flex;align-items:center;gap:10px;flex-shrink:0}
    .hdr-title-main{font-size:1em;font-weight:700;color:#fff;white-space:nowrap}
    .hdr-title-main b{color:rgba(255,255,255,.5)}
    .hdr-title-sub{font-size:.65em;color:rgba(255,255,255,.5);margin-top:1px;white-space:nowrap}
    [data-theme="dark"] .hdr-title-main{color:var(--text)}[data-theme="dark"] .hdr-title-main b{color:var(--td)}[data-theme="dark"] .hdr-title-sub{color:var(--tm)}
    .hdr-chip{font-size:.68em;font-weight:700;background:rgba(255,255,255,.15);color:#fff;border:1px solid rgba(255,255,255,.3);padding:1px 9px;border-radius:var(--rp);white-space:nowrap;flex-shrink:0}
    [data-theme="dark"] .hdr-chip{background:rgba(211,178,118,.1);color:var(--text);border-color:rgba(211,178,118,.25)}
    .hdr-input{background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.25);color:#fff;padding:5px 14px;border-radius:var(--rp);font-size:.82em;font-family:inherit;outline:none;width:100%;min-width:100px}
    .hdr-input:focus{background:rgba(255,255,255,.2);border-color:rgba(255,255,255,.5);box-shadow:0 0 0 3px rgba(255,255,255,.08)}
    .hdr-input::placeholder{color:rgba(255,255,255,.45)}
    [data-theme="dark"] .hdr-input{background:rgba(211,178,118,.08);border-color:rgba(211,178,118,.2);color:var(--text)}
    [data-theme="dark"] .hdr-input:focus{background:rgba(211,178,118,.15);border-color:rgba(211,178,118,.45)}
    [data-theme="dark"] .hdr-input::placeholder{color:var(--td)}
    .hdr-btn{background:rgba(255,255,255,.15);color:#fff;border:1px solid rgba(255,255,255,.3);padding:5px 11px;border-radius:var(--rs);cursor:pointer;font-size:.78em;font-weight:500;font-family:inherit;transition:all .2s;white-space:nowrap;flex-shrink:0}
    .hdr-btn:hover{background:rgba(255,255,255,.25)}
    [data-theme="dark"] .hdr-btn{background:rgba(211,178,118,.1);color:var(--text);border-color:rgba(211,178,118,.25)}
    [data-theme="dark"] .hdr-btn:hover{background:rgba(211,178,118,.2)}
    .hdr-av{width:28px;height:28px;border-radius:50%;background:rgba(255,255,255,.2);border:2px solid rgba(255,255,255,.35);color:#fff;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:.78em;flex-shrink:0}
    [data-theme="dark"] .hdr-av{background:rgba(211,178,118,.2);border-color:rgba(211,178,118,.35);color:var(--text)}
    .hdr-un{font-size:.78em;font-weight:600;color:#fff;white-space:nowrap;max-width:130px;overflow:hidden;text-overflow:ellipsis}
    .hdr-us{font-size:.68em;color:rgba(255,255,255,.55);white-space:nowrap}
    [data-theme="dark"] .hdr-un{color:var(--text)}[data-theme="dark"] .hdr-us{color:var(--tm)}
    .hdr-logo{cursor:pointer;display:flex;align-items:center;flex-shrink:0;color:#fff;transition:opacity .2s}
    .hdr-logo:hover{opacity:.75}
    .hdr-logo svg{height:24px;width:auto;display:block;color:inherit}
    [data-theme="dark"] .hdr-logo{color:#D3B276}
    .hdr-r2{display:flex;align-items:center;gap:5px;flex-wrap:wrap;border-top:1px solid rgba(255,255,255,.1);padding-top:6px}
    [data-theme="dark"] .hdr-r2{border-top-color:rgba(211,178,118,.1)}
    .sp{display:inline-flex;align-items:center;gap:4px;font-size:.72em;color:rgba(255,255,255,.75);padding:2px 8px;background:rgba(255,255,255,.1);border-radius:var(--rp);white-space:nowrap}
    .sp b{color:#fff;font-weight:700}
    [data-theme="dark"] .sp{background:rgba(211,178,118,.08);color:var(--tm)}[data-theme="dark"] .sp b{color:var(--text)}
    .sp-od{background:rgba(0,120,212,.25);border:1px solid rgba(0,120,212,.5)}
    [data-theme="dark"] .sp-od{background:rgba(77,163,255,.15);border-color:rgba(77,163,255,.4)}
    .vdg{width:1px;height:14px;background:rgba(255,255,255,.2);flex-shrink:0}
    [data-theme="dark"] .vdg{background:rgba(211,178,118,.2)}
    .bcp{background:rgba(255,255,255,.08);border-radius:var(--rp);padding:3px;display:flex;align-items:center;gap:2px;margin-left:auto}
    [data-theme="dark"] .bcp{background:rgba(211,178,118,.05)}
    .bcp-a{color:rgba(255,255,255,.7);text-decoration:none;padding:2px 8px;border-radius:var(--rp);font-size:.78em;cursor:pointer;transition:all .15s;white-space:nowrap}
    .bcp-a:hover{background:rgba(255,255,255,.15);color:#fff}
    .bcp-a.c{background:rgba(255,255,255,.18);color:#fff;font-weight:600}
    [data-theme="dark"] .bcp-a{color:var(--tm)}[data-theme="dark"] .bcp-a:hover,[data-theme="dark"] .bcp-a.c{background:rgba(211,178,118,.15);color:var(--text)}
    .bcp-sep{color:rgba(255,255,255,.3);font-size:.78em}[data-theme="dark"] .bcp-sep{color:var(--td)}
    .alert-compact{display:flex;align-items:center;flex-wrap:wrap;padding:5px 24px;background:var(--card);border-bottom:1px solid var(--border);font-size:.8em;min-height:32px}
    .ac-item{display:inline-flex;align-items:center;gap:5px;padding:3px 12px}
    .ac-dn{color:var(--danger)}.ac-un{color:var(--unscanned)}.ac-ex{color:var(--warn)}.ac-od{color:var(--od)}
    .ac-sep{color:var(--border);padding:0 4px;font-size:1.1em;user-select:none}
    .ac-hint{color:var(--al);cursor:pointer;text-decoration:underline;font-size:.88em;margin-left:6px;font-weight:500}
    .ac-hint:hover{opacity:.75}
    .alert-det{display:none;padding:8px 24px 10px 40px;background:var(--card);border-bottom:1px solid var(--border);font-size:.8em;color:var(--tm)}
    .alert-det p{padding:3px 0}.alert-det ul{margin:0;padding:0;list-style:none}.alert-det li{padding:2px 0}
    .alert-det code{padding:1px 5px;border-radius:3px;background:rgba(192,32,46,.08);color:var(--danger)}
    .statsbar{background:var(--surface);padding:9px 24px;border-bottom:1px solid var(--border);display:flex;gap:7px;flex-wrap:wrap;align-items:center}
    .chip{display:inline-flex;align-items:center;gap:5px;background:var(--card);border:1px solid var(--border);border-radius:var(--rp);padding:4px 12px;font-size:.78em}
    .chip-label{color:var(--tm)}.chip-value{color:var(--al);font-weight:600}
    .chip-warn{border-color:var(--danger)}.chip-warn .chip-value{color:var(--danger)}
    .chip-excl{border-color:var(--warn)}.chip-excl .chip-value{color:var(--warn)}
    .chip-uns{border-color:var(--unscanned)}.chip-uns .chip-value{color:var(--unscanned)}
    .chip-den{border-color:var(--danger)}.chip-den .chip-value{color:var(--danger)}
    .table-wrap{padding:14px 24px;overflow-x:auto}
    table{width:100%;border-collapse:collapse;font-size:.875em;background:var(--surface);border-radius:var(--r);overflow:hidden;box-shadow:var(--shadow)}
    thead th{background:var(--card);color:var(--tm);padding:11px 14px;text-align:left;border-bottom:2px solid var(--al);font-weight:600;font-size:.73em;letter-spacing:.5px;text-transform:uppercase;cursor:pointer;user-select:none;white-space:nowrap}
    thead th:hover{background:var(--card-h);color:var(--al)}
    thead th.sa::after{content:' \25b2';color:var(--al)}thead th.sd::after{content:' \25bc';color:var(--al)}
    tbody tr{border-bottom:1px solid var(--border-l);transition:background .1s}
    tbody tr:hover{background:var(--card)}tbody tr:last-child{border-bottom:none}
    tbody tr.row-ex{background:rgba(176,107,0,.05)}tbody tr.row-ex:hover{background:rgba(176,107,0,.1)}
    tbody tr.row-un{background:rgba(94,61,143,.05)}tbody tr.row-un:hover{background:rgba(94,61,143,.1)}
    tbody tr.row-dn{background:rgba(192,32,46,.05)}tbody tr.row-dn:hover{background:rgba(192,32,46,.1)}
    td{padding:9px 14px;vertical-align:middle}
    .c-icon{width:28px;text-align:center}.c-size{width:110px;text-align:right;font-family:monospace;font-size:.82em;color:var(--tm)}
    .c-bar{width:210px}.c-type{width:65px;text-align:center}
    .dir-l{color:var(--al);text-decoration:none;font-weight:500;cursor:pointer}.dir-l:hover{color:var(--text);text-decoration:underline}
    .dir-lu{color:var(--unscanned);text-decoration:none;font-weight:500;cursor:pointer}.dir-lu:hover{color:var(--text);text-decoration:underline}
    .dir-ld{color:var(--danger);font-weight:500;cursor:default}.file-n{color:var(--text)}
    .tbadge{font-size:.68em;padding:2px 6px;border-radius:3px;font-weight:600}
    .tb-dir{background:rgba(0,99,190,.1);color:var(--al)}.tb-file{background:var(--card);color:var(--td)}
    .tb-ex{background:rgba(176,107,0,.12);color:var(--warn)}.tb-un{background:rgba(94,61,143,.12);color:var(--unscanned)}.tb-dn{background:rgba(192,32,46,.12);color:var(--danger)}
    .ex-lbl{font-size:.68em;background:rgba(176,107,0,.12);color:var(--warn);padding:2px 6px;border-radius:3px;margin-left:5px;font-weight:500}
    .dn-lbl{font-size:.68em;background:rgba(192,32,46,.12);color:var(--danger);padding:2px 6px;border-radius:3px;margin-left:5px;font-weight:500}
    .sz-unk{color:var(--warn);font-weight:600;font-style:italic}.sz-uns{color:var(--unscanned);font-weight:600;font-style:italic}.sz-dn{color:var(--danger);font-weight:600;font-style:italic}
    .bar-wrap{display:flex;align-items:center;gap:7px}
    .bar-bg{background:var(--card);border-radius:4px;height:7px;flex:1;overflow:hidden;border:1px solid var(--border-l)}
    .bar-fill{height:100%;border-radius:4px;transition:width .3s}
    .bar-pct{font-size:.72em;color:var(--td);white-space:nowrap;min-width:36px;text-align:right}
    .bar-unk{font-size:.72em;color:var(--warn);font-weight:600}.bar-uns-t{font-size:.72em;color:var(--unscanned);font-weight:600;font-style:italic}.bar-dn-t{font-size:.72em;color:var(--danger);font-weight:600}
    .bar-ex{height:7px;background:repeating-linear-gradient(45deg,rgba(176,107,0,.1),rgba(176,107,0,.1) 4px,rgba(176,107,0,.2) 4px,rgba(176,107,0,.2) 8px)}
    .bar-uns{height:7px;background:repeating-linear-gradient(45deg,rgba(94,61,143,.1),rgba(94,61,143,.1) 4px,rgba(94,61,143,.2) 4px,rgba(94,61,143,.2) 8px)}
    .bar-dn{height:7px;background:repeating-linear-gradient(45deg,rgba(192,32,46,.1),rgba(192,32,46,.1) 4px,rgba(192,32,46,.2) 4px,rgba(192,32,46,.2) 8px)}
    .msg{padding:50px 24px;text-align:center}.msg .ic{font-size:2.8em;margin-bottom:14px;display:block}
    .msg h3{color:var(--text);margin-bottom:7px;font-weight:500;font-size:1.05em}.msg p{color:var(--tm);font-size:.85em;margin-top:5px}
    .msg .hint{margin-top:12px;font-size:.78em;background:rgba(176,107,0,.1);color:var(--warn);border:1px solid rgba(176,107,0,.3);padding:7px 14px;border-radius:var(--rs);display:inline-block}
    .msg .hint-u{margin-top:12px;font-size:.78em;background:rgba(94,61,143,.1);color:var(--unscanned);border:1px solid rgba(94,61,143,.3);padding:7px 14px;border-radius:var(--rs);display:inline-block}
    .msg .hint-d{margin-top:12px;font-size:.78em;background:rgba(192,32,46,.1);color:var(--danger);border:1px solid rgba(192,32,46,.3);padding:7px 14px;border-radius:var(--rs);display:inline-block}
    .msg code{font-family:monospace;font-size:.82em;color:var(--al);background:var(--card);padding:2px 7px;border-radius:3px;margin-top:7px;display:inline-block}
    .footer{text-align:center;padding:13px 24px;color:var(--td);font-size:.76em;border-top:1px solid var(--border);background:var(--surface);margin-top:14px;line-height:1.8}
    .footer a{color:var(--al);text-decoration:none}.footer a:hover{text-decoration:underline}
    .footer-sup{margin-top:5px;font-size:.83em;color:var(--tm)}
    ::-webkit-scrollbar{width:6px;height:6px}::-webkit-scrollbar-track{background:var(--bg)}::-webkit-scrollbar-thumb{background:var(--border);border-radius:3px}::-webkit-scrollbar-thumb:hover{background:var(--al)}
    @media(max-width:768px){.c-bar,.c-type{display:none}.hdr-r1-mid{flex:3}.hdr-un,.hdr-us,.hdr-r2{display:none}}
    </style>
</head>
<body>
<div class="header-outer">
  <div class="header-inner">
    <div class="hdr-r1">
      <div class="hdr-r1-left">
        <div><div class="hdr-title-main">PS-<b>NCDU</b></div><div class="hdr-title-sub">Disk Usage Analyzer</div></div>
        <span class="hdr-chip">v$SCRIPT_VERSION</span>
      </div>
      <div class="hdr-r1-mid">
        <input class="hdr-input" id="pathInput" value="$rootEnc"
               placeholder="C:\chemin ou \\serveur\partage"
               onkeydown="if(event.key==='Enter')navigateTo(this.value)">
        <button class="hdr-btn" onclick="navigateTo(document.getElementById('pathInput').value)">&#128269;</button>
        <button class="hdr-btn" onclick="goUp()">&#11014;</button>
        <button class="hdr-btn" onclick="navigateTo(ROOT_PATH)">&#127968;</button>
      </div>
      <div class="hdr-r1-right">
        <button class="hdr-btn" onclick="toggleTheme()" id="themeBtn">&#9728; Light</button>
        <div class="hdr-av" id="userAv">?</div>
        <div id="userFullName" data-name="$fullUserEnc">
          <div class="hdr-un" title="$fullUserEnc">$fullUserEnc</div>
          <div class="hdr-us">$machineEnc &middot; $usernameEnc</div>
        </div>
        <div class="hdr-logo" onclick="navigateTo(ROOT_PATH)" title="Racine">$logoSvg</div>
      </div>
    </div>
    <div class="hdr-r2">
      <div class="sp">&#128197; <b>$scanDTEnc</b></div>
      <div class="sp">$pathIcon <b>$rootEnc</b></div>
      <div class="sp">&#9201; <b>${ElapsedSec}s</b></div>
      <div class="sp">&#128193; <b>$nbScans</b> dossiers</div>
      $(if ($JunctionsSkipped -gt 0) { "<div class='sp'>&#128279; <b>$JunctionsSkipped</b> jonctions</div>" })
      $(if ($IsOneDrive -and $CloudOnlyCount -gt 0) { "<div class='sp sp-od'>&#9729; <b>$CloudOnlyCount</b> cloud-only</div>" })
      <div class="vdg"></div>
      <div class="sp">$(if($Unlimited){'&#8734;'}else{"&#128269; $MaxDepth niv."}) <b>$modeEncAcc</b></div>
      <div class="bcp" id="breadcrumb"><a class="bcp-a c">$rootEnc</a></div>
    </div>
  </div>
</div>

$alertBarHtml
$alertDetailsHtml

<div class="statsbar" id="stats"></div>
<div id="mainContent"></div>

<div class="footer">
    PS-NCDU v$SCRIPT_VERSION &middot; $pathIcon $pathTypeEnc &middot; $depthLabelEnc &middot; $nbScans dossiers &middot; $scanDTEnc &middot; $authorEnc$footerExtra
    <div class="footer-sup">Support : <a href="mailto:$emailEnc">$emailEnc</a> &nbsp;&middot;&nbsp; Log : <code>$logEnc</code></div>
</div>

<script>
$jsCode
init($jsonScans,$jsonExcluded,"$rootPathSafe",$maxDepthJs,$unlimitedJs);
</script>
</body>
</html>
"@

    Write-Progress -Activity $ACT -Completed
    Write-Step "[HTML] Generation complete" $t
    return $html
}

# ============================================================
# POINT D ENTREE
# ============================================================
Clear-Host
Write-Host ""
Write-Host "  +-----------------------------------------+" -ForegroundColor Cyan
Write-Host "  |  ____  ____      _   _  ____ ____  _   _ |" -ForegroundColor DarkYellow
Write-Host "  | |  _ \/ ___|    | \ | |/ ___|  _ \| | | ||" -ForegroundColor DarkYellow
Write-Host "  | | |_) \___ \    |  \| | |   | | | | | | ||" -ForegroundColor DarkYellow
Write-Host "  | |  __/ ___) |   | |\  | |___| |_| | |_| ||" -ForegroundColor DarkYellow
Write-Host "  | |_|   |____/    |_| \_|\____|____/ \___/ |" -ForegroundColor DarkYellow
Write-Host "  +-----------------------------------------+" -ForegroundColor Cyan
Write-Host "  | Disk Usage Analyzer          v$SCRIPT_VERSION       |" -ForegroundColor Cyan
Write-Host "  | $SCRIPT_DATE                              |" -ForegroundColor DarkGray
Write-Host "  | $SCRIPT_AUTHOR                    |" -ForegroundColor DarkGray
Write-Host "  | $USER_EMAIL            |" -ForegroundColor DarkGray
Write-Host "  +-----------------------------------------+" -ForegroundColor Cyan
Write-Host ""

$fullUserName = Get-FullUserName
Write-Log "FullUserName : $fullUserName"
Write-Host "  Utilisateur : $fullUserName ($($env:USERNAME))" -ForegroundColor Cyan
Write-Host ""

# ✅ v3.6 - Detection et proposition OneDrive
$oneDrivePaths = Get-OneDrivePaths
if ($oneDrivePaths.Count -gt 0) {
    Write-Host "  ☁️  ONEDRIVE DETECTE :" -ForegroundColor Cyan
    $idx = 1
    foreach ($od in $oneDrivePaths) {
        Write-Host "  [$idx] $($od.Icon) $($od.Type)" -ForegroundColor White
        Write-Host "       $($od.Path)" -ForegroundColor DarkGray
        $idx++
    }
    Write-Host ""
}

Write-Host "  MODES DE SCAN :" -ForegroundColor White
Write-Host ""
foreach ($k in ($SCAN_MODES.Keys | Sort-Object)) {
    $m=$SCAN_MODES[$k]; $sc=switch($k){1{"Yellow"}2{"Yellow"}3{"Red"}default{"White"}}
    Write-Host "  [$k] " -NoNewline -ForegroundColor White
    Write-Host "$($m["Name"])" -NoNewline -ForegroundColor Cyan
    Write-Host " | $($m["Speed"])" -NoNewline -ForegroundColor $sc
    Write-Host " | $($m["Accuracy"])" -ForegroundColor DarkGray
}
Write-Host ""
Write-Host "  Mode (defaut=1) : " -NoNewline -ForegroundColor Yellow
$inputMode=Read-Host; $modeKey=1
if(-not[string]::IsNullOrWhiteSpace($inputMode)){try{$mk=[int]$inputMode;if($SCAN_MODES.ContainsKey($mk)){$modeKey=$mk}}catch{}}
$selectedMode=$SCAN_MODES[$modeKey]
Write-Host "  >> $($selectedMode["Name"])" -ForegroundColor Green
Write-Host ""

Write-Host "  Profondeur : 1-10 = limite | 0 = ILLIMITEE" -ForegroundColor DarkGray
Write-Host "  Profondeur (defaut=$DEFAULT_DEPTH) : " -NoNewline -ForegroundColor Yellow
$inputDepth=Read-Host; $maxDepth=$DEFAULT_DEPTH
if(-not[string]::IsNullOrWhiteSpace($inputDepth)){try{$p=[int]$inputDepth;if($p -ge 0 -and $p -le 10){$maxDepth=$p}}catch{}}

$isUnlimited=($maxDepth -eq 0)
if($isUnlimited){
    Write-Host ""; Write-Host "  !! ATTENTION : Profondeur ILLIMITEE !!" -ForegroundColor Yellow
    Write-Host "  Peut prendre 30-60 min sur C:\." -ForegroundColor Yellow
    Write-Host "  Continuer ? (O/N, defaut=O) : " -NoNewline -ForegroundColor Yellow
    $confirm=Read-Host
    if($confirm -match '^[Nn]$'){Write-Host "  Annule." -ForegroundColor Red;exit 0}
}

# ✅ v3.6 - Chemin avec suggestion OneDrive
if ($oneDrivePaths.Count -gt 0) {
    Write-Host "  Chemin (ENTER=$DEFAULT_PATH, od1/od2...=OneDrive) : " -NoNewline -ForegroundColor Yellow
} else {
    Write-Host "  Chemin (ENTER = $DEFAULT_PATH) : " -NoNewline -ForegroundColor Yellow
}
$inputPath=Read-Host; $inputPath=$inputPath.Trim()

# Raccourcis OneDrive od1, od2...
$isOneDriveScan = $false
if ($inputPath -match '^od(\d+)$' -and $oneDrivePaths.Count -gt 0) {
    $odIdx = [int]$Matches[1] - 1
    if ($odIdx -ge 0 -and $odIdx -lt $oneDrivePaths.Count) {
        $inputPath = $oneDrivePaths[$odIdx].Path
        $isOneDriveScan = $true
        Write-Host "  >> OneDrive selectionne : $inputPath" -ForegroundColor Cyan
    }
}

if([string]::IsNullOrWhiteSpace($inputPath)){
    $startPath=$DEFAULT_PATH
} else {
    $inputPath=$inputPath -replace '/','\' 
    if($inputPath -match '^[A-Za-z]:$'){$inputPath+='\'}

    # Detecter automatiquement si c est un chemin OneDrive
    if(-not $isOneDriveScan){
        foreach($od in $oneDrivePaths){
            if($inputPath.TrimEnd('\').ToLower() -eq $od.Path.TrimEnd('\').ToLower() -or
               $inputPath.ToLower().StartsWith($od.Path.TrimEnd('\').ToLower()+'\')){
                $isOneDriveScan = $true
                Write-Host "  ☁️  Chemin OneDrive detecte automatiquement" -ForegroundColor Cyan
                break
            }
        }
        # Detection par nom de dossier
        if(-not $isOneDriveScan -and $inputPath -match 'OneDrive'){
            $isOneDriveScan = $true
            Write-Host "  ☁️  Chemin OneDrive detecte (nom)" -ForegroundColor Cyan
        }
    }

    Write-Host "  Verification du chemin..." -ForegroundColor Gray
    $pathOk = $false

    if($inputPath -match '^\\\\'){
        Write-Host "  Type : UNC reseau" -ForegroundColor Cyan
        try { if(Test-Path -Path $inputPath -ErrorAction Stop){ $pathOk=$true; Write-Host "  OK (Test-Path)." -ForegroundColor Green } } catch {}
        if(-not $pathOk){
            try { $di=New-Object System.IO.DirectoryInfo($inputPath); $null=$di.GetDirectories(); $pathOk=$true; Write-Host "  OK (DirectoryInfo)." -ForegroundColor Green } catch {}
        }
        if(-not $pathOk){
            try {
                $prevEnc=[Console]::OutputEncoding; [Console]::OutputEncoding=[System.Text.Encoding]::UTF8
                $cmdOut=& cmd /c "dir `"$inputPath`" 2>nul"; [Console]::OutputEncoding=$prevEnc
                if($null -ne $cmdOut -and $cmdOut.Count -gt 0){ $pathOk=$true; Write-Host "  OK (cmd dir)." -ForegroundColor Green }
            } catch {}
        }
        if(-not $pathOk){
            Write-Host "  Chemin UNC inaccessible." -ForegroundColor Red
            Write-Host "  Utilisation de $DEFAULT_PATH" -ForegroundColor Yellow
            $startPath=$DEFAULT_PATH
        } else { $startPath=$inputPath }
    } else {
        Write-Host "  Type : Local / Lecteur mappe$(if($isOneDriveScan){' / OneDrive'})" -ForegroundColor Cyan
        if(Test-Path $inputPath -ErrorAction SilentlyContinue){
            $pathOk=$true; Write-Host "  OK." -ForegroundColor Green; $startPath=$inputPath
        } else {
            Write-Host "  Chemin invalide." -ForegroundColor Red
            Write-Host "  Utilisation de $DEFAULT_PATH" -ForegroundColor Yellow
            $startPath=$DEFAULT_PATH; $isOneDriveScan=$false
        }
    }
}

$startPath    = Normalize-Path $startPath
$depthDisplay = if($isUnlimited){"ILLIMITEE"}else{"$maxDepth niveaux"}
$scanDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

Write-Log "Lancement v$SCRIPT_VERSION : '$startPath' mode=$modeKey profondeur=$depthDisplay OneDrive=$isOneDriveScan"
Write-Host ""
Write-Host "  Chemin     : $startPath" -ForegroundColor Cyan
Write-Host "  Type       : $(if($startPath -match '^\\\\'){'UNC Reseau'}elseif($isOneDriveScan){'OneDrive'}else{'Local/Mappe'})" -ForegroundColor Cyan
if($isOneDriveScan){ Write-Host "  ☁️  Mode OneDrive : fichiers cloud-only exclus du calcul" -ForegroundColor Cyan }
Write-Host "  Mode       : $($selectedMode["Name"])"     -ForegroundColor Cyan
Write-Host "  Precision  : $($selectedMode["Accuracy"])" -ForegroundColor Yellow
Write-Host "  Profondeur : $depthDisplay" -ForegroundColor $(if($isUnlimited){"Yellow"}else{"Cyan"})
Write-Host "  Date scan  : $scanDateTime" -ForegroundColor Cyan
Write-Host ""

try {
    $globalStart=Get-Date
    $result=Start-FastScan -RootPath $startPath -MaxDepth $maxDepth -Mode $selectedMode -IsOneDrive $isOneDriveScan
    $allScans        = $result["Scans"]
    $excluded        = $result["Excluded"]
    $unscanned       = $result["Unscanned"]
    $accessDenied    = $result["AccessDenied"]
    $elapsed         = $result["ElapsedSec"]
    $pathType        = $result["PathType"]
    $unlimited       = $result["Unlimited"]
    $methodStats     = $result["MethodStats"]
    $junctionsSkipped= $result["JunctionsSkipped"]
    $cloudOnlyCount  = $result["CloudOnlyCount"]
    $localSize       = $result["LocalSize"]

    Write-Host "  Scan : ${elapsed}s - $($allScans.Count) dossiers | $($unscanned.Count) non scannes | $($accessDenied.Count) proteges | $junctionsSkipped junctions$(if($cloudOnlyCount -gt 0){" | ☁️ $cloudOnlyCount cloud-only"})" -ForegroundColor Green
    Write-Host "  Generation HTML..." -ForegroundColor Cyan

    $tHtml=Get-Date
    $html=Get-HtmlReport `
        -RootPath         $startPath `
        -AllScans         $allScans `
        -UnscannedDirs    $unscanned `
        -AccessDeniedDirs $accessDenied `
        -MaxDepth         $maxDepth `
        -Unlimited        $unlimited `
        -ModeName         $selectedMode["Name"] `
        -ModeAccuracy     $selectedMode["Accuracy"] `
        -ElapsedSec       $elapsed `
        -PathType         $pathType `
        -FullUserName     $fullUserName `
        -ScanDateTime     $scanDateTime `
        -MethodStats      $methodStats `
        -JunctionsSkipped $junctionsSkipped `
        -IsOneDrive       $isOneDriveScan `
        -CloudOnlyCount   $cloudOnlyCount `
        -LocalSize        $localSize

    Write-Log "[MAIN] Set-Content..."
    $html | Set-Content -Path $htmlFile -Encoding UTF8 -ErrorAction Stop
    Write-Log "[MAIN] Set-Content OK en $([int]((Get-Date)-$tHtml).TotalSeconds)s"

    $total=[int]((Get-Date)-$globalStart).TotalSeconds
    Write-Log "[MAIN] Total : ${total}s"

    Start-Process $htmlFile
    Write-Log "Navigateur ouvert - fin automatique"
    Write-Host "  Total : ${total}s | Rapport ouvert." -ForegroundColor Green
    Write-Host "  Fichier : $htmlFile" -ForegroundColor DarkGray
    Write-Host ""
}
catch {
    Write-Progress -Activity "PS-NCDU v$SCRIPT_VERSION" -Completed
    Write-Log "ERREUR : $_" -Level ERROR
    Write-Log "Ligne  : $($_.InvocationInfo.ScriptLineNumber)" -Level ERROR
    Write-Host "  ERREUR : $_" -ForegroundColor Red
    Write-Host "  Log    : $logFile" -ForegroundColor Yellow
    Write-Host "  Appuyez sur ENTER..." -ForegroundColor DarkGray
    $null = Read-Host
}

exit 0
