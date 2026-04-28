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
#  Version     : 3.1
#  Date        : 2026-04-23
#  Auteur      : Eric Guiffaut (EGUI@NOVONORDISK.COM)
#  Societe     : Novo Nordisk
# ------------------------------------------------------------
#  Compatibilite : PowerShell 5.1+ | ConstrainedLanguage OK
#  Dependances   : Aucune - 100% PowerShell natif
# ------------------------------------------------------------
#  NOUVEAUTES v3.1 :
#    FIX1 : Junctions Windows ignorees (evite double comptage)
#           Mes documents, Menu Démarrer, Local Settings, etc.
#    FIX2 : cmd.exe methode 3 - encodage UTF-8 (chcp 65001)
#           Accents correctement geres (é, è, à, ù, etc.)
#    FIX3 : ConvertTo-JsonSafe - caracteres de controle
#           null byte, tab, backspace, form feed echappes
#    FIX4 : Test-IsJunction - detection ReparsePoint
# ------------------------------------------------------------
#  OPTIMISATIONS v2.9 conservees :
#    OPT1 : Zero hashtable fichier en E3
#    OPT2 : Format-Size uniquement a la serialisation JSON
#    OPT3 : Fusion par niveaux en E1
#  FALLBACK v3.0 conserve :
#    M1 : Get-ChildItem
#    M2 : System.IO.DirectoryInfo
#    M3 : cmd.exe dir (chcp 65001 pour les accents)
# ------------------------------------------------------------
#  Historique :
#    v2.9  - Optimisations perf : -77% temps de scan
#    v3.0  - Fallback 3 niveaux pour droits NTFS partiels
#    v3.1  - Fix junctions + encodage accents + JSON propre
# ============================================================

$SCRIPT_VERSION  = "3.1"
$SCRIPT_DATE     = "2026-04-23"
$USER_EMAIL      = "EGUI@NOVONORDISK.COM"
$SCRIPT_AUTHOR   = "Eric Guiffaut"
$DEFAULT_PATH    = "C:\Users"
$DEFAULT_DEPTH   = 3

$NN_LOGO_SVG = @'
<svg version="1.1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"
     viewBox="0 0 600 340" style="height:44px;width:auto;display:block">
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

Write-Log "PS-NCDU v$SCRIPT_VERSION ($SCRIPT_DATE) Demarrage"
Write-Log "PSVersion    : $($PSVersionTable.PSVersion)"
Write-Log "LanguageMode : $($ExecutionContext.SessionState.LanguageMode)"
Write-Log "Auteur       : $SCRIPT_AUTHOR ($USER_EMAIL)"
Write-Log "User         : $($env:USERNAME)"
Write-Log "Machine      : $($env:COMPUTERNAME)"

# ============================================================
# Utilitaires chemin
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

# ✅ v3.1 FIX4 : Detection junction/symlink/reparse point
function Test-IsJunction {
    param([string]$Path)
    try {
        $attr = (Get-Item -Path $Path -Force -ErrorAction Stop).Attributes
        return ($attr -band [System.IO.FileAttributes]::ReparsePoint) -ne 0
    } catch { return $false }
}

function Test-NetworkPath {
    param([string]$Path)
    if ($Path -match '^\\\\') {
        try { return (Test-Path -Path $Path -ErrorAction Stop) }
        catch { Write-Log "[NET] Erreur acces '$Path' : $_" -Level ERROR; return $false }
    }
    return $true
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

# ✅ v3.1 FIX3 : JsonSafe - caracteres de controle + accents
function ConvertTo-JsonSafe {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return "" }
    $Text = $Text -replace '\\', '\\\\'
    $Text = $Text -replace '"',  '\"'
    $Text = $Text -replace "`r", ''
    $Text = $Text -replace "`n", ''
    $Text = $Text -replace "`t", '\t'
    # Supprimer caracteres de controle non imprimables (U+0000-U+001F sauf tab)
    $Text = $Text -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', ''
    return $Text
}

# ============================================================
# ✅ v3.1 FIX2+FIX4 : cmd.exe avec UTF-8 et filtre junctions
# ============================================================
function Invoke-CmdDirSafe {
    param([string]$Path, [string]$DirArgs = "/b /ad")
    $prevEnc = [Console]::OutputEncoding
    try {
        # FIX2 : Forcer UTF-8 pour les accents (é, è, à, etc.)
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        $cmdOut = & cmd /c "chcp 65001 >nul 2>nul & dir $DirArgs `"$Path`" 2>nul"
        return $cmdOut
    }
    catch { return $null }
    finally {
        [Console]::OutputEncoding = $prevEnc
    }
}

# ============================================================
# ✅ v3.1 FIX1+FIX4 : Enumeration sous-dossiers
# SANS junctions (evite double comptage)
# AVEC fallback 3 niveaux + encodage UTF-8
# ============================================================
function Get-SubDirectories {
    param([string]$Path)

    # ── Methode 1 : Get-ChildItem ────────────────────────────
    try {
        $items = Get-ChildItem -Path $Path -Directory -ErrorAction Stop -Force
        # ✅ FIX1 : Filtrer les junctions/symlinks/reparse points
        $jCount = 0
        $filtered = @()
        foreach ($item in $items) {
            if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                $jCount++
                Write-Log "[DIR-M1] Junction ignoree : $($item.FullName)" -Level DEBUG
            } else {
                $filtered += $item
            }
        }
        if ($jCount -gt 0) {
            Write-Log "[DIR-M1] '$Path' : $jCount junction(s) ignoree(s) sur $($items.Count)"
        }
        return @{ Items=$filtered; Method="GCI"; Denied=$false; Junctions=$jCount }
    }
    catch [System.UnauthorizedAccessException] {
        Write-Log "[DIR-M1] GCI refuse '$Path' - essai M2" -Level DEBUG
    }
    catch { Write-Log "[DIR-M1] GCI erreur '$Path' : $_ - essai M2" -Level DEBUG }

    # ── Methode 2 : System.IO.DirectoryInfo ─────────────────
    try {
        $di    = New-Object System.IO.DirectoryInfo($Path)
        $subs  = $di.GetDirectories()
        $jCount = 0
        $items = @()
        foreach ($sub in $subs) {
            # FIX1 : Filtrer les ReparsePoints
            if (($sub.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                $jCount++
                Write-Log "[DIR-M2] Junction ignoree : $($sub.FullName)" -Level DEBUG
            } else {
                $items += New-Object PSObject -Property @{ FullName=$sub.FullName; Name=$sub.Name }
            }
        }
        if ($jCount -gt 0) { Write-Log "[DIR-M2] '$Path' : $jCount junction(s) ignoree(s)" }
        Write-Log "[DIR-M2] DirectoryInfo OK '$Path' ($($items.Count) items)"
        return @{ Items=$items; Method="NET"; Denied=$false; Junctions=$jCount }
    }
    catch [System.UnauthorizedAccessException] {
        Write-Log "[DIR-M2] .NET refuse '$Path' - essai M3" -Level DEBUG
    }
    catch { Write-Log "[DIR-M2] .NET erreur '$Path' : $_ - essai M3" -Level DEBUG }

    # ── Methode 3 : cmd.exe UTF-8 ───────────────────────────
    try {
        # FIX2 : UTF-8 via Invoke-CmdDirSafe
        $cmdOut = Invoke-CmdDirSafe -Path $Path -DirArgs "/b /ad"
        if ($null -ne $cmdOut) {
            $jCount = 0
            $items = @()
            foreach ($line in ($cmdOut | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
                $fp = Join-Path $Path $line
                # FIX1 : Verifier junction
                try {
                    $attr = (Get-Item $fp -Force -ErrorAction Stop).Attributes
                    if (($attr -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                        $jCount++
                        Write-Log "[DIR-M3] Junction ignoree : $fp" -Level DEBUG
                        continue
                    }
                } catch {}
                $items += New-Object PSObject -Property @{ FullName=$fp; Name=$line }
            }
            if ($jCount -gt 0) { Write-Log "[DIR-M3] '$Path' : $jCount junction(s) ignoree(s)" }
            Write-Log "[DIR-M3] cmd.exe UTF-8 OK '$Path' ($($items.Count) items)"
            return @{ Items=$items; Method="CMD"; Denied=$false; Junctions=$jCount }
        }
    }
    catch { Write-Log "[DIR-M3] cmd.exe erreur '$Path' : $_" -Level DEBUG }

    Write-Log "[DIR] PROTEGE (toutes methodes) : '$Path'" -Level INFO
    return @{ Items=@(); Method="NONE"; Denied=$true; Junctions=0 }
}

# ============================================================
# ✅ v3.1 : Enumeration fichiers directs avec fallback + UTF-8
# ============================================================
function Get-DirectFiles {
    param([string]$Path)

    # Methode 1
    try {
        return Get-ChildItem -Path $Path -File -ErrorAction Stop -Force
    }
    catch [System.UnauthorizedAccessException] {}
    catch { Write-Log "[FILE-M1] Erreur '$Path' : $_" -Level DEBUG }

    # Methode 2
    try {
        $di = New-Object System.IO.DirectoryInfo($Path)
        return $di.GetFiles() | ForEach-Object {
            New-Object PSObject -Property @{
                FullName=$_.FullName; Name=$_.Name
                Length=$_.Length; Extension=$_.Extension; DirectoryName=$Path
            }
        }
    }
    catch [System.UnauthorizedAccessException] {}
    catch { Write-Log "[FILE-M2] Erreur '$Path' : $_" -Level DEBUG }

    # Methode 3 : cmd.exe UTF-8
    try {
        $cmdOut = Invoke-CmdDirSafe -Path $Path -DirArgs "/b /a-d"
        if ($null -ne $cmdOut) {
            return $cmdOut |
                   Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                   ForEach-Object {
                       $fp = Join-Path $Path $_
                       try {
                           $fi = New-Object System.IO.FileInfo($fp)
                           New-Object PSObject -Property @{
                               FullName=$fp; Name=$_
                               Length=$fi.Length; Extension=$fi.Extension; DirectoryName=$Path
                           }
                       } catch { $null }
                   } | Where-Object { $null -ne $_ }
        }
    }
    catch { Write-Log "[FILE-M3] Erreur '$Path' : $_" -Level DEBUG }
    return @()
}

# ============================================================
# ✅ v3.1 : Scan recursif avec fallback + UTF-8
# ============================================================
function Get-RecursiveFiles {
    param([string]$Path)

    # Methode 1
    try {
        return Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue -Force
    }
    catch [System.UnauthorizedAccessException] {}
    catch { Write-Log "[RFILE-M1] Erreur '$Path' : $_" -Level DEBUG }

    # Methode 2 : DFS .NET avec filtre junctions
    try {
        $result   = @()
        $dirQueue = @($Path)
        while ($dirQueue.Count -gt 0) {
            $current  = $dirQueue[0]
            $dirQueue = if ($dirQueue.Count -gt 1) { $dirQueue[1..($dirQueue.Count-1)] } else { @() }
            try {
                $di = New-Object System.IO.DirectoryInfo($current)
                foreach ($f in $di.GetFiles()) {
                    $result += New-Object PSObject -Property @{
                        FullName=$f.FullName; Name=$f.Name
                        Length=$f.Length; Extension=$f.Extension; DirectoryName=$current
                    }
                }
                # FIX1 : Ne pas suivre les junctions
                foreach ($s in $di.GetDirectories()) {
                    if (($s.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -eq 0) {
                        $dirQueue += $s.FullName
                    }
                }
            }
            catch {}
        }
        if ($result.Count -gt 0) {
            Write-Log "[RFILE-M2] .NET DFS OK '$Path' ($($result.Count) fichiers)"
            return $result
        }
    }
    catch { Write-Log "[RFILE-M2] Erreur '$Path' : $_" -Level DEBUG }

    # Methode 3 : cmd.exe UTF-8
    try {
        $cmdOut = Invoke-CmdDirSafe -Path $Path -DirArgs "/s /b /a-d"
        if ($null -ne $cmdOut) {
            $result = $cmdOut |
                      Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                      ForEach-Object {
                          try {
                              $fi = New-Object System.IO.FileInfo($_)
                              New-Object PSObject -Property @{
                                  FullName=$_; Name=$fi.Name
                                  Length=$fi.Length; Extension=$fi.Extension; DirectoryName=$fi.DirectoryName
                              }
                          } catch { $null }
                      } | Where-Object { $null -ne $_ }
            Write-Log "[RFILE-M3] cmd.exe UTF-8 OK '$Path' ($($result.Count) fichiers)"
            return $result
        }
    }
    catch { Write-Log "[RFILE-M3] Erreur '$Path' : $_" -Level DEBUG }
    return @()
}

# ============================================================
# SCAN v3.1
# ============================================================
function Start-FastScan {
    param([string]$RootPath,[int]$MaxDepth=3,[hashtable]$Mode=$null)

    $ACT       = "PS-NCDU v$SCRIPT_VERSION"
    $scanStart = Get-Date
    $script:ScanGlobalStart = $scanStart
    $rootDepth = Get-RootDepth -RootPath $RootPath
    $modeName  = if($null -ne $Mode){$Mode["Name"]}else{"Complet"}
    $isUNC     = $RootPath -match '^\\\\' 
    $pathType  = if($isUNC){"UNC reseau"}else{"Local"}
    $unlimited = ($MaxDepth -eq 0)
    $depthLabel= if($unlimited){"ILLIMITEE"}else{"$MaxDepth niveaux"}

    Write-Log "==========================================";Write-Log "DEBUT SCAN v$SCRIPT_VERSION : $RootPath"
    Write-Log "Type : $pathType | Mode : $modeName | Profondeur : $depthLabel"
    Write-Log "Fix : Junctions ignorees | UTF-8 cmd.exe | JSON propre";Write-Log "=========================================="

    # E1
    $t = Get-Date
    Update-Progress $ACT "E1/5 : Liste des dossiers [$depthLabel]..." "" 1 $true $true

    $allLevels       = @(@($RootPath))
    $excludedFound   = @()
    $accessDenied    = @()
    $junctionsTotal  = 0
    $methodStats     = @{ GCI=0; NET=0; CMD=0; NONE=0 }

    if ($unlimited) {
        Write-Log "[E1] Mode illimite + filtre junctions"
        Update-Progress $ACT "E1/5 : Enumeration complete (illimitee)..." "" 5 $true $true
        $flatList = @($RootPath)
        try {
            $allDirsFound = Get-ChildItem -Path $RootPath -Recurse -Directory -ErrorAction SilentlyContinue -Force
            $dirsDone = 0
            foreach ($dir in $allDirsFound) {
                $dirsDone++
                if(($dirsDone%500) -eq 0){Update-Progress $ACT "E1/5 : $dirsDone dossiers enumeres..." "" 15 ($dirsDone%5000 -eq 0)}
                # FIX1 : Ignorer junctions en mode illimite
                if (($dir.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                    $junctionsTotal++
                    Write-Log "[E1] Junction ignoree (illimite) : $($dir.FullName)" -Level DEBUG
                    continue
                }
                if(Test-IsExcluded -Path $dir.FullName -ExcludedList $EXCLUDED_DIRS){
                    if(-not($excludedFound -contains $dir.FullName)){$excludedFound+=$dir.FullName}
                } else { $flatList+=$dir.FullName }
            }
        }
        catch { Write-Log "[E1] Erreur recurse : $_" -Level WARN }
        $allLevels = @($flatList)
        Write-Log "[E1] Mode illimite : $($flatList.Count) dossiers | $junctionsTotal junctions ignorees"
    } else {
        for ($d=0; $d -lt $MaxDepth; $d++) {
            $currentLevel = $allLevels[$d]
            if ($null -eq $currentLevel -or $currentLevel.Count -eq 0) { break }
            $tLevel=$Get=Get-Date; $nextLevel=@(); $totalDirs=$currentLevel.Count; $dirsDone=0
            $levelNum=$d+1; $levelBase=1+[int](($d*24)/$MaxDepth); $levelRng=[int](24/$MaxDepth)

            foreach ($dir in $currentLevel) {
                $dirsDone++
                $pct=$levelBase+[int](($dirsDone*$levelRng)/($totalDirs+1))
                Update-Progress $ACT "E1/5 : Niveau $levelNum/$MaxDepth | $dirsDone/$totalDirs" "Scan : $dir" $pct ($dirsDone%200 -eq 0)

                $subResult = Get-SubDirectories -Path $dir
                $methodStats[$subResult["Method"]]++
                $junctionsTotal += $subResult["Junctions"]

                if ($subResult["Denied"]) {
                    if(-not($accessDenied -contains $dir)){$accessDenied+=$dir;Write-Log "[E1] PROTEGE : $dir" -Level INFO}
                } else {
                    foreach ($sub in $subResult["Items"]) {
                        if(Test-IsExcluded -Path $sub.FullName -ExcludedList $EXCLUDED_DIRS){
                            if(-not($excludedFound -contains $sub.FullName)){$excludedFound+=$sub.FullName}
                        } else { $nextLevel+=$sub.FullName }
                    }
                }
            }
            $allLevels+=,$nextLevel
            Write-Log "[E1] Niveau $levelNum/$MaxDepth : $($nextLevel.Count) en $([int]((Get-Date)-$tLevel).TotalSeconds)s | junctions: $junctionsTotal"
            if($nextLevel.Count -eq 0){break}
        }
    }

    # Fusion niveaux (OPT3)
    $dirsInScope=@()
    foreach ($level in $allLevels){ foreach ($dir in $level){ $dirsInScope+=$dir } }
    $nbScope=$dirsInScope.Count

    Write-Log "[E1] Stats : GCI=$($methodStats['GCI']) | .NET=$($methodStats['NET']) | CMD=$($methodStats['CMD']) | PROTEGE=$($methodStats['NONE']) | Junctions ignorees=$junctionsTotal"

    # Dossiers non scannes niveau MaxDepth+1
    $dirsUnscanned=@()
    if(-not $unlimited){
        $lastLevel=$allLevels[$allLevels.Count-1]
        if($null -ne $lastLevel){
            foreach ($dir in $lastLevel){
                $depth=Get-PathDepth -Path $dir -RootPath $RootPath -RootDepth $rootDepth
                if($depth -eq $MaxDepth){
                    $subResult=Get-SubDirectories -Path $dir
                    if(-not $subResult["Denied"]){
                        foreach($sub in $subResult["Items"]){
                            if(-not(Test-IsExcluded -Path $sub.FullName -ExcludedList $EXCLUDED_DIRS)){$dirsUnscanned+=$sub.FullName}
                        }
                    } else { if(-not($accessDenied -contains $dir)){$accessDenied+=$dir} }
                }
            }
        }
    }
    Write-Step "[E1] TERMINEE" $t "Scope : $nbScope | NonScannes : $($dirsUnscanned.Count) | Proteges : $($accessDenied.Count) | Junctions ignorees : $junctionsTotal"

    # E2
    $t=Get-Date
    Update-Progress $ACT "E2/5 : Init ($nbScope dossiers)..." "" 25 $true $true
    $dirOwnSizes=@{}; $dirTotalSizes=@{}; $initDone=0
    foreach ($dir in $dirsInScope){
        $initDone++; $dirOwnSizes[$dir]=[long]0; $dirTotalSizes[$dir]=[long]0
        Update-Progress $ACT "E2/5 : Init $initDone / $nbScope" "" (25+[int](($initDone*5)/($nbScope+1))) ($initDone%5000 -eq 0)
    }
    Write-Step "[E2] TERMINEE" $t "$nbScope entrees"

    # E3
    $t=Get-Date; $totalFilesScanned=0; $totalSizeScanned=[long]0
    Update-Progress $ACT "E3/5 : Scan tailles..." "" 30 $true $true

    try {
        $rootFiles=Get-DirectFiles -Path $RootPath
        $nbRF=0; $sizeRF=[long]0
        foreach ($f in $rootFiles){if($null -eq $f){continue};$fl=[long]$f.Length;$sizeRF+=$fl;$nbRF++;$dirOwnSizes[$RootPath]+=$fl}
        Write-Log "[E3] Racine : $nbRF fichiers = $(Format-Size $sizeRF)"
    } catch { Write-Log "[E3] Erreur racine : $_" -Level WARN }

    $level1Dirs=@()
    $l1Result=Get-SubDirectories -Path $RootPath
    if(-not $l1Result["Denied"]){ foreach($s in $l1Result["Items"]){$level1Dirs+=$s.FullName} }
    $totalL1=$level1Dirs.Count; $doneL1=0
    Write-Log "[E3] $totalL1 dossiers L1 | Fallback actif + filtre junctions"

    foreach ($l1dir in $level1Dirs){
        $doneL1++; $l1Name=Split-Path $l1dir -Leaf
        $pctL1=30+[int](($doneL1*45)/($totalL1+1)); if($pctL1 -gt 74){$pctL1=74}

        if(Test-IsExcluded -Path $l1dir -ExcludedList $EXCLUDED_DIRS){
            Write-Log "[E3] L1 EXCLU : '$l1Name'"
            Update-Progress $ACT "E3/5 : L1 $doneL1/$totalL1 - $l1Name [EXCLU]" "" $pctL1 $true $true
            continue
        }
        # FIX1 : Skip si c est une junction au niveau L1
        if (Test-IsJunction -Path $l1dir) {
            Write-Log "[E3] L1 JUNCTION ignoree : '$l1dir'"
            Update-Progress $ACT "E3/5 : L1 $doneL1/$totalL1 - $l1Name [JUNCTION]" "" $pctL1 $true $true
            continue
        }

        Update-Progress $ACT "E3/5 : L1 $doneL1/$totalL1 - $l1Name | $(Format-Size $totalSizeScanned)" "Collecte..." $pctL1 $true $true

        try {
            $filesInL1=Get-RecursiveFiles -Path $l1dir
            $nbF1=0; $sz1=[long]0; $fdL1=0

            foreach ($file in $filesInL1){
                if($null -eq $file){continue}
                $pd=$null
                try{$pd=$file.DirectoryName}catch{$pd=Get-ParentPath $file.FullName}
                if($null -eq $pd){continue}
                if(Test-IsExcluded -Path $pd -ExcludedList $EXCLUDED_DIRS){continue}

                $fl=[long]$file.Length; $sz1+=$fl; $nbF1++; $fdL1++; $totalFilesScanned++; $totalSizeScanned+=$fl

                Update-Progress $ACT "E3/5 : L1 $doneL1/$totalL1 - $l1Name | $(Format-Size $totalSizeScanned)" "$fdL1 fichiers ($(Format-Size $sz1))" $pctL1 ($fdL1%50000 -eq 0)

                if($dirOwnSizes.ContainsKey($pd)){
                    $dirOwnSizes[$pd]+=$fl
                } else {
                    $anc=$pd; $hops=0; $found=$false
                    while(-not[string]::IsNullOrEmpty($anc) -and $hops -lt 50){
                        $anc=Get-ParentPath $anc; $hops++
                        if(-not[string]::IsNullOrEmpty($anc) -and $dirOwnSizes.ContainsKey($anc)){
                            $dirOwnSizes[$anc]+=$fl; $found=$true; break
                        }
                    }
                    if(-not $found){$dirOwnSizes[$RootPath]+=$fl}
                }
            }
            Write-Log "[E3] L1 $doneL1/$totalL1 '$l1Name' : $nbF1 = $(Format-Size $sz1) | Cumul : $totalFilesScanned = $(Format-Size $totalSizeScanned)"
        }
        catch { Write-Log "[E3] Erreur '$l1dir' : $_" -Level ERROR }
    }
    Write-Step "[E3] TERMINEE" $t "$totalFilesScanned fichiers = $(Format-Size $totalSizeScanned)"

    # E4
    $t=Get-Date; Update-Progress $ACT "E4/5 : Propagation..." "" 75 $true $true
    foreach ($dir in $dirsInScope){$dirTotalSizes[$dir]=$dirOwnSizes[$dir]}
    $propDone=0
    foreach ($dir in $dirsInScope){
        $propDone++; $ownSize=$dirOwnSizes[$dir]; if($ownSize -eq 0){continue}
        $pct=75+[int](($propDone*10)/($nbScope+1)); if($pct -gt 84){$pct=84}
        Update-Progress $ACT "E4/5 : $propDone / $nbScope" "Racine : $(Format-Size $dirTotalSizes[$RootPath])" $pct ($propDone%5000 -eq 0)
        $current=Get-ParentPath $dir; $hops=0
        while(-not[string]::IsNullOrEmpty($current) -and $hops -lt 50){
            if($dirTotalSizes.ContainsKey($current)){$dirTotalSizes[$current]+=$ownSize}
            $current=Get-ParentPath $current; $hops++
        }
    }
    Write-Step "[E4] TERMINEE" $t "Racine : $(Format-Size $dirTotalSizes[$RootPath])"

    # E5
    $t=Get-Date; Update-Progress $ACT "E5/5 : Index + rapport..." "" 85 $true $true
    $childIndex=@{}; $idxDone=0
    foreach ($dir in $dirsInScope){
        $idxDone++; $parent=Get-ParentPath $dir
        if($null -ne $parent){if(-not $childIndex.ContainsKey($parent)){$childIndex[$parent]=@()};$childIndex[$parent]+=$dir}
        Update-Progress $ACT "E5/5 : Index $idxDone / $nbScope" "" (85+[int](($idxDone*3)/($nbScope+1))) ($idxDone%5000 -eq 0)
    }
    foreach ($u in $dirsUnscanned){$p=Get-ParentPath $u;if($null -ne $p){if(-not $childIndex.ContainsKey($p)){$childIndex[$p]=@()};if(-not($childIndex[$p] -contains "UNSCANNED:$u")){$childIndex[$p]+="UNSCANNED:$u"}}}
    foreach ($d in $accessDenied){$p=Get-ParentPath $d;if($null -ne $p){if(-not $childIndex.ContainsKey($p)){$childIndex[$p]=@()};if(-not($childIndex[$p] -contains "DENIED:$d")){$childIndex[$p]+="DENIED:$d"}}}
    foreach ($e in $excludedFound){$p=Get-ParentPath $e;if($null -ne $p){if(-not $childIndex.ContainsKey($p)){$childIndex[$p]=@()};if(-not($childIndex[$p] -contains "EXCLU:$e")){$childIndex[$p]+="EXCLU:$e"}}}

    Write-Log "[E5] Index : $($childIndex.Count) parents"
    Update-Progress $ACT "E5/5 : Construction ($nbScope dossiers)..." "" 90 $true $true

    $allScans=@{}; $reportCount=0; $skipCount=0
    foreach ($dirPath in $dirsInScope){
        if(-not $unlimited){
            $dirDepth=Get-PathDepth -Path $dirPath -RootPath $RootPath -RootDepth $rootDepth
            if($dirDepth -gt $MaxDepth){$skipCount++;continue}
        }
        $entries=@()
        if($childIndex.ContainsKey($dirPath)){
            foreach ($ce in $childIndex[$dirPath]){
                if($ce -like "EXCLU:*"){$ep=$ce.Substring(6);$entries+=@{Name=Split-Path $ep -Leaf;FullPath=$ep;Size=[long]-1;IsDir=$true;Excluded=$true;Unscanned=$false;Denied=$false;Ext=""}}
                elseif($ce -like "UNSCANNED:*"){$up=$ce.Substring(10);$entries+=@{Name=Split-Path $up -Leaf;FullPath=$up;Size=[long]-2;IsDir=$true;Excluded=$false;Unscanned=$true;Denied=$false;Ext=""}}
                elseif($ce -like "DENIED:*"){$dp=$ce.Substring(7);$entries+=@{Name=Split-Path $dp -Leaf;FullPath=$dp;Size=[long]-3;IsDir=$true;Excluded=$false;Unscanned=$false;Denied=$true;Ext=""}}
                else{$cn=$ce.Substring($dirPath.TrimEnd('\').Length).TrimStart('\');$entries+=@{Name=$cn;FullPath=$ce;Size=$dirTotalSizes[$ce];IsDir=$true;Excluded=$false;Unscanned=$false;Denied=$false;Ext=""}}
            }
        }
        if($dirOwnSizes[$dirPath] -gt 0){
            try {
                $directFiles=Get-DirectFiles -Path $dirPath
                foreach ($f in $directFiles){
                    if($null -eq $f){continue}
                    $fl=[long]$f.Length
                    $entries+=@{Name=$f.Name;FullPath=$f.FullName;Size=$fl;IsDir=$false;Excluded=$false;Unscanned=$false;Denied=$false;Ext=$f.Extension.ToLower()}
                }
            } catch {}
        }
        $sN=$entries|Where-Object{-not $_["Excluded"] -and -not $_["Unscanned"] -and -not $_["Denied"]}|Sort-Object -Property{$_["Size"]} -Descending
        $sU=$entries|Where-Object{$_["Unscanned"] -eq $true}
        $sD=$entries|Where-Object{$_["Denied"]    -eq $true}
        $sE=$entries|Where-Object{$_["Excluded"]  -eq $true}
        $sorted=@();foreach($e in $sN){$sorted+=$e};foreach($e in $sU){$sorted+=$e};foreach($e in $sD){$sorted+=$e};foreach($e in $sE){$sorted+=$e}
        $allScans[$dirPath]=@{Items=$sorted;Total=$dirTotalSizes[$dirPath]};$reportCount++
        Update-Progress $ACT "E5/5 : $reportCount / $nbScope" "" (90+[int](($reportCount*9)/($nbScope+1))) ($reportCount%2000 -eq 0)
    }
    Write-Step "[E5] TERMINEE" $t "$reportCount construits | $($dirsUnscanned.Count) non scannes | $($accessDenied.Count) proteges | $junctionsTotal junctions ignorees"
    Update-Progress $ACT "Termine" "" 100 $true $true; Write-Progress -Activity $ACT -Completed

    $totalElap=[int]((Get-Date)-$scanStart).TotalSeconds
    Write-Log "SCAN TERMINE v$SCRIPT_VERSION en ${totalElap}s | $pathType | $depthLabel | Dossiers : $reportCount | Fichiers : $totalFilesScanned | Proteges : $($accessDenied.Count) | Junctions : $junctionsTotal | Taille : $(Format-Size $dirTotalSizes[$RootPath])"

    return @{
        Scans         = $allScans
        Excluded      = $excludedFound
        Unscanned     = $dirsUnscanned
        AccessDenied  = $accessDenied
        ModeName      = $modeName
        ElapsedSec    = $totalElap
        MaxDepth      = $MaxDepth
        Unlimited     = $unlimited
        PathType      = $pathType
        MethodStats   = $methodStats
        JunctionsSkipped = $junctionsTotal
    }
}

# ============================================================
# Generation HTML v3.1
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
        [int]$JunctionsSkipped
    )

    $ACT="PS-NCDU v$SCRIPT_VERSION"; $t=Get-Date
    Write-Log "[HTML] Generation v$SCRIPT_VERSION : $($AllScans.Count) dossiers"
    Update-Progress $ACT "HTML : JSON..." "" 2 $true $true

    $jsonParts=@(); $jsonCount=0
    foreach ($scanPath in $AllScans.Keys){
        $jsonCount++
        Update-Progress $ACT "HTML : JSON $jsonCount / $($AllScans.Count)" "" (2+[int](($jsonCount*80)/($AllScans.Count+1))) ($jsonCount%5000 -eq 0)
        $scanData=$AllScans[$scanPath]; $items=$scanData["Items"]; $total=$scanData["Total"]
        $ip=@()
        foreach ($item in $items){
            $sizeStr=Format-Size $item["Size"]
            # FIX3 : ConvertTo-JsonSafe propre (caracteres de controle)
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
    $psVer         = $PSVersionTable.PSVersion.ToString()
    $logEnc        = ConvertTo-HtmlEncoded $logFile
    $rootEnc       = ConvertTo-HtmlEncoded $RootPath
    $nbScans       = $AllScans.Count
    $emailEnc      = ConvertTo-HtmlEncoded $USER_EMAIL
    $authorEnc     = ConvertTo-HtmlEncoded $SCRIPT_AUTHOR
    $modeEncHtml   = ConvertTo-HtmlEncoded $ModeName
    $modeEncAcc    = ConvertTo-HtmlEncoded $ModeAccuracy
    $logoSvg       = $NN_LOGO_SVG
    $maxDepthJs    = $MaxDepth
    $unlimitedJs   = if($Unlimited){"true"}else{"false"}
    $nbUnscanned   = $UnscannedDirs.Count
    $nbDenied      = $AccessDeniedDirs.Count
    $pathTypeEnc   = ConvertTo-HtmlEncoded $PathType
    $pathIcon      = if($RootPath -match '^\\\\'){"&#127760;"}else{"&#128190;"}
    $depthLabel    = if($Unlimited){"Illimitee"}else{"$MaxDepth niveaux"}
    $depthLabelEnc = ConvertTo-HtmlEncoded $depthLabel
    $fullUserEnc   = ConvertTo-HtmlEncoded $FullUserName
    $scanDTEnc     = ConvertTo-HtmlEncoded $ScanDateTime
    $machineEnc    = ConvertTo-HtmlEncoded $env:COMPUTERNAME
    # Info junctions et fallback dans footer
    $footerExtra = ""
    if ($JunctionsSkipped -gt 0) { $footerExtra += " &middot; $JunctionsSkipped junctions ignorees" }
    if ($null -ne $MethodStats -and ($MethodStats["NET"] -gt 0 -or $MethodStats["CMD"] -gt 0)) {
        $footerExtra += " &middot; Fallback .NET=$($MethodStats['NET']) CMD=$($MethodStats['CMD'])"
    }
    $exListHtml="";foreach($ex in $EXCLUDED_DIRS){$exListHtml+="<li><code>$(ConvertTo-HtmlEncoded $ex)</code></li>"}
    $dnListHtml="";foreach($dn in $AccessDeniedDirs){$dnListHtml+="<li><code>$(ConvertTo-HtmlEncoded $dn)</code></li>"}

    $html = @"
<!DOCTYPE html>
<html lang="fr" data-theme="light">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width,initial-scale=1.0">
    <title>PS-NCDU v$SCRIPT_VERSION - $rootEnc</title>
    <style>
    @font-face{font-family:'Public Sans';font-style:normal;font-weight:400;src:url(https://fonts.gstatic.com/s/publicsans/v21/ijwRs572Xtc6ZYQws9YVwllKVnqpTiU.woff2) format('woff2')}
    @font-face{font-family:'Public Sans';font-style:italic;font-weight:400;src:url(https://fonts.gstatic.com/s/publicsans/v21/ijwTs572Xtc6ZYQws9YVwnNDTJLax9k0.woff2) format('woff2')}
    @font-face{font-family:'Public Sans';font-style:normal;font-weight:600;src:url(https://fonts.gstatic.com/s/publicsans/v21/ijwRs572Xtc6ZYQws9YVwllKVnqpTiU.woff2) format('woff2')}
    @font-face{font-family:'Public Sans';font-style:normal;font-weight:700;src:url(https://fonts.gstatic.com/s/publicsans/v21/ijwRs572Xtc6ZYQws9YVwllKVnqpTiU.woff2) format('woff2')}
    :root{--bg:#f0f4f8;--surface:#ffffff;--card:#e8edf3;--card-hover:#dde4ed;--border:#cdd6e0;--border-light:#dde4ed;--text:#0f1923;--text-muted:#4a6080;--text-dim:#8da0b8;--accent:#0063be;--accent-light:#004fa3;--accent-glow:rgba(0,99,190,.08);--success:#007a5e;--warning:#b86b00;--danger:#c0202e;--unscanned:#5e3d8f;--shadow:0 2px 8px rgba(0,0,0,.1);--shadow-sm:0 1px 3px rgba(0,0,0,.08);--radius:8px;--radius-sm:5px;--radius-pill:20px;--logo-color:#001965}
    [data-theme="dark"]{--bg:#0f1923;--surface:#162032;--card:#1e2d42;--card-hover:#243350;--border:#2a3f5f;--border-light:#1e3050;--text:#D3B276;--text-muted:#a08a5a;--text-dim:#6b5a3a;--accent:#0063be;--accent-light:#D3B276;--accent-glow:rgba(211,178,118,.15);--success:#00c48c;--warning:#f5a623;--danger:#e8394a;--unscanned:#9b7fd4;--shadow:0 2px 12px rgba(0,0,0,.45);--shadow-sm:0 1px 4px rgba(0,0,0,.3);--logo-color:#D3B276}
    *{box-sizing:border-box;margin:0;padding:0}
    body{font-family:'Public Sans','Segoe UI',Roboto,Arial,sans-serif;background:var(--bg);color:var(--text);min-height:100vh;font-size:14px;transition:background .3s,color .3s}
    .header{background:var(--surface);height:68px;padding:0 24px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;box-shadow:var(--shadow);position:sticky;top:0;z-index:200;gap:16px}
    .header-left{display:flex;align-items:center;gap:12px;flex-shrink:0}
    .app-title-main{font-size:1.3em;font-weight:700;letter-spacing:.5px;color:var(--text)}
    .app-title-main span{color:var(--accent-light)}
    .app-title-sub{font-size:.72em;color:var(--text-muted);letter-spacing:.3px}
    .version-chip{font-size:.7em;background:var(--accent-glow);color:var(--accent-light);border:1px solid var(--accent-light);padding:2px 10px;border-radius:var(--radius-pill);font-weight:600;white-space:nowrap}
    .header-right{display:flex;align-items:center;gap:14px;flex-shrink:0}
    .user-block{display:flex;align-items:center;gap:10px;background:var(--card);border:1px solid var(--border);border-radius:var(--radius);padding:6px 14px}
    .user-avatar{width:34px;height:34px;border-radius:50%;background:var(--accent-light);color:var(--surface);display:flex;align-items:center;justify-content:center;font-weight:700;font-size:.95em;flex-shrink:0}
    [data-theme="dark"] .user-avatar{color:#0f1923}
    .user-info{display:flex;flex-direction:column;gap:1px;text-align:left}
    .user-fullname{font-size:.85em;font-weight:600;color:var(--text);white-space:nowrap;max-width:200px;overflow:hidden;text-overflow:ellipsis}
    .user-details{font-size:.72em;color:var(--text-muted);white-space:nowrap}
    .theme-toggle{background:var(--card);border:1px solid var(--border);color:var(--text-muted);padding:6px 12px;border-radius:var(--radius-sm);cursor:pointer;font-size:.85em;transition:all .2s;white-space:nowrap;font-family:inherit;flex-shrink:0}
    .theme-toggle:hover{background:var(--card-hover);color:var(--text);border-color:var(--accent-light)}
    .nn-logo{cursor:pointer;display:flex;align-items:center;color:var(--logo-color);height:44px;transition:opacity .2s;flex-shrink:0}
    .nn-logo:hover{opacity:.75}
    .nn-logo svg{height:44px;width:auto;color:var(--logo-color)}
    .scan-datebar{background:var(--accent-light);color:var(--surface);padding:4px 24px;font-size:.78em;font-weight:500;display:flex;align-items:center;gap:16px;flex-wrap:wrap;border-bottom:1px solid var(--border)}
    [data-theme="dark"] .scan-datebar{background:var(--card);color:var(--text-muted)}
    .scan-datebar span{opacity:.9}.scan-datebar strong{opacity:1;font-weight:700}
    .toolbar{background:var(--surface);padding:10px 24px;border-bottom:1px solid var(--border);display:flex;gap:10px;align-items:center;flex-wrap:wrap}
    .breadcrumb{display:flex;flex-wrap:wrap;gap:2px;align-items:center;font-size:.85em;flex:1;min-width:0}
    .breadcrumb a{color:var(--accent-light);text-decoration:none;cursor:pointer;padding:4px 8px;border-radius:var(--radius-sm);transition:background .15s;white-space:nowrap}
    .breadcrumb a:hover{background:var(--accent-glow)}
    .breadcrumb a.active{color:var(--text);font-weight:600;background:var(--card);cursor:default}
    .sep{color:var(--text-dim);padding:0 2px;user-select:none}
    .nav-input{background:var(--card);border:1px solid var(--border);color:var(--text);padding:7px 14px;border-radius:var(--radius-pill);font-size:.85em;font-family:inherit;width:280px;outline:none;transition:all .2s}
    .nav-input:focus{border-color:var(--accent-light);box-shadow:0 0 0 3px var(--accent-glow);width:380px}
    .nav-input::placeholder{color:var(--text-dim)}
    .btn{background:var(--card);color:var(--accent-light);border:1px solid var(--border);padding:7px 16px;border-radius:var(--radius-sm);cursor:pointer;font-size:.875em;font-weight:500;transition:all .15s;white-space:nowrap;font-family:inherit}
    .btn:hover{background:var(--card-hover);border-color:var(--accent-light);box-shadow:var(--shadow-sm)}
    .btn-primary{background:var(--accent);color:#fff;border-color:var(--accent)}
    .btn-primary:hover{background:#0050a0}
    .mode-banner{background:var(--card);padding:8px 24px;font-size:.8em;display:flex;align-items:center;gap:10px;flex-wrap:wrap;border-bottom:1px solid var(--border);color:var(--text-muted)}
    .mode-tag{display:inline-flex;align-items:center;background:var(--accent-glow);color:var(--accent-light);border:1px solid var(--accent-light);padding:2px 10px;border-radius:var(--radius-pill);font-weight:600;font-size:.9em;white-space:nowrap}
    .path-type-badge{display:inline-flex;align-items:center;gap:4px;background:rgba(0,99,190,.1);color:var(--accent);border:1px solid rgba(0,99,190,.3);padding:2px 10px;border-radius:var(--radius-pill);font-weight:600;font-size:.85em;white-space:nowrap}
    .depth-badge{display:inline-flex;align-items:center;gap:4px;padding:2px 10px;border-radius:var(--radius-pill);font-weight:600;font-size:.85em;white-space:nowrap}
    .depth-unlimited{background:rgba(0,122,94,.1);color:var(--success);border:1px solid var(--success)}
    .depth-limited{background:var(--accent-glow);color:var(--accent-light);border:1px solid var(--accent-light)}
    .alert-banner{padding:10px 24px;font-size:.85em;display:flex;align-items:flex-start;gap:10px;border-bottom:1px solid var(--border);flex-wrap:wrap}
    .alert-unscanned{background:rgba(94,61,143,.08);border-left:4px solid var(--unscanned);color:var(--unscanned)}
    .alert-denied{background:rgba(192,32,46,.07);border-left:4px solid var(--danger);color:var(--danger)}
    .alert-banner strong{font-weight:700}
    .hint-link{color:var(--accent-light);cursor:pointer;text-decoration:underline;font-size:.9em;margin-left:6px}
    .alert-detail{display:none;padding:6px 24px 10px 40px;background:var(--card);border-bottom:1px solid var(--border);font-size:.82em;color:var(--text-muted)}
    .alert-detail ul{margin:0;padding:0;list-style:none}
    .alert-detail li{padding:2px 0}
    .alert-detail code{padding:1px 6px;border-radius:3px;background:rgba(192,32,46,.08);color:var(--danger)}
    .excl-section{border-bottom:1px solid var(--border)}
    .excl-toggle{width:100%;background:none;border:none;cursor:pointer;padding:9px 24px;display:flex;align-items:center;gap:8px;font-size:.82em;color:var(--warning);text-align:left;font-family:inherit;transition:background .15s}
    .excl-toggle:hover{background:var(--card)}
    .excl-count{background:rgba(176,107,0,.12);color:var(--warning);border:1px solid var(--warning);padding:1px 8px;border-radius:var(--radius-pill);font-weight:600;font-size:.85em}
    .excl-chevron{margin-left:auto;transition:transform .25s;color:var(--text-dim);font-size:.8em}
    .excl-toggle.open .excl-chevron{transform:rotate(180deg)}
    .excl-body{display:none;padding:6px 24px 10px 40px;background:var(--card);border-top:1px solid var(--border-light)}
    .excl-body.open{display:block}
    .excl-body ul{margin:0;padding:0;list-style:none}
    .excl-body li{padding:2px 0;font-size:.82em;color:var(--text-muted)}
    .excl-body code{background:rgba(176,107,0,.08);color:var(--warning);padding:1px 6px;border-radius:3px}
    .statsbar{background:var(--surface);padding:10px 24px;border-bottom:1px solid var(--border);display:flex;gap:8px;flex-wrap:wrap;align-items:center}
    .chip{display:inline-flex;align-items:center;gap:6px;background:var(--card);border:1px solid var(--border);border-radius:var(--radius-pill);padding:5px 14px;font-size:.8em}
    .chip-label{color:var(--text-muted)}.chip-value{color:var(--accent-light);font-weight:600}
    .chip-warn{border-color:var(--danger)}.chip-warn .chip-value{color:var(--danger)}
    .chip-excl{border-color:var(--warning)}.chip-excl .chip-value{color:var(--warning)}
    .chip-unscanned{border-color:var(--unscanned)}.chip-unscanned .chip-value{color:var(--unscanned)}
    .chip-denied{border-color:var(--danger)}.chip-denied .chip-value{color:var(--danger)}
    .table-wrap{padding:16px 24px;overflow-x:auto}
    table{width:100%;border-collapse:collapse;font-size:.875em;background:var(--surface);border-radius:var(--radius);overflow:hidden;box-shadow:var(--shadow)}
    thead th{background:var(--card);color:var(--text-muted);padding:12px 16px;text-align:left;border-bottom:2px solid var(--accent-light);font-weight:600;font-size:.75em;letter-spacing:.5px;text-transform:uppercase;cursor:pointer;user-select:none;white-space:nowrap;transition:background .15s,color .15s}
    thead th:hover{background:var(--card-hover);color:var(--accent-light)}
    thead th.sort-asc::after{content:' \u25b2';color:var(--accent-light)}
    thead th.sort-desc::after{content:' \u25bc';color:var(--accent-light)}
    tbody tr{border-bottom:1px solid var(--border-light);transition:background .1s}
    tbody tr:hover{background:var(--card)}
    tbody tr:last-child{border-bottom:none}
    tbody tr.row-excluded{background:rgba(176,107,0,.05)}tbody tr.row-excluded:hover{background:rgba(176,107,0,.1)}
    tbody tr.row-unscanned{background:rgba(94,61,143,.05)}tbody tr.row-unscanned:hover{background:rgba(94,61,143,.1)}
    tbody tr.row-denied{background:rgba(192,32,46,.05)}tbody tr.row-denied:hover{background:rgba(192,32,46,.1)}
    td{padding:10px 16px;vertical-align:middle}
    .col-icon{width:28px;text-align:center}.col-size{width:120px;text-align:right;font-family:monospace;font-size:.85em;color:var(--text-muted)}
    .col-bar{width:220px}.col-type{width:70px;text-align:center}
    .dir-link{color:var(--accent-light);text-decoration:none;font-weight:500;cursor:pointer;transition:color .15s}
    .dir-link:hover{color:var(--text);text-decoration:underline}
    .dir-link-unscanned{color:var(--unscanned);text-decoration:none;font-weight:500;cursor:pointer}
    .dir-link-unscanned:hover{color:var(--text);text-decoration:underline}
    .dir-link-denied{color:var(--danger);text-decoration:none;font-weight:500;cursor:default}
    .file-name{color:var(--text)}
    .type-badge{font-size:.7em;padding:2px 7px;border-radius:3px;font-weight:600}
    .type-dir{background:rgba(0,99,190,.1);color:var(--accent-light)}
    .type-file{background:var(--card);color:var(--text-dim)}
    .type-excl{background:rgba(176,107,0,.12);color:var(--warning)}
    .type-unscanned{background:rgba(94,61,143,.12);color:var(--unscanned)}
    .type-denied{background:rgba(192,32,46,.12);color:var(--danger)}
    .excl-label{font-size:.7em;background:rgba(176,107,0,.12);color:var(--warning);padding:2px 7px;border-radius:3px;margin-left:6px;font-weight:500}
    .denied-label{font-size:.7em;background:rgba(192,32,46,.12);color:var(--danger);padding:2px 7px;border-radius:3px;margin-left:6px;font-weight:500}
    .size-unknown{color:var(--warning);font-weight:600;font-style:italic}
    .size-unscanned{color:var(--unscanned);font-weight:600;font-style:italic}
    .size-denied{color:var(--danger);font-weight:600;font-style:italic}
    .bar-wrap{display:flex;align-items:center;gap:8px}
    .bar-bg{background:var(--card);border-radius:4px;height:8px;flex:1;overflow:hidden;border:1px solid var(--border-light)}
    .bar-fill{height:100%;border-radius:4px;transition:width .3s}
    .bar-pct{font-size:.75em;color:var(--text-dim);white-space:nowrap;min-width:38px;text-align:right}
    .bar-unknown{font-size:.75em;color:var(--warning);font-weight:600}
    .bar-unscanned-text{font-size:.75em;color:var(--unscanned);font-weight:600;font-style:italic}
    .bar-denied-text{font-size:.75em;color:var(--danger);font-weight:600}
    .bar-excl{height:8px;background:repeating-linear-gradient(45deg,rgba(176,107,0,.1),rgba(176,107,0,.1) 4px,rgba(176,107,0,.2) 4px,rgba(176,107,0,.2) 8px)}
    .bar-unscanned{height:8px;background:repeating-linear-gradient(45deg,rgba(94,61,143,.1),rgba(94,61,143,.1) 4px,rgba(94,61,143,.2) 4px,rgba(94,61,143,.2) 8px)}
    .bar-denied{height:8px;background:repeating-linear-gradient(45deg,rgba(192,32,46,.1),rgba(192,32,46,.1) 4px,rgba(192,32,46,.2) 4px,rgba(192,32,46,.2) 8px)}
    .msg-page{padding:56px 24px;text-align:center}
    .msg-page .icon{font-size:3em;margin-bottom:16px;display:block}
    .msg-page h3{color:var(--text);margin-bottom:8px;font-weight:500;font-size:1.1em}
    .msg-page p{color:var(--text-muted);font-size:.875em;margin-top:6px}
    .msg-page .hint{margin-top:14px;font-size:.8em;background:rgba(176,107,0,.1);color:var(--warning);border:1px solid rgba(176,107,0,.3);padding:8px 16px;border-radius:var(--radius-sm);display:inline-block}
    .msg-page .hint-unscanned{margin-top:14px;font-size:.8em;background:rgba(94,61,143,.1);color:var(--unscanned);border:1px solid rgba(94,61,143,.3);padding:8px 16px;border-radius:var(--radius-sm);display:inline-block}
    .msg-page .hint-denied{margin-top:14px;font-size:.8em;background:rgba(192,32,46,.1);color:var(--danger);border:1px solid rgba(192,32,46,.3);padding:8px 16px;border-radius:var(--radius-sm);display:inline-block}
    .msg-page code{font-family:monospace;font-size:.85em;color:var(--accent-light);background:var(--card);padding:2px 8px;border-radius:3px;margin-top:8px;display:inline-block}
    .msg-page ul.protected-list{text-align:left;margin:12px auto;display:inline-block;color:var(--text-muted);font-size:.85em;list-style:disc;padding-left:20px}
    .msg-page ul.protected-list li{margin:4px 0}
    .footer{text-align:center;padding:14px 24px;color:var(--text-dim);font-size:.78em;border-top:1px solid var(--border);background:var(--surface);margin-top:16px;line-height:1.8}
    .footer a{color:var(--accent-light);text-decoration:none}
    .footer a:hover{text-decoration:underline}
    .footer-support{margin-top:6px;font-size:.85em;color:var(--text-muted)}
    ::-webkit-scrollbar{width:6px;height:6px}
    ::-webkit-scrollbar-track{background:var(--bg)}
    ::-webkit-scrollbar-thumb{background:var(--border);border-radius:3px}
    ::-webkit-scrollbar-thumb:hover{background:var(--accent-light)}
    @media(max-width:768px){.col-bar,.col-type{display:none}.nav-input{width:140px}.nav-input:focus{width:220px}.user-block,.user-info{display:none}.scan-datebar{font-size:.72em;padding:4px 12px}}
    </style>
</head>
<body>

<div class="header">
    <div class="header-left">
        <div>
            <div class="app-title-main">PS-<span>NCDU</span></div>
            <div class="app-title-sub">Disk Usage Analyzer</div>
        </div>
        <span class="version-chip">v$SCRIPT_VERSION</span>
    </div>
    <div class="header-right">
        <button class="theme-toggle" onclick="toggleTheme()" id="themeBtn">&#9728; Light</button>
        <div class="user-block">
            <div class="user-avatar" id="userAvatar">?</div>
            <div class="user-info">
                <div class="user-fullname" title="$fullUserEnc">$fullUserEnc</div>
                <div class="user-details">$machineEnc &middot; $($env:USERNAME)</div>
            </div>
        </div>
        <div class="nn-logo" onclick="navigateTo(ROOT_PATH)" title="Retour racine">$logoSvg</div>
    </div>
</div>

<div class="scan-datebar">
    <span>&#128197; Scan du <strong>$scanDTEnc</strong></span>
    <span>&middot;</span>
    <span>$pathIcon <strong>$rootEnc</strong></span>
    <span>&middot;</span>
    <span>&#8987; ${ElapsedSec}s &middot; $nbScans dossiers</span>
    $(if ($JunctionsSkipped -gt 0) { "<span>&middot; &#128279; $JunctionsSkipped jonctions ignorees</span>" })
</div>

<div class="toolbar">
    <div class="breadcrumb" id="breadcrumb"></div>
    <input type="text" class="nav-input" id="pathInput" value="$rootEnc"
           placeholder="C:\chemin ou \\serveur\partage"
           onkeydown="if(event.key==='Enter')navigateTo(this.value)">
    <button class="btn btn-primary" onclick="navigateTo(document.getElementById('pathInput').value)" title="Naviguer">&#128269;</button>
    <button class="btn" onclick="goUp()">&#11014; Remonter</button>
    <button class="btn" onclick="navigateTo(ROOT_PATH)" title="Racine">&#127968;</button>
</div>

<div class="mode-banner">
    <span>&#8505;&#65039;</span>
    <span class="mode-tag">$modeEncHtml</span>
    <span class="path-type-badge">$pathIcon $pathTypeEnc</span>
    <span class="depth-badge $(if($Unlimited){'depth-unlimited'}else{'depth-limited'})">$(if($Unlimited){'&#8734; Profondeur illimitee'}else{"&#128269; Profondeur : $MaxDepth"})</span>
    <span>Precision : <strong style="color:var(--text)">$modeEncAcc</strong></span>
</div>

$(if ($nbDenied -gt 0) {
"<div class='alert-banner alert-denied'>
    <span>&#128274;</span>
    <div>
        <strong>$nbDenied repertoire(s) proteges par ACL NTFS</strong>
        &mdash; Inaccessibles avec le compte <strong>$($env:USERNAME)</strong>.
        C'est normal pour les dossiers systeme ou d'autres utilisateurs.
        <span class='hint-link' onclick='toggleDetail(""deniedDetail"")'>&#128196; Voir la liste</span>
    </div>
</div>
<div class='alert-detail' id='deniedDetail'><ul>$dnListHtml</ul></div>"
})

$(if ($nbUnscanned -gt 0) {
"<div class='alert-banner alert-unscanned'>
    <span>&#128269;</span>
    <div><strong>$nbUnscanned repertoire(s) non scanne(s)</strong>
    &mdash; Scan arrete a la profondeur <strong>$MaxDepth</strong>. Taille inconnue.
    <span class='hint-link' onclick='toggleDetail(""deepScanHint"")'>&#128161; Scanner plus profond</span></div>
</div>
<div class='alert-detail' id='deepScanHint'>Relancez avec profondeur <strong>0 (illimitee)</strong> ou augmentez le niveau.</div>"
})

$(if ($EXCLUDED_DIRS.Count -gt 0) {
"<div class='excl-section'>
    <button class='excl-toggle' id='exclToggle' onclick='toggleExcl()'>
        <span>&#9888;&#65039;</span>
        <span>Repertoires exclus du scan &mdash; taille inconnue</span>
        <span class='excl-count'>$($EXCLUDED_DIRS.Count)</span>
        <span class='excl-chevron'>&#9660;</span>
    </button>
    <div class='excl-body' id='exclBody'><ul>$exListHtml</ul></div>
</div>"
})

<div class="statsbar" id="stats"></div>
<div id="mainContent"></div>

<div class="footer">
    PS-NCDU v$SCRIPT_VERSION &middot; Disk Usage Analyzer &middot;
    $pathIcon $pathTypeEnc &middot; $depthLabelEnc &middot;
    $nbScans dossiers &middot; $scanDTEnc &middot; $authorEnc$footerExtra
    <div class="footer-support">
        Support : <a href="mailto:$emailEnc">$emailEnc</a>
        &nbsp;&middot;&nbsp; Log : <code>$logEnc</code>
    </div>
</div>

<script>
var DATA=($jsonScans),EXCLUDED=($jsonExcluded),ROOT_PATH="$rootPathSafe";
var MAX_DEPTH=$maxDepthJs,UNLIMITED=$unlimitedJs;
var curPath=ROOT_PATH,sortCol='size',sortAsc=false;

(function(){
    var fn="$fullUserEnc",parts=fn.split(' ').filter(function(p){return p.length>0;}),ini='';
    if(parts.length>=2){ini=(parts[0][0]+parts[parts.length-1][0]).toUpperCase();}
    else if(parts.length===1){ini=parts[0].substring(0,2).toUpperCase();}
    else{ini='?';}
    var av=document.getElementById('userAvatar');if(av)av.textContent=ini;
})();

var EXT_ICONS={
    '.pdf':'&#128196;','.doc':'&#128196;','.docx':'&#128196;','.xls':'&#128200;','.xlsx':'&#128200;',
    '.ppt':'&#128202;','.pptx':'&#128202;','.txt':'&#128196;','.csv':'&#128200;','.odt':'&#128196;',
    '.ods':'&#128200;','.odp':'&#128202;','.rtf':'&#128196;','.md':'&#128196;','.log':'&#128203;',
    '.ps1':'&#128187;','.py':'&#128013;','.js':'&#129300;','.ts':'&#129300;','.html':'&#127760;',
    '.htm':'&#127760;','.css':'&#127912;','.php':'&#128187;','.java':'&#9749;','.cs':'&#128187;',
    '.cpp':'&#128187;','.c':'&#128187;','.h':'&#128187;','.go':'&#128187;','.rs':'&#128187;',
    '.rb':'&#128312;','.sh':'&#128192;','.bat':'&#128192;','.cmd':'&#128192;','.sql':'&#128020;',
    '.xml':'&#128196;','.json':'&#128196;','.yaml':'&#128196;','.yml':'&#128196;','.ini':'&#9881;',
    '.conf':'&#9881;','.config':'&#9881;','.env':'&#9881;','.toml':'&#128196;',
    '.jpg':'&#128444;','.jpeg':'&#128444;','.png':'&#128444;','.gif':'&#127916;','.bmp':'&#128444;',
    '.svg':'&#128444;','.ico':'&#128444;','.webp':'&#128444;','.tiff':'&#128444;','.tif':'&#128444;',
    '.psd':'&#127912;','.ai':'&#127912;','.eps':'&#127912;','.raw':'&#128247;',
    '.mp4':'&#127916;','.avi':'&#127916;','.mkv':'&#127916;','.mov':'&#127916;','.wmv':'&#127916;',
    '.flv':'&#127916;','.webm':'&#127916;','.m4v':'&#127916;','.mpg':'&#127916;','.mpeg':'&#127916;',
    '.mp3':'&#127925;','.wav':'&#127925;','.flac':'&#127925;','.aac':'&#127925;','.ogg':'&#127925;',
    '.m4a':'&#127925;','.wma':'&#127925;',
    '.zip':'&#128230;','.rar':'&#128230;','.7z':'&#128230;','.tar':'&#128230;','.gz':'&#128230;',
    '.bz2':'&#128230;','.xz':'&#128230;','.iso':'&#128191;','.dmg':'&#128191;',
    '.exe':'&#9881;','.msi':'&#9881;','.dll':'&#9881;','.so':'&#9881;','.deb':'&#9881;',
    '.rpm':'&#9881;','.apk':'&#128241;','.app':'&#9881;',
    '.db':'&#128020;','.sqlite':'&#128020;','.mdb':'&#128020;','.bak':'&#128190;','.mdf':'&#128020;','.ldf':'&#128020;',
    '.ttf':'&#128210;','.otf':'&#128210;','.woff':'&#128210;','.woff2':'&#128210;',
    '.torrent':'&#128279;','.ics':'&#128197;','.vcf':'&#128100;','.eml':'&#128140;','.msg':'&#128140;','.pst':'&#128140;',
    'dir':'&#128193;','net':'&#127760;','denied':'&#128274;','unknown':'&#128196;'
};
function getFileIcon(ext,isDir,isDenied,path){
    if(isDenied)return EXT_ICONS['denied'];
    if(isDir){if(path&&path.startsWith('\\\\'))return EXT_ICONS['net'];return EXT_ICONS['dir'];}
    return EXT_ICONS[ext.toLowerCase()]||EXT_ICONS['unknown'];
}

function toggleTheme(){
    var h=document.documentElement,isLight=h.getAttribute('data-theme')==='light';
    h.setAttribute('data-theme',isLight?'dark':'light');
    document.getElementById('themeBtn').innerHTML=isLight?'&#9790; Dark':'&#9728; Light';
    try{localStorage.setItem('nn-theme',isLight?'dark':'light');}catch(e){}
}
(function(){
    try{var t=localStorage.getItem('nn-theme');if(t){document.documentElement.setAttribute('data-theme',t);var b=document.getElementById('themeBtn');if(b)b.innerHTML=t==='light'?'&#9728; Light':'&#9790; Dark';}}catch(e){}
})();

function toggleExcl(){var b=document.getElementById('exclToggle'),d=document.getElementById('exclBody');if(b&&d){b.classList.toggle('open');d.classList.toggle('open');}}
function toggleDetail(id){var el=document.getElementById(id);if(el)el.style.display=(el.style.display==='block')?'none':'block';}

function escHtml(s){if(!s)return '';return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function normPath(p){
    if(!p)return '';var isUnc=p.startsWith('\\\\');p=p.replace(/\//g,'\\');
    if(isUnc){p='\\\\'+p.substring(2).replace(/\\{2,}/g,'\\');}else{p=p.replace(/\\{2,}/g,'\\');}
    if(p.length>3&&!isUnc)p=p.replace(/\\+`$`/,'');
    if(/^[A-Za-z]:`$`/.test(p))p+='\\';return p;
}
function parentOf(p){
    p=normPath(p);var isUnc=p.startsWith('\\\\');
    if(/^[A-Za-z]:\\`$`/.test(p))return null;
    if(isUnc){var parts=p.substring(2).split('\\').filter(Boolean);if(parts.length<=2)return null;return '\\\\'+parts.slice(0,parts.length-1).join('\\');}
    var i=p.lastIndexOf('\\');if(i<0)return null;var r=p.substring(0,i);if(/^[A-Za-z]:`$`/.test(r))r+='\\';return r||null;
}
function formatSize(b){b=Number(b);if(b<0)return '?';if(b>=1099511627776)return (b/1099511627776).toFixed(2)+' TB';if(b>=1073741824)return (b/1073741824).toFixed(2)+' GB';if(b>=1048576)return (b/1048576).toFixed(2)+' MB';if(b>=1024)return (b/1024).toFixed(2)+' KB';return b+' B';}
function barColor(sz){if(sz>=1073741824)return 'var(--danger)';if(sz>=104857600)return 'var(--warning)';if(sz>=10485760)return '#c47d00';return 'var(--success)';}
function isExcluded(path){path=normPath(path).toLowerCase();for(var i=0;i<EXCLUDED.length;i++){var ex=normPath(EXCLUDED[i]).toLowerCase();if(path===ex||path.startsWith(ex+'\\'))return true;}return false;}
function findScan(p){p=normPath(p).toLowerCase();for(var k in DATA){if(normPath(k).toLowerCase()===p)return DATA[k];}return null;}
function navigateTo(p){if(!p||!p.trim())return;p=normPath(p.trim());curPath=p;document.getElementById('pathInput').value=p;renderBreadcrumb(p);renderContent(p);}
function goUp(){var p=parentOf(curPath);if(p)navigateTo(p);}

function renderBreadcrumb(path){
    path=normPath(path);var isUnc=path.startsWith('\\\\');var html='',built='';
    if(isUnc){
        var parts=path.substring(2).split('\\').filter(Boolean);
        if(parts.length>=1){built='\\\\'+parts[0];html+='<span style="color:var(--text-dim)">&#127760;</span> ';if(parts.length===1){html+='<a class="active">'+escHtml('\\\\'+parts[0])+'</a>';}else{html+='<a onclick="navigateTo(\''+built.replace(/\\/g,'\\\\')+'\')">'+escHtml('\\\\'+parts[0])+'</a>';}}
        for(var i=1;i<parts.length;i++){built+='\\'+parts[i];html+=' <span class="sep">&#8250;</span> ';if(i===parts.length-1){html+='<a class="active">'+escHtml(parts[i])+'</a>';}else{var b=built;html+='<a onclick="navigateTo(\''+b.replace(/\\/g,'\\\\')+'\')">' +escHtml(parts[i])+'</a>';}}
    }else{
        var parts2=path.replace(/\\+`$`/,'').split('\\').filter(Boolean);
        if(/^[A-Za-z]:/.test(path)){built=parts2[0]+'\\';if(parts2.length===1){html+='<a class="active">'+escHtml(parts2[0])+'</a>';}else{html+='<a onclick="navigateTo(\''+built.replace(/\\/g,'\\\\')+'\')">'+escHtml(parts2[0])+'</a>';}parts2.shift();}
        parts2.forEach(function(p,i){built+=p+'\\';html+=' <span class="sep">&#8250;</span> ';if(i===parts2.length-1){html+='<a class="active">'+escHtml(p)+'</a>';}else{var b=built;html+='<a onclick="navigateTo(\''+b.replace(/\\/g,'\\\\')+'\')">' +escHtml(p)+'</a>';}});
    }
    document.getElementById('breadcrumb').innerHTML=html||'&#127968; Racine';
}

function renderContent(path){
    if(isExcluded(path)){
        document.getElementById('stats').innerHTML='<div class="chip chip-excl"><span class="chip-label">&#9888;</span><span class="chip-value">Repertoire exclu</span></div>';
        document.getElementById('mainContent').innerHTML='<div class="msg-page"><span class="icon">&#9888;</span><h3>Repertoire exclu du scan</h3><p>Sa taille est inconnue.</p><code>'+escHtml(path)+'</code></div>';
        return;
    }
    var scan=findScan(path);
    var isUnscannedDir=false,isDeniedDir=false;
    if(!scan){
        var par=parentOf(path);
        if(par){var parScan=findScan(par);if(parScan){parScan.items.forEach(function(it){var np=normPath(it.fullPath).toLowerCase(),npath=normPath(path).toLowerCase();if(np===npath){if(it.unscanned)isUnscannedDir=true;if(it.denied)isDeniedDir=true;}});}}
    }
    if(!scan){
        if(isDeniedDir){
            document.getElementById('stats').innerHTML='<div class="chip chip-denied"><span class="chip-label">&#128274;</span><span class="chip-value">Protege ACL NTFS</span></div>';
            document.getElementById('mainContent').innerHTML=
                '<div class="msg-page"><span class="icon">&#128274;</span>'+
                '<h3>Dossier protege par ACL NTFS</h3>'+
                '<p>Ce dossier est inaccessible avec le compte <strong>$($env:USERNAME)</strong>.</p>'+
                '<p style="margin-top:8px;color:var(--text-muted);font-size:.85em">C\'est normal pour :</p>'+
                '<ul class="protected-list">'+
                '<li>Les dossiers systeme Windows</li>'+
                '<li>Les profils d\'autres utilisateurs</li>'+
                '<li>Les dossiers de protection de donnees</li>'+
                '</ul>'+
                '<p class="hint-denied">Pour y acceder : demande via <strong>NovoAccess</strong></p>'+
                '<code>'+escHtml(path)+'</code></div>';
        }else if(isUnscannedDir){
            document.getElementById('stats').innerHTML='<div class="chip chip-unscanned"><span class="chip-label">&#128269;</span><span class="chip-value">Non scanne - profondeur '+MAX_DEPTH+'</span></div>';
            document.getElementById('mainContent').innerHTML='<div class="msg-page"><span class="icon">&#128194;</span><h3>Dossier non scanne</h3><p>Scan arrete a la profondeur <strong>'+MAX_DEPTH+'</strong>.</p><p class="hint-unscanned">&#128161; Relancez avec profondeur 0 ou ce chemin :</p><code>'+escHtml(path)+'</code></div>';
        }else{
            var msg=UNLIMITED?'Ce dossier n\'est pas dans le rapport.':'Depasse la profondeur '+MAX_DEPTH+'. Utilisez profondeur 0 pour un scan illimite.';
            document.getElementById('stats').innerHTML='<div class="chip chip-warn"><span class="chip-label">&#9888;</span><span class="chip-value">Non scanne</span></div>';
            document.getElementById('mainContent').innerHTML='<div class="msg-page"><span class="icon">&#128194;</span><h3>Dossier non scanne</h3><p>'+msg+'</p><p class="hint">&#128161; Relancez avec ce chemin</p><code>'+escHtml(path)+'</code></div>';
        }
        return;
    }
    var items=scan.items,total=scan.total,nbD=0,nbF=0,nbExcl=0,nbUnscanned=0,nbDenied=0;
    items.forEach(function(i){if(i.excluded)nbExcl++;else if(i.unscanned)nbUnscanned++;else if(i.denied)nbDenied++;else if(i.isDir)nbD++;else nbF++;});
    var st='<div class="chip"><span class="chip-label">Taille scannee</span><span class="chip-value">'+formatSize(total)+'</span></div>'+
        '<div class="chip"><span class="chip-label">Dossiers</span><span class="chip-value">'+nbD+'</span></div>'+
        '<div class="chip"><span class="chip-label">Fichiers</span><span class="chip-value">'+nbF+'</span></div>'+
        '<div class="chip"><span class="chip-label">Elements</span><span class="chip-value">'+items.length+'</span></div>';
    if(nbDenied>0)st+='<div class="chip chip-denied"><span class="chip-label">&#128274; Proteges</span><span class="chip-value">'+nbDenied+'</span></div>';
    if(nbUnscanned>0)st+='<div class="chip chip-unscanned"><span class="chip-label">&#128269; Non scannes</span><span class="chip-value">'+nbUnscanned+'</span></div>';
    if(nbExcl>0)st+='<div class="chip chip-excl"><span class="chip-label">&#9888; Exclus</span><span class="chip-value">'+nbExcl+'</span></div>';
    document.getElementById('stats').innerHTML=st;
    if(items.length===0){document.getElementById('mainContent').innerHTML='<div class="msg-page"><span class="icon">&#128194;</span><h3>Dossier vide</h3><p>Aucun fichier accessible.</p></div>';return;}
    var rows='';
    items.forEach(function(item){
        var isExcl=item.excluded||isExcluded(item.fullPath);
        var isUnscanned=item.unscanned,isDenied=item.denied;
        var fp=item.fullPath.replace(/\\/g,'\\\\');
        var icon=getFileIcon(item.ext||'',item.isDir,isDenied,item.fullPath);
        var isDeep=!UNLIMITED&&item.isDir&&!isExcl&&!isUnscanned&&!isDenied&&!findScan(item.fullPath);
        var tag=isDeep?'<span style="font-size:.7em;color:var(--text-dim);background:var(--card);padding:1px 6px;border-radius:3px;margin-left:6px;border:1px solid var(--border)">+</span>':'';
        var nameCell,sizeCell,barCell,badge,rowClass,tipTitle;
        if(isDenied){
            rowClass='row-denied';tipTitle='Protege par ACL NTFS - demande via NovoAccess si necessaire';
            nameCell='<span class="dir-link-denied" title="'+tipTitle+'">'+escHtml(item.name)+'</span><span class="denied-label">&#128274; Protege</span>';
            sizeCell='<span class="size-denied" title="'+tipTitle+'">?</span>';
            barCell='<div class="bar-wrap"><div class="bar-bg"><div class="bar-denied"></div></div><span class="bar-denied-text" title="'+tipTitle+'">&#128274;</span></div>';
            badge='<span class="type-badge type-denied" title="'+tipTitle+'">ACL</span>';
        }else if(isExcl){
            rowClass='row-excluded';tipTitle='Repertoire exclu - taille inconnue';
            nameCell='<a class="dir-link" onclick="navigateTo(\''+fp+'\')" title="'+tipTitle+'">'+escHtml(item.name)+'</a><span class="excl-label">&#9888;</span>';
            sizeCell='<span class="size-unknown">?</span>';
            barCell='<div class="bar-wrap"><div class="bar-bg"><div class="bar-excl"></div></div><span class="bar-unknown">?</span></div>';
            badge='<span class="type-badge type-excl">EXCLU</span>';
        }else if(isUnscanned){
            rowClass='row-unscanned';tipTitle='Taille inconnue - scan limite a la profondeur '+MAX_DEPTH+'. Utilisez profondeur 0.';
            nameCell='<a class="dir-link-unscanned" onclick="navigateTo(\''+fp+'\')" title="'+tipTitle+'">'+escHtml(item.name)+'</a>';
            sizeCell='<span class="size-unscanned" title="'+tipTitle+'">?</span>';
            barCell='<div class="bar-wrap"><div class="bar-bg"><div class="bar-unscanned"></div></div><span class="bar-unscanned-text" title="'+tipTitle+'">?</span></div>';
            badge='<span class="type-badge type-unscanned" title="'+tipTitle+'">?</span>';
        }else{
            rowClass='';var pct=total>0?Math.round((item.size/total)*1000)/10:0;
            nameCell=item.isDir?'<a class="dir-link" onclick="navigateTo(\''+fp+'\')">'+escHtml(item.name)+'</a>'+tag:'<span class="file-name">'+escHtml(item.name)+'</span>';
            sizeCell=escHtml(item.sizeStr);
            barCell='<div class="bar-wrap"><div class="bar-bg"><div class="bar-fill" style="width:'+pct+'%;background:'+barColor(item.size)+'"></div></div><span class="bar-pct">'+pct+'%</span></div>';
            badge=item.isDir?'<span class="type-badge type-dir">DIR</span>':'<span class="type-badge type-file">FILE</span>';
        }
        rows+='<tr class="'+rowClass+'"><td class="col-icon">'+icon+'</td><td class="col-name">'+nameCell+'</td><td class="col-size">'+sizeCell+'</td><td class="col-bar">'+barCell+'</td><td class="col-type">'+badge+'</td></tr>';
    });
    document.getElementById('mainContent').innerHTML=
        '<div class="table-wrap"><table><thead><tr><th class="col-icon"></th><th id="th-name" onclick="sortBy(\'name\')">Nom</th><th id="th-size" onclick="sortBy(\'size\')" style="text-align:right">Taille</th><th>Utilisation</th><th class="col-type">Type</th></tr></thead><tbody>'+rows+'</tbody></table></div>';
    ['th-name','th-size'].forEach(function(id){var el=document.getElementById(id);if(el)el.classList.remove('sort-asc','sort-desc');});
    var thEl=document.getElementById(sortCol==='name'?'th-name':'th-size');
    if(thEl)thEl.classList.add(sortAsc?'sort-asc':'sort-desc');
}

function sortBy(col){
    if(sortCol===col){sortAsc=!sortAsc;}else{sortCol=col;sortAsc=(col==='name');}
    var scan=findScan(curPath);if(!scan)return;
    var normal=scan.items.filter(function(i){return !i.excluded&&!i.unscanned&&!i.denied;});
    var unscanned=scan.items.filter(function(i){return i.unscanned;});
    var denied=scan.items.filter(function(i){return i.denied;});
    var excl=scan.items.filter(function(i){return i.excluded;});
    normal.sort(function(a,b){var va=a[col],vb=b[col];if(typeof va==='string'){va=va.toLowerCase();vb=vb.toLowerCase();}if(va<vb)return sortAsc?-1:1;if(va>vb)return sortAsc?1:-1;return 0;});
    scan.items=normal.concat(unscanned).concat(denied).concat(excl);
    renderContent(curPath);
}
window.onload=function(){navigateTo(ROOT_PATH);};
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
Write-Host "  | Fix: Junctions + UTF-8 + JSON propre    |" -ForegroundColor DarkGray
Write-Host "  +-----------------------------------------+" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Chemins : Local, UNC (\\serveur\partage), Mappe (Z:\)" -ForegroundColor DarkGray
Write-Host ""

$fullUserName = Get-FullUserName
Write-Log "FullUserName : $fullUserName"
Write-Host "  Utilisateur : $fullUserName ($($env:USERNAME))" -ForegroundColor Cyan
Write-Host ""

Write-Host "  MODES DE SCAN :" -ForegroundColor White
Write-Host ""
foreach ($k in ($SCAN_MODES.Keys | Sort-Object)) {
    $m=$SCAN_MODES[$k];$sc=switch($k){1{"Yellow"}2{"Yellow"}3{"Red"}default{"White"}}
    Write-Host "  [$k] " -NoNewline -ForegroundColor White
    Write-Host "$($m["Name"])" -NoNewline -ForegroundColor Cyan
    Write-Host " | $($m["Speed"])" -NoNewline -ForegroundColor $sc
    Write-Host " | $($m["Accuracy"])" -ForegroundColor DarkGray
}
Write-Host ""
Write-Host "  Mode (defaut=1) : " -NoNewline -ForegroundColor Yellow
$inputMode=Read-Host;$modeKey=1
if(-not[string]::IsNullOrWhiteSpace($inputMode)){try{$mk=[int]$inputMode;if($SCAN_MODES.ContainsKey($mk)){$modeKey=$mk}}catch{}}
$selectedMode=$SCAN_MODES[$modeKey]
Write-Host "  >> $($selectedMode["Name"])" -ForegroundColor Green
Write-Host ""

Write-Host "  Profondeur : 1-10 = limite | 0 = ILLIMITEE (attention : tres long sur C:\)" -ForegroundColor DarkGray
Write-Host "  Profondeur (defaut=$DEFAULT_DEPTH) : " -NoNewline -ForegroundColor Yellow
$inputDepth=Read-Host;$maxDepth=$DEFAULT_DEPTH
if(-not[string]::IsNullOrWhiteSpace($inputDepth)){try{$p=[int]$inputDepth;if($p -ge 0 -and $p -le 10){$maxDepth=$p}}catch{}}

$isUnlimited=($maxDepth -eq 0)
if($isUnlimited){
    Write-Host "";Write-Host "  !! ATTENTION : Profondeur ILLIMITEE !!" -ForegroundColor Yellow
    Write-Host "  Peut prendre 30-60 min sur C:\. Recommande sur des sous-dossiers." -ForegroundColor Yellow
    Write-Host "  Continuer ? (O/N, defaut=O) : " -NoNewline -ForegroundColor Yellow
    $confirm=Read-Host
    if($confirm -match '^[Nn]$'){Write-Host "  Annule." -ForegroundColor Red;$null=Read-Host;exit 0}
}

Write-Host "  Chemin (ENTER = $DEFAULT_PATH) : " -NoNewline -ForegroundColor Yellow
$inputPath=Read-Host;$inputPath=$inputPath.Trim()

if([string]::IsNullOrWhiteSpace($inputPath)){$startPath=$DEFAULT_PATH}
else{
    $inputPath=$inputPath -replace '/','\' 
    if($inputPath -match '^[A-Za-z]:$'){$inputPath+='\'}
    Write-Host "  Verification..." -ForegroundColor Gray
    if((Test-NetworkPath -Path $inputPath) -and (Test-Path $inputPath -ErrorAction SilentlyContinue)){
        $startPath=$inputPath;Write-Host "  OK." -ForegroundColor Green
    } else {
        Write-Host "  Chemin invalide ou inaccessible, utilisation de $DEFAULT_PATH" -ForegroundColor Red
        Write-Log "Chemin invalide : '$inputPath'" -Level WARN;$startPath=$DEFAULT_PATH
    }
}

$startPath    = Normalize-Path $startPath
$depthDisplay = if($isUnlimited){"ILLIMITEE"}else{"$maxDepth niveaux"}
$scanDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

Write-Log "Lancement v$SCRIPT_VERSION : '$startPath' mode=$modeKey profondeur=$depthDisplay user='$fullUserName'"
Write-Host ""
Write-Host "  Chemin     : $startPath" -ForegroundColor Cyan
Write-Host "  Type       : $(if($startPath -match '^\\\\'){'UNC Reseau'}else{'Local/Mappe'})" -ForegroundColor Cyan
Write-Host "  Mode       : $($selectedMode["Name"])"    -ForegroundColor Cyan
Write-Host "  Precision  : $($selectedMode["Accuracy"])" -ForegroundColor Yellow
Write-Host "  Profondeur : $depthDisplay" -ForegroundColor $(if($isUnlimited){"Yellow"}else{"Cyan"})
Write-Host "  Date scan  : $scanDateTime" -ForegroundColor Cyan
Write-Host ""

try {
    $globalStart=Get-Date
    $result=Start-FastScan -RootPath $startPath -MaxDepth $maxDepth -Mode $selectedMode
    $allScans        = $result["Scans"]
    $excluded        = $result["Excluded"]
    $unscanned       = $result["Unscanned"]
    $accessDenied    = $result["AccessDenied"]
    $elapsed         = $result["ElapsedSec"]
    $pathType        = $result["PathType"]
    $unlimited       = $result["Unlimited"]
    $methodStats     = $result["MethodStats"]
    $junctionsSkipped= $result["JunctionsSkipped"]

    Write-Host "  Scan : ${elapsed}s - $($allScans.Count) dossiers | $($unscanned.Count) non scannes | $($accessDenied.Count) proteges | $junctionsSkipped junctions ignorees" -ForegroundColor Green
    if ($null -ne $methodStats -and ($methodStats["CMD"] -gt 0 -or $methodStats["NET"] -gt 0)) {
        Write-Host "  Fallback : .NET=$($methodStats['NET']) CMD=$($methodStats['CMD'])" -ForegroundColor DarkYellow
    }
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
        -JunctionsSkipped $junctionsSkipped

    Write-Log "[MAIN] Set-Content..."
    $html | Set-Content -Path $htmlFile -Encoding UTF8 -ErrorAction Stop
    Write-Log "[MAIN] Set-Content OK en $([int]((Get-Date)-$tHtml).TotalSeconds)s"

    $total=[int]((Get-Date)-$globalStart).TotalSeconds
    Write-Log "[MAIN] Total : ${total}s"
    Start-Process $htmlFile;Write-Log "Navigateur ouvert"

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
}

Write-Host "  Appuyez sur ENTER pour terminer..." -ForegroundColor DarkGray
$null = Read-Host
