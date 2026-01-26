# GitHub-Version des Skripts
$HtmlPath = "index.html"
$IstPath  = "ProgrammlisteIST.csv"
$GlobalResults = New-Object System.Collections.Generic.List[PSCustomObject]

function Write-To-ProgramList {
    param($Name, $Version, $Bemerkung)
    $IstVersion = "---"
    if (Test-Path $IstPath) {
        $IstData = Import-Csv $IstPath -Delimiter ";" -Encoding UTF8
        $Match = $IstData | Where-Object { $_.Programm -eq $Name }
        if ($Match) { $IstVersion = $Match.Version }
    }
    $Status = "NEU"
    if ($IstVersion -ne "---") {
        if ($Version -eq $IstVersion) { $Status = "OK" } else { $Status = "UPDATE" }
    }
    $GlobalResults.Add([PSCustomObject]@{Programm=$Name; IST=$IstVersion; AKTUELL=$Version; Status=$Status; Bemerkung=$Bemerkung})
}

# ============================================================
# --- HIER DEINE SOFTWARE-ABFRAGEN (Chrome, Firefox, etc.) ---
# ============================================================

# --- 1. GOOGLE CHROME ENTERPRISE (API) ---
try {
    Write-Host "Chrome Enterprise..." -NoNewline
    $ChromeApiUrl = "https://versionhistory.googleapis.com/v1/chrome/platforms/win/channels/stable/versions"
    $ChromeResponse = Invoke-RestMethod -Uri $ChromeApiUrl -Method Get
    $ChromeVersion = $ChromeResponse.versions[0].version 
    Write-To-ProgramList -Name "Google Chrome Enterprise" -Version $ChromeVersion -Bemerkung "Stable Channel (Index 0)"
    Write-Host " [OK: $ChromeVersion]" -ForegroundColor Green
} catch { Write-Warning " Fehler bei Chrome: $($_.Exception.Message)" }


# --- 2. MOZILLA FIREFOX (API) ---
try {
    Write-Host "Mozilla Firefox..." -NoNewline
    $FirefoxApiUrl = "https://product-details.mozilla.org/1.0/firefox_versions.json"
    $FirefoxResponse = Invoke-RestMethod -Uri $FirefoxApiUrl -Method Get
    $FirefoxVersion = $FirefoxResponse.LATEST_FIREFOX_VERSION
    Write-To-ProgramList -Name "Mozilla Firefox" -Version $FirefoxVersion -Bemerkung "Stable Release (Official API)"
    Write-Host " [OK: $FirefoxVersion]" -ForegroundColor Green
} catch { Write-Warning " Fehler bei Firefox: $($_.Exception.Message)" }


# --- 3. ADOBE ACROBAT READER (CHOCOLATEY) ---
try {
    Write-Host "Adobe Reader DC..." -NoNewline
    $ChocoUrl = "https://community.chocolatey.org/packages/adobereader"
    $ChocoResponse = Invoke-WebRequest -Uri $ChocoUrl -UseBasicParsing -UserAgent "Mozilla/5.0"
    if ($ChocoResponse.Content -match 'Adobe Acrobat Reader DC\s+(\d{4}\.\d+\.\d+)') {
        $AdobeVersion = $Matches[1]
        Write-To-ProgramList -Name "Adobe Acrobat Reader" -Version $AdobeVersion -Bemerkung "Quelle: Chocolatey"
        Write-Host " [OK: $AdobeVersion]" -ForegroundColor Green
    }
} catch { Write-Warning " Fehler bei Adobe Reader: $($_.Exception.Message)" }


# --- 4. ADOBE AIR (CHOCOLATEY) ---
try {
    Write-Host "Adobe AIR..." -NoNewline
    $AirUrl = "https://community.chocolatey.org/packages/AdobeAIR"
    $AirResponse = Invoke-WebRequest -Uri $AirUrl -UseBasicParsing -UserAgent "Mozilla/5.0"
    # Sucht nach dem Muster "Adobe AIR Runtime" gefolgt von der Versionsnummer
    if ($AirResponse.Content -match 'Adobe AIR Runtime\s+([\d\.]+)') {
        $AirVersion = $Matches[1]
        Write-To-ProgramList -Name "Adobe AIR" -Version $AirVersion -Bemerkung "Quelle: Chocolatey (Harman)"
        Write-Host " [OK: $AirVersion]" -ForegroundColor Green
    } else {
        Write-Host " [FEHLER]" -ForegroundColor Red
        Write-Warning " Adobe AIR Version nicht gefunden."
    }
} catch { Write-Warning " Fehler bei Adobe AIR: $($_.Exception.Message)" }


# --- 5. JAVA 8 (ORACLE API & CHOCO FALLBACK) ---
try {
    Write-Host "Java 8..." -NoNewline
    $JavaVersion = $null
    $Quelle = ""

    # Versuch A: Oracle API
    try {
        $JavaResp = Invoke-RestMethod -Uri "https://java.oraclecloud.com/currentJavaReleases/8" -Method Get -TimeoutSec 5
        if ($JavaResp.releaseVersion) {
            $JavaVersion = $JavaResp.releaseVersion
            $Quelle = "Oracle Cloud API"
        }
    } catch { $JavaVersion = $null }

    # Versuch B: Chocolatey Fallback
    if (-not $JavaVersion) {
        $ChocoResp = Invoke-WebRequest -Uri "https://community.chocolatey.org/packages/jre8" -UseBasicParsing -UserAgent "Mozilla/5.0"
        if ($ChocoResp.Content -match 'Java SE Runtime Environment\s+(\d+\.\d+\.\d+)') {
            $JavaVersion = $Matches[1]
            $Quelle = "Chocolatey (Fallback)"
        }
    }

    if ($JavaVersion) {
        # --- NEU: FORMATIERUNG ---
        # Wandelt "1.8.0_471" oder "1.8.0.471" in "8.471" um
        $JavaVersion = $JavaVersion -replace '^1\.8\.0[_.]', '8.'
        # Falls die API direkt "8.0.471" liefert, wird auch das zu "8.471"
        $JavaVersion = $JavaVersion -replace '^8\.0\.', '8.'
        # -------------------------

        Write-To-ProgramList -Name "Java 8" -Version $JavaVersion -Bemerkung "Quelle: $Quelle"
        Write-Host " [OK: $JavaVersion]" -ForegroundColor Green
    } else { 
        Write-Host " [FEHLER]" -ForegroundColor Red 
    }
} catch { 
    Write-Warning " Fehler bei Java: $($_.Exception.Message)" 
}


# --- 6. PDF24 CREATOR (OFFIZIELLER CHANGELOG) ---
try {
    Write-Host "PDF24 Creator..." -NoNewline
    $PdfUrl = "https://creator.pdf24.org/changelog/de.html"
    
    # Seite abrufen (User-Agent sorgt dafür, dass wir nicht blockiert werden)
    $PdfResponse = Invoke-WebRequest -Uri $PdfUrl -UseBasicParsing -UserAgent "Mozilla/5.0"
    
    # Wir suchen nach dem ersten Vorkommen von "vX.X.X" (z.B. v11.29.0)
    # Da die neueste Version immer ganz oben steht, nimmt -match automatisch den ersten Treffer.
    if ($PdfResponse.Content -match 'v(\d+\.\d+\.\d+)') {
        $PdfVersion = $Matches[1]
        
        Write-To-ProgramList -Name "PDF24 Creator" -Version $PdfVersion -Bemerkung "Offizieller Changelog"
        Write-Host " [OK: $PdfVersion]" -ForegroundColor Green
    } else {
        Write-Host " [FEHLER]" -ForegroundColor Red
        Write-Warning " PDF24 Version konnte im Changelog nicht gefunden werden."
    }
} catch { 
    Write-Host " [FEHLER]" -ForegroundColor Red
    Write-Warning " Fehler bei PDF24: $($_.Exception.Message)" 
}


# --- 7. FOXIT PDF READER (CHOCOLATEY) ---
try {
    Write-Host "Foxit PDF Reader..." -NoNewline
    $FoxitUrl = "https://community.chocolatey.org/packages/foxitreader"
    
    # Abruf der Seite mit Browser-Tarnung
    $FoxitResponse = Invoke-WebRequest -Uri $FoxitUrl -UseBasicParsing -UserAgent "Mozilla/5.0"
    
    # Suche nach "Foxit PDF Reader" gefolgt von der Versionsnummer
    # Der Ausdruck [\d\.]+ findet jede Versionsnummer, egal ob sie mit 2025 oder 2026 beginnt
    if ($FoxitResponse.Content -match 'Foxit PDF Reader\s+([\d\.]+)') {
        $FoxitVersion = $Matches[1]
        
        Write-To-ProgramList -Name "Foxit PDF Reader" -Version $FoxitVersion -Bemerkung "Quelle: Chocolatey"
        Write-Host " [OK: $FoxitVersion]" -ForegroundColor Green
    } else {
        Write-Host " [FEHLER]" -ForegroundColor Red
        Write-Warning " Foxit Version konnte bei Chocolatey nicht gefunden werden."
    }
} catch { 
    Write-Host " [FEHLER]" -ForegroundColor Red
    Write-Warning " Fehler bei Foxit: $($_.Exception.Message)" 
}


# --- 8. LEGO MINDSTORMS EV3 CLASSROOM (APPLE STORE) ---
try {
    Write-Host "Lego EV3 Classroom..." -NoNewline
    
    # WICHTIG: Reset, damit kein "Foxit-Müll" in der Variable bleibt
    $Matches = $null
    $LegoVersion = $null

    # Deine funktionierende URL und Abfrage
    $LegoUrl = "https://apps.apple.com/us/app/ev3-classroom-lego-education/id1502412247"
    $LegoResponse = Invoke-WebRequest -Uri $LegoUrl -UseBasicParsing
    
    # Dein Regex-Muster: Sucht nach 'Version History' und extrahiert die erste Versionsnummer
    if ($LegoResponse.Content -match '(?s)Version History.*?(?<!\d)(\d+\.\d+\.\d+)') {
        $LegoVersion = $Matches[1]
        
        # In die Liste schreiben
        Write-To-ProgramList -Name "Lego Mindstorms EV3 Classroom" -Version $LegoVersion -Bemerkung "Quelle: Apple App Store (Proxy)"
        Write-Host " [OK: $LegoVersion]" -ForegroundColor Green
    } 
    else {
        Write-Host " [FEHLER]" -ForegroundColor Red
        Write-Warning " Version konnte im App Store Quelltext nicht gefunden werden."
    }
} catch { 
    Write-Host " [FEHLER]" -ForegroundColor Red
    Write-Warning " Fehler bei Lego: $($_.Exception.Message)" 
}


# --- 9. WORKSHEET CRAFTER (OFFIZIELLE DOWNLOAD-SEITE) ---
try {
    Write-Host "Worksheet Crafter..." -NoNewline
    
    # WICHTIG: Variablen-Reset
    $Matches = $null
    $WscVersion = $null

    # Die offizielle Vollversions-Download-Seite
    $WscUrl = "https://worksheetcrafter.com/de/downloads/vollversion"
    $WscResponse = Invoke-WebRequest -Uri $WscUrl -UseBasicParsing -UserAgent "Mozilla/5.0"
    
    # Wir suchen nach dem Begriff "Version" gefolgt von einer Nummer (z.B. 2025.2.7)
    # Das Muster \d{4} stellt sicher, dass wir die Jahreszahl-Versionen finden
    if ($WscResponse.Content -match 'Version\s+(\d{4}\.[\d\.]+)') {
        $WscVersion = $Matches[1]
        
        Write-To-ProgramList -Name "Worksheet Crafter" -Version $WscVersion -Bemerkung "Offizielle Download-Seite"
        Write-Host " [OK: $WscVersion]" -ForegroundColor Green
    } 
    else {
        Write-Host " [FEHLER]" -ForegroundColor Red
        Write-Warning " Version konnte auf der Download-Seite nicht gefunden werden."
    }
} catch { 
    Write-Host " [FEHLER]" -ForegroundColor Red
    Write-Warning " Fehler bei Worksheet Crafter: $($_.Exception.Message)" 
}


# --- 10. PAINT.NET (CHOCOLATEY) ---
try {
    Write-Host "Paint.NET..." -NoNewline
    
    # WICHTIG: Variablen-Reset
    $Matches = $null
    $PaintVersion = $null

    $PaintNetUrl = "https://community.chocolatey.org/packages/paint.net"
    $PaintNetResponse = Invoke-WebRequest -Uri $PaintNetUrl -UseBasicParsing -UserAgent "Mozilla/5.0"
    
    # Suche nach "paint.net" gefolgt von der Versionsnummer
    if ($PaintNetResponse.Content -match 'paint.net\s+([\d\.]+)') {
        $PaintVersion = $Matches[1]
        
        Write-To-ProgramList -Name "Paint.NET" -Version $PaintVersion -Bemerkung "Quelle: Chocolatey"
        Write-Host " [OK: $PaintVersion]" -ForegroundColor Green
    } 
    else {
        Write-Host " [FEHLER]" -ForegroundColor Red
        Write-Warning " Paint.NET Version konnte bei Chocolatey nicht gefunden werden."
    }
} catch { 
    Write-Host " [FEHLER]" -ForegroundColor Red
    Write-Warning " Fehler bei Paint.NET: $($_.Exception.Message)" 
}

# --- 11. SHOTCUT (RELEASE NOTES Liste) ---
try {
    Write-Host "Shotcut..." -NoNewline

    $Matches = $null
    $ShotcutVersion = $null

    $ShotcutUrl = "https://www.shotcut.org/download/releasenotes/"
    $ShotcutResponse = Invoke-WebRequest -Uri $ShotcutUrl -UseBasicParsing -UserAgent "Mozilla/5.0"

    # Offizielle "New Version YY.MM" nehmen (stable) [web:98]
    if ($ShotcutResponse.Content -match 'Release\s+(\d+\.\d+\.\d+)\b') {
    $ShotcutVersion = $Matches[1]

    Write-To-ProgramList -Name "Shotcut" -Version $ShotcutVersion -Bemerkung "Quelle: shotcut.org (Release Notes)"
    Write-Host " [OK: $ShotcutVersion]" -ForegroundColor Green
	}
	elseif ($ShotcutResponse.Content -match 'New Version\s+(\d+\.\d+)\b') {
		$ShotcutVersion = $Matches[1]

		Write-To-ProgramList -Name "Shotcut" -Version $ShotcutVersion -Bemerkung "Quelle: shotcut.org (Release Notes)"
		Write-Host " [OK: $ShotcutVersion]" -ForegroundColor Green
	}
	else {
		Write-Host " [FEHLER]" -ForegroundColor Red
		Write-Warning " Shotcut Version konnte in den Release Notes nicht gefunden werden."
	}
} catch {
    Write-Host " [FEHLER]" -ForegroundColor Red
    Write-Warning " Fehler bei Shotcut: $($_.Exception.Message)"
}


# --- 12. SWEET HOME 3D (CHOCOLATEY) ---
try {
    Write-Host "Sweet Home 3D..." -NoNewline
    
    # WICHTIG: Variablen-Reset
    $Matches = $null
    $SweetHomeVersion = $null

    # Korrigierte URL mit Bindestrichen
    $SweetHomeUrl = "https://community.chocolatey.org/packages/sweet-home-3d"
    $SweetHomeResponse = Invoke-WebRequest -Uri $SweetHomeUrl -UseBasicParsing -UserAgent "Mozilla/5.0"
    
    # Suche nach dem Anzeigenamen im Quelltext
    if ($SweetHomeResponse.Content -match 'Sweet Home 3D\s+([\d\.]+)') {
        $SweetHomeVersion = $Matches[1]
        
        Write-To-ProgramList -Name "Sweet Home 3D" -Version $SweetHomeVersion -Bemerkung "Quelle: Chocolatey"
        Write-Host " [OK: $SweetHomeVersion]" -ForegroundColor Green
    } 
    else {
        Write-Host " [FEHLER]" -ForegroundColor Red
        Write-Warning " Sweet Home 3D Version konnte bei Chocolatey nicht gefunden werden."
    }
} catch { 
    Write-Host " [FEHLER]" -ForegroundColor Red
    Write-Warning " Fehler bei Sweet Home 3D: $($_.Exception.Message)" 
}

# --- 13. VLC MEDIA PLAYER (DOWNLOAD-LINK VARIANTE) ---
try {
    Write-Host "VLC Media Player..." -NoNewline
    
    # WICHTIG: Variablen-Reset
    $Matches = $null
    $VlcVersion = $null

    # Deine effiziente URL
    $VlcUrl = "https://www.videolan.org/vlc/download-windows.html"
    $VlcResponse = Invoke-WebRequest -Uri $VlcUrl -UseBasicParsing -UserAgent "Mozilla/5.0"
    
    # Deine Regex-Logik: Sucht nach dem Muster im Dateinamen
    if ($VlcResponse.Content -match 'vlc-([\d\.]+)-win') {
        $VlcVersion = $Matches[1]
        
        Write-To-ProgramList -Name "VLC Media Player" -Version $VlcVersion -Bemerkung "Quelle: Offizielle Download-Seite"
        Write-Host " [OK: $VlcVersion]" -ForegroundColor Green
    } 
    else {
        Write-Host " [FEHLER]" -ForegroundColor Red
        Write-Warning " VLC Version konnte im Download-Link nicht gefunden werden."
    }
} catch { 
    Write-Host " [FEHLER]" -ForegroundColor Red
    Write-Warning " Fehler bei VLC: $($_.Exception.Message)" 
}

# --- 14. LEGO SPIKE APP (LEGO EDUCATION RELEASE NOTES) ---
try {
    Write-Host "Lego SPIKE App..." -NoNewline
    
    # WICHTIG: Variablen-Reset
    $Matches = $null
    $SpikeVersion = $null

    # Deine funktionierende URL
    $SpikeUrl = "https://legoeducation.atlassian.net/servicedesk/customer/article/38611681568"
    $SpikeResponse = Invoke-WebRequest -Uri $SpikeUrl -UseBasicParsing -UserAgent "Mozilla/5.0"
    
    # Wir kombinieren deine Suchmuster in einer Abfrage
    # Zuerst suchen wir nach der spezifischen Kombination "SPIKE App version X.X.X"
    if ($SpikeResponse.Content -match 'SPIKE.*?App.*?version\s+(\d+\.\d+\.\d+)' -or 
        $SpikeResponse.Content -match 'version\s+(\d+\.\d+\.\d+)') {
        
        $SpikeVersion = $Matches[1]
        
        # In die Liste schreiben
        Write-To-ProgramList -Name "Lego SPIKE App" -Version $SpikeVersion -Bemerkung "Quelle: Lego Education Release Notes"
        Write-Host " [OK: $SpikeVersion]" -ForegroundColor Green
    } 
    else {
        Write-Host " [FEHLER]" -ForegroundColor Red
        Write-Warning " SPIKE Version konnte in den Release Notes nicht gefunden werden."
    }
} catch { 
    Write-Host " [FEHLER]" -ForegroundColor Red
    Write-Warning " Fehler bei Lego SPIKE: $($_.Exception.Message)" 
}

# --- 15. SMART NOTEBOOK (SMART TECH UPDATES) ---
try {
    Write-Host "SMART Notebook..." -NoNewline
    
    # WICHTIG: Variablen-Reset
    $Matches = $null
    $SmartVersion = $null

    # Deine neu gefundene, funktionierende URL
    $SmartUrl = "https://techupdates.smarttech.com/article/smart-notebook-25-1"
    $SmartResponse = Invoke-WebRequest -Uri $SmartUrl -UseBasicParsing -UserAgent "Mozilla/5.0"
    
    # Wir nutzen deine Regex-Logik
    # Sie sucht nach "Notebook" oder "Version" gefolgt von der Nummer (z.B. 25.1)
    if ($SmartResponse.Content -match 'Notebook\s+(\d+\.\d+)' -or 
        $SmartResponse.Content -match 'Version\s+(\d+\.\d+)') {
        
        $SmartVersion = $Matches[1]
        
        Write-To-ProgramList -Name "SMART Notebook" -Version $SmartVersion -Bemerkung "Quelle: SMART Tech Updates"
        Write-Host " [OK: $SmartVersion]" -ForegroundColor Green
    } 
    else {
        Write-Host " [FEHLER]" -ForegroundColor Red
        Write-Warning " SMART Notebook Version konnte auf der Update-Seite nicht gefunden werden."
    }
} catch { 
    Write-Host " [FEHLER]" -ForegroundColor Red
    Write-Warning " Fehler bei SMART Notebook: $($_.Exception.Message)" 
}


# --- 16. LYNX WHITEBOARD (GOOGLE PLAY STORE QUELLE) ---
try {
    Write-Host "LYNX Whiteboard..." -NoNewline
    
    # WICHTIG: Variablen-Reset
    $Matches = $null
    $LynxVersion = $null

    # Deine funktionierende Play-Store-URL
    $LynxUrl = "https://play.google.com/store/apps/details?id=com.clevertouch.lynx&hl=de"
    $LynxResponse = Invoke-WebRequest -Uri $LynxUrl -UseBasicParsing -UserAgent "Mozilla/5.0"
    
    # Deine Logik: Suche nach allen 8.x Versionen im Quelltext
    # Wir extrahieren alle Treffer, die wie eine Version aussehen
    $Pattern = '8\.\d+\.\d+(?:\.\d+)?'
    $AllMatches = [regex]::Matches($LynxResponse.Content, $Pattern)
    
    if ($AllMatches.Count -gt 0) {
        # Wir sortieren die gefundenen Nummern und nehmen die höchste
        $LynxVersion = $AllMatches.Value | Sort-Object { [version]$_ } -Descending | Select-Object -First 1
        
        Write-To-ProgramList -Name "LYNX Whiteboard" -Version $LynxVersion -Bemerkung "Quelle: Google Play Store"
        Write-Host " [OK: $LynxVersion]" -ForegroundColor Green
    } 
    else {
        Write-Host " [FEHLER]" -ForegroundColor Red
        Write-Warning " LYNX Version konnte im Play Store nicht identifiziert werden."
    }
} catch { 
    Write-Host " [FEHLER]" -ForegroundColor Red
    Write-Warning " Fehler bei LYNX Whiteboard: $($_.Exception.Message)" 
}


# --- 17. OPENBOARD (GITHUB API) ---
try {
    Write-Host "OpenBoard..." -NoNewline
    
    # WICHTIG: Variablen-Reset
    $Matches = $null
    $OpenBoardVersion = $null

    # Wir nutzen die offizielle GitHub API für das aktuellste Release
    $ObApiUrl = "https://api.github.com/repos/OpenBoard-org/OpenBoard/releases/latest"
    
    # Invoke-RestMethod ist ideal für APIs, da es direkt ein Objekt zurückgibt
    $ObResponse = Invoke-RestMethod -Uri $ObApiUrl
    
    # Wir extrahieren die Version aus dem Tag-Namen (z.B. v1.7.3 -> 1.7.3)
    if ($ObResponse.tag_name -match 'v?(\d+\.\d+\.\d+)') {
        $OpenBoardVersion = $Matches[1]
        
        Write-To-ProgramList -Name "OpenBoard" -Version $OpenBoardVersion -Bemerkung "Quelle: GitHub API"
        Write-Host " [OK: $OpenBoardVersion]" -ForegroundColor Green
    } 
    else {
        Write-Host " [FEHLER]" -ForegroundColor Red
        Write-Warning " OpenBoard Version konnte via GitHub API nicht ermittelt werden."
    }
} catch { 
    Write-Host " [FEHLER]" -ForegroundColor Red
    Write-Warning " Fehler bei OpenBoard: $($_.Exception.Message)" 
}


# --- 18. 7-ZIP (OFFIZIELLE WEBSEITE) ---
try {
    Write-Host "7-Zip..." -NoNewline
    
    # WICHTIG: Variablen-Reset
    $Matches = $null
    $ZipVersion = $null

    # Die offizielle Download-Seite ist extrem leichtgewichtig
    $ZipUrl = "https://www.7-zip.org/"
    $ZipResponse = Invoke-WebRequest -Uri $ZipUrl -UseBasicParsing -UserAgent "Mozilla/5.0"
    
    # Wir suchen nach "Download 7-Zip" gefolgt von der Versionsnummer (z.B. 24.08)
    # Da die aktuellste Version immer ganz oben steht, nimmt -match den ersten Treffer.
    if ($ZipResponse.Content -match 'Download 7-Zip\s+([\d\.]+)') {
        $ZipVersion = $Matches[1]
        
        Write-To-ProgramList -Name "7-Zip" -Version $ZipVersion -Bemerkung "Quelle: 7-zip.org"
        Write-Host " [OK: $ZipVersion]" -ForegroundColor Green
    } 
    else {
        Write-Host " [FEHLER]" -ForegroundColor Red
        Write-Warning " 7-Zip Version konnte auf der Webseite nicht gefunden werden."
    }
} catch { 
    Write-Host " [FEHLER]" -ForegroundColor Red
    Write-Warning " Fehler bei 7-Zip: $($_.Exception.Message)" 
}


# --- 19. GIMP (OFFIZIELLE JSON-API - NUR LATEST) ---
try {
    Write-Host "GIMP..." -NoNewline
    
    $GimpVersion = $null
    $GimpApiUrl = "https://www.gimp.org/gimp_versions.json"
    
    # Wir laden das JSON-Objekt
    $GimpData = Invoke-RestMethod -Uri $GimpApiUrl
    
    # Wir schauen nach der stabilen Version. 
    # Falls es eine Liste ist, nehmen wir mit [0] nur das allererste (neueste) Element.
    if ($GimpData.stable -is [array]) {
        $GimpVersion = $GimpData.stable[0].version
    } 
    elseif ($GimpData.stable.windows) {
        $GimpVersion = $GimpData.stable.windows[0].version
    }
    else {
        $GimpVersion = $GimpData.stable.version
    }

    # Sicherheits-Check: Falls immer noch mehrere Versionen drinstecken (wie eben),
    # erzwingen wir die Auswahl des ersten Elements.
    if ($GimpVersion -is [array]) {
        $GimpVersion = $GimpVersion[0]
    }

    if ($GimpVersion) {
        Write-To-ProgramList -Name "GIMP" -Version $GimpVersion -Bemerkung "Quelle: gimp.org API (Latest)"
        Write-Host " [OK: $GimpVersion]" -ForegroundColor Green
    } 
    else {
        Write-Host " [FEHLER]" -ForegroundColor Red
        Write-Warning " GIMP Version konnte nicht eindeutig bestimmt werden."
    }
} catch { 
    Write-Host " [FEHLER]" -ForegroundColor Red
    Write-Warning " Fehler bei GIMP: $($_.Exception.Message)" 
}


# --- 20. INKSCAPE (OFFIZIELLE WEBSEITE) ---
try {
    Write-Host "Inkscape..." -NoNewline
    
    # WICHTIG: Variablen-Reset
    $Matches = $null
    $InkVersion = $null

    # Die offizielle Seite für die aktuelle stabile Version
    $InkUrl = "https://inkscape.org/de/release/"
    $InkResponse = Invoke-WebRequest -Uri $InkUrl -UseBasicParsing -UserAgent "Mozilla/5.0"
    
    # Wir suchen nach "Inkscape" gefolgt von der Versionsnummer (z.B. Inkscape 1.3.2)
    # Wir suchen im Seitentitel oder in den Hauptüberschriften
    if ($InkResponse.Content -match 'Inkscape\s+([\d\.]+)') {
        $InkVersion = $Matches[1]
        
        Write-To-ProgramList -Name "Inkscape" -Version $InkVersion -Bemerkung "Quelle: inkscape.org"
        Write-Host " [OK: $InkVersion]" -ForegroundColor Green
    } 
    else {
        # Fallback: Falls die Nummer in einem anderen Format steht
        if ($InkResponse.Content -match 'release-([\d-]+)/') {
            $InkVersion = $Matches[1] -replace '-', '.'
            Write-To-ProgramList -Name "Inkscape" -Version $InkVersion -Bemerkung "Quelle: inkscape.org (Fallback)"
            Write-Host " [OK: $InkVersion]" -ForegroundColor Green
        } else {
            Write-Host " [FEHLER]" -ForegroundColor Red
            Write-Warning " Inkscape Version konnte auf der Webseite nicht gefunden werden."
        }
    }
} catch { 
    Write-Host " [FEHLER]" -ForegroundColor Red
    Write-Warning " Fehler bei Inkscape: $($_.Exception.Message)" 
}


# --- 21. IRFANVIEW (OFFIZIELLE WEBSEITE) ---
try {
    Write-Host "IrfanView..." -NoNewline
    
    # WICHTIG: Variablen-Reset
    $Matches = $null
    $IrfanVersion = $null

    # Die Startseite von IrfanView ist extrem leichtgewichtig
    $IrfanUrl = "https://www.irfanview.com/"
    $IrfanResponse = Invoke-WebRequest -Uri $IrfanUrl -UseBasicParsing -UserAgent "Mozilla/5.0"
    
    # Wir suchen nach "Current version" oder "Version" gefolgt von der Nummer (z.B. 4.70)
    if ($IrfanResponse.Content -match '(?:Current version|Version)\s+([\d\.]+)') {
        $IrfanVersion = $Matches[1]
        
        Write-To-ProgramList -Name "IrfanView" -Version $IrfanVersion -Bemerkung "Quelle: irfanview.com"
        Write-Host " [OK: $IrfanVersion]" -ForegroundColor Green
    } 
    else {
        Write-Host " [FEHLER]" -ForegroundColor Red
        Write-Warning " IrfanView Version konnte auf der Webseite nicht gefunden werden."
    }
} catch { 
    Write-Host " [FEHLER]" -ForegroundColor Red
    Write-Warning " Fehler bei IrfanView: $($_.Exception.Message)" 
}

# --- 22. AUDACITY (GITHUB API) ---
try {
    Write-Host "Audacity..." -NoNewline
    
    # WICHTIG: Variablen-Reset
    $Matches = $null
    $AudacityVersion = $null

    # Die offizielle GitHub API für das neueste Release von Audacity
    $AudacityApiUrl = "https://api.github.com/repos/audacity/audacity/releases/latest"
    
    # Abruf der Daten als Objekt
    $AudacityResponse = Invoke-RestMethod -Uri $AudacityApiUrl
    
    # Der Tag-Name bei Audacity sieht oft so aus: "Audacity-3.4.2" oder "v3.4.2"
    # Wir extrahieren nur die Versionsnummer
    if ($AudacityResponse.tag_name -match '(\d+\.\d+\.\d+)') {
        $AudacityVersion = $Matches[1]
        
        Write-To-ProgramList -Name "Audacity" -Version $AudacityVersion -Bemerkung "Quelle: GitHub API"
        Write-Host " [OK: $AudacityVersion]" -ForegroundColor Green
    } 
    else {
        Write-Host " [FEHLER]" -ForegroundColor Red
        Write-Warning " Audacity Version konnte nicht aus dem API-Tag extrahiert werden."
    }
} catch { 
    Write-Host " [FEHLER]" -ForegroundColor Red
    Write-Warning " Fehler bei Audacity: $($_.Exception.Message)" 
}


# --- 23. MUSESCORE (GITHUB API) ---
try {
    Write-Host "MuseScore..." -NoNewline
    
    # WICHTIG: Variablen-Reset
    $Matches = $null
    $MuseVersion = $null

    # Die offizielle GitHub API für das neueste Release von MuseScore
    $MuseApiUrl = "https://api.github.com/repos/musescore/MuseScore/releases/latest"
    
    # Abruf der Daten als Objekt
    $MuseResponse = Invoke-RestMethod -Uri $MuseApiUrl
    
    # Der Tag-Name bei MuseScore sieht oft so aus: "v4.2.1" oder "4.2.1"
    # Wir extrahieren die reine Versionsnummer
    if ($MuseResponse.tag_name -match '(\d+\.\d+\.\d+)') {
        $MuseVersion = $Matches[1]
        
        Write-To-ProgramList -Name "MuseScore" -Version $MuseVersion -Bemerkung "Quelle: GitHub API"
        Write-Host " [OK: $MuseVersion]" -ForegroundColor Green
    } 
    else {
        Write-Host " [FEHLER]" -ForegroundColor Red
        Write-Warning " MuseScore Version konnte nicht aus dem API-Tag extrahiert werden."
    }
} catch { 
    Write-Host " [FEHLER]" -ForegroundColor Red
    Write-Warning " Fehler bei MuseScore: $($_.Exception.Message)" 
}


# --- 24. MINECRAFT EDUCATION (DIRECT DOWNLOAD REDIRECT) ---
try {
    Write-Host "Minecraft Education..." -NoNewline
    
    $Matches = $null
    $McVersion = $null

    # Dein direkter Download-Link
    $AkaUrl = "https://aka.ms/downloadmee-desktopApp"

    # Wir erstellen eine Web-Anfrage, folgen dem Redirect, laden aber die Datei NICHT herunter
    $Request = [System.Net.HttpWebRequest]::Create($AkaUrl)
    $Request.AllowAutoRedirect = $true
    $Request.Method = "HEAD" # Wir wollen nur den Header (die Adresse), nicht den Inhalt
    
    $Response = $Request.GetResponse()
    $RealUrl = $Response.ResponseUri.ToString()
    $Response.Close()

    # Die RealUrl sieht meistens so aus: .../MinecraftEducation_x64_1.21.05.0.msi
    # Wir suchen nach der Zahlenfolge (z.B. 1.21.05.0)
    if ($RealUrl -match '(\d+\.\d+\.\d+(?:\.\d+)?)') {
        $McVersion = $Matches[1]
        
        Write-To-ProgramList -Name "Minecraft Education" -Version $McVersion -Bemerkung "Quelle: aka.ms Redirect"
        Write-Host " [OK: $McVersion]" -ForegroundColor Green
    } 
    else {
        Write-Host " [FEHLER]" -ForegroundColor Red
        Write-Warning " Version konnte nicht aus der Ziel-URL extrahiert werden: $RealUrl"
    }
} catch { 
    Write-Host " [FEHLER]" -ForegroundColor Red
    Write-Warning " Fehler bei Minecraft Education: $($_.Exception.Message)" 
}


# --- 25. AFFINITY (FREE-CODECS) ---
try {
    Write-Host "Affinity Suite..." -NoNewline
    
    # WICHTIG: Variablen-Reset
    $Matches = $null
    $AffVersion = $null

    # Wir nutzen die Drittanbieter-Quelle Free-Codecs (sehr stabil für Metadaten)
    $AffUrl = "https://www.free-codecs.com/download/affinity.htm"
    $AffResponse = Invoke-WebRequest -Uri $AffUrl -UseBasicParsing -UserAgent "Mozilla/5.0"
    
    # Wir suchen spezifisch nach "Affinity" gefolgt von der Versionsnummer
    # Der Regex erkennt Formate wie 3.0.2 oder auch vierstellige Versionen
    if ($AffResponse.Content -match 'Affinity[^\d]*?(\d+\.\d+\.\d+(?:\.\d+)?)') {
        $AffVersion = $Matches[1]
        
        Write-To-ProgramList -Name "Affinity Suite" -Version $AffVersion -Bemerkung "Quelle: Free-Codecs"
        Write-Host " [OK: $AffVersion]" -ForegroundColor Green
    } 
    else {
        Write-Host " [FEHLER]" -ForegroundColor Red
        Write-Warning " Affinity Version konnte auf Free-Codecs nicht gefunden werden."
    }
} catch { 
    Write-Host " [FEHLER]" -ForegroundColor Red
    Write-Warning " Fehler bei Affinity: $($_.Exception.Message)" 
}


# --- 26. HELBLING MEDIA APP (PARTIAL DOWNLOAD) ---
try {
    Write-Host "Helbling Media App..." -NoNewline
    
    $HelVersion = $null
    $Url = "https://mediaapp.helbling.com/downloads/OU34DJKB/latest/HELBLING%20Media%20App%20Setup.exe"
    $TempPath = Join-Path $env:TEMP "HelblingCheck.exe"

    # Wir laden nur die ersten 2 MB – das reicht für die Metadaten völlig aus
    $Request = [System.Net.HttpWebRequest]::Create($Url)
    $Request.AddRange(0, 2MB - 1)
    $Response = $Request.GetResponse()
    
    $FileStream = [System.IO.File]::Create($TempPath)
    $Response.GetResponseStream().CopyTo($FileStream)
    $FileStream.Close()
    $Response.Close()

    # Version direkt aus der Datei auslesen
    $HelVersion = (Get-Item $TempPath).VersionInfo.FileVersion
    
    # Datei sofort wieder löschen
    if (Test-Path $TempPath) { Remove-Item $TempPath -Force }

    if ($HelVersion) {
        Write-To-ProgramList -Name "Helbling Media App" -Version $HelVersion -Bemerkung "Quelle: File-Header (Partial Download)"
        Write-Host " [OK: $HelVersion]" -ForegroundColor Green
    } 
    else {
        Write-Host " [FEHLER]" -ForegroundColor Red
        Write-Warning " Version konnte nicht aus dem Datei-Header gelesen werden."
    }
} catch { 
    Write-Host " [FEHLER]" -ForegroundColor Red
    Write-Warning " Fehler bei Helbling: $($_.Exception.Message)" 
}

# --- 27. FFmpeg (GITHUB API - GyanD Builds) ---
try {
    Write-Host "FFmpeg..." -NoNewline
    
    # WICHTIG: Variablen-Reset
    $Matches = $null
    $FfmpegVersion = $null

    # Wir nutzen die GitHub API von GyanD (offizieller Windows-Build-Provider)
    $FfmpegApiUrl = "https://api.github.com/repos/GyanD/codexffmpeg/releases/latest"
    $FfmpegResponse = Invoke-RestMethod -Uri $FfmpegApiUrl
    
    # Die Tags dort sehen meist so aus: "7.1" oder "2024-12-10-git-..."
    # Wir suchen nach der ersten Versionsnummer (z.B. 7.1)
    if ($FfmpegResponse.tag_name -match '(\d+\.\d+(?:\.\d+)?)') {
        $FfmpegVersion = $Matches[1]
        
        Write-To-ProgramList -Name "FFmpeg" -Version $FfmpegVersion -Bemerkung "Quelle: GitHub (GyanD Builds)"
        Write-Host " [OK: $FfmpegVersion]" -ForegroundColor Green
    } 
    else {
        Write-Host " [FEHLER]" -ForegroundColor Red
        Write-Warning " FFmpeg Version konnte via GitHub API nicht extrahiert werden."
    }
} catch { 
    Write-Host " [FEHLER]" -ForegroundColor Red
    Write-Warning " Fehler bei FFmpeg: $($_.Exception.Message)" 
}

# ============================================================
# ENDE SOFTWARE_ABFRAGEN -------------------------------------
# ============================================================

# --- HTML GENERIERUNG (Das Dashboard) ---
function Export-To-Html {
    $GlobalResults = $GlobalResults | Sort-Object Programm
    $TimeNow = (Get-Date).AddHours(1).ToString("dd.MM.yyyy HH:mm")
    $Total = $GlobalResults.Count
    $UpToDate = ($GlobalResults | Where-Object { $_.Status -eq "OK" }).Count
    $Updates = ($GlobalResults | Where-Object { $_.Status -eq "UPDATE" }).Count
    $Percent = if ($Total -gt 0) { [math]::Round(($UpToDate / $Total) * 100) } else { 0 }

    $Html = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <title>Software Dashboard</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #121212; color: #e0e0e0; margin: 0; padding: 20px; }
        .container { max-width: 1100px; margin: auto; background: #1e1e1e; padding: 30px; border-radius: 12px; box-shadow: 0 4px 30px rgba(0,0,0,0.5); }
        h1 { color: #ffffff; margin-top: 0; }
        .info-bar { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; font-size: 0.9em; color: #aaa; }
        .status-badge { padding: 4px 10px; border-radius: 20px; font-size: 0.85em; font-weight: bold; text-transform: uppercase; display: inline-block; }
        .status-OK { background-color: #1b4332; color: #75f0a0; }
        .status-UPDATE { background-color: #641220; color: #ffb3c1; border: 1px solid #801b2d; animation: pulse 2s infinite; }
        .status-NEU { background-color: #5e503f; color: #ffd60a; }
        @keyframes pulse { 0% { opacity: 1; } 50% { opacity: 0.5; } 100% { opacity: 1; } }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; }
        th { background-color: #2d2d2d; color: #ffffff; text-align: left; padding: 12px; border-bottom: 2px solid #444; }
        td { padding: 12px; border-bottom: 1px solid #333; }
        tr:hover { background-color: #2a2a2a; }
        .btn-accept { background-color: #2b5797; color: white; border: none; padding: 6px 12px; border-radius: 4px; cursor: pointer; font-size: 0.8em; margin-left: 10px; transition: background 0.3s; }
        .btn-accept:hover { background-color: #3e78cc; }
        .btn-accept:disabled { background-color: #444; cursor: not-allowed; opacity: 0.5; }
        .progress-container { background: #333; border-radius: 10px; height: 12px; margin: 15px 0; overflow: hidden; }
        .progress-bar { background: #4ade80; height: 100%; transition: width 1s; box-shadow: 0 0 10px rgba(74, 222, 128, 0.3); }
		.btn-add { background-color: #d4a017; /* Gold/Gelb */ color: black; border: none; padding: 6px 12px; border-radius: 4px; cursor: pointer; font-size: 0.8em; margin-left: 10px; font-weight: bold; transition: background 0.3s; }
		.btn-add:hover { background-color: #ffd60a; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Software Monitor (franz@vobs.at - Jan26)</h1>
        <div class="info-bar">
            <span>Stand: <strong>$TimeNow Uhr</strong></span>
            <span>$UpToDate von $Total Programmen aktuell</span>
        </div>
        
        <div class="progress-container">
            <div class="progress-bar" style="width: $($Percent)%"></div>
        </div>

        <table>
            <thead>
                <tr>
                    <th>Programm</th>
                    <th>IST</th>
                    <th>AKTUELL</th>
                    <th>Status</th>
                </tr>
            </thead>
            <tbody>
"@

# Jedes Programm wird hier verarbeitet
foreach ($R in $GlobalResults) {
    $BadgeClass = "status-" + $R.Status
    $StatusZelle = "<span class='status-badge $BadgeClass'>$($R.Status)</span>"

    # Logik für Buttons:
    if ($R.Status -eq "UPDATE") {
        # Button für bestehende Programme (Update)
        $StatusZelle += " <button class='btn-accept' onclick='acceptUpdate(`"$($R.Programm)`", `"$($R.AKTUELL)`")'>Übernehmen</button>"
    } 
    elseif ($R.Status -eq "NEU") {
        # NEUER BUTTON für neue Programme (Hinzufügen)
        $StatusZelle += " <button class='btn-add' onclick='acceptUpdate(`"$($R.Programm)`", `"$($R.AKTUELL)`")'>Hinzufügen</button>"
    }

    $Html += @"
            <tr>
                <td>$($R.Programm)</td>
                <td>$($R.IST)</td>
                <td>$($R.AKTUELL)</td>
                <td>$StatusZelle</td>
            </tr>
"@
}

    $Html += @"
            </tbody>
        </table>
        <p style="margin-top: 30px; font-size: 0.8em; color: #adb5bd; text-align: center;">GitHub Actions Automatisierung</p>
    </div>

<script>
async function acceptUpdate(softwareName, newVersion) {
    const token = localStorage.getItem('github_token') || prompt("Bitte gib deinen GitHub Personal Access Token ein:");
    if (!token) return;
    localStorage.setItem('github_token', token);

    const button = event.target;
    button.disabled = true;
    button.innerText = "...";

    const url = 'https://api.github.com/repos/loufrance/Software-Monitor/dispatches';
    
    try {
        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'Authorization': 'Bearer ' + token, 
                'Accept': 'application/vnd.github.v3+json',
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                event_type: 'update_software_csv',
                client_payload: {
                    name: softwareName,
                    version: newVersion
                }
            })
        });

        if (response.ok) {
            button.parentElement.innerHTML = '<span class="status-badge status-OK" style="background-color: #2b5797;">Wird aktualisiert...</span>';
        } else {
            alert("Fehler: " + response.status);
            button.disabled = false;
            button.innerText = "Übernehmen";
        }
    } catch (e) {
        alert("Verbindungsfehler.");
        button.disabled = false;
    }
}
</script>
</body>
</html>
"@
    $Html | Out-File -FilePath $HtmlPath -Encoding UTF8
}
Export-To-Html
