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



# ============================================================
# ENDE SOFTWARE_ABFRAGEN -------------------------------------
# ============================================================

# --- HTML GENERIERUNG (Das Dashboard) ---
function Export-To-Html {
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
    <title>Software Monitor</title>
    <style>
        body { font-family: sans-serif; background: #f0f2f5; padding: 20px; }
        .card { background: white; padding: 20px; border-radius: 10px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); max-width: 1000px; margin: auto; }
        .progress-bg { background: #eee; height: 20px; border-radius: 10px; margin: 15px 0; }
        .progress-fill { background: #27ae60; height: 100%; width: $($Percent)%; border-radius: 10px; transition: width 1s; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th { text-align: left; background: #2c3e50; color: white; padding: 12px; }
        td { padding: 12px; border-bottom: 1px solid #eee; }
        .status-UPDATE { color: #e74c3c; font-weight: bold; background: #ffebeb; padding: 4px; border-radius: 4px; }
        .status-OK { color: #27ae60; font-weight: bold; }
    </style>
</head>
<body>
    <div class="card">
        <h1>🚀 Software Version Dashboard</h1>
        <p>Letzte Prüfung: <strong>$TimeNow Uhr</strong> (Automatisch alle 60 Min)</p>
        <div class="progress-bg"><div class="progress-fill"></div></div>
        <p>Status: $UpToDate von $Total Programmen sind aktuell ($Percent%)</p>
        <table>
            <tr><th>Programm</th><th>IST</th><th>AKTUELL</th><th>Status</th></tr>
"@
    foreach ($R in $GlobalResults) {
        $Html += "<tr><td>$($R.Programm)</td><td>$($R.IST)</td><td>$($R.AKTUELL)</td><td><span class='status-$($R.Status)'>$($R.Status)</span></td></tr>"
    }
    $Html += "</table></div></body></html>"
    $Html | Out-File -FilePath $HtmlPath -Encoding UTF8
}
Export-To-Html
