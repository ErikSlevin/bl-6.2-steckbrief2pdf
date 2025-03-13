function Set-StatusColor {
    param ([Parameter(Mandatory=$true)]$value)
    if ($value) { return 'Cyan' } else { return 'DarkGray' }
}

function Write-StatusLine {
    param (
        [Parameter(Mandatory=$true)]$label,
        [Parameter(Mandatory=$true)]$value,
        [Parameter(Mandatory=$true)]$color
    )
    $foregroundColor = if ($value) { 'White' } else { 'DarkGray' }
    Write-Host -ForegroundColor $foregroundColor $label -NoNewline
    Write-Host -ForegroundColor $color "$value"
}
function SetupVars {
    param (
        [bool]$lastLine = $true,  # Standardwert ist $true
        [switch]$Without          # Neuer Parameter ohne Ausgabe
    )
    Clear-Host
    $searchPathColor = Set-StatusColor -value $searchPath
    $pdfOutputDirColor = Set-StatusColor -value $pdfOutputDir
    $pdfcpuColor = Set-StatusColor -value $pdfcpu
    $deleteOldFilesColor = Set-StatusColor -value $deleteOldFiles
    Write-Host -ForegroundColor Green -NoNewline @"
    _____ ____   ____     _  _       ____  ____  _____
   | ____/ ___| / ___|   | || |     |  _ \/ ___|| ____|
   |  _| \___ \| |  _    | || |_    | | | \___ \|  _|
   | |___ ___) | |_| |   |__   _|   | |_| |___) | |___
   |_____|____/ \____|      |_|     |____/|____/|_____|
"@
Write-Host -ForegroundColor DarkGray " by HF Wahl"


    if (-not $Without) {
        Write-Host -ForegroundColor DarkGray "__________________________________________________________________`n"
        # Zeige die Eingaben dynamisch an
        Write-StatusLine -label "      Steckbriefe-Ordner Pfad: " -value $searchPathSummary -color $searchPathColor
        Write-StatusLine -label "              PDF-Ordner Pfad: " -value $pdfOutputDir -color $pdfOutputDirColor
        Write-StatusLine -label "              pdfcpu.exe Pfad: " -value $pdfcpu -color $pdfcpuColor
        Write-StatusLine -label "            AiO PDF + Anlagen: " -value $summary -color $pdfcpuColor
        Write-StatusLine -label "    Excel Steckbriefe löschen: " -value $deleteOldFiles -color $deleteOldFilesColor
        Write-Host ($null, "__________________________________________________________________")[$lastLine] -ForegroundColor DarkGray
    } else {
        Write-Host -ForegroundColor DarkGray "__________________________________________________________________`n"
    }
}

function ConvertToPDF {
    param (
        [string]$inputExcel,  # Pfad zur Excel-Datei
        [string]$OutputPDF     # Pfad zur Ausgabe-PDF
    )

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false

    try {
        $workbook = $excel.Workbooks.Open($inputExcel)
        $workbook.ExportAsFixedFormat(0,$OutputPDF)
        return $true
    } catch {
        Write-Host "❌ Fehler beim Konvertieren: $_"
        return $false
    } finally {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
}

# Setze alle Variablen auf Null
$searchPath, $searchPathSummary, $pdfOutputDir, $pdfcpu, $summary, $Deckblatt, $deleteOldFiles = "", "", "", "", "", "", ""

Clear-Host
SetupVars # Statusanzeige

# Initiale Eingaben
Write-Host -ForegroundColor DarkCyan "`n📄 Hinweistext: Verzeichnis der Excel-Dateien, die aus dem Masterdatenblatt abgeleitet wurden"
$searchPath = (Read-Host "📂 Bitte geben Sie den Suchpfad ein").Trim().Replace('"', '')
$pdfOutputDir = $searchPath + "\000_Generated_PDF\Steckbriefe"

$regexPattern = "\d{8}\.xlsx$"
$steckbriefe = Get-ChildItem -Path $searchPath -Filter "*.xlsx" -Recurse | Where-Object { $_.Name -match $regexPattern }
$searchPathSummary = "$searchPath ($($steckbriefe.Count) Steckbriefe_YYYYMMDD gefunden)"
SetupVars  # Statusanzeige

Write-Host -ForegroundColor DarkCyan "`n📄  Sollen alle PDF-Steckbriefe zusätzlich in einer AiO-Steckbriefdatei zusammengeführt werden und als Anhang? (Empfohlen)"
$summary = Read-Host "📂  Zusammenfassung aller Steckbriefe? (Ja/Nein)"
if ($summary -eq "Ja") {
    Write-Host -ForegroundColor DarkCyan "`n📄  Es muss ein (absoluter) Pfad zu einer PDF-Datei angegeben werden, die als Deckblatt dient. Die Steckbriefe werden in dieser als Anlage hinzugefügt"
    $Deckblatt = (Read-Host "📂  Pfad zum Steckbrief-Deckblatt").Trim().Replace('"', '')
    $summary = "$pdfOutputDir\000_AiO_mit_Anlagen_HF_WAHL.pdf"

    Write-Host -ForegroundColor DarkCyan "`n📄  Das Tool pdfcpu (https://github.com/pdfcpu/pdfcpu) wird für die Umwandlung benötigt. Bitte geben Sie den vollständigen (absoluten) Pfad zu der pdfcpu.exe an"
    $pdfcpu = (Read-Host "📂  Pfad zu der EXE-Datei").Trim().Replace('"', '')
} else {
    $summary = "Nein"
    $pdfcpu = "Wird nicht benötigt"
}
SetupVars  # Statusanzeige

Write-Host -ForegroundColor DarkCyan "`n📄  Möchten Sie die durch das Masterdatenblatt abgeleiteten Excel-Steckbriefe (im Format _YYYYMMDD) nach Abschluss des Vorgangs löschen? (Nicht empfohlen)"
$deleteOldFiles = Read-Host "📂  Generierte Excel-Steckbriefe im Anschluss löschen? (Ja/Nein)"
SetupVars -lastLine $false  # Statusanzeige

$steckbriefeArray_Old = @()
if ($steckbriefe.Count -ge 0) {
    Write-Host -ForegroundColor Cyan "_________________ $($steckbriefe.Count) Gefundene Steckbriefe _______________________`n"
    $steckbriefe.Name | ForEach-Object { 
        Write-Host "📄 $_" 
        $steckbriefeArray_Old += "$_"
    }
    $Deckblatt = (Read-Host "`n❔ Sollen diese Steckbriefe in PDFs umgewandelt werden? (Ja/Nein)")
    if ($Deckblatt -ne "Ja") {exit}
} else {
    Write-Host "`n❌ Keine Steckbriefe mit Datumsformat YYYYMMDD gefunden!"
    exit
}

SetupVars -lastLine $false  # Statusanzeige
$i = 0
$err = @()
$steckbriefeArray_New = $steckbriefeArray_Old
# Konvertiere Steckbriefe in PDF
foreach ($steckbrief in $steckbriefe) {
    SetupVars -Without  # Statusanzeige

    Write-Host -ForegroundColor DarkCyan "📄  Konvertiere $i von $($steckbriefe.count) `n"
    $steckbriefeArray_New

    # Versuche, ein Datum im Format YYYYMMDD aus dem Dateinamen zu extrahieren
    if ($steckbrief.BaseName -notmatch "(\d{4})(\d{2})(\d{2})$") {
        Write-Host -ForegroundColor Red "❌ Überspringe $($steckbrief.Name) – Kein gültiges Datum (YYYYMMDD) gefunden!"
        continue  # Überspringe diese Datei und fahre mit der nächsten fort
    }
    $datefolder = "$($matches[1])-$($matches[2])-$($matches[3])"

    # Setze den Zielpfad für das Verzeichnis, das nur einmal korrekt gesetzt wird
    $pdfTargetDir = Join-Path -Path $pdfOutputDir -ChildPath $dateFolder

    # Falls das Verzeichnis nicht existiert, wird es erstellt
    New-Item -ItemType Directory -Path $pdfTargetDir -Force | Out-Null

    # Versuche, die restlichen Dateiinformationen aus dem Dateinamen zu extrahieren
    if ($steckbrief.BaseName -notmatch "(\d{3})_(.*?)_(\d{8})$") {
        Write-Host -ForegroundColor Red "❌ Überspringe $($steckbrief.Name) – Format passt nicht!"
        continue  # Überspringe diese Datei und fahre mit der nächsten fort
    }
    
    # Setze den neuen Dateinamen für das PDF (Beispiel: 123_Name.pdf)
    $SteckbriefFileName = "$($matches[1])_$($matches[2]).pdf"
    
    # Setze den vollständigen Zielpfad für das PDF
    $pdfOutputFile = Join-Path -Path $pdfTargetDir -ChildPath $SteckbriefFileName
    $conversionResult = ConvertToPDF -inputExcel $steckbrief.FullName -OutputPDF $pdfOutputFile
    if ($conversionResult) {
        $index = $steckbriefeArray_New.IndexOf($steckbrief.Name)
        $steckbriefeArray_New[$index] = "✅  $($steckbrief.Name) --> kopiert: $pdfOutputFile"
        $i++
    } else {
        Write-Host "❌ Fehler: Konvertierung von $($steckbrief.FullName) fehlgeschlagen! "
    }
}

SetupVars -Without  # Statusanzeige

Write-Host -ForegroundColor DarkCyan "📄  $i von $($steckbriefe.count) `n"

$steckbriefeArray_New