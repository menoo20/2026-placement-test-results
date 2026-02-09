# Robust Extraction Script for Main Dashboard (User's Students)
# Extracts from root folder, excludes 'Mostafa's Students' (by non-recursive), handles COM errors

# Extracts from "Mohammed Ameen's students" folder
# Handles COM errors and ensures robust extraction

$rootPath = "f:\black gold\Student Analysis designed html\Mohammed Ameen's students"
$files = Get-ChildItem $rootPath -Filter "placementtest*.xlsx" | Where-Object { $_.Name -notlike "~$*" }
$outputData = @()
$tempDir = Join-Path $env:TEMP "main_extraction"
if (Test-Path $tempDir) { Remove-Item -Recurse -Force $tempDir -ErrorAction SilentlyContinue }
New-Item -ItemType Directory -Force -Path $tempDir | Out-Null

Write-Host "Found $($files.Count) files to process."

foreach ($file in $files) {
    Write-Host "Processing $($file.Name)..."
    $baseName = $file.BaseName
    $fileTempDir = Join-Path $tempDir $baseName
    New-Item -ItemType Directory -Force -Path $fileTempDir | Out-Null
    
    $htmlFile = Join-Path $fileTempDir "export.htm"
    $tempXlsx = Join-Path $fileTempDir "temp_source.xlsx"
    
    # Copy to temp to avoid lock/modification issues
    Copy-Item $file.FullName $tempXlsx -Force
    
    # 1. Convert to HTML using fresh Excel instance per file (Robustness)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    try {
        $workbook = $excel.Workbooks.Open($tempXlsx, $null, $true) # ReadOnly
        $workbook.SaveAs($htmlFile, 44) # xlHtml
        $workbook.Close($false)
    }
    catch {
        Write-Error "Failed to convert $($file.Name): $_"
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        Remove-Variable excel -ErrorAction SilentlyContinue
        continue
    }
    
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Remove-Variable excel -ErrorAction SilentlyContinue
    
    # 2. Analyze Styles
    $supportDir = Join-Path $fileTempDir "export_files"
    $cssFile = Join-Path $supportDir "stylesheet.css"
    $sheetFile = Join-Path $supportDir "sheet001.htm"
    
    if (-not (Test-Path $sheetFile)) {
        Write-Warning "Sheet file not found for $($file.Name)"
        continue
    }
    
    # Parse CSS
    $cssContent = Get-Content $cssFile -Raw
    $greenClasses = @()
    $cssMatches = [Regex]::Matches($cssContent, "\.(xl\d+)[^\{]*\{[^\}]*background:#D9F7ED", "IgnoreCase")
    foreach ($m in $cssMatches) {
        $greenClasses += $m.Groups[1].Value
    }
    
    # 3. Parse Sheet HTML
    $sheetContent = Get-Content $sheetFile -Raw
    $htmlDoc = New-Object -ComObject "HTMLFile"
    $htmlDoc.IHTMLDocument2_write($sheetContent)
    
    $table = $htmlDoc.getElementsByTagName("table") | Select-Object -First 1
    $rows = $table.rows
    
    # Student Map
    $studentMap = @{} 
    $headerRow = $rows.item(0) 
    $cells = $headerRow.cells
    
    for ($i = 12; $i -lt $cells.length; $i++) {
        $name = $cells.item($i).innerText
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            $studentMap[$i] = @{
                Name     = $name
                ColIndex = $i
                Topics   = @{
                    "Simple MCQ"          = @{Total = 0; Correct = 0 }
                    "Verb to Be"          = @{Total = 0; Correct = 0 }
                    "Tenses"              = @{Total = 0; Correct = 0 }
                    "Passive Voice"       = @{Total = 0; Correct = 0 }
                    "Basic Vocabulary"    = @{Total = 0; Correct = 0 }
                    "Advanced Vocabulary" = @{Total = 0; Correct = 0 }
                    "Advanced Grammar"    = @{Total = 0; Correct = 0 }
                }
                Accuracy = 0 
            }
        }
    }
    
    # Iterate Rows
    for ($r = 1; $r -lt $rows.length; $r++) {
        $row = $rows.item($r)
        if ($row.cells.length -lt 2) { continue }
        
        $qCell = $row.cells.item(1)
        $qText = $qCell.innerText
        
        # Check Accuracy Row
        $firstStudentCol = $studentMap.Keys | Sort-Object | Select-Object -First 1
        if ($null -ne $firstStudentCol -and $row.cells.length -gt $firstStudentCol) {
            $val = $row.cells.item($firstStudentCol).innerText
            if ($val -match "^\d+%$") {
                foreach ($key in $studentMap.Keys) {
                    if ($key -lt $row.cells.length) {
                        $accText = $row.cells.item($key).innerText
                        $studentMap[$key].Accuracy = [int]($accText -replace "%", "")
                    }
                }
                continue 
            }
        }
        
        $currentQNum = $r
        
        if ($currentQNum -gt 100) { continue }
        
        # Handle empty/missing for Q89-100
        if ($currentQNum -ge 89 -and $currentQNum -le 100) {
            if ([string]::IsNullOrWhiteSpace($qText)) {
                $qText = "Question $currentQNum"
            }
        }
        
        if ([string]::IsNullOrWhiteSpace($qText)) { continue }
        
        # Classify
        $topic = "Simple MCQ"

        $isVerbBe = $qText -match "\b(am|is|are|was|were)\b"
        $isPassive = $qText -match "passive|by\b"
        $isTense = $qText -match "yesterday|tomorrow|ago|next|usually|always|never|sometimes|will|did\b"
        
        if ($currentQNum -le 5) { 
            $topic = "Verb to Be" 
        }
        elseif ($currentQNum -le 32) { 
            $topic = "Tenses"
            if ($qText -match "tag") { $topic = "Advanced Grammar" }
        }
        elseif ($currentQNum -le 50) { 
            $topic = "Basic Vocabulary"
            if ($qText -match "active|passive") { $topic = "Passive Voice" }
        }
        elseif ($currentQNum -le 54) {
            $topic = "Advanced Grammar" 
        }
        elseif ($currentQNum -le 60) { 
            $topic = "Advanced Grammar" 
        }
        elseif ($currentQNum -le 62) { 
            $topic = "Passive Voice" 
        }
        elseif ($currentQNum -le 67) { 
            $topic = "Simple MCQ" 
        }
        elseif ($currentQNum -le 87) { 
            $topic = "Advanced Vocabulary"
        }
        else {
            $topic = "Simple MCQ"
        }

        if ($isPassive -and $topic -ne "Passive Voice" -and $currentQNum -lt 89) { $topic = "Passive Voice" }
        
        # Count Stats
        foreach ($key in $studentMap.Keys) {
            if ($key -lt $row.cells.length) {
                $cell = $row.cells.item($key)
                $className = $cell.className
                $isCorrect = $false
                if ($greenClasses -contains $className) { $isCorrect = $true }
                
                $studentMap[$key].Topics[$topic].Total++
                if ($isCorrect) { $studentMap[$key].Topics[$topic].Correct++ }
            }
        }
    }
    
    # Add to output
    foreach ($s in $studentMap.Values) {
        $acc = $s.Accuracy
        $lvl = 0
        if ($acc -lt 55) { $lvl = 1 }
        elseif ($acc -le 69) { $lvl = 2 }
        elseif ($acc -le 79) { $lvl = 3 }
        else { $lvl = 4 }
        
        $outputData += [PSCustomObject]@{
            Name     = $s.Name
            Accuracy = $acc
            Level    = $lvl
            Topics   = $s.Topics
        }
    }
    
    # Manual Cleanup
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($htmlDoc) | Out-Null
    Remove-Variable htmlDoc -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 200
}

# Cleanup Temp
Remove-Item -Recurse -Force $tempDir -ErrorAction SilentlyContinue

# Export
$outputData | Select-Object Name, Accuracy, Level | Export-Csv "f:\black gold\Student Analysis designed html\students_summary.csv" -NoTypeInformation -Encoding UTF8
$jsonString = $outputData | ConvertTo-Json -Depth 5
$jsContent = "const studentData = $jsonString;"
Set-Content "f:\black gold\Student Analysis designed html\dashboard_data.js" -Value $jsContent -Encoding UTF8

Write-Host "Done. Extracted $($outputData.Count) students."
