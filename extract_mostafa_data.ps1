$files = Get-ChildItem "f:\black gold\Student Analysis designed html\Mostafa's Students" -Filter "placementtest*.xlsx" | Where-Object { $_.Name -notlike "~$*" }
$outputData = @()
$tempDir = Join-Path $env:TEMP "mostafa_extraction"
if (Test-Path $tempDir) { Remove-Item -Recurse -Force $tempDir -ErrorAction SilentlyContinue }
New-Item -ItemType Directory -Force -Path $tempDir | Out-Null

foreach ($file in $files) {
    Write-Host "Processing $($file.Name)..."
    
    # 1. Isolated Processing
    # Create valid temp path
    $baseName = $file.BaseName
    $fileTempDir = Join-Path $tempDir $baseName
    New-Item -ItemType Directory -Force -Path $fileTempDir | Out-Null
    $htmlFile = Join-Path $fileTempDir "export.htm"
    
    # Copy file to temp to avoid locks/path issues
    $tempXlsx = Join-Path $fileTempDir "temp_source.xlsx"
    Copy-Item $file.FullName $tempXlsx -Force

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
        continue
    }
    
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Start-Sleep -Milliseconds 200 # Brief pause
    
    # 2. Analyze Styles (Same logic as before)
    $supportDir = Join-Path $fileTempDir "export_files"
    $cssFile = Join-Path $supportDir "stylesheet.css"
    $sheetFile = Join-Path $supportDir "sheet001.htm"
    
    if (-not (Test-Path $sheetFile)) {
        # Try sheet002.htm just in case
        $sheetFile2 = Join-Path $supportDir "sheet002.htm"
        if (Test-Path $sheetFile2) {
            $sheetFile = $sheetFile2
            Write-Host "Using sheet002.htm for $($file.Name)"
        }
        else {
            Write-Warning "Sheet file not found for $($file.Name)"
            continue
        }
    }
    
    # Parse CSS
    $cssContent = Get-Content $cssFile -Raw
    $greenClasses = @()
    $cssMatches = [Regex]::Matches($cssContent, "\.(xl\d+)[^\{]*\{[^\}]*background:#D9F7ED", "IgnoreCase")
    foreach ($m in $cssMatches) {
        $greenClasses += $m.Groups[1].Value
    }
    # Write-Host "  Green Classes: $($greenClasses.Count)"
    
    # 3. Parse Sheet HTML
    $sheetContent = Get-Content $sheetFile -Raw
    
    $htmlDoc = New-Object -ComObject "HTMLFile"
    $htmlDoc.IHTMLDocument2_write($sheetContent)
    
    $table = $htmlDoc.getElementsByTagName("table") | Select-Object -First 1
    if ($null -eq $table) {
        Write-Warning "No table found in $($file.Name)"
        continue
    }
    $rows = $table.rows
    Write-Host "  Rows: $($rows.length)"
    
    $studentMap = @{} # ColIndex -> StudentName
    
    $headerRow = $rows.item(0) # Row 1
    $cells = $headerRow.cells
    
    for ($i = 12; $i -lt $cells.length; $i++) {
        $name = $cells.item($i).innerText
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            Write-Host "  Found Student Candidate: $name (Col $i)"
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
    
    # Track Question Number manually based on row index since we know the structure
    # Row 1 (index 0) is Header. Row 2 (index 1) is Q1. So Q_Num = $r.
    
    for ($r = 1; $r -lt $rows.length; $r++) {
        $row = $rows.item($r)
        
        # Check if row is valid for students
        if ($row.cells.length -lt 2) { continue }
        
        $qCell = $row.cells.item(1)
        $qText = $qCell.innerText
        
        # Check Accuracy Row (heuristic: if first student col has % value)
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
        
        # Question Number Logic
        $currentQNum = $r
        
        # Handle empty/missing questions (especially Q89-100)
        # We process them so they are counted in the total (100 Qs)
        if ($currentQNum -gt 100) { continue }
        
        # Handle empty/missing questions (especially Q89-100)
        # We process them so they are counted in the total (100 Qs)
        if ($currentQNum -ge 89 -and $currentQNum -le 100) {
            if ([string]::IsNullOrWhiteSpace($qText)) {
                $qText = "Question $currentQNum"
            }
        }
        
        if ([string]::IsNullOrWhiteSpace($qText)) { continue }
        
        # Classify Topic
        $topic = "Simple MCQ"

        # Regex Helpers
        $isVerbBe = $qText -match "\b(am|is|are|was|were)\b"
        $isPassive = $qText -match "passive|by\b"
        $isTense = $qText -match "yesterday|tomorrow|ago|next|usually|always|never|sometimes|will|did\b"
        $isGrammarAdv = $qText -match "perfect|had\s+\w+|have\s+\w+|has\s+\w+|if\s+|conditional"

        # Logic with precedence/ranges
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
            # Q88-100 fall here -> Simple MCQ (or generic)
            $topic = "Simple MCQ" 
        }

        # Override for explicit keywords
        if ($isPassive -and $topic -ne "Passive Voice" -and $currentQNum -lt 89) { $topic = "Passive Voice" }
        
        # Check Students
        foreach ($key in $studentMap.Keys) {
            if ($key -lt $row.cells.length) {
                $cell = $row.cells.item($key)
                $className = $cell.className
                
                $isCorrect = $false
                if ($greenClasses -contains $className) {
                    $isCorrect = $true
                }
                
                $studentMap[$key].Topics[$topic].Total++
                if ($isCorrect) {
                    $studentMap[$key].Topics[$topic].Correct++
                }
            }
        }
    }
    
    # Collect
    foreach ($s in $studentMap.Values) {
        $acc = $s.Accuracy
        $lvl = 0
        if ($acc -lt 55) { $lvl = 1 }
        elseif ($acc -le 69) { $lvl = 2 }
        elseif ($acc -le 79) { $lvl = 3 }
        else { $lvl = 4 }
        
        $existing = $outputData | Where-Object { $_.Name -eq $s.Name }
        if ($null -eq $existing) {
            $outputData += [PSCustomObject]@{
                Name     = $s.Name
                Accuracy = $acc
                Level    = $lvl
                Topics   = $s.Topics
            }
        }
    }
}

# Cleanup
Remove-Item -Recurse -Force $tempDir -ErrorAction SilentlyContinue

# Export
$outputData | Select-Object Name, Accuracy, Level | Export-Csv "f:\black gold\Student Analysis designed html\mostafa_students_summary.csv" -NoTypeInformation -Encoding UTF8

$jsonString = $outputData | ConvertTo-Json -Depth 5
$jsContent = "const studentData = $jsonString;"
Set-Content "f:\black gold\Student Analysis designed html\mostafa_dashboard_data.js" -Value $jsContent -Encoding UTF8
Write-Host "Done. Extracted $($outputData.Count) students."
