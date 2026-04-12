# NetSuite Import Processor - Version 5
# Reads xlsx file as ZIP archive with robust error handling

$ErrorActionPreference = "Stop"

$csvFile = "C:\Users\matan\OneDrive\desktop\csv.csv"
$mappingFile = "C:\Users\matan\OneDrive\desktop\ST mapping.xlsx"
$outputPath = "C:\Users\matan\OneDrive\desktop\NetSuite_to_be_imported.csv"
$logFile = "C:\Users\matan\OneDrive\desktop\NetSuite import\process_log.txt"

# Clear log file
"" | Out-File -FilePath $logFile -Encoding UTF8

function Write-Log($msg, $color = "White") {
    Write-Host $msg -ForegroundColor $color
    $msg | Out-File -FilePath $logFile -Append -Encoding UTF8
    [Console]::Out.Flush()
}

Write-Log "NetSuite Import Processor v5" "Cyan"
Write-Log ("=" * 50)

try {
    Write-Log "`nStep 1: Loading mapping file..." "Yellow"
    
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    
    # Copy the file using stream-based approach
    $tempFile = [System.IO.Path]::GetTempFileName() + ".xlsx"
    Write-Log "Copying to: $tempFile" "Gray"
    
    $sourceStream = [System.IO.File]::Open($mappingFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    $destStream = [System.IO.File]::Create($tempFile)
    $sourceStream.CopyTo($destStream)
    $destStream.Close()
    $sourceStream.Close()
    Write-Log "File copied" "Gray"
    
    # Open ZIP
    Write-Log "Opening ZIP..." "Gray"
    $zip = [System.IO.Compression.ZipFile]::OpenRead($tempFile)
    Write-Log "ZIP opened, entries: $($zip.Entries.Count)" "Gray"
    
    # Read shared strings
    Write-Log "Reading shared strings..." "Gray"
    $sharedStringsEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/sharedStrings.xml" }
    $sharedStrings = @()
    if ($sharedStringsEntry) {
        $reader = [System.IO.StreamReader]::new($sharedStringsEntry.Open())
        $content = $reader.ReadToEnd()
        $reader.Close()
        $sharedStringsXml = [xml]$content
        $sharedStrings = @($sharedStringsXml.sst.si | ForEach-Object { 
            if ($_.t -is [string]) { $_.t } 
            elseif ($_.t) { $_.t.'#text' }
            else { "" }
        })
        Write-Log "Loaded $($sharedStrings.Count) shared strings" "Gray"
    }
    
    # Read sheet
    Write-Log "Reading sheet1.xml..." "Gray"
    $sheetEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/worksheets/sheet1.xml" }
    if (-not $sheetEntry) { throw "Could not find sheet1.xml" }
    
    $reader = [System.IO.StreamReader]::new($sheetEntry.Open())
    $sheetContent = $reader.ReadToEnd()
    $reader.Close()
    $sheetXml = [xml]$sheetContent
    
    $zip.Dispose()
    Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
    
    Write-Log "Sheet loaded" "Gray"
    
    # Get rows
    $rows = @($sheetXml.worksheet.sheetData.row)
    Write-Log "Found $($rows.Count) rows" "Gray"
    
    # Helper functions
    function Get-CellValue($cell, $strings) {
        if (-not $cell) { return $null }
        $value = $cell.v
        if ($cell.t -eq "s" -and $strings.Count -gt 0 -and $value) {
            $index = [int]$value
            if ($index -lt $strings.Count) {
                return $strings[$index]
            }
        }
        return $value
    }
    
    function Get-ColumnLetter($cellRef) {
        return ($cellRef -replace '\d+', '')
    }
    
    # Parse header
    Write-Log "Parsing header row..." "Gray"
    $headerRow = $rows[0]
    $headerCells = @($headerRow.c)
    Write-Log "Header has $($headerCells.Count) cells" "Gray"
    
    $internalCol = $null
    $intercompanyCol = $null
    
    foreach ($cell in $headerCells) {
        $colLetter = Get-ColumnLetter $cell.r
        $value = Get-CellValue $cell $sharedStrings
        Write-Log "  $colLetter : '$value'" "Gray"
        
        if ($value) {
            $valueLower = $value.ToString().Trim().ToLower()
            # Match "internal" or "ls internal"
            if ($valueLower -like "*internal") {
                $internalCol = $colLetter
                Write-Log "    -> Found Internal column!" "Green"
            }
            elseif ($valueLower -eq "intercompany") {
                $intercompanyCol = $colLetter
                Write-Log "    -> Found Intercompany column!" "Green"
            }
        }
    }
    
    Write-Log "After header parsing: Internal=$internalCol, Intercompany=$intercompanyCol" "Gray"
    
    if (-not $internalCol) { throw "Could not find 'Internal' column" }
    if (-not $intercompanyCol) { throw "Could not find 'Intercompany' column" }
    
    Write-Log "Internal column: $internalCol" "Green"
    Write-Log "Intercompany column: $intercompanyCol" "Green"
    
    # Build mapping
    Write-Log "`nStep 2: Building mapping dictionary..." "Yellow"
    $mappingDict = @{}
    $totalRows = $rows.Count - 1
    Write-Log "Processing $totalRows data rows..." "Gray"
    
    for ($i = 1; $i -lt $rows.Count; $i++) {
        try {
            $row = $rows[$i]
            $internalValue = $null
            $intercompanyValue = $null
            
            $rowCells = @($row.c)
            foreach ($cell in $rowCells) {
                $colLetter = Get-ColumnLetter $cell.r
                $value = Get-CellValue $cell $sharedStrings
                
                if ($colLetter -eq $internalCol) { $internalValue = $value }
                elseif ($colLetter -eq $intercompanyCol) { $intercompanyValue = $value }
            }
            
            if ($null -ne $internalValue -and $internalValue.ToString().Trim() -ne "") {
                $isIC = $intercompanyValue -and $intercompanyValue.ToString().Trim().ToLower() -eq "yes"
                $mappingDict[$internalValue.ToString()] = $isIC
            }
            
            if ($i % 100 -eq 0) { Write-Log "  Row $i/$totalRows..." "Gray" }
        } catch {
            Write-Log "  Error at row $i : $_" "Red"
        }
    }
    Write-Log "Mapping loop completed" "Gray"
    
    Write-Log "Loaded $($mappingDict.Count) mappings" "Green"
    $icCount = ($mappingDict.Values | Where-Object { $_ -eq $true }).Count
    Write-Log "  I/C accounts: $icCount" "Gray"
    Write-Log "  Non-I/C accounts: $($mappingDict.Count - $icCount)" "Gray"
    
    # Sample I/C accounts
    Write-Log "`nSample I/C accounts:" "Gray"
    $mappingDict.GetEnumerator() | Where-Object { $_.Value -eq $true } | Select-Object -First 5 | ForEach-Object {
        Write-Log "  $($_.Key)" "Gray"
    }
    
    # Load CSV
    Write-Log "`nStep 3: Loading CSV file..." "Yellow"
    $csvData = @(Import-Csv -Path $csvFile)
    Write-Log "Loaded $($csvData.Count) rows" "Green"
    
    # Get column names
    $headers = $csvData[0].PSObject.Properties.Name
    $internalColName = $headers | Where-Object { $_.Trim().ToLower() -eq "internal" } | Select-Object -First 1
    $subsidiaryColName = $headers | Where-Object { $_.Trim().ToLower() -eq "subsidiary" } | Select-Object -First 1
    
    if (-not $internalColName) { throw "Could not find 'Internal' column in CSV" }
    if (-not $subsidiaryColName) { throw "Could not find 'Subsidiary' column in CSV" }
    
    Write-Log "CSV Internal: '$internalColName'" "Gray"
    Write-Log "CSV Subsidiary: '$subsidiaryColName'" "Gray"
    
    # Process rows
    Write-Log "`nStep 4: Processing rows..." "Yellow"
    
    $processedRows = [System.Collections.ArrayList]::new()
    $icTransactions = [System.Collections.ArrayList]::new()
    $rowNum = 0
    
    foreach ($row in $csvData) {
        $rowNum++
        $internalValue = $row.$internalColName.ToString().Trim()
        
        $isIC = $false
        if ($mappingDict.ContainsKey($internalValue)) {
            $isIC = $mappingDict[$internalValue]
        }
        
        $newRow = [ordered]@{}
        foreach ($prop in $row.PSObject.Properties) {
            $newRow[$prop.Name] = $prop.Value
        }
        $newRow["Eliminate"] = if ($isIC) { "Yes" } else { "No" }
        $newRow["To Subsidiary / From Subsidiary"] = "Lsports Data Ltd"
        
        $rowObj = [PSCustomObject]$newRow
        [void]$processedRows.Add($rowObj)
        
        if ($isIC) {
            [void]$icTransactions.Add($rowObj)
        }
        
        if ($rowNum % 10000 -eq 0) {
            Write-Log "  Processed $rowNum rows..." "Gray"
        }
    }
    
    Write-Log "Found $($icTransactions.Count) I/C transactions" "Green"
    
    # Duplicate I/C transactions
    Write-Log "`nStep 5: Duplicating I/C transactions..." "Yellow"
    foreach ($icRow in $icTransactions) {
        $dupRow = [ordered]@{}
        foreach ($prop in $icRow.PSObject.Properties) {
            $dupRow[$prop.Name] = $prop.Value
        }
        
        if ($dupRow[$subsidiaryColName].Trim().ToLower() -eq "statscore") {
            $dupRow[$subsidiaryColName] = "Lsports Data Ltd"
        }
        
        $dupRow["To Subsidiary / From Subsidiary"] = "Statscore"
        
        [void]$processedRows.Add([PSCustomObject]$dupRow)
    }
    
    Write-Log "Final row count: $($processedRows.Count)" "Green"
    
    # Save
    Write-Log "`nStep 6: Saving to CSV..." "Yellow"
    $processedRows | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
    Write-Log "Saved to: $outputPath" "Green"
    
    Write-Log ("`n" + ("=" * 50)) "Cyan"
    Write-Log "DONE!" "Cyan"
    Write-Log "Output: $outputPath" "Green"
    Write-Log "`nNote: Duplicate I/C rows are at the bottom." "Yellow"
    
} catch {
    Write-Log "`nERROR: $_" "Red"
    Write-Log $_.ScriptStackTrace "Red"
}

