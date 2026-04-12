# NetSuite Import Processor - Version 4
# Reads xlsx file as ZIP archive (native PowerShell approach)

$csvFile = "C:\Users\matan\OneDrive\desktop\csv.csv"
$mappingFile = "C:\Users\matan\OneDrive\desktop\ST mapping.xlsx"
$outputPath = "C:\Users\matan\OneDrive\desktop\NetSuite_to_be_imported.csv"

Write-Host "NetSuite Import Processor v4" -ForegroundColor Cyan
Write-Host ("=" * 50)

try {
    # Read xlsx file as ZIP and parse XML
    Write-Host "`nLoading mapping file (ZIP/XML method)..." -ForegroundColor Yellow
    
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    
    # Copy the file using stream-based approach (works even when file is open)
    $tempFile = [System.IO.Path]::GetTempFileName() + ".xlsx"
    Write-Host "Copying mapping file to temp location (stream method)..." -ForegroundColor Gray
    
    try {
        # Use FileShare.ReadWrite to read even when file is locked
        $sourceStream = [System.IO.File]::Open($mappingFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $destStream = [System.IO.File]::Create($tempFile)
        $sourceStream.CopyTo($destStream)
        $destStream.Close()
        $sourceStream.Close()
        Write-Host "File copied successfully" -ForegroundColor Gray
    } catch {
        Write-Host "Stream copy failed: $_" -ForegroundColor Yellow
        Write-Host "Trying direct copy..." -ForegroundColor Yellow
        Copy-Item -Path $mappingFile -Destination $tempFile -Force
    }
    
    Write-Host "Opening ZIP archive..." -ForegroundColor Gray
    $zip = [System.IO.Compression.ZipFile]::OpenRead($tempFile)
    Write-Host "ZIP opened successfully, entries: $($zip.Entries.Count)" -ForegroundColor Gray
    
    # Get shared strings (for string values in cells)
    Write-Host "Looking for sharedStrings.xml..." -ForegroundColor Gray
    $sharedStringsEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/sharedStrings.xml" }
    $sharedStrings = @()
    if ($sharedStringsEntry) {
        $reader = [System.IO.StreamReader]::new($sharedStringsEntry.Open())
        $sharedStringsXml = [xml]$reader.ReadToEnd()
        $reader.Close()
        $sharedStrings = $sharedStringsXml.sst.si | ForEach-Object { 
            if ($_.t -is [string]) { $_.t } 
            else { $_.t.'#text' }
        }
        Write-Host "Loaded $($sharedStrings.Count) shared strings" -ForegroundColor Gray
    }
    
    # Get sheet data
    $sheetEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/worksheets/sheet1.xml" }
    $reader = [System.IO.StreamReader]::new($sheetEntry.Open())
    $sheetXml = [xml]$reader.ReadToEnd()
    $reader.Close()
    
    $zip.Dispose()
    
    # Clean up temp file
    Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
    
    # Parse rows
    $rows = $sheetXml.worksheet.sheetData.row
    Write-Host "Found $($rows.Count) rows in mapping file" -ForegroundColor Gray
    
    # Get cell value helper function
    function Get-CellValue($cell, $sharedStrings) {
        if ($null -eq $cell) { return $null }
        $value = $cell.v
        if ($cell.t -eq "s" -and $sharedStrings.Count -gt 0) {
            # It's a shared string index
            $index = [int]$value
            return $sharedStrings[$index]
        }
        return $value
    }
    
    # Get column letter from cell reference (e.g., "A1" -> "A")
    function Get-ColumnLetter($cellRef) {
        return ($cellRef -replace '\d+', '')
    }
    
    # Parse header row to find Internal and Intercompany columns
    $headerRow = $rows[0]
    Write-Host "Header row type: $($headerRow.GetType().Name)" -ForegroundColor Gray
    Write-Host "Header row cells count: $($headerRow.c.Count)" -ForegroundColor Gray
    $internalCol = $null
    $intercompanyCol = $null
    
    Write-Host "Header row cells:" -ForegroundColor Gray
    $cells = @($headerRow.c)
    Write-Host "Cells array count: $($cells.Count)" -ForegroundColor Gray
    foreach ($cell in $cells) {
        $colLetter = Get-ColumnLetter $cell.r
        $value = Get-CellValue $cell $sharedStrings
        Write-Host "  $colLetter : '$value'" -ForegroundColor Gray
        
        # Match "internal" or "LS internal" (case insensitive, contains match)
        if ($value -and $value.ToString().Trim().ToLower() -match "internal$") {
            $internalCol = $colLetter
        }
        elseif ($value -and $value.ToString().Trim().ToLower() -eq "intercompany") {
            $intercompanyCol = $colLetter
        }
    }
    
    if (-not $internalCol) { throw "Could not find 'Internal' or 'LS internal' column" }
    if (-not $intercompanyCol) { throw "Could not find 'Intercompany' column" }
    
    Write-Host "Internal column: $internalCol" -ForegroundColor Gray
    Write-Host "Intercompany column: $intercompanyCol" -ForegroundColor Gray
    
    # Build mapping dictionary
    Write-Host "Building mapping dictionary..." -ForegroundColor Gray
    $mappingDict = @{}
    
    for ($i = 1; $i -lt $rows.Count; $i++) {
        if ($i % 100 -eq 0) { Write-Host "  Processing mapping row $i..." -ForegroundColor Gray }
        $row = $rows[$i]
        $internalValue = $null
        $intercompanyValue = $null
        
        foreach ($cell in $row.c) {
            $colLetter = Get-ColumnLetter $cell.r
            $value = Get-CellValue $cell $sharedStrings
            
            if ($colLetter -eq $internalCol) {
                $internalValue = $value
            }
            elseif ($colLetter -eq $intercompanyCol) {
                $intercompanyValue = $value
            }
        }
        
        if ($null -ne $internalValue) {
            $isIC = $intercompanyValue -and $intercompanyValue.ToString().Trim().ToLower() -eq "yes"
            $mappingDict[$internalValue.ToString()] = $isIC
        }
    }
    
    Write-Host "Loaded $($mappingDict.Count) mappings" -ForegroundColor Green
    $icCount = ($mappingDict.Values | Where-Object { $_ -eq $true }).Count
    Write-Host "  I/C accounts: $icCount" -ForegroundColor Gray
    Write-Host "  Non-I/C accounts: $($mappingDict.Count - $icCount)" -ForegroundColor Gray
    
    # Show sample I/C accounts
    Write-Host "`nSample I/C accounts:" -ForegroundColor Gray
    $mappingDict.GetEnumerator() | Where-Object { $_.Value -eq $true } | Select-Object -First 5 | ForEach-Object {
        Write-Host "  $($_.Key) = I/C" -ForegroundColor Gray
    }
    
    # Load CSV
    Write-Host "`nLoading CSV file..." -ForegroundColor Yellow
    $csvData = Import-Csv -Path $csvFile
    Write-Host "Loaded $($csvData.Count) rows" -ForegroundColor Green
    
    # Get column names
    $headers = $csvData[0].PSObject.Properties.Name
    $internalColName = $headers | Where-Object { $_.Trim().ToLower() -eq "internal" } | Select-Object -First 1
    $subsidiaryColName = $headers | Where-Object { $_.Trim().ToLower() -eq "subsidiary" } | Select-Object -First 1
    
    if (-not $internalColName) { throw "Could not find 'Internal' column in CSV" }
    if (-not $subsidiaryColName) { throw "Could not find 'Subsidiary' column in CSV" }
    
    Write-Host "CSV Internal column: '$internalColName'" -ForegroundColor Gray
    Write-Host "CSV Subsidiary column: '$subsidiaryColName'" -ForegroundColor Gray
    
    # Process rows
    Write-Host "`nProcessing rows..." -ForegroundColor Yellow
    
    $processedRows = [System.Collections.ArrayList]::new()
    $icTransactions = [System.Collections.ArrayList]::new()
    $rowNum = 0
    
    foreach ($row in $csvData) {
        $rowNum++
        $internalValue = $row.$internalColName.ToString().Trim()
        
        # Determine if I/C
        $isIC = $false
        if ($mappingDict.ContainsKey($internalValue)) {
            $isIC = $mappingDict[$internalValue]
        }
        
        # Create new row with added columns
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
            Write-Host "  Processed $rowNum rows..." -ForegroundColor Gray
        }
    }
    
    Write-Host "Found $($icTransactions.Count) I/C transactions to duplicate" -ForegroundColor Green
    
    # Duplicate I/C transactions (add at the bottom)
    Write-Host "`nDuplicating I/C transactions..." -ForegroundColor Yellow
    foreach ($icRow in $icTransactions) {
        $dupRow = [ordered]@{}
        foreach ($prop in $icRow.PSObject.Properties) {
            $dupRow[$prop.Name] = $prop.Value
        }
        
        # Swap subsidiary: Statscore -> Lsports Data Ltd
        if ($dupRow[$subsidiaryColName].Trim().ToLower() -eq "statscore") {
            $dupRow[$subsidiaryColName] = "Lsports Data Ltd"
        }
        
        # Set "To Subsidiary / From Subsidiary" to Statscore
        $dupRow["To Subsidiary / From Subsidiary"] = "Statscore"
        
        [void]$processedRows.Add([PSCustomObject]$dupRow)
    }
    
    Write-Host "Final row count: $($processedRows.Count)" -ForegroundColor Green
    
    # Save to CSV
    Write-Host "`nSaving to CSV..." -ForegroundColor Yellow
    $processedRows | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
    Write-Host "Saved to: $outputPath" -ForegroundColor Green
    
    Write-Host "`n" + ("=" * 50) -ForegroundColor Cyan
    Write-Host "DONE!" -ForegroundColor Cyan
    Write-Host "Output file: $outputPath" -ForegroundColor Green
    Write-Host "`nNote: Duplicate I/C rows are at the bottom of the file." -ForegroundColor Yellow
    Write-Host "      To apply bold formatting, open in Excel and format manually." -ForegroundColor Yellow
    
} catch {
    Write-Host "`nERROR: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
}

