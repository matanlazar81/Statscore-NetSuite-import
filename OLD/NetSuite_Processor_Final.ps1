# NetSuite Import Processor - Final Version with GUI
# Processes SAP export CSV files for NetSuite import
# Adds Eliminate and To Subsidiary / From Subsidiary columns
# Duplicates I/C transactions with swapped subsidiaries

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.IO.Compression.FileSystem

# GUI File Selection
function Select-File($title, $filter) {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = $title
    $dialog.Filter = $filter
    $dialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    }
    return $null
}

function Show-Message($msg, $title = "NetSuite Processor") {
    [System.Windows.Forms.MessageBox]::Show($msg, $title, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

function Show-Error($msg) {
    [System.Windows.Forms.MessageBox]::Show($msg, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}

Write-Host "NetSuite Import Processor" -ForegroundColor Cyan
Write-Host ("=" * 50)

# Select files
Show-Message "Select the CSV file exported from SAP"
$csvFile = Select-File "Select CSV File" "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
if (-not $csvFile) { Show-Error "No CSV file selected."; exit }

Show-Message "Select the Excel mapping file (ST mapping.xlsx)"
$mappingFile = Select-File "Select Mapping Excel File" "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
if (-not $mappingFile) { Show-Error "No mapping file selected."; exit }

$outputPath = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($csvFile), "NetSuite_to_be_imported.csv")

Write-Host "CSV: $csvFile" -ForegroundColor Green
Write-Host "Mapping: $mappingFile" -ForegroundColor Green
Write-Host "Output: $outputPath" -ForegroundColor Green

try {
    # Step 1: Load mapping from Excel
    Write-Host "`nStep 1: Loading mapping file..." -ForegroundColor Yellow
    
    # Copy file to temp (in case it's open in Excel)
    $tempFile = [System.IO.Path]::GetTempFileName() + ".xlsx"
    $sourceStream = [System.IO.File]::Open($mappingFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    $destStream = [System.IO.File]::Create($tempFile)
    $sourceStream.CopyTo($destStream)
    $destStream.Close()
    $sourceStream.Close()
    
    # Open as ZIP and read XML
    $zip = [System.IO.Compression.ZipFile]::OpenRead($tempFile)
    
    # Read shared strings
    $sharedStringsEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/sharedStrings.xml" }
    $sharedStrings = @()
    if ($sharedStringsEntry) {
        $reader = [System.IO.StreamReader]::new($sharedStringsEntry.Open())
        $sharedStringsXml = [xml]$reader.ReadToEnd()
        $reader.Close()
        $sharedStrings = @($sharedStringsXml.sst.si | ForEach-Object { 
            if ($_.t -is [string]) { $_.t } elseif ($_.t) { $_.t.'#text' } else { "" }
        })
    }
    
    # Read sheet
    $sheetEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/worksheets/sheet1.xml" }
    $reader = [System.IO.StreamReader]::new($sheetEntry.Open())
    $sheetXml = [xml]$reader.ReadToEnd()
    $reader.Close()
    $zip.Dispose()
    Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
    
    $rows = @($sheetXml.worksheet.sheetData.row)
    Write-Host "  Found $($rows.Count) rows in mapping" -ForegroundColor Gray
    
    # Helper functions
    function Get-CellValue($cell, $strings) {
        if (-not $cell) { return $null }
        $value = $cell.v
        if ($cell.t -eq "s" -and $strings.Count -gt 0 -and $value) {
            $index = [int]$value
            if ($index -lt $strings.Count) { return $strings[$index] }
        }
        return $value
    }
    
    function Get-ColumnLetter($cellRef) { return ($cellRef -replace '\d+', '') }
    
    # Find columns
    $headerCells = @($rows[0].c)
    $internalCol = $null
    $intercompanyCol = $null
    
    foreach ($cell in $headerCells) {
        $colLetter = Get-ColumnLetter $cell.r
        $value = Get-CellValue $cell $sharedStrings
        if ($value) {
            $valueLower = $value.ToString().Trim().ToLower()
            if ($valueLower -like "*internal") { $internalCol = $colLetter }
            elseif ($valueLower -eq "intercompany") { $intercompanyCol = $colLetter }
        }
    }
    
    if (-not $internalCol -or -not $intercompanyCol) {
        throw "Could not find required columns in mapping file"
    }
    
    Write-Host "  Internal column: $internalCol, Intercompany column: $intercompanyCol" -ForegroundColor Gray
    
    # Build mapping dictionary
    $mappingDict = @{}
    for ($i = 1; $i -lt $rows.Count; $i++) {
        $row = $rows[$i]
        $internalValue = $null
        $intercompanyValue = $null
        
        foreach ($cell in @($row.c)) {
            $colLetter = Get-ColumnLetter $cell.r
            $value = Get-CellValue $cell $sharedStrings
            if ($colLetter -eq $internalCol) { $internalValue = $value }
            elseif ($colLetter -eq $intercompanyCol) { $intercompanyValue = $value }
        }
        
        if ($null -ne $internalValue -and $internalValue.ToString().Trim() -ne "") {
            $isIC = $intercompanyValue -and $intercompanyValue.ToString().Trim().ToLower() -eq "yes"
            $mappingDict[$internalValue.ToString()] = $isIC
        }
    }
    
    $icCount = ($mappingDict.Values | Where-Object { $_ -eq $true }).Count
    Write-Host "  Loaded $($mappingDict.Count) mappings ($icCount I/C accounts)" -ForegroundColor Green
    
    # Step 2: Load and process CSV
    Write-Host "`nStep 2: Loading CSV file..." -ForegroundColor Yellow
    $csvData = @(Import-Csv -Path $csvFile)
    Write-Host "  Loaded $($csvData.Count) rows" -ForegroundColor Green
    
    # Find column names
    $headers = $csvData[0].PSObject.Properties.Name
    $internalColName = $headers | Where-Object { $_.Trim().ToLower() -eq "internal" } | Select-Object -First 1
    $subsidiaryColName = $headers | Where-Object { $_.Trim().ToLower() -eq "subsidiary" } | Select-Object -First 1
    
    if (-not $internalColName -or -not $subsidiaryColName) {
        throw "Could not find required columns in CSV file"
    }
    
    # Step 3: Process rows
    Write-Host "`nStep 3: Processing rows..." -ForegroundColor Yellow
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
        
        if ($isIC) { [void]$icTransactions.Add($rowObj) }
        
        if ($rowNum % 10000 -eq 0) {
            Write-Host "  Processed $rowNum rows..." -ForegroundColor Gray
        }
    }
    
    Write-Host "  Found $($icTransactions.Count) I/C transactions" -ForegroundColor Green
    
    # Step 4: Duplicate I/C transactions
    Write-Host "`nStep 4: Duplicating I/C transactions..." -ForegroundColor Yellow
    foreach ($icRow in $icTransactions) {
        $dupRow = [ordered]@{}
        foreach ($prop in $icRow.PSObject.Properties) {
            $dupRow[$prop.Name] = $prop.Value
        }
        
        # Swap subsidiary: Statscore -> Lsports Data Ltd
        if ($dupRow[$subsidiaryColName].Trim().ToLower() -eq "statscore") {
            $dupRow[$subsidiaryColName] = "Lsports Data Ltd"
        }
        
        $dupRow["To Subsidiary / From Subsidiary"] = "Statscore"
        [void]$processedRows.Add([PSCustomObject]$dupRow)
    }
    
    Write-Host "  Final row count: $($processedRows.Count)" -ForegroundColor Green
    
    # Step 5: Save
    Write-Host "`nStep 5: Saving to CSV..." -ForegroundColor Yellow
    $processedRows | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Saved to: $outputPath" -ForegroundColor Green
    
    Write-Host "`n" -NoNewline
    Write-Host ("=" * 50) -ForegroundColor Cyan
    Write-Host "DONE!" -ForegroundColor Cyan
    
    Show-Message "Processing complete!`n`nOutput saved to:`n$outputPath`n`nTotal rows: $($processedRows.Count)`nI/C transactions duplicated: $($icTransactions.Count)"
    
} catch {
    Write-Host "`nERROR: $_" -ForegroundColor Red
    Show-Error "An error occurred:`n$_"
}

