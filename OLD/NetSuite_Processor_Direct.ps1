# NetSuite Import Processor - Direct Version (no GUI)
# Uses the known file paths directly

Add-Type -AssemblyName Microsoft.Office.Interop.Excel

$csvFile = "C:\Users\matan\OneDrive\desktop\csv.csv"
$mappingFile = "C:\Users\matan\OneDrive\desktop\ST mapping.xlsx"

Write-Host "NetSuite Import Processor" -ForegroundColor Cyan
Write-Host ("=" * 50)
Write-Host "CSV file: $csvFile" -ForegroundColor Green
Write-Host "Mapping file: $mappingFile" -ForegroundColor Green

try {
    # Load mapping from Excel using COM
    Write-Host "`nLoading mapping file..." -ForegroundColor Yellow
    
    # Check if file is locked
    if (Test-Path ($mappingFile -replace '\.xlsx$', '.xlsx').Replace('\', '\~$')) {
        Write-Host "Warning: Excel file might be open. Trying to read anyway..." -ForegroundColor Yellow
    }
    
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    Write-Host "Opening Excel workbook..." -ForegroundColor Gray
    $workbook = $excel.Workbooks.Open($mappingFile, $false, $true)  # Read-only mode
    $worksheet = $workbook.Sheets.Item(1)
    
    # Find the used range
    $usedRange = $worksheet.UsedRange
    $rowCount = $usedRange.Rows.Count
    $colCount = $usedRange.Columns.Count
    
    Write-Host "Mapping file has $rowCount rows and $colCount columns" -ForegroundColor Gray
    
    # Show header row
    Write-Host "Headers in mapping file:" -ForegroundColor Gray
    for ($col = 1; $col -le $colCount; $col++) {
        $headerValue = $worksheet.Cells.Item(1, $col).Text
        Write-Host "  Column $col : '$headerValue'" -ForegroundColor Gray
    }
    
    # Find column indices (Internal and Intercompany)
    $internalCol = -1
    $intercompanyCol = -1
    
    for ($col = 1; $col -le $colCount; $col++) {
        $headerValue = $worksheet.Cells.Item(1, $col).Text.Trim().ToLower()
        if ($headerValue -eq "internal") {
            $internalCol = $col
        }
        elseif ($headerValue -eq "intercompany") {
            $intercompanyCol = $col
        }
    }
    
    if ($internalCol -eq -1) {
        throw "Could not find 'Internal' column in mapping file"
    }
    if ($intercompanyCol -eq -1) {
        throw "Could not find 'Intercompany' column in mapping file"
    }
    
    Write-Host "Found Internal column at position $internalCol" -ForegroundColor Gray
    Write-Host "Found Intercompany column at position $intercompanyCol" -ForegroundColor Gray
    
    # Build mapping hashtable
    $mappingDict = @{}
    for ($row = 2; $row -le $rowCount; $row++) {
        $internalValue = $worksheet.Cells.Item($row, $internalCol).Value2
        $intercompanyValue = $worksheet.Cells.Item($row, $intercompanyCol).Text.Trim().ToLower()
        
        if ($null -ne $internalValue) {
            $mappingDict[$internalValue.ToString()] = ($intercompanyValue -eq "yes")
        }
    }
    
    Write-Host "Loaded $($mappingDict.Count) mappings from Excel file" -ForegroundColor Green
    $icCount = ($mappingDict.Values | Where-Object { $_ -eq $true }).Count
    Write-Host "I/C accounts: $icCount" -ForegroundColor Gray
    Write-Host "Non-I/C accounts: $($mappingDict.Count - $icCount)" -ForegroundColor Gray
    
    $workbook.Close($false)
    
    # Load CSV
    Write-Host "`nLoading CSV file..." -ForegroundColor Yellow
    $csvData = Import-Csv -Path $csvFile
    
    Write-Host "Loaded $($csvData.Count) rows from CSV" -ForegroundColor Green
    
    # Find column names (handle potential whitespace)
    $headers = $csvData[0].PSObject.Properties.Name
    Write-Host "CSV Headers: $($headers -join ', ')" -ForegroundColor Gray
    
    $internalColName = $headers | Where-Object { $_.Trim().ToLower() -eq "internal" } | Select-Object -First 1
    $subsidiaryColName = $headers | Where-Object { $_.Trim().ToLower() -eq "subsidiary" } | Select-Object -First 1
    
    if (-not $internalColName) {
        throw "Could not find 'Internal' column in CSV file"
    }
    if (-not $subsidiaryColName) {
        throw "Could not find 'Subsidiary' column in CSV file"
    }
    
    Write-Host "Internal column: '$internalColName'" -ForegroundColor Gray
    Write-Host "Subsidiary column: '$subsidiaryColName'" -ForegroundColor Gray
    
    # Process rows and add new columns
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
        
        # Create a new hashtable for the row with new columns
        $newRow = [ordered]@{}
        foreach ($prop in $row.PSObject.Properties) {
            $newRow[$prop.Name] = $prop.Value
        }
        $newRow["Eliminate"] = if ($isIC) { "Yes" } else { "No" }
        $newRow["To Subsidiary / From Subsidiary"] = "Lsports Data Ltd"
        $newRow["_IsDuplicate"] = $false
        
        [void]$processedRows.Add([PSCustomObject]$newRow)
        
        if ($isIC) {
            [void]$icTransactions.Add([PSCustomObject]$newRow)
        }
        
        if ($rowNum % 5000 -eq 0) {
            Write-Host "  Processed $rowNum rows..." -ForegroundColor Gray
        }
    }
    
    Write-Host "Found $($icTransactions.Count) I/C transactions to duplicate" -ForegroundColor Green
    
    # Duplicate I/C transactions
    foreach ($icRow in $icTransactions) {
        # Create a copy of the row
        $dupRow = [ordered]@{}
        foreach ($prop in $icRow.PSObject.Properties) {
            $dupRow[$prop.Name] = $prop.Value
        }
        
        # Swap subsidiary: Statscore -> Lsports Data Ltd
        $currentSubsidiary = $dupRow[$subsidiaryColName]
        if ($currentSubsidiary.Trim().ToLower() -eq "statscore") {
            $dupRow[$subsidiaryColName] = "Lsports Data Ltd"
        }
        
        # Set To Subsidiary / From Subsidiary to Statscore
        $dupRow["To Subsidiary / From Subsidiary"] = "Statscore"
        $dupRow["_IsDuplicate"] = $true
        
        [void]$processedRows.Add([PSCustomObject]$dupRow)
    }
    
    Write-Host "Final row count: $($processedRows.Count)" -ForegroundColor Green
    
    # Save to Excel with formatting
    Write-Host "`nSaving result to Excel..." -ForegroundColor Yellow
    
    $outputPath = "C:\Users\matan\OneDrive\desktop\NetSuite_to_be_imported.xlsx"
    
    # Delete existing file if it exists
    if (Test-Path $outputPath) {
        Remove-Item $outputPath -Force
        Write-Host "Removed existing output file" -ForegroundColor Gray
    }
    
    # Create new workbook
    $outputWorkbook = $excel.Workbooks.Add()
    $outputSheet = $outputWorkbook.Sheets.Item(1)
    $outputSheet.Name = "NetSuite Import"
    
    # Get column names (excluding _IsDuplicate)
    $outputColumns = $processedRows[0].PSObject.Properties.Name | Where-Object { $_ -ne "_IsDuplicate" }
    
    Write-Host "Writing $($outputColumns.Count) columns..." -ForegroundColor Gray
    
    # Write header row (bold)
    $col = 1
    foreach ($colName in $outputColumns) {
        $cell = $outputSheet.Cells.Item(1, $col)
        $cell.Value2 = $colName
        $cell.Font.Bold = $true
        $col++
    }
    
    # Write data rows
    Write-Host "Writing data rows..." -ForegroundColor Gray
    $rowNum = 2
    foreach ($dataRow in $processedRows) {
        $col = 1
        $isDuplicate = $dataRow._IsDuplicate
        
        foreach ($colName in $outputColumns) {
            $cell = $outputSheet.Cells.Item($rowNum, $col)
            $cell.Value2 = $dataRow.$colName
            
            # Bold the duplicated I/C rows
            if ($isDuplicate) {
                $cell.Font.Bold = $true
            }
            $col++
        }
        $rowNum++
        
        # Show progress every 5000 rows
        if (($rowNum - 1) % 5000 -eq 0) {
            Write-Host "  Written $($rowNum - 1) rows..." -ForegroundColor Gray
        }
    }
    
    # Auto-fit columns
    Write-Host "Auto-fitting columns..." -ForegroundColor Gray
    $outputSheet.UsedRange.Columns.AutoFit() | Out-Null
    
    # Save and close
    $outputWorkbook.SaveAs($outputPath)
    $outputWorkbook.Close($true)
    
    Write-Host "`nOutput saved to: $outputPath" -ForegroundColor Green
    
    # Clean up Excel COM
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outputSheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outputWorkbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Host "`nDone!" -ForegroundColor Cyan
    
} catch {
    Write-Host "`nError: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    
    # Clean up Excel if it's still open
    if ($excel) {
        try {
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        } catch {}
    }
}

