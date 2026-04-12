# NetSuite Import Processor - Version 2
# Uses ImportExcel module or COM as fallback

$csvFile = "C:\Users\matan\OneDrive\desktop\csv.csv"
$mappingFile = "C:\Users\matan\OneDrive\desktop\ST mapping.xlsx"
$outputPath = "C:\Users\matan\OneDrive\desktop\NetSuite_to_be_imported.xlsx"

Write-Host "NetSuite Import Processor v2" -ForegroundColor Cyan
Write-Host ("=" * 50)
Write-Host "CSV file: $csvFile" -ForegroundColor Green
Write-Host "Mapping file: $mappingFile" -ForegroundColor Green

# First, try to install ImportExcel module if not available
$hasImportExcel = $false
try {
    Import-Module ImportExcel -ErrorAction Stop
    $hasImportExcel = $true
    Write-Host "Using ImportExcel module" -ForegroundColor Gray
} catch {
    Write-Host "ImportExcel module not available, will use COM" -ForegroundColor Yellow
}

try {
    # Load mapping from Excel
    Write-Host "`nLoading mapping file..." -ForegroundColor Yellow
    
    $mappingDict = @{}
    
    if ($hasImportExcel) {
        # Use ImportExcel module
        $mappingData = Import-Excel -Path $mappingFile
        
        foreach ($row in $mappingData) {
            $internalValue = $row.Internal
            $intercompanyValue = $row.Intercompany
            
            if ($null -ne $internalValue) {
                $mappingDict[$internalValue.ToString()] = ($intercompanyValue.ToString().Trim().ToLower() -eq "yes")
            }
        }
    } else {
        # Use COM
        $excel = $null
        $workbook = $null
        try {
            $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            
            Write-Host "Opening Excel workbook (read-only)..." -ForegroundColor Gray
            $workbook = $excel.Workbooks.Open($mappingFile, 0, $true)  # Read-only
            $worksheet = $workbook.Sheets.Item(1)
            
            $usedRange = $worksheet.UsedRange
            $rowCount = $usedRange.Rows.Count
            $colCount = $usedRange.Columns.Count
            
            Write-Host "Mapping file has $rowCount rows and $colCount columns" -ForegroundColor Gray
            
            # Find column indices
            $internalCol = -1
            $intercompanyCol = -1
            
            for ($col = 1; $col -le $colCount; $col++) {
                $headerValue = $worksheet.Cells.Item(1, $col).Text.Trim().ToLower()
                Write-Host "  Column $col : '$headerValue'" -ForegroundColor Gray
                if ($headerValue -eq "internal") { $internalCol = $col }
                elseif ($headerValue -eq "intercompany") { $intercompanyCol = $col }
            }
            
            if ($internalCol -eq -1) { throw "Could not find 'Internal' column in mapping file" }
            if ($intercompanyCol -eq -1) { throw "Could not find 'Intercompany' column in mapping file" }
            
            # Build mapping
            for ($row = 2; $row -le $rowCount; $row++) {
                $internalValue = $worksheet.Cells.Item($row, $internalCol).Value2
                $intercompanyValue = $worksheet.Cells.Item($row, $intercompanyCol).Text.Trim().ToLower()
                
                if ($null -ne $internalValue) {
                    $mappingDict[$internalValue.ToString()] = ($intercompanyValue -eq "yes")
                }
            }
            
            $workbook.Close($false)
            $excel.Quit()
        } catch {
            Write-Host "COM Error: $_" -ForegroundColor Red
            throw
        } finally {
            if ($workbook) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null }
            if ($excel) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
    
    Write-Host "Loaded $($mappingDict.Count) mappings" -ForegroundColor Green
    $icCount = ($mappingDict.Values | Where-Object { $_ -eq $true }).Count
    Write-Host "  I/C accounts: $icCount" -ForegroundColor Gray
    Write-Host "  Non-I/C accounts: $($mappingDict.Count - $icCount)" -ForegroundColor Gray
    
    # Sample some I/C accounts
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
    
    Write-Host "Internal column: '$internalColName'" -ForegroundColor Gray
    Write-Host "Subsidiary column: '$subsidiaryColName'" -ForegroundColor Gray
    
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
        $newRow["_IsDuplicate"] = $false
        
        [void]$processedRows.Add([PSCustomObject]$newRow)
        
        if ($isIC) {
            [void]$icTransactions.Add([PSCustomObject]$newRow)
        }
        
        if ($rowNum % 10000 -eq 0) {
            Write-Host "  Processed $rowNum rows..." -ForegroundColor Gray
        }
    }
    
    Write-Host "Found $($icTransactions.Count) I/C transactions to duplicate" -ForegroundColor Green
    
    # Duplicate I/C transactions
    foreach ($icRow in $icTransactions) {
        $dupRow = [ordered]@{}
        foreach ($prop in $icRow.PSObject.Properties) {
            $dupRow[$prop.Name] = $prop.Value
        }
        
        # Swap subsidiary
        if ($dupRow[$subsidiaryColName].Trim().ToLower() -eq "statscore") {
            $dupRow[$subsidiaryColName] = "Lsports Data Ltd"
        }
        
        $dupRow["To Subsidiary / From Subsidiary"] = "Statscore"
        $dupRow["_IsDuplicate"] = $true
        
        [void]$processedRows.Add([PSCustomObject]$dupRow)
    }
    
    Write-Host "Final row count: $($processedRows.Count)" -ForegroundColor Green
    
    # Save to CSV first (as backup), then try Excel
    Write-Host "`nSaving to CSV (backup)..." -ForegroundColor Yellow
    $csvOutputPath = $outputPath -replace '\.xlsx$', '.csv'
    $outputColumns = $processedRows[0].PSObject.Properties.Name | Where-Object { $_ -ne "_IsDuplicate" }
    
    # Create output with selected columns
    $outputData = $processedRows | Select-Object -Property $outputColumns
    $outputData | Export-Csv -Path $csvOutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "Saved to: $csvOutputPath" -ForegroundColor Green
    
    # Try to save as Excel with formatting
    Write-Host "`nSaving to Excel with formatting..." -ForegroundColor Yellow
    
    if ($hasImportExcel) {
        # Use ImportExcel
        if (Test-Path $outputPath) { Remove-Item $outputPath -Force }
        
        $outputData | Export-Excel -Path $outputPath -AutoSize -BoldTopRow -WorksheetName "NetSuite Import"
        
        # Apply bold to duplicate rows (would require more complex logic with ImportExcel)
        Write-Host "Saved to: $outputPath" -ForegroundColor Green
    } else {
        # Use COM
        $excel = $null
        $workbook = $null
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            
            $workbook = $excel.Workbooks.Add()
            $sheet = $workbook.Sheets.Item(1)
            $sheet.Name = "NetSuite Import"
            
            # Write headers
            $col = 1
            foreach ($colName in $outputColumns) {
                $sheet.Cells.Item(1, $col).Value2 = $colName
                $sheet.Cells.Item(1, $col).Font.Bold = $true
                $col++
            }
            
            # Write data
            $rowNum = 2
            foreach ($dataRow in $processedRows) {
                $col = 1
                $isDuplicate = $dataRow._IsDuplicate
                
                foreach ($colName in $outputColumns) {
                    $cell = $sheet.Cells.Item($rowNum, $col)
                    $cell.Value2 = $dataRow.$colName
                    if ($isDuplicate) { $cell.Font.Bold = $true }
                    $col++
                }
                $rowNum++
                
                if (($rowNum - 1) % 10000 -eq 0) {
                    Write-Host "  Written $($rowNum - 1) rows..." -ForegroundColor Gray
                }
            }
            
            $sheet.UsedRange.Columns.AutoFit() | Out-Null
            
            if (Test-Path $outputPath) { Remove-Item $outputPath -Force }
            $workbook.SaveAs($outputPath)
            $workbook.Close($true)
            $excel.Quit()
            
            Write-Host "Saved to: $outputPath" -ForegroundColor Green
        } catch {
            Write-Host "Excel save error: $_" -ForegroundColor Red
            Write-Host "CSV output is available at: $csvOutputPath" -ForegroundColor Yellow
        } finally {
            if ($workbook) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null }
            if ($excel) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }
            [System.GC]::Collect()
        }
    }
    
    Write-Host "`nDone!" -ForegroundColor Cyan
    
} catch {
    Write-Host "`nFATAL ERROR: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
}

