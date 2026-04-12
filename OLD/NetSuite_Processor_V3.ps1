# NetSuite Import Processor - Version 3
# Uses OLE DB to read Excel without COM automation

$csvFile = "C:\Users\matan\OneDrive\desktop\csv.csv"
$mappingFile = "C:\Users\matan\OneDrive\desktop\ST mapping.xlsx"
$outputPath = "C:\Users\matan\OneDrive\desktop\NetSuite_to_be_imported.csv"

Write-Host "NetSuite Import Processor v3" -ForegroundColor Cyan
Write-Host ("=" * 50)
Write-Host "CSV file: $csvFile" -ForegroundColor Green
Write-Host "Mapping file: $mappingFile" -ForegroundColor Green

try {
    # Read Excel file using OLE DB
    Write-Host "`nLoading mapping file using OLE DB..." -ForegroundColor Yellow
    
    $connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$mappingFile;Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1'"
    
    $connection = New-Object System.Data.OleDb.OleDbConnection($connectionString)
    $connection.Open()
    
    # Get the first sheet name
    $schemaTable = $connection.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Tables, $null)
    $sheetName = $schemaTable.Rows[0]["TABLE_NAME"]
    Write-Host "Reading sheet: $sheetName" -ForegroundColor Gray
    
    # Query the data
    $query = "SELECT * FROM [$sheetName]"
    $command = New-Object System.Data.OleDb.OleDbCommand($query, $connection)
    $adapter = New-Object System.Data.OleDb.OleDbDataAdapter($command)
    $dataTable = New-Object System.Data.DataTable
    $adapter.Fill($dataTable) | Out-Null
    
    $connection.Close()
    
    Write-Host "Loaded $($dataTable.Rows.Count) rows from mapping file" -ForegroundColor Green
    Write-Host "Columns: $($dataTable.Columns.ColumnName -join ', ')" -ForegroundColor Gray
    
    # Build mapping dictionary
    $mappingDict = @{}
    
    # Find column names
    $internalCol = $dataTable.Columns | Where-Object { $_.ColumnName.Trim().ToLower() -eq "internal" } | Select-Object -First 1 -ExpandProperty ColumnName
    $intercompanyCol = $dataTable.Columns | Where-Object { $_.ColumnName.Trim().ToLower() -eq "intercompany" } | Select-Object -First 1 -ExpandProperty ColumnName
    
    if (-not $internalCol) { throw "Could not find 'Internal' column in mapping file" }
    if (-not $intercompanyCol) { throw "Could not find 'Intercompany' column in mapping file" }
    
    Write-Host "Internal column: '$internalCol'" -ForegroundColor Gray
    Write-Host "Intercompany column: '$intercompanyCol'" -ForegroundColor Gray
    
    foreach ($row in $dataTable.Rows) {
        $internalValue = $row[$internalCol]
        $intercompanyValue = $row[$intercompanyCol]
        
        if ($null -ne $internalValue -and $internalValue -ne [DBNull]::Value) {
            $isIC = $intercompanyValue.ToString().Trim().ToLower() -eq "yes"
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
    
    Write-Host "`nDone!" -ForegroundColor Cyan
    Write-Host "Note: Duplicate I/C rows are at the bottom of the file." -ForegroundColor Yellow
    Write-Host "      Open in Excel to apply bold formatting if needed." -ForegroundColor Yellow
    
} catch {
    Write-Host "`nERROR: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    
    # If OLE DB failed, try alternative approach
    if ($_.ToString() -match "OleDb|ACE") {
        Write-Host "`nOLE DB not available. Trying alternative approach..." -ForegroundColor Yellow
        Write-Host "Please save the mapping file as CSV and place it at:" -ForegroundColor Yellow
        Write-Host "C:\Users\matan\OneDrive\desktop\ST mapping.csv" -ForegroundColor Yellow
    }
}

