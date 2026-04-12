# NetSuite Import Processor
# - Opens a GUI to select CSV file and Excel mapping file
# - Adds "Eliminate" column based on Intercompany status
# - Adds "To Subsidiary / From Subsidiary" column
# - Duplicates I/C transactions with swapped subsidiaries

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.Office.Interop.Excel

# Function to select file
function Select-File {
    param (
        [string]$Title,
        [string]$Filter
    )
    
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = $Title
    $dialog.Filter = $Filter
    $dialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    }
    return $null
}

# Show info message
function Show-Message {
    param (
        [string]$Message,
        [string]$Title = "NetSuite Import Processor"
    )
    [System.Windows.Forms.MessageBox]::Show($Message, $Title, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

# Show error message
function Show-Error {
    param (
        [string]$Message,
        [string]$Title = "Error"
    )
    [System.Windows.Forms.MessageBox]::Show($Message, $Title, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}

Write-Host "NetSuite Import Processor" -ForegroundColor Cyan
Write-Host "=" * 50

# Select CSV file
Show-Message "First, select the CSV file exported from SAP"
$csvFile = Select-File -Title "Select CSV File" -Filter "CSV files (*.csv)|*.csv|All files (*.*)|*.*"

if (-not $csvFile) {
    Show-Error "No CSV file selected. Exiting."
    exit
}

Write-Host "CSV file: $csvFile" -ForegroundColor Green

# Select mapping Excel file
Show-Message "Now, select the Excel mapping file"
$mappingFile = Select-File -Title "Select Mapping Excel File" -Filter "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*"

if (-not $mappingFile) {
    Show-Error "No mapping file selected. Exiting."
    exit
}

Write-Host "Mapping file: $mappingFile" -ForegroundColor Green

try {
    # Load mapping from Excel using COM
    Write-Host "`nLoading mapping file..." -ForegroundColor Yellow
    
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    $workbook = $excel.Workbooks.Open($mappingFile)
    $worksheet = $workbook.Sheets.Item(1)
    
    # Find the used range
    $usedRange = $worksheet.UsedRange
    $rowCount = $usedRange.Rows.Count
    $colCount = $usedRange.Columns.Count
    
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
        
        if ($internalValue -ne $null) {
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
    
    $processedRows = @()
    $icTransactions = @()
    
    foreach ($row in $csvData) {
        $internalValue = $row.$internalColName.ToString().Trim()
        
        # Determine if I/C
        $isIC = $false
        if ($mappingDict.ContainsKey($internalValue)) {
            $isIC = $mappingDict[$internalValue]
        }
        
        # Add new columns
        $row | Add-Member -NotePropertyName "Eliminate" -NotePropertyValue $(if ($isIC) { "Yes" } else { "No" }) -Force
        $row | Add-Member -NotePropertyName "To Subsidiary / From Subsidiary" -NotePropertyValue "Lsports Data Ltd" -Force
        $row | Add-Member -NotePropertyName "_IsDuplicate" -NotePropertyValue $false -Force
        
        $processedRows += $row
        
        if ($isIC) {
            $icTransactions += $row
        }
    }
    
    Write-Host "Found $($icTransactions.Count) I/C transactions to duplicate" -ForegroundColor Green
    
    # Duplicate I/C transactions
    foreach ($icRow in $icTransactions) {
        # Create a copy of the row
        $dupRow = $icRow.PSObject.Copy()
        
        # Swap subsidiary: Statscore -> Lsports Data Ltd
        $currentSubsidiary = $dupRow.$subsidiaryColName
        if ($currentSubsidiary.Trim().ToLower() -eq "statscore") {
            $dupRow.$subsidiaryColName = "Lsports Data Ltd"
        }
        
        # Set To Subsidiary / From Subsidiary to Statscore
        $dupRow."To Subsidiary / From Subsidiary" = "Statscore"
        $dupRow._IsDuplicate = $true
        
        $processedRows += $dupRow
    }
    
    Write-Host "Final row count: $($processedRows.Count)" -ForegroundColor Green
    
    # Save to Excel with formatting
    Write-Host "`nSaving result to Excel..." -ForegroundColor Yellow
    
    $outputPath = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($csvFile), "NetSuite_to_be_imported.xlsx")
    
    # Create new workbook
    $outputWorkbook = $excel.Workbooks.Add()
    $outputSheet = $outputWorkbook.Sheets.Item(1)
    $outputSheet.Name = "NetSuite Import"
    
    # Get column names (excluding _IsDuplicate)
    $outputColumns = $processedRows[0].PSObject.Properties.Name | Where-Object { $_ -ne "_IsDuplicate" }
    
    # Write header row (bold)
    $col = 1
    foreach ($colName in $outputColumns) {
        $cell = $outputSheet.Cells.Item(1, $col)
        $cell.Value2 = $colName
        $cell.Font.Bold = $true
        $col++
    }
    
    # Write data rows
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
        
        # Show progress every 1000 rows
        if ($rowNum % 1000 -eq 0) {
            Write-Host "  Processed $rowNum rows..." -ForegroundColor Gray
        }
    }
    
    # Auto-fit columns
    $outputSheet.UsedRange.Columns.AutoFit() | Out-Null
    
    # Save and close
    if (Test-Path $outputPath) {
        Remove-Item $outputPath -Force
    }
    $outputWorkbook.SaveAs($outputPath)
    $outputWorkbook.Close($true)
    
    Write-Host "`nOutput saved to: $outputPath" -ForegroundColor Green
    
    # Clean up Excel COM
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Show-Message "Processing complete!`n`nOutput saved to:`n$outputPath"
    
    Write-Host "`nDone!" -ForegroundColor Cyan
    
} catch {
    Write-Host "`nError: $_" -ForegroundColor Red
    Show-Error "An error occurred:`n$_"
    
    # Clean up Excel if it's still open
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

