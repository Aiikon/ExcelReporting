Function Import-Excel
{
    Param
    (
        [Parameter(Mandatory=$true, Position=0)] [string] $FilePath,
        [Parameter()] [string] $SheetName,
        [Parameter()] [int] $HeaderRow,
        [Parameter()] [string[]] $Header
    )
    End
    {
        $importCsvArgs = @{}
        if ($Header)
        {
            $HeaderRow += 1
            $importCsvArgs.Header = $Header
        }

        $finalPath = $null

        try
        {
            $finalPath = Resolve-Path $FilePath -ErrorAction Stop | Select-Object -ExpandProperty ProviderPath
        }
        catch
        {
            throw "File $FilePath not found."
            return
        }

        $tempFile = [System.IO.Path]::GetTempFileName()

        $excel = New-Object -ComObject Excel.Application
        $excel.Application.Workbooks.Open($finalPath, $null, $true) | Out-Null
        $excel.DisplayAlerts = $false

        if ($SheetName)
        {
            $activeSheet = $excel.Application.ActiveWorkbook.Sheets |
                Where-Object { $_.Name -ieq $SheetName } |
                Select-Object -First 1

            if (!$activeSheet) { throw "Unable to find worksheet: $SheetName." }
        }
        else
        {
            $activeSheet = $excel.ActiveSheet
        }

        if ($HeaderRow -and $HeaderRow -ne 1)
        {
            $prev = $HeaderRow - 1
            [void]$activeSheet.Range("1:$prev").Delete()
        }

        $activeSheet.SaveAs($tempFile, 6)
        $excel.Application.ActiveWorkbook.Close()
        $excel.Quit()

        while ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) -gt 0) { }
        Remove-Variable excel, activeSheet
        [GC]::Collect()

        Import-Csv $tempFile -Encoding ASCII @importCsvArgs

        Remove-Item $tempFile
    }
}

Function Export-Excel
{
    Param
    (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)] [object] $InputObject,
        [Parameter(Mandatory=$true, Position=0)] [string] $FilePath,
        [Parameter()] [switch] $Open
    )
    Begin
    {
        if (Test-Path $FilePath)
        {
            try
            {
                Remove-Item -Force -Path $FilePath
            }
            catch
            {
                throw $_
                return
            }
        }

        $finalPath = $null
        try
        {
            '' | Out-File $FilePath -ErrorAction Stop
            $finalPath = Resolve-Path $FilePath | Select-Object -ExpandProperty Path
            Remove-Item $FilePath -ErrorAction Stop
        }
        catch
        {
            throw $_
            return
        }

        $objects = New-Object System.Collections.Generic.List``1[System.Object]
    }
    Process
    {
        $objects.Add($InputObject)
    }
    End
    {
        $tempFile = [System.IO.Path]::GetTempFileName()
        $objects | Export-Csv $tempFile -NoTypeInformation

        $excel = New-Object -ComObject Excel.Application
        #                                    (Filename,  Origin, StartRow, DataType, TextQualifier, Consecutive, Tab   , Semic , Comma)
        $excel.Application.Workbooks.OpenText($tempFile, $null , 1  , 1       , 1            , $false     , $false, $false, $true)
        $columns = $excel.Application.ActiveSheet.Range("A:Z").Columns
        $columns.AutoFit() | Out-Null
        $excel.Application.ActiveWorkbook.SaveAs($finalPath, 51)

        if ($Open)
        {
            $excel.Visible = $true
        }
        else
        {
            $excel.ActiveWorkbook.Close()
            $excel.Workbooks.Close()
            $excel.Quit()
        }

        while ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) -gt 0) { }
        Remove-Variable excel, columns
        [GC]::Collect()
        Remove-Item $tempFile
    }
}

Function Read-ExcelWindow
{
    Param
    (
        [Parameter()] [switch] $NoHidden,
        [Parameter(Position=0)] [object] $Workbook,
        [Parameter(Position=1)] [object] $Sheet,
        [Parameter()] [int] $HeaderRow = 1
    )
    End
    {
        $workbookIsString = $Workbook -and $Workbook.GetType().FullName -ieq 'System.String'
        $sheetIsString = $Sheet -and $Sheet.GetType().FullName -ieq 'System.String'
        if ($workbookIsString)
        {
            $excel = [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')
            $activeWorkbook = $excel.Workbooks |
                Where-Object { $_.Name -like "$workbook.*" } |
                Select-Object -First 1

            if (!$activeWorkbook) { throw "Unable to find workbook: $Workbook." }

            if ($Sheet -and -not $sheetIsString)
            { throw "Sheet must be a string if workbook is specified." }

            if ($Sheet)
            {
                $activeSheet = $activeWorkbook.Sheets |
                    Where-Object { $_.Name -ieq $Sheet } |
                    Select-Object -First 1

                if (!$activeSheet) { throw "Unable to find worksheet: $Sheet." }
            }
            else
            {
                $activeSheet = $activeWorkbook.ActiveSheet
            }
        }
        else
        {
            if ($Sheet -and -not $sheetIsString)
            {
                $activeSheet = $Sheet
            }
            elseif ($Sheet -and $sheetIsString)
            {
                $excel = [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')
                $activeSheet = $excel.ActiveWorkbook.Sheets |
                    Where-Object { $_.Name -ieq $Sheet } |
                    Select-Object -First 1

                if (!$activeSheet) { throw "Unable to find worksheet: $Sheet." }
            }
            else
            {
                $excel = [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')
                $activeSheet = $excel.ActiveSheet
            }
        }

        if (-not $activeSheet)
        {
            throw "Unable to access Excel sheet. Cell may be in edit mode."
        }

        $region = $activeSheet.Cells.Item($HeaderRow,1).CurrentRegion

        if (!$NoHidden)
        {
            $hiddenColumns = $region.Columns |
                Where-Object Hidden
            $hiddenColumns | ForEach-Object { $_.Hidden = $false }
            $hiddenRows = $region.Rows |
                Where-Object Hidden
            $hiddenRows | ForEach-Object { $_.Hidden = $false }
        }
        if ($HeaderRow -ne 1)
        {
            $firstCell = $region.Cells.Item($HeaderRow, 1)
            $lastCell = $region.Cells.Item($region.EntireRow.Count - $HeaderRow + 1, $region.EntireColumn.Count)
            $region = $activeSheet.Range($firstCell, $lastCell)
        }

        $oldClipboard = [System.Windows.Forms.Clipboard]::GetText()

        [void]$region.Copy()

        $clipboard = [System.Windows.Forms.Clipboard]::GetText()

        if ($oldClipboard)
        {
            [System.Windows.Forms.Clipboard]::SetText($oldClipboard)
        }
        else { [System.Windows.Forms.Clipboard]::Clear() }

        if (!$NoHidden)
        {
            $hiddenColumns | ForEach-Object { $_.Hidden = $true }
            $hiddenRows | ForEach-Object { $_.Hidden = $true }
        }

        if (-not ($sheetIsString -or $workbookIsString))
        {
            #Remove-Variable Workbook, Sheet, activeWorkbook, activeSheet, excel, hiddenColumns, hiddenRows, region  -ErrorAction Ignore
            #[GC]::Collect()
        }

        $clipboard | ConvertFrom-Csv -Delimiter "`t"
    }
}

Function Update-ExcelWindow
{
    Param
    (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)] [object] $InputObject,
        [Parameter(Mandatory=$true, Position=0)] [string] $InputKey,
        [Parameter(Position=1)] [string] $ExcelKey,
        [Parameter()] [string] $TimestampColumn,
        [Parameter()] [string[]] $TimestampTestColumns,
        [Parameter()] [switch] $NoHidden
    )
    Begin
    {
        if (!(Get-Command "ConvertTo-Hashtable")) { throw "The Data module is required." }

        if (!$ExcelKey) { $ExcelKey = $InputKey }

        $excel = [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')
        $sheet = $excel.ActiveSheet

        $excelData = Read-ExcelWindow -NoHidden:$NoHidden

        $excelIndices = @{}
        $excelHash = @{}

        $i = 2
        foreach ($data in $excelData)
        {
            $excelKeyValue = $data.$Excelkey

            if (-not $excelIndices.ContainsKey($excelKeyValue))
            {
                $excelIndices.Add($excelKeyValue, (New-Object System.Collections.Generic.List[object]))
                $excelHash.Add($excelKeyValue, (New-Object System.Collections.Generic.List[object]))
            }

            $excelIndices.$excelKeyValue.Add($i)
            $excelHash.$excelKeyValue.Add($data)

            $i += 1
        }

        $excelPropertyHash = @{}
        $i = 1
        foreach ($oldProperty in ($data.psobject.Properties | Select-Object -ExpandProperty Name))
        {
            $excelPropertyHash.Add($oldProperty, $i)
            $i += 1
        }

        $inputProperties = $null
    }
    Process
    {
        if (!$inputProperties)
        {
            $inputProperties = @{}
            $InputObject | Get-PropertyName | % { $inputProperties.Add($_, $null) }
        }
        $keyValue = "$($InputObject.$InputKey)"
        if ($excelHash.ContainsKey($keyValue))
        {
            $i = 0
            foreach ($excelRecord in $excelHash[$keyValue])
            {
                $updated = $false
                foreach ($excelProperty in $excelPropertyHash.Keys)
                {
                    $row = $excelIndices[$keyValue][$i]
                    if ($inputProperties.ContainsKey($excelProperty) -and "$($InputObject.$excelProperty)" -ne $excelRecord.$excelProperty)
                    {
                        $column = $excelPropertyHash[$excelProperty]
                        $sheet.Cells.Item($row, $column) = $InputObject.$excelProperty
                        Write-Verbose "Updating ($row) $keyValue column $excelProperty with $($InputObject.$excelProperty)."
                        if (-not $TimestampTestColumns -or $excelProperty -iin $TimestampTestColumns)
                        {
                            $updated = $true
                        }
                    }
                }
                if ($updated -and $TimestampColumn)
                {
                    Write-Verbose "Updating ($row) $keyValue column $TimestampColumn with $([DateTime]::Now)."
                    $sheet.Cells.Item($row, $excelPropertyHash[$TimestampColumn]) = [DateTime]::Now
                }
                $i += 1
            }
        }
    }
    End
    {
        #[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
        #while ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) -gt 0) { }
        #while ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) -gt 0) { }
        #Remove-Variable excel, sheet -ErrorAction Ignore
        #[GC]::Collect()
    }
}

Function Out-Excel
{
    Param
    (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)] [object] $InputObject,
        [Parameter()] [switch] $NoAutoFit
    )
    Begin
    {
        Add-Type -AssemblyName 'presentationframework'
        Set-StrictMode -Version Latest
        $objects = New-Object System.Collections.Generic.List``1[System.Object]
    }
    Process
    {
        $objects.Add($InputObject)
    }
    End
    {
        if ($objects.Count -eq 0)
        {
            Write-Warning "No objects were passed to Out-Excel. Exiting."
            return
        }

        $excel = New-Object -ComObject Excel.Application
        $workbook = $excel.Workbooks.Add()
        $range = $workbook.ActiveSheet.Range("A1")
        $clipText = ($objects | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation | Out-String)

        $oldClipboard = [System.Windows.Forms.Clipboard]::GetText()

        [System.Windows.Forms.Clipboard]::SetText($clipText)
        [void]$workbook.ActiveSheet.Paste($range, $false)

        if ($oldClipboard)
        {
            [System.Windows.Forms.Clipboard]::SetText($oldClipboard)
        }
        else { [System.Windows.Forms.Clipboard]::Clear() }

        if (-not $NoAutoFit.IsPresent)
        {
            [void]$workbook.ActiveSheet.Range("A1").CurrentRegion.Columns.AutoFit()
        }
        $excel.Visible = $true

        while ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) -gt 0) { }
        Remove-Variable excel, workbook, range -ErrorAction Ignore
        [GC]::Collect()
    }
}
