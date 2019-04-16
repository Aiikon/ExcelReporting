Function Import-Excel
{
    Param
    (
        [Parameter(Mandatory=$true, Position=0)] [string] $FilePath,
        [Parameter()] [string] $SheetName
    )
    End
    {
        trap { $PSCmdlet.ThrowTerminatingError($_) }

        $finalPath = $PSCmdlet.GetResolvedProviderPathFromPSPath($FilePath)
        $tempFile = [System.IO.Path]::GetTempFileName()

        $excel = New-Object -ComObject Excel.Application
        [void]$excel.Application.Workbooks.Open($finalPath, $null, $true)
        $excel.DisplayAlerts = $false

        if ($SheetName)
        {
            $activeSheet = $excel.Application.ActiveWorkbook.Sheets |
                Where-Object { $_.Name -eq $SheetName } |
                Select-Object -First 1

            if (!$activeSheet) { throw "Unable to find a worksheet titled '$SheetName'." }
        }
        else
        {
            $activeSheet = $excel.ActiveSheet
        }

        $activeSheet.SaveAs($tempFile, 6)
        $excel.Application.ActiveWorkbook.Close()
        $excel.Quit()

        while ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) -gt 0) { }
        Remove-Variable excel, activeSheet
        [GC]::Collect()

        Import-Csv $tempFile -Encoding ASCII
        [System.IO.File]::Delete($tempFile)
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
        trap { $PSCmdlet.ThrowTerminatingError($_) }

        $finalPath = $PSCmdlet.GetResolvedProviderPathFromPSPath($FilePath)

        $objects = New-Object System.Collections.Generic.List[object]
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
        $excel.Application.Workbooks.OpenText($tempFile, $null, 1, 1, 1, $false, $false, $false, $true)
        $columns = $excel.Application.ActiveSheet.Range("A:Z").Columns
        [void]$columns.AutoFit()
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
        [System.IO.File]::Delete($tempFile)
    }
}

Function Get-ExcelWorkbook
{
    try { [void][Rhodium.ExcelReporting.ExcelHelper] } catch {
        Add-Type -ReferencedAssemblies "Microsoft.Office.Interop.Excel" '
        using System;
        using System.Diagnostics;
        using System.Collections.Generic;
        using System.Text;
        using Excel = Microsoft.Office.Interop.Excel;
        using System.Runtime.InteropServices;

        namespace Rhodium.ExcelReporting
        {
            public class ExcelHelper
            {
                [DllImport("Oleacc.dll")]
                public static extern int AccessibleObjectFromWindow(
                    int hwnd, uint dwObjectID, byte[] riid,
                    ref Microsoft.Office.Interop.Excel.Window ptr);

                public delegate bool EnumChildCallback(int hwnd, ref int lParam);

                [DllImport("User32.dll")]
                public static extern bool EnumChildWindows(
                    int hWndParent, EnumChildCallback lpEnumFunc,
                    ref int lParam);


                [DllImport("User32.dll")]
                public static extern int GetClassName(
                    int hWnd, StringBuilder lpClassName, int nMaxCount);

                public static bool EnumChildProc(int hwndChild, ref int lParam)
                {
                    StringBuilder buf = new StringBuilder(128);
                    GetClassName(hwndChild, buf, 128);
                    if (buf.ToString() == "EXCEL7")
                    {
                        lParam = hwndChild;
                        return false;
                    }
                    return true;
                }

                public static List<object> GetExcelApplicationList()
                {
                    List<object> results = new List<object>();

                    EnumChildCallback cb;
                    List<Process> procs = new List<Process>();
                    procs.AddRange(Process.GetProcessesByName("excel"));

                    foreach (Process p in procs)
                    {
                        if ((int)p.MainWindowHandle > 0)
                        {
                            int childWindow = 0;
                            cb = new EnumChildCallback(EnumChildProc);
                            EnumChildWindows((int)p.MainWindowHandle, cb, ref childWindow);

                            if (childWindow > 0)
                            {
                                const uint OBJID_NATIVEOM = 0xFFFFFFF0;
                                Guid IID_IDispatch = new Guid("{00020400-0000-0000-C000-000000000046}");
                                Excel.Window window = null;
                                int res = AccessibleObjectFromWindow(childWindow, OBJID_NATIVEOM, IID_IDispatch.ToByteArray(), ref window);
                                if (res >= 0)
                                {
                                    results.Add(window.Application);
                                }
                            }
                        }
                    }

                    return results;
                }
            }
        }
        '
    }

    foreach ($application in [Rhodium.ExcelReporting.ExcelHelper]::GetExcelApplicationList())
    {
        foreach ($workbook in $application.Workbooks)
        {
            $result = [ordered]@{}
            $result.WorkbookName = $workbook.Name
            $result.Handles = [pscustomobject]@{
                Application = $application
                Workbook = $workbook
            }
            [pscustomobject]$result
        }
    }
}

Function Read-ExcelWindow
{
    Param
    (
        [Parameter()] [switch] $GetHidden,
        [Parameter()] [switch] $Ask,
        [Parameter()] [object] $Workbook
    )
    End
    {
        trap { $PSCmdlet.ThrowTerminatingError($_) }

        if (!$Workbook)
        {
            $workbookList = Get-ExcelWorkbook
            if (!$workbookList) { throw "No workbooks are open. If a cell is in Edit mode the workbook won't show." }
            if ($Ask) { $workbook = $workbookList | Out-GridView -Title "Select a workbook" -OutputMode Single }
            else { $workbook = $workbookList | Select-Object -First 1 }
            if (!$Workbook) { throw "No workbook selected." }
        }
        
        $activeSheet = $Workbook.Handles.Workbook.ActiveSheet

        if ($GetHidden)
        {
            $tempWorkbook = $activeSheet.Application.Workbooks.Add()
            $activeSheet.Copy($tempWorkbook.ActiveSheet)
            $activeSheet = $tempWorkbook.ActiveSheet
            $activeSheet.Rows.Hidden = $false
            $activeSheet.Columns.Hidden = $false
            $activeSheet.AutoFilterMode = $false
        }

        $range = $activeSheet.Range("A1").CurrentRegion
        $oldClipboard = [System.Windows.Forms.Clipboard]::GetText()
        [void]$range.Copy()

        $clipboard = [System.Windows.Forms.Clipboard]::GetText()

        [void]$activeSheet.Range("A1").Copy()

        if ($GetHidden)
        {
            $tempWorkbook.Close($false)
        }

        if ($oldClipboard)
        {
            [System.Windows.Forms.Clipboard]::SetText($oldClipboard)
        }
        else { [System.Windows.Forms.Clipboard]::Clear() }

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
        [Parameter()] [string] $Ask,
        [Parameter()] [object] $Workbook,
        [Parameter()] [switch] $AddInputColumns
    )
    Begin
    {
        if (!$ExcelKey) { $ExcelKey = $InputKey }

        if (!$Workbook)
        {
            $workbookList = Get-ExcelWorkbook
            if (!$workbookList) { throw "No workbooks are open. If a cell is in Edit mode the workbook won't show." }
            if ($Ask) { $workbook = $workbookList | Out-GridView -Title "Select a workbook" -OutputMode Single }
            else { $workbook = $workbookList | Select-Object -First 1 }
            if (!$Workbook) { throw "No workbook selected." }
        }

        $excelData = Read-ExcelWindow -Workbook $Workbook -GetHidden
        
        $excelIndices = @{}
        $excelHash = @{}

        $i = 2
        foreach ($data in $excelData)
        {
            $excelKeyValue = & $Script:GetKeyValue $data $Excelkey

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
        foreach ($oldProperty in ($data.PSObject.Properties.Name))
        {
            $excelPropertyHash.Add($oldProperty, $i)
            $i += 1
        }

        $inputProperties = $null
        $usedKeys = @{}
        $activeSheet = $Workbook.Handle.Workbook.ActiveSheet
    }
    Process
    {
        if (!$inputProperties)
        {
            $inputProperties = @{}
            foreach ($property in $InputObject.PSObject.Properties.Name)
            {
                $inputProperties.Add($property, $null)
            }

            if ($AddInputColumns)
            {
                $row1 = $activeSheet.Rows[1]
                $columnList = foreach ($cell in $row1.Cells)
                {
                    $cellObj = [ordered]@{}
                    $cellObj.Column = $cell.Column
                    $cellObj.Header = $cell.Value2
                    if (!$cellObj.Header) { break }
                    [pscustomobject]$cellObj
                }
                $currentColumnNames = $columnList.Header
                $newColumnNames = $InputObject.PSObject.Properties.Name |
                    Where-Object { $_ -notin $currentColumnNames } |
                    Where-Object { $_ -notin $InputKey }
                if ($newColumnNames)
                {
                    $i = $currentColumnNames.Count + 1
                    foreach ($newColumnName in $newColumnNames)
                    {
                        Write-Verbose "Adding column '$newColumnName'"
                        $activeSheet.Cells(1, $i).Value2 = $newColumnName
                        $excelPropertyHash[$newColumnName] = $i
                        $i += 1
                    }
                }
            }
        }
        $keyValue = & $Script:GetKeyValue $InputObject $InputKey
        if ($excelHash.ContainsKey($keyValue))
        {
            if ($usedKeys.$keyValue)
            {
                Write-Warning "The key '$keyValue' has already been used and the previous values may be overwritten."
            }
            $usedKeys.$keyValue = $true
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
                        $activeSheet.Cells.Item($row, $column) = $InputObject.$excelProperty
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
                    $activeSheet.Cells.Item($row, $excelPropertyHash[$TimestampColumn]) = [DateTime]::Now
                }
                $i += 1
            }
        }
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
            $range = $workbook.ActiveSheet.Range("A1")
            [void]$range.Select()
            [void]$range.AutoFilter()
            $workbook.ActiveSheet.Rows[1].Font.FontStyle = 'Bold'
        }
        $excel.Visible = $true

        while ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) -gt 0) { }
        Remove-Variable excel, workbook, range -ErrorAction Ignore
        [GC]::Collect()
    }
}
