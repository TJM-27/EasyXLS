Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

Set-StrictMode -Version 2.0

function Release-ComObject {
    param(
        [Parameter(ValueFromPipeline = $true)]
        [object]$ComObject
    )

    process {
        if ($null -ne $ComObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
        }
    }
}

function ConvertTo-SafeColumnName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [hashtable]$SeenNames
    )

    $baseName = ($Name -replace '\s+', ' ').Trim()
    if ([string]::IsNullOrWhiteSpace($baseName)) {
        $baseName = 'Sheet'
    }

    if (-not $SeenNames.ContainsKey($baseName)) {
        $SeenNames[$baseName] = 1
        return $baseName
    }

    $SeenNames[$baseName]++
    return '{0} ({1})' -f $baseName, $SeenNames[$baseName]
}

function Test-DateNumberFormat {
    param([string]$NumberFormat)

    if ([string]::IsNullOrWhiteSpace($NumberFormat)) {
        return $false
    }

    $format = $NumberFormat.ToLowerInvariant()
    return ($format -match '[dmy]' -or $format -match 'yyyy' -or $format -match 'mmm')
}

function Get-CellDate {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Cell
    )

    $value = $Cell.Value2
    $text = [string]$Cell.Text
    $text = $text.Trim()

    if ($null -ne $value) {
        if ($value -is [datetime]) {
            return $value.Date
        }

        if ($value -is [double] -or $value -is [int] -or $value -is [decimal]) {
            $serial = [double]$value
            if ($serial -ge 1 -and $serial -le 60000 -and (Test-DateNumberFormat -NumberFormat ([string]$Cell.NumberFormat))) {
                try {
                    return [datetime]::FromOADate($serial).Date
                }
                catch {
                    return $null
                }
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    $parsed = [datetime]::MinValue
    $looksDateLike = $text -match '\d{1,2}[/-]\d{1,2}[/-]\d{2,4}' -or
        $text -match '\d{4}[/-]\d{1,2}[/-]\d{1,2}' -or
        $text -match '(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)'

    if ($looksDateLike -and [datetime]::TryParse($text, [ref]$parsed)) {
        return $parsed.Date
    }

    return $null
}

function Test-CurrencyNumberFormat {
    param([string]$NumberFormat)

    if ([string]::IsNullOrWhiteSpace($NumberFormat)) {
        return $false
    }

    return $NumberFormat -match '[$]' -or $NumberFormat -match '\baccounting\b' -or $NumberFormat -match '#,##0'
}

function Get-CellCurrencyValue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Cell
    )

    $value = $Cell.Value2
    $text = ([string]$Cell.Text).Trim()
    $numberFormat = [string]$Cell.NumberFormat
    $hasCurrencySignal = $text -match '[$]' -or $numberFormat -match '[$]' -or (Test-CurrencyNumberFormat -NumberFormat $numberFormat)

    if (-not $hasCurrencySignal) {
        return $null
    }

    if ($null -ne $value -and ($value -is [double] -or $value -is [int] -or $value -is [decimal])) {
        return [decimal]$value
    }

    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    $isNegative = $text -match '^\s*\(.*\)\s*$'
    $clean = $text -replace '[$,\s]', ''
    $clean = $clean -replace '[()]', ''

    $parsed = [decimal]::Zero
    if ([decimal]::TryParse($clean, [System.Globalization.NumberStyles]::Number, [System.Globalization.CultureInfo]::CurrentCulture, [ref]$parsed)) {
        if ($isNegative) {
            return -1 * $parsed
        }

        return $parsed
    }

    return $null
}

function Find-DateColumn {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Worksheet,

        [Parameter(Mandatory = $true)]
        [object]$UsedRange
    )

    $startRow = [int]$UsedRange.Row
    $startColumn = [int]$UsedRange.Column
    $rowCount = [int]$UsedRange.Rows.Count
    $columnCount = [int]$UsedRange.Columns.Count
    $bestColumn = $null
    $bestScore = 0

    for ($columnOffset = 0; $columnOffset -lt $columnCount; $columnOffset++) {
        $column = $startColumn + $columnOffset
        $score = 0

        for ($rowOffset = 0; $rowOffset -lt $rowCount; $rowOffset++) {
            $row = $startRow + $rowOffset
            $cell = $Worksheet.Cells.Item($row, $column)
            try {
                if ($null -ne (Get-CellDate -Cell $cell)) {
                    $score++
                }
            }
            finally {
                Release-ComObject $cell
            }
        }

        if ($score -gt $bestScore) {
            $bestScore = $score
            $bestColumn = $column
        }
    }

    if ($bestScore -lt 1) {
        return $null
    }

    return $bestColumn
}

function Get-FirstCurrencyInRow {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Worksheet,

        [Parameter(Mandatory = $true)]
        [int]$Row,

        [Parameter(Mandatory = $true)]
        [int]$StartColumn,

        [Parameter(Mandatory = $true)]
        [int]$ColumnCount,

        [Parameter(Mandatory = $true)]
        [int]$DateColumn
    )

    for ($columnOffset = 0; $columnOffset -lt $ColumnCount; $columnOffset++) {
        $column = $StartColumn + $columnOffset
        if ($column -eq $DateColumn) {
            continue
        }

        $cell = $Worksheet.Cells.Item($Row, $column)
        try {
            $currencyValue = Get-CellCurrencyValue -Cell $cell
            if ($null -ne $currencyValue) {
                return $currencyValue
            }
        }
        finally {
            Release-ComObject $cell
        }
    }

    return $null
}

function Get-ZipEntryText {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Compression.ZipArchive]$Archive,

        [Parameter(Mandatory = $true)]
        [string]$EntryName
    )

    $entry = $Archive.GetEntry($EntryName)
    if ($null -eq $entry) {
        return $null
    }

    $stream = $entry.Open()
    try {
        $reader = New-Object System.IO.StreamReader($stream)
        try {
            return $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
}

function Get-ExcelColumnIndex {
    param([string]$CellReference)

    $letters = ($CellReference -replace '[^A-Za-z]', '').ToUpperInvariant()
    $index = 0

    foreach ($character in $letters.ToCharArray()) {
        $index = ($index * 26) + ([int][char]$character - [int][char]'A' + 1)
    }

    return $index
}

function Get-ObjectPropertyValue {
    param(
        [object]$Object,
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if ($null -eq $Object) {
        return $null
    }

    if ($Object.PSObject.Properties.Name -contains $Name) {
        return $Object.$Name
    }

    return $null
}

function Test-XlsxDateFormatId {
    param([int]$NumberFormatId)

    return @(
        14, 15, 16, 17, 18, 19, 20, 21, 22,
        27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
        45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58
    ) -contains $NumberFormatId
}

function Test-XlsxCurrencyFormatId {
    param([int]$NumberFormatId)

    return @(5, 6, 7, 8, 37, 38, 39, 40, 41, 42, 43, 44) -contains $NumberFormatId
}

function Get-XlsxStyleMaps {
    param([xml]$StylesXml)

    $customFormats = @{}
    $styleDateFlags = @{}
    $styleCurrencyFlags = @{}

    $styleSheet = Get-ObjectPropertyValue -Object $StylesXml -Name 'styleSheet'
    $numFmts = Get-ObjectPropertyValue -Object $styleSheet -Name 'numFmts'
    $cellXfs = Get-ObjectPropertyValue -Object $styleSheet -Name 'cellXfs'

    if ($null -ne $numFmts) {
        foreach ($format in @(Get-ObjectPropertyValue -Object $numFmts -Name 'numFmt')) {
            $id = [int]$format.numFmtId
            $customFormats[$id] = [string]$format.formatCode
        }
    }

    if ($null -ne $cellXfs) {
        $styleIndex = 0
        foreach ($format in @(Get-ObjectPropertyValue -Object $cellXfs -Name 'xf')) {
            $numberFormatId = [int]$format.numFmtId
            $formatCode = ''
            if ($customFormats.ContainsKey($numberFormatId)) {
                $formatCode = $customFormats[$numberFormatId]
            }

            $styleDateFlags[$styleIndex] = (Test-XlsxDateFormatId -NumberFormatId $numberFormatId) -or (Test-DateNumberFormat -NumberFormat $formatCode)
            $styleCurrencyFlags[$styleIndex] = (Test-XlsxCurrencyFormatId -NumberFormatId $numberFormatId) -or (Test-CurrencyNumberFormat -NumberFormat $formatCode)
            $styleIndex++
        }
    }

    return [pscustomobject]@{
        IsDate = $styleDateFlags
        IsCurrency = $styleCurrencyFlags
    }
}

function Get-XlsxSharedStrings {
    param([xml]$SharedStringsXml)

    $strings = New-Object System.Collections.Generic.List[string]
    $sst = Get-ObjectPropertyValue -Object $SharedStringsXml -Name 'sst'
    if ($null -eq $sst) {
        return $strings
    }

    foreach ($item in @(Get-ObjectPropertyValue -Object $sst -Name 'si')) {
        $textNode = Get-ObjectPropertyValue -Object $item -Name 't'
        if ($null -ne $textNode) {
            [void]$strings.Add([string]$textNode)
            continue
        }

        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($run in @(Get-ObjectPropertyValue -Object $item -Name 'r')) {
            $runTextNode = Get-ObjectPropertyValue -Object $run -Name 't'
            if ($null -ne $runTextNode) {
                [void]$parts.Add([string]$runTextNode)
            }
        }

        [void]$strings.Add(($parts -join ''))
    }

    return $strings
}

function Convert-XlsxCellValue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Cell,

        [object]$SharedStrings,

        [Parameter(Mandatory = $true)]
        [object]$StyleMaps
    )

    $styleIndex = 0
    $styleAttribute = Get-ObjectPropertyValue -Object $Cell -Name 's'
    if ($null -ne $styleAttribute) {
        $styleIndex = [int]$styleAttribute
    }

    $type = [string](Get-ObjectPropertyValue -Object $Cell -Name 't')
    $rawValue = ''
    $valueNode = Get-ObjectPropertyValue -Object $Cell -Name 'v'
    if ($null -ne $valueNode) {
        $rawValue = [string]$valueNode
    }

    $text = $rawValue
    if ($type -eq 's' -and $rawValue -match '^\d+$') {
        $sharedIndex = [int]$rawValue
        if ($sharedIndex -ge 0 -and $sharedIndex -lt $SharedStrings.Count) {
            $text = $SharedStrings[$sharedIndex]
        }
    }
    elseif ($type -eq 'inlineStr') {
        $inlineString = Get-ObjectPropertyValue -Object $Cell -Name 'is'
        $inlineText = Get-ObjectPropertyValue -Object $inlineString -Name 't'
        if ($null -ne $inlineText) {
            $text = [string]$inlineText
        }
    }

    $number = $null
    $parsedNumber = [double]::NaN
    if ([double]::TryParse($rawValue, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsedNumber)) {
        $number = $parsedNumber
    }

    $isDateStyle = $StyleMaps.IsDate.ContainsKey($styleIndex) -and $StyleMaps.IsDate[$styleIndex]
    $isCurrencyStyle = $StyleMaps.IsCurrency.ContainsKey($styleIndex) -and $StyleMaps.IsCurrency[$styleIndex]

    $date = $null
    if ($null -ne $number -and $isDateStyle -and $number -ge 1 -and $number -le 60000) {
        try {
            $date = [datetime]::FromOADate($number).Date
        }
        catch {
            $date = $null
        }
    }
    elseif (-not [string]::IsNullOrWhiteSpace($text)) {
        $parsedDate = [datetime]::MinValue
        $looksDateLike = $text -match '\d{1,2}[/-]\d{1,2}[/-]\d{2,4}' -or
            $text -match '\d{4}[/-]\d{1,2}[/-]\d{1,2}' -or
            $text -match '(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)'

        if ($looksDateLike -and [datetime]::TryParse($text, [ref]$parsedDate)) {
            $date = $parsedDate.Date
        }
    }

    $currency = $null
    if ($null -ne $number -and ($isCurrencyStyle -or $text -match '[$]')) {
        $currency = [decimal]$number
    }
    elseif ($text -match '[$]') {
        $isNegative = $text -match '^\s*\(.*\)\s*$'
        $clean = $text -replace '[$,\s]', ''
        $clean = $clean -replace '[()]', ''
        $parsedCurrency = [decimal]::Zero

        if ([decimal]::TryParse($clean, [System.Globalization.NumberStyles]::Number, [System.Globalization.CultureInfo]::CurrentCulture, [ref]$parsedCurrency)) {
            if ($isNegative) {
                $currency = -1 * $parsedCurrency
            }
            else {
                $currency = $parsedCurrency
            }
        }
    }

    return [pscustomobject]@{
        ColumnIndex = Get-ExcelColumnIndex -CellReference ([string]$Cell.r)
        Text = $text
        Number = $number
        Date = $date
        Currency = $currency
    }
}

function Import-XlsxTransactions {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [scriptblock]$StatusCallback
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Workbook not found: $Path"
    }

    $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
    $sheetResults = New-Object System.Collections.Generic.List[object]
    $allDates = New-Object 'System.Collections.Generic.SortedSet[datetime]'
    $sheetNameCounts = @{
        Date = 1
    }
    $warnings = New-Object System.Collections.Generic.List[string]

    try {
        if ($StatusCallback) {
            & $StatusCallback 'Reading XLSX workbook...'
        }

        [xml]$workbookXml = Get-ZipEntryText -Archive $archive -EntryName 'xl/workbook.xml'
        [xml]$relationshipsXml = Get-ZipEntryText -Archive $archive -EntryName 'xl/_rels/workbook.xml.rels'
        $sharedStringsText = Get-ZipEntryText -Archive $archive -EntryName 'xl/sharedStrings.xml'
        $stylesText = Get-ZipEntryText -Archive $archive -EntryName 'xl/styles.xml'

        $sharedStringsXml = $null
        if (-not [string]::IsNullOrWhiteSpace($sharedStringsText)) {
            [xml]$sharedStringsXml = $sharedStringsText
        }

        $stylesXml = $null
        if (-not [string]::IsNullOrWhiteSpace($stylesText)) {
            [xml]$stylesXml = $stylesText
        }

        $sharedStrings = Get-XlsxSharedStrings -SharedStringsXml $sharedStringsXml
        $styleMaps = Get-XlsxStyleMaps -StylesXml $stylesXml

        $relationshipTargets = @{}
        foreach ($relationship in @($relationshipsXml.Relationships.Relationship)) {
            $target = [string]$relationship.Target
            if ($target.StartsWith('/')) {
                $target = $target.TrimStart('/')
            }
            else {
                $target = 'xl/' + $target
            }

            $relationshipTargets[[string]$relationship.Id] = $target
        }

        foreach ($sheet in @($workbookXml.workbook.sheets.sheet)) {
            $sheetName = ConvertTo-SafeColumnName -Name ([string]$sheet.name) -SeenNames $sheetNameCounts
            $relationshipId = [string]$sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')

            if (-not $relationshipTargets.ContainsKey($relationshipId)) {
                [void]$warnings.Add(("Skipped '{0}': worksheet relationship not found." -f $sheetName))
                continue
            }

            if ($StatusCallback) {
                & $StatusCallback ("Scanning sheet: {0}" -f $sheetName)
            }

            [xml]$sheetXml = Get-ZipEntryText -Archive $archive -EntryName $relationshipTargets[$relationshipId]
            $rows = New-Object System.Collections.Generic.List[object]
            $dateScores = @{}

            foreach ($row in @($sheetXml.worksheet.sheetData.row)) {
                $cells = New-Object System.Collections.Generic.List[object]

                foreach ($cell in @($row.c)) {
                    $converted = Convert-XlsxCellValue -Cell $cell -SharedStrings $sharedStrings -StyleMaps $styleMaps
                    [void]$cells.Add($converted)

                    if ($null -ne $converted.Date) {
                        if (-not $dateScores.ContainsKey($converted.ColumnIndex)) {
                            $dateScores[$converted.ColumnIndex] = 0
                        }
                        $dateScores[$converted.ColumnIndex]++
                    }
                }

                [void]$rows.Add($cells)
            }

            if ($dateScores.Count -eq 0) {
                [void]$warnings.Add(("Skipped '{0}': no date column found." -f $sheetName))
                continue
            }

            $dateColumn = ($dateScores.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 1).Key
            $transactions = @{}

            foreach ($cells in $rows) {
                $dateCell = $cells | Where-Object { $_.ColumnIndex -eq $dateColumn -and $null -ne $_.Date } | Select-Object -First 1
                if ($null -eq $dateCell) {
                    continue
                }

                $currencyCell = $cells | Where-Object { $_.ColumnIndex -ne $dateColumn -and $null -ne $_.Currency } | Sort-Object -Property ColumnIndex | Select-Object -First 1
                if ($null -eq $currencyCell) {
                    $currencyCell = $cells | Where-Object { $_.ColumnIndex -ne $dateColumn -and $null -ne $_.Number -and $null -eq $_.Date } | Sort-Object -Property ColumnIndex | Select-Object -First 1
                }

                if ($null -eq $currencyCell) {
                    continue
                }

                $dateKey = $dateCell.Date.Date
                if (-not $transactions.ContainsKey($dateKey)) {
                    if ($null -ne $currencyCell.Currency) {
                        $transactions[$dateKey] = $currencyCell.Currency
                    }
                    else {
                        $transactions[$dateKey] = [decimal]$currencyCell.Number
                    }
                }

                [void]$allDates.Add($dateKey)
            }

            if ($transactions.Count -eq 0) {
                continue
            }

            [void]$sheetResults.Add([pscustomobject]@{
                Name = $sheetName
                Transactions = $transactions
            })
        }
    }
    finally {
        $archive.Dispose()
    }

    return [pscustomobject]@{
        Sheets = $sheetResults
        Dates = $allDates
        Warnings = $warnings
    }
}

function Get-LocalWorkbookCopy {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [scriptblock]$StatusCallback
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Workbook not found: $Path"
    }

    $sourcePath = (Resolve-Path -LiteralPath $Path).ProviderPath
    $baseDirectory = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($baseDirectory)) {
        $baseDirectory = (Get-Location).Path
    }

    $workingDirectory = Join-Path $baseDirectory '_EasyXLS_WorkingCopies'
    if (-not (Test-Path -LiteralPath $workingDirectory -PathType Container)) {
        [void](New-Item -Path $workingDirectory -ItemType Directory -Force)
    }

    $extension = [System.IO.Path]::GetExtension($sourcePath)
    $copyPath = Join-Path $workingDirectory ('EasyXLS_{0}{1}' -f ([guid]::NewGuid().ToString('N')), $extension)

    if ($StatusCallback) {
        & $StatusCallback 'Copying workbook locally...'
    }

    Copy-Item -LiteralPath $sourcePath -Destination $copyPath -Force
    return $copyPath
}

function Open-ExcelWorkbookReadOnly {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Excel,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $workbooks = $null
    try {
        $workbooks = $Excel.Workbooks

        try {
            return $workbooks.Open(
                $Path,
                0,
                $true,
                [Type]::Missing,
                [Type]::Missing,
                [Type]::Missing,
                $true,
                [Type]::Missing,
                [Type]::Missing,
                $false,
                $false,
                [Type]::Missing,
                $false,
                $true,
                1
            )
        }
        catch {
            return $workbooks.Open($Path)
        }
    }
    finally {
        if ($null -ne $workbooks) {
            Release-ComObject $workbooks
        }
    }
}

function Import-XlsTransactions {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [scriptblock]$StatusCallback
    )

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($extension -eq '.xlsx' -or $extension -eq '.xlsm') {
        return Import-XlsxTransactions -Path $Path -StatusCallback $StatusCallback
    }

    $excel = $null
    $workbook = $null
    $localWorkbookPath = $null
    $sheetResults = New-Object System.Collections.Generic.List[object]
    $allDates = New-Object 'System.Collections.Generic.SortedSet[datetime]'
    $sheetNameCounts = @{
        Date = 1
    }
    $warnings = New-Object System.Collections.Generic.List[string]

    try {
        if ($StatusCallback) {
            & $StatusCallback 'Starting Excel...'
        }

        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $localWorkbookPath = Get-LocalWorkbookCopy -Path $Path -StatusCallback $StatusCallback
        $workbook = Open-ExcelWorkbookReadOnly -Excel $excel -Path $localWorkbookPath

        foreach ($worksheet in $workbook.Worksheets) {
            $usedRange = $null
            $sheetName = ConvertTo-SafeColumnName -Name ([string]$worksheet.Name) -SeenNames $sheetNameCounts
            $transactions = @{}

            try {
                if ($StatusCallback) {
                    & $StatusCallback ("Scanning sheet: {0}" -f $sheetName)
                }

                $usedRange = $worksheet.UsedRange
                $startRow = [int]$usedRange.Row
                $startColumn = [int]$usedRange.Column
                $rowCount = [int]$usedRange.Rows.Count
                $columnCount = [int]$usedRange.Columns.Count

                $dateColumn = Find-DateColumn -Worksheet $worksheet -UsedRange $usedRange
                if ($null -eq $dateColumn) {
                    [void]$warnings.Add(("Skipped '{0}': no date column found." -f $sheetName))
                    continue
                }

                for ($rowOffset = 0; $rowOffset -lt $rowCount; $rowOffset++) {
                    $row = $startRow + $rowOffset
                    $dateCell = $worksheet.Cells.Item($row, $dateColumn)

                    try {
                        $date = Get-CellDate -Cell $dateCell
                    }
                    finally {
                        Release-ComObject $dateCell
                    }

                    if ($null -eq $date) {
                        continue
                    }

                    $currencyValue = Get-FirstCurrencyInRow -Worksheet $worksheet -Row $row -StartColumn $startColumn -ColumnCount $columnCount -DateColumn $dateColumn
                    if ($null -eq $currencyValue) {
                        continue
                    }

                    $dateKey = $date.Date
                    if (-not $transactions.ContainsKey($dateKey)) {
                        $transactions[$dateKey] = $currencyValue
                    }

                    [void]$allDates.Add($dateKey)
                }

                if ($transactions.Count -eq 0) {
                    continue
                }

                [void]$sheetResults.Add([pscustomobject]@{
                    Name = $sheetName
                    Transactions = $transactions
                })
            }
            finally {
                if ($null -ne $usedRange) {
                    Release-ComObject $usedRange
                }

                Release-ComObject $worksheet
            }
        }
    }
    finally {
        if ($null -ne $workbook) {
            $workbook.Close($false)
            Release-ComObject $workbook
        }

        if ($null -ne $excel) {
            $excel.Quit()
            Release-ComObject $excel
        }

        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()

        if ($null -ne $localWorkbookPath -and (Test-Path -LiteralPath $localWorkbookPath -PathType Leaf)) {
            try {
                Remove-Item -LiteralPath $localWorkbookPath -Force
            }
            catch {
                Write-EasyXlsLog -Message ("Unable to remove working copy '{0}': {1}" -f $localWorkbookPath, $_.Exception.Message)
            }
        }
    }

    return [pscustomobject]@{
        Sheets = $sheetResults
        Dates = $allDates
        Warnings = $warnings
    }
}

function ConvertTo-TransactionDataTable {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ImportResult
    )

    $table = New-Object System.Data.DataTable
    [void]$table.Columns.Add('Date', [string])

    foreach ($sheet in $ImportResult.Sheets) {
        [void]$table.Columns.Add($sheet.Name, [string])
    }

    foreach ($date in $ImportResult.Dates) {
        $row = $table.NewRow()
        $row['Date'] = $date.ToString('yyyy-MM-dd')

        foreach ($sheet in $ImportResult.Sheets) {
            if ($sheet.Transactions.ContainsKey($date)) {
                $row[$sheet.Name] = ('{0:C2}' -f $sheet.Transactions[$date])
            }
            else {
                $row[$sheet.Name] = ''
            }
        }

        [void]$table.Rows.Add($row)
    }

    return ,$table
}

function Write-EasyXlsLog {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    try {
        $baseDirectory = $PSScriptRoot
        if ([string]::IsNullOrWhiteSpace($baseDirectory)) {
            $baseDirectory = (Get-Location).Path
        }

        $logPath = Join-Path $baseDirectory 'EasyXLS-error.log'
        $line = '{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
        Add-Content -Path $logPath -Value $line
    }
    catch {
        # Logging should never prevent the GUI from showing the actual error.
    }
}

function ConvertTo-PdfEscapedText {
    param([string]$Text)

    if ($null -eq $Text) {
        return ''
    }

    return ($Text -replace '\\', '\\' -replace '\(', '\(' -replace '\)', '\)' -replace "`r", ' ' -replace "`n", ' ')
}

function Add-PdfText {
    param(
        [System.Text.StringBuilder]$Builder,
        [string]$Text,
        [double]$X,
        [double]$Y,
        [double]$FontSize
    )

    $escaped = ConvertTo-PdfEscapedText -Text $Text
    [void]$Builder.AppendFormat([System.Globalization.CultureInfo]::InvariantCulture, "BT /F1 {0:0.##} Tf {1:0.##} {2:0.##} Td ({3}) Tj ET`n", $FontSize, $X, $Y, $escaped)
}

function Add-PdfRectangle {
    param(
        [System.Text.StringBuilder]$Builder,
        [double]$X,
        [double]$Y,
        [double]$Width,
        [double]$Height,
        [System.Drawing.Color]$FillColor,
        [bool]$Stroke
    )

    if ($null -ne $FillColor -and $FillColor -ne [System.Drawing.Color]::Empty) {
        [void]$Builder.AppendFormat(
            [System.Globalization.CultureInfo]::InvariantCulture,
            "q {0:0.###} {1:0.###} {2:0.###} rg {3:0.##} {4:0.##} {5:0.##} {6:0.##} re f Q`n",
            ($FillColor.R / 255),
            ($FillColor.G / 255),
            ($FillColor.B / 255),
            $X,
            $Y,
            $Width,
            $Height
        )
    }

    if ($Stroke) {
        [void]$Builder.AppendFormat([System.Globalization.CultureInfo]::InvariantCulture, "q 0.65 0.65 0.65 RG {0:0.##} {1:0.##} {2:0.##} {3:0.##} re S Q`n", $X, $Y, $Width, $Height)
    }
}

function Get-TruncatedPdfText {
    param(
        [string]$Text,
        [double]$ColumnWidth,
        [double]$FontSize
    )

    if ([string]::IsNullOrEmpty($Text)) {
        return ''
    }

    $maxCharacters = [Math]::Max(1, [int][Math]::Floor(($ColumnWidth - 6) / ($FontSize * 0.52)))
    if ($Text.Length -le $maxCharacters) {
        return $Text
    }

    if ($maxCharacters -le 3) {
        return $Text.Substring(0, $maxCharacters)
    }

    return $Text.Substring(0, $maxCharacters - 3) + '...'
}

function Export-GridToPdf {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.DataGridView]$Grid,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [System.Drawing.Color]$HeaderColor,

        [Parameter(Mandatory = $true)]
        [System.Drawing.Color]$HighlightColor
    )

    $columns = @($Grid.Columns | Where-Object { $_.Visible } | Sort-Object -Property DisplayIndex)
    if ($columns.Count -eq 0) {
        throw 'There are no visible columns to export.'
    }

    $rows = @($Grid.Rows | Where-Object { -not $_.IsNewRow })
    if ($rows.Count -eq 0) {
        throw 'There are no rows to export.'
    }

    $pageWidth = 792.0
    $pageHeight = 612.0
    $margin = 28.0
    $tableWidth = $pageWidth - (2 * $margin)
    $headerHeight = 20.0
    $rowHeight = 16.0
    $fontSize = 7.0
    $titleFontSize = 11.0
    $usableHeight = $pageHeight - (2 * $margin) - 24.0 - $headerHeight
    $rowsPerPage = [Math]::Max(1, [int][Math]::Floor($usableHeight / $rowHeight))

    $totalGridWidth = 0.0
    foreach ($column in $columns) {
        $totalGridWidth += [Math]::Max(35, $column.Width)
    }

    $columnWidths = @{}
    foreach ($column in $columns) {
        $columnWidths[$column.Index] = ([Math]::Max(35, $column.Width) / $totalGridWidth) * $tableWidth
    }

    $contentStreams = New-Object System.Collections.Generic.List[string]
    for ($rowStart = 0; $rowStart -lt $rows.Count; $rowStart += $rowsPerPage) {
        $builder = New-Object System.Text.StringBuilder
        $pageNumber = $contentStreams.Count + 1
        Add-PdfText -Builder $builder -Text ("EasyXLS Export - Page {0}" -f $pageNumber) -X $margin -Y ($pageHeight - $margin + 4) -FontSize $titleFontSize

        $x = $margin
        $headerY = $pageHeight - $margin - 22
        foreach ($column in $columns) {
            $width = $columnWidths[$column.Index]
            $headerBackColor = $HeaderColor
            if ($column.HeaderCell.Style.BackColor.ToArgb() -eq $HighlightColor.ToArgb()) {
                $headerBackColor = $HighlightColor
            }

            Add-PdfRectangle -Builder $builder -X $x -Y ($headerY - $headerHeight + 4) -Width $width -Height $headerHeight -FillColor $headerBackColor -Stroke $true
            Add-PdfText -Builder $builder -Text (Get-TruncatedPdfText -Text $column.HeaderText -ColumnWidth $width -FontSize $fontSize) -X ($x + 3) -Y ($headerY - 10) -FontSize $fontSize
            $x += $width
        }

        $rowsThisPage = [Math]::Min($rowsPerPage, $rows.Count - $rowStart)
        for ($offset = 0; $offset -lt $rowsThisPage; $offset++) {
            $row = $rows[$rowStart + $offset]
            $rowY = $headerY - $headerHeight - (($offset + 1) * $rowHeight) + 4
            $x = $margin

            foreach ($column in $columns) {
                $cell = $row.Cells[$column.Index]
                $width = $columnWidths[$column.Index]
                $backColor = [System.Drawing.Color]::White
                if ($cell.Style.BackColor.ToArgb() -eq $HighlightColor.ToArgb()) {
                    $backColor = $HighlightColor
                }

                Add-PdfRectangle -Builder $builder -X $x -Y $rowY -Width $width -Height $rowHeight -FillColor $backColor -Stroke $true
                $valueText = ''
                if ($null -ne $cell.Value) {
                    $valueText = [string]$cell.Value
                }

                Add-PdfText -Builder $builder -Text (Get-TruncatedPdfText -Text $valueText -ColumnWidth $width -FontSize $fontSize) -X ($x + 3) -Y ($rowY + 5) -FontSize $fontSize
                $x += $width
            }
        }

        [void]$contentStreams.Add($builder.ToString())
    }

    $objects = New-Object System.Collections.Generic.List[string]
    [void]$objects.Add('<< /Type /Catalog /Pages 2 0 R >>')
    $kids = New-Object System.Collections.Generic.List[string]
    for ($index = 0; $index -lt $contentStreams.Count; $index++) {
        $pageObjectNumber = 3 + ($index * 2)
        [void]$kids.Add(('{0} 0 R' -f $pageObjectNumber))
    }
    [void]$objects.Add(('<< /Type /Pages /Kids [{0}] /Count {1} >>' -f ($kids -join ' '), $contentStreams.Count))

    for ($index = 0; $index -lt $contentStreams.Count; $index++) {
        $contentObjectNumber = 4 + ($index * 2)
        $streamBytes = [System.Text.Encoding]::ASCII.GetBytes($contentStreams[$index])
        [void]$objects.Add(('<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {0} {1}] /Resources << /Font << /F1 {2} 0 R >> >> /Contents {3} 0 R >>' -f $pageWidth, $pageHeight, (3 + ($contentStreams.Count * 2)), $contentObjectNumber))
        [void]$objects.Add(("<< /Length {0} >>`nstream`n{1}endstream" -f $streamBytes.Length, $contentStreams[$index]))
    }

    [void]$objects.Add('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>')

    $pdf = New-Object System.Text.StringBuilder
    [void]$pdf.Append("%PDF-1.4`n")
    $offsets = New-Object System.Collections.Generic.List[int]
    for ($index = 0; $index -lt $objects.Count; $index++) {
        [void]$offsets.Add([System.Text.Encoding]::ASCII.GetByteCount($pdf.ToString()))
        [void]$pdf.AppendFormat("{0} 0 obj`n{1}`nendobj`n", ($index + 1), $objects[$index])
    }

    $xrefOffset = [System.Text.Encoding]::ASCII.GetByteCount($pdf.ToString())
    [void]$pdf.AppendFormat("xref`n0 {0}`n", ($objects.Count + 1))
    [void]$pdf.Append("0000000000 65535 f `n")
    foreach ($offset in $offsets) {
        [void]$pdf.AppendFormat("{0:0000000000} 00000 n `n", $offset)
    }
    [void]$pdf.AppendFormat("trailer`n<< /Size {0} /Root 1 0 R >>`nstartxref`n{1}`n%%EOF", ($objects.Count + 1), $xrefOffset)

    [System.IO.File]::WriteAllBytes($Path, [System.Text.Encoding]::ASCII.GetBytes($pdf.ToString()))
}

function ConvertTo-NumericSearchText {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ''
    }

    return ($Text -replace '\D', '')
}

function ConvertTo-SearchDecimal {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $null
    }

    $clean = $Text.Trim()
    $isNegative = $clean -match '^\s*\(.*\)\s*$' -or $clean.StartsWith('-')
    $clean = $clean -replace '[$,\s]', ''
    $clean = $clean -replace '[()]', ''

    $parsed = [decimal]::Zero
    if ([decimal]::TryParse($clean, [System.Globalization.NumberStyles]::Number, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) {
        if ($isNegative) {
            return -1 * $parsed
        }

        return $parsed
    }

    return $null
}

function Test-NumericSearchMatch {
    param(
        [string]$Text,
        [decimal]$SearchValue
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $false
    }

    $numberPattern = '\(?-?\$?\d{1,3}(?:,\d{3})*(?:\.\d+)?\)?|\(?-?\$?\d+(?:\.\d+)?\)?'
    foreach ($match in [regex]::Matches($Text, $numberPattern)) {
        $candidate = ConvertTo-SearchDecimal -Text $match.Value
        if ($null -ne $candidate -and $candidate -eq $SearchValue) {
            return $true
        }
    }

    return $false
}

function Show-EasyXlsGui {
    $greenBackColor = [System.Drawing.Color]::FromArgb(242, 248, 242)
    $greenPanelColor = [System.Drawing.Color]::FromArgb(225, 239, 225)
    $greenHeaderColor = [System.Drawing.Color]::FromArgb(196, 225, 196)
    $greenAccentColor = [System.Drawing.Color]::FromArgb(76, 132, 86)
    $greenButtonColor = [System.Drawing.Color]::FromArgb(214, 235, 214)
    $greenButtonHoverColor = [System.Drawing.Color]::FromArgb(199, 225, 199)
    $highlightColor = [System.Drawing.Color]::FromArgb(255, 242, 120)
    $defaultCellColor = [System.Drawing.Color]::White
    $alternateCellColor = [System.Drawing.Color]::FromArgb(248, 252, 248)
    $baseFont = New-Object System.Drawing.Font('Segoe UI', 9)
    $smallFont = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $headerFont = New-Object System.Drawing.Font('Segoe UI Semibold', 9, [System.Drawing.FontStyle]::Bold)

    $styleButton = {
        param([System.Windows.Forms.Button]$Button)

        $Button.FlatStyle = 'Flat'
        $Button.BackColor = $greenButtonColor
        $Button.ForeColor = [System.Drawing.Color]::FromArgb(35, 65, 40)
        $Button.Font = $baseFont
        $Button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(156, 196, 156)
        $Button.FlatAppearance.MouseOverBackColor = $greenButtonHoverColor
        $Button.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(184, 216, 184)
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'EasyXLS Transaction Table'
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object System.Drawing.Size(1000, 650)
    $form.MinimumSize = New-Object System.Drawing.Size(800, 500)
    $form.BackColor = $greenBackColor
    $form.Font = $baseFont
    $iconPath = Join-Path $PSScriptRoot 'EasyXLS.ico'
    $formIcon = $null
    if (Test-Path -LiteralPath $iconPath -PathType Leaf) {
        $formIcon = New-Object System.Drawing.Icon($iconPath)
        $form.Icon = $formIcon
        $form.Add_FormClosed({
            if ($null -ne $formIcon) {
                $formIcon.Dispose()
            }
        })
    }

    $topPanel = New-Object System.Windows.Forms.Panel
    $topPanel.Dock = 'Top'
    $topPanel.Height = 94
    $topPanel.BackColor = $greenPanelColor

    $openButton = New-Object System.Windows.Forms.Button
    $openButton.Text = 'Open XLS...'
    $openButton.Width = 120
    $openButton.Height = 32
    $openButton.Left = 14
    $openButton.Top = 12
    & $styleButton $openButton

    $fileLabel = New-Object System.Windows.Forms.Label
    $fileLabel.AutoSize = $false
    $fileLabel.Left = 150
    $fileLabel.Top = 19
    $fileLabel.Width = 805
    $fileLabel.Height = 22
    $fileLabel.Anchor = 'Top,Left,Right'
    $fileLabel.Text = 'Select an .xls workbook to build the transaction table.'
    $fileLabel.BackColor = $greenPanelColor
    $fileLabel.ForeColor = [System.Drawing.Color]::FromArgb(45, 75, 45)
    $fileLabel.Font = $smallFont

    $searchLabel = New-Object System.Windows.Forms.Label
    $searchLabel.Text = 'Search:'
    $searchLabel.AutoSize = $true
    $searchLabel.Left = 16
    $searchLabel.Top = 62
    $searchLabel.BackColor = $greenPanelColor
    $searchLabel.ForeColor = [System.Drawing.Color]::FromArgb(45, 75, 45)
    $searchLabel.Font = $smallFont

    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Left = 76
    $searchBox.Top = 58
    $searchBox.Width = 260
    $searchBox.Height = 26
    $searchBox.Anchor = 'Top,Left'
    $searchBox.BorderStyle = 'FixedSingle'
    $searchBox.Font = $baseFont

    $searchButton = New-Object System.Windows.Forms.Button
    $searchButton.Text = 'Find'
    $searchButton.Left = 348
    $searchButton.Top = 56
    $searchButton.Width = 70
    $searchButton.Height = 30
    & $styleButton $searchButton

    $clearSearchButton = New-Object System.Windows.Forms.Button
    $clearSearchButton.Text = 'Clear'
    $clearSearchButton.Left = 426
    $clearSearchButton.Top = 56
    $clearSearchButton.Width = 70
    $clearSearchButton.Height = 30
    & $styleButton $clearSearchButton

    $exportPdfButton = New-Object System.Windows.Forms.Button
    $exportPdfButton.Text = 'Export PDF...'
    $exportPdfButton.Left = 504
    $exportPdfButton.Top = 56
    $exportPdfButton.Width = 105
    $exportPdfButton.Height = 30
    & $styleButton $exportPdfButton

    $searchResultLabel = New-Object System.Windows.Forms.Label
    $searchResultLabel.AutoSize = $false
    $searchResultLabel.Left = 620
    $searchResultLabel.Top = 62
    $searchResultLabel.Width = 335
    $searchResultLabel.Height = 22
    $searchResultLabel.Anchor = 'Top,Left,Right'
    $searchResultLabel.Text = 'No search active.'
    $searchResultLabel.BackColor = $greenPanelColor
    $searchResultLabel.ForeColor = [System.Drawing.Color]::FromArgb(45, 75, 45)
    $searchResultLabel.Font = $smallFont

    $statusStrip = New-Object System.Windows.Forms.StatusStrip
    $statusStrip.BackColor = $greenPanelColor
    $statusStrip.SizingGrip = $false
    $statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $statusLabel.Text = 'Ready.'
    $statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(45, 75, 45)
    [void]$statusStrip.Items.Add($statusLabel)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = 'Fill'
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AllowUserToOrderColumns = $true
    $grid.AllowUserToResizeColumns = $true
    $grid.ReadOnly = $true
    $grid.AutoSizeColumnsMode = 'None'
    $grid.ColumnHeadersVisible = $true
    $grid.ColumnHeadersHeight = 34
    $grid.ColumnHeadersHeightSizeMode = 'DisableResizing'
    $grid.SelectionMode = 'FullRowSelect'
    $grid.RowHeadersVisible = $false
    $grid.ScrollBars = 'Both'
    $grid.BackgroundColor = $greenBackColor
    $grid.BorderStyle = 'None'
    $grid.CellBorderStyle = 'SingleHorizontal'
    $grid.ColumnHeadersBorderStyle = 'Single'
    $grid.GridColor = [System.Drawing.Color]::FromArgb(218, 232, 218)
    $grid.EnableHeadersVisualStyles = $false
    $grid.Font = $baseFont
    $grid.RowTemplate.Height = 26
    $grid.DefaultCellStyle.BackColor = $defaultCellColor
    $grid.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(35, 35, 35)
    $grid.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(5, 2, 5, 2)
    $grid.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(183, 222, 183)
    $grid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
    $grid.AlternatingRowsDefaultCellStyle.BackColor = $alternateCellColor
    $grid.ColumnHeadersDefaultCellStyle.BackColor = $greenHeaderColor
    $grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(25, 65, 32)
    $grid.ColumnHeadersDefaultCellStyle.Font = $headerFont
    $grid.ColumnHeadersDefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(5, 4, 5, 4)
    $grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = $greenHeaderColor
    $grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(25, 65, 32)

    $columnHeaderMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $hideColumnMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $hideColumnMenuItem.Text = 'Hide This Column'
    $unhideAllMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $unhideAllMenuItem.Text = 'Unhide All Columns'
    [void]$columnHeaderMenu.Items.Add($hideColumnMenuItem)
    [void]$columnHeaderMenu.Items.Add($unhideAllMenuItem)
    $searchMatchedCellKeys = @{}
    $searchMatchedHeaderIndexes = @{}

    $openDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openDialog.Filter = 'Excel Workbooks (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|Modern Excel Workbooks (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|Excel 97-2003 Workbook (*.xls)|*.xls|All files (*.*)|*.*'
    $openDialog.FilterIndex = 1
    $openDialog.Title = 'Select an XLS workbook'

    $savePdfDialog = New-Object System.Windows.Forms.SaveFileDialog
    $savePdfDialog.Filter = 'PDF Files (*.pdf)|*.pdf|All files (*.*)|*.*'
    $savePdfDialog.FilterIndex = 1
    $savePdfDialog.DefaultExt = 'pdf'
    $savePdfDialog.AddExtension = $true
    $savePdfDialog.OverwritePrompt = $true
    $savePdfDialog.Title = 'Export EasyXLS Table to PDF'

    $setStatus = {
        param([string]$Message)
        $statusLabel.Text = $Message
        [System.Windows.Forms.Application]::DoEvents()
    }

    $clearSearchHighlights = {
        $searchMatchedCellKeys.Clear()
        $searchMatchedHeaderIndexes.Clear()

        foreach ($column in $grid.Columns) {
            $column.HeaderCell.Style.BackColor = $greenHeaderColor
            $column.HeaderCell.Style.SelectionBackColor = $greenHeaderColor
        }

        foreach ($row in $grid.Rows) {
            foreach ($cell in $row.Cells) {
                $cell.Style.BackColor = [System.Drawing.Color]::Empty
                $cell.Style.SelectionBackColor = [System.Drawing.Color]::Empty
            }
        }

        $searchResultLabel.Text = 'No search active.'
        $grid.Invalidate()
    }

    $applySearch = {
        & $clearSearchHighlights

        $query = $searchBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($query)) {
            return
        }

        $searchValue = ConvertTo-SearchDecimal -Text $query
        if ($null -eq $searchValue) {
            $searchResultLabel.Text = 'Enter numbers to search.'
            return
        }

        $cellHits = 0

        foreach ($column in $grid.Columns) {
            if (-not $column.Visible) {
                continue
            }

            if (Test-NumericSearchMatch -Text $column.HeaderText -SearchValue $searchValue) {
                $column.HeaderCell.Style.BackColor = $highlightColor
                $column.HeaderCell.Style.SelectionBackColor = $highlightColor
                $searchMatchedHeaderIndexes[$column.Index] = $true
            }
        }

        foreach ($row in $grid.Rows) {
            foreach ($cell in $row.Cells) {
                if (-not $cell.OwningColumn.Visible) {
                    continue
                }

                $valueText = ''
                if ($null -ne $cell.Value) {
                    $valueText = [string]$cell.Value
                }

                if (Test-NumericSearchMatch -Text $valueText -SearchValue $searchValue) {
                    $cellKey = '{0}:{1}' -f $cell.RowIndex, $cell.ColumnIndex
                    $searchMatchedCellKeys[$cellKey] = $true
                    $cell.Style.BackColor = $highlightColor
                    $cell.Style.SelectionBackColor = $highlightColor
                    $cell.OwningColumn.HeaderCell.Style.BackColor = $highlightColor
                    $cell.OwningColumn.HeaderCell.Style.SelectionBackColor = $highlightColor
                    $searchMatchedHeaderIndexes[$cell.ColumnIndex] = $true
                    $cellHits++
                }
            }
        }

        $headerHits = $searchMatchedHeaderIndexes.Count
        $totalHits = $cellHits + $headerHits
        if ($totalHits -eq 0) {
            $searchResultLabel.Text = "0 hits for '$query'."
        }
        else {
            $searchResultLabel.Text = "{0} hit(s) for '{1}' ({2} cell, {3} header)." -f $totalHits, $query, $cellHits, $headerHits
        }

        $grid.Invalidate()
    }

    $grid.Add_CellFormatting({
        param($sender, $eventArgs)

        if ($eventArgs.RowIndex -lt 0 -or $eventArgs.ColumnIndex -lt 0) {
            return
        }

        $cellKey = '{0}:{1}' -f $eventArgs.RowIndex, $eventArgs.ColumnIndex
        if ($searchMatchedCellKeys.ContainsKey($cellKey)) {
            $eventArgs.CellStyle.BackColor = $highlightColor
            $eventArgs.CellStyle.SelectionBackColor = $highlightColor
            $eventArgs.CellStyle.SelectionForeColor = [System.Drawing.Color]::Black
        }
        else {
            $eventArgs.CellStyle.BackColor = $defaultCellColor
        }
    })

    $openButton.Add_Click({
        try {
            if ($openDialog.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) {
                return
            }

            $path = $openDialog.FileName
            $openButton.Enabled = $false
            $grid.DataSource = $null
            $fileLabel.Text = $path

            & $setStatus 'Loading workbook...'
            $importResult = Import-XlsTransactions -Path $path -StatusCallback $setStatus
            $table = ConvertTo-TransactionDataTable -ImportResult $importResult
            $grid.DataSource = $table
            $grid.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::DisplayedCells)

            foreach ($column in $grid.Columns) {
                $column.Resizable = [System.Windows.Forms.DataGridViewTriState]::True
                $column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
                $column.HeaderCell.Style.BackColor = $greenHeaderColor
                $column.HeaderCell.Style.SelectionBackColor = $greenHeaderColor
            }

            foreach ($row in $grid.Rows) {
                $row.Frozen = $false
            }

            & $clearSearchHighlights

            if ($table.Rows.Count -eq 0) {
                & $setStatus 'Finished. No transactions were found.'
            }
            else {
                & $setStatus ("Finished. {0} date row(s), {1} sheet column(s)." -f $table.Rows.Count, $importResult.Sheets.Count)
            }

            if ($importResult.Warnings.Count -gt 0) {
                [System.Windows.Forms.MessageBox]::Show(($importResult.Warnings -join [Environment]::NewLine), 'Workbook Warnings', 'OK', 'Warning') | Out-Null
            }
        }
        catch {
            & $setStatus 'Error while processing workbook.'
            Write-EasyXlsLog -Message $_.Exception.ToString()
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'EasyXLS Error', 'OK', 'Error') | Out-Null
        }
        finally {
            $openButton.Enabled = $true
        }
    })

    $searchButton.Add_Click({
        & $applySearch
    })

    $clearSearchButton.Add_Click({
        $searchBox.Text = ''
        & $clearSearchHighlights
    })

    $searchBox.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $_.SuppressKeyPress = $true
            & $applySearch
        }
    })

    $exportPdfButton.Add_Click({
        try {
            if ($null -eq $grid.DataSource -or $grid.Rows.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show('Load a workbook before exporting to PDF.', 'EasyXLS', 'OK', 'Information') | Out-Null
                return
            }

            $defaultName = 'EasyXLS-Export.pdf'
            if (-not [string]::IsNullOrWhiteSpace($fileLabel.Text) -and (Test-Path -LiteralPath $fileLabel.Text -PathType Leaf)) {
                $defaultName = '{0}-EasyXLS.pdf' -f [System.IO.Path]::GetFileNameWithoutExtension($fileLabel.Text)
                $savePdfDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($fileLabel.Text)
            }

            $savePdfDialog.FileName = $defaultName
            if ($savePdfDialog.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) {
                return
            }

            & $setStatus 'Exporting PDF...'
            Export-GridToPdf -Grid $grid -Path $savePdfDialog.FileName -HeaderColor $greenHeaderColor -HighlightColor $highlightColor
            & $setStatus ("Exported PDF: {0}" -f $savePdfDialog.FileName)
            [System.Windows.Forms.MessageBox]::Show(("PDF exported successfully:`r`n{0}" -f $savePdfDialog.FileName), 'EasyXLS', 'OK', 'Information') | Out-Null
        }
        catch {
            & $setStatus 'Error while exporting PDF.'
            Write-EasyXlsLog -Message $_.Exception.ToString()
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'EasyXLS PDF Export Error', 'OK', 'Error') | Out-Null
        }
    })

    $getVisibleColumnCount = {
        $count = 0
        foreach ($column in $grid.Columns) {
            if ($column.Visible) {
                $count++
            }
        }

        return $count
    }

    $getHiddenColumnCount = {
        $count = 0
        foreach ($column in $grid.Columns) {
            if (-not $column.Visible) {
                $count++
            }
        }

        return $count
    }

    $grid.Add_CellMouseDown({
        param($sender, $eventArgs)

        if ($eventArgs.Button -ne [System.Windows.Forms.MouseButtons]::Right -or $eventArgs.RowIndex -ne -1 -or $eventArgs.ColumnIndex -lt 0) {
            return
        }

        $columnHeaderMenu.Tag = $eventArgs.ColumnIndex
        $column = $grid.Columns[$eventArgs.ColumnIndex]
        $hideColumnMenuItem.Text = "Hide '$($column.HeaderText)'"
        $hideColumnMenuItem.Enabled = ((& $getVisibleColumnCount) -gt 1)
        $unhideAllMenuItem.Enabled = ((& $getHiddenColumnCount) -gt 0)
        $columnHeaderMenu.Show($grid, $eventArgs.Location)
    })

    $hideColumnMenuItem.Add_Click({
        $clickedHeaderColumnIndex = [int]$columnHeaderMenu.Tag
        if ($clickedHeaderColumnIndex -lt 0 -or $clickedHeaderColumnIndex -ge $grid.Columns.Count) {
            return
        }

        if ((& $getVisibleColumnCount) -le 1) {
            [System.Windows.Forms.MessageBox]::Show('At least one column must remain visible.', 'EasyXLS', 'OK', 'Information') | Out-Null
            return
        }

        $grid.Columns[$clickedHeaderColumnIndex].Visible = $false
        if ([string]::IsNullOrWhiteSpace($searchBox.Text)) {
            & $clearSearchHighlights
        }
        else {
            & $applySearch
        }
    })

    $unhideAllMenuItem.Add_Click({
        foreach ($column in $grid.Columns) {
            $column.Visible = $true
        }

        if ([string]::IsNullOrWhiteSpace($searchBox.Text)) {
            & $clearSearchHighlights
        }
        else {
            & $applySearch
        }
    })

    [void]$topPanel.Controls.Add($openButton)
    [void]$topPanel.Controls.Add($fileLabel)
    [void]$topPanel.Controls.Add($searchLabel)
    [void]$topPanel.Controls.Add($searchBox)
    [void]$topPanel.Controls.Add($searchButton)
    [void]$topPanel.Controls.Add($clearSearchButton)
    [void]$topPanel.Controls.Add($exportPdfButton)
    [void]$topPanel.Controls.Add($searchResultLabel)
    [void]$form.Controls.Add($grid)
    [void]$form.Controls.Add($topPanel)
    [void]$form.Controls.Add($statusStrip)

    [void]$form.ShowDialog()
}

Show-EasyXlsGui
