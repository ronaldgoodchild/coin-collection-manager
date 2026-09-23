#Requires -Version 5.1
<#
.SYNOPSIS
    US Coin Collection Dashboard - WPF GUI Application
.DESCRIPTION
    Full mouse-driven Windows GUI to track your coin collection.
    Dashboard-style layout matching the US Coin Collection spreadsheet.
    Supports multiple collections, slab/certification tracking,
    coin images, backup/restore, duplicate detection, and
    Excel import from US Coin Collection spreadsheets.
    Run from PowerShell: .\CoinCollection.ps1
#>

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Web

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# ImagePath -> BitmapImage converter (C# for WPF binding)
# Cache compiled DLL for fast startup on subsequent launches
# ---------------------------------------------------------------------------
$_converterDll = Join-Path $PSScriptRoot 'ImagePathConverter.dll'
if ((Test-Path $_converterDll) -and -not ([System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetType('ImagePathConverter', $false) })) {
    try {
        Add-Type -Path $_converterDll -ErrorAction Stop
    } catch {
        Remove-Item $_converterDll -Force -ErrorAction SilentlyContinue
    }
}
if (-not ([System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetType('ImagePathConverter', $false) })) {
    $_presCore = [System.Windows.Media.Imaging.BitmapImage].Assembly.Location
    $_winBase  = [System.Windows.Threading.Dispatcher].Assembly.Location
    $_presFwk  = [System.Windows.Data.IValueConverter].Assembly.Location
    $_sysXaml  = Join-Path ([System.Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()) 'System.Xaml.dll'
    $_compilerParams = [System.CodeDom.Compiler.CompilerParameters]::new()
    $_sysAssembly = [System.Uri].Assembly.Location
    $_compilerParams.ReferencedAssemblies.AddRange(@($_presCore, $_winBase, $_presFwk, $_sysXaml, $_sysAssembly))
    $_compilerParams.OutputAssembly = $_converterDll
    $_compilerParams.GenerateInMemory = $false
    $provider = [Microsoft.CSharp.CSharpCodeProvider]::new()
    $result = $provider.CompileAssemblyFromSource($_compilerParams, @'
using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media.Imaging;

public class ImagePathConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        string path = value as string;
        if (string.IsNullOrWhiteSpace(path) || !System.IO.File.Exists(path))
            return null;
        try
        {
            var bi = new BitmapImage();
            bi.BeginInit();
            bi.UriSource = new Uri(path, UriKind.Absolute);
            bi.DecodePixelWidth = 50;
            bi.CacheOption = BitmapCacheOption.OnLoad;
            bi.EndInit();
            bi.Freeze();
            return bi;
        }
        catch { return null; }
    }
    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}
'@)
    if ($result.Errors.HasErrors) {
        # Fallback: compile in-memory the old way
        Add-Type -TypeDefinition @'
using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media.Imaging;

public class ImagePathConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        string path = value as string;
        if (string.IsNullOrWhiteSpace(path) || !System.IO.File.Exists(path))
            return null;
        try
        {
            var bi = new BitmapImage();
            bi.BeginInit();
            bi.UriSource = new Uri(path, UriKind.Absolute);
            bi.DecodePixelWidth = 50;
            bi.CacheOption = BitmapCacheOption.OnLoad;
            bi.EndInit();
            bi.Freeze();
            return bi;
        }
        catch { return null; }
    }
    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}
'@ -ReferencedAssemblies $_presCore, $_winBase, $_presFwk, $_sysXaml
    }
    $provider.Dispose()
}

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
$script:Version         = '4.1.0'
$script:CollectionsDir  = Join-Path $PSScriptRoot 'Collections'
$script:ActiveCollection = 'MyCoins'
$script:DataFile        = Join-Path $script:CollectionsDir 'MyCoins.csv'
$script:CsvHeader       = 'ID,Year,Mint,Denomination,Type,Variety,Country,Grade,FaceValue,BookValue,PurchasePrice,CurrentValue,KeyDates,Error,Metal,Weight,DupCount,DupFaceValue,DupGradedValue,Location,DupLoc,Comments,Notes,DateRangeMinted,DataStandard,InCollection,DateAdded,CertCompany,CertNumber,SlabGrade,Mintage,RarityScore,ImagePath,Series,VAMNumber,FSNumber,ErrorType,VarietyType'
$script:ValueHistoryFile = Join-Path $PSScriptRoot 'ValueHistory.csv'
$script:CoinImageDir    = Join-Path $PSScriptRoot 'CoinImages'

# Denomination -> coin type image file mapping (27 types from workbook)
$script:CoinTypeImageMap = @{
    'Half Cent'                   = 'image1.png'
    'Large Cent'                  = 'image2.png'
    'Flying Eagle Cent'           = 'image3.png'
    'Indian Head Cent'            = 'image3.png'
    'Lincoln Cent'                = 'image3.png'
    'Two Cent Piece'              = 'image4.png'
    'Three Cent Piece'            = 'image5.png'
    'Buffalo Nickel'              = 'image6.png'
    'Jefferson Nickel'            = 'image6.png'
    'Liberty Nickel'              = 'image6.png'
    'Shield Nickel'               = 'image6.png'
    'Early Half Dime'             = 'image7.png'
    'Liberty Seated Half Dime'    = 'image7.png'
    'Barber Dime'                 = 'image8.png'
    'Bust Dime'                   = 'image8.png'
    'Early Dime'                  = 'image8.png'
    'Mercury Dime'                = 'image8.png'
    'Roosevelt Dime'              = 'image8.png'
    'Seated Liberty Dime'         = 'image8.png'
    'Twenty Cent Piece'           = 'image9.png'
    'Barber Quarter'              = 'image10.png'
    'Capped Bust Quarter'         = 'image10.png'
    'Draped Bust Quarters'        = 'image10.png'
    'Seated Liberty Quarter'      = 'image10.png'
    'Standing Liberty Quarter'    = 'image10.png'
    'Washington Quarter'          = 'image10.png'
    'Capped Bust Half Dollar'     = 'image14.png'
    'Draped Bust Half Dollar'     = 'image14.png'
    'Flowing Hair Half Dollar'    = 'image14.png'
    'Seated Liberty Half Dollar'  = 'image14.png'
    'Barber Half Dollar'          = 'image15.png'
    'Franklin Half Dollar'        = 'image15.png'
    'Kennedy Half Dollar'         = 'image15.png'
    'Walking Liberty Half Dollar' = 'image15.png'
    'Draped Bust Dollar'          = 'image16.png'
    'Flowing Hair Dollar'         = 'image16.png'
    'Seated Liberty Dollar'       = 'image17.png'
    'Trade Dollar'                = 'image18.png'
    'Morgan Dollar'               = 'image19.png'
    'Peace Dollar'                = 'image19.png'
    'Eisenhower Dollar'           = 'image20.png'
    'Susan B. Anthony Dollar'     = 'image21.png'
    'Sacagawea Dollar'            = 'image22.png'
    'Presidential Dollar'         = 'image23.png'
    'American Eagle Dollar'       = 'image24.png'
    'Gold ($1.00)'                = 'image25.png'
    'Gold ($2.50)'                = 'image25.png'
    'Gold ($3.00)'                = 'image25.png'
    'Gold Half Eagle ($5.00)'     = 'image25.png'
    'Gold Eagle ($10.00)'         = 'image25.png'
    'Gold Quarter Eagle ($2.50)'  = 'image26.png'
    'Gold Double Eagle ($20.00)'  = 'image26.png'
    'Commemorative'               = 'image27.png'
    'Commemorative Half Dollar'   = 'image27.png'
    'Mint Set'                    = 'image27.png'
    'Proof Set'                   = 'image27.png'
}

# ---------------------------------------------------------------------------
# CSV Data Layer (Hardened + Self-Healing)
# ---------------------------------------------------------------------------
# Pre-split header into array for fast iteration
$script:CsvFields = $script:CsvHeader -split ','
# Default values for blank fields (cached, not re-evaluated per row)
$script:FieldDefaults = @{
    PurchasePrice = '0'; CurrentValue = '0'; BookValue = '0'; FaceValue = '0'
    DupCount = '0'; DupFaceValue = '0'; DupGradedValue = '0'; InCollection = 'Y'
}
# Fields requiring numeric validation
$script:NumericFields = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@('PurchasePrice','CurrentValue','BookValue','FaceValue','DupFaceValue','DupGradedValue'),
    [System.StringComparer]::Ordinal)

function Repair-CsvRow {
    param([pscustomobject]$Row)
    # Build HashSet of available properties ONCE per row (instead of 38 -contains checks)
    $propSet = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]@($Row.PSObject.Properties.Name),
        [System.StringComparer]::OrdinalIgnoreCase)

    $fixed = [ordered]@{}
    foreach ($fld in $script:CsvFields) {
        if ($propSet.Contains($fld)) {
            $v = $Row.$fld
            $fixed[$fld] = if ($null -ne $v) { [string]$v } else { '' }
        } else {
            $fixed[$fld] = ''
        }
    }
    # Apply defaults for blank fields
    foreach ($k in @($fixed.Keys)) {
        if ([string]::IsNullOrWhiteSpace($fixed[$k])) {
            if ($script:FieldDefaults.ContainsKey($k)) {
                $fixed[$k] = $script:FieldDefaults[$k]
            } elseif ($k -eq 'DateAdded') {
                $fixed[$k] = (Get-Date -Format 'yyyy-MM-dd')
            }
        }
    }
    # Numeric validation
    foreach ($n in $script:NumericFields) {
        if ($fixed[$n] -notmatch '^\d+(\.\d{1,4})?$') { $fixed[$n] = '0' }
    }
    if ($fixed['DupCount'] -notmatch '^\d+$') { $fixed['DupCount'] = '0' }
    return [pscustomobject]$fixed
}

function Get-Collection {
    if (-not (Test-Path $script:DataFile)) {
        $script:CsvHeader | Set-Content -Path $script:DataFile -Encoding UTF8
        return @()
    }
    $raw = Import-Csv -Path $script:DataFile -Encoding UTF8
    if ($null -eq $raw) { return @() }
    # Use List<T> to avoid slow array += reallocation on every row
    $clean = [System.Collections.Generic.List[pscustomobject]]::new(1500)
    foreach ($row in $raw) {
        if ($null -eq $row) { continue }
        $clean.Add((Repair-CsvRow $row))
    }
    return @($clean)
}

function Save-Collection {
    param([object[]]$Coins)
    if ($null -eq $Coins -or $Coins.Count -eq 0) {
        $script:CsvHeader | Set-Content -Path $script:DataFile -Encoding UTF8
        return
    }
    $fixed = $Coins | ForEach-Object { Repair-CsvRow $_ }
    $fixed | Export-Csv -Path $script:DataFile -NoTypeInformation -Encoding UTF8
}

function New-CoinID {
    param([object[]]$Coins)
    if ($null -eq $Coins -or $Coins.Count -eq 0) { return 1 }
    return (($Coins | ForEach-Object { [int]$_.ID } | Measure-Object -Maximum).Maximum + 1)
}

# ---------------------------------------------------------------------------
# Multi-Collection Helpers
# ---------------------------------------------------------------------------
function Initialize-Collections {
    if (-not (Test-Path $script:CollectionsDir)) {
        New-Item -Path $script:CollectionsDir -ItemType Directory -Force | Out-Null
    }
    $legacyFile  = Join-Path $PSScriptRoot 'CoinCollection.csv'
    $defaultFile = Join-Path $script:CollectionsDir 'MyCoins.csv'
    if ((Test-Path $legacyFile) -and -not (Test-Path $defaultFile)) {
        Copy-Item -Path $legacyFile -Destination $defaultFile -Force
    }
    if (-not (Test-Path $defaultFile)) {
        $script:CsvHeader | Set-Content -Path $defaultFile -Encoding UTF8
    }
    $script:DataFile         = $defaultFile
    $script:ActiveCollection = 'MyCoins'
}

function Get-CollectionNames {
    if (-not (Test-Path $script:CollectionsDir)) { return @('MyCoins') }
    $files = Get-ChildItem -Path $script:CollectionsDir -Filter '*.csv' | Sort-Object Name
    return @($files | ForEach-Object { [System.IO.Path]::GetFileNameWithoutExtension($_.Name) })
}

# ---------------------------------------------------------------------------
# Backup / Restore
# ---------------------------------------------------------------------------
function Invoke-Backup {
    param([switch]$Force)
    $backupDir = Join-Path $PSScriptRoot 'Backups'
    if (-not (Test-Path $backupDir)) {
        New-Item -Path $backupDir -ItemType Directory -Force | Out-Null
    }
    # Skip if a backup from today already exists (unless -Force)
    if (-not $Force) {
        $todayStamp = Get-Date -Format 'yyyyMMdd'
        $existing = @(Get-ChildItem -Path $backupDir -Directory -Filter "Collections_${todayStamp}_*" -ErrorAction SilentlyContinue)
        if ($existing.Count -gt 0) { return $existing[0].FullName }
    }
    $stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
    $destDir = Join-Path $backupDir "Collections_$stamp"
    if (Test-Path $script:CollectionsDir) {
        Copy-Item -Path $script:CollectionsDir -Destination $destDir -Recurse -Force
    }
    $allBackups = @(Get-ChildItem -Path $backupDir -Directory -Filter 'Collections_*' |
        Sort-Object LastWriteTime -Descending)
    if ($allBackups.Count -gt 10) {
        $allBackups | Select-Object -Skip 10 | Remove-Item -Recurse -Force
    }
    return $destDir
}

# ---------------------------------------------------------------------------
# Import from Excel (.xlsm / .xlsx)
# ---------------------------------------------------------------------------
function Import-FromExcel {
    param(
        [string]$ExcelPath,
        [bool]$OnlyInCollection = $true
    )

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    try {
        $wb = $excel.Workbooks.Open($ExcelPath)
        $ws = $wb.Worksheets.Item('Database')
        $data = $ws.UsedRange.Value2
        $rowMax = $data.GetUpperBound(0)
        $colMax = [Math]::Min($data.GetUpperBound(1), 23)

        $cInColl    = 1;  $cCoinName  = 2;  $cType      = 3;  $cYear      = 4
        $cMintMark  = 5;  $cError     = 6;  $cVariety   = 7;  $cGrade     = 8
        $cBookVal   = 9;  $cKeyDates  = 10; $cFaceVal   = 11; $cDupCount  = 12
        $cDupFace   = 13; $cMetal     = 14; $cDupGrade  = 15; $cLocation  = 16
        $cDupLoc    = 17; $cComments  = 18; $cNotes     = 19; $cDateRange = 20
        $cDataStd   = 21; $cWeight    = 22; $cID        = 23

        function GetCell([int]$r, [int]$c) {
            if ($c -lt 1 -or $c -gt $colMax) { return '' }
            $v = $data[$r, $c]
            if ($null -eq $v) { return '' }
            return ([string]$v).Trim()
        }
        function CleanNum([string]$val) {
            if ([string]::IsNullOrWhiteSpace($val)) { return '0' }
            $c = $val -replace '[\$\s,]', ''
            if ($c -match '^\d+(\.\d+)?$') { return $c }
            return '0'
        }

        $coins = [System.Collections.Generic.List[PSCustomObject]]::new()

        for ($r = 2; $r -le $rowMax; $r++) {
            $coinName = GetCell $r $cCoinName
            if ([string]::IsNullOrWhiteSpace($coinName)) { continue }

            $inColl = GetCell $r $cInColl
            if ($OnlyInCollection -and $inColl -ne 'Y') { continue }

            $rawID = GetCell $r $cID
            $id = if ($rawID -match '^\d') { [string][int][double]$rawID } else { '' }
            $rawYear = GetCell $r $cYear
            $year = if ($rawYear -match '^\d') { [string][int][double]$rawYear } else { $rawYear }
            $rawDup = GetCell $r $cDupCount
            $dupCnt = if ($rawDup -match '^\d') { [string][int][double]$rawDup } else { '0' }

            $coin = [PSCustomObject]@{
                ID              = $id
                Year            = $year
                Mint            = GetCell $r $cMintMark
                Denomination    = $coinName
                Type            = GetCell $r $cType
                Variety         = GetCell $r $cVariety
                Country         = 'USA'
                Grade           = GetCell $r $cGrade
                FaceValue       = CleanNum (GetCell $r $cFaceVal)
                BookValue       = CleanNum (GetCell $r $cBookVal)
                PurchasePrice   = '0'
                CurrentValue    = CleanNum (GetCell $r $cBookVal)
                KeyDates        = GetCell $r $cKeyDates
                Error           = GetCell $r $cError
                Metal           = GetCell $r $cMetal
                Weight          = GetCell $r $cWeight
                DupCount        = $dupCnt
                DupFaceValue    = CleanNum (GetCell $r $cDupFace)
                DupGradedValue  = CleanNum (GetCell $r $cDupGrade)
                Location        = GetCell $r $cLocation
                DupLoc          = GetCell $r $cDupLoc
                Comments        = GetCell $r $cComments
                Notes           = GetCell $r $cNotes
                DateRangeMinted = GetCell $r $cDateRange
                DataStandard    = GetCell $r $cDataStd
                InCollection    = $inColl
                DateAdded       = Get-Date -Format 'yyyy-MM-dd'
                CertCompany     = ''
                CertNumber      = ''
                SlabGrade       = ''
                Mintage         = ''
                RarityScore     = ''
                ImagePath       = ''
                Series          = ''
                VAMNumber       = ''
                FSNumber        = ''
                ErrorType       = ''
                VarietyType     = ''
            }

            $coins.Add($coin)
        }

        $wb.Close($false)
        return $coins.ToArray()
    }
    catch {
        throw
    }
    finally {
        try { $excel.Quit() } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { }
    }
}

# ---------------------------------------------------------------------------
# Coin Type Image Helpers
# ---------------------------------------------------------------------------
function Extract-CoinImages {
    param([string]$ExcelPath)
    if (-not (Test-Path $script:CoinImageDir)) {
        New-Item -Path $script:CoinImageDir -ItemType Directory -Force | Out-Null
    }
    $zipCopy = Join-Path $env:TEMP 'coin_workbook_extract.zip'
    $tempDir = Join-Path $env:TEMP 'coin_excel_extract'
    try {
        if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }
        Copy-Item $ExcelPath $zipCopy -Force
        Expand-Archive $zipCopy $tempDir -Force
        $mediaDir = Join-Path $tempDir 'xl\media'
        if (Test-Path $mediaDir) {
            $imgs = Get-ChildItem $mediaDir -Filter '*.png'
            foreach ($img in $imgs) {
                Copy-Item $img.FullName (Join-Path $script:CoinImageDir $img.Name) -Force
            }
            return $imgs.Count
        }
        return 0
    } catch { return 0 }
    finally {
        if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
        if (Test-Path $zipCopy) { Remove-Item $zipCopy -Force -ErrorAction SilentlyContinue }
    }
}

function Get-CoinTypeImage {
    param([string]$Denomination)
    if ([string]::IsNullOrWhiteSpace($Denomination)) { return '' }
    $imgFile = $script:CoinTypeImageMap[$Denomination]
    if ($imgFile) {
        $fullPath = Join-Path $script:CoinImageDir $imgFile
        if (Test-Path $fullPath) { return $fullPath }
    }
    return ''
}

# ---------------------------------------------------------------------------
# Image Viewer (click to enlarge)
# ---------------------------------------------------------------------------
function Show-ImageViewer {
    param([string]$ImagePath, [string]$Title = 'Coin Image')
    if ([string]::IsNullOrWhiteSpace($ImagePath) -or -not (Test-Path $ImagePath)) { return }
    [xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="PLACEHOLDER" Height="620" Width="620"
        WindowStartupLocation="CenterOwner" ResizeMode="CanResize"
        Background="#1A1A2E">
    <Grid>
        <Image Name="imgFull" Stretch="Uniform" Margin="10"/>
    </Grid>
</Window>
'@
    $xaml.Window.Title = $Title
    $rdr = [System.Xml.XmlNodeReader]::new($xaml)
    $win = [System.Windows.Markup.XamlReader]::Load($rdr)
    try {
        $bi = [System.Windows.Media.Imaging.BitmapImage]::new()
        $bi.BeginInit()
        $bi.UriSource = [Uri]::new($ImagePath, [UriKind]::Absolute)
        $bi.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bi.EndInit(); $bi.Freeze()
        $win.FindName('imgFull').Source = $bi
    } catch { return }
    $prevEAP = $ErrorActionPreference
    try { $ErrorActionPreference = 'Continue'; $null = $win.ShowDialog() }
    finally { $ErrorActionPreference = $prevEAP }
}

# ---------------------------------------------------------------------------
# Stat Card Export (click card -> printable text file)
# ---------------------------------------------------------------------------
function Export-StatCard {
    param([string]$Category, [object[]]$AllCoins)
    if ($null -eq $AllCoins -or $AllCoins.Count -eq 0) { return }
    $lines = [System.Collections.Generic.List[string]]::new()
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $subset = @()
    switch ($Category) {
        'Holding' {
            $subset = @($AllCoins | Where-Object { $_.InCollection -eq 'Y' })
            $lines.Add("COINS IN COLLECTION (Holding)  -  $stamp")
        }
        'Missing' {
            $subset = @($AllCoins | Where-Object { $_.InCollection -ne 'Y' })
            $lines.Add("COINS NOT IN COLLECTION (Missing)  -  $stamp")
        }
        'Minted' {
            $subset = @($AllCoins)
            $lines.Add("ALL COINS (Minted Total)  -  $stamp")
        }
        'HoldPct' {
            $hold = @($AllCoins | Where-Object { $_.InCollection -eq 'Y' }).Count
            $total = $AllCoins.Count
            $pct = if ($total -gt 0) { [Math]::Round(($hold / $total) * 100, 1) } else { 0 }
            $lines.Add("HOLD PERCENTAGE REPORT  -  $stamp")
            $lines.Add(('=' * 80))
            $lines.Add(("Holding: {0}   Missing: {1}   Total: {2}   Pct: {3}%" -f $hold, ($total - $hold), $total, $pct))
            $lines.Add('')
            $byDenom = $AllCoins | Group-Object Denomination | Sort-Object Count -Descending
            $lines.Add(("{0,-35} {1,8} {2,8} {3,8}" -f 'Denomination', 'Have', 'Total', 'Pct%'))
            $lines.Add(('-' * 65))
            foreach ($g in $byDenom) {
                $h = @($g.Group | Where-Object { $_.InCollection -eq 'Y' }).Count
                $t = $g.Count
                $p = if ($t -gt 0) { [Math]::Round(($h/$t)*100,1) } else { 0 }
                $lines.Add(("{0,-35} {1,8} {2,8} {3,7}%" -f $g.Name, $h, $t, $p))
            }
            $subset = $null
        }
        'BookValue' {
            $lines.Add("BOOK VALUE REPORT  -  $stamp")
            $lines.Add(('=' * 100))
            $total = [Math]::Round(($AllCoins | ForEach-Object { [double]$_.BookValue } | Measure-Object -Sum).Sum, 2)
            $lines.Add(("Total Book Value: `${0:N2}" -f $total))
            $lines.Add('')
            $byDenom = $AllCoins | Group-Object Denomination | Sort-Object @{E={($_.Group | ForEach-Object {[double]$_.BookValue} | Measure-Object -Sum).Sum}} -Descending
            $lines.Add(("{0,-35} {1,10} {2,8}" -f 'Denomination', 'Book Value', 'Count'))
            $lines.Add(('-' * 58))
            foreach ($g in $byDenom) {
                $bv = [Math]::Round(($g.Group | ForEach-Object {[double]$_.BookValue} | Measure-Object -Sum).Sum, 2)
                $lines.Add(("{0,-35} `${1,9:N2} {2,8}" -f $g.Name, $bv, $g.Count))
            }
            $subset = $null
        }
        'FaceValue' {
            $lines.Add("FACE VALUE REPORT  -  $stamp")
            $lines.Add(('=' * 100))
            $total = [Math]::Round(($AllCoins | ForEach-Object { [double]$_.FaceValue } | Measure-Object -Sum).Sum, 2)
            $lines.Add(("Total Face Value: `${0:N2}" -f $total))
            $lines.Add('')
            $byDenom = $AllCoins | Group-Object Denomination | Sort-Object @{E={($_.Group | ForEach-Object {[double]$_.FaceValue} | Measure-Object -Sum).Sum}} -Descending
            $lines.Add(("{0,-35} {1,10} {2,8}" -f 'Denomination', 'Face Value', 'Count'))
            $lines.Add(('-' * 58))
            foreach ($g in $byDenom) {
                $fv = [Math]::Round(($g.Group | ForEach-Object {[double]$_.FaceValue} | Measure-Object -Sum).Sum, 2)
                $lines.Add(("{0,-35} `${1,9:N2} {2,8}" -f $g.Name, $fv, $g.Count))
            }
            $subset = $null
        }
        'Graded' {
            $subset = @($AllCoins | Where-Object { $_.Grade -ne '' -and $null -ne $_.Grade })
            $lines.Add("GRADED COINS  -  $stamp")
        }
        'KeyDates' {
            $subset = @($AllCoins | Where-Object { $_.KeyDates -ne '' -and $null -ne $_.KeyDates })
            $lines.Add("KEY DATE COINS  -  $stamp")
        }
        'Duplicates' {
            $subset = @($AllCoins | Where-Object { [int]$_.DupCount -gt 0 })
            $lines.Add("DUPLICATE COINS  -  $stamp")
        }
    }

    if ($null -ne $subset) {
        $lines.Add(('=' * 120))
        $lines.Add(("Total: {0} coins" -f $subset.Count))
        $lines.Add('')
        $lines.Add(("{0,-5} {1,-5} {2,-5} {3,-28} {4,-16} {5,-14} {6,8} {7,8} {8,-8} {9,-10}" -f
            'ID','Year','Mint','Denomination','Type','Grade','Face$','Book$','Key','Metal'))
        $lines.Add(('-' * 120))
        foreach ($c in ($subset | Sort-Object @{E={[int]$_.Year}}, Mint)) {
            $lines.Add(("{0,-5} {1,-5} {2,-5} {3,-28} {4,-16} {5,-14} {6,8:N2} {7,8:N2} {8,-8} {9,-10}" -f
                $c.ID, $c.Year, $c.Mint, $c.Denomination, $c.Type, $c.Grade,
                ([double]$c.FaceValue), ([double]$c.BookValue), $c.KeyDates, $c.Metal))
        }
    }
    $lines.Add('')
    $lines.Add(('=' * 120))

    $reportDir = Join-Path $PSScriptRoot 'Reports'
    if (-not (Test-Path $reportDir)) { New-Item $reportDir -ItemType Directory -Force | Out-Null }
    $safeCategory = $Category -replace '[^a-zA-Z0-9]', ''
    $fileName = "{0}_{1}.txt" -f $safeCategory, (Get-Date -Format 'yyyyMMdd_HHmmss')
    $filePath = Join-Path $reportDir $fileName
    $lines | Set-Content -Path $filePath -Encoding UTF8
    Start-Process notepad.exe $filePath
    return $filePath
}

# ---------------------------------------------------------------------------
# Online Research Links
# ---------------------------------------------------------------------------
function Show-ResearchMenu {
    param([hashtable]$Coin)
    $yr   = if ($Coin) { $Coin['Year'] } else { '' }
    $mint = if ($Coin) { $Coin['Mint'] } else { '' }
    $den  = if ($Coin) { $Coin['Denomination'] } else { '' }
    $cert = if ($Coin) { $Coin['CertNumber'] } else { '' }
    $certCo = if ($Coin) { $Coin['CertCompany'] } else { '' }
    $query  = [Uri]::EscapeDataString("$yr $mint $den coin")
    $denEnc = [Uri]::EscapeDataString($den)

    $links = [ordered]@{
        'PCGS CoinFacts (coin encyclopedia)'    = "https://www.pcgs.com/coinfacts/search?query=$query"
        'PCGS Price Guide'                       = "https://www.pcgs.com/prices/us"
        'PCGS Photograde (grading photos)'       = "https://www.pcgs.com/photograde"
        'PCGS Auction Prices Realized'           = "https://www.pcgs.com/auctionprices/search?query=$query"
        'PCGS Population Report'                 = "https://www.pcgs.com/pop/uscoins"
        '---1'                                   = ''
        'NGC Coin Explorer'                      = "https://www.ngccoin.com/coin-explorer/united-states/"
        'NGC Price Guide'                        = "https://www.ngccoin.com/price-guide/united-states/"
        '---2'                                   = ''
        'Heritage Auctions (auction results)'    = "https://coins.ha.com/c/search.zx?N=0&Ntt=$query"
        'USA CoinBook (free price guide)'        = "https://www.usacoinbook.com/"
        'CoinTrackers (quick valuations)'        = "https://cointrackers.com/"
        'Greysheet (wholesale pricing)'          = "https://www.greysheet.com/"
        '---3'                                   = ''
        'USA CoinBook Melt Values'               = "https://www.usacoinbook.com/coin-melt-values/"
        'MeltValue.com (spot price calc)'        = "https://meltvalue.com/"
        '---4'                                   = ''
        'CONECA (error coin reference)'          = "https://conecaonline.org/"
        'Numista (world coin database)'          = "https://en.numista.com/catalogue/index.php?r=$denEnc&ct=coin"
        'ANA (education/resources)'              = "https://www.money.org/"
    }
    if (-not [string]::IsNullOrWhiteSpace($cert)) {
        if ($certCo -eq 'PCGS') {
            $links['---5'] = ''
            $links["PCGS Cert Verify: $cert"] = "https://www.pcgs.com/cert/$cert"
        } elseif ($certCo -eq 'NGC') {
            $links['---5'] = ''
            $links["NGC Cert Verify: $cert"] = "https://www.ngccoin.com/certlookup/$cert"
        }
    }

    $ctxResearch = [System.Windows.Controls.ContextMenu]::new()
    $ctxResearch.FontSize = 12
    foreach ($kv in $links.GetEnumerator()) {
        if ($kv.Key -like '---*') {
            $ctxResearch.Items.Add([System.Windows.Controls.Separator]::new()) | Out-Null
        } else {
            $mi = [System.Windows.Controls.MenuItem]::new()
            $mi.Header = $kv.Key
            $mi.Tag    = $kv.Value
            $mi.Add_Click({ param($s,$e); Start-Process $s.Tag })
            $ctxResearch.Items.Add($mi) | Out-Null
        }
    }
    $ctxResearch.IsOpen = $true
    $ctxResearch.PlacementTarget = $btnResearch
}

# ---------------------------------------------------------------------------
# Melt Value Calculator
# ---------------------------------------------------------------------------
# Known compositions: metal content in troy ounces for common US coin types
$script:MeltCompositions = @{
    'Half Cent'                   = @{ Silver = 0; Gold = 0; Copper = 0.00480 }
    'Large Cent'                  = @{ Silver = 0; Gold = 0; Copper = 0.00989 }
    'Flying Eagle Cent'           = @{ Silver = 0; Gold = 0; Copper = 0.00252 }
    'Indian Head Cent'            = @{ Silver = 0; Gold = 0; Copper = 0.00222 }
    'Lincoln Cent'                = @{ Silver = 0; Gold = 0; Copper = 0.00222 }
    'Two Cent Piece'              = @{ Silver = 0; Gold = 0; Copper = 0.00444 }
    'Three Cent Piece'            = @{ Silver = 0.02172; Gold = 0; Copper = 0 }
    'Shield Nickel'               = @{ Silver = 0; Gold = 0; Copper = 0.00360 }
    'Liberty Nickel'              = @{ Silver = 0; Gold = 0; Copper = 0.00360 }
    'Buffalo Nickel'              = @{ Silver = 0; Gold = 0; Copper = 0.00360 }
    'Jefferson Nickel'            = @{ Silver = 0; Gold = 0; Copper = 0.00360 }
    'Early Half Dime'             = @{ Silver = 0.03876; Gold = 0; Copper = 0 }
    'Liberty Seated Half Dime'    = @{ Silver = 0.03876; Gold = 0; Copper = 0 }
    'Barber Dime'                 = @{ Silver = 0.07234; Gold = 0; Copper = 0 }
    'Bust Dime'                   = @{ Silver = 0.07234; Gold = 0; Copper = 0 }
    'Early Dime'                  = @{ Silver = 0.07234; Gold = 0; Copper = 0 }
    'Mercury Dime'                = @{ Silver = 0.07234; Gold = 0; Copper = 0 }
    'Roosevelt Dime'              = @{ Silver = 0.07234; Gold = 0; Copper = 0 }
    'Seated Liberty Dime'         = @{ Silver = 0.07234; Gold = 0; Copper = 0 }
    'Twenty Cent Piece'           = @{ Silver = 0.14468; Gold = 0; Copper = 0 }
    'Barber Quarter'              = @{ Silver = 0.18084; Gold = 0; Copper = 0 }
    'Capped Bust Quarter'         = @{ Silver = 0.18084; Gold = 0; Copper = 0 }
    'Draped Bust Quarters'        = @{ Silver = 0.18084; Gold = 0; Copper = 0 }
    'Seated Liberty Quarter'      = @{ Silver = 0.18084; Gold = 0; Copper = 0 }
    'Standing Liberty Quarter'    = @{ Silver = 0.18084; Gold = 0; Copper = 0 }
    'Washington Quarter'          = @{ Silver = 0.18084; Gold = 0; Copper = 0 }
    'Capped Bust Half Dollar'     = @{ Silver = 0.36169; Gold = 0; Copper = 0 }
    'Draped Bust Half Dollar'     = @{ Silver = 0.36169; Gold = 0; Copper = 0 }
    'Flowing Hair Half Dollar'    = @{ Silver = 0.36169; Gold = 0; Copper = 0 }
    'Seated Liberty Half Dollar'  = @{ Silver = 0.36169; Gold = 0; Copper = 0 }
    'Barber Half Dollar'          = @{ Silver = 0.36169; Gold = 0; Copper = 0 }
    'Franklin Half Dollar'        = @{ Silver = 0.36169; Gold = 0; Copper = 0 }
    'Kennedy Half Dollar'         = @{ Silver = 0.14792; Gold = 0; Copper = 0 }
    'Walking Liberty Half Dollar' = @{ Silver = 0.36169; Gold = 0; Copper = 0 }
    'Draped Bust Dollar'          = @{ Silver = 0.77344; Gold = 0; Copper = 0 }
    'Flowing Hair Dollar'         = @{ Silver = 0.77344; Gold = 0; Copper = 0 }
    'Seated Liberty Dollar'       = @{ Silver = 0.77344; Gold = 0; Copper = 0 }
    'Trade Dollar'                = @{ Silver = 0.78750; Gold = 0; Copper = 0 }
    'Morgan Dollar'               = @{ Silver = 0.77344; Gold = 0; Copper = 0 }
    'Peace Dollar'                = @{ Silver = 0.77344; Gold = 0; Copper = 0 }
    'Eisenhower Dollar'           = @{ Silver = 0; Gold = 0; Copper = 0.00771 }
    'Susan B. Anthony Dollar'     = @{ Silver = 0; Gold = 0; Copper = 0.00571 }
    'Sacagawea Dollar'            = @{ Silver = 0; Gold = 0; Copper = 0.00571 }
    'Presidential Dollar'         = @{ Silver = 0; Gold = 0; Copper = 0.00571 }
    'American Eagle Dollar'       = @{ Silver = 1.00000; Gold = 0; Copper = 0 }
    'Gold ($1.00)'                = @{ Silver = 0; Gold = 0.04837; Copper = 0 }
    'Gold ($2.50)'                = @{ Silver = 0; Gold = 0.12094; Copper = 0 }
    'Gold ($3.00)'                = @{ Silver = 0; Gold = 0.14512; Copper = 0 }
    'Gold Half Eagle ($5.00)'     = @{ Silver = 0; Gold = 0.24187; Copper = 0 }
    'Gold Eagle ($10.00)'         = @{ Silver = 0; Gold = 0.48375; Copper = 0 }
    'Gold Quarter Eagle ($2.50)'  = @{ Silver = 0; Gold = 0.12094; Copper = 0 }
    'Gold Double Eagle ($20.00)'  = @{ Silver = 0; Gold = 0.96750; Copper = 0 }
}

function Show-MeltCalculator {
    param([object[]]$Coins)
    if ($null -eq $Coins -or $Coins.Count -eq 0) {
        [System.Windows.MessageBox]::Show('No coins in collection.', 'Melt Calculator') | Out-Null; return
    }

    # Attempt to fetch live spot prices (fallback to defaults)
    $silverSpot = 32.50; $goldSpot = 2650.00; $copperSpot = 4.20
    try {
        $lblStatus.Text = 'Fetching live metal spot prices...'
        $mainWin.Dispatcher.Invoke([Action]{}, 'Render')
        $resp = Invoke-RestMethod -Uri 'https://api.metalpriceapi.com/v1/latest?api_key=demo&base=USD&currencies=XAU,XAG,XCU' -TimeoutSec 5 -ErrorAction SilentlyContinue
        if ($resp -and $resp.rates) {
            if ($resp.rates.USDXAG) { $silverSpot = [Math]::Round(1 / $resp.rates.USDXAG, 2) }
            if ($resp.rates.USDXAU) { $goldSpot   = [Math]::Round(1 / $resp.rates.USDXAU, 2) }
        }
    } catch { }

    $totalSilverOz = 0.0; $totalGoldOz = 0.0; $totalCopperLbs = 0.0
    $denomResults  = @{}

    foreach ($c in $Coins) {
        $den = $c.Denomination
        $comp = $script:MeltCompositions[$den]
        if ($null -eq $comp) { continue }
        $dupMult = [Math]::Max(1, [int]$c.DupCount + 1)
        $sOz = $comp.Silver * $dupMult
        $gOz = $comp.Gold   * $dupMult
        $cLb = $comp.Copper * $dupMult
        $totalSilverOz += $sOz
        $totalGoldOz   += $gOz
        $totalCopperLbs += $cLb
        if (-not $denomResults.ContainsKey($den)) {
            $denomResults[$den] = @{ Count = 0; Silver = 0.0; Gold = 0.0; Copper = 0.0 }
        }
        $denomResults[$den].Count  += $dupMult
        $denomResults[$den].Silver += $sOz
        $denomResults[$den].Gold   += $gOz
        $denomResults[$den].Copper += $cLb
    }

    $silverVal = [Math]::Round($totalSilverOz * $silverSpot, 2)
    $goldVal   = [Math]::Round($totalGoldOz * $goldSpot, 2)
    $copperVal = [Math]::Round($totalCopperLbs * $copperSpot, 2)
    $totalMelt = $silverVal + $goldVal + $copperVal

    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add('=' * 80)
    $lines.Add('  MELT VALUE CALCULATOR')
    $lines.Add(("  Generated: {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')))
    $lines.Add('=' * 80)
    $lines.Add('')
    $lines.Add('  SPOT PRICES USED:')
    $lines.Add(("    Silver: `${0:N2}/oz   Gold: `${1:N2}/oz   Copper: `${2:N2}/lb" -f $silverSpot, $goldSpot, $copperSpot))
    $lines.Add('')
    $lines.Add('  TOTAL METAL CONTENT:')
    $lines.Add(("    Silver: {0:N4} troy oz  =  `${1:N2}" -f $totalSilverOz, $silverVal))
    $lines.Add(("    Gold:   {0:N4} troy oz  =  `${1:N2}" -f $totalGoldOz, $goldVal))
    $lines.Add(("    Copper: {0:N4} lbs      =  `${1:N2}" -f $totalCopperLbs, $copperVal))
    $lines.Add('')
    $lines.Add(("  >>> TOTAL MELT VALUE: `${0:N2} <<<" -f $totalMelt))
    $lines.Add('')
    $lines.Add(('-' * 80))
    $lines.Add(("{0,-32} {1,6} {2,12} {3,12} {4,12}" -f 'Denomination','Count','Silver$','Gold$','Total$'))
    $lines.Add(('-' * 80))
    foreach ($kv in ($denomResults.GetEnumerator() | Sort-Object { $_.Value.Silver * $silverSpot + $_.Value.Gold * $goldSpot } -Descending)) {
        $sv = [Math]::Round($kv.Value.Silver * $silverSpot, 2)
        $gv = [Math]::Round($kv.Value.Gold * $goldSpot, 2)
        $tv = $sv + $gv + [Math]::Round($kv.Value.Copper * $copperSpot, 2)
        if ($tv -gt 0) {
            $lines.Add(("{0,-32} {1,6} {2,12:N2} {3,12:N2} {4,12:N2}" -f $kv.Key, $kv.Value.Count, $sv, $gv, $tv))
        }
    }
    $lines.Add(('-' * 80))
    $lines.Add('')
    $lines.Add('  Note: Melt values based on known compositions for 90% silver, 40% silver,')
    $lines.Add('  and gold coins. Post-1964 clad coins have negligible melt value.')
    $lines.Add('  Silver/gold American Eagles use 1 oz / standard bullion weight.')
    $lines.Add('=' * 80)

    $reportDir = Join-Path $PSScriptRoot 'Reports'
    if (-not (Test-Path $reportDir)) { New-Item $reportDir -ItemType Directory -Force | Out-Null }
    $filePath = Join-Path $reportDir ("MeltValue_{0}.txt" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $lines | Set-Content -Path $filePath -Encoding UTF8
    Start-Process notepad.exe $filePath
    $lblStatus.Text = ("Melt value: `${0:N2} (Silver `${1:N2} + Gold `${2:N2})" -f $totalMelt, $silverVal, $goldVal)
}

# ---------------------------------------------------------------------------
# Charts & Analytics (generates HTML report opened in browser)
# ---------------------------------------------------------------------------
function Show-ChartsReport {
    param([object[]]$Coins)
    if ($null -eq $Coins -or $Coins.Count -eq 0) {
        [System.Windows.MessageBox]::Show('No coins to chart.', 'Charts') | Out-Null; return
    }

    $holdCount = @($Coins | Where-Object { $_.InCollection -eq 'Y' }).Count
    $missCount = @($Coins | Where-Object { $_.InCollection -ne 'Y' }).Count
    $totalBook = [Math]::Round(($Coins | ForEach-Object { [double]$_.BookValue } | Measure-Object -Sum).Sum, 2)
    $totalFace = [Math]::Round(($Coins | ForEach-Object { [double]$_.FaceValue } | Measure-Object -Sum).Sum, 2)

    # By denomination (top 15)
    $byDenom = $Coins | Group-Object Denomination | Sort-Object Count -Descending | Select-Object -First 15
    $denomLabels = ($byDenom | ForEach-Object { $n = $_.Name -replace "'",""; "'$n'" }) -join ','
    $denomCounts = ($byDenom | ForEach-Object { $_.Count }) -join ','
    $denomColors = @('#104861','#C62828','#2E7D32','#E65100','#6A1B9A','#00897B','#F9A825','#AD1457','#795548','#37474F','#1565C0','#558B2F','#BF360C','#4527A0','#00695C')
    $denomColStr = ($denomColors | ForEach-Object { "'$_'" }) -join ','

    # By denomination value (top 15)
    $byDenomVal = $Coins | Group-Object Denomination | ForEach-Object {
        $bv = [Math]::Round(($_.Group | ForEach-Object { [double]$_.BookValue } | Measure-Object -Sum).Sum, 2)
        [PSCustomObject]@{ Name = $_.Name; Value = $bv }
    } | Sort-Object Value -Descending | Select-Object -First 15
    $dvLabels = ($byDenomVal | ForEach-Object { $n = $_.Name -replace "'",""; "'$n'" }) -join ','
    $dvValues = ($byDenomVal | ForEach-Object { $_.Value }) -join ','

    # Grade distribution
    $byGrade = $Coins | Where-Object { $_.Grade -ne '' } | Group-Object Grade | Sort-Object Count -Descending | Select-Object -First 12
    $gradeLabels = ($byGrade | ForEach-Object { $n = $_.Name -replace "'",""; "'$n'" }) -join ','
    $gradeCounts = ($byGrade | ForEach-Object { $_.Count }) -join ','

    # By mint
    $byMint = $Coins | Group-Object Mint | Sort-Object Count -Descending
    $mintLabels = ($byMint | ForEach-Object { $n = if ([string]::IsNullOrWhiteSpace($_.Name)) { 'None' } else { $_.Name }; "'$n'" }) -join ','
    $mintCounts = ($byMint | ForEach-Object { $_.Count }) -join ','

    # By decade
    $byDecade = $Coins | Where-Object { $_.Year -match '^\d{4}$' } | Group-Object { [Math]::Floor([int]$_.Year / 10) * 10 } | Sort-Object Name
    $decadeLabels = ($byDecade | ForEach-Object { "'$($_.Name)s'" }) -join ','
    $decadeCounts = ($byDecade | ForEach-Object { $_.Count }) -join ','

    # By metal
    $byMetal = $Coins | Where-Object { $_.Metal -ne '' } | Group-Object Metal | Sort-Object Count -Descending
    $metalLabels = ($byMetal | ForEach-Object { $n = $_.Name -replace "'",""; "'$n'" }) -join ','
    $metalCounts = ($byMetal | ForEach-Object { $_.Count }) -join ','

    $html = @"
<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<title>Coin Collection Analytics</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
  body { font-family: 'Segoe UI', sans-serif; background: #E8ECF0; margin: 0; padding: 20px; color: #1A1A2E; }
  h1 { color: #104861; border-bottom: 3px solid #104861; padding-bottom: 8px; }
  h2 { color: #104861; margin-top: 30px; }
  .stats-row { display: flex; gap: 15px; flex-wrap: wrap; margin: 15px 0; }
  .stat-card { background: white; border-radius: 8px; padding: 15px 25px; text-align: center; border: 1px solid #B0B8C4; min-width: 120px; }
  .stat-card .num { font-size: 28px; font-weight: bold; color: #104861; }
  .stat-card .lbl { font-size: 12px; color: #666; }
  .chart-row { display: flex; gap: 20px; flex-wrap: wrap; margin: 20px 0; }
  .chart-box { background: white; border-radius: 8px; padding: 15px; border: 1px solid #B0B8C4; flex: 1; min-width: 400px; max-width: 600px; }
  canvas { max-height: 350px; }
  .footer { text-align: center; color: #888; font-size: 11px; margin-top: 30px; padding: 15px; border-top: 1px solid #B0B8C4; }
</style></head><body>
<h1>United States Coin Collection Analytics</h1>
<p>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Collection: $($script:ActiveCollection)</p>
<div class="stats-row">
  <div class="stat-card"><div class="num">$holdCount</div><div class="lbl">Holding</div></div>
  <div class="stat-card"><div class="num" style="color:#C62828">$missCount</div><div class="lbl">Missing</div></div>
  <div class="stat-card"><div class="num">$($Coins.Count)</div><div class="lbl">Total</div></div>
  <div class="stat-card"><div class="num" style="color:#2E7D32">$(if($Coins.Count -gt 0){[Math]::Round(($holdCount/$Coins.Count)*100,1)}else{0})%</div><div class="lbl">Complete</div></div>
  <div class="stat-card"><div class="num">`$$($totalBook.ToString('N0'))</div><div class="lbl">Book Value</div></div>
  <div class="stat-card"><div class="num">`$$($totalFace.ToString('N0'))</div><div class="lbl">Face Value</div></div>
</div>
<div class="chart-row">
  <div class="chart-box"><h2>Coins by Denomination (Top 15)</h2><canvas id="denomChart"></canvas></div>
  <div class="chart-box"><h2>Book Value by Denomination</h2><canvas id="denomValChart"></canvas></div>
</div>
<div class="chart-row">
  <div class="chart-box"><h2>Grade Distribution</h2><canvas id="gradeChart"></canvas></div>
  <div class="chart-box"><h2>Coins by Mint Mark</h2><canvas id="mintChart"></canvas></div>
</div>
<div class="chart-row">
  <div class="chart-box"><h2>Coins by Decade</h2><canvas id="decadeChart"></canvas></div>
  <div class="chart-box"><h2>Coins by Metal</h2><canvas id="metalChart"></canvas></div>
</div>
<div class="chart-row">
  <div class="chart-box"><h2>Collection Status</h2><canvas id="statusChart"></canvas></div>
</div>
<div class="footer">US Coin Collection Dashboard v4.0.0 | Analytics Report</div>
<script>
const colors = [$denomColStr];
new Chart(document.getElementById('denomChart'),{type:'bar',data:{labels:[$denomLabels],datasets:[{label:'Count',data:[$denomCounts],backgroundColor:colors}]},options:{indexAxis:'y',plugins:{legend:{display:false}}}});
new Chart(document.getElementById('denomValChart'),{type:'bar',data:{labels:[$dvLabels],datasets:[{label:'Book Value `$',data:[$dvValues],backgroundColor:'#104861'}]},options:{indexAxis:'y',plugins:{legend:{display:false}}}});
new Chart(document.getElementById('gradeChart'),{type:'bar',data:{labels:[$gradeLabels],datasets:[{label:'Count',data:[$gradeCounts],backgroundColor:'#6A1B9A'}]},options:{plugins:{legend:{display:false}}}});
new Chart(document.getElementById('mintChart'),{type:'doughnut',data:{labels:[$mintLabels],datasets:[{data:[$mintCounts],backgroundColor:colors}]},options:{plugins:{legend:{position:'right'}}}});
new Chart(document.getElementById('decadeChart'),{type:'bar',data:{labels:[$decadeLabels],datasets:[{label:'Count',data:[$decadeCounts],backgroundColor:'#00897B'}]},options:{plugins:{legend:{display:false}}}});
new Chart(document.getElementById('metalChart'),{type:'doughnut',data:{labels:[$metalLabels],datasets:[{data:[$metalCounts],backgroundColor:colors}]},options:{plugins:{legend:{position:'right'}}}});
new Chart(document.getElementById('statusChart'),{type:'doughnut',data:{labels:['Holding','Missing'],datasets:[{data:[$holdCount,$missCount],backgroundColor:['#2E7D32','#C62828']}]},options:{plugins:{legend:{position:'right'}}}});
</script></body></html>
"@
    $reportDir = Join-Path $PSScriptRoot 'Reports'
    if (-not (Test-Path $reportDir)) { New-Item $reportDir -ItemType Directory -Force | Out-Null }
    $filePath = Join-Path $reportDir ("Charts_{0}.html" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $html | Set-Content -Path $filePath -Encoding UTF8
    Start-Process $filePath
    $lblStatus.Text = 'Charts report opened in browser'
}

# ---------------------------------------------------------------------------
# 2x2 Flip Label Printing (generates printable HTML)
# ---------------------------------------------------------------------------
function Show-PrintLabels {
    param([object[]]$Coins)
    if ($null -eq $Coins -or $Coins.Count -eq 0) {
        [System.Windows.MessageBox]::Show('No coins selected for labels.', 'Labels') | Out-Null; return
    }

    $labelHtml = [System.Text.StringBuilder]::new()
    $null = $labelHtml.Append(@"
<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Coin Labels</title>
<style>
  @page { margin: 0.4in; }
  body { font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 10px; }
  .label-grid { display: flex; flex-wrap: wrap; gap: 4px; }
  .label {
    width: 1.85in; height: 1.85in; border: 1px dashed #999;
    padding: 4px; box-sizing: border-box; page-break-inside: avoid;
    display: flex; flex-direction: column; justify-content: center;
    align-items: center; text-align: center; font-size: 9pt;
  }
  .label .denom { font-weight: bold; font-size: 10pt; color: #104861; margin-bottom: 2px; }
  .label .year-mint { font-size: 12pt; font-weight: bold; margin: 2px 0; }
  .label .grade { font-size: 9pt; color: #444; }
  .label .value { font-size: 8pt; color: #666; margin-top: 2px; }
  .label .cert { font-size: 7pt; color: #888; margin-top: 1px; }
  .label .metal { font-size: 7pt; color: #888; }
  .label .key { font-size: 8pt; color: #C62828; font-weight: bold; }
  h1 { font-size: 14pt; color: #104861; margin: 5px 0 10px 0; }
  .instructions { font-size: 9pt; color: #666; margin-bottom: 10px; }
  @media print { h1, .instructions { display: none; } .label { border: 1px dashed #CCC; } }
</style></head><body>
<h1>Coin Labels - $($Coins.Count) labels</h1>
<p class="instructions">Print at 100% scale. Cut along dashed lines. Labels sized for standard 2x2 holders (1.85" x 1.85").</p>
<div class="label-grid">
"@)

    foreach ($c in ($Coins | Sort-Object @{E={[int]$_.Year}}, Mint)) {
        $certLine = ''
        if (-not [string]::IsNullOrWhiteSpace($c.CertCompany)) {
            $certLine = "$($c.CertCompany)"
            if (-not [string]::IsNullOrWhiteSpace($c.CertNumber)) { $certLine += " #$($c.CertNumber)" }
            if (-not [string]::IsNullOrWhiteSpace($c.SlabGrade)) { $certLine += " ($($c.SlabGrade))" }
        }
        $keyLine = if (-not [string]::IsNullOrWhiteSpace($c.KeyDates)) { $c.KeyDates } else { '' }
        $mintDisp = if ([string]::IsNullOrWhiteSpace($c.Mint)) { '' } else { "-$($c.Mint)" }
        $typeDisp = if ([string]::IsNullOrWhiteSpace($c.Type)) { '' } else { "<div class=`"grade`">$([System.Web.HttpUtility]::HtmlEncode($c.Type))</div>" }
        $gradeDisp = if ([string]::IsNullOrWhiteSpace($c.Grade)) { '' } else { "<div class=`"grade`">$([System.Web.HttpUtility]::HtmlEncode($c.Grade))</div>" }
        $metalDisp = if ([string]::IsNullOrWhiteSpace($c.Metal)) { '' } else { "<div class=`"metal`">$([System.Web.HttpUtility]::HtmlEncode($c.Metal))</div>" }
        $null = $labelHtml.Append(@"
<div class="label">
  <div class="denom">$([System.Web.HttpUtility]::HtmlEncode($c.Denomination))</div>
  <div class="year-mint">$($c.Year)$mintDisp</div>
  $typeDisp
  $gradeDisp
  <div class="value">FV: `$$($c.FaceValue)  BV: `$$($c.BookValue)</div>
  $metalDisp
  $(if($certLine){"<div class=`"cert`">$([System.Web.HttpUtility]::HtmlEncode($certLine))</div>"})
  $(if($keyLine){"<div class=`"key`">$([System.Web.HttpUtility]::HtmlEncode($keyLine))</div>"})
</div>
"@)
    }

    $null = $labelHtml.Append('</div></body></html>')

    $reportDir = Join-Path $PSScriptRoot 'Reports'
    if (-not (Test-Path $reportDir)) { New-Item $reportDir -ItemType Directory -Force | Out-Null }
    $filePath = Join-Path $reportDir ("Labels_{0}.html" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $labelHtml.ToString() | Set-Content -Path $filePath -Encoding UTF8
    Start-Process $filePath
    $lblStatus.Text = ("Generated {0} coin labels - opened in browser for printing" -f $Coins.Count)
}

# ---------------------------------------------------------------------------
# Insurance / Appraisal Report (professional HTML format)
# ---------------------------------------------------------------------------
function Show-InsuranceReport {
    param([object[]]$Coins)
    if ($null -eq $Coins -or $Coins.Count -eq 0) {
        [System.Windows.MessageBox]::Show('No coins to report.', 'Insurance Report') | Out-Null; return
    }
    $holding = @($Coins | Where-Object { $_.InCollection -eq 'Y' })
    $totalBook = [Math]::Round(($holding | ForEach-Object { [double]$_.BookValue } | Measure-Object -Sum).Sum, 2)
    $totalCurr = [Math]::Round(($holding | ForEach-Object { [double]$_.CurrentValue } | Measure-Object -Sum).Sum, 2)
    $totalPaid = [Math]::Round(($holding | ForEach-Object { [double]$_.PurchasePrice } | Measure-Object -Sum).Sum, 2)
    $appValue  = [Math]::Max($totalBook, $totalCurr)
    $stamp     = Get-Date -Format 'MMMM dd, yyyy'

    $html = [System.Text.StringBuilder]::new()
    $null = $html.Append(@"
<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Insurance Appraisal Report</title>
<style>
  @page { margin: 0.5in; }
  body { font-family: 'Segoe UI',Georgia,serif; color: #1A1A2E; margin: 0; padding: 20px; font-size: 11pt; }
  h1 { color: #104861; border-bottom: 3px double #104861; padding-bottom: 8px; font-size: 20pt; }
  h2 { color: #104861; border-bottom: 1px solid #B0B8C4; padding-bottom: 4px; margin-top: 25px; }
  .header-info { display: flex; justify-content: space-between; margin: 10px 0 20px; }
  .header-info div { font-size: 10pt; }
  .summary-box { background: #F0F4F7; border: 1px solid #B0B8C4; border-radius: 5px; padding: 15px; margin: 15px 0; }
  .summary-row { display: flex; justify-content: space-between; padding: 4px 0; }
  .summary-row .label { font-weight: bold; color: #104861; }
  .summary-row .value { font-weight: bold; font-size: 12pt; }
  table { width: 100%; border-collapse: collapse; margin: 10px 0; font-size: 9pt; }
  th { background: #104861; color: white; padding: 6px 8px; text-align: left; font-size: 8pt; }
  td { padding: 5px 8px; border-bottom: 1px solid #DDD; }
  tr:nth-child(even) { background: #F8F8F8; }
  .cert-badge { background: #E3F2FD; color: #1565C0; padding: 1px 6px; border-radius: 3px; font-size: 8pt; }
  .key-badge { background: #FBE9E7; color: #BF360C; padding: 1px 6px; border-radius: 3px; font-size: 8pt; font-weight: bold; }
  .footer { margin-top: 40px; padding-top: 15px; border-top: 2px solid #104861; font-size: 9pt; color: #666; }
  .sig-line { margin-top: 50px; border-top: 1px solid #333; width: 250px; }
  .sig-label { font-size: 9pt; color: #666; }
  @media print { body { padding: 0; } }
</style></head><body>
<h1>Coin Collection Insurance &amp; Appraisal Report</h1>
<div class="header-info">
  <div><strong>Collection:</strong> $($script:ActiveCollection)<br><strong>Date:</strong> $stamp</div>
  <div><strong>Total Items:</strong> $($holding.Count) coins<br><strong>Appraised Value:</strong> `$$($appValue.ToString('N2'))</div>
</div>
<div class="summary-box">
  <div class="summary-row"><span class="label">Total Book Value:</span><span class="value">`$$($totalBook.ToString('N2'))</span></div>
  <div class="summary-row"><span class="label">Total Current Market Value:</span><span class="value">`$$($totalCurr.ToString('N2'))</span></div>
  <div class="summary-row"><span class="label">Total Purchase Cost:</span><span class="value">`$$($totalPaid.ToString('N2'))</span></div>
  <div class="summary-row"><span class="label">Recommended Insurance Coverage:</span><span class="value" style="color:#C62828;">`$$($appValue.ToString('N2'))</span></div>
</div>
<h2>Detailed Inventory</h2>
<table><tr><th>#</th><th>Year</th><th>Mint</th><th>Denomination</th><th>Type</th><th>Grade</th><th>Cert</th><th>Face `$</th><th>Book `$</th><th>Value `$</th><th>Key</th><th>Notes</th></tr>
"@)
    $i = 0
    foreach ($c in ($holding | Sort-Object @{E={[double]$_.BookValue}} -Descending)) {
        $i++
        $certHtml = ''
        if (-not [string]::IsNullOrWhiteSpace($c.CertCompany)) {
            $certHtml = "<span class='cert-badge'>$([System.Web.HttpUtility]::HtmlEncode($c.CertCompany))"
            if (-not [string]::IsNullOrWhiteSpace($c.SlabGrade)) { $certHtml += " $([System.Web.HttpUtility]::HtmlEncode($c.SlabGrade))" }
            $certHtml += '</span>'
        }
        $keyHtml = if (-not [string]::IsNullOrWhiteSpace($c.KeyDates)) { "<span class='key-badge'>$([System.Web.HttpUtility]::HtmlEncode($c.KeyDates))</span>" } else { '' }
        $null = $html.Append("<tr><td>$i</td><td>$($c.Year)</td><td>$($c.Mint)</td><td>$([System.Web.HttpUtility]::HtmlEncode($c.Denomination))</td><td>$([System.Web.HttpUtility]::HtmlEncode($c.Type))</td><td>$([System.Web.HttpUtility]::HtmlEncode($c.Grade))</td><td>$certHtml</td><td>$($c.FaceValue)</td><td>$($c.BookValue)</td><td>$($c.CurrentValue)</td><td>$keyHtml</td><td>$([System.Web.HttpUtility]::HtmlEncode($c.Comments))</td></tr>`n")
    }
    $null = $html.Append(@"
</table>
<div class="footer">
  <p>This report represents the fair market value of the items listed above as of $stamp.
  Values are based on current price guide data and recent market transactions.</p>
  <div class="sig-line"></div>
  <div class="sig-label">Owner Signature / Date</div>
  <br><br>
  <div class="sig-line"></div>
  <div class="sig-label">Appraiser Signature / Date (if applicable)</div>
</div>
</body></html>
"@)
    $reportDir = Join-Path $PSScriptRoot 'Reports'
    if (-not (Test-Path $reportDir)) { New-Item $reportDir -ItemType Directory -Force | Out-Null }
    $filePath = Join-Path $reportDir ("InsuranceReport_{0}.html" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $html.ToString() | Set-Content -Path $filePath -Encoding UTF8
    Start-Process $filePath
    $lblStatus.Text = ("Insurance report: {0} coins, appraised value `${1:N2}" -f $holding.Count, $appValue)
}

# ---------------------------------------------------------------------------
# Set Completion Tracking (popular US coin series)
# ---------------------------------------------------------------------------
$script:CoinSets = [ordered]@{
    'Lincoln Cent (1909-2024)'       = @{ Denom = 'Lincoln Cent'; YearMin = 1909; YearMax = 2024 }
    'Indian Head Cent (1859-1909)'   = @{ Denom = 'Indian Head Cent'; YearMin = 1859; YearMax = 1909 }
    'Jefferson Nickel (1938-2024)'   = @{ Denom = 'Jefferson Nickel'; YearMin = 1938; YearMax = 2024 }
    'Buffalo Nickel (1913-1938)'     = @{ Denom = 'Buffalo Nickel'; YearMin = 1913; YearMax = 1938 }
    'Roosevelt Dime (1946-2024)'     = @{ Denom = 'Roosevelt Dime'; YearMin = 1946; YearMax = 2024 }
    'Mercury Dime (1916-1945)'       = @{ Denom = 'Mercury Dime'; YearMin = 1916; YearMax = 1945 }
    'Washington Quarter (1932-2024)' = @{ Denom = 'Washington Quarter'; YearMin = 1932; YearMax = 2024 }
    'Standing Liberty Quarter'       = @{ Denom = 'Standing Liberty Quarter'; YearMin = 1916; YearMax = 1930 }
    'Kennedy Half Dollar (1964-2024)'= @{ Denom = 'Kennedy Half Dollar'; YearMin = 1964; YearMax = 2024 }
    'Franklin Half Dollar (1948-63)' = @{ Denom = 'Franklin Half Dollar'; YearMin = 1948; YearMax = 1963 }
    'Walking Liberty Half (1916-47)' = @{ Denom = 'Walking Liberty Half Dollar'; YearMin = 1916; YearMax = 1947 }
    'Morgan Dollar (1878-1921)'      = @{ Denom = 'Morgan Dollar'; YearMin = 1878; YearMax = 1921 }
    'Peace Dollar (1921-1935)'       = @{ Denom = 'Peace Dollar'; YearMin = 1921; YearMax = 1935 }
    'Eisenhower Dollar (1971-1978)'  = @{ Denom = 'Eisenhower Dollar'; YearMin = 1971; YearMax = 1978 }
    'Barber Dime (1892-1916)'        = @{ Denom = 'Barber Dime'; YearMin = 1892; YearMax = 1916 }
    'Barber Quarter (1892-1916)'     = @{ Denom = 'Barber Quarter'; YearMin = 1892; YearMax = 1916 }
    'Barber Half Dollar (1892-1915)' = @{ Denom = 'Barber Half Dollar'; YearMin = 1892; YearMax = 1915 }
    'Seated Liberty Dime (1837-91)'  = @{ Denom = 'Seated Liberty Dime'; YearMin = 1837; YearMax = 1891 }
    'Seated Liberty Quarter (1838-91)'= @{ Denom = 'Seated Liberty Quarter'; YearMin = 1838; YearMax = 1891 }
    'Seated Liberty Half (1839-91)'  = @{ Denom = 'Seated Liberty Half Dollar'; YearMin = 1839; YearMax = 1891 }
}

function Show-SetTracker {
    param([object[]]$Coins)
    if ($null -eq $Coins -or $Coins.Count -eq 0) {
        [System.Windows.MessageBox]::Show('No coins to track.', 'Set Tracker') | Out-Null; return
    }

    $html = [System.Text.StringBuilder]::new()
    $null = $html.Append(@"
<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Set Completion Tracker</title>
<style>
  body { font-family: 'Segoe UI', sans-serif; background: #E8ECF0; margin: 0; padding: 20px; color: #1A1A2E; }
  h1 { color: #104861; border-bottom: 3px solid #104861; padding-bottom: 8px; }
  .set-card { background: white; border: 1px solid #B0B8C4; border-radius: 8px; padding: 15px; margin: 12px 0; }
  .set-title { font-size: 14pt; font-weight: bold; color: #104861; margin-bottom: 8px; }
  .progress-bar { background: #E0E0E0; border-radius: 10px; height: 24px; overflow: hidden; margin: 6px 0; }
  .progress-fill { height: 100%; border-radius: 10px; display: flex; align-items: center; justify-content: center; color: white; font-weight: bold; font-size: 11px; min-width: 40px; transition: width 0.3s; }
  .stats { display: flex; gap: 20px; font-size: 10pt; color: #555; margin-top: 6px; }
  .missing-list { font-size: 9pt; color: #888; margin-top: 6px; max-height: 60px; overflow-y: auto; }
  .missing-list summary { cursor: pointer; color: #104861; font-weight: bold; }
  .footer { text-align: center; color: #888; font-size: 11px; margin-top: 20px; }
</style></head><body>
<h1>Set Completion Tracker</h1>
<p>Tracking $($script:CoinSets.Count) popular US coin series against your collection.</p>
"@)

    foreach ($kv in $script:CoinSets.GetEnumerator()) {
        $setName = $kv.Key
        $def     = $kv.Value
        $matching = @($Coins | Where-Object { $_.Denomination -eq $def.Denom -and $_.InCollection -eq 'Y' })
        $totalInSet = @($Coins | Where-Object { $_.Denomination -eq $def.Denom })
        $have = $matching.Count
        $total = $totalInSet.Count
        if ($total -eq 0) { $total = 1 }
        $pct = [Math]::Round(($have / $total) * 100, 1)
        $color = if ($pct -ge 80) { '#2E7D32' } elseif ($pct -ge 50) { '#E65100' } elseif ($pct -ge 20) { '#F9A825' } else { '#C62828' }
        $bookVal = [Math]::Round(($matching | ForEach-Object { [double]$_.BookValue } | Measure-Object -Sum).Sum, 2)
        $haveYears = @($matching | ForEach-Object { "$($_.Year)-$($_.Mint)" } | Sort-Object -Unique)
        $allYears  = @($totalInSet | ForEach-Object { "$($_.Year)-$($_.Mint)" } | Sort-Object -Unique)
        $missingYears = @($allYears | Where-Object { $_ -notin $haveYears }) | Select-Object -First 20
        $missingText = if ($missingYears.Count -gt 0) { ($missingYears -join ', ') } else { 'None - Set Complete!' }
        $null = $html.Append(@"
<div class="set-card">
  <div class="set-title">$([System.Web.HttpUtility]::HtmlEncode($setName))</div>
  <div class="progress-bar"><div class="progress-fill" style="width:$([Math]::Max(5,$pct))%;background:$color;">$pct%</div></div>
  <div class="stats">
    <span>Have: <strong>$have</strong></span>
    <span>Total: <strong>$($totalInSet.Count)</strong></span>
    <span>Missing: <strong>$($totalInSet.Count - $have)</strong></span>
    <span>Book Value: <strong>`$$($bookVal.ToString('N2'))</strong></span>
  </div>
  <div class="missing-list"><details><summary>Missing coins (click to expand)</summary>$([System.Web.HttpUtility]::HtmlEncode($missingText))</details></div>
</div>
"@)
    }

    $null = $html.Append('<div class="footer">US Coin Collection Dashboard - Set Completion Tracker</div></body></html>')
    $reportDir = Join-Path $PSScriptRoot 'Reports'
    if (-not (Test-Path $reportDir)) { New-Item $reportDir -ItemType Directory -Force | Out-Null }
    $filePath = Join-Path $reportDir ("SetTracker_{0}.html" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $html.ToString() | Set-Content -Path $filePath -Encoding UTF8
    Start-Process $filePath
    $lblStatus.Text = 'Set completion tracker opened in browser'
}

# ---------------------------------------------------------------------------
# PCGS Cert Lookup (Public API integration)
# ---------------------------------------------------------------------------
function Invoke-PCGSLookup {
    param([string]$CertNumber)
    if ([string]::IsNullOrWhiteSpace($CertNumber)) { return $null }
    $lblStatus.Text = ("Looking up PCGS cert {0}..." -f $CertNumber)
    $mainWin.Dispatcher.Invoke([Action]{}, 'Render')
    try {
        $url = "https://www.pcgs.com/cert/$CertNumber"
        $response = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
        $html = $response.Content

        $result = @{}
        # Parse grade from page
        if ($html -match 'Grade:\s*</[^>]+>\s*<[^>]+>([^<]+)') { $result['Grade'] = $Matches[1].Trim() }
        if ($html -match 'Denomination:\s*</[^>]+>\s*<[^>]+>([^<]+)') { $result['Denomination'] = $Matches[1].Trim() }
        if ($html -match 'Year:\s*</[^>]+>\s*<[^>]+>([^<]+)') { $result['Year'] = $Matches[1].Trim() }
        if ($html -match 'Mint Mark:\s*</[^>]+>\s*<[^>]+>([^<]+)') { $result['Mint'] = $Matches[1].Trim() }
        if ($html -match 'Variety:\s*</[^>]+>\s*<[^>]+>([^<]+)') { $result['Variety'] = $Matches[1].Trim() }
        if ($html -match 'Trueview') { $result['HasTrueView'] = $true }

        if ($result.Count -gt 0) {
            $result['CertNumber'] = $CertNumber
            $result['Source'] = $url
            return $result
        }
    } catch { }
    return $null
}

# ---------------------------------------------------------------------------
# Value History Tracking
# ---------------------------------------------------------------------------
function Record-ValueSnapshot {
    param([object[]]$Coins)
    if ($null -eq $Coins -or $Coins.Count -eq 0) { return }
    $stamp = Get-Date -Format 'yyyy-MM-dd'
    $totalBook = [Math]::Round(($Coins | ForEach-Object { [double]$_.BookValue } | Measure-Object -Sum).Sum, 2)
    $totalCurr = [Math]::Round(($Coins | ForEach-Object { [double]$_.CurrentValue } | Measure-Object -Sum).Sum, 2)
    $totalFace = [Math]::Round(($Coins | ForEach-Object { [double]$_.FaceValue } | Measure-Object -Sum).Sum, 2)
    $holdCount = @($Coins | Where-Object { $_.InCollection -eq 'Y' }).Count
    $totalCount = $Coins.Count

    $entry = [PSCustomObject]@{
        Date        = $stamp
        Collection  = $script:ActiveCollection
        TotalCoins  = $totalCount
        Holding     = $holdCount
        BookValue   = $totalBook
        CurrentValue= $totalCurr
        FaceValue   = $totalFace
    }

    if (-not (Test-Path $script:ValueHistoryFile)) {
        'Date,Collection,TotalCoins,Holding,BookValue,CurrentValue,FaceValue' |
            Set-Content -Path $script:ValueHistoryFile -Encoding UTF8
    }
    $entry | Export-Csv -Path $script:ValueHistoryFile -NoTypeInformation -Encoding UTF8 -Append
}

function Show-ValueHistory {
    if (-not (Test-Path $script:ValueHistoryFile)) {
        [System.Windows.MessageBox]::Show(
            "No value history recorded yet.`n`nThe app will now record a snapshot. Click again later to see trends.",
            'Value History') | Out-Null
        Record-ValueSnapshot -Coins @(Get-Collection)
        return
    }
    $history = @(Import-Csv -Path $script:ValueHistoryFile -Encoding UTF8 |
        Where-Object { $_.Collection -eq $script:ActiveCollection } |
        Sort-Object Date)

    if ($history.Count -eq 0) {
        Record-ValueSnapshot -Coins @(Get-Collection)
        $history = @(Import-Csv -Path $script:ValueHistoryFile -Encoding UTF8 |
            Where-Object { $_.Collection -eq $script:ActiveCollection } |
            Sort-Object Date)
    }

    $dates     = ($history | ForEach-Object { "'$($_.Date)'" }) -join ','
    $bookVals  = ($history | ForEach-Object { $_.BookValue }) -join ','
    $currVals  = ($history | ForEach-Object { $_.CurrentValue }) -join ','
    $faceVals  = ($history | ForEach-Object { $_.FaceValue }) -join ','
    $coinCounts= ($history | ForEach-Object { $_.TotalCoins }) -join ','
    $holdCounts= ($history | ForEach-Object { $_.Holding }) -join ','

    $latest = $history[-1]
    $first  = $history[0]
    $bookChange = [Math]::Round([double]$latest.BookValue - [double]$first.BookValue, 2)
    $changeColor = if ($bookChange -ge 0) { '#2E7D32' } else { '#C62828' }
    $changeSign  = if ($bookChange -ge 0) { '+' } else { '' }

    $html = @"
<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Collection Value History</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
  body { font-family: 'Segoe UI', sans-serif; background: #E8ECF0; margin: 0; padding: 20px; color: #1A1A2E; }
  h1 { color: #104861; border-bottom: 3px solid #104861; padding-bottom: 8px; }
  .stats-row { display: flex; gap: 15px; flex-wrap: wrap; margin: 15px 0; }
  .stat-card { background: white; border-radius: 8px; padding: 15px 25px; text-align: center; border: 1px solid #B0B8C4; min-width: 130px; }
  .stat-card .num { font-size: 24px; font-weight: bold; color: #104861; }
  .stat-card .lbl { font-size: 11px; color: #666; }
  .chart-box { background: white; border-radius: 8px; padding: 20px; border: 1px solid #B0B8C4; margin: 15px 0; }
  table { width: 100%; border-collapse: collapse; margin: 15px 0; font-size: 10pt; }
  th { background: #104861; color: white; padding: 6px 10px; text-align: left; }
  td { padding: 5px 10px; border-bottom: 1px solid #DDD; }
  tr:nth-child(even) { background: #F8F8F8; }
</style></head><body>
<h1>Collection Value History - $($script:ActiveCollection)</h1>
<p>Tracking $($history.Count) snapshots from $($first.Date) to $($latest.Date)</p>
<div class="stats-row">
  <div class="stat-card"><div class="num">`$$([double]$latest.BookValue)</div><div class="lbl">Current Book Value</div></div>
  <div class="stat-card"><div class="num">`$$([double]$latest.CurrentValue)</div><div class="lbl">Current Market Value</div></div>
  <div class="stat-card"><div class="num" style="color:$changeColor">$changeSign`$$bookChange</div><div class="lbl">Book Value Change</div></div>
  <div class="stat-card"><div class="num">$($latest.Holding)</div><div class="lbl">Coins Held</div></div>
  <div class="stat-card"><div class="num">$($history.Count)</div><div class="lbl">Snapshots</div></div>
</div>
<div class="chart-box"><h2 style="color:#104861">Value Over Time</h2><canvas id="valueChart" height="80"></canvas></div>
<div class="chart-box"><h2 style="color:#104861">Collection Size Over Time</h2><canvas id="countChart" height="60"></canvas></div>
<h2 style="color:#104861">History Log</h2>
<table><tr><th>Date</th><th>Coins</th><th>Holding</th><th>Book Value</th><th>Current Value</th><th>Face Value</th></tr>
"@
    foreach ($h in ($history | Sort-Object Date -Descending)) {
        $html += "<tr><td>$($h.Date)</td><td>$($h.TotalCoins)</td><td>$($h.Holding)</td><td>`$$($h.BookValue)</td><td>`$$($h.CurrentValue)</td><td>`$$($h.FaceValue)</td></tr>`n"
    }
    $html += @"
</table>
<script>
new Chart(document.getElementById('valueChart'),{type:'line',data:{labels:[$dates],datasets:[
  {label:'Book Value',data:[$bookVals],borderColor:'#104861',backgroundColor:'rgba(16,72,97,0.1)',fill:true,tension:0.3},
  {label:'Current Value',data:[$currVals],borderColor:'#2E7D32',backgroundColor:'rgba(46,125,50,0.1)',fill:true,tension:0.3},
  {label:'Face Value',data:[$faceVals],borderColor:'#E65100',borderDash:[5,5],fill:false,tension:0.3}
]},options:{plugins:{legend:{position:'top'}},scales:{y:{beginAtZero:false}}}});
new Chart(document.getElementById('countChart'),{type:'line',data:{labels:[$dates],datasets:[
  {label:'Total Coins',data:[$coinCounts],borderColor:'#104861',fill:false,tension:0.3},
  {label:'Holding',data:[$holdCounts],borderColor:'#2E7D32',fill:false,tension:0.3}
]},options:{plugins:{legend:{position:'top'}},scales:{y:{beginAtZero:true}}}});
</script></body></html>
"@
    $reportDir = Join-Path $PSScriptRoot 'Reports'
    if (-not (Test-Path $reportDir)) { New-Item $reportDir -ItemType Directory -Force | Out-Null }
    $filePath = Join-Path $reportDir ("ValueHistory_{0}.html" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $html | Set-Content -Path $filePath -Encoding UTF8
    Start-Process $filePath
    $lblStatus.Text = 'Value history chart opened in browser'
}

# ---------------------------------------------------------------------------
# Restore from Backup
# ---------------------------------------------------------------------------
function Show-RestoreDialog {
    $backupDir = Join-Path $PSScriptRoot 'Backups'
    if (-not (Test-Path $backupDir)) {
        [System.Windows.MessageBox]::Show('No backups found.', 'Restore') | Out-Null; return
    }
    $backups = @(Get-ChildItem $backupDir -Directory -Filter 'Collections_*' |
        Sort-Object LastWriteTime -Descending | Select-Object -First 30)
    if ($backups.Count -eq 0) {
        [System.Windows.MessageBox]::Show('No backups found.', 'Restore') | Out-Null; return
    }
    [xml]$dlgXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Restore Backup" Height="360" Width="480"
        WindowStartupLocation="CenterOwner" ResizeMode="CanResize"
        Background="#F0F2F5">
    <DockPanel Margin="14">
        <TextBlock DockPanel.Dock="Top" Text="Select a backup snapshot to restore:"
                   FontSize="13" FontWeight="Bold" Foreground="#104861" Margin="0,0,0,8"/>
        <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal"
                    HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button Name="btnRestoreOK"     Content="Restore" Width="90" Height="30"
                    Background="#4527A0" Foreground="White" Margin="0,0,8,0" BorderThickness="0"/>
            <Button Name="btnRestoreCancel" Content="Cancel"  Width="80" Height="30"
                    Background="#78909C" Foreground="White" BorderThickness="0"/>
        </StackPanel>
        <ListBox Name="lbBackups" FontSize="12" FontFamily="Consolas"/>
    </DockPanel>
</Window>
'@
    $rdr = [System.Xml.XmlNodeReader]::new($dlgXaml)
    $dlg = [System.Windows.Markup.XamlReader]::Load($rdr)
    $lb      = $dlg.FindName('lbBackups')
    $btnOK   = $dlg.FindName('btnRestoreOK')
    $btnCncl = $dlg.FindName('btnRestoreCancel')
    foreach ($b in $backups) { $lb.Items.Add($b.Name) | Out-Null }
    $lb.SelectedIndex = 0
    $script:RestoreChoice = $null
    $btnCncl.Add_Click({ $dlg.Close() })
    $btnOK.Add_Click({
        if ($lb.SelectedIndex -lt 0) { return }
        $script:RestoreChoice = $backups[$lb.SelectedIndex]
        $dlg.Close()
    })
    $prevEAP = $ErrorActionPreference
    try { $ErrorActionPreference = 'Continue'; $null = $dlg.ShowDialog() }
    finally { $ErrorActionPreference = $prevEAP }
    if ($null -eq $script:RestoreChoice) { return }
    $confirm = [System.Windows.MessageBox]::Show(
        "Restore '$($script:RestoreChoice.Name)'?`nYour current data will be backed up first.",
        'Confirm Restore', 'YesNo', 'Warning')
    if ($confirm -ne 'Yes') { return }
    Invoke-Backup -Force | Out-Null
    # Backup folder may contain CSV files directly or inside a Collections sub-folder
    $srcDir = $script:RestoreChoice.FullName
    $subDir = Join-Path $srcDir 'Collections'
    if (Test-Path $subDir) { $srcDir = $subDir }
    Get-ChildItem -Path $srcDir -Filter '*.csv' |
        ForEach-Object { Copy-Item $_.FullName $script:CollectionsDir -Force }
    return $true
}

# ---------------------------------------------------------------------------
# Statistics Window
# ---------------------------------------------------------------------------
function Show-StatsWindow {
    param([object[]]$Coins)

    if ($null -eq $Coins -or $Coins.Count -eq 0) {
        [System.Windows.MessageBox]::Show('No coins found.', 'Empty') | Out-Null; return
    }

    $totalPurchase = [Math]::Round(($Coins | ForEach-Object { [double]$_.PurchasePrice } | Measure-Object -Sum).Sum, 2)
    $totalCurrent  = [Math]::Round(($Coins | ForEach-Object { [double]$_.CurrentValue  } | Measure-Object -Sum).Sum, 2)
    $totalBook     = [Math]::Round(($Coins | ForEach-Object { [double]$_.BookValue     } | Measure-Object -Sum).Sum, 2)
    $totalFace     = [Math]::Round(($Coins | ForEach-Object { [double]$_.FaceValue     } | Measure-Object -Sum).Sum, 2)
    $gainLoss      = [Math]::Round($totalCurrent - $totalPurchase, 2)
    $glColor       = if ($gainLoss -ge 0) { '#2E7D32' } else { '#C62828' }
    $years         = @($Coins | ForEach-Object { if ($_.Year) { [int]$_.Year } } | Sort-Object)
    $earliest      = if ($years.Count -gt 0) { $years[0]  } else { 'N/A' }
    $latest        = if ($years.Count -gt 0) { $years[-1] } else { 'N/A' }

    $byDenom   = ($Coins | Group-Object Denomination | Sort-Object Count -Descending |
        ForEach-Object { "  $($_.Name): $($_.Count)" }) -join "`n"
    $byMint    = ($Coins | Group-Object Mint | Sort-Object Count -Descending |
        ForEach-Object { $n = if ([string]::IsNullOrWhiteSpace($_.Name)) { '(none)' } else { $_.Name }; "  $n`: $($_.Count)" }) -join "`n"
    $byCountry = ($Coins | Group-Object Country | Sort-Object Count -Descending |
        ForEach-Object { "  $($_.Name): $($_.Count)" }) -join "`n"
    $byMetal   = ($Coins | Where-Object { $_.Metal -ne '' } | Group-Object Metal | Sort-Object Count -Descending |
        ForEach-Object { "  $($_.Name): $($_.Count)" }) -join "`n"
    if (-not $byMetal) { $byMetal = '  (none)' }
    $byGrade   = ($Coins | Where-Object { $_.Grade -ne '' } | Group-Object Grade | Sort-Object Count -Descending |
        ForEach-Object { "  $($_.Name): $($_.Count)" }) -join "`n"
    if (-not $byGrade) { $byGrade = '  (none)' }
    $keyInfo = ($Coins | Where-Object { $_.KeyDates -ne '' } | Group-Object KeyDates | Sort-Object Count -Descending |
        ForEach-Object { "  $($_.Name): $($_.Count)" }) -join "`n"
    if (-not $keyInfo) { $keyInfo = '  (none)' }
    $totalDups  = ($Coins | ForEach-Object { [int]$_.DupCount } | Measure-Object -Sum).Sum
    $totalDupFV = [Math]::Round(($Coins | ForEach-Object { [double]$_.DupFaceValue } | Measure-Object -Sum).Sum, 2)
    $slabbedCoins = @($Coins | Where-Object { $_.CertCompany -ne '' })
    $byCert = ($slabbedCoins | Group-Object CertCompany | Sort-Object Count -Descending |
        ForEach-Object { "  $($_.Name): $($_.Count)" }) -join "`n"
    if (-not $byCert) { $byCert = '  (none)' }
    $inCollY = @($Coins | Where-Object { $_.InCollection -eq 'Y' }).Count
    $inCollN = @($Coins | Where-Object { $_.InCollection -ne 'Y' }).Count

    [xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Collection Statistics" Height="720" Width="520"
        WindowStartupLocation="CenterOwner" ResizeMode="CanResize"
        Background="#F0F2F5">
    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="#1A1A2E"/>
            <Setter Property="FontSize"   Value="12"/>
        </Style>
    </Window.Resources>
    <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="16">
      <StackPanel>
        <TextBlock Text="OVERVIEW" FontSize="15" FontWeight="Bold" Foreground="#104861" Margin="0,0,0,6"/>
        <Grid>
          <Grid.ColumnDefinitions><ColumnDefinition Width="200"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
          <Grid.RowDefinitions>
            <RowDefinition/><RowDefinition/><RowDefinition/><RowDefinition/>
            <RowDefinition/><RowDefinition/><RowDefinition/><RowDefinition/>
            <RowDefinition/><RowDefinition/>
          </Grid.RowDefinitions>
          <TextBlock Grid.Row="0" Grid.Column="0" Text="Total Coins:"/>
          <TextBlock Grid.Row="0" Grid.Column="1" Name="lblTotal"/>
          <TextBlock Grid.Row="1" Grid.Column="0" Text="In Collection (Y):"/>
          <TextBlock Grid.Row="1" Grid.Column="1" Name="lblInCollY"/>
          <TextBlock Grid.Row="2" Grid.Column="0" Text="Not In Collection (N):"/>
          <TextBlock Grid.Row="2" Grid.Column="1" Name="lblInCollN"/>
          <TextBlock Grid.Row="3" Grid.Column="0" Text="Total Face Value:"/>
          <TextBlock Grid.Row="3" Grid.Column="1" Name="lblFace"/>
          <TextBlock Grid.Row="4" Grid.Column="0" Text="Total Book Value:"/>
          <TextBlock Grid.Row="4" Grid.Column="1" Name="lblBook"/>
          <TextBlock Grid.Row="5" Grid.Column="0" Text="Total Paid:"/>
          <TextBlock Grid.Row="5" Grid.Column="1" Name="lblPaid"/>
          <TextBlock Grid.Row="6" Grid.Column="0" Text="Total Current Value:"/>
          <TextBlock Grid.Row="6" Grid.Column="1" Name="lblValue"/>
          <TextBlock Grid.Row="7" Grid.Column="0" Text="Gain / Loss:"/>
          <TextBlock Grid.Row="7" Grid.Column="1" Name="lblGL"/>
          <TextBlock Grid.Row="8" Grid.Column="0" Text="Earliest Year:"/>
          <TextBlock Grid.Row="8" Grid.Column="1" Name="lblEarliest"/>
          <TextBlock Grid.Row="9" Grid.Column="0" Text="Latest Year:"/>
          <TextBlock Grid.Row="9" Grid.Column="1" Name="lblLatest"/>
        </Grid>
        <TextBlock Text="DUPLICATES" FontSize="13" FontWeight="Bold" Foreground="#104861" Margin="0,14,0,3"/>
        <TextBlock Name="lblDupInfo" TextWrapping="Wrap"/>
        <TextBlock Text="BY DENOMINATION" FontSize="13" FontWeight="Bold" Foreground="#104861" Margin="0,14,0,3"/>
        <TextBlock Name="lblDenom" TextWrapping="Wrap"/>
        <TextBlock Text="BY MINT" FontSize="13" FontWeight="Bold" Foreground="#104861" Margin="0,14,0,3"/>
        <TextBlock Name="lblMint" TextWrapping="Wrap"/>
        <TextBlock Text="BY COUNTRY" FontSize="13" FontWeight="Bold" Foreground="#104861" Margin="0,14,0,3"/>
        <TextBlock Name="lblCountry" TextWrapping="Wrap"/>
        <TextBlock Text="BY METAL" FontSize="13" FontWeight="Bold" Foreground="#104861" Margin="0,14,0,3"/>
        <TextBlock Name="lblMetal" TextWrapping="Wrap"/>
        <TextBlock Text="BY GRADE" FontSize="13" FontWeight="Bold" Foreground="#104861" Margin="0,14,0,3"/>
        <TextBlock Name="lblGrade" TextWrapping="Wrap"/>
        <TextBlock Text="KEY DATES" FontSize="13" FontWeight="Bold" Foreground="#104861" Margin="0,14,0,3"/>
        <TextBlock Name="lblKeyDates" TextWrapping="Wrap"/>
        <TextBlock Text="CERTIFICATION" FontSize="13" FontWeight="Bold" Foreground="#104861" Margin="0,14,0,3"/>
        <TextBlock Name="lblSlabCount" TextWrapping="Wrap"/>
        <TextBlock Name="lblCertBreakdown" TextWrapping="Wrap"/>
      </StackPanel>
    </ScrollViewer>
</Window>
'@
    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $win    = [System.Windows.Markup.XamlReader]::Load($reader)
    $win.FindName('lblTotal').Text    = [string]$Coins.Count
    $win.FindName('lblInCollY').Text  = [string]$inCollY
    $win.FindName('lblInCollN').Text  = [string]$inCollN
    $win.FindName('lblFace').Text     = '$' + $totalFace
    $win.FindName('lblBook').Text     = '$' + $totalBook
    $win.FindName('lblPaid').Text     = '$' + $totalPurchase
    $win.FindName('lblValue').Text    = '$' + $totalCurrent
    $glSign = if ($gainLoss -ge 0) { '+$' } else { '-$' }
    $gl = $win.FindName('lblGL')
    $gl.Text       = $glSign + ([Math]::Abs($gainLoss).ToString('N2'))
    $gl.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($glColor)
    $win.FindName('lblEarliest').Text = [string]$earliest
    $win.FindName('lblLatest').Text   = [string]$latest
    $win.FindName('lblDupInfo').Text  = "  Total duplicates: $totalDups  |  Dup face value: `$$totalDupFV"
    $win.FindName('lblDenom').Text    = $byDenom
    $win.FindName('lblMint').Text     = $byMint
    $win.FindName('lblCountry').Text  = $byCountry
    $win.FindName('lblMetal').Text    = $byMetal
    $win.FindName('lblGrade').Text    = $byGrade
    $win.FindName('lblKeyDates').Text = $keyInfo
    $win.FindName('lblSlabCount').Text     = "Slabbed coins: $($slabbedCoins.Count) of $($Coins.Count)"
    $win.FindName('lblCertBreakdown').Text = $byCert
    $prevEAP = $ErrorActionPreference
    try { $ErrorActionPreference = 'Continue'; $null = $win.ShowDialog() }
    finally { $ErrorActionPreference = $prevEAP }
}

# ---------------------------------------------------------------------------
# Export Report
# ---------------------------------------------------------------------------
function Export-Report {
    param([object[]]$Coins)
    $sorted        = @($Coins | Sort-Object @{ Expression = { [int]$_.Year } }, Mint)
    $stamp         = Get-Date -Format 'yyyyMMdd_HHmmss'
    $reportFile    = Join-Path $PSScriptRoot "CoinReport_$stamp.txt"
    $totalPurchase = ($sorted | ForEach-Object { [double]$_.PurchasePrice } | Measure-Object -Sum).Sum
    $totalCurrent  = ($sorted | ForEach-Object { [double]$_.CurrentValue  } | Measure-Object -Sum).Sum
    $totalBook     = ($sorted | ForEach-Object { [double]$_.BookValue     } | Measure-Object -Sum).Sum
    $totalFace     = ($sorted | ForEach-Object { [double]$_.FaceValue     } | Measure-Object -Sum).Sum
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add('=' * 160)
    $lines.Add("  US COIN COLLECTION REPORT  -  Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $lines.Add("  Total Coins: $($sorted.Count)")
    $lines.Add('=' * 160)
    $lines.Add('')
    $lines.Add(("{0,-5} {1,-5} {2,-6} {3,-22} {4,-20} {5,-25} {6,-14} {7,8} {8,8} {9,8} {10,-8} {11,-10} {12,-6} {13,-10}" -f
        'ID','Year','Mint','Denomination','Type','Variety','Grade','Face$','Book$','Value$','Key','Metal','Cert','Slab'))
    $lines.Add('-' * 160)
    foreach ($c in $sorted) {
        $lines.Add(("{0,-5} {1,-5} {2,-6} {3,-22} {4,-20} {5,-25} {6,-14} {7,8:N2} {8,8:N2} {9,8:N2} {10,-8} {11,-10} {12,-6} {13,-10}" -f
            $c.ID, $c.Year, $c.Mint, $c.Denomination, $c.Type, $c.Variety, $c.Grade,
            ([double]$c.FaceValue), ([double]$c.BookValue), ([double]$c.CurrentValue),
            $c.KeyDates, $c.Metal, $c.CertCompany, $c.SlabGrade))
    }
    $lines.Add('-' * 160)
    $lines.Add(("  TOTAL FACE VALUE : {0,10:N2}" -f $totalFace))
    $lines.Add(("  TOTAL BOOK VALUE : {0,10:N2}" -f $totalBook))
    $lines.Add(("  TOTAL PURCHASE   : {0,10:N2}" -f $totalPurchase))
    $lines.Add(("  TOTAL VALUE      : {0,10:N2}" -f $totalCurrent))
    $lines.Add(("  GAIN / LOSS      : {0,10:N2}" -f ($totalCurrent - $totalPurchase)))
    $lines.Add('')
    $lines.Add('=' * 160)
    $lines | Set-Content -Path $reportFile -Encoding UTF8
    return $reportFile
}


# ===========================================================================
# MAIN WINDOW XAML  -  Dashboard Theme (matching Excel) + Left-Side Form
# ===========================================================================
[xml]$mainXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="United States Coin Collection Dashboard"
        Height="850" Width="1500"
        MinHeight="600" MinWidth="1100"
        WindowStartupLocation="CenterScreen"
        Background="#E8ECF0">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Height"      Value="32"/>
            <Setter Property="FontSize"    Value="12"/>
            <Setter Property="FontWeight"  Value="SemiBold"/>
            <Setter Property="Cursor"      Value="Hand"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Margin"      Value="3,0"/>
            <Setter Property="Padding"     Value="10,0"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background"    Value="White"/>
            <Setter Property="Foreground"    Value="#1A1A2E"/>
            <Setter Property="BorderBrush"   Value="#B0B8C4"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding"       Value="5,3"/>
            <Setter Property="FontSize"      Value="12"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        <Style TargetType="DataGrid">
            <Setter Property="Background"              Value="White"/>
            <Setter Property="Foreground"              Value="#1A1A2E"/>
            <Setter Property="BorderBrush"             Value="#B0B8C4"/>
            <Setter Property="BorderThickness"         Value="1"/>
            <Setter Property="RowBackground"           Value="White"/>
            <Setter Property="AlternatingRowBackground" Value="#F2F2F2"/>
            <Setter Property="GridLinesVisibility"     Value="Horizontal"/>
            <Setter Property="HorizontalGridLinesBrush" Value="#D9D9D9"/>
            <Setter Property="FontSize"                Value="11"/>
            <Setter Property="FontFamily"              Value="Segoe UI"/>
            <Setter Property="CanUserAddRows"          Value="False"/>
            <Setter Property="CanUserDeleteRows"       Value="False"/>
            <Setter Property="IsReadOnly"              Value="True"/>
            <Setter Property="SelectionMode"           Value="Single"/>
            <Setter Property="SelectionUnit"           Value="FullRow"/>
            <Setter Property="AutoGenerateColumns"     Value="False"/>
            <Setter Property="HeadersVisibility"       Value="Column"/>
        </Style>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background"   Value="#104861"/>
            <Setter Property="Foreground"   Value="White"/>
            <Setter Property="FontWeight"   Value="Bold"/>
            <Setter Property="FontSize"     Value="11"/>
            <Setter Property="Padding"      Value="6,5"/>
            <Setter Property="BorderBrush"  Value="#0D3A4D"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
        </Style>
        <Style TargetType="DataGridRow">
            <Setter Property="Cursor" Value="Hand"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#C8E0EC"/>
                    <Setter Property="Foreground" Value="#104861"/>
                </Trigger>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#E0EEF5"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="DataGridCell">
            <Setter Property="Padding"        Value="3,1"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="Transparent"/>
                    <Setter Property="BorderBrush" Value="Transparent"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="Label">
            <Setter Property="Foreground" Value="#1A1A2E"/>
            <Setter Property="FontSize"   Value="11"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Background"  Value="White"/>
            <Setter Property="Foreground"  Value="#1A1A2E"/>
            <Setter Property="BorderBrush" Value="#B0B8C4"/>
            <Setter Property="FontSize"    Value="12"/>
            <Setter Property="Height"      Value="26"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
    </Window.Resources>

    <DockPanel>
        <!-- ===== TITLE BANNER (dark teal, matching Excel Dashboard) ===== -->
        <Border DockPanel.Dock="Top" Padding="18,12">
            <Border.Background>
                <LinearGradientBrush StartPoint="0,0" EndPoint="1,0">
                    <GradientStop Color="#0D3A4D" Offset="0"/>
                    <GradientStop Color="#104861" Offset="0.5"/>
                    <GradientStop Color="#186080" Offset="1"/>
                </LinearGradientBrush>
            </Border.Background>
            <DockPanel>
                <TextBlock Name="lblCoinCount" DockPanel.Dock="Right"
                           FontSize="13" Foreground="#80C8E0"
                           HorizontalAlignment="Right" VerticalAlignment="Center"
                           Margin="0,0,10,0" Text="0 coins"/>
                <StackPanel DockPanel.Dock="Right" Orientation="Horizontal"
                            Margin="0,0,20,0" VerticalAlignment="Center">
                    <TextBlock Text="Collection:" Foreground="#A0D0E0" FontSize="11"
                               VerticalAlignment="Center" Margin="0,0,5,0"/>
                    <ComboBox Name="cbCollections" Width="150" Height="26" IsEditable="False"
                              FontSize="11"/>
                    <Button Name="btnNewCollection" Content="+" Width="26" Height="26"
                            Background="#2E8B57" Foreground="White" FontSize="13" FontWeight="Bold"
                            Margin="3,0,0,0" Padding="0"/>
                    <Button Name="btnDelCollection" Content="-" Width="26" Height="26"
                            Background="#C62828" Foreground="White" FontSize="13" FontWeight="Bold"
                            Margin="3,0,0,0" Padding="0"/>
                </StackPanel>
                <StackPanel>
                    <TextBlock Text="United States Coin Collection Dashboard"
                               FontSize="19" FontWeight="Bold" Foreground="White"/>
                    <TextBlock Text="Circulated Condition" FontSize="11" Foreground="#80C8E0"
                               Margin="0,1,0,0"/>
                </StackPanel>
            </DockPanel>
        </Border>

        <!-- ===== STATUS BAR ===== -->
        <Border DockPanel.Dock="Bottom" Background="#104861" Padding="10,4">
            <DockPanel>
                <TextBlock DockPanel.Dock="Right" Text="v4.0.0" Foreground="#5A9AB5" FontSize="11"
                           VerticalAlignment="Center" Margin="10,0,4,0"/>
                <TextBlock DockPanel.Dock="Right" Foreground="#5A9AB5" FontSize="11"
                           VerticalAlignment="Center"
                           Text="c2026 US Coin Inventory System"/>
                <Button Name="btnUndo" Content="↩ Undo Delete" Visibility="Collapsed"
                        Background="#BF360C" Foreground="White" Margin="0,0,12,0"
                        FontSize="11" Padding="6,0" BorderThickness="0"/>
                <TextBlock Name="lblStatus" Foreground="#80C8E0" FontSize="12" Text="Ready"/>
            </DockPanel>
        </Border>

        <!-- ===== MAIN CONTENT: LEFT FORM + RIGHT DASHBOARD ===== -->
        <Grid>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="335"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- ===== LEFT: COIN FORM PANEL (Always Visible) ===== -->
            <Border Grid.Column="0" Background="#F0F2F5" BorderBrush="#B0B8C4" BorderThickness="0,0,1,0">
                <DockPanel>
                    <!-- Form Header -->
                    <Border DockPanel.Dock="Top" Padding="12,8">
                        <Border.Background>
                            <LinearGradientBrush StartPoint="0,0" EndPoint="1,0">
                                <GradientStop Color="#0D3A4D" Offset="0"/>
                                <GradientStop Color="#104861" Offset="1"/>
                            </LinearGradientBrush>
                        </Border.Background>
                        <DockPanel>
                            <TextBlock Name="lblFormTitle" Text="Add New Coin"
                                       Foreground="White" FontSize="14" FontWeight="Bold"
                                       VerticalAlignment="Center"/>
                        </DockPanel>
                    </Border>
                    <!-- Bottom: Save / New buttons -->
                    <Border DockPanel.Dock="Bottom" Background="#D9DFE3" Padding="8,6">
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                            <Button Name="btnClearForm" Content="Clear / New" Width="95" Height="30"
                                    Background="#78909C" Foreground="White" Margin="0,0,6,0"/>
                            <Button Name="btnSaveCoin" Content="Save Coin" Width="100" Height="30"
                                    Background="#2E8B57" Foreground="White"/>
                        </StackPanel>
                    </Border>
                    <!-- Scrollable Form Fields -->
                    <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="8,4">
                        <Grid Margin="2">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="92"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>

                            <!-- Row 0: Section - Coin Info -->
                            <TextBlock Grid.Row="0" Grid.ColumnSpan="2" Text="Coin Information"
                                       Foreground="#104861" FontSize="11" FontWeight="Bold" Margin="0,4,0,3"/>
                            <Label Grid.Row="1" Grid.Column="0" Content="Year *" FontSize="11"/>
                            <TextBox Grid.Row="1" Grid.Column="1" Name="tbYear" MaxLength="4" Height="24" FontSize="11"/>
                            <Label Grid.Row="2" Grid.Column="0" Content="Mint Mark" FontSize="11"/>
                            <ComboBox Grid.Row="2" Grid.Column="1" Name="cbMint" IsEditable="True" Height="24" FontSize="11">
                                <ComboBoxItem Content=""/><ComboBoxItem Content="P"/><ComboBoxItem Content="D"/>
                                <ComboBoxItem Content="S"/><ComboBoxItem Content="W"/><ComboBoxItem Content="CC"/>
                                <ComboBoxItem Content="O"/><ComboBoxItem Content="C"/><ComboBoxItem Content="Plain"/>
                            </ComboBox>
                            <Label Grid.Row="3" Grid.Column="0" Content="Denom *" FontSize="11"/>
                            <ComboBox Grid.Row="3" Grid.Column="1" Name="cbDenom" IsEditable="True" Height="24" FontSize="11">
                                <ComboBoxItem Content="American Eagle Dollar"/><ComboBoxItem Content="Barber Dime"/>
                                <ComboBoxItem Content="Barber Half Dollar"/><ComboBoxItem Content="Barber Quarter"/>
                                <ComboBoxItem Content="Buffalo Nickel"/><ComboBoxItem Content="Bust Dime"/>
                                <ComboBoxItem Content="Capped Bust Half Dollar"/><ComboBoxItem Content="Capped Bust Quarter"/>
                                <ComboBoxItem Content="Commemorative"/><ComboBoxItem Content="Commemorative Half Dollar"/>
                                <ComboBoxItem Content="Draped Bust Dollar"/><ComboBoxItem Content="Draped Bust Half Dollar"/>
                                <ComboBoxItem Content="Draped Bust Quarters"/><ComboBoxItem Content="Early Dime"/>
                                <ComboBoxItem Content="Early Half Dime"/><ComboBoxItem Content="Eisenhower Dollar"/>
                                <ComboBoxItem Content="Flowing Hair Dollar"/><ComboBoxItem Content="Flowing Hair Half Dollar"/>
                                <ComboBoxItem Content="Flying Eagle Cent"/><ComboBoxItem Content="Franklin Half Dollar"/>
                                <ComboBoxItem Content="Gold ($1.00)"/><ComboBoxItem Content="Gold ($2.50)"/>
                                <ComboBoxItem Content="Gold ($3.00)"/><ComboBoxItem Content="Gold Double Eagle ($20.00)"/>
                                <ComboBoxItem Content="Gold Eagle ($10.00)"/><ComboBoxItem Content="Gold Half Eagle ($5.00)"/>
                                <ComboBoxItem Content="Gold Quarter Eagle ($2.50)"/><ComboBoxItem Content="Half Cent"/>
                                <ComboBoxItem Content="Indian Head Cent"/><ComboBoxItem Content="Jefferson Nickel"/>
                                <ComboBoxItem Content="Kennedy Half Dollar"/><ComboBoxItem Content="Large Cent"/>
                                <ComboBoxItem Content="Liberty Nickel"/><ComboBoxItem Content="Liberty Seated Half Dime"/>
                                <ComboBoxItem Content="Lincoln Cent"/><ComboBoxItem Content="Mercury Dime"/>
                                <ComboBoxItem Content="Mint Set"/><ComboBoxItem Content="Morgan Dollar"/>
                                <ComboBoxItem Content="Peace Dollar"/><ComboBoxItem Content="Presidential Dollar"/>
                                <ComboBoxItem Content="Proof Set"/><ComboBoxItem Content="Roosevelt Dime"/>
                                <ComboBoxItem Content="Sacagawea Dollar"/><ComboBoxItem Content="Seated Liberty Dime"/>
                                <ComboBoxItem Content="Seated Liberty Dollar"/><ComboBoxItem Content="Seated Liberty Half Dollar"/>
                                <ComboBoxItem Content="Seated Liberty Quarter"/><ComboBoxItem Content="Shield Nickel"/>
                                <ComboBoxItem Content="Standing Liberty Quarter"/><ComboBoxItem Content="Susan B. Anthony Dollar"/>
                                <ComboBoxItem Content="Three Cent Piece"/><ComboBoxItem Content="Trade Dollar"/>
                                <ComboBoxItem Content="Twenty Cent Piece"/><ComboBoxItem Content="Two Cent Piece"/>
                                <ComboBoxItem Content="Walking Liberty Half Dollar"/><ComboBoxItem Content="Washington Quarter"/>
                            </ComboBox>
                            <Label Grid.Row="4" Grid.Column="0" Content="Type" FontSize="11"/>
                            <TextBox Grid.Row="4" Grid.Column="1" Name="tbType" Height="24" FontSize="11"/>
                            <Label Grid.Row="5" Grid.Column="0" Content="Variety" FontSize="11"/>
                            <TextBox Grid.Row="5" Grid.Column="1" Name="tbVariety" Height="24" FontSize="11"/>
                            <Label Grid.Row="6" Grid.Column="0" Content="Country" FontSize="11"/>
                            <ComboBox Grid.Row="6" Grid.Column="1" Name="cbCountry" IsEditable="True" Height="24" FontSize="11">
                                <ComboBoxItem Content="USA"/><ComboBoxItem Content="Canada"/><ComboBoxItem Content="Mexico"/>
                                <ComboBoxItem Content="UK"/><ComboBoxItem Content="Australia"/><ComboBoxItem Content="Other"/>
                            </ComboBox>
                            <Label Grid.Row="7" Grid.Column="0" Content="Series" FontSize="11"/>
                            <TextBox Grid.Row="7" Grid.Column="1" Name="tbSeries" Height="24" FontSize="11"/>
                            <Label Grid.Row="8" Grid.Column="0" Content="Error" FontSize="11"/>
                            <TextBox Grid.Row="8" Grid.Column="1" Name="tbError" Height="24" FontSize="11"/>

                            <!-- Row 9: Section - Grading & Value -->
                            <TextBlock Grid.Row="9" Grid.ColumnSpan="2" Text="Grading &amp; Value"
                                       Foreground="#104861" FontSize="11" FontWeight="Bold" Margin="0,8,0,3"/>
                            <Label Grid.Row="10" Grid.Column="0" Content="Grade" FontSize="11"/>
                            <ComboBox Grid.Row="10" Grid.Column="1" Name="cbGrade" IsEditable="True" Height="24" FontSize="11">
                                <ComboBoxItem Content=""/><ComboBoxItem Content="Poor"/><ComboBoxItem Content="Fair"/>
                                <ComboBoxItem Content="About Good"/><ComboBoxItem Content="Good"/>
                                <ComboBoxItem Content="Very Good"/><ComboBoxItem Content="Fine"/>
                                <ComboBoxItem Content="Very Fine"/><ComboBoxItem Content="Extremely Fine"/>
                                <ComboBoxItem Content="About Uncirculated"/><ComboBoxItem Content="Uncirculated"/>
                                <ComboBoxItem Content="Mint State"/><ComboBoxItem Content="Proof"/>
                                <ComboBoxItem Content="Poor (P-1)"/><ComboBoxItem Content="Good (G-4)"/>
                                <ComboBoxItem Content="Very Good (VG-8)"/><ComboBoxItem Content="Fine (F-12)"/>
                                <ComboBoxItem Content="Very Fine (VF-20)"/><ComboBoxItem Content="Extremely Fine (EF-40)"/>
                                <ComboBoxItem Content="About Uncirculated (AU-50)"/>
                                <ComboBoxItem Content="Mint State (MS-60)"/><ComboBoxItem Content="Mint State (MS-63)"/>
                                <ComboBoxItem Content="Mint State (MS-65)"/><ComboBoxItem Content="Mint State (MS-70)"/>
                                <ComboBoxItem Content="Proof (PR-65)"/><ComboBoxItem Content="Proof (PR-70)"/>
                            </ComboBox>
                            <Label Grid.Row="11" Grid.Column="0" Content="Face Value $" FontSize="11"/>
                            <TextBox Grid.Row="11" Grid.Column="1" Name="tbFaceValue" Text="0" Height="24" FontSize="11"/>
                            <Label Grid.Row="12" Grid.Column="0" Content="Book Value $" FontSize="11"/>
                            <TextBox Grid.Row="12" Grid.Column="1" Name="tbBookValue" Text="0" Height="24" FontSize="11"/>
                            <Label Grid.Row="13" Grid.Column="0" Content="Purchase $" FontSize="11"/>
                            <TextBox Grid.Row="13" Grid.Column="1" Name="tbPurchasePrice" Text="0" Height="24" FontSize="11"/>
                            <Label Grid.Row="14" Grid.Column="0" Content="Current Val $" FontSize="11"/>
                            <TextBox Grid.Row="14" Grid.Column="1" Name="tbCurrentValue" Text="0" Height="24" FontSize="11"/>
                            <Label Grid.Row="15" Grid.Column="0" Content="Key Dates" FontSize="11"/>
                            <ComboBox Grid.Row="15" Grid.Column="1" Name="cbKeyDates" IsEditable="True" Height="24" FontSize="11">
                                <ComboBoxItem Content=""/><ComboBoxItem Content="KeyDate"/>
                                <ComboBoxItem Content="SemiKey"/><ComboBoxItem Content="BetterDate"/>
                            </ComboBox>

                            <!-- Row 16: Section - Physical -->
                            <TextBlock Grid.Row="16" Grid.ColumnSpan="2" Text="Physical Properties"
                                       Foreground="#104861" FontSize="11" FontWeight="Bold" Margin="0,8,0,3"/>
                            <Label Grid.Row="17" Grid.Column="0" Content="Metal" FontSize="11"/>
                            <ComboBox Grid.Row="17" Grid.Column="1" Name="cbMetal" IsEditable="True" Height="24" FontSize="11">
                                <ComboBoxItem Content=""/><ComboBoxItem Content="copper"/><ComboBoxItem Content="copper/nickel"/>
                                <ComboBoxItem Content="zinc/copper"/><ComboBoxItem Content="Nickel"/><ComboBoxItem Content="Bronze"/>
                                <ComboBoxItem Content="Brass"/><ComboBoxItem Content="Steel"/><ComboBoxItem Content="Aluminum"/>
                                <ComboBoxItem Content="Silver"/><ComboBoxItem Content="35% Silver"/><ComboBoxItem Content="40% Silver"/>
                                <ComboBoxItem Content="90% Silver"/><ComboBoxItem Content="99.9% Silver"/>
                                <ComboBoxItem Content="Gold"/><ComboBoxItem Content="90% Gold"/>
                            </ComboBox>
                            <Label Grid.Row="18" Grid.Column="0" Content="Weight" FontSize="11"/>
                            <TextBox Grid.Row="18" Grid.Column="1" Name="tbWeight" Height="24" FontSize="11"/>

                            <!-- Row 19: Section - Collection Status -->
                            <TextBlock Grid.Row="19" Grid.ColumnSpan="2" Text="Collection Status"
                                       Foreground="#104861" FontSize="11" FontWeight="Bold" Margin="0,8,0,3"/>
                            <Label Grid.Row="20" Grid.Column="0" Content="In Collection" FontSize="11"/>
                            <ComboBox Grid.Row="20" Grid.Column="1" Name="cbInCollection" IsEditable="False" Height="24" FontSize="11">
                                <ComboBoxItem Content="Y"/><ComboBoxItem Content="N"/>
                            </ComboBox>
                            <Label Grid.Row="21" Grid.Column="0" Content="Location" FontSize="11"/>
                            <TextBox Grid.Row="21" Grid.Column="1" Name="tbLocation" Height="24" FontSize="11"/>
                            <Label Grid.Row="22" Grid.Column="0" Content="Date Added" FontSize="11"/>
                            <TextBox Grid.Row="22" Grid.Column="1" Name="tbDateAdded" IsReadOnly="True"
                                     Foreground="#888" Height="24" FontSize="11"/>
                            <Label Grid.Row="23" Grid.Column="0" Content="Date Range" FontSize="11"/>
                            <TextBox Grid.Row="23" Grid.Column="1" Name="tbDateRangeMinted" Height="24" FontSize="11"/>
                            <Label Grid.Row="24" Grid.Column="0" Content="Data Std" FontSize="11"/>
                            <ComboBox Grid.Row="24" Grid.Column="1" Name="cbDataStandard" IsEditable="True" Height="24" FontSize="11">
                                <ComboBoxItem Content=""/><ComboBoxItem Content="OK"/>
                                <ComboBoxItem Content="Missing Data"/><ComboBoxItem Content="Record Disabled"/>
                            </ComboBox>

                            <!-- Row 25: Section - Duplicates -->
                            <TextBlock Grid.Row="25" Grid.ColumnSpan="2" Text="Duplicates"
                                       Foreground="#104861" FontSize="11" FontWeight="Bold" Margin="0,8,0,3"/>
                            <Label Grid.Row="26" Grid.Column="0" Content="Dup Count" FontSize="11"/>
                            <TextBox Grid.Row="26" Grid.Column="1" Name="tbDupCount" Text="0" Height="24" FontSize="11"/>
                            <Label Grid.Row="27" Grid.Column="0" Content="Dup Face $" FontSize="11"/>
                            <TextBox Grid.Row="27" Grid.Column="1" Name="tbDupFaceValue" Text="0" Height="24" FontSize="11"/>
                            <Label Grid.Row="28" Grid.Column="0" Content="Dup Grade $" FontSize="11"/>
                            <TextBox Grid.Row="28" Grid.Column="1" Name="tbDupGradedValue" Text="0" Height="24" FontSize="11"/>
                            <Label Grid.Row="29" Grid.Column="0" Content="Dup Location" FontSize="11"/>
                            <TextBox Grid.Row="29" Grid.Column="1" Name="tbDupLoc" Height="24" FontSize="11"/>

                            <!-- Row 30: Section - Certification -->
                            <TextBlock Grid.Row="30" Grid.ColumnSpan="2" Text="Slab / Certification"
                                       Foreground="#104861" FontSize="11" FontWeight="Bold" Margin="0,8,0,3"/>
                            <Label Grid.Row="31" Grid.Column="0" Content="Cert Company" FontSize="11"/>
                            <ComboBox Grid.Row="31" Grid.Column="1" Name="cbCertCompany" IsEditable="False" Height="24" FontSize="11">
                                <ComboBoxItem Content=""/><ComboBoxItem Content="PCGS"/>
                                <ComboBoxItem Content="NGC"/><ComboBoxItem Content="ANACS"/>
                                <ComboBoxItem Content="ICG"/><ComboBoxItem Content="Other"/>
                            </ComboBox>
                            <Label Grid.Row="32" Grid.Column="0" Content="Cert Number" FontSize="11"/>
                            <DockPanel Grid.Row="32" Grid.Column="1">
                                <Button DockPanel.Dock="Right" Name="btnVerify" Content="Verify"
                                        Width="50" Height="24" Margin="3,0,0,0"
                                        Background="#104861" Foreground="White" FontSize="10"/>
                                <TextBox Name="tbCertNumber" Height="24" FontSize="11"/>
                            </DockPanel>
                            <Label Grid.Row="33" Grid.Column="0" Content="Slab Grade" FontSize="11"/>
                            <TextBox Grid.Row="33" Grid.Column="1" Name="tbSlabGrade" Height="24" FontSize="11"/>
                            <Label Grid.Row="34" Grid.Column="0" Content="Mintage" FontSize="11"/>
                            <TextBox Grid.Row="34" Grid.Column="1" Name="tbMintage" Height="24" FontSize="11"/>
                            <Label Grid.Row="35" Grid.Column="0" Content="Rarity (1-10)" FontSize="11"/>
                            <ComboBox Grid.Row="35" Grid.Column="1" Name="cbRarityScore" IsEditable="False" Height="24" FontSize="11">
                                <ComboBoxItem Content=""/><ComboBoxItem Content="1"/><ComboBoxItem Content="2"/>
                                <ComboBoxItem Content="3"/><ComboBoxItem Content="4"/><ComboBoxItem Content="5"/>
                                <ComboBoxItem Content="6"/><ComboBoxItem Content="7"/><ComboBoxItem Content="8"/>
                                <ComboBoxItem Content="9"/><ComboBoxItem Content="10"/>
                            </ComboBox>

                            <!-- Row 36: Section - Notes -->
                            <TextBlock Grid.Row="36" Grid.ColumnSpan="2" Text="Notes"
                                       Foreground="#104861" FontSize="11" FontWeight="Bold" Margin="0,8,0,3"/>
                            <Label Grid.Row="37" Grid.Column="0" Content="Comments" FontSize="11"/>
                            <TextBox Grid.Row="37" Grid.Column="1" Name="tbComments" Height="24" FontSize="11"/>
                            <Label Grid.Row="38" Grid.Column="0" Content="Notes" FontSize="11"/>
                            <TextBox Grid.Row="38" Grid.Column="1" Name="tbNotes" Height="24" FontSize="11"/>

                            <!-- Row 39: Section - Varieties & Errors -->
                            <TextBlock Grid.Row="39" Grid.ColumnSpan="2" Text="Varieties &amp; Errors"
                                       Foreground="#104861" FontSize="11" FontWeight="Bold" Margin="0,8,0,3"/>
                            <Label Grid.Row="40" Grid.Column="0" Content="VAM #" FontSize="11"/>
                            <TextBox Grid.Row="40" Grid.Column="1" Name="tbVAMNumber" Height="24" FontSize="11"
                                     ToolTip="Van Allen/Mallis variety number (Morgan/Peace dollars)"/>
                            <Label Grid.Row="41" Grid.Column="0" Content="FS #" FontSize="11"/>
                            <TextBox Grid.Row="41" Grid.Column="1" Name="tbFSNumber" Height="24" FontSize="11"
                                     ToolTip="Fivaz-Stanton number from Cherrypicker's Guide"/>
                            <Label Grid.Row="42" Grid.Column="0" Content="Error Type" FontSize="11"/>
                            <ComboBox Grid.Row="42" Grid.Column="1" Name="cbErrorType" IsEditable="True" Height="24" FontSize="11">
                                <ComboBoxItem Content=""/><ComboBoxItem Content="DDO (Doubled Die Obverse)"/>
                                <ComboBoxItem Content="DDR (Doubled Die Reverse)"/>
                                <ComboBoxItem Content="RPM (Repunched Mint Mark)"/>
                                <ComboBoxItem Content="Off-Center Strike"/>
                                <ComboBoxItem Content="Broadstrike"/><ComboBoxItem Content="Clipped Planchet"/>
                                <ComboBoxItem Content="Die Crack"/><ComboBoxItem Content="Die Break"/>
                                <ComboBoxItem Content="Cud Error"/><ComboBoxItem Content="Wrong Planchet"/>
                                <ComboBoxItem Content="Double Struck"/><ComboBoxItem Content="Brockage"/>
                                <ComboBoxItem Content="Lamination Error"/><ComboBoxItem Content="Missing Edge Lettering"/>
                            </ComboBox>
                            <Label Grid.Row="43" Grid.Column="0" Content="Variety" FontSize="11"/>
                            <ComboBox Grid.Row="43" Grid.Column="1" Name="cbVarietyType" IsEditable="True" Height="24" FontSize="11">
                                <ComboBoxItem Content=""/><ComboBoxItem Content="Overdate"/>
                                <ComboBoxItem Content="Over Mint Mark"/><ComboBoxItem Content="Large Date"/>
                                <ComboBoxItem Content="Small Date"/><ComboBoxItem Content="Large Letters"/>
                                <ComboBoxItem Content="Small Letters"/><ComboBoxItem Content="Proof"/>
                                <ComboBoxItem Content="Deep Cameo"/><ComboBoxItem Content="Full Bands"/>
                                <ComboBoxItem Content="Full Bell Lines"/><ComboBoxItem Content="Full Steps"/>
                                <ComboBoxItem Content="Full Head"/><ComboBoxItem Content="Full Torch"/>
                                <ComboBoxItem Content="CAC Verified"/>
                            </ComboBox>

                            <!-- Row 44: Section - Image -->
                            <TextBlock Grid.Row="44" Grid.ColumnSpan="2" Text="Coin Image"
                                       Foreground="#104861" FontSize="11" FontWeight="Bold" Margin="0,8,0,3"/>
                            <Label Grid.Row="45" Grid.Column="0" Content="Image" FontSize="11"/>
                            <DockPanel Grid.Row="45" Grid.Column="1">
                                <Button DockPanel.Dock="Right" Name="btnBrowseImage" Content="..."
                                        Width="32" Height="24" Margin="3,0,0,0"
                                        Background="#104861" Foreground="White" FontSize="10"/>
                                <TextBox Name="tbImagePath" IsReadOnly="True" Foreground="#888" Height="24" FontSize="10"/>
                            </DockPanel>
                            <Image Grid.Row="46" Grid.Column="1" Name="imgPreview"
                                   Width="80" Height="80" Stretch="Uniform"
                                   HorizontalAlignment="Left" Margin="0,4,0,0"
                                   AllowDrop="True"
                                   ToolTip="Click to enlarge · Right-click to clear · Drag &amp; drop image here"/>
                            <TextBlock Grid.Row="47" Grid.ColumnSpan="2" Name="lblFormError"
                                       Foreground="#C62828" FontSize="11" TextWrapping="Wrap"
                                       Margin="0,4,0,0" Visibility="Collapsed"/>
                        </Grid>
                    </ScrollViewer>
                </DockPanel>
            </Border>

            <!-- ===== RIGHT: DASHBOARD CONTENT ===== -->
            <DockPanel Grid.Column="1">
                <!-- Dashboard Stats Cards -->
                <Border DockPanel.Dock="Top" Background="#D9DFE3" Padding="6,8">
                    <WrapPanel HorizontalAlignment="Center">
                        <Border Name="cardHolding" Background="White" CornerRadius="5" Padding="14,6" Margin="4,2" BorderBrush="#B0B8C4" BorderThickness="1" Cursor="Hand">
                            <StackPanel HorizontalAlignment="Center" Width="80">
                                <TextBlock Name="statHolding" Text="0" FontSize="22" FontWeight="Bold" Foreground="#104861" HorizontalAlignment="Center"/>
                                <TextBlock Text="Holding" FontSize="10" Foreground="#666" HorizontalAlignment="Center"/>
                                <TextBlock Text="click to print" FontSize="7" Foreground="#AAA" HorizontalAlignment="Center"/>
                            </StackPanel>
                        </Border>
                        <Border Name="cardMissing" Background="White" CornerRadius="5" Padding="14,6" Margin="4,2" BorderBrush="#B0B8C4" BorderThickness="1" Cursor="Hand">
                            <StackPanel HorizontalAlignment="Center" Width="80">
                                <TextBlock Name="statMissing" Text="0" FontSize="22" FontWeight="Bold" Foreground="#C62828" HorizontalAlignment="Center"/>
                                <TextBlock Text="Missing" FontSize="10" Foreground="#666" HorizontalAlignment="Center"/>
                                <TextBlock Text="click to print" FontSize="7" Foreground="#AAA" HorizontalAlignment="Center"/>
                            </StackPanel>
                        </Border>
                        <Border Name="cardMinted" Background="White" CornerRadius="5" Padding="14,6" Margin="4,2" BorderBrush="#B0B8C4" BorderThickness="1" Cursor="Hand">
                            <StackPanel HorizontalAlignment="Center" Width="80">
                                <TextBlock Name="statMinted" Text="0" FontSize="22" FontWeight="Bold" Foreground="#333" HorizontalAlignment="Center"/>
                                <TextBlock Text="Minted" FontSize="10" Foreground="#666" HorizontalAlignment="Center"/>
                                <TextBlock Text="click to print" FontSize="7" Foreground="#AAA" HorizontalAlignment="Center"/>
                            </StackPanel>
                        </Border>
                        <Border Name="cardHoldPct" Background="White" CornerRadius="5" Padding="14,6" Margin="4,2" BorderBrush="#B0B8C4" BorderThickness="1" Cursor="Hand">
                            <StackPanel HorizontalAlignment="Center" Width="80">
                                <TextBlock Name="statHoldPct" Text="0%" FontSize="22" FontWeight="Bold" Foreground="#2E7D32" HorizontalAlignment="Center"/>
                                <TextBlock Text="Hold %" FontSize="10" Foreground="#666" HorizontalAlignment="Center"/>
                                <TextBlock Text="click to print" FontSize="7" Foreground="#AAA" HorizontalAlignment="Center"/>
                            </StackPanel>
                        </Border>
                        <Border Name="cardBookVal" Background="White" CornerRadius="5" Padding="14,6" Margin="4,2" BorderBrush="#B0B8C4" BorderThickness="1" Cursor="Hand">
                            <StackPanel HorizontalAlignment="Center" Width="90">
                                <TextBlock Name="statBookVal" Text="$0" FontSize="20" FontWeight="Bold" Foreground="#104861" HorizontalAlignment="Center"/>
                                <TextBlock Text="Book Value" FontSize="10" Foreground="#666" HorizontalAlignment="Center"/>
                                <TextBlock Text="click to print" FontSize="7" Foreground="#AAA" HorizontalAlignment="Center"/>
                            </StackPanel>
                        </Border>
                        <Border Name="cardFaceVal" Background="White" CornerRadius="5" Padding="14,6" Margin="4,2" BorderBrush="#B0B8C4" BorderThickness="1" Cursor="Hand">
                            <StackPanel HorizontalAlignment="Center" Width="90">
                                <TextBlock Name="statFaceVal" Text="$0" FontSize="20" FontWeight="Bold" Foreground="#104861" HorizontalAlignment="Center"/>
                                <TextBlock Text="Face Value" FontSize="10" Foreground="#666" HorizontalAlignment="Center"/>
                                <TextBlock Text="click to print" FontSize="7" Foreground="#AAA" HorizontalAlignment="Center"/>
                            </StackPanel>
                        </Border>
                        <Border Name="cardGraded" Background="White" CornerRadius="5" Padding="14,6" Margin="4,2" BorderBrush="#B0B8C4" BorderThickness="1" Cursor="Hand">
                            <StackPanel HorizontalAlignment="Center" Width="80">
                                <TextBlock Name="statGraded" Text="0" FontSize="22" FontWeight="Bold" Foreground="#6A1B9A" HorizontalAlignment="Center"/>
                                <TextBlock Text="Graded" FontSize="10" Foreground="#666" HorizontalAlignment="Center"/>
                                <TextBlock Text="click to print" FontSize="7" Foreground="#AAA" HorizontalAlignment="Center"/>
                            </StackPanel>
                        </Border>
                        <Border Name="cardKeys" Background="White" CornerRadius="5" Padding="14,6" Margin="4,2" BorderBrush="#B0B8C4" BorderThickness="1" Cursor="Hand">
                            <StackPanel HorizontalAlignment="Center" Width="80">
                                <TextBlock Name="statKeys" Text="0" FontSize="22" FontWeight="Bold" Foreground="#E65100" HorizontalAlignment="Center"/>
                                <TextBlock Text="Key Dates" FontSize="10" Foreground="#666" HorizontalAlignment="Center"/>
                                <TextBlock Text="click to print" FontSize="7" Foreground="#AAA" HorizontalAlignment="Center"/>
                            </StackPanel>
                        </Border>
                        <Border Name="cardDups" Background="White" CornerRadius="5" Padding="14,6" Margin="4,2" BorderBrush="#B0B8C4" BorderThickness="1" Cursor="Hand">
                            <StackPanel HorizontalAlignment="Center" Width="80">
                                <TextBlock Name="statDups" Text="0" FontSize="22" FontWeight="Bold" Foreground="#795548" HorizontalAlignment="Center"/>
                                <TextBlock Text="Duplicates" FontSize="10" Foreground="#666" HorizontalAlignment="Center"/>
                                <TextBlock Text="click to print" FontSize="7" Foreground="#AAA" HorizontalAlignment="Center"/>
                            </StackPanel>
                        </Border>
                    </WrapPanel>
                </Border>

                <!-- Toolbar -->
                <Border DockPanel.Dock="Top" Padding="8,5">
                    <Border.Background>
                        <LinearGradientBrush StartPoint="0,0" EndPoint="0,1">
                            <GradientStop Color="#1C4A5E" Offset="0"/>
                            <GradientStop Color="#104861" Offset="1"/>
                        </LinearGradientBrush>
                    </Border.Background>
                    <WrapPanel>
                        <Button Name="btnAdd"      Content="+ New Coin"
                                Background="#2E8B57" Foreground="White"/>
                        <Button Name="btnDelete"   Content="Delete"
                                Background="#C62828" Foreground="White"/>
                        <Separator Width="12" Background="Transparent"/>
                        <Button Name="btnStats"    Content="Statistics"
                                Background="#E65100" Foreground="White"/>
                        <Button Name="btnExport"   Content="Export Report"
                                Background="#6A1B9A" Foreground="White"/>
                        <Button Name="btnRefresh"  Content="Refresh"
                                Background="#37474F" Foreground="White"/>
                        <Separator Width="12" Background="Transparent"/>
                        <Button Name="btnBackup"   Content="Backup"
                                Background="#F9A825" Foreground="#1A1A2E"/>
                        <Button Name="btnImport"   Content="Import CSV"
                                Background="#00897B" Foreground="White"/>
                        <Button Name="btnImportExcel" Content="Import Excel"
                                Background="#1565C0" Foreground="White"/>
                        <Separator Width="12" Background="Transparent"/>
                        <Button Name="btnClone"      Content="Clone Coin"
                                Background="#00695C" Foreground="White"/>
                        <Button Name="btnExportCsv"  Content="Export CSV"
                                Background="#558B2F" Foreground="White"/>
                        <Button Name="btnRestore"    Content="Restore Backup"
                                Background="#4527A0" Foreground="White"/>
                        <Separator Width="12" Background="Transparent"/>
                        <Button Name="btnResearch"   Content="&#x1F50D; Research"
                                Background="#0277BD" Foreground="White"/>
                        <Button Name="btnMeltCalc"   Content="&#x2696; Melt Value"
                                Background="#BF360C" Foreground="White"/>
                        <Button Name="btnCharts"     Content="&#x1F4CA; Charts"
                                Background="#AD1457" Foreground="White"/>
                        <Button Name="btnPrintLabel" Content="&#x1F3F7; Print Labels"
                                Background="#4E342E" Foreground="White"/>
                        <Separator Width="12" Background="Transparent"/>
                        <Button Name="btnInsurance"  Content="&#x1F4CB; Insurance"
                                Background="#1B5E20" Foreground="White"/>
                        <Button Name="btnSetTracker" Content="&#x1F3AF; Set Tracker"
                                Background="#880E4F" Foreground="White"/>
                        <Button Name="btnValHistory" Content="&#x1F4C8; Value History"
                                Background="#311B92" Foreground="White"/>
                    </WrapPanel>
                </Border>

                <!-- Search / Filter Bar -->
                <Border DockPanel.Dock="Top" Background="#D9D9D9" Padding="8,5">
                    <WrapPanel VerticalAlignment="Center">
                        <TextBlock Text="Search:" Foreground="#104861" FontWeight="Bold" FontSize="12"
                                   VerticalAlignment="Center" Margin="0,0,6,0"/>
                        <Label Content="Year:"/>
                        <TextBox Name="tbFilterYear"    Width="50"  Margin="1,0,5,0"/>
                        <Label Content="Mint:"/>
                        <TextBox Name="tbFilterMint"    Width="40"  Margin="1,0,5,0"/>
                        <Label Content="Denomination:"/>
                        <TextBox Name="tbFilterDenom"   Width="110" Margin="1,0,5,0"/>
                        <Label Content="Grade:"/>
                        <TextBox Name="tbFilterGrade"   Width="75"  Margin="1,0,5,0"/>
                        <Label Content="Metal:"/>
                        <TextBox Name="tbFilterMetal"   Width="65"  Margin="1,0,5,0"/>
                        <Label Content="Key:"/>
                        <TextBox Name="tbFilterKey"     Width="55"  Margin="1,0,5,0"/>
                        <Button Name="btnSearch"  Content="Search" Width="68"
                                Background="#104861" Foreground="White" Margin="4,0"/>
                        <Button Name="btnClear"   Content="Clear"  Width="58"
                                Background="#78909C" Foreground="White" Margin="4,0"/>
                    </WrapPanel>
                </Border>

                <!-- Data Grid -->
                <DataGrid Name="dgCoins" Margin="0" CanUserSortColumns="True">
                    <DataGrid.Columns>
                        <DataGridTextColumn Header="ID"           Binding="{Binding ID}"            Width="42"/>
                        <DataGridTextColumn Header="Year"         Binding="{Binding Year}"           Width="50"/>
                        <DataGridTextColumn Header="Mint"         Binding="{Binding Mint}"           Width="48"/>
                        <DataGridTextColumn Header="Denomination" Binding="{Binding Denomination}"   Width="140"/>
                        <DataGridTextColumn Header="Type"         Binding="{Binding Type}"           Width="110"/>
                        <DataGridTextColumn Header="Variety"      Binding="{Binding Variety}"        Width="130"/>
                        <DataGridTextColumn Header="Grade"        Binding="{Binding Grade}"          Width="82"/>
                        <DataGridTextColumn Header="Face$"        Binding="{Binding FaceValue}"      Width="50"/>
                        <DataGridTextColumn Header="Book$"        Binding="{Binding BookValue}"      Width="55"/>
                        <DataGridTextColumn Header="Key"          Binding="{Binding KeyDates}"       Width="68"/>
                        <DataGridTextColumn Header="Metal"        Binding="{Binding Metal}"          Width="78"/>
                        <DataGridTextColumn Header="Dup#"         Binding="{Binding DupCount}"       Width="40"/>
                        <DataGridTextColumn Header="Location"     Binding="{Binding Location}"       Width="72"/>
                        <DataGridTextColumn Header="Comments"     Binding="{Binding Comments}"       Width="*"/>
                    </DataGrid.Columns>
                    <DataGrid.RowDetailsTemplate>
                        <DataTemplate>
                            <Border Background="#E8F0F5" Padding="10,6" BorderBrush="#B0C4D0" BorderThickness="0,1,0,0">
                                <StackPanel>
                                    <WrapPanel>
                                        <TextBlock Foreground="#444" FontSize="11"
                                                   Text="{Binding Error, StringFormat='Error: {0}'}" Margin="0,0,16,0"/>
                                        <TextBlock Foreground="#444" FontSize="11"
                                                   Text="{Binding DateRangeMinted, StringFormat='Minted: {0}'}" Margin="0,0,16,0"/>
                                        <TextBlock Foreground="#444" FontSize="11"
                                                   Text="{Binding Weight, StringFormat='Weight: {0}'}" Margin="0,0,16,0"/>
                                        <TextBlock Foreground="#444" FontSize="11"
                                                   Text="{Binding InCollection, StringFormat='InColl: {0}'}" Margin="0,0,16,0"/>
                                        <TextBlock Foreground="#444" FontSize="11"
                                                   Text="{Binding DataStandard, StringFormat='Standard: {0}'}" Margin="0,0,16,0"/>
                                        <TextBlock Foreground="#888" FontSize="11"
                                                   Text="{Binding Country, StringFormat='Country: {0}'}" Margin="0,0,16,0"/>
                                    </WrapPanel>
                                    <WrapPanel Margin="0,3,0,0">
                                        <TextBlock Foreground="#444" FontSize="11"
                                                   Text="{Binding PurchasePrice, StringFormat='Paid: ${0}'}" Margin="0,0,16,0"/>
                                        <TextBlock Foreground="#444" FontSize="11"
                                                   Text="{Binding CurrentValue, StringFormat='Value: ${0}'}" Margin="0,0,16,0"/>
                                        <TextBlock Foreground="#444" FontSize="11"
                                                   Text="{Binding DupFaceValue, StringFormat='DupFace$: {0}'}" Margin="0,0,16,0"/>
                                        <TextBlock Foreground="#444" FontSize="11"
                                                   Text="{Binding DupGradedValue, StringFormat='DupGrade$: {0}'}" Margin="0,0,16,0"/>
                                        <TextBlock Foreground="#444" FontSize="11"
                                                   Text="{Binding DupLoc, StringFormat='DupLoc: {0}'}" Margin="0,0,16,0"/>
                                    </WrapPanel>
                                    <WrapPanel Margin="0,3,0,0">
                                        <TextBlock Foreground="#444" FontSize="11"
                                                   Text="{Binding CertCompany, StringFormat='Cert: {0}'}" Margin="0,0,16,0"/>
                                        <TextBlock Foreground="#444" FontSize="11"
                                                   Text="{Binding CertNumber, StringFormat='Cert#: {0}'}" Margin="0,0,16,0"/>
                                        <TextBlock Foreground="#444" FontSize="11"
                                                   Text="{Binding SlabGrade, StringFormat='Slab: {0}'}" Margin="0,0,16,0"/>
                                        <TextBlock Foreground="#444" FontSize="11"
                                                   Text="{Binding Mintage, StringFormat='Mintage: {0}'}" Margin="0,0,16,0"/>
                                        <TextBlock Foreground="#444" FontSize="11"
                                                   Text="{Binding RarityScore, StringFormat='Rarity: {0}'}" Margin="0,0,16,0"/>
                                    </WrapPanel>
                                    <WrapPanel Margin="0,3,0,0">
                                        <TextBlock Foreground="#444" FontSize="11"
                                                   Text="{Binding Notes, StringFormat='Notes: {0}'}" Margin="0,0,16,0"/>
                                        <TextBlock Foreground="#888" FontSize="11"
                                                   Text="{Binding DateAdded, StringFormat='Added: {0}'}" Margin="0,0,16,0"/>
                                    </WrapPanel>
                                </StackPanel>
                            </Border>
                        </DataTemplate>
                    </DataGrid.RowDetailsTemplate>
                </DataGrid>
            </DockPanel>
        </Grid>
    </DockPanel>
</Window>
'@

# ---------------------------------------------------------------------------
# Build & Wire Up Main Window
# ---------------------------------------------------------------------------
$reader  = [System.Xml.XmlNodeReader]::new($mainXaml)
$mainWin = [System.Windows.Markup.XamlReader]::Load($reader)

# Right-side controls
$dgCoins          = $mainWin.FindName('dgCoins')
$lblCoinCount     = $mainWin.FindName('lblCoinCount')
$lblStatus        = $mainWin.FindName('lblStatus')
$btnAdd           = $mainWin.FindName('btnAdd')
$btnDelete        = $mainWin.FindName('btnDelete')
$btnStats         = $mainWin.FindName('btnStats')
$btnExport        = $mainWin.FindName('btnExport')
$btnRefresh       = $mainWin.FindName('btnRefresh')
$btnBackup        = $mainWin.FindName('btnBackup')
$btnImport        = $mainWin.FindName('btnImport')
$btnImportExcel   = $mainWin.FindName('btnImportExcel')
$btnClone         = $mainWin.FindName('btnClone')
$btnExportCsv     = $mainWin.FindName('btnExportCsv')
$btnRestore       = $mainWin.FindName('btnRestore')
$btnResearch      = $mainWin.FindName('btnResearch')
$btnMeltCalc      = $mainWin.FindName('btnMeltCalc')
$btnCharts        = $mainWin.FindName('btnCharts')
$btnPrintLabel    = $mainWin.FindName('btnPrintLabel')
$btnInsurance     = $mainWin.FindName('btnInsurance')
$btnSetTracker    = $mainWin.FindName('btnSetTracker')
$btnValHistory    = $mainWin.FindName('btnValHistory')
$btnUndo          = $mainWin.FindName('btnUndo')
$btnSearch        = $mainWin.FindName('btnSearch')
$btnClear         = $mainWin.FindName('btnClear')
$tbFilterYear     = $mainWin.FindName('tbFilterYear')
$tbFilterMint     = $mainWin.FindName('tbFilterMint')
$tbFilterDenom    = $mainWin.FindName('tbFilterDenom')
$tbFilterGrade    = $mainWin.FindName('tbFilterGrade')
$tbFilterMetal    = $mainWin.FindName('tbFilterMetal')
$tbFilterKey      = $mainWin.FindName('tbFilterKey')
$cbCollections    = $mainWin.FindName('cbCollections')
$btnNewCollection = $mainWin.FindName('btnNewCollection')
$btnDelCollection = $mainWin.FindName('btnDelCollection')

# Dashboard stat card references (text + clickable border)
$statHolding = $mainWin.FindName('statHolding')
$statMissing = $mainWin.FindName('statMissing')
$statMinted  = $mainWin.FindName('statMinted')
$statHoldPct = $mainWin.FindName('statHoldPct')
$statBookVal = $mainWin.FindName('statBookVal')
$statFaceVal = $mainWin.FindName('statFaceVal')
$statGraded  = $mainWin.FindName('statGraded')
$statKeys    = $mainWin.FindName('statKeys')
$statDups    = $mainWin.FindName('statDups')
$cardHolding = $mainWin.FindName('cardHolding')
$cardMissing = $mainWin.FindName('cardMissing')
$cardMinted  = $mainWin.FindName('cardMinted')
$cardHoldPct = $mainWin.FindName('cardHoldPct')
$cardBookVal = $mainWin.FindName('cardBookVal')
$cardFaceVal = $mainWin.FindName('cardFaceVal')
$cardGraded  = $mainWin.FindName('cardGraded')
$cardKeys    = $mainWin.FindName('cardKeys')
$cardDups    = $mainWin.FindName('cardDups')

# Left-side form controls
$lblFormTitle     = $mainWin.FindName('lblFormTitle')
$lblFormError     = $mainWin.FindName('lblFormError')
$btnClearForm     = $mainWin.FindName('btnClearForm')
$btnSaveCoin      = $mainWin.FindName('btnSaveCoin')
$tbYear           = $mainWin.FindName('tbYear')
$cbMint           = $mainWin.FindName('cbMint')
$cbDenom          = $mainWin.FindName('cbDenom')
$tbType           = $mainWin.FindName('tbType')
$tbVariety        = $mainWin.FindName('tbVariety')
$cbCountry        = $mainWin.FindName('cbCountry')
$tbSeries         = $mainWin.FindName('tbSeries')
$tbError          = $mainWin.FindName('tbError')
$cbGrade          = $mainWin.FindName('cbGrade')
$tbFaceValue      = $mainWin.FindName('tbFaceValue')
$tbBookValue      = $mainWin.FindName('tbBookValue')
$tbPurchasePrice  = $mainWin.FindName('tbPurchasePrice')
$tbCurrentValue   = $mainWin.FindName('tbCurrentValue')
$cbKeyDates       = $mainWin.FindName('cbKeyDates')
$cbMetal          = $mainWin.FindName('cbMetal')
$tbWeight         = $mainWin.FindName('tbWeight')
$cbInCollection   = $mainWin.FindName('cbInCollection')
$tbLocation       = $mainWin.FindName('tbLocation')
$tbDateAdded      = $mainWin.FindName('tbDateAdded')
$tbDateRangeMinted= $mainWin.FindName('tbDateRangeMinted')
$cbDataStandard   = $mainWin.FindName('cbDataStandard')
$tbDupCount       = $mainWin.FindName('tbDupCount')
$tbDupFaceValue   = $mainWin.FindName('tbDupFaceValue')
$tbDupGradedValue = $mainWin.FindName('tbDupGradedValue')
$tbDupLoc         = $mainWin.FindName('tbDupLoc')
$cbCertCompany    = $mainWin.FindName('cbCertCompany')
$tbCertNumber     = $mainWin.FindName('tbCertNumber')
$btnVerify        = $mainWin.FindName('btnVerify')
$tbSlabGrade      = $mainWin.FindName('tbSlabGrade')
$tbMintage        = $mainWin.FindName('tbMintage')
$cbRarityScore    = $mainWin.FindName('cbRarityScore')
$tbComments       = $mainWin.FindName('tbComments')
$tbNotes          = $mainWin.FindName('tbNotes')
$tbVAMNumber      = $mainWin.FindName('tbVAMNumber')
$tbFSNumber       = $mainWin.FindName('tbFSNumber')
$cbErrorType      = $mainWin.FindName('cbErrorType')
$cbVarietyType    = $mainWin.FindName('cbVarietyType')
$tbImagePath      = $mainWin.FindName('tbImagePath')
$btnBrowseImage   = $mainWin.FindName('btnBrowseImage')
$imgPreview       = $mainWin.FindName('imgPreview')

# Thumbnail column via code-behind
$imgColFactory = [System.Windows.FrameworkElementFactory]::new([System.Windows.Controls.Image])
$imgBinding    = [System.Windows.Data.Binding]::new('ImagePath')
$imgBinding.Converter = [ImagePathConverter]::new()
$imgColFactory.SetBinding([System.Windows.Controls.Image]::SourceProperty, $imgBinding)
$imgColFactory.SetValue([System.Windows.FrameworkElement]::WidthProperty,  [double]40)
$imgColFactory.SetValue([System.Windows.FrameworkElement]::HeightProperty, [double]40)
$imgColFactory.SetValue([System.Windows.Controls.Image]::StretchProperty,
    [System.Windows.Media.Stretch]::UniformToFill)
$imgTemplate = [System.Windows.DataTemplate]::new()
$imgTemplate.VisualTree = $imgColFactory
$imgColumn = [System.Windows.Controls.DataGridTemplateColumn]::new()
$imgColumn.Header     = ''
$imgColumn.Width      = 48
$imgColumn.IsReadOnly = $true
$imgColumn.CellTemplate = $imgTemplate
$dgCoins.Columns.Insert(0, $imgColumn)

# State
$script:EditingCoinID = $null
$script:LastDeleted   = $null
$script:AllColumns = @('ID','Year','Mint','Denomination','Type','Variety','Country','Grade',
    'FaceValue','BookValue','PurchasePrice','CurrentValue','KeyDates','Error','Metal','Weight',
    'DupCount','DupFaceValue','DupGradedValue','Location','DupLoc','Comments','Notes',
    'DateRangeMinted','DataStandard','InCollection','DateAdded','CertCompany','CertNumber',
    'SlabGrade','Mintage','RarityScore','ImagePath','Series',
    'VAMNumber','FSNumber','ErrorType','VarietyType')

# ---------------------------------------------------------------------------
# Form Helper Functions
# ---------------------------------------------------------------------------
function Get-ComboText {
    param([System.Windows.Controls.ComboBox]$cb)
    $t = $cb.Text
    if ([string]::IsNullOrWhiteSpace($t) -and $null -ne $cb.SelectedItem) {
        $si = $cb.SelectedItem
        if ($si -is [System.Windows.Controls.ComboBoxItem]) { $t = $si.Content }
        else { $t = [string]$si }
    }
    if ($null -eq $t) { return '' }
    return $t.Trim()
}

function Set-ComboValue {
    param([System.Windows.Controls.ComboBox]$cb, [string]$val)
    $found = $false
    for ($i = 0; $i -lt $cb.Items.Count; $i++) {
        $item = $cb.Items[$i]
        if ($item -is [System.Windows.Controls.ComboBoxItem] -and $item.Content -eq $val) {
            $cb.SelectedIndex = $i; $found = $true; break
        }
    }
    if (-not $found -and $cb.IsEditable) { $cb.Text = $val }
    elseif (-not $found) { $cb.SelectedIndex = -1 }
}

function Clear-CoinForm {
    $script:EditingCoinID = $null
    $lblFormTitle.Text   = 'Add New Coin'
    $lblFormError.Visibility = 'Collapsed'
    $tbYear.Text         = ''
    $cbMint.Text         = '';  $cbMint.SelectedIndex = -1
    $cbDenom.Text        = '';  $cbDenom.SelectedIndex = -1
    $tbType.Text         = ''
    $tbVariety.Text      = ''
    $cbCountry.Text      = 'USA'
    $tbSeries.Text       = ''
    $tbError.Text        = ''
    $cbGrade.Text        = '';  $cbGrade.SelectedIndex = -1
    $tbFaceValue.Text    = '0'
    $tbBookValue.Text    = '0'
    $tbPurchasePrice.Text= '0'
    $tbCurrentValue.Text = '0'
    $cbKeyDates.Text     = '';  $cbKeyDates.SelectedIndex = -1
    $cbMetal.Text        = '';  $cbMetal.SelectedIndex = -1
    $tbWeight.Text       = ''
    Set-ComboValue $cbInCollection 'Y'
    $tbLocation.Text     = ''
    $tbDateAdded.Text    = (Get-Date -Format 'yyyy-MM-dd')
    $tbDateRangeMinted.Text = ''
    $cbDataStandard.Text = '';  $cbDataStandard.SelectedIndex = -1
    $tbDupCount.Text     = '0'
    $tbDupFaceValue.Text = '0'
    $tbDupGradedValue.Text = '0'
    $tbDupLoc.Text       = ''
    $cbCertCompany.SelectedIndex = -1
    $tbCertNumber.Text   = ''
    $tbSlabGrade.Text    = ''
    $tbMintage.Text      = ''
    $cbRarityScore.SelectedIndex = -1
    $tbComments.Text     = ''
    $tbNotes.Text        = ''
    $tbVAMNumber.Text    = ''
    $tbFSNumber.Text     = ''
    $cbErrorType.Text    = '';  $cbErrorType.SelectedIndex = -1
    $cbVarietyType.Text  = '';  $cbVarietyType.SelectedIndex = -1
    $tbImagePath.Text    = ''
    $imgPreview.Source    = $null
    $lblStatus.Text      = 'Ready - New coin form cleared'
}

function Populate-CoinForm {
    param([hashtable]$Coin)
    if ($null -eq $Coin) { Clear-CoinForm; return }

    $script:EditingCoinID = $Coin['ID']
    $lblFormTitle.Text    = "Edit Coin [ID: $($Coin['ID'])]"
    $lblFormError.Visibility = 'Collapsed'

    $tbYear.Text          = $Coin['Year']
    Set-ComboValue $cbMint  $Coin['Mint']
    Set-ComboValue $cbDenom $Coin['Denomination']
    $tbType.Text          = $Coin['Type']
    $tbVariety.Text       = $Coin['Variety']
    Set-ComboValue $cbCountry $Coin['Country']
    $tbSeries.Text        = $Coin['Series']
    $tbError.Text         = $Coin['Error']
    Set-ComboValue $cbGrade   $Coin['Grade']
    $tbFaceValue.Text     = $Coin['FaceValue']
    $tbBookValue.Text     = $Coin['BookValue']
    $tbPurchasePrice.Text = $Coin['PurchasePrice']
    $tbCurrentValue.Text  = $Coin['CurrentValue']
    Set-ComboValue $cbKeyDates $Coin['KeyDates']
    Set-ComboValue $cbMetal    $Coin['Metal']
    $tbWeight.Text        = $Coin['Weight']
    Set-ComboValue $cbInCollection $Coin['InCollection']
    $tbLocation.Text      = $Coin['Location']
    $tbDateAdded.Text     = $Coin['DateAdded']
    $tbDateRangeMinted.Text = $Coin['DateRangeMinted']
    Set-ComboValue $cbDataStandard $Coin['DataStandard']
    $tbDupCount.Text      = $Coin['DupCount']
    $tbDupFaceValue.Text  = $Coin['DupFaceValue']
    $tbDupGradedValue.Text= $Coin['DupGradedValue']
    $tbDupLoc.Text        = $Coin['DupLoc']
    Set-ComboValue $cbCertCompany  $Coin['CertCompany']
    $tbCertNumber.Text    = $Coin['CertNumber']
    $tbSlabGrade.Text     = $Coin['SlabGrade']
    $tbMintage.Text       = $Coin['Mintage']
    Set-ComboValue $cbRarityScore $Coin['RarityScore']
    $tbComments.Text      = $Coin['Comments']
    $tbNotes.Text         = $Coin['Notes']
    $tbVAMNumber.Text     = $Coin['VAMNumber']
    $tbFSNumber.Text      = $Coin['FSNumber']
    Set-ComboValue $cbErrorType   $Coin['ErrorType']
    Set-ComboValue $cbVarietyType $Coin['VarietyType']
    $tbImagePath.Text     = $Coin['ImagePath']

    # Load image preview; fall back to denomination type image if no personal photo
    $imgPreview.Source = $null
    $imgPath = $Coin['ImagePath']
    if ([string]::IsNullOrWhiteSpace($imgPath) -or -not (Test-Path $imgPath)) {
        $imgPath = Get-CoinTypeImage -Denomination $Coin['Denomination']
    }
    if (-not [string]::IsNullOrWhiteSpace($imgPath) -and (Test-Path $imgPath)) {
        try {
            $bi = [System.Windows.Media.Imaging.BitmapImage]::new()
            $bi.BeginInit()
            $bi.UriSource = [Uri]::new($imgPath, [UriKind]::Absolute)
            $bi.DecodePixelWidth = 100
            $bi.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
            $bi.EndInit(); $bi.Freeze()
            $imgPreview.Source = $bi
        } catch { }
    }
}

# ---------------------------------------------------------------------------
# Grid / Collection Helpers
# ---------------------------------------------------------------------------
function Refresh-Grid {
    param([object[]]$Data = $null)
    if ($null -eq $Data) { $Data = @(Get-Collection) }
    $sorted = @($Data | Sort-Object -Property 'Year', 'Mint')

    $dt = [System.Data.DataTable]::new()
    foreach ($col in $script:AllColumns) { $null = $dt.Columns.Add($col, [string]) }
    foreach ($c in $sorted) {
        $row = $dt.NewRow()
        foreach ($col in $script:AllColumns) { $row[$col] = [string]$c.$col }
        $dt.Rows.Add($row)
    }

    $dgCoins.ItemsSource = $dt.DefaultView
    $count  = $dt.Rows.Count
    $plural = if ($count -ne 1) { 's' } else { '' }
    $lblCoinCount.Text = "$count coin$plural"
    $lblStatus.Text    = "Loaded $count coin$plural - $(Get-Date -Format 'HH:mm:ss')"

    # Update dashboard stats
    $holdingCount = @($sorted | Where-Object { $_.InCollection -eq 'Y' }).Count
    $missingCount = @($sorted | Where-Object { $_.InCollection -ne 'Y' }).Count
    $totalCount   = $sorted.Count
    $holdPct      = if ($totalCount -gt 0) { [Math]::Round(($holdingCount / $totalCount) * 100, 1) } else { 0 }
    $bookTotal    = [Math]::Round(($sorted | ForEach-Object { [double]$_.BookValue } | Measure-Object -Sum).Sum, 2)
    $faceTotal    = [Math]::Round(($sorted | ForEach-Object { [double]$_.FaceValue } | Measure-Object -Sum).Sum, 2)
    $gradedCount  = @($sorted | Where-Object { $_.Grade -ne '' -and $null -ne $_.Grade }).Count
    $keyCount     = @($sorted | Where-Object { $_.KeyDates -ne '' -and $null -ne $_.KeyDates }).Count
    $dupTotal     = ($sorted | ForEach-Object { [int]$_.DupCount } | Measure-Object -Sum).Sum

    $statHolding.Text = [string]$holdingCount
    $statMissing.Text = [string]$missingCount
    $statMinted.Text  = [string]$totalCount
    $statHoldPct.Text = "$holdPct%"
    $statBookVal.Text = '${0:N0}' -f $bookTotal
    $statFaceVal.Text = '${0:N0}' -f $faceTotal
    $statGraded.Text  = [string]$gradedCount
    $statKeys.Text    = [string]$keyCount
    $statDups.Text    = [string]$dupTotal
}

function Get-SelectedCoin {
    $drv = $dgCoins.SelectedItem
    if ($null -eq $drv) { return $null }
    $result = @{}
    foreach ($col in $script:AllColumns) { $result[$col] = $drv[$col] }
    return $result
}

function Refresh-CollectionDropdown {
    $cbCollections.Items.Clear()
    $names = Get-CollectionNames
    foreach ($n in $names) {
        $item = [System.Windows.Controls.ComboBoxItem]::new()
        $item.Content = $n
        $cbCollections.Items.Add($item) | Out-Null
    }
    for ($i = 0; $i -lt $cbCollections.Items.Count; $i++) {
        if ($cbCollections.Items[$i].Content -eq $script:ActiveCollection) {
            $cbCollections.SelectedIndex = $i; break
        }
    }
}

function Switch-Collection {
    param([string]$Name)
    $file = Join-Path $script:CollectionsDir "$Name.csv"
    if (-not (Test-Path $file)) {
        $script:CsvHeader | Set-Content -Path $file -Encoding UTF8
    }
    $script:DataFile = $file
    $script:ActiveCollection = $Name
    Refresh-Grid
    Refresh-CollectionDropdown
}

function Copy-CoinImage {
    param([string]$SourcePath, [string]$CoinID)
    if ([string]::IsNullOrWhiteSpace($SourcePath) -or -not (Test-Path $SourcePath)) { return '' }
    $imgDir = Join-Path $PSScriptRoot 'Images'
    if (-not (Test-Path $imgDir)) { New-Item -Path $imgDir -ItemType Directory -Force | Out-Null }
    $ext     = [System.IO.Path]::GetExtension($SourcePath)
    $destImg = Join-Path $imgDir "$CoinID$ext"
    if ($SourcePath -ne $destImg) { Copy-Item -Path $SourcePath -Destination $destImg -Force }
    return $destImg
}

# ---------------------------------------------------------------------------
# Event Handlers
# ---------------------------------------------------------------------------

# Grid row selection -> auto-populate form + update count
$dgCoins.Add_SelectionChanged({
    $sel = Get-SelectedCoin
    if ($null -ne $sel) {
        Populate-CoinForm -Coin $sel
    }
    $total   = $dgCoins.Items.Count
    $selCnt  = $dgCoins.SelectedItems.Count
    $lblCoinCount.Text = if ($selCnt -gt 0) { "$total coins  |  $selCnt selected" } else { "$total coins" }
})

# Double-click toggles row details (expand/collapse extra info)
$dgCoins.Add_MouseDoubleClick({
    if ($null -eq $dgCoins.SelectedItem) { return }
    $dgCoins.RowDetailsVisibilityMode =
        if ($dgCoins.RowDetailsVisibilityMode -eq 'Collapsed') { 'VisibleWhenSelected' } else { 'Collapsed' }
})

# Click preview image -> open full-size viewer
$imgPreview.Add_MouseLeftButtonUp({
    $imgPath = $tbImagePath.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($imgPath) -or -not (Test-Path $imgPath)) {
        $sel = Get-SelectedCoin
        if ($null -ne $sel) { $imgPath = Get-CoinTypeImage -Denomination $sel['Denomination'] }
    }
    Show-ImageViewer -ImagePath $imgPath
})
$imgPreview.Cursor = [System.Windows.Input.Cursors]::Hand

# Stat card left-click = filter grid | right-click = export to file
foreach ($c in @($cardHolding,$cardMissing,$cardMinted,$cardHoldPct,$cardBookVal,$cardFaceVal,$cardGraded,$cardKeys,$cardDups)) {
    $c.ToolTip = 'Left-click: filter grid  |  Right-click: export to text file'
}
$cardHolding.Add_MouseLeftButtonUp({ Refresh-Grid -Data @(Get-Collection | Where-Object { $_.InCollection -eq 'Y' });  $lblStatus.Text = 'Filter: In Collection (Holding)' })
$cardMissing.Add_MouseLeftButtonUp({ Refresh-Grid -Data @(Get-Collection | Where-Object { $_.InCollection -ne 'Y' });  $lblStatus.Text = 'Filter: Not In Collection (Missing)' })
$cardMinted.Add_MouseLeftButtonUp({  Refresh-Grid;  $lblStatus.Text = 'Filter cleared - showing all coins' })
$cardHoldPct.Add_MouseLeftButtonUp({ Refresh-Grid;  $lblStatus.Text = 'Filter cleared - showing all coins' })
$cardBookVal.Add_MouseLeftButtonUp({ Refresh-Grid -Data @(Get-Collection | Sort-Object { [double]$_.BookValue } -Descending); $lblStatus.Text = 'Sorted by Book Value descending' })
$cardFaceVal.Add_MouseLeftButtonUp({ Refresh-Grid -Data @(Get-Collection | Sort-Object { [double]$_.FaceValue } -Descending); $lblStatus.Text = 'Sorted by Face Value descending' })
$cardGraded.Add_MouseLeftButtonUp({  Refresh-Grid -Data @(Get-Collection | Where-Object { $_.Grade -ne '' -and $null -ne $_.Grade });     $lblStatus.Text = 'Filter: Graded coins' })
$cardKeys.Add_MouseLeftButtonUp({    Refresh-Grid -Data @(Get-Collection | Where-Object { $_.KeyDates -ne '' -and $null -ne $_.KeyDates }); $lblStatus.Text = 'Filter: Key Date coins' })
$cardDups.Add_MouseLeftButtonUp({    Refresh-Grid -Data @(Get-Collection | Where-Object { [int]$_.DupCount -gt 0 });                       $lblStatus.Text = 'Filter: Duplicate coins' })
$cardHolding.Add_MouseRightButtonUp({ Export-StatCard -Category 'Holding'    -AllCoins @(Get-Collection) })
$cardMissing.Add_MouseRightButtonUp({ Export-StatCard -Category 'Missing'    -AllCoins @(Get-Collection) })
$cardMinted.Add_MouseRightButtonUp({  Export-StatCard -Category 'Minted'     -AllCoins @(Get-Collection) })
$cardHoldPct.Add_MouseRightButtonUp({ Export-StatCard -Category 'HoldPct'   -AllCoins @(Get-Collection) })
$cardBookVal.Add_MouseRightButtonUp({ Export-StatCard -Category 'BookValue'  -AllCoins @(Get-Collection) })
$cardFaceVal.Add_MouseRightButtonUp({ Export-StatCard -Category 'FaceValue'  -AllCoins @(Get-Collection) })
$cardGraded.Add_MouseRightButtonUp({  Export-StatCard -Category 'Graded'    -AllCoins @(Get-Collection) })
$cardKeys.Add_MouseRightButtonUp({    Export-StatCard -Category 'KeyDates'  -AllCoins @(Get-Collection) })
$cardDups.Add_MouseRightButtonUp({    Export-StatCard -Category 'Duplicates' -AllCoins @(Get-Collection) })

# + New Coin (clears form for adding)
$btnAdd.Add_Click({
    Clear-CoinForm
    $tbYear.Focus() | Out-Null
})

# Clear / New button on form panel
$btnClearForm.Add_Click({
    Clear-CoinForm
    $tbYear.Focus() | Out-Null
})

# Save Coin (handles both Add and Edit)
$btnSaveCoin.Add_Click({
    $lblFormError.Visibility = 'Collapsed'
    $errors = @()

    $yr = $tbYear.Text.Trim()
    if ($yr -notmatch '^\d{4}$') { $errors += 'Year must be 4 digits.' }
    elseif ([int]$yr -lt 1700 -or [int]$yr -gt 2099) { $errors += 'Year: 1700-2099.' }
    $denomText = Get-ComboText $cbDenom
    if ($denomText -eq '') { $errors += 'Denomination required.' }
    $pp = $tbPurchasePrice.Text.Trim()
    if ($pp -ne '' -and $pp -notmatch '^\d+(\.\d{1,2})?$') { $errors += 'Purchase Price invalid.' }
    $cv = $tbCurrentValue.Text.Trim()
    if ($cv -ne '' -and $cv -notmatch '^\d+(\.\d{1,2})?$') { $errors += 'Current Value invalid.' }
    $fv = $tbFaceValue.Text.Trim()
    if ($fv -ne '' -and $fv -notmatch '^\d+(\.\d{1,4})?$') { $errors += 'Face Value invalid.' }
    $bv = $tbBookValue.Text.Trim()
    if ($bv -ne '' -and $bv -notmatch '^\d+(\.\d{1,2})?$') { $errors += 'Book Value invalid.' }
    $mintageVal = $tbMintage.Text.Trim()
    if ($mintageVal -ne '' -and $mintageVal -notmatch '^\d+$') { $errors += 'Mintage must be whole number.' }
    $dupC = $tbDupCount.Text.Trim()
    if ($dupC -ne '' -and $dupC -notmatch '^\d+$') { $errors += 'Dup Count must be whole number.' }

    if ($errors.Count -gt 0) {
        $lblFormError.Text = ($errors -join '  |  ')
        $lblFormError.Visibility = 'Visible'
        return
    }

    $mintText    = (Get-ComboText $cbMint).ToUpper()
    $gradeText   = Get-ComboText $cbGrade
    $countryText = Get-ComboText $cbCountry
    if ($countryText -eq '') { $countryText = 'USA' }

    $coinData = [PSCustomObject]@{
        Year            = $yr
        Mint            = $mintText
        Denomination    = $denomText
        Type            = $tbType.Text.Trim()
        Variety         = $tbVariety.Text.Trim()
        Country         = $countryText
        Grade           = $gradeText
        FaceValue       = if ($fv -eq '') { '0' } else { $fv }
        BookValue       = if ($bv -eq '') { '0' } else { $bv }
        PurchasePrice   = if ($pp -eq '') { '0' } else { $pp }
        CurrentValue    = if ($cv -eq '') { '0' } else { $cv }
        KeyDates        = (Get-ComboText $cbKeyDates)
        Error           = $tbError.Text.Trim()
        Metal           = (Get-ComboText $cbMetal)
        Weight          = $tbWeight.Text.Trim()
        DupCount        = if ($dupC -eq '') { '0' } else { $dupC }
        DupFaceValue    = $tbDupFaceValue.Text.Trim()
        DupGradedValue  = $tbDupGradedValue.Text.Trim()
        Location        = $tbLocation.Text.Trim()
        DupLoc          = $tbDupLoc.Text.Trim()
        Comments        = $tbComments.Text.Trim()
        Notes           = $tbNotes.Text.Trim()
        DateRangeMinted = $tbDateRangeMinted.Text.Trim()
        DataStandard    = (Get-ComboText $cbDataStandard)
        InCollection    = (Get-ComboText $cbInCollection)
        DateAdded       = $tbDateAdded.Text
        CertCompany     = (Get-ComboText $cbCertCompany)
        CertNumber      = $tbCertNumber.Text.Trim()
        SlabGrade       = $tbSlabGrade.Text.Trim()
        Mintage         = $mintageVal
        RarityScore     = (Get-ComboText $cbRarityScore)
        ImagePath       = $tbImagePath.Text.Trim()
        Series          = $tbSeries.Text.Trim()
        VAMNumber       = $tbVAMNumber.Text.Trim()
        FSNumber        = $tbFSNumber.Text.Trim()
        ErrorType       = (Get-ComboText $cbErrorType)
        VarietyType     = (Get-ComboText $cbVarietyType)
    }

    $coins = @(Get-Collection)

    if ($null -eq $script:EditingCoinID -or $script:EditingCoinID -eq '') {
        # --- ADD MODE ---
        $dupExists = $coins | Where-Object {
            $_.Year -eq $coinData.Year -and $_.Mint.ToUpper() -eq $coinData.Mint -and $_.Denomination -eq $coinData.Denomination
        }
        if ($dupExists) {
            $confirm = [System.Windows.MessageBox]::Show(
                "A coin with Year $($coinData.Year), Mint $($coinData.Mint), Denomination $($coinData.Denomination) already exists.`nAdd anyway?",
                'Duplicate Detected', 'YesNo', 'Warning')
            if ($confirm -ne 'Yes') { $lblStatus.Text = 'Add cancelled'; return }
        }
        $newID = New-CoinID $coins
        $coinData | Add-Member -NotePropertyName 'ID' -NotePropertyValue $newID -Force
        $coinData.ImagePath = Copy-CoinImage -SourcePath $coinData.ImagePath -CoinID $newID
        $coins += $coinData
        Save-Collection $coins
        Refresh-Grid
        $lblStatus.Text = "Added: $yr $mintText $denomText (ID: $newID)"
        # Stay in add mode, clear form for next entry
        Clear-CoinForm
    }
    else {
        # --- EDIT MODE ---
        $editID = $script:EditingCoinID
        $coinData | Add-Member -NotePropertyName 'ID' -NotePropertyValue $editID -Force
        $coinData.ImagePath = Copy-CoinImage -SourcePath $coinData.ImagePath -CoinID $editID
        $idx = -1
        for ($i = 0; $i -lt $coins.Count; $i++) { if ($coins[$i].ID -eq $editID) { $idx = $i; break } }
        if ($idx -ge 0) {
            $coins[$idx] = $coinData
            Save-Collection $coins
            Refresh-Grid
            $lblStatus.Text = ("Updated ID {0}: {1} {2} {3}" -f $editID, $yr, $mintText, $denomText)
        } else {
            $lblStatus.Text = ("Error: Coin ID {0} not found" -f $editID)
        }
    }
})

# Research (online coin lookup links)
$btnResearch.Add_Click({
    $sel = Get-SelectedCoin
    Show-ResearchMenu -Coin $sel
})

# Melt Value Calculator
$btnMeltCalc.Add_Click({
    $coins = @(Get-Collection)
    if ($coins.Count -eq 0) {
        [System.Windows.MessageBox]::Show('Collection is empty.', 'Melt Calculator') | Out-Null; return
    }
    Show-MeltCalculator -Coins $coins
})

# Charts & Analytics
$btnCharts.Add_Click({
    $coins = @(Get-Collection)
    if ($coins.Count -eq 0) {
        [System.Windows.MessageBox]::Show('Collection is empty.', 'Charts') | Out-Null; return
    }
    Show-ChartsReport -Coins $coins
})

# Print 2x2 Flip Labels
$btnPrintLabel.Add_Click({
    $coins = @(Get-Collection)
    if ($coins.Count -eq 0) {
        [System.Windows.MessageBox]::Show('Collection is empty.', 'Labels') | Out-Null; return
    }
    $choice = [System.Windows.MessageBox]::Show(
        "Generate labels for:`n`nYes = All coins in collection ($($coins.Count))`nNo = Only the selected coin`nCancel = Abort",
        'Print Labels', 'YesNoCancel', 'Question')
    if ($choice -eq 'Cancel') { return }
    if ($choice -eq 'No') {
        $sel = Get-SelectedCoin
        if ($null -eq $sel) {
            [System.Windows.MessageBox]::Show('Select a coin first.', 'Labels') | Out-Null; return
        }
        $coinObj = [PSCustomObject]$sel
        Show-PrintLabels -Coins @($coinObj)
    } else {
        Show-PrintLabels -Coins $coins
    }
})

# Insurance Report
$btnInsurance.Add_Click({
    $coins = @(Get-Collection)
    if ($coins.Count -eq 0) {
        [System.Windows.MessageBox]::Show('Collection is empty.', 'Insurance Report') | Out-Null; return
    }
    Show-InsuranceReport -Coins $coins
})

# Set Completion Tracker
$btnSetTracker.Add_Click({
    $coins = @(Get-Collection)
    if ($coins.Count -eq 0) {
        [System.Windows.MessageBox]::Show('Collection is empty.', 'Set Tracker') | Out-Null; return
    }
    Show-SetTracker -Coins $coins
})

# Value History
$btnValHistory.Add_Click({
    Record-ValueSnapshot -Coins @(Get-Collection)
    Show-ValueHistory
})

# Verify certification online (enhanced with PCGS lookup)
$btnVerify.Add_Click({
    $certComp = Get-ComboText $cbCertCompany
    $certNum  = $tbCertNumber.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($certNum)) {
        [System.Windows.MessageBox]::Show('Enter a certification number first.', 'Verify', 'OK', 'Information') | Out-Null
        return
    }
    # Try PCGS API lookup first, then open browser
    if ($certComp -eq 'PCGS') {
        $pcgsData = Invoke-PCGSLookup -CertNumber $certNum
        if ($null -ne $pcgsData -and $pcgsData.Count -gt 0) {
            $info = "PCGS Cert $certNum found:`n"
            if ($pcgsData['Denomination']) { $info += "  Denomination: $($pcgsData['Denomination'])`n" }
            if ($pcgsData['Year']) { $info += "  Year: $($pcgsData['Year'])`n" }
            if ($pcgsData['Grade']) { $info += "  Grade: $($pcgsData['Grade'])`n" }
            if ($pcgsData['Variety']) { $info += "  Variety: $($pcgsData['Variety'])`n" }
            $info += "`nOpen PCGS page in browser?"
            $open = [System.Windows.MessageBox]::Show($info, 'PCGS Lookup', 'YesNo', 'Information')
            if ($open -eq 'Yes') { Start-Process "https://www.pcgs.com/cert/$certNum" }
            $lblStatus.Text = ("PCGS lookup complete for cert {0}" -f $certNum)
            return
        }
    }
    switch ($certComp) {
        'PCGS' { Start-Process "https://www.pcgs.com/cert/$certNum" }
        'NGC'  { Start-Process "https://www.ngccoin.com/certlookup/$certNum" }
        default {
            [System.Windows.MessageBox]::Show("Online verification is available for PCGS and NGC only.", 'Verify', 'OK', 'Information') | Out-Null
        }
    }
})

# Browse for coin image
$btnBrowseImage.Add_Click({
    $ofd = [Microsoft.Win32.OpenFileDialog]::new()
    $ofd.Title  = 'Select Coin Image'
    $ofd.Filter = 'Image Files (*.jpg;*.jpeg;*.png;*.bmp;*.gif)|*.jpg;*.jpeg;*.png;*.bmp;*.gif|All Files (*.*)|*.*'
    if (-not $ofd.ShowDialog()) { return }
    $tbImagePath.Text = $ofd.FileName
    try {
        $bi = [System.Windows.Media.Imaging.BitmapImage]::new()
        $bi.BeginInit()
        $bi.UriSource = [Uri]::new($ofd.FileName, [UriKind]::Absolute)
        $bi.DecodePixelWidth = 100
        $bi.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bi.EndInit(); $bi.Freeze()
        $imgPreview.Source = $bi
    } catch { $imgPreview.Source = $null }
})

# Delete coin
$btnDelete.Add_Click({
    $sel = Get-SelectedCoin
    if ($null -eq $sel) {
        [System.Windows.MessageBox]::Show('Select a coin to delete.', 'No Selection') | Out-Null; return
    }
    $confirm = [System.Windows.MessageBox]::Show(
        "Delete ID $($sel.ID): $($sel.Year) $($sel.Mint) $($sel.Denomination)?", 'Confirm Delete', 'YesNo', 'Warning')
    if ($confirm -ne 'Yes') { return }
    $coins = @(Get-Collection)
    $script:LastDeleted = $coins | Where-Object { $_.ID -eq $sel.ID } | Select-Object -First 1
    $remaining = @($coins | Where-Object { $_.ID -ne $sel.ID })
    Save-Collection $remaining
    Clear-CoinForm
    Refresh-Grid
    $btnUndo.Visibility = 'Visible'
    $lblStatus.Text = "Deleted ID $($sel.ID)  — click Undo to restore"
})

# Statistics
$btnStats.Add_Click({
    $coins = @(Get-Collection)
    if ($coins.Count -eq 0) {
        [System.Windows.MessageBox]::Show('Collection is empty.', 'Statistics') | Out-Null; return
    }
    Show-StatsWindow -Coins $coins
})

# Export
$btnExport.Add_Click({
    $coins = @(Get-Collection)
    if ($coins.Count -eq 0) {
        [System.Windows.MessageBox]::Show('Nothing to export.', 'Export') | Out-Null; return
    }
    $path = Export-Report -Coins $coins
    $open = [System.Windows.MessageBox]::Show("Report saved:`n$path`n`nOpen now?", 'Export Complete', 'YesNo', 'Information')
    if ($open -eq 'Yes') { Start-Process notepad.exe $path }
    $lblStatus.Text = "Exported: $path"
})

# Refresh
$btnRefresh.Add_Click({ Refresh-Grid })

# Backup
$btnBackup.Add_Click({
    $path = Invoke-Backup -Force
    if ($path) {
        [System.Windows.MessageBox]::Show("Backup saved:`n$path", 'Backup Complete') | Out-Null
        $lblStatus.Text = "Backup: $path"
    }
})

# Import CSV
$btnImport.Add_Click({
    $ofd = [Microsoft.Win32.OpenFileDialog]::new()
    $ofd.Title = 'Select CSV File'
    $ofd.Filter = 'CSV (*.csv)|*.csv|All (*.*)|*.*'
    $ofd.InitialDirectory = $PSScriptRoot
    if (-not $ofd.ShowDialog()) { return }
    try { $testRows = Import-Csv -Path $ofd.FileName -Encoding UTF8 } catch {
        [System.Windows.MessageBox]::Show("Cannot read CSV:`n$($_.Exception.Message)", 'Error') | Out-Null; return
    }
    $confirm = [System.Windows.MessageBox]::Show(
        "Replace '$($script:ActiveCollection)' with imported file?`nA backup will be created first.",
        'Confirm Import', 'YesNo', 'Warning')
    if ($confirm -ne 'Yes') { return }
    Invoke-Backup -Force | Out-Null
    Copy-Item -Path $ofd.FileName -Destination $script:DataFile -Force
    Refresh-Grid
    $lblStatus.Text = "Imported: $($ofd.FileName)"
})

# Import Excel
$btnImportExcel.Add_Click({
    $ofd = [Microsoft.Win32.OpenFileDialog]::new()
    $ofd.Title = 'Select Excel Coin Collection'
    $ofd.Filter = 'Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All (*.*)|*.*'
    $ofd.InitialDirectory = $PSScriptRoot
    if (-not $ofd.ShowDialog()) { return }

    $choice = [System.Windows.MessageBox]::Show(
        "Yes = Import InCollection=Y only`nNo = Import ALL coins`nCancel = Abort",
        'Import Mode', 'YesNoCancel', 'Question')
    if ($choice -eq 'Cancel') { return }
    $onlyY = ($choice -eq 'Yes')

    $lblStatus.Text = 'Importing from Excel... please wait'
    $mainWin.Dispatcher.Invoke([Action]{}, 'Render')

    try {
        Invoke-Backup -Force | Out-Null
        $extractedCount = Extract-CoinImages -ExcelPath $ofd.FileName
        $imported = Import-FromExcel -ExcelPath $ofd.FileName -OnlyInCollection $onlyY
        if ($null -eq $imported -or $imported.Count -eq 0) {
            [System.Windows.MessageBox]::Show('No coins found.', 'Import') | Out-Null
            $lblStatus.Text = 'Import: 0 coins'; return
        }
        Save-Collection $imported
        Refresh-Grid
        $modeText = if ($onlyY) { 'InCollection=Y' } else { 'all coins' }
        $lblStatus.Text = "Imported $($imported.Count) coins ($modeText)"
        [System.Windows.MessageBox]::Show("Imported $($imported.Count) coins ($modeText).`nBackup created.",
            'Import Complete') | Out-Null
    } catch {
        [System.Windows.MessageBox]::Show("Import failed:`n$($_.Exception.Message)", 'Error') | Out-Null
        $lblStatus.Text = 'Excel import failed'
    }
})

# Search / Filter
$btnSearch.Add_Click({
    $all = @(Get-Collection)
    $fYear  = $tbFilterYear.Text.Trim()
    $fMint  = $tbFilterMint.Text.Trim().ToUpper()
    $fDenom = $tbFilterDenom.Text.Trim()
    $fGrade = $tbFilterGrade.Text.Trim()
    $fMetal = $tbFilterMetal.Text.Trim()
    $fKey   = $tbFilterKey.Text.Trim()
    $results = @($all | Where-Object {
        ($fYear  -eq '' -or $_.Year         -eq $fYear)        -and
        ($fMint  -eq '' -or $_.Mint         -like "*$fMint*")  -and
        ($fDenom -eq '' -or $_.Denomination -like "*$fDenom*") -and
        ($fGrade -eq '' -or $_.Grade        -like "*$fGrade*") -and
        ($fMetal -eq '' -or $_.Metal        -like "*$fMetal*") -and
        ($fKey   -eq '' -or $_.KeyDates     -like "*$fKey*")
    })
    Refresh-Grid -Data $results
    $lblStatus.Text = "Search: $($results.Count) result$(if ($results.Count -ne 1){'s'})"
})

$btnClear.Add_Click({
    $tbFilterYear.Text = ''; $tbFilterMint.Text = ''; $tbFilterDenom.Text = ''
    $tbFilterGrade.Text = ''; $tbFilterMetal.Text = ''; $tbFilterKey.Text = ''
    Refresh-Grid
})

foreach ($tb in @($tbFilterYear, $tbFilterMint, $tbFilterDenom, $tbFilterGrade, $tbFilterMetal, $tbFilterKey)) {
    $tb.Add_KeyDown({
        param($s, $e)
        if ($e.Key -eq 'Return') { $btnSearch.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent)) }
    })
}

# Collection switcher
$cbCollections.Add_SelectionChanged({
    $sel = $cbCollections.SelectedItem
    if ($null -eq $sel) { return }
    $name = $sel.Content
    if ($name -ne $script:ActiveCollection) {
        $script:DataFile = Join-Path $script:CollectionsDir "$name.csv"
        $script:ActiveCollection = $name
        Refresh-Grid
        Clear-CoinForm
        $lblStatus.Text = "Switched to: $name"
    }
})

$btnNewCollection.Add_Click({
    [xml]$inputXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="New Collection" Height="160" Width="350"
        WindowStartupLocation="CenterOwner" ResizeMode="NoResize"
        Background="#F0F2F5">
    <StackPanel Margin="18">
        <TextBlock Text="Collection Name:" Foreground="#1A1A2E" FontSize="13" Margin="0,0,0,6"/>
        <TextBox Name="tbCollName" Background="White" Foreground="#1A1A2E"
                 BorderBrush="#B0B8C4" FontSize="13" Height="28" Padding="4,3"/>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
            <Button Name="btnCollCancel" Content="Cancel" Width="80" Height="30"
                    Background="#78909C" Foreground="White" Margin="0,0,8,0"
                    BorderThickness="0" Cursor="Hand"/>
            <Button Name="btnCollOK" Content="Create" Width="80" Height="30"
                    Background="#2E8B57" Foreground="White"
                    BorderThickness="0" Cursor="Hand"/>
        </StackPanel>
    </StackPanel>
</Window>
'@
    $rdr = [System.Xml.XmlNodeReader]::new($inputXaml)
    $dlg = [System.Windows.Markup.XamlReader]::Load($rdr)
    $tbName  = $dlg.FindName('tbCollName')
    $btnOK   = $dlg.FindName('btnCollOK')
    $btnCncl = $dlg.FindName('btnCollCancel')
    $script:NewCollName = $null
    $btnCncl.Add_Click({ $dlg.Close() })
    $btnOK.Add_Click({
        $n = $tbName.Text.Trim() -replace '[\\/:*?"<>|]', ''
        if ($n -eq '') { return }
        $script:NewCollName = $n
        $dlg.Close()
    })
    $prevEAP = $ErrorActionPreference
    try { $ErrorActionPreference = 'Continue'; $null = $dlg.ShowDialog() }
    finally { $ErrorActionPreference = $prevEAP }
    if ($script:NewCollName) {
        $newFile = Join-Path $script:CollectionsDir "$($script:NewCollName).csv"
        if (Test-Path $newFile) {
            [System.Windows.MessageBox]::Show("'$($script:NewCollName)' already exists.", 'Duplicate') | Out-Null; return
        }
        $script:CsvHeader | Set-Content -Path $newFile -Encoding UTF8
        Switch-Collection $script:NewCollName
        $lblStatus.Text = "Created: $($script:NewCollName)"
    }
})

$btnDelCollection.Add_Click({
    $names = Get-CollectionNames
    if ($names.Count -le 1) {
        [System.Windows.MessageBox]::Show('Cannot delete the last collection.', 'Delete') | Out-Null; return
    }
    $confirm = [System.Windows.MessageBox]::Show(
        "Delete '$($script:ActiveCollection)' and all its data?", 'Confirm Delete', 'YesNo', 'Warning')
    if ($confirm -ne 'Yes') { return }
    $fileToDelete = $script:DataFile
    $otherName = $names | Where-Object { $_ -ne $script:ActiveCollection } | Select-Object -First 1
    Switch-Collection $otherName
    Remove-Item -Path $fileToDelete -Force
    Refresh-CollectionDropdown
    $lblStatus.Text = "Deleted collection, switched to: $otherName"
})

# ---------------------------------------------------------------------------
# Context Menu on DataGrid
# ---------------------------------------------------------------------------
$ctxMenu    = [System.Windows.Controls.ContextMenu]::new()
$ctxEdit    = [System.Windows.Controls.MenuItem]::new(); $ctxEdit.Header    = '_Edit'
$ctxCloneM  = [System.Windows.Controls.MenuItem]::new(); $ctxCloneM.Header  = 'Cl_one'
$ctxDeleteM = [System.Windows.Controls.MenuItem]::new(); $ctxDeleteM.Header = '_Delete'
$ctxSep1    = [System.Windows.Controls.Separator]::new()
$ctxViewImg = [System.Windows.Controls.MenuItem]::new(); $ctxViewImg.Header = '_View Image'
$ctxSep2    = [System.Windows.Controls.Separator]::new()
$ctxCopyRow = [System.Windows.Controls.MenuItem]::new(); $ctxCopyRow.Header = '_Copy Row to Clipboard'
foreach ($item in @($ctxEdit,$ctxCloneM,$ctxDeleteM,$ctxSep1,$ctxViewImg,$ctxSep2,$ctxCopyRow)) {
    $ctxMenu.Items.Add($item) | Out-Null
}
$dgCoins.ContextMenu = $ctxMenu

$ctxEdit.Add_Click({
    $sel = Get-SelectedCoin
    if ($null -ne $sel) { Populate-CoinForm -Coin $sel }
})
$ctxCloneM.Add_Click({
    $sel = Get-SelectedCoin
    if ($null -eq $sel) { return }
    Populate-CoinForm -Coin $sel
    $script:EditingCoinID = $null
    $lblFormTitle.Text = 'Add New Coin (Clone)'
    $lblStatus.Text = 'Clone ready — edit fields then Save'
})
$ctxDeleteM.Add_Click({
    $btnDelete.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
})
$ctxViewImg.Add_Click({
    $sel = Get-SelectedCoin
    if ($null -eq $sel) { return }
    $imgPath = $sel['ImagePath']
    if ([string]::IsNullOrWhiteSpace($imgPath) -or -not (Test-Path $imgPath)) {
        $imgPath = Get-CoinTypeImage -Denomination $sel['Denomination']
    }
    Show-ImageViewer -ImagePath $imgPath -Title "$($sel['Year']) $($sel['Mint']) $($sel['Denomination'])"
})
$ctxCopyRow.Add_Click({
    $sel = Get-SelectedCoin
    if ($null -eq $sel) { return }
    $text = "$($sel.ID)`t$($sel.Year)`t$($sel.Mint)`t$($sel.Denomination)`t$($sel.Type)`t$($sel.Grade)`t$($sel.BookValue)`t$($sel.Location)"
    [System.Windows.Clipboard]::SetText($text)
    $lblStatus.Text = 'Row copied to clipboard'
})

# ---------------------------------------------------------------------------
# Clone Coin Button
# ---------------------------------------------------------------------------
$btnClone.Add_Click({
    $sel = Get-SelectedCoin
    if ($null -eq $sel) {
        [System.Windows.MessageBox]::Show('Select a coin to clone.', 'Clone') | Out-Null; return
    }
    Populate-CoinForm -Coin $sel
    $script:EditingCoinID = $null
    $lblFormTitle.Text = 'Add New Coin (Clone)'
    $lblStatus.Text = 'Clone ready — edit fields then Save'
})

# ---------------------------------------------------------------------------
# Export CSV Button
# ---------------------------------------------------------------------------
$btnExportCsv.Add_Click({
    $coins = @(Get-Collection)
    if ($coins.Count -eq 0) {
        [System.Windows.MessageBox]::Show('Collection is empty.', 'Export CSV') | Out-Null; return
    }
    $sfd = [Microsoft.Win32.SaveFileDialog]::new()
    $sfd.Title      = 'Export Collection as CSV'
    $sfd.Filter     = 'CSV Files (*.csv)|*.csv|All Files (*.*)|*.*'
    $sfd.FileName   = "$($script:ActiveCollection)_$(Get-Date -Format 'yyyyMMdd').csv"
    if (-not $sfd.ShowDialog()) { return }
    $coins | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8
    $open = [System.Windows.MessageBox]::Show(
        "Saved $($coins.Count) coins to:`n$($sfd.FileName)`n`nOpen now?",
        'Export CSV Complete', 'YesNo', 'Information')
    if ($open -eq 'Yes') { Start-Process $sfd.FileName }
    $lblStatus.Text = "CSV exported: $($sfd.FileName)"
})

# ---------------------------------------------------------------------------
# Restore Backup Button
# ---------------------------------------------------------------------------
$btnRestore.Add_Click({
    $result = Show-RestoreDialog
    if ($result) {
        Refresh-CollectionDropdown
        Refresh-Grid
        Clear-CoinForm
        $lblStatus.Text = 'Collection restored from backup'
    }
})

# ---------------------------------------------------------------------------
# Undo Last Delete
# ---------------------------------------------------------------------------
$btnUndo.Add_Click({
    if ($null -eq $script:LastDeleted) { $btnUndo.Visibility = 'Collapsed'; return }
    $coins = [System.Collections.Generic.List[object]]::new(@(Get-Collection))
    $coins.Add($script:LastDeleted)
    Save-Collection $coins.ToArray()
    $script:LastDeleted = $null
    $btnUndo.Visibility = 'Collapsed'
    Refresh-Grid
    $lblStatus.Text = 'Delete undone — coin restored'
})

# ---------------------------------------------------------------------------
# Keyboard Shortcuts
# ---------------------------------------------------------------------------
$mainWin.Add_KeyDown({
    param($s, $e)
    $mod = [System.Windows.Input.Keyboard]::Modifiers
    if ($mod -eq 'Ctrl') {
        switch ($e.Key) {
            'S' { $btnSaveCoin.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent)); $e.Handled = $true }
            'N' { $btnAdd.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent));      $e.Handled = $true }
            'Z' { if ($btnUndo.Visibility -eq 'Visible') { $btnUndo.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent)) }; $e.Handled = $true }
        }
    }
    if ($mod -eq 'None') {
        switch ($e.Key) {
            'Escape' { $btnClearForm.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent)); $e.Handled = $true }
            'F5'     { Refresh-Grid; $e.Handled = $true }
        }
    }
})

# ---------------------------------------------------------------------------
# Live Search (debounced — fires 400 ms after last keystroke)
# ---------------------------------------------------------------------------
$script:SearchTimer = [System.Windows.Threading.DispatcherTimer]::new()
$script:SearchTimer.Interval = [TimeSpan]::FromMilliseconds(400)
$script:SearchTimer.Add_Tick({
    $script:SearchTimer.Stop()
    $all    = @(Get-Collection)
    $fYear  = $tbFilterYear.Text.Trim()
    $fMint  = $tbFilterMint.Text.Trim().ToUpper()
    $fDenom = $tbFilterDenom.Text.Trim()
    $fGrade = $tbFilterGrade.Text.Trim()
    $fMetal = $tbFilterMetal.Text.Trim()
    $fKey   = $tbFilterKey.Text.Trim()
    $anyFilter = $fYear -or $fMint -or $fDenom -or $fGrade -or $fMetal -or $fKey
    if (-not $anyFilter) { Refresh-Grid; return }
    $results = @($all | Where-Object {
        ($fYear  -eq '' -or $_.Year         -eq $fYear)        -and
        ($fMint  -eq '' -or $_.Mint         -like "*$fMint*")  -and
        ($fDenom -eq '' -or $_.Denomination -like "*$fDenom*") -and
        ($fGrade -eq '' -or $_.Grade        -like "*$fGrade*") -and
        ($fMetal -eq '' -or $_.Metal        -like "*$fMetal*") -and
        ($fKey   -eq '' -or $_.KeyDates     -like "*$fKey*")
    })
    Refresh-Grid -Data $results
    $lblStatus.Text = "Live search: $($results.Count) result$(if ($results.Count -ne 1){'s'})"
})
foreach ($tb in @($tbFilterYear, $tbFilterMint, $tbFilterDenom, $tbFilterGrade, $tbFilterMetal, $tbFilterKey)) {
    $tb.Add_TextChanged({ $script:SearchTimer.Stop(); $script:SearchTimer.Start() })
}

# ---------------------------------------------------------------------------
# Drag-and-Drop Image onto Preview
# ---------------------------------------------------------------------------
$imgPreview.Add_DragOver({
    param($s, $e)
    if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
        $e.Effects = [System.Windows.DragDropEffects]::Copy
    } else {
        $e.Effects = [System.Windows.DragDropEffects]::None
    }
    $e.Handled = $true
})
$imgPreview.Add_Drop({
    param($s, $e)
    $files = $e.Data.GetData([System.Windows.DataFormats]::FileDrop)
    if ($null -eq $files -or $files.Count -eq 0) { return }
    $file = $files[0]
    $ext  = [System.IO.Path]::GetExtension($file).ToLower()
    if ($ext -notin @('.jpg','.jpeg','.png','.bmp','.gif')) {
        $lblStatus.Text = 'Drop rejected — not an image file'; return
    }
    $tbImagePath.Text = $file
    try {
        $bi = [System.Windows.Media.Imaging.BitmapImage]::new()
        $bi.BeginInit()
        $bi.UriSource = [Uri]::new($file, [UriKind]::Absolute)
        $bi.DecodePixelWidth = 100
        $bi.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bi.EndInit(); $bi.Freeze()
        $imgPreview.Source = $bi
    } catch { }
    $lblStatus.Text = "Image set: $([System.IO.Path]::GetFileName($file))"
})

# Right-click preview to clear the image
$imgPreview.Add_MouseRightButtonUp({
    $tbImagePath.Text  = ''
    $imgPreview.Source = $null
    $lblStatus.Text    = 'Image cleared'
})

# ---------------------------------------------------------------------------
# Initialization & Launch  (show window fast, load data after)
# ---------------------------------------------------------------------------
Initialize-Collections
Clear-CoinForm
$lblStatus.Text    = 'Loading collection...'
$lblCoinCount.Text = '...'

# Show window immediately, then load data via Dispatcher for responsive UI
$mainWin.Add_ContentRendered({
    $mainWin.Dispatcher.Invoke([Action]{
        try {
            Invoke-Backup | Out-Null
            Refresh-CollectionDropdown
            # Load collection once and reuse for both grid and snapshot
            $_coins = @(Get-Collection)
            Refresh-Grid -Data $_coins
            # Auto-record value snapshot for history tracking
            try { Record-ValueSnapshot -Coins $_coins } catch { }
        } catch {
            $lblStatus.Text = "Load error: $($_.Exception.Message)"
        }
    }, [System.Windows.Threading.DispatcherPriority]::Background)
})

try {
    $ErrorActionPreference = 'Continue'
    $null = $mainWin.ShowDialog()
} finally {
    $ErrorActionPreference = 'Stop'
}

# SIG # Begin signature block
# MIIFagYJKoZIhvcNAQcCoIIFWzCCBVcCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUby8jV44WwjbNxQvePgQQGrH5
# s5WgggMGMIIDAjCCAeqgAwIBAgIQfT/Qb+FrZJxFNj8fYAcaIzANBgkqhkiG9w0B
# AQsFADAZMRcwFQYDVQQDDA5Db2luQ29sbGVjdGlvbjAeFw0yNjAyMjYxMDM2MzBa
# Fw0yNzAyMjYxMDU2MzBaMBkxFzAVBgNVBAMMDkNvaW5Db2xsZWN0aW9uMIIBIjAN
# BgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA4p5lGsblyI/UryEqIMNLFKFbC2a4
# Lb7z4kyvGRRowta2Uxz+6dcFJWZT/j1jXIZlH+Qsc54N1dS0temkhW34LaJXMCSi
# XMw5EhNQgS1r72iD7EjvulkkhvawpxEEgSaPPZGlfj7ESWg77ziqGJcfI1V6aECt
# tkkHzQktFWEmrTDmIDOex2dmXjKSWF5HWcAIZpXPQ8R6R+hCmSZ/rD0EeiHWoKGg
# UqS2Oc206JZ6T6z6ayd+pk9xxkjfBDVx0lQXLUdTnryMg8Ga9889Wb9XcsSA92hH
# oNdCxyPAcGpXgyx/eOeEh3MIJ6pqp0x9i+ijjdEbt2nUsC9OJRvpwxd/PQIDAQAB
# o0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwHQYDVR0O
# BBYEFOe/3dNhbCmrzjO1ZghqqDuTkXvvMA0GCSqGSIb3DQEBCwUAA4IBAQB8cTqb
# hfiFosxveJ5hFrj/4oYjT8oiN7lYjVhRzfhRtiF9vnZxaWxr/oQXhrgNlLYw6VP9
# 2jv48fzjKxfd670ErGQOjRth9Rmb6fAFhDLyFWc6gSvAG3KoGUBsUxdoAwbsJHzF
# jcppEsSFoeBSVFBbtX+IYIJr6i+W7o2wIcXj5G6lprgbWvNJpII4a9YEHP/6D+Fu
# opp5aieoA/HCYiEVjlqXROla1Lvf0ijZ83u2+Vs79nk5yRWi1qKYeEG6azvOtOFW
# QORVdbHjtpI2xPoQrN3XcHtvsJT9qTIMT8lvZPsx9gJj2dDqEzSR8kglDVKCXC3z
# efTqtiglf/ILiweXMYIBzjCCAcoCAQEwLTAZMRcwFQYDVQQDDA5Db2luQ29sbGVj
# dGlvbgIQfT/Qb+FrZJxFNj8fYAcaIzAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIB
# DDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEE
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUiMuhUkr3+pQT
# 9da3wHvZ1sPrnHkwDQYJKoZIhvcNAQEBBQAEggEAeDTPyV+IYoSWoCS8wHZ588cz
# AnJzAEqdZKUmbVSc1iFQDsmOqq/yQT0nC6q8UwrskDw6xYk/fLFKNVzDyFzxiFrF
# Nitz7TAqqFbobhVX6ZJGvIP8OcRddigxG8sprbilMEyDjDQFd/E8JrvX5pL271kh
# vPn3GsDWlQ9H8kuagaSnV5MT/ocdMdZHPGflTHHgmld3yiIL8HFT0jmkEgNdiM2H
# P/HqEGeDoD1hVjuCo1a/t0NnhUP1qUaM++GcYiciH3fhIUivozqeS5DXAmeBhbMP
# 028zM7Gi+X/zjPBeJQPtHO72QXyoMnGT5p0lYVWhV2KVOrvckCPjxQa4Zks6zw==
# SIG # End signature block
