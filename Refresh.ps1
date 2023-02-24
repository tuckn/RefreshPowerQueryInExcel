# File Encoding must be SJIS or UTF-8 with BOM.

Param(
    [Parameter(Position = 0, Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ })]
    [String] $BookPath,

    [Parameter(Position = 1)]
    [Switch] $ShowsWindow
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

# Setting IsReadOnly property to $False
Get-Item -Path $BookPath | Set-ItemProperty -Name IsReadOnly -Value $false

$excel = New-Object -ComObject Excel.Application
$excel.Visible = !($ShowsWindow)
# $excel.DisplayAlerts = $false

$book = $excel.Workbooks.Open($BookPath, $null, $false, $true)
$book.refreshall()
# $book.Unprotect() # Not working
$book.Save()
# $book.Protect()
$book.Close()

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)

# Setting IsReadOnly property to $True
Get-Item -Path $BookPath | Set-ItemProperty -Name IsReadOnly -Value $true