param([Parameter(Mandatory=$true)][string]$DocumentPath)
$ErrorActionPreference = 'Stop'
$resolvedReport = (Resolve-Path -LiteralPath $DocumentPath).Path
$pdfReport = [IO.Path]::ChangeExtension($resolvedReport, '.pdf')
$wordApp = $null
$wordDoc = $null
try {
    $wordApp = New-Object -ComObject Word.Application
    $wordApp.Visible = $false
    $wordApp.DisplayAlerts = 0
    $wordDoc = $wordApp.Documents.Open($resolvedReport, $false, $false)
    $null = $wordDoc.Fields.Update()
    foreach ($toc in $wordDoc.TablesOfContents) { $toc.Update() }
    $wordDoc.Repaginate()
    foreach ($toc in $wordDoc.TablesOfContents) { $toc.UpdatePageNumbers() }
    $wordDoc.Save()
    $wordDoc.ExportAsFixedFormat($pdfReport, 17)
    Write-Output ('Pages: ' + $wordDoc.ComputeStatistics(2))
    Write-Output $pdfReport
} finally {
    if ($null -ne $wordDoc) { $wordDoc.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordDoc) }
    if ($null -ne $wordApp) { $wordApp.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordApp) }
}
