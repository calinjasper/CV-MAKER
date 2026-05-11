Param(
    [string]$SourcePath = "d:\PROGRAM\CV maker\Calin_Jasper_Resume_Improved.doc",
    [string]$OutBase = "d:\PROGRAM\CV maker\Calin_Jasper_Resume_Improved"
)

$wdFormatPDF = 17
$wdFormatXMLDocument = 12

Write-Host "Source:" $SourcePath
if (-not (Test-Path $SourcePath)) {
    Write-Error "Source file not found: $SourcePath"
    exit 1
}

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Open($SourcePath)
    $doc.SaveAs([ref]("$OutBase.docx"), [ref]$wdFormatXMLDocument)
    $doc.ExportAsFixedFormat("$OutBase.pdf", $wdFormatPDF)
    $doc.Close()
    $word.Quit()
    Write-Host "Generated:" "$OutBase.docx" "and" "$OutBase.pdf"
} catch {
    Write-Error "Conversion failed: $($_.Exception.Message)"
    exit 1
}
