$dir = $PSScriptRoot
$files = @(
    "Bao_cao_Giua_ki_Springer_Dinh_Van_Lam",
    "Bao_cao_Cuoi_ki_DNC_309_Dinh_Van_Lam"
)

foreach ($name in $files) {
    $docxPath = Join-Path $dir "$name.docx"
    $pdfPath = Join-Path $dir "$name.pdf"
    
    Write-Output "Processing: $docxPath"
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    try {
        $doc = $word.Documents.Open($docxPath)
        $doc.ExportAsFixedFormat($pdfPath, 17) # 17 = wdExportFormatPDF
        $doc.Close(0) # 0 = wdDoNotSaveChanges
        Write-Output "Generated: $pdfPath"
    } catch {
        Write-Error $_
    } finally {
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}
