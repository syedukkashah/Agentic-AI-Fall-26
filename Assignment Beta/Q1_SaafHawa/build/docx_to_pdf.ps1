# Convert .docx files to PDF with Microsoft Word.
# Usage (from Q1_SaafHawa):  powershell -File build/docx_to_pdf.ps1 q1_report.docx [more.docx ...]
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    foreach ($f in $args) {
        $in = (Resolve-Path $f).Path
        $out = [System.IO.Path]::ChangeExtension($in, ".pdf")
        $doc = $word.Documents.Open($in, $false, $true)
        $doc.SaveAs2($out, 17)   # 17 = wdFormatPDF
        $doc.Close($false)
        Write-Output "wrote $out"
    }
} finally {
    $word.Quit()
}
