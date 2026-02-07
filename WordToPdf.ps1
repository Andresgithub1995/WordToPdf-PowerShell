# =========================
# CONFIGURACIÓN (solo cambias esto)
# =========================

$folder  = "C:\Users\andre\OneDrive - UNIR\2026 - 1\Seguridad en los Sistemas de Informacion"
$docName = "Actividad2.docx"

# =========================
# LÓGICA (no tocar)
# =========================

$inputDoc  = Join-Path $folder $docName
$outputPdf = [System.IO.Path]::ChangeExtension($inputDoc, "pdf")

if (-not (Test-Path $inputDoc)) {
    throw "No existe el archivo: $inputDoc"
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

$doc = $null

try {
    $doc = $word.Documents.Open($inputDoc, $false, $true)
    $doc.ExportAsFixedFormat($outputPdf, 17)
    $doc.Close($false)
}
finally {
    if ($doc)  { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null }
    if ($word) { 
        $word.Quit() | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null 
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "✅ PDF creado en:"
Write-Host $outputPdf
