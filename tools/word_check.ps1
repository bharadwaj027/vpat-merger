# Opens a .docx in Microsoft Word (hidden), reports page count / table count,
# and exports a PDF so the layout can be inspected. Word shows a repair prompt
# (and fails to open in read-only, no-repair mode) when the package is invalid.
param([Parameter(Mandatory = $true)][string]$Path, [string]$Pdf)

$full = (Resolve-Path $Path).Path
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate,
    #      Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible, OpenAndRepair, ..., NoEncodingDialog)
    $doc = $word.Documents.Open($full, $false, $true, $false, "", "", $false, "", "", 0, 0, $false, $false, 0, $true)
    "Opened: $full"
    "Pages: " + $doc.ComputeStatistics(2)
    "Tables: " + $doc.Tables.Count
    "Paragraphs: " + $doc.Paragraphs.Count
    if ($Pdf) {
        $doc.ExportAsFixedFormat($Pdf, 17)
        "PDF: $Pdf"
    }
    $doc.Close($false)
}
catch {
    "FAILED to open: $($_.Exception.Message)"
    exit 1
}
finally {
    $word.Quit()
}
