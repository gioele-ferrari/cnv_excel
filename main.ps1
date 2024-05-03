Param(
    $inputFile,
    $outFile
)

# Combinare il percorso dello script con il nome del file Excel
$currentDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$excelFile = Join-Path -Path $currentDirectory -ChildPath $inputFile

# Percorso di output del file di testo
$textFile = $outFile

# Crea un oggetto Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

# Apri il file Excel
$workbook = $excel.Workbooks.Open($excelFile)
$worksheet = $workbook.Sheets.Item(1)
$range = $worksheet.UsedRange

# Ottieni il numero di righe e colonne
$rows = $range.Rows.Count
$columns = $range.Columns.Count

# Apri il file di testo in modalità di scrittura
$output = [System.IO.File]::CreateText($textFile)

# Ciclo per leggere e scrivere i dati
for ($row = 1; $row -le $rows; $row++) {
    $line = ""
    for ($col = 1; $col -le $columns; $col++) {
        $value = $range.Cells.Item($row, $col).Text
        $line += $value
        if ($col -lt $columns) {
            $line += ","
        }
    }
    $output.WriteLine($line)
}

# Chiudi il file di testo e Excel
$output.Close()
$workbook.Close()
$excel.Quit()

# Rilascia le risorse COM
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Remove-Variable excel

Write-Host "Conversione completata."
