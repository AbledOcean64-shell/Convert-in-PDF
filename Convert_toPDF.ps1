$folder = "."

Write-Host "=== Conversor para PDF - Versão Estável ===" -ForegroundColor Cyan
Write-Host "Pasta: $folder" -ForegroundColor White

if (-not (Test-Path $folder)) {
    Write-Host "ERRO: Pasta não encontrada!" -ForegroundColor Red
    pause
    exit
}

$files = Get-ChildItem -Path $folder -File | 
         Where-Object { $_.Extension -match '\.(doc|docx|xls|xlsx|xlsm|ppt|pptx|rtf)$' }

if ($files.Count -eq 0) {
    Write-Host "Nenhum arquivo encontrado." -ForegroundColor Yellow
    pause
    exit
}

Write-Host "$($files.Count) arquivo(s) encontrado(s)..." -ForegroundColor Green

$word = $null
$excel = $null
$ppt = $null

function Convert-With-Retry {
    param($File, $MaxAttempts = 3)
    
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            if ($attempt -gt 1) {
                Write-Host "  Tentativa $attempt de $MaxAttempts..." -ForegroundColor Yellow
                Start-Sleep -Seconds 2
            }

            $pdfName = [System.IO.Path]::ChangeExtension($File.FullName, ".pdf")

            switch -Regex ($File.Extension.ToLower()) {
                '\.(doc|docx|rtf)$' {
                    if (-not $word) {
                        $word = New-Object -ComObject Word.Application
                        $word.Visible = $false
                    }
                    $doc = $word.Documents.Open($File.FullName)
                    $doc.SaveAs([ref]$pdfName, [ref]17)
                    $doc.Close()
                }
                '\.(xls|xlsx|xlsm)$' {
                    if (-not $excel) {
                        $excel = New-Object -ComObject Excel.Application
                        $excel.Visible = $false
                        $excel.DisplayAlerts = $false
                        $excel.ScreenUpdating = $false
                    }
                    $wb = $excel.Workbooks.Open($File.FullName, $false, $true)
                    $wb.ExportAsFixedFormat(0, $pdfName)
                    $wb.Close($false)
                }
                '\.(ppt|pptx)$' {
                    if (-not $ppt) {
                        $ppt = New-Object -ComObject PowerPoint.Application
                        $ppt.Visible = $false
                    }
                    $pres = $ppt.Presentations.Open($File.FullName)
                    $pres.SaveAs($pdfName, 32)
                    $pres.Close()
                }
            }
            return $true
        }
        catch {
            if ($attempt -eq $MaxAttempts) {
                Write-Host " → ERRO após $MaxAttempts tentativas: $($_.Exception.Message)" -ForegroundColor Red
                return $false
            }
        }
    }
}

foreach ($file in $files) {
    Write-Host "Convertendo: $($file.Name)" -NoNewline
    $success = Convert-With-Retry -File $file
    
    if ($success) {
        Write-Host " → OK" -ForegroundColor Green
    }
}

# Fecha aplicativos
if ($word) { try { $word.Quit() } catch {} }
if ($excel) { try { $excel.Quit() } catch {} }
if ($ppt) { try { $ppt.Quit() } catch {} }

Write-Host "`nConversão finalizada!" -ForegroundColor Green
pause
