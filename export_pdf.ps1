param(
    [string]$HtmlPath = ".\output\relatorio_atendimentos.html",
    [string]$PdfPath = ".\output\relatorio_atendimentos.pdf"
)

$resolvedHtml = Resolve-Path -LiteralPath $HtmlPath -ErrorAction Stop
$targetPdf = [System.IO.Path]::GetFullPath($PdfPath)

$browserCandidates = @(
    "$env:ProgramFiles(x86)\Microsoft\Edge\Application\msedge.exe",
    "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
    "$env:ProgramFiles(x86)\Google\Chrome\Application\chrome.exe",
    "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
)

$browser = $browserCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $browser) {
    Write-Error "Nenhum Edge ou Chrome compativel foi encontrado. Abra o HTML no navegador e use Imprimir > Salvar como PDF."
    exit 1
}

$htmlUri = [System.Uri]::new($resolvedHtml.Path).AbsoluteUri

& $browser --headless --disable-gpu "--print-to-pdf=$targetPdf" $htmlUri

if ($LASTEXITCODE -ne 0) {
    Write-Error "Falha ao gerar o PDF."
    exit $LASTEXITCODE
}

Write-Output "PDF gerado em: $targetPdf"
