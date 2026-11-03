param(
    [string]$SourceRoot = (Join-Path $PSScriptRoot '..\desenvolvimento'),
    [string]$TestRoot = (Join-Path $env:TEMP ('VBA_Auditoria_' + [guid]::NewGuid().ToString('N'))),
    [switch]$RunTests
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# Nao altera ExecutionPolicy, AccessVBOM, Trust Center ou seguranca de macros.
# O XLSM original nunca e aberto. Todas as planilhas sao novas e ficticias.
$sourcePath = (Resolve-Path -LiteralPath $SourceRoot).Path
$testPath = [IO.Path]::GetFullPath($TestRoot)
$tempPrefix = [IO.Path]::GetFullPath($env:TEMP).TrimEnd('\') + '\'
if (-not $testPath.StartsWith($tempPrefix, [StringComparison]::OrdinalIgnoreCase) -or
    $testPath -match '(?i)OneDrive') { throw 'O destino deve estar dentro de TEMP, fora do OneDrive.' }
if (Test-Path -LiteralPath $testPath) { throw 'O destino ja existe. Escolha uma pasta nova.' }
[void](New-Item -ItemType Directory -Path $testPath)
$reportPath = Join-Path $testPath 'resultado_preparacao.json'
$result = [ordered]@{ Timestamp = (Get-Date).ToString('o'); Status = 'STARTING'; TestRoot = $testPath }
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $result.ExcelVersion = $excel.Version
    $result.AutomationSecurityOriginal = $excel.AutomationSecurity
    $book = $excel.Workbooks.Add()
    # A leitura direta pode retornar null sem erro no PowerShell. Reflection
    # conserva a mensagem real do Office e impede um falso sucesso de importacao.
    $project = $book.GetType().InvokeMember('VBProject', [Reflection.BindingFlags]::GetProperty, $null, $book, $null)
    if ($null -eq $project) { throw 'VBProject indisponivel; nenhuma protecao sera alterada.' }
    $result.References = @($project.References | ForEach-Object {
        [ordered]@{Name=$_.Name; IsBroken=$_.IsBroken; Guid=$_.Guid; Major=$_.Major; Minor=$_.Minor}
    })
    foreach ($folder in @('MDL', 'Form')) {
        $filter = if ($folder -eq 'MDL') {'*.bas'} else {'*.frm'}
        foreach ($file in Get-ChildItem -LiteralPath (Join-Path $sourcePath $folder) -Filter $filter) {
            [void]$project.VBComponents.Import($file.FullName)
        }
    }
    [void]$project.VBComponents.Import((Join-Path $PSScriptRoot 'AuditoriaListaFake.cls'))
    [void]$project.VBComponents.Import((Join-Path $PSScriptRoot 'modTestesAuditoria.bas'))
    $names = @('Menu Principal','Entradas','Transferidos','Entradas Anual','Demonstração',
        'Saidas','Todas Saidas','Entradas Filial','Entradas Anual Filial','Saidas Filial',
        'Todas Saidas Filial','Carros PF','Clientes','Veículos','Doc. para fora','NCM e CEST','Listas')
    $book.Worksheets.Item(1).Name = $names[0]
    foreach ($name in $names[1..($names.Count - 1)]) {
        $sheet = $book.Worksheets.Add()
        $sheet.Name = $name
    }
    $book.Worksheets.Item('Menu Principal').Range('Z1').Value2 = 'AUDITORIA_VBA_FICTICIO'
    $fixture = Join-Path $testPath ('Fixture_Auditoria_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.xlsm')
    $book.SaveAs($fixture, 52)
    $result.Fixture = $fixture
    $result.Status = 'PREPARED_NOT_COMPILED'
    if ($RunTests) {
        # Opcional: pode revelar erros de compilacao do Office. Sem RunTests,
        # a preparacao nunca executa procedimentos VBA do projeto.
        $result.Tests = $excel.Run(("'" + $book.Name + "'!Auditoria_ExecutarTestesSeguros"))
        $result.Status = 'TESTS_EXECUTED'
    }
} catch {
    $result.Status = 'BLOCKED_OR_FAILED'
    $result.Error = $_.Exception.Message
    $result.HResult = $_.Exception.HResult
} finally {
    $result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $reportPath -Encoding UTF8
    if ($null -ne $book) { try { $book.Close($false) } catch {} }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
}
$result | ConvertTo-Json -Depth 8
if ($result.Status -eq 'BLOCKED_OR_FAILED') { exit 1 }
