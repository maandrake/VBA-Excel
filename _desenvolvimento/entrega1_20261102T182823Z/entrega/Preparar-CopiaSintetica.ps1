[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][switch]$ConfirmarLaboratorioSintetico,
    [string]$Pacote = $PSScriptRoot
)
$ErrorActionPreference = 'Stop'
if (-not $ConfirmarLaboratorioSintetico) { throw 'Exige confirmacao explicita de copia sintetica.' }
$Pacote = (Resolve-Path -LiteralPath $Pacote).Path
$manifesto = Get-Content -LiteralPath (Join-Path $Pacote 'manifesto_artefatos.json') -Raw | ConvertFrom-Json
foreach ($item in $manifesto.files) {
    if ($item.name -notin @('Estoque_Fera_Veiculos_Candidato_Integrado.xlsm','Banco_Candidato_Sintetico.accdb')) { throw 'Nome inesperado no manifesto.' }
    $origem = Join-Path $Pacote $item.name
    if ((Get-FileHash -LiteralPath $origem -Algorithm SHA256).Hash -ne $item.sha256) { throw "Hash divergente: $($item.name)" }
}
$destino = Join-Path ([IO.Path]::GetTempPath()) ('Entrega1_Ensaio_' + [Guid]::NewGuid().ToString('N'))
$raizTemp = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
$destino = [IO.Path]::GetFullPath($destino)
if (-not $destino.StartsWith($raizTemp,[StringComparison]::OrdinalIgnoreCase) -or $destino -match 'OneDrive') { throw 'Destino fora do TEMP isolado.' }
if (Test-Path -LiteralPath $destino) { throw 'Destino ja existe; nada sera sobrescrito.' }
[void][IO.Directory]::CreateDirectory($destino)
foreach ($item in $manifesto.files) {
    Copy-Item -LiteralPath (Join-Path $Pacote $item.name) -Destination (Join-Path $destino $item.name) -ErrorAction Stop
}
foreach ($nomeXml in @('valido.xml','transferencia_b.xml','transferencia_b2.xml','transferencia_duplicada.xml')) {
    $xml = Join-Path $Pacote $nomeXml
    if (Test-Path -LiteralPath $xml) { Copy-Item -LiteralPath $xml -Destination (Join-Path $destino $nomeXml) }
}
$banco = Join-Path $destino 'Banco_Candidato_Sintetico.accdb'
$conexao = New-Object -ComObject ADODB.Connection
try {
    $conexao.Open('Provider=Microsoft.ACE.OLEDB.16.0;Data Source=' + $banco)
    function Ler-Escalar([string]$sql) {
        $r = $conexao.Execute($sql)
        try { if ($r.EOF) { throw 'Consulta sem resultado.' }; return $r.Fields.Item(0).Value } finally { $r.Close() }
    }
    if ((Ler-Escalar 'SELECT SchemaNo FROM Meta WHERE ID=1') -ne 5) { throw 'Esquema inesperado.' }
    if ((Ler-Escalar 'SELECT StatusText FROM WriterLease WHERE ID=1') -ne 'FREE') { throw 'Lease ocupado: conciliacao manual obrigatoria; nao ha desbloqueio automatico.' }
    if ((Ler-Escalar 'SELECT Count(*) FROM Pending') -ne 0) { throw 'Base possui operacoes pendentes. Preservar e conciliar.' }
    if ((Ler-Escalar 'SELECT WriteState FROM Meta WHERE ID=1') -ne 'READY') { throw 'Base nao esta READY.' }
    $baseId = Ler-Escalar 'SELECT BaseID FROM Meta WHERE ID=1'
    if ($baseId -ne $manifesto.baseId) { throw 'Linhagem divergente.' }
    # Vinculacao explicita apenas da nova copia sintetica. Os fatos e IDs sao preservados.
    $comando = New-Object -ComObject ADODB.Command
    $comando.ActiveConnection = $conexao
    $comando.CommandText = "UPDATE CandidateSettings SET AuthorityPath=?,WorkbookID=?,ProjectionState='PENDING' WHERE ID=1 AND BaseID=?"
    [void]$comando.Parameters.Append($comando.CreateParameter('caminho',202,1,1024,$banco))
    [void]$comando.Parameters.Append($comando.CreateParameter('workbook',202,1,40,[Guid]::NewGuid().ToString().ToUpperInvariant()))
    [void]$comando.Parameters.Append($comando.CreateParameter('base',202,1,40,$baseId))
    $conexao.BeginTrans() | Out-Null
    try { $comando.Execute() | Out-Null; $conexao.CommitTrans() | Out-Null } catch { $conexao.RollbackTrans() | Out-Null; throw }
    [ordered]@{Mode='COPIA_SINTETICA_INDEPENDENTE_NAO_OPERACIONAL';SourcePackage=$Pacote;BaseID=$baseId;AuthorityPath=$banco;CreatedUTC=[DateTime]::UtcNow.ToString('o');NextAction='Abrir XLSM e executar CandidatoReconstruir; sem mudar seguranca do Office.'} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $destino 'vinculacao_laboratorio.json') -Encoding UTF8
} finally {
    if ($conexao.State -ne 0) { $conexao.Close() }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($conexao)
}
Write-Output $destino
Write-Output 'A copia nao e uma segunda autoridade operacional. Usar exclusivamente no laboratorio sintetico.'
