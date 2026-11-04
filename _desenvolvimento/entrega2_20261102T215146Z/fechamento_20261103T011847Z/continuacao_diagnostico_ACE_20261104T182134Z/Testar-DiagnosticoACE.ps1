param([Parameter(Mandatory)][string]$Laboratorio)
$ErrorActionPreference = 'Stop'
$lab = (Get-Item -LiteralPath $Laboratorio).FullName
$temp = [IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'Temp')) + [IO.Path]::DirectorySeparatorChar
if (-not $lab.StartsWith($temp, [StringComparison]::OrdinalIgnoreCase) -or
    [IO.Path]::GetFileName($lab) -notlike 'Entrega2_Diagnostico_ACE_*') { throw 'Somente laboratorio descartavel TEMP' }
$file = Join-Path $lab 'Registro_Descartavel.accdb'
$output = Join-Path $lab 'resultado.json'
if (Test-Path -LiteralPath $output) { throw 'Resultado existente: nao repetir nem sobrescrever' }
$events = [Collections.Generic.List[object]]::new()
$checks = [Collections.Generic.List[object]]::new()
function Step([string]$Name, $Connection, [scriptblock]$Action) {
    $Connection.Errors.Clear()
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $entry = [ordered]@{ Etapa=$Name; Sucesso=$false }
    try { $entry.Retorno = & $Action; $entry.Sucesso=$true }
    catch {
        $entry.Excecao=$_.Exception.Message
        $entry.HResult=$_.Exception.HResult
        $entry.ADO=@(foreach ($e in $Connection.Errors) {
            [ordered]@{ Numero=$e.Number; Nativo=$e.NativeError; SQLState=$e.SQLState; Origem=$e.Source; Descricao=$e.Description }
        })
    }
    $entry.Segundos=$watch.Elapsed.TotalSeconds
    $events.Add($entry)
    return $entry
}
function Read-State($Connection) {
    $r=$Connection.Execute('SELECT * FROM AuthorityControl ORDER BY ID')
    try {
        $rows=@(while (-not $r.EOF) {
            $row=[ordered]@{}
            foreach ($field in $r.Fields) { $row[$field.Name]=if ($field.Value -is [DBNull]) { $null } else { $field.Value } }
            $row; $r.MoveNext()
        })
        return ConvertTo-Json -InputObject $rows -Depth 4 -Compress
    } finally { $r.Close() }
}
function Update-Owner($Connection, [string]$Owner) {
    $q=New-Object -ComObject ADODB.Command
    $q.ActiveConnection=$Connection
    $q.CommandType=1
    $q.CommandTimeout=2
    $q.CommandText='UPDATE AuthorityControl SET OwnerID=? WHERE ID=1 AND OwnerID Is Null'
    $q.Parameters.Append($q.CreateParameter('p0',202,1,$Owner.Length,$Owner))
    $n=0
    $r=$q.Execute([ref]$n)
    return [int]$n
}
function Check([string]$ID, [bool]$OK, [string]$Observed) {
    $checks.Add([ordered]@{ ID=$ID; Resultado=if ($OK) {'PASS'} else {'FAIL'}; Observado=$Observed })
}
$a=New-Object -ComObject ADODB.Connection
$b=New-Object -ComObject ADODB.Connection
$result=[ordered]@{ Modalidade='COM_ADO_ACE_DUAS_CONEXOES_UM_PROCESSO'; Banco=$file; AutoridadeOperacional=$false; VBAExecutado=$false; InicioUTC=[DateTime]::UtcNow.ToString('o') }
try {
    foreach ($c in @($a,$b)) { $c.CommandTimeout=2; $c.Open("Provider=Microsoft.ACE.OLEDB.16.0;Data Source=$file;") }
    $result.Antes=Read-State $a
    $result.Provedor=$a.Provider
    $r=$a.Execute('SELECT OwnerID FROM AuthorityControl WHERE ID=1')
    $free=$r.Fields.Item(0).Value -is [DBNull]; $r.Close()
    if (-not $free) { throw 'Copia ja ocupada: nao liberar automaticamente' }
    [void]$b.BeginTrans()
    $baseline=Step 'Sem_conflito_UPDATE' $b { Update-Owner $b 'QA_B' }
    $b.RollbackTrans()
    Check 'ACE_BASELINE' ($baseline.Sucesso -and $baseline.Retorno -eq 1) "Linhas=$($baseline.Retorno)"
    [void]$a.BeginTrans()
    if ((Update-Owner $a 'QA_A') -ne 1) { throw 'Primeiro escritor nao adquiriu a linha' }
    [void]$b.BeginTrans()
    $conflict=Step 'Conflito_UPDATE_parametrizado' $b { Update-Owner $b 'QA_B' }
    Check 'ACE_CONFLITO_RECUSADO' (-not $conflict.Sucesso) ($conflict | ConvertTo-Json -Depth 5 -Compress)
    if ($conflict.Sucesso) { throw 'Segundo escritor atualizou durante lock ativo' }
    # A tentativa seguinte e diagnostica, somente na copia descartavel, sem fatos gravados por B.
    $commit=Step 'Commit_depois_do_UPDATE_recusado' $b { $b.CommitTrans() }
    $cleanup=Step 'Rollback_depois_do_conflito' $b { $b.RollbackTrans() }
    $a.RollbackTrans()
    [void]$b.BeginTrans()
    $retry=Step 'Nova_transacao_apos_rollback_do_primeiro' $b { Update-Owner $b 'QA_B' }
    $b.RollbackTrans()
    Check 'ACE_NOVA_TRANSACAO' ($retry.Sucesso -and $retry.Retorno -eq 1) "Linhas=$($retry.Retorno)"
    [void]$a.BeginTrans()
    if ((Update-Owner $a 'QA_A') -ne 1) { throw 'Primeiro escritor nao adquiriu a linha na segunda sequencia' }
    [void]$b.BeginTrans()
    $directConflict=Step 'Sequencia_da_candidata_UPDATE_recusado' $b { Update-Owner $b 'QA_B' }
    $directCleanup=Step 'Sequencia_da_candidata_rollback_sem_commit' $b { $b.RollbackTrans() }
    Check 'ACE_CONFLITO_SEM_COMMIT' (-not $directConflict.Sucesso) ($directConflict | ConvertTo-Json -Depth 5 -Compress)
    Check 'ACE_LIMPEZA_SEM_COMMIT' $directCleanup.Sucesso ($directCleanup | ConvertTo-Json -Depth 5 -Compress)
    $a.RollbackTrans()
    $result.Depois=Read-State $a
    Check 'ACE_ESTADO_RESTAURADO' ($result.Antes -ceq $result.Depois) 'Todos os campos de AuthorityControl comparados'
} catch { $result.ErroFatal=$_.Exception.Message }
finally {
    foreach ($c in @($b,$a)) {
        if ($c.State -ne 0) { try { $c.RollbackTrans() } catch {}; $c.Close() }
    }
    $result.Etapas=@($events.ToArray()); $result.Verificacoes=@($checks.ToArray())
    $result.FimUTC=[DateTime]::UtcNow.ToString('o')
    $json=$result | ConvertTo-Json -Depth 10
    [IO.File]::WriteAllText($output,$json,[Text.UTF8Encoding]::new($false))
    $json
}
if ($result.Contains('ErroFatal') -or $checks.Count -ne 6 -or @($checks | Where-Object Resultado -eq 'FAIL').Count) { throw 'Ensaio incompleto ou reprovado; preservar resultado.json' }
