param([Parameter(Mandatory)][string]$Registro,[Parameter(Mandatory)][string]$Laboratorio,[Parameter(Mandatory)][string]$Evidencia)
$ErrorActionPreference='Stop'
$raizTemp=[IO.Path]::GetFullPath((Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'Temp'))+[IO.Path]::DirectorySeparatorChar
$lab=(Get-Item -LiteralPath $Laboratorio).FullName
$registroAbsoluto=(Get-Item -LiteralPath $Registro).FullName
if(-not $lab.StartsWith($raizTemp,[StringComparison]::OrdinalIgnoreCase) -or -not $registroAbsoluto.StartsWith($raizTemp,[StringComparison]::OrdinalIgnoreCase) -or $lab -notmatch 'Entrega2_Restauracao_Controlada_'){throw 'Ensaio restrito ao laboratorio sintetico TEMP'}
$ativo=Join-Path $lab 'lock_ativo.flag'
$liberar=Join-Path $lab 'liberar_concorrencia.flag'
if((Test-Path -LiteralPath $ativo) -or (Test-Path -LiteralPath $liberar)){throw 'Ensaio ja iniciado; preserve as evidencias e nao repita automaticamente'}
$estado=[ordered]@{Modalidade='ADO_ACE_PROCESSO_EXTERNO_LOCK_TRANSACIONAL_REGISTRO';PID=$PID;InicioUTC=[DateTime]::UtcNow.ToString('o');Registro=$registroAbsoluto;Laboratorio=$lab;CommitExecutado=$false;SimulaDoisExcel=$false}
$c=New-Object -ComObject ADODB.Connection
$tx=$false
try{
 $c.CommandTimeout=2
 $c.Open("Provider=Microsoft.ACE.OLEDB.16.0;Data Source=$registroAbsoluto;")
 $r=$c.Execute('SELECT Generation,OwnerID FROM AuthorityControl WHERE ID=1')
 if($null -eq $r -or $r.EOF){throw 'Registro de autoridade ausente'}
 $estado.Geracao=$r.Fields.Item('Generation').Value
 $v=$r.Fields.Item('OwnerID').Value
 if($null -ne $v -and $v -isnot [DBNull] -and [string]$v -ne ''){throw 'Registro ja ocupado; nao recuperar automaticamente'}
 $r.Close()
 [void]$c.BeginTrans();$tx=$true
 [void]$c.Execute("UPDATE AuthorityControl SET OwnerID='QA_PROCESSO_EXTERNO_LOCK' WHERE ID=1 AND OwnerID Is Null")
 [IO.File]::WriteAllText($ativo,[DateTime]::UtcNow.ToString('o'))
 $limite=[DateTime]::UtcNow.AddSeconds(180)
 while(-not (Test-Path -LiteralPath $liberar) -and [DateTime]::UtcNow -lt $limite){Start-Sleep -Milliseconds 200}
 $estado.LiberacaoPeloTeste=Test-Path -LiteralPath $liberar
 $c.RollbackTrans();$tx=$false
 $r=$c.Execute('SELECT Generation,OwnerID FROM AuthorityControl WHERE ID=1')
 if($null -eq $r -or $r.EOF){throw 'Registro posterior ausente'}
 $estado.GeracaoFinal=$r.Fields.Item('Generation').Value
 $v=$r.Fields.Item('OwnerID').Value
 $estado.OwnerFinal=if($null -eq $v -or $v -is [DBNull]){$null}else{[string]$v}
 $r.Close()
 $estado.Resultado=if($estado.LiberacaoPeloTeste -and $null -eq $estado.OwnerFinal){'LOCK_LIBERADO_POR_ROLLBACK'}else{'ENSAIO_NAO_CONCLUIDO'}
}catch{$estado.Erro=$_.Exception.Message;throw}finally{
 if($tx){try{$c.RollbackTrans()}catch{}}
 if($c.State -ne 0){$c.Close()}
 $estado.FimUTC=[DateTime]::UtcNow.ToString('o')
 $estado | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $Evidencia -Encoding utf8
}
