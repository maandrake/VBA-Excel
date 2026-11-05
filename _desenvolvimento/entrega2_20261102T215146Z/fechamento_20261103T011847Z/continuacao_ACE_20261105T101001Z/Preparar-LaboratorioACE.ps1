param([Parameter(Mandatory)][string]$Checkpoint,[Parameter(Mandatory)][string]$Laboratorio,[Parameter(Mandatory)][string]$Evidencia)
$ErrorActionPreference='Stop'
$tempRoot=[IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'Temp'))+[IO.Path]::DirectorySeparatorChar
$lab=[IO.Path]::GetFullPath($Laboratorio)
if(-not $lab.StartsWith($tempRoot,[StringComparison]::OrdinalIgnoreCase) -or $lab -notmatch 'Entrega2_Restauracao_Controlada_ACE_' -or (Test-Path -LiteralPath $lab)){throw 'Exige laboratorio TEMP novo e exclusivo'}
$manifest=Get-Content -LiteralPath (Join-Path $Checkpoint 'sha256.json') -Raw | ConvertFrom-Json
foreach($f in $manifest){if((Get-FileHash -LiteralPath $f.copia -Algorithm SHA256).Hash -ne $f.sha256){throw 'Checkpoint divergente'}}
[void][IO.Directory]::CreateDirectory($lab)
$book=Join-Path $lab 'Estoque_Fera_Veiculos_Candidato_Entrega2_Final.xlsm'
$db=Join-Path $lab 'Banco_Candidato_Entrega2_Final.accdb'
$registry=Join-Path $lab 'Controle_Laboratorio_ACE.accdb'
foreach($f in $manifest){$target=if($f.copia.EndsWith('.xlsm')){$book}elseif((Split-Path $f.copia -Leaf) -like 'Banco_*'){$db}else{$registry};[IO.File]::Copy($f.copia,$target,$false)}
function Open-Db([string]$file){$c=New-Object -ComObject ADODB.Connection;$c.Open("Provider=Microsoft.ACE.OLEDB.16.0;Data Source=$file;");return ,$c}
function Exec([object]$c,[string]$sql,[object[]]$Values=@()){$q=New-Object -ComObject ADODB.Command;$q.ActiveConnection=$c;$q.CommandType=1;$q.CommandText=$sql;foreach($v in $Values){$type=if($v -is [int]){3}else{202};$size=if($type -eq 3){0}else{[Math]::Max(1,([string]$v).Length)};$q.Parameters.Append($q.CreateParameter('p',$type,1,$size,$v))};[void]$q.Execute()}
function Scalar([object]$c,[string]$sql){$r=$c.Execute($sql);try{return $r.Fields.Item(0).Value}finally{$r.Close()}}
$base=[guid]::NewGuid().ToString().ToUpperInvariant();$lineage=[guid]::NewGuid().ToString().ToUpperInvariant();$authority=[guid]::NewGuid().ToString().ToUpperInvariant();$workbook=[guid]::NewGuid().ToString().ToUpperInvariant()
$c=$null;$r=$null;$tx=$false;$rtx=$false
$tables=@('ReconciliationReceipts','ReconciliationPlans','Pending','SheetProjections','Receipts','ChangeLog','Attempts','Projections','Movements','Operations','Provenance','Episodes','Vehicles','ProjectionAcks')
try{
 $c=Open-Db $db
 if((Scalar $c 'SELECT RevisionNo FROM Meta WHERE ID=1') -ne 26 -or (Scalar $c 'SELECT SchemaNo FROM Meta WHERE ID=1') -ne 7 -or (Scalar $c 'SELECT Count(*) FROM Pending') -ne 0 -or (Scalar $c 'SELECT StatusText FROM WriterLease WHERE ID=1') -ne 'FREE'){throw 'Referencia sintetica inesperada; nao continuar'}
 $schema=Scalar $c 'SELECT SchemaHash FROM Meta WHERE ID=1'
 [void]$c.BeginTrans();$tx=$true
 Exec $c "UPDATE Episodes SET StatusText='NEW',CurrentPID=Null,CurrentEpisodeID=Null"
 foreach($table in $tables){Exec $c "DELETE FROM [$table]"}
 Exec $c "UPDATE Meta SET BaseID=?,LineageID=?,RevisionNo=0,WriteState='READY' WHERE ID=1" @($base,$lineage)
 Exec $c "UPDATE WriterLease SET OwnerID='',FenceNo=0,StatusText='FREE' WHERE ID=1"
 Exec $c "UPDATE CandidateSettings SET BaseID=?,AuthorityPath=?,WorkbookID=?,ProjectionState='PENDING' WHERE ID=1" @($base,$db,$workbook)
 Exec $c 'UPDATE AuthorityBinding SET AuthorityID=?,RegistryPath=?,Generation=1,WorkbookPath=? WHERE ID=1' @($authority,$registry,$book)
 $c.CommitTrans();$tx=$false
 foreach($table in $tables){if((Scalar $c "SELECT Count(*) FROM [$table]") -ne 0){throw "Tabela sintetica nao vazia: $table"}}
 if((Scalar $c 'SELECT SchemaHash FROM Meta WHERE ID=1') -ne $schema){throw 'Esquema divergente'}
 $c.Close()
 $r=Open-Db $registry;[void]$r.BeginTrans();$rtx=$true
 Exec $r 'DELETE FROM AuthorityHistory'
 Exec $r 'UPDATE AuthorityControl SET AuthorityID=?,BaseID=?,LineageID=?,Generation=1,DataPath=?,WorkbookPath=?,OwnerID=Null WHERE ID=1' @($authority,$base,$lineage,$db,$book)
 $canon="LABORATORIO_SINTETICO_NOVO|$base|$lineage|$authority|$db|$book|Codex|Conflito ACE; sem migracao operacional"
 $eventHash=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($canon))).ToLowerInvariant()
 Exec $r 'INSERT INTO AuthorityHistory (EventID,EventCanon,EventHash) VALUES (?,?,?)' @([guid]::NewGuid().ToString(),$canon,$eventHash)
 $r.CommitTrans();$rtx=$false;$r.Close()
 $hashes=@(foreach($file in @($book,$db,$registry)){[pscustomobject]@{arquivo=$file;sha256=(Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash}})
 [ordered]@{modalidade='PREPARACAO_SINTETICA_ADO_NAO_TESTE_VBA';dataUTC=[datetime]::UtcNow.ToString('o');BaseID=$base;LineageID=$lineage;AuthorityID=$authority;WorkbookID=$workbook;SchemaHash=$schema;SchemaNo=7;Revision=0;TabelasVazias=$tables;arquivos=$hashes;referenciaPreservada=$Checkpoint} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $Evidencia -Encoding utf8
 Write-Output $lab
}finally{if($tx -and $null -ne $c){$c.RollbackTrans()};if($rtx -and $null -ne $r){$r.RollbackTrans()};if($null -ne $c -and $c.State -ne 0){$c.Close()};if($null -ne $r -and $r.State -ne 0){$r.Close()}}
