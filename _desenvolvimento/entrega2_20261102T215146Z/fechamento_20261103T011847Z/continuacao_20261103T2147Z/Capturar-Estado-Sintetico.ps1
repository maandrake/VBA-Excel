param([Parameter(Mandatory)][string]$Banco,[Parameter(Mandatory)][string]$Saida)
$ErrorActionPreference='Stop'
if (-not (Test-Path -LiteralPath $Banco -PathType Leaf)) { throw 'Banco ausente' }
$c=New-Object -ComObject ADODB.Connection
$c.Mode=1
$c.Open("Provider=Microsoft.ACE.OLEDB.16.0;Data Source=$Banco;Persist Security Info=False;")
$estado=[ordered]@{ UTC=[DateTime]::UtcNow.ToString('o'); Modalidade='LEITURA_ADO_BANCO_PERSISTIDO'; Banco=$Banco; Tabelas=[ordered]@{} }
try {
 foreach($t in @('Meta','WriterLease','CandidateSettings','AuthorityBinding','Operations','Episodes','Vehicles','Movements','Projections','SheetProjections','Receipts','ChangeLog','Attempts','Pending','ReconciliationPlans','ReconciliationReceipts','ProjectionAcks')) {
  $r=$c.Execute("SELECT * FROM [$t]")
  if($null -eq $r){throw "Recordset nulo: $t"}
  $rows=@()
  while(-not $r.EOF){
   $row=[ordered]@{}
   for($i=0;$i -lt $r.Fields.Count;$i++){
    $v=$r.Fields.Item($i).Value
    if($null -eq $v -or $v -is [DBNull]){$value=$null}else{$value=[Convert]::ToString($v,[Globalization.CultureInfo]::InvariantCulture)}
    $row[$r.Fields.Item($i).Name]=[ordered]@{Tipo=if($null -eq $v){'Null'}else{$v.GetType().Name};Valor=$value}
   }
   $rows+=,$row
   $r.MoveNext()
  }
  $r.Close()
  $estado.Tabelas[$t]=[ordered]@{Quantidade=$rows.Count;Registros=$rows}
 }
} finally { if($c.State -ne 0){$c.Close()} }
$estado.SHA256=(Get-FileHash -LiteralPath $Banco -Algorithm SHA256).Hash.ToLowerInvariant()
$estado | ConvertTo-Json -Depth 14 | Set-Content -LiteralPath $Saida -Encoding utf8
$estado.Tabelas.GetEnumerator() | ForEach-Object { '{0}: {1}' -f $_.Key,$_.Value.Quantidade }
