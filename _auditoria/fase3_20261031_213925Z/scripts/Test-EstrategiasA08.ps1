$ErrorActionPreference='Stop'
$work='C:\Users\MARCOSALEXANDREPERDO\AppData\Local\Temp\vscode-agent-codex-sandbox-JTJOQt\VBA_Fase3_20261031_213925Z'
$base='C:\Users\MARCOSALEXANDREPERDO\OneDrive - MARCOS DESPACHANTE\Marcos Despachante\Software\VBA-Excel\_auditoria\fase3_20261031_213925Z'
$dst=Join-Path $work 'A08'
if(Test-Path -LiteralPath $dst){throw 'Destino existente.'}
[void][IO.Directory]::CreateDirectory($dst)
$r=[ordered]@{Timestamp=(Get-Date).ToString('o');Modality='Operacao equivalente COM e prototipo de identidade em memoria; nenhum VBA';Cases=@();ProductionPolicyChanged=$false}
function Key([int]$n){return $n.ToString().PadLeft(44,'0')}
function Seed($book){
 $book.Worksheets.Item(1).Name='Entradas Filial'
 $h=$book.Worksheets.Add();$h.Name='Entradas Anual Filial'
 $s=$book.Worksheets.Item('Entradas Filial')
 for($row=4;$row -le 6;$row++){
  $i=$row-3
  $s.Range('A'+$row).NumberFormat='@';$s.Range('A'+$row).Value2=[string](1000+$i)
  $s.Range('F'+$row).Value2=[string]('VEICULO_FICTICIO_'+$i)
  $s.Range('I'+$row).Value2=[double](10000+$i)
  $s.Range('J'+$row).Value2=([datetime]'2026-01-01').ToOADate()
  $s.Range('L'+$row).NumberFormat='@';$s.Range('L'+$row).Value2=[string](Key $i)
  $s.Range('M'+$row).Formula=[string]('=IF(J'+$row+'="","",MONTH(J'+$row+'))')
  $s.Range('N'+$row).Formula=[string]('=IF(J'+$row+'="","",TODAY()-J'+$row+')')
  $s.Range('O'+$row).Formula=[string]('=IF(A'+$row+'="","",F'+$row+')')
  $s.Range('T'+$row).Value2=[string]('LATERAL_'+$row)
  $h.Range('A'+$row).Value2=[string](1000+$i)
  $h.Range('F'+$row).Value2=[string]('VEICULO_FICTICIO_'+$i)
 }
 $h.Range('K4').Formula="='Entradas Filial'!F4"
 $h.Range('P4').Formula=('=INDEX(''Entradas Filial''!F4:F6,MATCH("'+(Key 1)+'",''Entradas Filial''!L4:L6,0))')
}
function Observe($book){
 $s=$book.Worksheets.Item('Entradas Filial');$h=$book.Worksheets.Item('Entradas Anual Filial')
 return [ordered]@{A4=$s.Range('A4').Value2;F4=$s.Range('F4').Value2;L4=$s.Range('L4').Value2;M4=$s.Range('M4').Formula;N4=$s.Range('N4').Formula;O4=$s.Range('O4').Formula;T4=$s.Range('T4').Value2;HistoryPlate=$h.Range('F4').Value2;DirectFormula=$h.Range('K4').Formula;DirectValue=$h.Range('K4').Value2;KeyFormula=$h.Range('P4').Formula;KeyValue=$h.Range('P4').Value2}
}
function Active-Map($sheet){
 $map=New-Object 'Collections.Generic.Dictionary[string,int]'
 for($row=4;$row -le 6;$row++){
  $id=[string]$sheet.Range('L'+$row).Value2
  if($id -notmatch '^[0-9]{44}$'){throw 'IDENTIDADE_AUSENTE_OU_INVALIDA'}
  if($map.ContainsKey($id)){throw 'IDENTIDADE_AMBIGUA'}
  $map.Add($id,$row)
 }
 return ,$map
}
$x=$null;$b=$null
try{
 $x=New-Object -ComObject Excel.Application
 $x.Visible=$false;$x.DisplayAlerts=$false;$x.EnableEvents=$false
 $r.ExcelVersion=$x.Version;$r.OwnHwnd=$x.Hwnd
 foreach($strategy in @('A_Parcial','A_LinhaInteira','B_Limpeza','C_IdentidadeMemoria')){
  $b=$x.Workbooks.Add();Seed $b
  $s=$b.Worksheets.Item('Entradas Filial');$before=Observe $b
  $case=[ordered]@{Strategy=$strategy;Before=$before}
  if($strategy -eq 'A_Parcial'){
   [void]$s.Range('A4:M4').Delete(-4162);$x.CalculateFull()
   $case.After=Observe $b
   $case.Observation='Delete A:M deslocou campos e quebrou referencias; historico de valores mantido, referencia direta quebrada.'
   $case.OraclePass=($case.After.A4 -eq '1002' -and $case.After.N4 -like '*#REF!*')
   $case.IntegrityApproved=$false
  }elseif($strategy -eq 'A_LinhaInteira'){
   [void]$s.Range('A4').EntireRow.Delete();$x.CalculateFull()
   $case.After=Observe $b
   $case.Observation='Linha inteira removeu sentinela lateral; nao e substituicao universal.'
   $case.OraclePass=($case.After.T4 -ne 'LATERAL_4')
   $case.IntegrityApproved=$false
  }elseif($strategy -eq 'B_Limpeza'){
   [void]$s.Range('A4:L4').ClearContents();$x.CalculateFull()
   $case.AfterClear=Observe $b
   $legacyCount=0;$scanCount=0;$row=4
   while($row -le 6 -and -not [string]::IsNullOrEmpty([string]$s.Range('A'+$row).Value2)){$legacyCount++;$row++}
   for($row=4;$row -le 6;$row++){if(-not [string]::IsNullOrEmpty([string]$s.Range('A'+$row).Value2)){$scanCount++}}
   $case.SearchComparison=@{WhileFirstEmpty=$legacyCount;ScanAllRows=$scanCount;Scope='Modelos equivalentes; pesquisa rapida do projeto ja percorre lacunas. Filtros legados de saidas/NCM usam While.'}
   $s.Range('A4').NumberFormat='@';$s.Range('A4').Value2='9999';$s.Range('F4').Value2='VEICULO_FICTICIO_NOVO';$s.Range('L4').NumberFormat='@';$s.Range('L4').Value2=[string](Key 9)
   $x.CalculateFull();$case.AfterReuse=Observe $b
   $case.OraclePass=($case.AfterClear.M4 -eq $before.M4 -and $scanCount -eq 2 -and $legacyCount -eq 0 -and $case.AfterReuse.DirectValue -eq 'VEICULO_FICTICIO_NOVO' -and $case.AfterReuse.HistoryPlate -eq 'VEICULO_FICTICIO_1')
   $case.Observation='Sem Delete, formulas preservadas; reutilizacao troca silenciosamente veiculo da referencia direta.'
   $case.IntegrityApproved=$false
  }else{
   $active=Active-Map $s
   [void]$active.Remove((Key 1))
   $case.ActiveRecords=$active.Count;$case.ExistingUnadaptedSearchRecords=3
   $x.CalculateFull();$case.After=Observe $b
   $case.DurableLedger=$false
   $case.PrototypeIdentity='Chave ficticia em L, confirmada somente na fixture. Nao e contrato da aplicacao.'
   $s.Range('L6').Value2=[string](Key 2)
   $duplicateRejected=$false
   try{$ignored=Active-Map $s}catch{$duplicateRejected=($_.Exception.Message -eq 'IDENTIDADE_AMBIGUA')}
   [void]$s.Range('L6').ClearContents()
   $missingRejected=$false
   try{$ignored=Active-Map $s}catch{$missingRejected=($_.Exception.Message -eq 'IDENTIDADE_AUSENTE_OU_INVALIDA')}
   $s.Range('L6').NumberFormat='@';$s.Range('L6').Value2=[string](Key 3)
   $case.DuplicateRejected=$duplicateRejected;$case.MissingRejected=$missingRejected
   $case.OraclePass=($active.Count -eq 2 -and $case.After.DirectValue -eq 'VEICULO_FICTICIO_1' -and $case.After.KeyValue -eq 'VEICULO_FICTICIO_1' -and $duplicateRejected -and $missingRejected)
   $case.Observation='Identidade e historico preservados no prototipo, mas exige identidade completa, leitores adaptados e persistencia aprovada.'
   $case.IntegrityApproved='SOMENTE_PROTOTIPO_SEM_INTEGRACAO'
  }
  $file=Join-Path $dst ($strategy+'.xlsx')
  $b.SaveAs($file,51);$b.Close($false);$b=$null
  $case.Fixture=$file;$case.SHA256=(Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash
  $r.Cases+=$case
 }
}catch{$r.Error=$_.Exception.Message;$r.Position=$_.InvocationInfo.PositionMessage}
finally{if($null -ne $b){try{$b.Close($false)}catch{}};if($null -ne $x){try{$x.Quit()}catch{};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($x)}}
$r|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $base 'a08-estrategias-com.json') -Encoding UTF8
$r|ConvertTo-Json -Depth 12
