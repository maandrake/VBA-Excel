$ErrorActionPreference='Stop'
$work='C:\Users\MARCOSALEXANDREPERDO\AppData\Local\Temp\vscode-agent-codex-sandbox-JTJOQt\VBA_Fase3_20261031_213925Z'
$base='C:\Users\MARCOSALEXANDREPERDO\OneDrive - MARCOS DESPACHANTE\Marcos Despachante\Software\VBA-Excel\_auditoria\fase3_20261031_213925Z'
$dst=Join-Path $work 'A09'
if(Test-Path -LiteralPath $dst){throw 'Destino existente.'}
[void][IO.Directory]::CreateDirectory($dst)
$r=[ordered]@{Timestamp=(Get-Date).ToString('o');Modality='COM equivalente com estado de fronteira modelado em PowerShell; nao executa VBA';Cases=@();SaveFailureMethod='SaveAs para pasta inexistente; nao e teste do ThisWorkbook.Save da macro'}
$names=@('Saidas','Entradas Filial','Entradas Anual Filial','Todas Saidas','Entradas')
$targets=$names[0..3]
function Snapshot($book,$sheetNames){
 $cells=@()
 foreach($name in $sheetNames){
  foreach($c in $book.Worksheets.Item($name).Range('A4:T5').Cells){
   $cells+=[pscustomobject]@{Sheet=$name;Address=$c.Address($false,$false);Value=$c.Value2;Formula=$c.Formula;HasFormula=[bool]$c.HasFormula;Format=[string]$c.NumberFormat}
  }
 }
 return ,$cells
}
function Restore($book,$cells){
 $errors=@()
 for($i=$cells.Count-1;$i -ge 0;$i--){
  $t=$cells[$i]
  try{
   $c=$book.Worksheets.Item($t.Sheet).Range($t.Address);[void]$c.ClearContents()
   if($t.HasFormula){$c.NumberFormat=[string]$t.Format;$c.Formula=[string]$t.Formula}
   elseif($t.Value -is [string]){$c.NumberFormat='@';$c.Value2=[string]$t.Value;$c.NumberFormat=[string]$t.Format}
   elseif($null -ne $t.Value){$c.NumberFormat=[string]$t.Format;$c.Value2=$t.Value}
   else{$c.NumberFormat=[string]$t.Format}
  }catch{$errors+=@{Sheet=$t.Sheet;Address=$t.Address;Error=$_.Exception.Message}}
 }
 return ,$errors
}
function Hash-Shared([string]$p){
 $s=[IO.File]::Open($p,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
 $h=[Security.Cryptography.SHA256]::Create()
 try{return [BitConverter]::ToString($h.ComputeHash($s)).Replace('-','')}finally{$h.Dispose();$s.Dispose()}
}
function Seed($book){
 $book.Worksheets.Item(1).Name=$names[0]
 foreach($n in $names[1..4]){$s=$book.Worksheets.Add();$s.Name=$n}
 foreach($n in $names){
  $s=$book.Worksheets.Item($n)
  $s.Range('A5').Value2='REGISTRO_FICTICIO_EXISTENTE'
  $s.Range('F5').Value2='VEICULO_ANTERIOR'
  $s.Range('K4').Formula='=1+2';$s.Range('K4').NumberFormat='0'
  $s.Range('M4').Formula='=1+12';$s.Range('N4').Formula='=2+12';$s.Range('O4').Formula='=3+12'
  $s.Range('Q4').Formula='=LEN(A4)';$s.Range('T4').Value2='LATERAL'
  $s.Range('B4').NumberFormat='@';$s.Range('B4').Value2='=TEXTO_FICTICIO'
  $s.Range('L4').NumberFormat='@';$s.Range('L4').Value2='00000000000000000000000000000000000000000001'
 }
 $book.Worksheets.Item('Entradas').Range('A4').Value2='ORIGEM_FICTICIA'
 $book.Worksheets.Item('Entradas').Range('F4').Value2='VEICULO_ORIGEM'
 $book.Worksheets.Item('Entradas Anual Filial').Range('S4').Formula="=Entradas!F4"
}
function Write-Target($book,[int]$i,[bool]$partial){
 $s=$book.Worksheets.Item($targets[$i]);$s.Range('A4').Value2='NF_FICTICIA_1'
 if($partial){throw 'FALHA_CONTROLADA_DENTRO_DA_ESCRITA'}
 $s.Range('F4').Value2='VEICULO_FICTICIO_1'
 $s.Range('I4').Value2=10000.0
 if($targets[$i] -in @('Saidas','Todas Saidas')){
  $s.Range('J4').Value2=12000.0;$s.Range('L4').Value2=([datetime]'2026-01-01').ToOADate()
 }elseif($targets[$i] -eq 'Entradas Filial'){
  $s.Range('J4').Value2=([datetime]'2026-01-01').ToOADate()
  $s.Range('L4').NumberFormat='@';$s.Range('L4').Value2='00000000000000000000000000000000000000000002'
 }else{$s.Range('J4').Value2=([datetime]'2026-01-01').ToOADate()}
}
$cases=@(@{Name='AntesPrimeira';Mode='Before'})
for($i=0;$i -lt 4;$i++){$cases+=@{Name=('Durante'+($i+1));Mode='During';Index=$i}}
for($i=0;$i -lt 3;$i++){$cases+=@{Name=('Entre'+($i+1)+'e'+($i+2));Mode='Between';Index=$i}}
$cases+=@(@{Name='AntesDelete';Mode='BeforeDelete'},@{Name='DeleteRecusado';Mode='DeleteFail'},@{Name='AntesSaveAposDelete';Mode='BeforeSaveAfterDelete'},@{Name='AntesSaveSemDelete';Mode='BeforeSaveNoDelete'},@{Name='FalhaSaveAsSemDelete';Mode='SaveFail'},@{Name='ErroDepoisSave';Mode='AfterSave'},@{Name='FalhaNaCompensacao';Mode='RestoreFail'})
$x=$null;$b=$null;$other=$null
try{
 $x=New-Object -ComObject Excel.Application
 $x.Visible=$false;$x.DisplayAlerts=$false;$x.EnableEvents=$false;$r.OwnHwnd=$x.Hwnd
 $other=$x.Workbooks.Add();$other.Worksheets.Item(1).Range('A4').Value2='OUTRA_PASTA_FICTICIA'
 foreach($case in $cases){
  $x.DisplayAlerts=$false;$x.EnableEvents=$false
  $b=$x.Workbooks.Add();Seed $b
  $p=Join-Path $dst ($case.Name+'.xlsx');$b.SaveAs($p,51)
  $diskBefore=Hash-Shared $p;$other.Activate()
  $x.EnableEvents=$true;$x.ScreenUpdating=$true;$x.Calculation=-4135
  $appBefore=@($x.EnableEvents,$x.ScreenUpdating,$x.DisplayAlerts,$x.Calculation)
  $x.CalculateFull();$before=Snapshot $b $names
  $targetSnapshot=@($before|Where-Object {$_.Sheet -ne 'Entradas'})
  $x.EnableEvents=$false;$x.ScreenUpdating=$false
  $boundaryDelete=$false;$boundarySave=$false;$restoreErrors=@();$errorText=$null;$injected=$false
  try{
   if($case.Mode -eq 'Before'){throw 'FALHA_CONTROLADA_ANTES_DA_PRIMEIRA'}
   for($i=0;$i -lt 4;$i++){
    Write-Target $b $i ($case.Mode -eq 'During' -and $case.Index -eq $i)
    if($case.Mode -eq 'Between' -and $case.Index -eq $i){throw 'FALHA_CONTROLADA_ENTRE_ESCRITAS'}
   }
   if($case.Mode -in @('BeforeDelete','BeforeSaveNoDelete','RestoreFail')){throw 'FALHA_CONTROLADA_ANTES_DA_FRONTEIRA'}
   if($case.Mode -eq 'DeleteFail'){
    $b.Worksheets.Item('Entradas').Protect('FICTICIO')
    $boundaryDelete=$true
    [void]$b.Worksheets.Item('Entradas').Range('A4:N4').Delete(-4162)
    throw 'ERRO: DELETE NAO FOI RECUSADO'
   }
   if($case.Mode -eq 'BeforeSaveAfterDelete'){
    $boundaryDelete=$true;[void]$b.Worksheets.Item('Entradas').Range('A4:N4').Delete(-4162)
    throw 'FALHA_CONTROLADA_DEPOIS_DELETE_ANTES_SAVE'
   }
   $boundarySave=$true
   if($case.Mode -eq 'SaveFail'){
    $b.SaveAs((Join-Path $dst 'PastaInexistente\nao_gravar.xlsx'),51)
    throw 'ERRO: SAVEAS NAO FALHOU'
   }
   $b.Save()
   throw 'FALHA_CONTROLADA_APOS_SAVE'
  }catch{
   $errorText=$_.Exception.Message;$injected=$true
   if(-not $boundaryDelete -and -not $boundarySave){
    if($case.Mode -eq 'RestoreFail'){$b.Worksheets.Item('Saidas').Protect('FICTICIO')}
    $restoreErrors=Restore $b $targetSnapshot
   }
  }finally{
   $x.Calculation=$appBefore[3];$x.DisplayAlerts=$appBefore[2];$x.ScreenUpdating=$appBefore[1];$x.EnableEvents=$appBefore[0]
  }
  $x.CalculateFull();$after=Snapshot $b $names
  $appAfter=@($x.EnableEvents,$x.ScreenUpdating,$x.DisplayAlerts,$x.Calculation)
  $same=(($before|ConvertTo-Json -Depth 5 -Compress) -ceq ($after|ConvertTo-Json -Depth 5 -Compress))
  $appSame=(($appBefore|ConvertTo-Json -Compress) -ceq ($appAfter|ConvertTo-Json -Compress))
  $record=[ordered]@{Name=$case.Name;Error=$errorText;DeleteStarted=$boundaryDelete;SaveStarted=$boundarySave;CellsEqual=$same;AppStateEqual=$appSame;OtherWorkbookUntouched=($other.Worksheets.Item(1).Range('A4').Value2 -eq 'OUTRA_PASTA_FICTICIA');Before=$before;After=$after;AppBefore=$appBefore;AppAfter=$appAfter;RecoveryErrors=$restoreErrors;DiskUnchanged=((Hash-Shared $p) -eq $diskBefore)}
  if($boundaryDelete -or $boundarySave){$record.Recovery='NAO_TENTADA_APOS_FRONTEIRA';$record.Integrity='CONCILIACAO_NECESSARIA';$record.OraclePass=(-not $same -and $appSame -and $errorText -notlike 'ERRO:*')}
  elseif($case.Mode -eq 'RestoreFail'){$record.Recovery='INCOMPLETA';$record.Integrity='FALHA_DE_RECUPERACAO_REPRODUZIDA';$record.OraclePass=($restoreErrors.Count -gt 0 -and -not $same -and $appSame)}
  else{$record.Recovery='COMPLETA_NO_EQUIVALENTE_COM';$record.Integrity='RESTABELECIDA_NO_ESCOPO';$record.OraclePass=($injected -and $same -and $appSame)}
  if($case.Name -eq 'Durante2'){
   for($i=0;$i -lt 4;$i++){Write-Target $b $i $false}
   $counts=@($targets|ForEach-Object {$x.WorksheetFunction.CountIf($b.Worksheets.Item($_).Range('A4:A5'),'NF_FICTICIA_1')})
   $record.RetryCounts=$counts;$record.RetryPass=(@($counts|Where-Object {$_ -ne 1}).Count -eq 0)
  }
  $r.Cases+=$record
  $b.Close($false);$b=$null
 }
}catch{$r.InfrastructureError=$_.Exception.Message;$r.Position=$_.InvocationInfo.PositionMessage}
finally{if($null -ne $b){try{$b.Close($false)}catch{}};if($null -ne $other){try{$other.Close($false)}catch{}};if($null -ne $x){try{$x.Quit()}catch{};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($x)}}
$r|ConvertTo-Json -Depth 14|Set-Content -LiteralPath (Join-Path $base 'a09-fronteiras-com.json') -Encoding UTF8
$r.Cases|Select-Object Name,OraclePass,CellsEqual,AppStateEqual,Recovery,Integrity,RetryPass,DiskUnchanged|ConvertTo-Json -Depth 5
if($r.Contains('InfrastructureError')){$r.InfrastructureError}
