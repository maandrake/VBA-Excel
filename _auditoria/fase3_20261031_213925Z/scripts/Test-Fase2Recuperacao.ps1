param([string]$TestRoot=(Join-Path $env:TEMP ('VBA_Fase2_Recuperacao_'+[guid]::NewGuid().ToString('N'))))
$ErrorActionPreference='Stop'
$prefix=[IO.Path]::GetFullPath($env:TEMP).TrimEnd('\')+'\'
$dst=[IO.Path]::GetFullPath($TestRoot)
if(-not $dst.StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase) -or $dst -match '(?i)OneDrive'){throw 'Somente TEMP fora do OneDrive.'}
if(Test-Path -LiteralPath $dst){throw 'Destino existente.'}
[void][IO.Directory]::CreateDirectory($dst)
$r=[ordered]@{Timestamp=(Get-Date).ToString('o');Kind='Excel COM equivalente da compensacao por celula; nao executa clsOperacaoPlanilha nem eventos VBA';Tests=@();TestRoot=$dst}
function Snapshot($book,$names){
 $cells=@()
 foreach($name in $names){$s=$book.Worksheets.Item($name)
  foreach($addr in @('A4','B4','C4','D4')){
   $c=$s.Range($addr)
   $cells+= [pscustomobject]@{Sheet=$name;Address=$addr;Value=$c.Value2;Formula=$c.Formula;HasFormula=[bool]$c.HasFormula;Format=[string]$c.NumberFormat}
  }
 }
 return ,$cells
}
function Restore($book,$cells){
 for($i=$cells.Count-1;$i -ge 0;$i--){
  $t=$cells[$i];$c=$book.Worksheets.Item($t.Sheet).Range($t.Address);$c.ClearContents()
  if($t.HasFormula){$c.NumberFormat=[string]$t.Format;$c.Formula=[string]$t.Formula}
  elseif($t.Value -is [string]){$c.NumberFormat='@';$c.Value2=[string]$t.Value;$c.NumberFormat=[string]$t.Format}
  else{$c.NumberFormat=[string]$t.Format;if($null -ne $t.Value){$c.Value2=$t.Value}}
 }
}
function Initialize($book,$names){
 foreach($name in $names){
  $s=$book.Worksheets.Item($name);$s.Range('A4:D10').ClearContents();$s.Range('A4:D4').ClearFormats();$defaultFormat=[string]$s.Range('B4').NumberFormat
  $s.Range('B4').NumberFormat='@';$s.Range('B4').Value2='=TEXTO_FICTICIO';$s.Range('B4').NumberFormat=$defaultFormat
  $s.Range('C4').Formula='=1+2';$s.Range('D4').NumberFormat='@';$s.Range('D4').Value2='00000000000000000000000000000000000000000001'
 }
}
function Add-Result($name,$ok,$details){$r.Tests+= [ordered]@{Name=$name;Pass=[bool]$ok;Details=$details}}
$x=$null;$b=$null
try{
 $x=New-Object -ComObject Excel.Application
 $x.Visible=$false;$x.DisplayAlerts=$false;$x.EnableEvents=$false
 $r.AutomationSecurityObserved=$x.AutomationSecurity
 $b=$x.Workbooks.Add()
 $all=@('Entradas','Entradas Anual','Entradas Filial','Entradas Anual Filial','Saidas','Todas Saidas','Saidas Filial','Todas Saidas Filial')
 $b.Worksheets.Item(1).Name=$all[0]
 foreach($name in $all[1..7]){$s=$b.Worksheets.Add();$s.Name=$name}
 $flows=@(
  @{Name='Entrada';Dest=@('Entradas','Entradas Anual')},
  @{Name='EntradaFilial';Dest=@('Entradas Filial','Entradas Anual Filial')},
  @{Name='Venda';Dest=@('Saidas','Todas Saidas')},
  @{Name='Transferencia';Dest=@('Saidas','Entradas Filial','Entradas Anual Filial','Todas Saidas')},
  @{Name='Devolucao';Dest=@('Saidas Filial','Todas Saidas Filial','Entradas','Entradas Anual')},
  @{Name='Cancelado';Dest=@('Saidas','Todas Saidas')}
 )
 foreach($flow in $flows){
  $names=$flow.Dest
  foreach($partial in @($false,$true)){
   for($failAt=0;$failAt -lt $names.Count;$failAt++){
    Initialize $b $names
    $before=Snapshot $b $names
    $appBefore=@($x.EnableEvents,$x.ScreenUpdating,$x.DisplayAlerts,$x.Calculation)
    $x.EnableEvents=$false;$x.ScreenUpdating=$false
    $injected=$false
    try{
     for($i=0;$i -lt $names.Count;$i++){
      $s=$b.Worksheets.Item($names[$i]);$s.Range('A4').Value2='NF_FICTICIA_1'
      if($i -eq $failAt -and $partial){throw 'FALHA_CONTROLADA'}
      $s.Range('B4').Value2='ALTERADO';$s.Range('C4').Value2=99.0;$s.Range('D4').Value2='ALTERADO'
      if($i -eq $failAt){throw 'FALHA_CONTROLADA'}
     }
    }catch{
     if($_.Exception.Message -ne 'FALHA_CONTROLADA'){throw}
     $injected=$true;Restore $b $before
    }finally{
     $x.Calculation=$appBefore[3];$x.DisplayAlerts=$appBefore[2];$x.ScreenUpdating=$appBefore[1];$x.EnableEvents=$appBefore[0]
    }
    $after=Snapshot $b $names
    $appAfter=@($x.EnableEvents,$x.ScreenUpdating,$x.DisplayAlerts,$x.Calculation)
    $same=(($before|ConvertTo-Json -Depth 5 -Compress) -ceq ($after|ConvertTo-Json -Depth 5 -Compress))
    $sameApp=(($appBefore|ConvertTo-Json -Compress) -ceq ($appAfter|ConvertTo-Json -Compress))
    Add-Result ($flow.Name+' falha '+$failAt+' parcial='+$partial) ($injected -and $same -and $sameApp) @{Before=$before;After=$after;AppBefore=$appBefore;AppAfter=$appAfter}
    foreach($name in $names){$b.Worksheets.Item($name).Range('A4').Value2='NF_FICTICIA_1'}
    $counts=@($names | ForEach-Object {$x.WorksheetFunction.CountIf($b.Worksheets.Item($_).Range('A4:A10'),'NF_FICTICIA_1')})
    Add-Result ($flow.Name+' nova tentativa '+$failAt+' parcial='+$partial) (@($counts|Where-Object {$_ -ne 1}).Count -eq 0) @{Counts=$counts}
   }
  }
  Initialize $b $names
  foreach($name in $names){$b.Worksheets.Item($name).Range('A4').Value2='NF_FICTICIA_1'}
  $counts=@($names | ForEach-Object {$x.WorksheetFunction.CountIf($b.Worksheets.Item($_).Range('A4:A10'),'NF_FICTICIA_1')})
  Add-Result ($flow.Name+' sucesso antes de baixa/Save') (@($counts|Where-Object {$_ -ne 1}).Count -eq 0) @{Counts=$counts}
 }
 $p=Join-Path $dst 'recuperacao_ficticia.xlsx';$b.SaveAs($p,51);$r.Fixture=$p
}catch{$r.Error=$_.Exception.Message;$r.Position=$_.InvocationInfo.PositionMessage}
finally{if($null -ne $b){try{$b.Close($false)}catch{}};if($null -ne $x){try{$x.Quit()}catch{};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($x)}}
$r|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $dst 'resultado_recuperacao.json') -Encoding UTF8
[ordered]@{Tests=$r.Tests.Count;Pass=@($r.Tests|Where-Object {$_.Pass}).Count;Fail=@($r.Tests|Where-Object {-not $_.Pass}).Count;Error=$r['Error'];Evidence=(Join-Path $dst 'resultado_recuperacao.json')}|ConvertTo-Json
