$ErrorActionPreference='Stop'
$base='C:\Users\MARCOSALEXANDREPERDO\OneDrive - MARCOS DESPACHANTE\Marcos Despachante\Software\VBA-Excel\_auditoria\fase3_20261031_213925Z'
$x=$null;$b=$null;$r=@()
try{
 $x=New-Object -ComObject Excel.Application;$x.Visible=$false;$x.DisplayAlerts=$false;$x.EnableEvents=$false
 $b=$x.Workbooks.Add();$s=$b.Worksheets.Item(1)
 foreach($kind in @('Vazio','Formula','Texto','Chave')){
  $c=$s.Range('A1');[void]$c.ClearContents()
  if($kind -eq 'Formula'){$c.Formula='=1+2'}
  elseif($kind -eq 'Texto'){$c.NumberFormat='@';$c.Value2='=TEXTO_FICTICIO'}
  elseif($kind -eq 'Chave'){$c.NumberFormat='@';$c.Value2='00000000000000000000000000000000000000000001';$c.Errors.Item(3).Ignore=$true}
  $old=[pscustomobject]@{Value=$c.Value2;Formula=$c.Formula;HasFormula=$c.HasFormula;Format=[string]$c.NumberFormat;Ignore=[bool]$c.Errors.Item(3).Ignore}
  $c.Value2='ALTERADO'
  [void]$c.ClearContents()
  if($old.HasFormula){$c.NumberFormat=$old.Format;$c.Formula=[string]$old.Formula}
  elseif($old.Value -is [string]){$c.NumberFormat='@';$c.Value2=[string]$old.Value;$c.NumberFormat=$old.Format}
  else{$c.NumberFormat=$old.Format;if($null -ne $old.Value){$c.Value2=$old.Value}}
  $c.Errors.Item(3).Ignore=[bool]$old.Ignore
  $after=[pscustomobject]@{Value=$c.Value2;Formula=$c.Formula;HasFormula=$c.HasFormula;Format=[string]$c.NumberFormat;Ignore=[bool]$c.Errors.Item(3).Ignore}
  $same=(($old|ConvertTo-Json -Compress) -ceq ($after|ConvertTo-Json -Compress))
  $r+=@{Case=$kind;Before=$old;After=$after;Status=if($same){'PASS'}else{'FAIL'}}
 }
}finally{if($null -ne $b){$b.Close($false)};if($null -ne $x){$x.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($x)}}
@{Timestamp=(Get-Date).ToString('o');Modality='COM equivalente do snapshot corrigido com metadado; nao executa classe VBA';Cases=$r}|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $base 'a09-metadado-corrigido-com.json') -Encoding UTF8
$r|ForEach-Object {[pscustomobject]@{Case=$_['Case'];Status=$_['Status']}}|ConvertTo-Json
