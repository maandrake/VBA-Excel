$ErrorActionPreference='Stop'
$work='C:\Users\MARCOSALEXANDREPERDO\AppData\Local\Temp\vscode-agent-codex-sandbox-JTJOQt\VBA_Fase2_20261031_201621Z\BackupContextoV4'
if(Test-Path -LiteralPath $work){throw 'Destino ja existe.'}
[void][IO.Directory]::CreateDirectory($work)
$out='C:\Users\MARCOSALEXANDREPERDO\OneDrive - MARCOS DESPACHANTE\Marcos Despachante\Software\VBA-Excel\_auditoria\fase2_20261031_201621Z\testes-backup-contexto-v4-com.json'
$r=[ordered]@{Timestamp=(Get-Date).ToString('o');Kind='Excel/MSXML/WScript COM equivalentes; nenhuma macro VBA executada';Tests=@();TestRoot=$work}
function Hash-Shared([string]$p) {
 $stream=[IO.File]::Open($p,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
 $algo=[Security.Cryptography.SHA256]::Create()
 try{return [BitConverter]::ToString($algo.ComputeHash($stream)).Replace('-','')}finally{$algo.Dispose();$stream.Dispose()}
}
function Check([string]$name,[bool]$ok,$details) { $r.Tests+= [ordered]@{Name=$name;Pass=$ok;Details=$details} }
function Strip-Limit([string]$v) {
 $v=$v.Trim()
 if($v.StartsWith('=')) {$v=$v.Substring(1)}
 if($v.Length -ge 2 -and $v.StartsWith('"') -and $v.EndsWith('"')) {$v=$v.Substring(1,$v.Length-2).Replace('""','"')}
 return $v
}
$y=$null; $x=$null; $a=$null; $b=$null; $restore=$null; $child=$null
try {
 $y=New-Object -ComObject Excel.Application
 $y.Visible=$false; $y.DisplayAlerts=$false; $y.EnableEvents=$false
 $x=New-Object -ComObject Excel.Application
 $x.Visible=$false; $x.DisplayAlerts=$false; $x.EnableEvents=$false
 $r.AutomationSecurityObserved=$x.AutomationSecurity
 $a=$x.Workbooks.Add(); $a.Worksheets.Item(1).Range('A1').Value2='FICTICIO_A'
 $pa=Join-Path $work 'contexto_A.xlsx'
 if(Test-Path -LiteralPath $pa){throw 'Arquivo de teste ja existe.'}
 $a.SaveAs($pa,51)
 $b=$x.Workbooks.Add(); $b.Worksheets.Item(1).Range('A1').Value2='FICTICIO_B'
 $pb=Join-Path $work 'contexto_B.xlsx'
 if(Test-Path -LiteralPath $pb){throw 'Arquivo de teste ja existe.'}
 $b.SaveAs($pb,51)
 $a.Worksheets.Item(1).Range('A2').Value2='A_ALTERADO'
 $b.Worksheets.Item(1).Range('A2').Value2='B_NAO_SALVO'
 $b.Activate()
 $a.Save()
 Check 'Salvar proprietario com outra pasta ativa' ($a.Saved -and -not $b.Saved -and $x.ActiveWorkbook.Name -eq $b.Name) @{SavedA=$a.Saved;SavedB=$b.Saved;Active=$x.ActiveWorkbook.Name}
 $diskBefore=(Hash-Shared $pa)
 $key='00000000000000000000000000000000000000000001'
 $a.Worksheets.Item(1).Range('D4').NumberFormat='@'
 $a.Worksheets.Item(1).Range('D4').Value2=[string]$key
 $a.Worksheets.Item(1).Range('C4').Formula='=1+2'
 $a.Worksheets.Item(1).Range('C4').NumberFormat='0.00'
 $formatSnapshot=[string]$a.Worksheets.Item(1).Range('C4').NumberFormat
 $r.NumberFormatObservedBeforeBackup=$formatSnapshot
 $a.Worksheets.Item(1).Range('B4').Value2='VERSAO_1'
 $bkRoot=Join-Path $work 'Backups_Ficticios'
 [void][IO.Directory]::CreateDirectory($bkRoot)
 $paths=@();$hashes=@()
 for($i=1;$i -le 2;$i++) {
  $a.Worksheets.Item(1).Range('B4').Value2=[string]('VERSAO_'+$i)
  $tag='contexto_A_'+(Get-Date -Format yyyyMMdd_HHmmss)
  $fso=New-Object -ComObject Scripting.FileSystemObject
  $created=$false
  for($n=1;$n -le 999;$n++){
   $candidate=Join-Path $bkRoot ($tag+'_'+$n.ToString('000'))
   if($fso.FolderExists($candidate)){continue}
   [void]$fso.CreateFolder($candidate);$created=$true;break
  }
  if(-not $created){throw 'Sem destino unico.'}
  $p=Join-Path $candidate 'contexto_A.xlsx'
  if(Test-Path -LiteralPath $p){throw 'Colisao inesperada.'}
  $a.SaveCopyAs($p)
  $bytes=[IO.File]::ReadAllBytes($p)
  Check ('Backup '+$i+' tem pacote ZIP') ($bytes.Length -gt 2 -and $bytes[0] -eq 80 -and $bytes[1] -eq 75) @{Bytes=$bytes.Length}
  $hashes+=(Get-FileHash -LiteralPath $p -Algorithm SHA256).Hash
  $paths+=$p
  $restore=$y.Workbooks.Open($p,0,$true)
  if($null -eq $restore){throw "Reabertura nao retornou pasta: $p"}
  Check ('Restaurar backup '+$i) ($restore.Worksheets.Item(1).Range('B4').Value2 -ceq ('VERSAO_'+$i) -and $restore.Worksheets.Item(1).Range('D4').Value2 -ceq $key -and $restore.Worksheets.Item(1).Range('C4').Formula -eq '=1+2' -and $restore.Worksheets.Item(1).Range('C4').NumberFormat -eq $formatSnapshot) @{KeyLength=$restore.Worksheets.Item(1).Range('D4').Value2.Length;Formula=$restore.Worksheets.Item(1).Range('C4').Formula;FormatBefore=$formatSnapshot;FormatAfter=$restore.Worksheets.Item(1).Range('C4').NumberFormat}
  $restore.Close($false);$restore=$null
 }
 Check 'Versionamento sem sobrescrita' ($paths[0] -cne $paths[1] -and (Get-FileHash -LiteralPath $paths[0] -Algorithm SHA256).Hash -eq $hashes[0]) @{Paths=$paths;Hashes=$hashes}
 Check 'SaveCopyAs nao salva nem modifica original em disco' (-not $a.Saved -and (Hash-Shared $pa) -eq $diskBefore) @{Saved=$a.Saved;DiskHash=$diskBefore}
 foreach($case in @(@('=5','5'),@('=10','10'),@('="ABC"','ABC'),@('="=ABC"','=ABC'),@('="A""B"','A"B'),@('ABC','ABC'),@('',''))) {
  Check ('Leitura de limite '+$case[0]) ((Strip-Limit $case[0]) -ceq $case[1]) @{Output=(Strip-Limit $case[0])}
 }
 $s=$a.Worksheets.Item(1)
 [void]$s.Range('F4').FormatConditions.Add(1,2,'5','10')
 $fc=$s.Range('F4').FormatConditions.Item(1);$fc.Interior.Color=255
 $low=[double](Strip-Limit ([string]$fc.Formula1));$high=[double](Strip-Limit ([string]$fc.Formula2))
 foreach($v in @(4,5,7,10,11)){
  $s.Range('F4').Value2=[double]$v;$x.CalculateFull()
  $expected=($v -lt $low -or $v -gt $high)
  $actual=($s.Range('F4').DisplayFormat.Interior.Color -eq 255)
  Check ('xlNotBetween com limites lidos: '+$v) ($expected -eq $actual) @{Formula1=$fc.Formula1;Formula2=$fc.Formula2;NativeExcel=$actual;Equivalent=$expected}
 }
 $wsh=New-Object -ComObject WScript.Shell
 $child=$wsh.Exec('powershell.exe -NoProfile -NonInteractive -Command "Start-Sleep -Seconds 3"')
 $timer=[Diagnostics.Stopwatch]::StartNew()
 while($child.Status -eq 0 -and $timer.ElapsedMilliseconds -lt 400){[Threading.Thread]::Sleep(15)}
 $terminated=$false
 if($child.Status -eq 0){$child.Terminate();$terminated=$true}
 Check 'Timeout encerra somente subprocesso criado' ($terminated -and $timer.ElapsedMilliseconds -lt 6000) @{PID=$child.ProcessID;Milliseconds=$timer.ElapsedMilliseconds}
} catch {$r.Error=$_.Exception.Message;$r.ErrorDetail=$_.ToString();$r.Position=$_.InvocationInfo.PositionMessage} finally {
 if($null -ne $restore){try{$restore.Close($false)}catch{}}
 if($null -ne $a){try{$a.Close($false)}catch{}}
 if($null -ne $b){try{$b.Close($false)}catch{}}
 if($null -ne $y){try{$y.Quit()}catch{};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($y)}
 if($null -ne $x){try{$x.Quit()}catch{};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($x)}
}
$r | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $out -Encoding UTF8
$r | ConvertTo-Json -Depth 10
