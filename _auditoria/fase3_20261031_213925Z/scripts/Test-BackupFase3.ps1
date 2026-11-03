$ErrorActionPreference='Stop'
$work='C:\Users\MARCOSALEXANDREPERDO\AppData\Local\Temp\vscode-agent-codex-sandbox-JTJOQt\VBA_Fase3_20261031_213925Z'
$base='C:\Users\MARCOSALEXANDREPERDO\OneDrive - MARCOS DESPACHANTE\Marcos Despachante\Software\VBA-Excel\_auditoria\fase3_20261031_213925Z'
$dst=Join-Path $work 'Backup'
if(Test-Path -LiteralPath $dst){throw 'Destino existente.'}
[void][IO.Directory]::CreateDirectory($dst)
$r=[ordered]@{Timestamp=(Get-Date).ToString('o');Modality='Excel/FSO COM equivalente; nao executa CriarBackupVersionado VBA';Cases=@();Copies=@();SecurityChanged=$false;RetentionCleanup=$false}
function Hash-Shared([string]$p){
 $s=[IO.File]::Open($p,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
 $h=[Security.Cryptography.SHA256]::Create()
 try{return [BitConverter]::ToString($h.ComputeHash($s)).Replace('-','')}finally{$h.Dispose();$s.Dispose()}
}
function Version($book,[string]$root,[bool]$injectLock){
 $fso=New-Object -ComObject Scripting.FileSystemObject
 if(-not $fso.FolderExists($root)){[void]$fso.CreateFolder($root)}
 $prefix=$fso.GetBaseName($book.Name)+'_TIMESTAMP_FICTICIO'
 $collisions=0
 for($n=1;$n -le 999;$n++){
  $folder=$fso.BuildPath($root,$prefix+'_'+$n.ToString('000'))
  try{[void]$fso.CreateFolder($folder);break}
  catch{$e=$_.Exception.GetBaseException();if(($e.HResult -band 65535) -eq 58){$collisions++;continue}else{throw}}
 }
 if($n -gt 999){throw 'Sem versao disponivel.'}
 $p=$fso.BuildPath($folder,$book.Name)
 if($fso.FileExists($p)){throw 'Destino existente.'}
 $lock=$null
 try{
  if($injectLock){$lock=[IO.File]::Open($p,[IO.FileMode]::CreateNew,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None);$marker=[Text.Encoding]::ASCII.GetBytes('TRAVA_FICTICIA');$lock.Write($marker,0,$marker.Length);$lock.Flush()}
  $book.SaveCopyAs($p)
 }finally{if($null -ne $lock){$lock.Dispose()}}
 $bytes=[IO.File]::ReadAllBytes($p)
 if($bytes.Length -lt 2 -or $bytes[0] -ne 80 -or $bytes[1] -ne 75){throw 'Cabecalho ZIP ausente.'}
 return [pscustomobject]@{Path=$p;Collisions=$collisions;SHA256=(Get-FileHash -LiteralPath $p -Algorithm SHA256).Hash;Size=$bytes.Length}
}
$x=$null;$y=$null;$a=$null;$b=$null;$restored=$null
try{
 $x=New-Object -ComObject Excel.Application;$x.Visible=$false;$x.DisplayAlerts=$false;$x.EnableEvents=$false
 $y=New-Object -ComObject Excel.Application;$y.Visible=$false;$y.DisplayAlerts=$false;$y.EnableEvents=$false
 $r.OwnHwnds=@($x.Hwnd,$y.Hwnd)
 $a=$x.Workbooks.Add();$s=$a.Worksheets.Item(1)
 $s.Range('A1').Value2='BASE_FICTICIA';$s.Range('C4').Formula='=1+2';$s.Range('C4').NumberFormat='0'
 $s.Range('D4').NumberFormat='@';$s.Range('D4').Value2='00000000000000000000000000000000000000000001'
 $pa=Join-Path $dst 'Origem.xlsx';$a.SaveAs($pa,51);$diskInitial=Hash-Shared $pa
 $b=$x.Workbooks.Add();$b.Worksheets.Item(1).Range('A1').Value2='OUTRA_FICTICIA'
 $root=Join-Path $dst 'RaizNova'
 $rootMissing=-not (Test-Path -LiteralPath $root);$formatBefore=[string]$s.Range('C4').NumberFormat
 for($i=1;$i -le 2;$i++){
  $s.Range('B4').Value2=[string]('VERSAO_'+$i);$b.Activate()
  $copy=Version $a $root $false
  $restored=$y.Workbooks.Open($copy.Path,0,$true)
  if($null -eq $restored){throw 'Restauracao retornou null.'}
  $rs=$restored.Worksheets.Item(1)
  $copy|Add-Member -NotePropertyName RestoreOK -NotePropertyValue ($rs.Range('B4').Value2 -eq ('VERSAO_'+$i) -and $rs.Range('D4').Value2 -ceq '00000000000000000000000000000000000000000001' -and $rs.Range('C4').Formula -eq '=1+2' -and $rs.Range('C4').NumberFormat -eq $formatBefore)
  $restored.Close($false);$restored=$null;$r.Copies+=$copy
 }
 $r.Cases+=@{Name='ProprietarioDuasPastasVersoesRestauracao';Status=if($r.Copies[0].RestoreOK -and $r.Copies[1].RestoreOK -and $r.Copies[1].Collisions -eq 1 -and (Get-FileHash -LiteralPath $r.Copies[0].Path -Algorithm SHA256).Hash -eq $r.Copies[0].SHA256 -and $x.ActiveWorkbook.Name -eq $b.Name){'PASS'}else{'FAIL'};RootCreated=$rootMissing -and (Test-Path -LiteralPath $root)}
 foreach($case in @(@{Name='PaiInexistente';Root=(Join-Path $dst 'PaiAusente\Filho')},@{Name='RaizEArquivo';Root=(Join-Path $dst 'RaizArquivo')})){
  if($case.Name -eq 'RaizEArquivo'){[IO.File]::WriteAllText($case.Root,'FICTICIO')}
  $caught=$false;$message=$null
  try{$ignored=Version $a $case.Root $false}catch{$caught=$true;$message=$_.Exception.Message}
  $r.Cases+=@{Name=$case.Name;Status=if($caught){'PASS'}else{'FAIL'};Meaning='Falha propagada no equivalente, sem copiar nem salvar a origem';OriginalError=$message}
 }
 $caught=$false;$message=$null
 try{$ignored=Version $a (Join-Path $dst 'ComTrava') $true}catch{$caught=$true;$message=$_.Exception.Message}
 $r.Cases+=@{Name='FalhaGravacaoSaveCopyAsComArquivoTravado';Status=if($caught){'PASS'}else{'FAIL'};OriginalError=$message;FailureEvidencePreserved=$true}
 $r.Cases+=@{Name='SemSaveInesperadoNaOrigem';Status=if(-not $a.Saved -and (Hash-Shared $pa) -eq $diskInitial -and $b.Worksheets.Item(1).Range('A1').Value2 -eq 'OUTRA_FICTICIA'){'PASS'}else{'FAIL'};OwnerSaved=$a.Saved;DiskSHA256=$diskInitial}
}catch{$r.InfrastructureError=$_.Exception.Message;$r.Position=$_.InvocationInfo.PositionMessage}
finally{
 if($null -ne $restored){try{$restored.Close($false)}catch{}}
 if($null -ne $a){try{$a.Close($false)}catch{}}
 if($null -ne $b){try{$b.Close($false)}catch{}}
 foreach($app in @($y,$x)){if($null -ne $app){try{$app.Quit()}catch{};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}
}
$r|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $base 'backup-fase3-com.json') -Encoding UTF8
$r|ConvertTo-Json -Depth 12
