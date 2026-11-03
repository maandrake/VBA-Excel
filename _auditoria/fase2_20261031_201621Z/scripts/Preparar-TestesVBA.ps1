param(
 [string]$SourceRoot=(Join-Path $PSScriptRoot '..\desenvolvimento'),
 [string]$TestRoot=(Join-Path $env:TEMP ('VBA_Fase2_' + [guid]::NewGuid().ToString('N'))),
 [switch]$RunTests
)
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
if($RunTests){throw 'Execute as suites manualmente no editor VBA somente depois da compilacao e da verificacao da fixture. Este preparador nao executa macros.'}
$src=(Resolve-Path -LiteralPath $SourceRoot).Path
$dst=[IO.Path]::GetFullPath($TestRoot)
$prefix=[IO.Path]::GetFullPath($env:TEMP).TrimEnd('\')+'\'
if(-not $dst.StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase) -or $dst -match '(?i)OneDrive'){throw 'Fixture exige TEMP fora do OneDrive.'}
if(Test-Path -LiteralPath $dst){throw 'Destino existente; nao sera sobrescrito.'}
[void][IO.Directory]::CreateDirectory($dst)
$stage=Join-Path $dst 'importacao_cp1252'
[void][IO.Directory]::CreateDirectory($stage)
$schema=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'fixture-esquema.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$r=[ordered]@{Timestamp=(Get-Date).ToString('o');Status='STARTING';VbaCompiled=$false;VbaExecuted=$false;Imported=@();References=@();Limits=$schema.limits;TestRoot=$dst}
$x=$null;$b=$null
function Stage-Source([IO.FileInfo]$f) {
 $bytes=[IO.File]::ReadAllBytes($f.FullName)
 try{$content=(New-Object Text.UTF8Encoding($false,$true)).GetString($bytes)}catch{$content=[Text.Encoding]::GetEncoding(1252).GetString($bytes)}
 $content=$content.TrimStart([char]0xFEFF)
 $encoder=[Text.Encoding]::GetEncoding(1252,[Text.EncoderExceptionFallback]::new(),[Text.DecoderExceptionFallback]::new())
 $p=Join-Path $stage $f.Name
 if(Test-Path -LiteralPath $p){throw "Nome de arquivo repetido no staging: $p"}
 [IO.File]::WriteAllBytes($p,$encoder.GetBytes($content))
 if($f.Extension -eq '.frm'){
  $frx=[IO.Path]::ChangeExtension($f.FullName,'.frx')
  if(-not (Test-Path -LiteralPath $frx)){throw "FRX ausente: $frx"}
  [IO.File]::Copy($frx,[IO.Path]::ChangeExtension($p,'.frx'),$false)
 }
 return $p
}
try{
 $x=New-Object -ComObject Excel.Application
 $x.Visible=$false;$x.DisplayAlerts=$false;$x.EnableEvents=$false;$x.AskToUpdateLinks=$false
 $r.ExcelVersion=$x.Version;$r.AutomationSecurityObserved=$x.AutomationSecurity
 $b=$x.Workbooks.Add()
 $project=$b.GetType().InvokeMember('VBProject',[Reflection.BindingFlags]::GetProperty,$null,$b,$null)
 if($null -eq $project){throw 'VBProject null. Importacao e compilacao nao realizadas.'}
 # Nenhum mecanismo altera AccessVBOM, ExecutionPolicy ou seguranca de macros.
 while($b.Worksheets.Count -gt 1){$b.Worksheets.Item($b.Worksheets.Count).Delete()}
 $b.Worksheets.Item(1).Name=$schema.worksheets[0].name
 for($i=1;$i -lt $schema.worksheets.Count;$i++){
  $ws=$b.Worksheets.Add([Type]::Missing,$b.Worksheets.Item($b.Worksheets.Count));$ws.Name=$schema.worksheets[$i].name
 }
 for($i=0;$i -lt $schema.worksheets.Count;$i++){
  $ws=$b.Worksheets.Item($schema.worksheets[$i].name)
  $project.VBComponents.Item($ws.CodeName).Name=('AuditoriaTmp'+$i)
 }
 foreach($s in $schema.worksheets){
  $ws=$b.Worksheets.Item($s.name)
  $component=$project.VBComponents.Item($ws.CodeName)
  $component.Name=$s.codeName
  if($component.CodeModule.CountOfLines -gt 0){$component.CodeModule.DeleteLines(1,$component.CodeModule.CountOfLines)}
  if(-not [string]::IsNullOrWhiteSpace($s.documentCode)){$component.CodeModule.AddFromString($s.documentCode)}
 }
 $hostComponent=$project.VBComponents.Item($b.CodeName)
 $hostComponent.Name=$schema.workbookCodeName
 if($hostComponent.CodeModule.CountOfLines -gt 0){$hostComponent.CodeModule.DeleteLines(1,$hostComponent.CodeModule.CountOfLines)}
 if(-not [string]::IsNullOrWhiteSpace($schema.workbookCode)){$hostComponent.CodeModule.AddFromString($schema.workbookCode)}
 foreach($reference in $schema.references){
  try{
   $found=@($project.References | Where-Object {$_.Guid -eq $reference.guid})
   if($found.Count -eq 0){$v=$reference.version.Split('.');[void]$project.References.AddFromGuid($reference.guid,[int]$v[0],[int]$v[1])}
   $r.References+=@{Guid=$reference.guid;Description=$reference.description;Status='AVAILABLE'}
  }catch{$r.References+=@{Guid=$reference.guid;Description=$reference.description;Status='UNAVAILABLE';Error=$_.Exception.Message}}
 }
 $sources=@(Get-ChildItem -LiteralPath (Join-Path $src 'MDL') -File | Where-Object {$_.Extension -in '.bas','.cls'})
 $sources+=@(Get-ChildItem -LiteralPath (Join-Path $src 'Form') -Filter '*.frm' -File)
 $sources+=Get-Item -LiteralPath (Join-Path $PSScriptRoot 'AuditoriaListaFake.cls')
 $sources+=Get-Item -LiteralPath (Join-Path $PSScriptRoot 'modTestesAuditoria.bas')
 $sources+=Get-Item -LiteralPath (Join-Path $PSScriptRoot 'modTestesFase2.bas')
 foreach($source in $sources){
  $staged=Stage-Source $source
  $component=$project.VBComponents.Import($staged)
  if($null -eq $component){throw "Importacao retornou null: $($source.Name)"}
  $r.Imported+=@{File=$source.Name;Component=$component.Name;Lines=$component.CodeModule.CountOfLines}
 }
 $b.Worksheets.Item('Menu Principal').Range('Z1').Value2='AUDITORIA_VBA_FICTICIO'
 $fixture=Join-Path $dst ('Fixture_Auditoria_'+(Get-Date -Format yyyyMMdd_HHmmss)+'.xlsm')
 $b.SaveAs($fixture,52)
 $r.Fixture=$fixture
 $r.Components=$project.VBComponents.Count
 $r.ResolvedReferences=@($project.References | ForEach-Object {@{Name=$_.Name;Guid=$_.Guid;IsBroken=$_.IsBroken}})
 $r.Status='PREPARED_NOT_COMPILED_NOT_FUNCTIONAL'
}catch{
 $r.Status='BLOCKED_OR_FAILED'
 $r.Error=$_.Exception.Message
 $base=$_.Exception
 while($null -ne $base.InnerException){$base=$base.InnerException}
 $r.OriginalOfficeError=$base.Message;$r.OriginalHResult=('0x{0:X8}' -f $base.HResult)
}finally{
 if($null -ne $b){try{$b.Close($false)}catch{}}
 if($null -ne $x){try{$x.Quit()}catch{};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($x)}
 $r | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $dst 'resultado_preparacao.json') -Encoding UTF8
}
$r | ConvertTo-Json -Depth 10
