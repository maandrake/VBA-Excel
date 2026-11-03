$ErrorActionPreference='Stop'
$work='C:\Users\MARCOSALEXANDREPERDO\AppData\Local\Temp\vscode-agent-codex-sandbox-JTJOQt\VBA_Fase3_20261031_213925Z'
$base='C:\Users\MARCOSALEXANDREPERDO\OneDrive - MARCOS DESPACHANTE\Marcos Despachante\Software\VBA-Excel\_auditoria\fase3_20261031_213925Z'
$schema=Get-Content -LiteralPath (Join-Path $base 'scripts\fixture-esquema.json') -Raw -Encoding UTF8|ConvertFrom-Json
$r=[ordered]@{Timestamp=(Get-Date).ToString('o');Kind='Sondagem Office COM e registro somente leitura';SecurityChanged=$false;Imported=0;VbaCompiled=$false;VbaExecuted=$false;Dependencies=@();DeclaredLibraries=@();Status='STARTING'}
$x=$null;$b=$null
try{
 $reg=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot,[Microsoft.Win32.RegistryView]::Registry64)
 foreach($prog in @('Forms.UserForm.1','MSComctlLib.TreeCtrl.2','Shell.Explorer.2','Selenium.ChromeDriver','MSXML2.DOMDocument.6.0','VBScript.RegExp')){
  $key=$reg.OpenSubKey($prog+'\CLSID',$false)
  $clsid=$null;if($null -ne $key){$clsid=$key.GetValue('');$key.Dispose()}
  $r.Dependencies+=@{ProgID=$prog;Registered64=($null -ne $clsid);CLSID=$clsid;Instantiated=$false}
 }
 foreach($ref in $schema.references){
  $key=$reg.OpenSubKey('TypeLib\'+$ref.guid+'\'+$ref.version,$false)
  $r.DeclaredLibraries+=@{Name=$ref.description;Guid=$ref.guid;Version=$ref.version;TypeLibVersionRegistered64=($null -ne $key);DeclaredFileExists=(Test-Path -LiteralPath $ref.declaredPath -PathType Leaf)}
  if($null -ne $key){$key.Dispose()}
 }
 $reg.Dispose()
 $x=New-Object -ComObject Excel.Application
 $r.ExcelVersion=$x.Version;$r.ExcelBuild=$x.Build;$r.ExcelPath=$x.Path;$r.OwnHwnd=$x.Hwnd
 $r.AutomationSecurityObserved=$x.AutomationSecurity
 $exe=Join-Path $x.Path 'EXCEL.EXE'
 $stream=[IO.File]::Open($exe,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
 $reader=New-Object IO.BinaryReader($stream)
 try{
  [void]$stream.Seek(60,[IO.SeekOrigin]::Begin);$pe=$reader.ReadInt32()
  [void]$stream.Seek($pe+4,[IO.SeekOrigin]::Begin);$machine=$reader.ReadUInt16()
  $r.ExcelMachine=('0x{0:X4}' -f $machine);$r.ExcelArchitecture=if($machine -eq 34404){'x64'}elseif($machine -eq 332){'x86'}else{'UNKNOWN'}
 }finally{$reader.Dispose();$stream.Dispose()}
 $x.Visible=$false;$x.DisplayAlerts=$false;$x.EnableEvents=$false;$x.AskToUpdateLinks=$false
 $b=$x.Workbooks.Add()
 try{
  $project=$b.GetType().InvokeMember('VBProject',[Reflection.BindingFlags]::GetProperty,$null,$b,$null)
  if($null -eq $project){throw 'VBProject retornou null; nenhuma importacao realizada.'}
  $r.Status='VBPROJECT_AVAILABLE';$r.CurrentComponents=$project.VBComponents.Count
  $r.ResolvedReferences=@($project.References|ForEach-Object {@{Name=$_.Name;IsBroken=$_.IsBroken;Guid=$_.Guid}})
 }catch{
  $e=$_.Exception;$r.ErrorOriginal=$e.ToString()
  while($null -ne $e.InnerException){$e=$e.InnerException}
  $r.RootMessage=$e.Message;$r.HResult=('0x{0:X8}' -f $e.HResult);$r.Status='BLOQUEADO'
 }
 # Esqueleto independente: zero fontes importados, sem dados reais.
 while($b.Worksheets.Count -gt 1){[void]$b.Worksheets.Item($b.Worksheets.Count).Delete()}
 $b.Worksheets.Item(1).Name=$schema.worksheets[0].name
 for($i=1;$i -lt $schema.worksheets.Count;$i++){
  $s=$b.Worksheets.Add([Type]::Missing,$b.Worksheets.Item($b.Worksheets.Count),1,-4167)
  if($null -eq $s){throw 'Criacao de aba retornou null.'}
  $s.Name=[string]$schema.worksheets[$i].name
 }
 $r.SkeletonSheets=@($schema.worksheets|ForEach-Object {$s=$b.Worksheets.Item($_.name);@{Name=$s.Name;ActualCodeName=$s.CodeName;RequiredCodeName=$_.codeName}})
 $b.Worksheets.Item('Menu Principal').Range('Z1').Value2='AUDITORIA_VBA_FICTICIO'
 $fixture=Join-Path $work 'Fixture_Auditoria_Fase3.xlsm'
 if(Test-Path -LiteralPath $fixture){throw 'Fixture ja existente; nao sobrescrever.'}
 $b.SaveAs($fixture,52)
 $r.Skeleton=$fixture;$r.SkeletonStatus='ESQUELETO_SEM_FONTES_NAO_COMPILADO'
 $b.Close($false);$b=$null
 $r.SkeletonSHA256=(Get-FileHash -LiteralPath $fixture -Algorithm SHA256).Hash
}catch{$r.InfrastructureError=$_.Exception.Message;$r.Position=$_.InvocationInfo.PositionMessage}
finally{
 if($null -ne $b){try{$b.Close($false)}catch{}}
 if($null -ne $x){try{$x.Quit()}catch{};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($x)}
 $r|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $base 'office-fase3.json') -Encoding UTF8
}
$r|ConvertTo-Json -Depth 12
