param([string]$TestRoot=(Join-Path $env:TEMP ('VBA_Fase2_A08_'+[guid]::NewGuid().ToString('N'))))
# PREPARADO PARA REPRODUCAO FUTURA. Os sete casos desta fase foram executados por comando inline, conforme JSON.
$ErrorActionPreference='Stop'
$dst=[IO.Path]::GetFullPath($TestRoot);$prefix=[IO.Path]::GetFullPath($env:TEMP).TrimEnd('\')+'\'
if(-not $dst.StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase) -or $dst -match '(?i)OneDrive'){throw 'Somente TEMP fora do OneDrive.'}
if(Test-Path -LiteralPath $dst){throw 'Destino existente.'}
[void][IO.Directory]::CreateDirectory($dst)
$cases=@(
 @{Id='D01';Sheet='Entradas';Address='A3:N3';Entire=$false},
 @{Id='D02';Sheet='Entradas Filial';Address='A3:M3';Entire=$false},
 @{Id='D03';Sheet='Entradas Anual Filial';Address='A3:P3';Entire=$false},
 @{Id='D04';Sheet='Saidas Filial';Address='A3';Entire=$true},
 @{Id='D05';Sheet='Entradas Filial';Address='A3:M3';Entire=$false},
 @{Id='D06';Sheet='Saidas';Address='A3';Entire=$true},
 @{Id='D07';Sheet='Entradas';Address='A3:N3';Entire=$false})
$r=@();$x=$null;$b=$null
try{
 $x=New-Object -ComObject Excel.Application
 $x.Visible=$false;$x.DisplayAlerts=$false;$x.EnableEvents=$false
 foreach($case in $cases){
  $b=$x.Workbooks.Add();$s=$b.Worksheets.Item(1);$s.Name=$case.Sheet
  for($i=2;$i -le 5;$i++){
   $s.Range('A'+$i).Value2=[string]('FICTICIO_'+$i)
   $s.Range('L'+$i).NumberFormat='@';$s.Range('L'+$i).Value2='00000000000000000000000000000000000000000001'
   $s.Range('N'+$i).Formula=('=A'+$i);$s.Range('O'+$i).Value2=[string]('AUX_'+$i)
   $s.Range('P'+$i).Formula=('=LEN(A'+$i+')+200')
   $s.Range('Q'+$i).Formula=('=A'+$i)
  }
  $s.Range('T3').Value2='LATERAL_INDEPENDENTE'
  if($case.Entire){[void]$s.Range($case.Address).EntireRow.Delete()}else{[void]$s.Range($case.Address).Delete(-4162)}
  $r+=@{Id=$case.Id;Kind='Reproducao COM de risco; nao aprovacao da rotina VBA';A3=$s.Range('A3').Value2;N3=$s.Range('N3').Formula;O3=$s.Range('O3').Value2;P3=$s.Range('P3').Formula;Q3=$s.Range('Q3').Formula;T3=$s.Range('T3').Value2}
  $b.Close($false);$b=$null
 }
}finally{
 if($null -ne $b){$b.Close($false)}
 if($null -ne $x){$x.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($x)}
}
$r|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $dst 'a08-reproducao.json') -Encoding UTF8
$r|ConvertTo-Json -Depth 6
