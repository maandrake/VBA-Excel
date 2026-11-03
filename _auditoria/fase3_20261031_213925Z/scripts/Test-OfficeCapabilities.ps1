param([switch]$TestWord)
$ErrorActionPreference = 'Stop'
$testRoot = Join-Path $env:TEMP ('VBA_OfficeProbe_' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $testRoot)
$result = [ordered]@{Timestamp=(Get-Date).ToString('o');TestRoot=$testRoot;Excel=@{};Word=@{Status='NOT_REQUESTED'}}
$xl=$null;$wb=$null;$word=$null;$doc=$null
try {
 $xl = New-Object -ComObject Excel.Application
 $xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.EnableEvents=$false; $xl.AutomationSecurity=3; $xl.AskToUpdateLinks=$false
 $result.Excel.Version=$xl.Version; $result.Excel.Build=$xl.Build
 $wb=$xl.Workbooks.Add()
 try {
  $auditProject=$wb.GetType().InvokeMember('VBProject',[Reflection.BindingFlags]::GetProperty,$null,$wb,$null)
  if($null -eq $auditProject){throw 'VBProject indisponivel'}
  $result.Excel.VBOM=$auditProject.VBComponents.Count
  $result.Excel.VBOMAccessible=$true
 } catch { $result.Excel.VBOMAccessible=$false;$result.Excel.VBOMError=$_.Exception.Message }
 $sheet=$wb.Worksheets.Item(1)
 $key='00000000000000000000000000000000000000000001'
 $sheet.Range('A1').NumberFormat='@';$sheet.Range('A1').Value2=$key
 $sheet.Range('B1').Formula='=1+2';$sheet.Range('C1').Value2='FICTICIO'
 $p=Join-Path $testRoot 'smoke_excel.xlsx';$wb.SaveAs($p,51);$wb.Close($false);$wb=$null
 $wb=$xl.Workbooks.Open($p,0,$true)
 $result.Excel.Text44=($wb.Worksheets.Item(1).Range('A1').Value2 -ceq $key)
 $result.Excel.Formula=($wb.Worksheets.Item(1).Range('B1').Value2 -eq 3)
 $result.Excel.NumberFormat=$wb.Worksheets.Item(1).Range('A1').NumberFormat
} catch { $result.Excel.Error=$_.Exception.Message } finally { if($null -ne $wb){$wb.Close($false)};if($null -ne $xl){$xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)} }
$result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $testRoot 'office_capabilities.json') -Encoding UTF8
# Word opcional: a inicializacao ficou pendente nesta maquina.
if($TestWord){
try {
 $word=New-Object -ComObject Word.Application
 $word.Visible=$false;$word.DisplayAlerts=0;$word.AutomationSecurity=3
 $result.Word.Version=$word.Version
 $doc=$word.Documents.Add();$doc.Content.Text='TESTE FICTICIO DE AUDITORIA VBA'
 $range=$doc.Range(0,5);[void]$doc.Bookmarks.Add('AuditoriaTeste',$range)
 $p=Join-Path $testRoot 'smoke_word.docx';$doc.SaveAs2($p,16);$doc.Close(0);$doc=$null
 $doc=$word.Documents.Open($p,$false,$true)
 $result.Word.Bookmark=$doc.Bookmarks.Exists('AuditoriaTeste');$result.Word.Text=$doc.Content.Text.StartsWith('TESTE FICTICIO')
} catch { $result.Word.Error=$_.Exception.Message } finally {if($null -ne $doc){$doc.Close(0)};if($null -ne $word){$word.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)}}
}
$result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $testRoot 'office_capabilities.json') -Encoding UTF8
$result | ConvertTo-Json -Depth 8
