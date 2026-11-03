$ErrorActionPreference='Stop'
$work='C:\Users\MARCOSALEXANDREPERDO\AppData\Local\Temp\vscode-agent-codex-sandbox-JTJOQt\VBA_Fase3_20261031_213925Z'
$base='C:\Users\MARCOSALEXANDREPERDO\OneDrive - MARCOS DESPACHANTE\Marcos Despachante\Software\VBA-Excel\_auditoria\fase3_20261031_213925Z'
$dst=Join-Path $work 'XML'
if(Test-Path -LiteralPath $dst){throw 'Destino existente.'}
[void][IO.Directory]::CreateDirectory($dst)
$regex=New-Object -ComObject VBScript.RegExp
$regex.Pattern='^-?[0-9]+(\.[0-9]+)?$';$regex.Global=$false;$regex.IgnoreCase=$true
function Num([string]$s){
 $s=$s.Trim();$v=0.0
 $ok=$regex.Test($s) -and [double]::TryParse($s,[Globalization.NumberStyles]::AllowLeadingSign -bor [Globalization.NumberStyles]::AllowDecimalPoint,[Globalization.CultureInfo]::InvariantCulture,[ref]$v)
 return [pscustomobject]@{Valid=$ok;Value=$v}
}
$key='00000000000000000000000000000000000000000001'
$valid='<NFe xmlns="http://www.portalfiscal.inf.br/nfe"><infNFe Id="NFe'+$key+'"><ide><nNF>1001</nNF></ide><det nItem="1"><prod><cProd>FICTICIO_1</cProd><vProd>100.00</vProd></prod><imposto><ICMS><ICMS00><vBC>50.00</vBC><vICMS>1.10</vICMS></ICMS00></ICMS><PIS><PISAliq><vPIS>0.10</vPIS></PISAliq></PIS><COFINS><COFINSAliq><vCOFINS>0.50</vCOFINS></COFINSAliq></COFINS></imposto></det><det nItem="2"><prod><cProd>FICTICIO_2</cProd><vProd>200.00</vProd></prod><imposto><ICMS><ICMS00><vBC>100.00</vBC><vICMS>2.20</vICMS></ICMS00></ICMS><PIS><PISAliq><vPIS>0.20</vPIS></PISAliq></PIS><COFINS><COFINSAliq><vCOFINS>1.00</vCOFINS></COFINSAliq></COFINS></imposto></det><total><ICMSTot><vProd>300.00</vProd><vBC>150.00</vBC><vICMS>3.30</vICMS><vPIS>0.30</vPIS><vCOFINS>1.50</vCOFINS></ICMSTot></total></infNFe></NFe>'
$cases=@(
 @{Name='ValidoEstruturalDoisItens';Xml=$valid;Structure=$true;Errors=0},
 @{Name='Malformado';Xml='<NFe><infNFe></NFe>';Structure=$false;Errors=0},
 @{Name='SemInfNFe';Xml='<NFe/>';Structure=$false;Errors=0},
 @{Name='MultiplosInfNFe';Xml=('<raiz><infNFe/><infNFe/></raiz>');Structure=$false;Errors=0},
 @{Name='ChaveInvalida';Xml=$valid.Replace($key,'123');Structure=$false;Errors=0},
 @{Name='NumeroInvalido';Xml=$valid.Replace('<nNF>1001</nNF>','<nNF>ABC</nNF>');Structure=$false;Errors=0},
 @{Name='NumeroZero';Xml=$valid.Replace('<nNF>1001</nNF>','<nNF>0</nNF>');Structure=$false;Errors=0},
 @{Name='NumeroLongo';Xml=$valid.Replace('<nNF>1001</nNF>','<nNF>1234567890</nNF>');Structure=$false;Errors=0},
 @{Name='ValorInvalidoSegundoItem';Xml=$valid.Replace('<vPIS>0.20</vPIS>','<vPIS>ABC</vPIS>');Structure=$true;Errors=1},
 @{Name='ValorVazioSegundoItem';Xml=$valid.Replace('<vPIS>0.20</vPIS>','<vPIS></vPIS>');Structure=$true;Errors=1},
 @{Name='OverflowSegundoItem';Xml=$valid.Replace('<vPIS>0.20</vPIS>',('<vPIS>'+('9'*400)+'</vPIS>'));Structure=$true;Errors=1},
 @{Name='DTD';Xml='<!DOCTYPE NFe [<!ENTITY teste "FICTICIO">]><NFe>&teste;</NFe>';Structure=$false;Errors=0})
$r=[ordered]@{Timestamp=(Get-Date).ToString('o');Modality='MSXML e RegExp COM; conversao/somatorio/erros modelados em .NET, nao e ValidarXMLNFe';FiscalRulesExecuted=$false;Results=@();ToleranceExisting=0.01}
foreach($case in $cases){
 $doc=New-Object -ComObject MSXML2.DOMDocument.6.0
 $doc.async=$false;$doc.validateOnParse=$false;$doc.resolveExternals=$false
 $doc.setProperty('ProhibitDTD',$true);$doc.setProperty('SelectionLanguage','XPath')
 $loaded=$doc.loadXML($case.Xml);$structure=$false;$reason=$null;$sums=@{};$errors=@()
 if($loaded){
  $nodes=$doc.selectNodes("//*[local-name()='infNFe']")
  if($nodes.length -eq 1){
   $inf=$nodes.item(0);$id=[string]$inf.getAttribute('Id');$n=$inf.selectSingleNode("./*[local-name()='ide']/*[local-name()='nNF']")
   $nn=if($null -eq $n){''}else{[string]$n.text}
   $structure=($id -match '^NFe[0-9]{44}$' -and $nn -match '^[0-9]{1,9}$' -and [double]$nn -gt 0 -and $null -eq $doc.doctype)
  }
 }else{$reason=$doc.parseError.reason}
 if($structure){
  foreach($field in @('vProd','vBC','vICMS','vPIS','vCOFINS')){
   $sum=0.0
   $nodes=$doc.selectNodes("//*[local-name()='det']//*[local-name()='"+$field+"']")
   foreach($node in $nodes){
    $number=Num ([string]$node.text)
    if($number.Valid){$sum+=$number.Value}else{$errors+=@{Field=$field;Message='Valor numerico no somatorio';Text=[string]$node.text;ExpectedVbaState='INCORRETO (pendente de execucao real)'}}
   }
   $total=$doc.selectSingleNode("//*[local-name()='total']/*[local-name()='ICMSTot']/*[local-name()='"+$field+"']")
   $converted=Num ([string]$total.text)
   $sums[$field]=@{Sum=$sum;XmlTotal=$converted.Value;Matches=($converted.Valid -and [math]::Abs($sum-$converted.Value) -le 0.01)}
  }
 }
 $pass=($structure -eq $case.Structure -and $errors.Count -eq $case.Errors)
 if($case.Name -eq 'ValidoEstruturalDoisItens'){$pass=$pass -and @($sums.Values|Where-Object {-not $_.Matches}).Count -eq 0}
 [IO.File]::WriteAllText((Join-Path $dst ($case.Name+'.xml')),$case.Xml,(New-Object Text.UTF8Encoding($false)))
 $r.Results+=@{Name=$case.Name;Status=if($pass){'PASS'}else{'FAIL'};Parsed=$loaded;StructureAccepted=$structure;NumericErrors=$errors;Sums=$sums;ParserReason=$reason;ReportVbaExecuted=$false}
 [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc)
}
$r|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $base 'xml-fase3-com.json') -Encoding UTF8
$r.Results|ForEach-Object {[pscustomobject]@{Name=$_['Name'];Status=$_['Status'];Structure=$_['StructureAccepted'];NumericErrors=@($_['NumericErrors']).Count}}|ConvertTo-Json
