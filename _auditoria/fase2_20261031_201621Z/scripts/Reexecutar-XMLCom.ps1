$ErrorActionPreference = 'Stop'
$result = [ordered]@{Timestamp=(Get-Date).ToString('o'); Kind='MSXML COM ficticio; nao executa VBA'; Tests=@()}
function New-TestDom {
    $dom = New-Object -ComObject MSXML2.DOMDocument.6.0
    $dom.async = $false
    $dom.validateOnParse = $false
    $dom.resolveExternals = $false
    $dom.setProperty('ProhibitDTD', $true)
    $dom.setProperty('SelectionLanguage', 'XPath')
    return $dom
}
try {
    $key = '00000000000000000000000000000000000000000001'
    $dom = New-TestDom
    $loaded = $dom.loadXML('<nfeProc xmlns="http://www.portalfiscal.inf.br/nfe"><NFe><infNFe Id="NFe' + $key + '"><ide><nNF>1</nNF></ide></infNFe></NFe></nfeProc>')
    $result.Tests += [ordered]@{Name='XML sintetico bem formado'; Pass=[bool]$loaded}
    $nodes = $dom.selectNodes("//*[local-name()='infNFe']")
    $result.Tests += [ordered]@{Name='XPath com namespace'; Pass=($nodes.length -eq 1)}
    $result.Tests += [ordered]@{Name='Chave no atributo preservada como texto'; Pass=($nodes.item(0).getAttribute('Id') -ceq ('NFe' + $key))}
    $bad = New-TestDom
    $result.Tests += [ordered]@{Name='XML malformado rejeitado'; Pass=(-not $bad.loadXML('<NFe><infNFe></NFe>'))}
    $dtd = New-TestDom
    $result.Tests += [ordered]@{Name='DTD interna rejeitada'; Pass=(-not $dtd.loadXML('<!DOCTYPE NFe [<!ENTITY teste "FICTICIO">]><NFe>&teste;</NFe>'))}
    $many = New-TestDom
    [void]$many.loadXML('<raiz><infNFe/><infNFe/></raiz>')
    $result.Tests += [ordered]@{Name='Dois infNFe identificados para recusa pelo carregador'; Pass=($many.selectNodes("//*[local-name()='infNFe']").length -eq 2)}
} catch { $result.Error = $_.Exception.Message }
$result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath 'C:\Users\MARCOSALEXANDREPERDO\OneDrive - MARCOS DESPACHANTE\Marcos Despachante\Software\VBA-Excel\_auditoria\fase2_20261031_201621Z\regressao-xml-com.json' -Encoding UTF8
$result | ConvertTo-Json -Depth 8
