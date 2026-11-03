param([string]$TestRoot = (Join-Path $env:TEMP ('VBA_ExcelCom_' + [guid]::NewGuid().ToString('N'))))
$ErrorActionPreference = 'Stop'
$testPath = [IO.Path]::GetFullPath($TestRoot)
$prefix = [IO.Path]::GetFullPath($env:TEMP).TrimEnd('\') + '\'
if (-not $testPath.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase) -or $testPath -match '(?i)OneDrive') {
    throw 'O teste exige pasta temporaria fora do OneDrive.'
}
if (Test-Path -LiteralPath $testPath) { throw 'Escolha uma pasta nova; nenhum arquivo existente sera sobrescrito.' }
[void](New-Item -ItemType Directory -Path $testPath)
$r = [ordered]@{Timestamp=(Get-Date).ToString('o'); Kind='Excel COM ficticio; nao executa VBA'; Tests=@(); TestRoot=$testPath}
$x = $null
$b = $null
try {
    $x = New-Object -ComObject Excel.Application
    $x.Visible=$false; $x.DisplayAlerts=$false; $x.EnableEvents=$false; $x.AutomationSecurity=3
    $b=$x.Workbooks.Add(); $s=$b.Worksheets.Item(1)
    $key='00000000000000000000000000000000000000000001'
    $large='01234567890123456789012345678901234567890123'
    [void]$s.Range('A1').ClearFormats(); $s.Range('A1').Value2=$large
    $r.Tests += [ordered]@{Name='Reproducao da perda no formato padrao'; Pass=([string]$s.Range('A1').Value2 -cne $large); Type=$s.Range('A1').Value2.GetType().FullName}
    foreach ($pair in @(@('A2',$key),@('A3',$large))) {
        $c=$s.Range($pair[0]); $c.NumberFormat='@'; $c.Value2=[string]$pair[1]
        $r.Tests += [ordered]@{Name=('Chave texto '+$pair[0]); Pass=($c.Value2 -ceq $pair[1]); Type=$c.Value2.GetType().FullName; Length=$c.Value2.Length}
    }
    $s.Range('B2').Formula='=1+2'; $s.Range('B2').NumberFormat='0.00'
    $formula=$s.Range('B2').Formula; $fmt=$s.Range('B2').NumberFormat
    $p=Join-Path $testPath 'regressao_excel.xlsx'
    $b.SaveAs($p,51); $b.Close($false); $b=$null
    $b=$x.Workbooks.Open($p,0,$true); $s=$b.Worksheets.Item(1)
    $r.Tests += [ordered]@{Name='Chave apos reabrir'; Pass=($s.Range('A3').Value2 -ceq $large)}
    $r.Tests += [ordered]@{Name='Formula e formato apos reabrir'; Pass=($s.Range('B2').Formula -eq $formula -and $s.Range('B2').Value2 -eq 3 -and $s.Range('B2').NumberFormat -eq $fmt)}
    $b.Close($false); $b=$null
    $b=$x.Workbooks.Add(); $s=$b.Worksheets.Item(1)
    $s.Range('A1').Value2='CABECALHO'; $s.Range('A2').Value2='FICTICIO_1'; $s.Range('A3').Value2='FICTICIO_2'
    $s.Range('N2').Formula='=A2'; $s.Range('O2').Value2='SENTINELA_1'; $s.Range('O3').Value2='SENTINELA_2'
    [void]$s.Range('A2:M2').Delete(-4162)
    $r.Tests += [ordered]@{Name='Reproducao do risco de exclusao parcial'; Pass=($s.Range('A2').Value2 -eq 'FICTICIO_2' -and $s.Range('O2').Value2 -eq 'SENTINELA_1'); FormulaN2=$s.Range('N2').Formula}
} catch { $r.Error=$_.Exception.Message } finally {
    if($null -ne $b){$b.Close($false)}
    if($null -ne $x){$x.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($x)}
}
$r | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $testPath 'resultado_excel.json') -Encoding UTF8
$r | ConvertTo-Json -Depth 8
if ($r.Contains('Error') -or @($r.Tests | Where-Object {-not $_.Pass}).Count -gt 0) { exit 1 }
