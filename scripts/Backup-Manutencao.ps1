#Requires -Version 7.0
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$DiretorioBackup,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string[]]$Arquivo
)

$ErrorActionPreference = 'Stop'
$raiz = Split-Path -Parent $PSScriptRoot
$prefixoRaiz = $raiz + [IO.Path]::DirectorySeparatorChar
$comparacao = [StringComparison]::OrdinalIgnoreCase

function Confirmar-CaminhoSemAtalhos([System.IO.FileSystemInfo]$Item) {
    while ($null -ne $Item) {
        if ($Item.LinkType) { throw "Atalho ou link no caminho: $($Item.FullName)" }
        if ($Item -is [IO.DirectoryInfo]) {
            $Item = $Item.Parent
        } else {
            $Item = $Item.Directory
        }
    }
}

$externo = Get-Item -LiteralPath $DiretorioBackup -Force
if ($externo -isnot [IO.DirectoryInfo] -or
    $externo.FullName.Equals($raiz, $comparacao) -or
    $externo.FullName.StartsWith($prefixoRaiz, $comparacao)) {
    throw 'Escolha uma pasta existente fora do repositório.'
}
Confirmar-CaminhoSemAtalhos $externo

$selecionados = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($relativo in $Arquivo) {
    if ([IO.Path]::IsPathRooted($relativo) -or $relativo.Contains(':')) {
        throw "Use caminhos relativos ao repositório: $relativo"
    }
    $origem = [IO.Path]::GetFullPath($relativo, $raiz)
    if (-not $origem.StartsWith($prefixoRaiz, $comparacao) -or $origem -match '(^|[\\/])\.git([\\/]|$)') {
        throw "Caminho fora da árvore de trabalho: $relativo"
    }
    $item = Get-Item -LiteralPath $origem -Force
    if ($item -isnot [IO.FileInfo]) { throw "Informe um arquivo: $relativo" }
    Confirmar-CaminhoSemAtalhos $item
    [void]$selecionados.Add($item.FullName)
}

$id = [DateTime]::UtcNow.ToString("yyyyMMdd'T'HHmmss'Z'") + '-' + [guid]::NewGuid().ToString('N').Substring(0,8)
$destino = New-Item -ItemType Directory -Path ([IO.Path]::Combine($externo.FullName, $id))
$arquivos = foreach ($origem in ($selecionados | Sort-Object)) {
    $relativo = [IO.Path]::GetRelativePath($raiz, $origem)
    $copia = [IO.Path]::Combine($destino.FullName, 'arquivos', $relativo)
    [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($copia))
    $hash = (Get-FileHash -LiteralPath $origem -Algorithm SHA256).Hash
    [IO.File]::Copy($origem, $copia, $false)
    if ((Get-FileHash -LiteralPath $copia -Algorithm SHA256).Hash -ne $hash) {
        throw "SHA-256 divergente para $relativo em $($destino.FullName)"
    }
    [ordered]@{ caminho = $relativo.Replace('\', '/'); sha256 = $hash }
}

$manifesto = [ordered]@{ id = $id; arquivos = @($arquivos) }
$manifesto | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $destino.FullName 'sha256.json') -Encoding utf8
$destino.FullName
