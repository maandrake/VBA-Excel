[CmdletBinding()]
param([string]$WorkspaceRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path)
$ErrorActionPreference = 'Stop'
$manifestPath = Join-Path $PSScriptRoot '../manifesto-final-fase8.json'
$digestPath = Join-Path $PSScriptRoot '../manifesto-final-fase8.sha256'
$expectedDigest = ((Get-Content -LiteralPath $digestPath -Raw).Trim() -split '\s+')[0]
$actualDigest = (Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash.ToLowerInvariant()
if ($actualDigest -ne $expectedDigest) { throw 'Digest do manifesto diverge.' }
$manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$resolvedRoot = (Resolve-Path -LiteralPath $WorkspaceRoot).Path.TrimEnd('\', '/')
$rootPrefix = $resolvedRoot + [IO.Path]::DirectorySeparatorChar
$failures = @()
foreach ($entry in $manifest.files) {
    $targetPath = [IO.Path]::GetFullPath((Join-Path $resolvedRoot $entry.relative))
    if (-not $targetPath.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase)) { throw 'Caminho fora do workspace.' }
    if (-not (Test-Path -LiteralPath $targetPath -PathType Leaf)) {
        $failures += [pscustomobject]@{relative=$entry.relative; reason='AUSENTE'}
        continue
    }
    $actualHash = (Get-FileHash -LiteralPath $targetPath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($actualHash -ne $entry.sha256) { $failures += [pscustomobject]@{relative=$entry.relative; reason='SHA256_DIVERGENTE'} }
}
[pscustomobject]@{manifest='PASS'; checked=@($manifest.files).Count; failures=$failures.Count}
if ($failures.Count -gt 0) { $failures | Format-Table -AutoSize; throw 'Falha na verificacao de integridade.' }
# Somente leitura. Nao abre Office, nao altera politicas, nao restaura nem exclui arquivos.
