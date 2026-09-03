<#
.SYNOPSIS
  Loads KEY=VALUE pairs from .env in the project root into a hashtable.
  Does not touch process/environment variables — callers read the returned map.
#>
param(
    [string]$EnvPath = "C:\Users\vmichalovska\Documents\VK_medijos_stebesena\.env"
)
$map = @{}
if (-not (Test-Path $EnvPath)) {
    throw ".env failas nerastas: $EnvPath — sukurkite jį pagal .env.example."
}
Get-Content $EnvPath | ForEach-Object {
    $line = $_.Trim()
    if (-not $line -or $line.StartsWith("#")) { return }
    $idx = $line.IndexOf("=")
    if ($idx -lt 1) { return }
    $key = $line.Substring(0, $idx).Trim()
    $val = $line.Substring($idx + 1).Trim()
    $map[$key] = $val
}
return $map
