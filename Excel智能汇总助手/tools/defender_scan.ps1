# Windows Defender custom scan for the onedir distribution (v1.5.4)
# ASCII-only: keep this file free of non-ASCII characters.
#
# Scans the release distribution directory passed as -DistDir, then reports
# detections that are NEW since the scan began (i.e. produced by THIS build),
# plus a summary of historical detections for context.
#
# -DistDir is REQUIRED and must be an existing directory. This script no longer
# guesses a directory under the project, and never picks "the first" subdirectory
# on its own. The current centralized release directory layout is documented in
# README; pass the concrete path explicitly.
#
# Nothing here runs automatically: run this script only when a deliberate
# Defender scan of a release is wanted.
#
# Argument handling runs before any scan call:
#   -Help / --help / -h  print usage only and exit 0; supplying a valid
#                        -DistDir as well does NOT turn the request into a scan.
#                        Accepted whether PowerShell binds the token to the
#                        -Help switch or leaves it in the argument list.
#   An unknown option or an extra positional argument is rejected with exit 2,
#   so a mistyped command can never fall through into a real scan.
#   PowerShell's simple parameter binder binds a lone "--bogus" / "-Bogus" to
#   -DistDir and leaves the argument list empty, so after the help branch a
#   -DistDir value starting with "-" is also rejected with exit 2 (unknown
#   option), before any directory validation. A directory whose name really
#   starts with "-" stays reachable by spelling it explicitly, e.g.
#   -DistDir .\--bogus or an absolute path.
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File tools\defender_scan.ps1 -DistDir <release dist directory>
#   powershell -ExecutionPolicy Bypass -File tools\defender_scan.ps1 -Help
#
# Exit codes: 0 = pass (or -Help), 1 = fail/warn, 2 = invalid usage (missing or
# invalid -DistDir, unknown option, extra positional argument). A PowerShell
# parameter binder error (for example the same parameter bound twice, or a
# missing value for -DistDir) exits 1 before this script body runs.
param(
    [string]$DistDir,
    [switch]$Help
)

$ErrorActionPreference = "Continue"

function Show-Usage {
    Write-Output "Usage: powershell -ExecutionPolicy Bypass -File tools\defender_scan.ps1 -DistDir <release dist directory>"
    Write-Output "       powershell -ExecutionPolicy Bypass -File tools\defender_scan.ps1 -Help"
    Write-Output "Scans exactly one release distribution directory with Windows Defender;"
    Write-Output "nothing is scanned unless -DistDir points at an existing directory."
}

$helpSpellings = @('-h', '-help', '--h', '--help')
$helpRequested = [bool]$Help

# The help spellings are recognized in every binding style: PowerShell binds
# them to -Help when the script is launched with -File, keeps them in the
# argument list when the script is called in-process, and can even bind a bare
# "--help" positionally to $DistDir. None of these paths may fall into a scan,
# so a -DistDir value that is itself a help spelling counts as a help request
# rather than as a directory name.
foreach ($a in @($args) + @($DistDir)) {
    if (($a -is [string]) -and ($helpSpellings -contains $a)) { $helpRequested = $true }
}

if ($helpRequested) {
    Show-Usage
    exit 0
}

# Unknown option bound to -DistDir by the simple parameter binder (--bogus,
# -Bogus, ...): the argument list is empty, so without this check the value
# would be treated as a directory name and a same-named directory would be
# scanned. Reject it here, before any path validation or resolution.
if ((-not [string]::IsNullOrWhiteSpace($DistDir)) -and $DistDir.StartsWith('-')) {
    Write-Output ("FAIL: -DistDir looks like an unknown option: " + $DistDir)
    Write-Output "Pass an explicit path (for example .\--bogus or an absolute path); nothing was scanned."
    Show-Usage
    exit 2
}

if ($args.Count -gt 0) {
    Write-Output ("FAIL: unexpected argument(s): " + ($args -join ' '))
    Write-Output "Unknown options and extra positional arguments are rejected; nothing was scanned."
    Show-Usage
    exit 2
}

if ([string]::IsNullOrWhiteSpace($DistDir)) {
    Write-Output "FAIL: missing -DistDir parameter."
    Show-Usage
    exit 2
}
if (-not (Test-Path -LiteralPath $DistDir -PathType Container)) {
    Write-Output "FAIL: -DistDir is not an existing directory: $DistDir"
    exit 2
}
$DistDir = (Resolve-Path -LiteralPath $DistDir).ProviderPath

Write-Output "=== Windows Defender custom scan ==="
Write-Output "Target: $DistDir"
Write-Output "Scan start (local): $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

$scanStart = Get-Date
$scanError = $null
try {
    Start-MpScan -ScanType CustomScan -ScanPath $DistDir -ErrorAction Stop
    Write-Output "Scanner finished without error."
} catch {
    $scanError = $_.Exception.Message
    Write-Output ("Scanner raised: " + $scanError)
}

# Detections created during/after this scan = threats from the current build
$new = @(Get-MpThreatDetection -ErrorAction SilentlyContinue | Where-Object {
    $_.InitialDetectionTime -ge $scanStart.AddMinutes(-1)
})
$all = @(Get-MpThreatDetection -ErrorAction SilentlyContinue)

Write-Output ("New detections since scan: " + $new.Count)
Write-Output ("Total historical detections: " + $all.Count)

foreach ($d in $all) {
    Write-Output ("  [history] " + $d.InitialDetectionTime.ToString("yyyy-MM-dd HH:mm") +
                  "  " + ($d.Resources -join '; '))
}

if ($new.Count -gt 0) {
    foreach ($d in $new) {
        Write-Output ("  [NEW] " + ($d.Resources -join '; ') + "  " + $d.InitialDetectionTime)
    }
    Write-Output "RESULT: FAIL - current build flagged by Defender"
    exit 1
}

if ($scanError) {
    Write-Output "RESULT: WARN - scan did not complete cleanly"
    exit 1
}

Write-Output "RESULT: PASS - current build not flagged by Defender"
exit 0
