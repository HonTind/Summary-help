<#
.SYNOPSIS
    E-help 打包入口：产物统一输出到仓库根目录下的集中本地产物目录。

.DESCRIPTION
    本脚本只是 PyInstaller 的轻量包装，不改动业务源码、spec、授权逻辑或版本号。

    - 使用项目目录中现有的两份 spec：
        onedir（默认，推荐）  E-help_onedir.spec
        onefile               E-help.spec
    - 产物目录：<仓库根>\本地产物\E-help\dist\<release>\
    - 工作目录：<仓库根>\本地产物\E-help\build\<release>\
    - 脚本从自身位置定位项目根与仓库根，并在项目目录中调用 PyInstaller，
      因此 spec 中的相对入口（start_gui.py）保持有效。
    - 严格校验 release 名称：拒绝路径分隔符、'.'/'..'、目录逃逸，以及
      Windows 保留设备名（含 CON.txt / NUL.v1 / COM1.log 这类带扩展名的写法）；
      解析后的输出目录必须是集中产物目录的直接子目录。
    - 目标目录已存在时拒绝构建，绝不覆盖、合并或递归删除旧产物；
      每个 release 单独一个目录，历史产物原样保留。
    - 不使用任何递归删除命令清理旧产物，也不向 PyInstaller 传 --noconfirm，
      避免任何形式的自动确认覆盖。
    - -PlanOnly 为无写入预览：只校验并打印最终目录与完整命令行，
      不创建目录或文件，也不调用 PyInstaller。
    - 预览中的「完整命令行」会把解释器与每个参数按 PowerShell 规则单引号引用
      并加 '&'，含空格或单引号的路径可直接复制到 PowerShell 执行。

.PARAMETER Release
    必填。本次构建的唯一 release 名称，例如 20261007。
    只允许 [A-Za-z0-9._-] 且必须以字母或数字开头。

.PARAMETER Mode
    onedir（默认，推荐）或 onefile。

.PARAMETER PythonExe
    可选。用于构建的 Python 解释器路径。
    省略时从 PATH 查找 python.exe。请勿把个人安装路径硬编码进本脚本或提交到仓库。

.PARAMETER PlanOnly
    无写入预览：打印解析后的路径、spec 与完整命令行后退出，不产生任何副作用。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\tools\build_release.ps1 -Release 20261007 -PlanOnly

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\tools\build_release.ps1 -Release 20261007 -Mode onefile

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\tools\build_release.ps1 -Release 20261007 `
        -PythonExe "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Release,

    [ValidateSet('onedir', 'onefile')]
    [string]$Mode = 'onedir',

    [string]$PythonExe,

    [switch]$PlanOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ArtifactDirName = '本地产物'
$ProjectDirName  = 'E-help'

$SpecByMode = @{
    'onedir'  = 'E-help_onedir.spec'
    'onefile' = 'E-help.spec'
}

function Assert-ReleaseName {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) {
        throw 'release 名称不能为空。'
    }
    if ($Name -match '[\x00-\x1F<>:"/\\|?*]') {
        throw "release 名称含非法字符（路径分隔符或 Windows 保留字符）：$Name"
    }
    if ($Name.Contains('..')) {
        throw "release 名称不允许包含 '..'（目录逃逸）：$Name"
    }
    if ($Name -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') {
        throw "release 名称只允许 [A-Za-z0-9._-] 且必须以字母或数字开头：$Name"
    }
    if ($Name.EndsWith('.')) {
        throw "release 名称不能以 '.' 结尾（Windows 会静默丢弃）：$Name"
    }
    # Windows 保留设备名即使带扩展名也仍然保留：CON.txt、NUL.v1、COM1.log 都会被系统
    # 解析成设备名。因此取第一个 '.' 之前的部分判断，而不是只匹配完整名称。
    $deviceBase = $Name.Split('.')[0]
    if ($deviceBase -match '^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$') {
        throw "release 名称是 Windows 保留设备名（含带扩展名的写法，如 CON.txt），不能作为目录名：$Name"
    }
}

# 生成可直接复制到 PowerShell 中执行的参数文本：始终加单引号，
# 内部单引号按 PowerShell 规则双写，因此含空格、单引号的路径也能安全复制使用。
function ConvertTo-PowerShellLiteral {
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Value)

    return "'" + ($Value -replace "'", "''") + "'"
}

# 解析后的路径必须是 Parent 的直接子项；否则视为目录逃逸并中止。
function Assert-DirectChildPath {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Parent,
        [Parameter(Mandatory = $true)][string]$Label
    )

    $fullPath   = [System.IO.Path]::GetFullPath($Path)
    $fullParent = [System.IO.Path]::GetFullPath($Parent).TrimEnd('\', '/')
    $leaf       = [System.IO.Path]::GetFileName($fullPath)
    if ([string]::IsNullOrEmpty($leaf)) {
        throw "$Label 不是有效的子目录路径：$fullPath"
    }
    $expected = [System.IO.Path]::Combine($fullParent, $leaf)
    if (-not [string]::Equals($fullPath, $expected, [StringComparison]::OrdinalIgnoreCase)) {
        throw "$Label 解析后不在预期父目录内（目录逃逸）：`n  $fullPath`n  期望父目录：$fullParent"
    }
    if ([string]::Equals($fullParent, $fullPath, [StringComparison]::OrdinalIgnoreCase)) {
        throw "$Label 不能与父目录相同：$fullPath"
    }
    return $fullPath
}

# ---------- 从脚本自身定位项目根与仓库根 ----------

$scriptDir = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($scriptDir)) {
    $scriptDir = Split-Path -Parent $PSCommandPath
}
$scriptDir   = [System.IO.Path]::GetFullPath($scriptDir)
$projectRoot = [System.IO.Path]::GetFullPath((Join-Path $scriptDir '..'))
$repoRoot    = [System.IO.Path]::GetFullPath((Join-Path $projectRoot '..'))

if ([System.IO.Path]::GetFileName($projectRoot) -ne $ProjectDirName) {
    throw "脚本位置异常：期望项目目录名为 '$ProjectDirName'，实际为 '$projectRoot'。"
}

# ---------- 校验 release 名称与输出路径 ----------

Assert-ReleaseName -Name $Release

$artifactRoot = Join-Path $repoRoot $ArtifactDirName
$localRoot    = Join-Path $artifactRoot $ProjectDirName
$distRoot     = Join-Path $localRoot 'dist'
$buildRoot    = Join-Path $localRoot 'build'

$distPath = Assert-DirectChildPath -Path (Join-Path $distRoot $Release) -Parent $distRoot -Label 'dist 产物目录'
$workPath = Assert-DirectChildPath -Path (Join-Path $buildRoot $Release) -Parent $buildRoot -Label 'build 工作目录'

$specName = $SpecByMode[$Mode]
$specPath = Join-Path $projectRoot $specName
if (-not (Test-Path -LiteralPath $specPath -PathType Leaf)) {
    throw "找不到 spec 文件：$specPath"
}

if ($Mode -eq 'onedir') {
    $expectedExe = Join-Path (Join-Path $distPath 'E-help') 'E-help.exe'
} else {
    $expectedExe = Join-Path $distPath 'E-help.exe'
}

# 已存在即拒绝：不覆盖、不合并、不递归删除。
$existing = @()
foreach ($candidate in @($distPath, $workPath)) {
    if (Test-Path -LiteralPath $candidate) {
        $existing += $candidate
    }
}
if ($existing.Count -gt 0) {
    throw ("目标已存在，拒绝覆盖（不会清理旧产物）：`n  " + ($existing -join "`n  ") +
           "`n请改用另一个 release 名称，或由人工确认后处理旧目录。")
}

# ---------- 解析 Python 解释器 ----------

$resolvedPython = $null
if (-not [string]::IsNullOrWhiteSpace($PythonExe)) {
    if (-not (Test-Path -LiteralPath $PythonExe -PathType Leaf)) {
        throw "指定的 Python 解释器不存在：$PythonExe"
    }
    $resolvedPython = (Resolve-Path -LiteralPath $PythonExe).ProviderPath
} else {
    $pythonCommand = Get-Command -Name 'python.exe' -CommandType Application -ErrorAction SilentlyContinue
    if ($null -ne $pythonCommand) {
        $resolvedPython = $pythonCommand.Source
    }
}

if ($null -eq $resolvedPython -and -not $PlanOnly) {
    throw ('未能在 PATH 中找到 python.exe，请用 -PythonExe 指定解释器，例如：' +
           "`n  -PythonExe `"`$env:LOCALAPPDATA\Programs\Python\Python312\python.exe`"")
}

$pythonResolved = $null -ne $resolvedPython
$pythonDisplay = if ($pythonResolved) {
    $resolvedPython
} else {
    '<未解析的 python.exe：运行时必须用 -PythonExe 指定>'
}

# ---------- 预览 ----------

# 不用 --noconfirm：它会在输出目录已存在时自动确认覆盖/清理。
# 本脚本已在上面直接拒绝已存在的目标目录，这里不再允许任何自动确认覆盖。
$pyInstallerArgs = @(
    '-m', 'PyInstaller',
    '--distpath', $distPath,
    '--workpath', $workPath,
    $specPath
)

$argumentText = (@($pyInstallerArgs | ForEach-Object { ConvertTo-PowerShellLiteral -Value $_ })) -join ' '
if ($pythonResolved) {
    $commandDisplay = '& ' + (ConvertTo-PowerShellLiteral -Value $resolvedPython) + ' ' + $argumentText
    $commandNote    = ''
} else {
    # 没有真实解释器可显示，第一项只能给占位符；此时预览命令不可直接执行。
    $commandDisplay = '& <未解析的 python.exe> ' + $argumentText
    $commandNote    = '（第一项是占位符：用 -PythonExe 指定真实解释器后，命令行才可复制执行。）'
}

Write-Host ''
Write-Host '=== E-help 打包预览 ==='
Write-Host ('模式 Mode         : {0}' -f $Mode)
Write-Host ('Release           : {0}' -f $Release)
Write-Host ('仓库根            : {0}' -f $repoRoot)
Write-Host ('项目根            : {0}' -f $projectRoot)
Write-Host ('spec 文件         : {0}' -f $specPath)
Write-Host ('Python 解释器     : {0}' -f $pythonDisplay)
Write-Host ('产物目录 distpath : {0}' -f $distPath)
Write-Host ('工作目录 workpath : {0}' -f $workPath)
Write-Host ('期望产物          : {0}' -f $expectedExe)
Write-Host ('调用目录          : {0}' -f $projectRoot)
Write-Host ('完整命令行        : {0}' -f $commandDisplay)
if (-not [string]::IsNullOrEmpty($commandNote)) {
    Write-Host ('                    {0}' -f $commandNote)
}
Write-Host ''

if ($PlanOnly) {
    Write-Host 'PlanOnly：以上仅为预览。未创建任何目录或文件，未调用 PyInstaller。'
    return
}

# ---------- 构建 ----------

Write-Host '开始构建……'
Push-Location -LiteralPath $projectRoot
try {
    & $resolvedPython @pyInstallerArgs
    $buildExitCode = $LASTEXITCODE
} finally {
    Pop-Location
}

if ($buildExitCode -ne 0) {
    throw "PyInstaller 退出码 $buildExitCode，构建失败。未删除或清理任何已有产物。"
}

if (-not (Test-Path -LiteralPath $expectedExe -PathType Leaf)) {
    throw "PyInstaller 返回成功，但未找到期望的 EXE：$expectedExe"
}

Write-Host ''
Write-Host ('构建完成：{0}' -f $expectedExe)
if ($Mode -eq 'onedir') {
    Write-Host '提醒：onedir 的 EXE 与同目录下的 _internal 文件夹必须放在一起，快捷方式请指向 EXE，不能只拷贝 EXE。'
}
Write-Host ('本次产物目录：{0}' -f $distPath)
return
