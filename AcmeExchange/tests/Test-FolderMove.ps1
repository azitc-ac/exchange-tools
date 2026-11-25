<#
.SYNOPSIS
    Checks that -Teardown moves the Posh-ACME state folder in one step or not at all.

.DESCRIPTION
    Move-Item on a folder is not atomic under Windows PowerShell 5.1: with a file open inside
    (FileShare Read, no Delete) it throws, but has already moved part of the files. Teardown then
    left the state folder (account keys, orders) split across two folders. Move-FolderAside uses
    [IO.Directory]::Move, which fails completely in the same situation.

    A working directory inside the folder does NOT reproduce this - it takes an open file.

    The test proves its own discriminating power in the same run:
      - counter-check: plain Move-Item in the same setup must move only PART of the files
        (Windows PowerShell 5.1 only; other versions report SKIP),
      - rollback mutant: Move-FolderAside with Move-Item instead of [IO.Directory]::Move must be
        caught by the open-file check.

.EXAMPLE
    .\tests\Test-FolderMove.ps1
#>
[CmdletBinding()]
param([string]$Engine)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if (-not $Engine) { $Engine = Join-Path $root 'Invoke-AcmeExchangeCert.ps1' }

$script:pass = 0; $script:fail = 0
function Check([string]$Name, [scriptblock]$Test) {
    try { if (& $Test) { Write-Host "[OK  ] $Name" -ForegroundColor Green; $script:pass++ }
          else { Write-Host "[FAIL] $Name" -ForegroundColor Red; $script:fail++ } }
    catch { Write-Host "[FAIL] $Name -> $($_.Exception.Message)" -ForegroundColor Red; $script:fail++ }
}

$ast = [Management.Automation.Language.Parser]::ParseFile($Engine, [ref]$null, [ref]$null)
function Get-Fn([string]$Name) { $ast.FindAll({ $args[0] -is [Management.Automation.Language.FunctionDefinitionAst] -and $args[0].Name -eq $Name }, $true) | Select-Object -First 1 }
$fn = Get-Fn 'Move-FolderAside'
$td = Get-Fn 'Invoke-Teardown'
Check 'Move-FolderAside exists' { $null -ne $fn }
Check 'Invoke-Teardown exists' { $null -ne $td }
if (-not $fn -or -not $td) { Write-Host "`n$script:pass passed, $script:fail failed"; exit 1 }

# --- static: Teardown moves the state folder only through Move-FolderAside
$cmds = @($td.FindAll({ $args[0] -is [Management.Automation.Language.CommandAst] }, $true) | ForEach-Object { $_.GetCommandName() })
Check 'Invoke-Teardown calls Move-FolderAside' { $cmds -contains 'Move-FolderAside' }
Check 'Invoke-Teardown has no Move-Item (not atomic on folders under PS 5.1)' { -not ($cmds | Where-Object { $_ -in 'Move-Item', 'move', 'mi', 'mv' }) }

# --- behaviour
$work = Join-Path $env:TEMP ('acme-foldermove-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
function New-Src([string]$Name) {
    $d = Join-Path $work $Name
    New-Item -ItemType Directory -Force $d | Out-Null
    0..99 | ForEach-Object { Set-Content -LiteralPath (Join-Path $d ('f{0:D3}.txt' -f $_)) 'x' }
    $d
}
function Count([string]$Dir) { @(Get-ChildItem -LiteralPath $Dir -Recurse -File -Force -ErrorAction SilentlyContinue).Count }
# Runs the open-file case against a given Move-FolderAside text; $true = no split state.
function Test-OpenFileCase([string]$FnText, [string]$Tag) {
    . ([scriptblock]::Create($FnText))
    $src = New-Src "open-$Tag"; $dst = "$src.removed"
    $fs = [IO.File]::Open((Join-Path $src 'f050.txt'), 'Open', 'Read', 'Read')
    $threw = $false
    try { Move-FolderAside -Path $src -Destination $dst } catch { $threw = $true } finally { $fs.Dispose() }
    $script:lastDetail = "threw=$threw, source $(Count $src)/100, target exists=$(Test-Path -LiteralPath $dst) ($(Count $dst) files)"
    $threw -and (Count $src) -eq 100 -and -not (Test-Path -LiteralPath $dst)
}

try {
    New-Item -ItemType Directory -Force $work | Out-Null
    . ([scriptblock]::Create($fn.Extent.Text))

    $a = New-Src 'plain'
    Move-FolderAside -Path $a -Destination "$a.removed"
    Check 'unlocked folder is moved completely' { -not (Test-Path -LiteralPath $a) -and (Count "$a.removed") -eq 100 }

    $b = New-Src 'exists'; New-Item -ItemType Directory -Force "$b.removed" | Out-Null
    $threw = $false; try { Move-FolderAside -Path $b -Destination "$b.removed" } catch { $threw = $true }
    Check 'existing target: throws, does not move INTO it' { $threw -and (Count $b) -eq 100 -and (Count "$b.removed") -eq 0 }

    $ok = Test-OpenFileCase $fn.Extent.Text 'fix'
    Write-Host "       $script:lastDetail"
    Check 'open file inside: fails completely, nothing moved' { $ok }

    # counter-check: the setup must be able to show the bug at all
    $c = New-Src 'counter'
    $fs = [IO.File]::Open((Join-Path $c 'f050.txt'), 'Open', 'Read', 'Read')
    try { Move-Item -LiteralPath $c -Destination "$c.removed" -ErrorAction Stop } catch { } finally { $fs.Dispose() }
    $moved = Count "$c.removed"
    if ($PSVersionTable.PSVersion.Major -eq 5) {
        Check "counter-check: Move-Item in the same setup moves only part ($moved of 100)" { $moved -gt 0 -and $moved -lt 100 }
    } else { Write-Host "[SKIP] counter-check (PS $($PSVersionTable.PSVersion); measured on 5.1 only): Move-Item moved $moved of 100" }

    # rollback mutant: the open-file check must catch Move-Item inside Move-FolderAside
    $needle = '[System.IO.Directory]::Move($src, $dst)'
    Check 'mutant can be built (fix line found)' { $fn.Extent.Text.Contains($needle) }
    $mutant = $fn.Extent.Text.Replace($needle, 'Move-Item -LiteralPath $src -Destination $dst -ErrorAction Stop')
    if ($PSVersionTable.PSVersion.Major -eq 5) {
        $caught = -not (Test-OpenFileCase $mutant 'mutant')
        Write-Host "       mutant: $script:lastDetail"
        Check 'rollback mutant (Move-Item) is caught by the open-file check' { $caught }
    } else { Write-Host '[SKIP] rollback mutant (PS 5.1 only)' }
} finally {
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "`n$script:pass passed, $script:fail failed"
exit ([int]($script:fail -gt 0))
