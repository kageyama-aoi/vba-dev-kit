#=============================================================================
# vba-import.ps1
# 用途 : git管理のUTF-8 .clsファイルをShift-JISに変換して _vbe/ に出力する
#        VBEにインポートする前にこのスクリプトを実行してください
#
# 使い方:
#   1. このスクリプトをPowerShellで実行
#      > .\scripts\vba-import.ps1
#   2. vba-files/_vbe/ に変換済みSJISファイルが生成される
#   3. ExcelのVBE（Alt+F11）→ ファイルからインポートで _vbe/ 内のファイルを選択
#=============================================================================

$srcDir = Join-Path $PSScriptRoot "..\vba-files\Class"
$dstDir = Join-Path $PSScriptRoot "..\vba-files\_vbe"

# 出力先フォルダを確保
if (-not (Test-Path $dstDir)) {
    New-Item -ItemType Directory -Path $dstDir | Out-Null
}

$files = Get-ChildItem -Path $srcDir -Filter "*.cls"

if ($files.Count -eq 0) {
    Write-Host "[警告] $srcDir に .cls ファイルが見つかりません" -ForegroundColor Yellow
    exit 1
}

$sjis = [System.Text.Encoding]::GetEncoding("shift_jis")
$utf8 = [System.Text.Encoding]::UTF8

$count = 0
foreach ($file in $files) {
    $content = [System.IO.File]::ReadAllText($file.FullName, $utf8)
    $dstPath = Join-Path $dstDir $file.Name
    [System.IO.File]::WriteAllText($dstPath, $content, $sjis)
    Write-Host "[変換完了] $($file.Name)  →  _vbe\$($file.Name)"
    $count++
}

Write-Host ""
Write-Host "$count 件のファイルをShift-JISに変換しました" -ForegroundColor Green
Write-Host "インポート先: $dstDir" -ForegroundColor Cyan
