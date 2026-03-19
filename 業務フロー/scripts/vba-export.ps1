#=============================================================================
# vba-export.ps1
# 用途 : VBEからエクスポートしたShift-JISの.clsをUTF-8に変換して
#        vba-files/Class/ に上書きする（gitコミット前に実行）
#
# 使い方:
#   1. ExcelのVBE（Alt+F11）→ 各モジュールを右クリック → ファイルのエクスポート
#      → 保存先を vba-files/_vbe/ に指定
#   2. このスクリプトをPowerShellで実行
#      > .\scripts\vba-export.ps1
#   3. vba-files/Class/ がUTF-8で上書きされる
#   4. git add / git commit でコミット
#=============================================================================

$srcDir = Join-Path $PSScriptRoot "..\vba-files\_vbe"
$dstDir = Join-Path $PSScriptRoot "..\vba-files\Class"

if (-not (Test-Path $srcDir)) {
    Write-Host "[エラー] $srcDir が見つかりません" -ForegroundColor Red
    Write-Host "VBEからのエクスポート先を vba-files/_vbe/ にしてください"
    exit 1
}

$files = Get-ChildItem -Path $srcDir -Filter "*.cls"

if ($files.Count -eq 0) {
    Write-Host "[警告] $srcDir に .cls ファイルが見つかりません" -ForegroundColor Yellow
    Write-Host "VBEからのエクスポート先を vba-files/_vbe/ にしてください"
    exit 1
}

# 出力先フォルダを確保
if (-not (Test-Path $dstDir)) {
    New-Item -ItemType Directory -Path $dstDir | Out-Null
}

$sjis   = [System.Text.Encoding]::GetEncoding("shift_jis")
$utf8nb = New-Object System.Text.UTF8Encoding $false  # BOMなしUTF-8

$count = 0
foreach ($file in $files) {
    $content = [System.IO.File]::ReadAllText($file.FullName, $sjis)
    $dstPath = Join-Path $dstDir $file.Name
    [System.IO.File]::WriteAllText($dstPath, $content, $utf8nb)
    Write-Host "[変換完了] _vbe\$($file.Name)  →  Class\$($file.Name)"
    $count++
}

Write-Host ""
Write-Host "$count 件のファイルをUTF-8に変換しました" -ForegroundColor Green
Write-Host "次のステップ: git add vba-files/Class/ してコミットしてください" -ForegroundColor Cyan
