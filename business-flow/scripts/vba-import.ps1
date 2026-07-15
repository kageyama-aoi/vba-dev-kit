#=============================================================================
# vba-import.ps1
# 用途 : git管理のUTF-8 VBAファイルをShift-JISに変換して _vbe/ に出力する
#        VBEにインポートする前にこのスクリプトを実行してください
#
# 対応ファイル:
#   vba-files/Class/*.cls    → クラスモジュール
#   vba-files/Module/*.bas   → 標準モジュール
#   vba-files/UserForm/*.frm → ユーザーフォーム
#
# 使い方:
#   1. このスクリプトをPowerShellで実行
#      > .\scripts\vba-import.ps1
#   2. vba-files/_vbe/ に変換済みSJISファイルが生成される
#   3. ExcelのVBE（Alt+F11）→ ファイルからインポートで _vbe/ 内のファイルを選択
#=============================================================================

$dstDir = Join-Path $PSScriptRoot "..\vba-files\_vbe"
$sjis   = [System.Text.Encoding]::GetEncoding("shift_jis")
$utf8   = [System.Text.Encoding]::UTF8

# 出力先フォルダを確保
if (-not (Test-Path $dstDir)) {
    New-Item -ItemType Directory -Path $dstDir | Out-Null
}

# 変換対象の定義（フォルダ → 拡張子）
$targets = @(
    @{ Dir = "Class";    Ext = "*.cls" },
    @{ Dir = "Module";   Ext = "*.bas" },
    @{ Dir = "UserForm"; Ext = "*.frm" }
)

$totalCount = 0

foreach ($target in $targets) {
    $srcDir = Join-Path $PSScriptRoot "..\vba-files\$($target.Dir)"

    if (-not (Test-Path $srcDir)) {
        Write-Host "[スキップ] $($target.Dir)/ フォルダが見つかりません" -ForegroundColor Yellow
        continue
    }

    $files = Get-ChildItem -Path $srcDir -Filter $target.Ext
    if ($files.Count -eq 0) { continue }

    foreach ($file in $files) {
        $content = [System.IO.File]::ReadAllText($file.FullName, $utf8)
        $dstPath = Join-Path $dstDir $file.Name
        [System.IO.File]::WriteAllText($dstPath, $content, $sjis)
        Write-Host "[変換完了] $($target.Dir)\$($file.Name)  →  _vbe\$($file.Name)"
        $totalCount++
    }
}

if ($totalCount -eq 0) {
    Write-Host "[警告] 変換対象ファイルが見つかりませんでした" -ForegroundColor Yellow
    exit 1
}

Write-Host ""
Write-Host "$totalCount 件のファイルをShift-JISに変換しました" -ForegroundColor Green
Write-Host "インポート先: $dstDir" -ForegroundColor Cyan
