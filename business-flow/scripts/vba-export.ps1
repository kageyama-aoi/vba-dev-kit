#=============================================================================
# vba-export.ps1
# 用途 : VBEからエクスポートしたShift-JISのVBAファイルをUTF-8に変換して
#        対応フォルダに上書きする（gitコミット前に実行）
#
# 対応ファイル:
#   _vbe/*.cls → vba-files/Class/   （クラスモジュール）
#   _vbe/*.bas → vba-files/Module/  （標準モジュール）
#   _vbe/*.frm → vba-files/UserForm/（ユーザーフォーム）
#
# 使い方:
#   1. ExcelのVBE（Alt+F11）→ 各モジュールを右クリック → ファイルのエクスポート
#      → 保存先を vba-files/_vbe/ に指定
#   2. このスクリプトをPowerShellで実行
#      > .\scripts\vba-export.ps1
#   3. 各フォルダがUTF-8で上書きされる
#   4. git add / git commit でコミット
#=============================================================================

$srcDir = Join-Path $PSScriptRoot "..\vba-files\_vbe"
$sjis   = [System.Text.Encoding]::GetEncoding("shift_jis")
$utf8nb = New-Object System.Text.UTF8Encoding $false  # BOMなしUTF-8

if (-not (Test-Path $srcDir)) {
    Write-Host "[エラー] $srcDir が見つかりません" -ForegroundColor Red
    Write-Host "VBEからのエクスポート先を vba-files/_vbe/ にしてください"
    exit 1
}

# 拡張子 → 出力先フォルダの対応
$extMap = @{
    ".cls" = "Class"
    ".bas" = "Module"
    ".frm" = "UserForm"
}

$allFiles = Get-ChildItem -Path $srcDir -File | Where-Object { $extMap.ContainsKey($_.Extension) }

if ($allFiles.Count -eq 0) {
    Write-Host "[警告] $srcDir に対象ファイル（.cls/.bas/.frm）が見つかりません" -ForegroundColor Yellow
    exit 1
}

$totalCount = 0

foreach ($file in $allFiles) {
    $subDir = $extMap[$file.Extension]
    $dstDir = Join-Path $PSScriptRoot "..\vba-files\$subDir"

    if (-not (Test-Path $dstDir)) {
        New-Item -ItemType Directory -Path $dstDir | Out-Null
    }

    $content = [System.IO.File]::ReadAllText($file.FullName, $sjis)
    $dstPath = Join-Path $dstDir $file.Name
    [System.IO.File]::WriteAllText($dstPath, $content, $utf8nb)
    Write-Host "[変換完了] _vbe\$($file.Name)  →  $subDir\$($file.Name)"
    $totalCount++
}

Write-Host ""
Write-Host "$totalCount 件のファイルをUTF-8に変換しました" -ForegroundColor Green
Write-Host "次のステップ: git add vba-files/ してコミットしてください" -ForegroundColor Cyan
