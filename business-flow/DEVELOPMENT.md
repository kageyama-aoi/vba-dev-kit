# Project: VBA フローチャートマクロ

* これは 2026-03-19 に自動生成されたプロジェクトである
* 初期構築担当ツール名：Claude Code Sonnet 4.6
* このプロジェクトでは、生成AIおよびスキルを積極的に活用して開発する

---

## ディレクトリ構成

```
vba-files/
  Class/          # git管理（UTF-8）Claude が編集・レビューする場所
  _vbe/           # gitignore（Shift-JIS）VBEとのやりとり用作業フォルダ
scripts/
  vba-import.ps1  # UTF-8 → Shift-JIS 変換（VBEに入れる前に実行）
  vba-export.ps1  # Shift-JIS → UTF-8 変換（VBEから出した後に実行）
```

---

## エンコーディングについて

VBA（Excel VBE）はファイルを **Shift-JIS** で扱いますが、
このリポジトリは **UTF-8** で管理しています。
`scripts/` の変換スクリプトを使ってギャップを吸収します。

---

## ワークフロー

### Claude が編集した VBA を VBE に反映する

```powershell
# 1. UTF-8 → Shift-JIS に変換
.\scripts\vba-import.ps1

# 2. VBE（Alt+F11）→ ファイルからインポート
#    vba-files\_vbe\ 内の .cls を選択してインポート
```

### VBE で編集した VBA を git に反映する

```powershell
# 1. VBE → ファイルのエクスポート
#    保存先: vba-files\_vbe\

# 2. Shift-JIS → UTF-8 に変換
.\scripts\vba-export.ps1

# 3. コミット
git add vba-files/Class/
git commit -m "..."
```
