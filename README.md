# マクロ開発環境

VBA（Excel）開発のための共有ナレッジ・スキル管理リポジトリ。

個別のVBAツールは各プロジェクトフォルダで独立したリポジトリとして管理し、
このリポジトリでは全プロジェクト共通のナレッジとAIスキルを一元管理する。

---

## ディレクトリ構成

```
マクロ開発環境/
├── CLAUDE.md                   # AIへの指示・スキル参照ルール
├── docs/                       # VBA開発ナレッジ
│   ├── VBA注意事項.md           # 実装ミス防止チェックリスト（常時参照）
│   ├── VBAツール標準テンプレート.md
│   ├── 関数命名ルール.md
│   ├── 雑多指示を整えるプロンプト.md
│   └── archive/                # ナレッジの更新履歴
└── .agent/skills/              # AIスキル定義
    ├── make-project/           # 新規プロジェクト作成
    ├── vba-spec/               # 要件整理
    ├── vba-new-module/         # 新規モジュール作成
    ├── vba-review/             # コードレビュー
    ├── vba-implement/          # 仕様〜.clsファイルまで一気通貫
    └── vba-knowledge-update/   # ナレッジ更新・アーカイブ
```

---

## 各プロジェクト（別リポジトリ）

| プロジェクト | リポジトリ | 概要 |
|------------|-----------|------|
| excel-image-importer | https://github.com/kageyama-aoi/excel-image-importer | 画像一括貼り付けツール |
| mysql-table-definition-to-excel | https://github.com/kageyama-aoi/vba-mysql-table-definition-to-excel | SQLからテーブル定義書を生成 |

> 各プロジェクトフォルダはこのリポジトリの管理対象外。

---

## 使い方

### スキルの発動方法

AIへの指示文にスキル名を含めるだけで自動的に発動する。
特別なコマンドは不要。

| やりたいこと | AIへの指示例 |
|------------|------------|
| 新規プロジェクトを作る | 「make-project スキルを使って新しいプロジェクトを作りたい」 |
| 要件を整理したい | 「vba-spec スキルで要件を整理して」 |
| モジュールをゼロから作る | 「vba-new-module スキルでモジュールを作って」 |
| 一気通貫で実装したい | 「vba-implement スキルでツールを作って」 |
| コードをレビューしてほしい | 「vba-review スキルでレビューして」 |
| ナレッジに知見を追加したい | 「vba-knowledge-update スキルでこれを追記して」 |

---

### 新規プロジェクトを立ち上げる

**事前準備（人間がやること）**

1. GitHubで新しいリポジトリを作成する
   - READMEなし・.gitignoreなし・ライセンスなしで作成すること
2. このフォルダ（マクロ開発環境）をClaude Codeで開く

**AIへの指示**

```
make-project スキルを使って新しいプロジェクトを作りたい
```

AIが以下を確認してくるので答える：
- プロジェクト名（フォルダ名）
- GitHubリポジトリURL
- 使用するVBAモジュール名（.cls / .bas）

あとはAIが自動で以下を行う：
- フォルダ・ファイル構成の作成
- CLAUDE.md / AGENTS.md の生成
- git init → first commit → GitHub push

**プロジェクト追加後にやること**

- `.gitignore` に新プロジェクトのフォルダ名を追記してコミット
- `README.md` のプロジェクト一覧を更新してコミット

---

### 既存プロジェクトで修正・開発をする

**手順**

1. 対象プロジェクトのフォルダをClaude Codeで開く
   （例：`excel-image-importer/` をワークスペースとして開く）
2. AIが自動でCLAUDE.md・AGENTS.md・MEMORY.md・HANDOFF.mdを読み込む
3. やりたいことを伝える

**修正の流れ（例）**

```
# 小さな修正の場合
「〇〇の処理がバグっているので直して」
→ AIが該当コードを読んで修正・vba-reviewで自己チェック

# 機能追加の場合
「vba-implement スキルで〇〇機能を追加して」
→ 要件整理 → 実装 → レビューまで一気通貫

# 仕様が曖昧な場合
「vba-spec スキルでまず要件を整理して」
→ ヒアリング → SPEC.md に書き込み → 確認後に実装へ
```

**セッション終了時**

```
/handoff
```

と入力すると、次回セッション用の引き継ぎファイルを自動作成する。

---

### ナレッジを更新する

開発中に「この注意事項をメモしておきたい」と思ったとき：

```
vba-knowledge-update スキルで、〇〇という注意事項をVBA注意事項.mdに追記して
```

AIが以下を自動で行う：
- 現在のナレッジファイルを `docs/archive/` にバックアップ
- 既存項目との重複チェック
- 優先度順を考慮した挿入位置を提案
- 確認後に反映

---

## スキル一覧

| スキル | SKILL.md | 役割 |
|-------|----------|------|
| make-project | `.agent/skills/make-project/` | 新規プロジェクト作成 |
| vba-spec | `.agent/skills/vba-spec/` | 要件整理・SPEC.md生成 |
| vba-new-module | `.agent/skills/vba-new-module/` | 新規モジュール作成 |
| vba-implement | `.agent/skills/vba-implement/` | 仕様〜.clsまで一気通貫 |
| vba-review | `.agent/skills/vba-review/` | チェックリスト駆動レビュー |
| vba-knowledge-update | `.agent/skills/vba-knowledge-update/` | ナレッジ追記・バックアップ |
