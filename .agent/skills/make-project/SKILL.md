---
name: make-project
description: 新規プロジェクトの初期構築を行うスキル。新規プロジェクトのセットアップ、初期ファイル作成、Git初期化、既存プロジェクトのアップデートを行う際に使用する。「プロジェクトを初期化して」「新規プロジェクトをセットアップして」「プロジェクトの初期構築をして」「プロジェクトをアップデートして」などのリクエストに対して必ず使用すること。
---

# make-project スキル

## ステップ1：モード選択

作業を開始する前に、ユーザーに必ず実行モードを選択してもらうこと。
選択式UI（Claude Code の AskUserQuestion 等）が使えるツールではそれを使い、
使えないツールでは以下を口頭で提示して番号で選択してもらう：

「make-project スキルを開始します。以下から実行モードを選択してください：

1. 新規プロジェクトを構成する（親フォルダ上の実行のみ有効）
2. 既存のプロジェクトをアップデートする（プロジェクトフォルダ内での実行のみ有効）
3. その他・相談する

番号を入力してください。」

- **1を選択** → 「モードA：新規プロジェクト作成」へ進む
- **2を選択** → 「モードB：既存プロジェクトのアップデート」へ進む
- **3を選択** → 「モードC：対話」へ進む

---

# モードA：新規プロジェクト作成

## A-1. 事前確認
以下の情報をユーザーに確認してから作業を開始する：
- プロジェクト名（フォルダ名）：[USER_INPUT]
- GitHubリポジトリURL：[USER_INPUT]
- デフォルトブランチ名：main（変更がある場合はユーザーに確認）
- 使用するVBAモジュール名（.cls / .bas ファイル名）：[USER_INPUT]（未定の場合は「未定」でOK）

※ 事前にGitHub上でリポジトリを作成しておくこと（README等は追加しない）

## A-2. プロジェクトフォルダの作成

\```bash
mkdir [PROJECT_NAME]
cd [PROJECT_NAME]
\```

以降の作業はすべてこのフォルダ内で実行する。

## A-3. サブフォルダ作成

\```bash
mkdir -p .agent/memory
mkdir -p .agent/handoff
mkdir -p .agent/workflows
mkdir -p .claude/commands
mkdir -p .spec
mkdir -p .output
mkdir -p .references
\```

※ `.agent/skills/` は親フォルダ（マクロ開発環境）で一元管理するため、各プロジェクトには作成しない。

## A-4. 初期ファイル作成

### README.md

README.mdが存在しない、あるいは中身が空の時のみ以下を実行する。
プロジェクト名はA-1で確認済みのものを使用する。日時はローカル時刻、ツール名は現在使用中のツール名＋モデル名（例：Claude Code + 使用モデル名）を記載する。

\```bash
cat << 'EOF' > README.md
# Project: [PROJECT_NAME]

* これは[YYYY-MM-DD HH:MM]に自動生成されたプロジェクトである
* 初期構築担当ツール名：[TOOL_NAME]
* このプロジェクトでは、生成AIおよびスキルを積極的に活用して開発する
EOF
\```

### .agent/memory/MEMORY.md
\```bash
cat << 'EOF' > .agent/memory/MEMORY.md
# MEMORY

## プロジェクト概要

## 学習した知識・教訓
EOF
\```

### .agent/handoff/HANDOFF.md
\```bash
cat << 'EOF' > .agent/handoff/HANDOFF.md
# HANDOFF

初回セットアップ完了。作業を開始してください。
EOF
\```

### CLAUDE.md（プロジェクトルート）
\```bash
cat << 'EOF' > CLAUDE.md
# [PROJECT_NAME] プロジェクト固有設定

- セッション開始時に AGENTS.md を読み込み、読み込んだことを報告すること
- 親フォルダのVBAナレッジ（docs/VBA注意事項.md / docs/VBAツール標準テンプレート.md / docs/関数命名ルール.md）が自動適用される

## このプロジェクトのモジュール
- `[MODULE_NAME]` … [MODULE_DESCRIPTION]（A-1で確認したモジュール名を記載。未定の場合は後で追記）

## プロジェクト固有の注意
- （プロジェクト特有の制約があればここに追記）
EOF
\```

### AGENTS.md（プロジェクトルート）
\```bash
cat << 'EOF' > AGENTS.md
# Project guide line

## プロジェクトの原則
- 本プロジェクトのプラン作成、および回答は全て日本語で行う

## プロジェクトの目的
- 

# Memory & Handoff Instructions

## 3ファイルの役割と哲学
- 本ファイル（AGENTS.md）は「厳格なルール」、人が作成
- MEMORY.mdは「積み上がる経験」、AIが作成・AIが利用
- HANDOFF.mdは「セッション間の引き継ぎ」、AIが作成・AIが利用、ただし人間がレビューし必要な情報をキュレーションする

## セッション開始時（必須）
セッション開始時、ユーザーへの最初の応答の前に、以下の2ファイルを読み込み、読み込んだことを報告すること：
- `.agent/memory/MEMORY.md`  （学習した知識・教訓）
- `.agent/handoff/HANDOFF.md` （前回の作業引き継ぎ）

## メモリ管理
- 新しい知識・教訓を記録する際は `.agent/memory/MEMORY.md` を更新
- 既存のMEMORY.mdを更新する前に、現在のファイルを`.agent/memory/YYYY-MM-DD.md` にアーカイブしてから新規作成
- ローカルの自動メモリ機能（~/.claude/ 配下）は使用しない
- MEMORY.mdは200行以内を維持すること
- 本ファイルと重複する内容はMEMORY.mdに書かない

## ハンドオフ管理
- ハンドオフは `/handoff` コマンドで作成（Claude Codeの場合）
- 保存先は `.agent/handoff/HANDOFF.md`（固定名）
- 作成時は既存ファイルを `.agent/handoff/YYYY-MM-DD-HHMM.md` にリネームしてからHANDOFF.mdを新規作成する
- 時刻はローカル時刻・24時間表記

## 仕様駆動開発（SDD）ルール
- コーディングや業務作業を開始する前に、必ず `.spec/` 配下の4ファイルを確認・更新すること
- 作業の順序：PLAN（目的確認）→ SPEC（要件確認）→ TODO（タスク確認）→ 実作業
- **PLAN.mdは人間の口頭メモ・自由記述**であり、箇条書き・口語・断片的な内容で構わない
- PLAN.mdを読んだら、そのまま実装に入らず、不明点をヒアリングしながらSPEC.mdを作成・確定させること
- SPEC.mdが確定してからTODO.mdのタスク分解を行い、ユーザーの承認を得てから実作業を開始する
- 作業完了後は TODO.md の該当タスクにチェックを入れ、KNOWLEDGE.md に学びを記録する
- 仕様が不明確な場合は作業を開始せず、ユーザーに確認してから SPEC.md を更新する
- 新しい開発サイクルを始める際は `/newplan` コマンドを使用する

## フォルダ用途
- `.spec/`：設計ドキュメント（PLAN / SPEC / TODO / KNOWLEDGE）
- `.output/`：成果物・アウトプット（記事MD、コード、資料など完成したもの）
- `.references/`：参考資料・素材（PDFや画像、URLメモ、サンプルコードなど作業の入力素材）
EOF
\```

## A-5. 仕様駆動開発ファイルの作成（.spec/）

### .spec/PLAN.md
\```bash
cat << 'EOF' > .spec/PLAN.md
# PLAN - やりたいこと

<!-- ここに思ったことを自由に書いてください。箇条書きでも口語でもOK -->
<!-- Claude がこの内容を読んでヒアリングし、SPEC.md を作成します -->
EOF
\```

### .spec/SPEC.md
\```bash
cat << 'EOF' > .spec/SPEC.md
# SPEC - 技術仕様・要件定義

## 機能要件
## 非機能要件
## 技術構成
EOF
\```

### .spec/TODO.md
\```bash
cat << 'EOF' > .spec/TODO.md
# TODO - タスクリスト

## 優先度：高
## 優先度：中
## 優先度：低
## 完了済み
- [x] 初期セットアップ
EOF
\```

### .spec/KNOWLEDGE.md
\```bash
cat << 'EOF' > .spec/KNOWLEDGE.md
# KNOWLEDGE - ドメイン知識・調査結果

## 業務・ドメイン知識
## 調査・リサーチ結果
## 技術的な知見
## 決定事項と理由
EOF
\```

## A-6. newplan / handoff コマンドの作成（テンプレートからコピー）

コマンドの本文は本スキルの `templates/` フォルダで一元管理している。
**内容をこのファイルや作業中に書き起こさず、必ずテンプレートファイルからコピーすること**
（内容を変えたい場合は `templates/` 側を修正する。ここに本文を書くと二重管理になる）。

プロジェクトフォルダは親フォルダ（マクロ開発環境）直下にあるため、
プロジェクトフォルダ内からは `../.agent/skills/make-project/templates/` で参照できる。

\```bash
cp ../.agent/skills/make-project/templates/newplan.md .claude/commands/newplan.md
cp ../.agent/skills/make-project/templates/newplan.md .agent/workflows/newplan.md
cp ../.agent/skills/make-project/templates/handoff.md .claude/commands/handoff.md
cp ../.agent/skills/make-project/templates/handoff.md .agent/workflows/handoff.md
\```

テンプレートが見つからない場合はパスを確認し、勝手に本文を創作しないこと。

## A-8. Git初期化

### .gitignore の作成
\```bash
cat << 'EOF' > .gitignore
# Logs
logs
*.log

node_modules
dist
dist-ssr
*.local

# Editor directories and files
.vscode/*
!.vscode/extensions.json
.idea
.DS_Store
.env
EOF
\```

### Git初期化とpush
\```bash
git init
git add .
git commit -m "first commit"
git remote add origin [USER_INPUT]
git push -u origin main
\```

## A-9. 完了報告

全手順完了後、以下を報告する：
- 作成したファイル・フォルダの一覧
- GitHubへのpush結果
- 次のステップの案内（「AGENTS.mdにプロジェクト概要を記載し、PLAN.mdにやりたいことを書いてください」など）

---

# モードB：既存プロジェクトのアップデート

## B-1. 現状の精査

本スキル（make-project）のモードAのA-3以降に記載されているすべての要素を正として、
現在のプロジェクトフォルダの状態と照合し、不足・未作成の要素をリストアップする。

精査完了後、以下を報告する：
「以下の差分が見つかりました。アップデートを適用してよいですか？
- 追加・作成するもの：[不足している要素の一覧]
- スキップするもの（既存）：[すでに存在する要素の一覧]」

ユーザーの承認を得てから B-2 に進む。

## B-2. 差分の適用

B-1で不足と判定された要素のみ、モードAの対応する手順を実行する。
既存ファイル・フォルダは上書きしない。
AGENTS.mdへの追記は既存内容と重複しないよう確認してから行う。

## B-3. 完了報告

適用した内容とスキップした内容の一覧を報告する。

---

# モードC：対話

ユーザーの相談内容をヒアリングし、このスキルの範囲でできることを提案する。
必要に応じてモードAまたはモードBへ誘導する。

