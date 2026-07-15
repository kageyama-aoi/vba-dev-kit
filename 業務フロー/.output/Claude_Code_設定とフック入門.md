# Claude Code 設定とフック入門

> 今回作った「Bashコマンドを自動ログする仕組み」を題材に、
> 関連する技術をまとめた資料です。

---

## 1. 設定ファイル（settings.json）

### ファイルの場所と優先順位

Claude Code は複数の設定ファイルを読み込み、後のものが前のものを上書きします。

| ファイル | スコープ | Git管理 | 用途 |
|---|---|---|---|
| `~/.claude/settings.json` | 全プロジェクト共通 | なし | 個人の好み・デフォルト設定 |
| `.claude/settings.json` | このプロジェクト専用 | **あり** | チーム共有のルール |
| `.claude/settings.local.json` | このプロジェクト専用 | なし（.gitignore推奨） | 個人のローカルオーバーライド |

今回作ったのは **プロジェクト専用の `.claude/settings.json`** です。

---

## 2. パーミッション設定（permissions）

Claude Code がツールを使うとき、ユーザーに確認を求めるかどうかを制御します。

```json
{
  "permissions": {
    "defaultMode": "acceptEdits",
    "allow": [
      "Bash(git *)",
      "Bash(gh *)"
    ]
  }
}
```

### defaultMode の選択肢

| 値 | 動作 |
|---|---|
| `"default"` | すべて確認あり（デフォルト） |
| `"acceptEdits"` | ファイル操作（Read/Write/Edit/Glob/Grep）は自動承認 |
| `"bypassPermissions"` | **すべて確認なし**（完全自動化） |
| `"plan"` | プランモードで起動 |

### allow リストの書き方

```
"Bash(git *)"         → git で始まるコマンド全般
"Bash(gh issue *)"    → gh issue で始まるコマンドのみ
"Bash(npm run test)"  → この1コマンドだけ
"Read"                → Read ツール全般
"Edit(/path/to/*)"    → 特定パス配下のファイル編集
```

`*` はワイルドカード（任意の文字列にマッチ）。

---

## 3. フック（hooks）

**フック**とは、Claude Code のツール実行の前後に自動で走るコマンドのことです。
「Claude が何かをしたとき、こっちも自動で動かす」という仕組みです。

### フックが発火するタイミング（主要なもの）

| イベント | タイミング |
|---|---|
| `PreToolUse` | ツール実行の**直前** |
| `PostToolUse` | ツール実行の**直後（成功時）** |
| `Stop` | Claude が応答を終えたとき |
| `SessionStart` | セッション開始時 |

### フックの構造

```json
{
  "hooks": {
    "PreToolUse": [
      {
        "matcher": "Bash",
        "hooks": [
          {
            "type": "command",
            "command": "ここにシェルコマンドを書く"
          }
        ]
      }
    ]
  }
}
```

| キー | 意味 |
|---|---|
| `matcher` | どのツールに反応するか（`"Bash"` / `"Write"` / `"Edit"` など） |
| `type` | フックの種類（`"command"` = シェルコマンドを実行） |
| `command` | 実際に実行するシェルコマンド |

---

## 4. フックが受け取るデータ（stdin JSON）

フックのコマンドは、**標準入力（stdin）にJSON形式で情報を受け取ります**。

### PreToolUse の場合

```json
{
  "session_id": "abc123",
  "tool_name": "Bash",
  "tool_input": {
    "command": "git status"
  }
}
```

`tool_input.command` に、Claude が実行しようとしているコマンドが入っています。
これを取り出してログに記録するのが今回の仕組みです。

---

## 5. 今回作ったログフック

```json
{
  "type": "command",
  "command": "python -c \"import sys,json;from datetime import datetime;d=json.load(sys.stdin);cmd=d.get('tool_input',{}).get('command','');open('C:/Users/.../bash-log.txt','a',encoding='utf-8').write('['+datetime.now().strftime('%Y-%m-%d %H:%M:%S')+'] '+cmd+'\\n')\" 2>/dev/null || true"
}
```

### やっていること（1行を分解すると）

```python
import sys, json
from datetime import datetime

# stdin から JSON を読み込む
d = json.load(sys.stdin)

# tool_input.command を取り出す
cmd = d.get('tool_input', {}).get('command', '')

# タイムスタンプ付きでログファイルに追記
with open('bash-log.txt', 'a', encoding='utf-8') as f:
    f.write('[' + datetime.now().strftime('%Y-%m-%d %H:%M:%S') + '] ' + cmd + '\n')
```

### `2>/dev/null || true` の意味

| 部分 | 意味 |
|---|---|
| `2>/dev/null` | エラーメッセージを捨てる（画面に出さない） |
| `\|\| true` | フックが失敗しても Claude の処理を止めない |

フックが失敗するとClaude全体がブロックされる可能性があるため、
ログ記録のような「補助的な処理」は必ず `|| true` をつけるのがベストプラクティスです。

---

## 6. 全体の流れ（今回の設定）

```
ユーザーが Claude に指示
        ↓
Claude が Bash コマンドを実行しようとする
        ↓
【PreToolUse フック発火】
  Python が stdin から JSON を受け取る
  → コマンド文字列を bash-log.txt に追記
        ↓
パーミッション確認
  → allow リストに一致 → 自動承認
  → defaultMode: acceptEdits → ファイル操作も自動承認
        ↓
Bash コマンドが実際に実行される
```

---

## 7. 応用例

### ファイル編集後に自動フォーマット（PostToolUse）

```json
{
  "hooks": {
    "PostToolUse": [{
      "matcher": "Write|Edit",
      "hooks": [{
        "type": "command",
        "command": "python -m black \"$(jq -r '.tool_input.file_path')\" 2>/dev/null || true"
      }]
    }]
  }
}
```

### 危険なコマンドをブロック（PreToolUse + JSON返却）

フックが `{"continue": false, "stopReason": "理由"}` を返すと、そのツール呼び出しをブロックできます。

```json
{
  "hooks": {
    "PreToolUse": [{
      "matcher": "Bash",
      "hooks": [{
        "type": "command",
        "command": "python -c \"import sys,json;d=json.load(sys.stdin);cmd=d['tool_input']['command'];print(json.dumps({'continue': False, 'stopReason': '要確認'}) if 'rm -rf' in cmd else '{}')\""
      }]
    }]
  }
}
```

---

## 8. まとめ

| 技術 | 役割 |
|---|---|
| `settings.json` | Claude Code の動作ルールを定義するファイル |
| `permissions.defaultMode` | ファイル操作の自動承認レベルを設定 |
| `permissions.allow` | 特定のBashコマンドを確認なしで実行 |
| `hooks.PreToolUse` | ツール実行前に自動で走るコマンドを登録 |
| stdin JSON | フックがツールの情報を受け取る仕組み |
| Python 1行スクリプト | JSONを読んでログに書き出す処理 |

この仕組みを応用すれば、**自動テスト・自動フォーマット・監査ログ・危険コマンドのブロック**など、
様々な自動化が `.claude/settings.json` 1ファイルで実現できます。
