# Mermaid 構造図 - Sql_Table_Formatter.cls
> 生成日: 2026-04-14
> 対象: mysql-table-definition-to-excel/src/Sql_Table_Formatter.cls

---

## 【図1】公開マクロ グループ図

> README や概要説明に貼る。全体像を一目で把握するための図。

```mermaid
flowchart TD
  subgraph PublicMacro ["■ 公開マクロ（Excelから直接実行）"]
    A["CreateTableDefinitionSheet\nSQLファイル → テーブル定義シート 自動生成"]
  end

  subgraph Phase1 ["フェーズ 1：ファイル入力"]
    B["SQLファイルをダイアログで選択\nUTF-8 / UTF-8 BOM / Shift-JIS に対応"]
  end

  subgraph Phase2 ["フェーズ 2：SQL解析（2パス）"]
    C["パス1 ─ CREATE TABLE ブロック\n　カラム名・型・NULL・DEFAULT を収集"]
    D["パス2 ─ ALTER TABLE ブロック\n　PK / UNIQUE / INDEX 情報を各カラムに反映"]
  end

  subgraph Phase3 ["フェーズ 3：Excel出力"]
    E["出力シートを削除して新規作成"]
    F["テーブル定義を2D配列で一括書き込み\n（スタイル・交互色・列幅自動調整）"]
  end

  A --> Phase1 --> Phase2 --> Phase3
```

**この図の読み方**
- このモジュールの公開マクロは `CreateTableDefinitionSheet` の1つのみです。
- 大きく「入力 → 解析 → 出力」の3フェーズで処理が進みます。
- 解析は2パス構成で、まずカラム定義、次にキー情報を取得します。

---

## 【図2】全体詳細図

> 技術解説・コードリーディング時の地図。内部構造まで把握するための図。

```mermaid
flowchart TD
  subgraph Public ["■ 公開マクロ"]
    MAIN["CreateTableDefinitionSheet"]
  end

  subgraph HelperInput ["▼ 内部ヘルパー：ファイル入力系"]
    SEL["SelectSqlFile\nファイル選択ダイアログ → パスを返す"]
    RFL["ReadFileLines\nADODB.Streamで UTF-8 読み込み\n失敗時は標準I/Oにフォールバック\nBOM自動除去・改行コード統一"]
  end

  subgraph HelperParse ["▼ 内部ヘルパー：SQL解析系"]
    PSF["ParseSqlFile\n2パスで全行を走査し\n並列配列にカラム情報を格納"]
    PCL["ParseColumnLine\nカラム定義行をトークン分解\n→ 名前・型・NULL・DEFAULT を抽出"]
    EBN["ExtractBacktickName\nバッククォート囲みの名前を取り出す"]
    EPC["ExtractParenCols\n括弧内のカラム名を配列で返す"]
    AKI["ApplyKeyInfo\n対象カラムの colKeys を更新\n（PK / UNIQUE / INDEX）"]
  end

  subgraph HelperOutput ["▼ 内部ヘルパー：Excel出力系"]
    WDF["WriteDefinitions\nテーブル単位でループし\n2D配列を一括書き込み"]
    ROS["RecreateOutputSheet\n既存シートを削除して新規作成"]
    ATS["ApplyTableHeaderStyle\nテーブル名行：濃青・白文字・太字"]
    ACS["ApplyColHeaderStyle\nカラムヘッダー行：水色・太字・下線"]
  end

  subgraph HelperUtil ["▼ 内部ヘルパー：ユーティリティ"]
    CUT["CountUniqueTables\nDictionaryで重複除去し\nテーブル数を返す"]
  end

  %% CreateTableDefinitionSheet の呼び出し関係
  MAIN -->|"① ファイル選択"| SEL
  MAIN -->|"② SQL解析"| PSF
  MAIN -->|"③ シート再作成"| ROS
  MAIN -->|"④ シート書き込み"| WDF
  MAIN -->|"⑤ 完了メッセージ"| CUT

  %% ParseSqlFile の内部呼び出し
  PSF -->|"ファイル読み込み"| RFL
  PSF -->|"パス1: CREATE TABLE"| EBN
  PSF -->|"パス1: カラム行解析"| PCL
  PSF -->|"パス2: ALTER TABLE"| EBN
  PSF -->|"パス2: キーカラム抽出"| EPC
  PSF -->|"パス2: キー情報反映"| AKI

  %% WriteDefinitions の内部呼び出し
  WDF -->|"テーブル数確認"| CUT
  WDF -->|"テーブル名行スタイル"| ATS
  WDF -->|"ヘッダー行スタイル"| ACS
```

**主要な処理フロー**

1. **エントリポイント** `CreateTableDefinitionSheet` が Application の描画・計算・イベントを停止し、高速処理モードに切り替える。
2. **ファイル読み込み** `ReadFileLines` は ADODB.Stream で UTF-8 読み込みを試み、失敗時は標準VBA I/Oにフォールバックする。BOM除去・改行コード統一も行う。
3. **2パス解析** `ParseSqlFile` はファイル全行を2周する。1周目で CREATE TABLE からカラム属性を収集、2周目で ALTER TABLE から PK / INDEX を対応カラムに追記する。
4. **一括書き込み** `WriteDefinitions` はカラムデータを 2D Variant 配列に詰めてから `Range.Value = array` で1回書き込む（セル単体へのCOM呼び出しを大幅削減）。
5. **後処理** エラー有無にかかわらず `Cleanup` ラベルで ScreenUpdating / Calculation / EnableEvents を必ず復元する。
