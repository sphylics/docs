
# Google Apps Script ライブラリ関数仕様書

>[!IMPORTANT]
>- すべてのスプレッドシート書き込み系関数は内部で `protection(value)` を実行し、インジェクション対策を行っています。
>- このライブラリは`db.`で宣言します。

## 目次

1. [help()](#1-help)
2. [get_row()](#2-get_rowsheet_name-string-row_number-number-any--null)
3. [get_col()](#3-get_colsheet_name-string-col_number-number-any--null)
4. [get_cell()](#4-get_cellsheet_name-string-row-number-col-number-any--null)
5. [get_range()](#5-get_rangesheet_name-string-start_row-number-start_col-number-num_rows-number-num_cols-number-any--null)
6. [add_row()](#5-add_rowvalue-any-row-number-sheet_name-string-void)
7. [add_col()](#6-add_colvalue-any-col-number-sheet_name-string-void)
8. [set_value()](#7-set_valuevalue-any-row-number-col-number-sheet_name-string-void)
9. [set_range()](#8-set_rangevalue-2dlist-start_row-number-start_col-number-sheet_name-string)
10. [sc_props()](#9-sc_propskey-string-string--null)
11. [time()](#10-timetype-datestringnumber-string)
12. [random()](#11-randomlength-number-string)
13. [追加補足](#追加補足)

...existing code...

## 1. `help()`
- **目的**: ライブラリ関数一覧と概要をログに出力し、開発者への簡易ヘルプを提供。
- **引数**: なし
- **戻り値**: なし（`Logger.log`にて関数説明を出力）
- **エラー処理**: なし
- **備考**: 運用開始前の関数理解・レビューに有効。

## 2. `get_row(sheet_name: string, row_number: number): any[] | null`

- **目的**: 指定スプレッドシートの特定行の全セル値を一括取得。
- **パラメータ**
  - `sheet_name`: 対象シート名（厳密に一致する必要あり）
  - `row_number`: 1ベースの行番号
- **戻り値**
  - 成功時: 行内のセル値を格納した1次元配列（`any[]`）
  - 失敗時: `null`（例：シート未存在、範囲外など）
- **例外**: 内部で`try-catch`制御済み。例外発生時はログへエラーメッセージを記録し、`null`返却。
- **依存関係**: `sc_props("server")`で取得したスプレッドシートIDに依存。
- **用途例**: データの一行丸ごと読み取りや行ベースの処理前取得。

## 3. `get_col(sheet_name: string, col_number: number): any[] | null`

- **目的**: 指定シートの指定列を縦一列すべて取得。
- **パラメータ**
  - `sheet_name`: 対象シート名
  - `col_number`: 1ベースの列番号
- **戻り値**
  - 成功時: 縦方向のセル値配列（`any[]`）
  - 失敗時: `null`
- **例外**: 同上。
- **注意**: データの途中空白セルも配列要素として含まれる。
- **利用シーン**: 列単位での集計や検索に最適。

## 4. `get_cell(sheet_name: string, row: number, col: number): any | null`

- **目的**: 指定セルの単一値を取得。
- **パラメータ**
  - `sheet_name`: シート名
  - `row`: 行番号（1-based）
  - `col`: 列番号（1-based）
- **戻り値**
  - 成功時: セルの値（型不問）
  - 失敗時: `null`
- **例外処理**: 例外は内部で捕捉し、エラーログ出力。
- **利用例**: 単一セルの状態チェックやトリガー条件判定。


## 5. `get_range(sheet_name: string, start_row: number, start_col: number, num_rows: number, num_cols: number): any[][] | null`
**概要:**
指定したシート範囲を2次元配列で取得します。

- **引数:**
  - `sheet_name`: シート名
  - `start_row`: 開始行（1始まり）
  - `start_col`: 開始列（1始まり）
  - `num_rows`: 取得行数
  - `num_cols`: 取得列数
- **戻り値:**
  - 成功時: 範囲の2次元配列
  - 失敗時: `null`
- **例外処理:** try-catchで例外時はログ出力し、`null`返却
- **用途例:**
```javascript
const data = get_range("社員一覧", 2, 1, 3, 2);
// 例: [["田中",28], ["鈴木",34], ["佐藤",41]]
```

## 6. `add_row(value: any, row: number, sheet_name: string): void`
**概要:**
指定した行の先頭空セル（横方向）に値を追加します。空セルがなければ右端に追加します。

- **引数:**
  - `value`: 追加する値
  - `row`: 1始まりの行番号
  - `sheet_name`: 対象シート名
- **戻り値:** なし
- **例外処理:** try-catchで例外時はログ出力のみ
- **注意:** 行の右端に追加される場合、他列との整合性に注意
- **用途例:** 横方向の履歴や動的データの追記

## 7. `add_col(value: any, col: number, sheet_name: string): void`
**概要:**
指定した列の先頭空セル（縦方向）に値を追加します。空セルがなければ最終行の下に追加します。

- **引数:**
  - `value`: 追加する値
  - `col`: 1始まりの列番号
  - `sheet_name`: 対象シート名
- **戻り値:** なし
- **例外処理:** try-catchで例外時はログ出力のみ
- **注意:** 他列との行整合性が崩れる可能性あり
- **用途例:** 縦方向のログや履歴の追加

## 8. `set_value(value: any, row: number, col: number, sheet_name: string): void`
**概要:**
指定したセルに値を上書きします。

- **引数:**
  - `value`: セットする値
  - `row`: 行番号（1始まり）
  - `col`: 列番号（1始まり）
  - `sheet_name`: シート名
- **戻り値:** なし
- **例外処理:** try-catchで例外時はログ出力のみ
- **用途例:** 任意セルの更新や修正

## 9. `set_range(value: any[][], start_row: number, start_col: number, sheet_name: string): void`
**概要:**
指定したシート範囲に2次元配列の値を一括でセットします。全セルに `protection()` を適用し、インジェクション対策済みです。

- **引数:**
  - `value`: 2次元配列（例：`[[1,2],[3,4]]`）
  - `start_row`: 書き込み開始行（1始まり）
  - `start_col`: 書き込み開始列（1始まり）
  - `sheet_name`: 対象シート名
- **戻り値:** なし
- **例外処理:** try-catchで例外時はログ出力のみ
- **注意:**
  - `value` は必ず2次元配列で渡すこと
  - 開始位置から範囲にセットされるため既存データが上書きされる
  - セル単位で `protection()` を適用
- **使用例:**
```javascript
let data = [
  ["名前", "年齢", "部署"],
  ["田中", 28, "営業"],
  ["鈴木", 34, "開発"]
];
set_range(data, 3, 2, "社員一覧");
```

## 10. `sc_props(key: string): string | null`
**概要:**
スクリプトプロパティから値を取得します。

- **引数:**
  - `key`: 取得キー（例: `"server"`）
- **戻り値:**
  - 成功時: プロパティ値（現在は固定ID文字列）
  - 未設定キー: `null`
- **利用制限:** 現状 `"server"` のみ対応
- **用途例:** スプレッドシートIDの一元管理

## 11. `time(time: Date | string | number): string`
**概要:**
任意の日時を `YYYY/MM/DD HH:mm:ss` 形式の文字列に変換します。

- **引数:**
  - `time`: Dateオブジェクト、日時文字列、またはタイムスタンプ
- **戻り値:** フォーマット済み日時文字列
- **例外処理:** 無効な日時の場合は `Invalid Date` となるため、入力検証推奨
- **用途例:** ログ記録や帳票作成時の統一日時フォーマット

## 12. `random(length: number): string`
**概要:**
指定した文字数のランダムな英数字（+`&`）文字列を生成します。

- **引数:**
  - `length`: 出力する文字数（正の整数推奨）
- **戻り値:** ランダム文字列（例: `aB3&z9Xy0Q`）
- **注意:** 乱数は `Math.random()` ベースで暗号学的強度はありません。セキュリティ用途には不向きです。
- **用途例:** 一時的な識別子や簡易トークン生成

## 補足事項

- **依存スプレッドシートID管理:** すべての関数は `sc_props("server")` で取得した固定IDのスプレッドシートにアクセスします。
- **例外の一元管理:** すべての関数で try-catch によりエラーはロギングし、`null` または void を返します。
- **パフォーマンス注意:** 大規模シートへの連続呼び出しはAPI制限や遅延リスクがあるため、バッチ処理を推奨します。
- **拡張性:** 入力チェックは限定的です。利用者側でパラメータバリデーションを行ってください。
