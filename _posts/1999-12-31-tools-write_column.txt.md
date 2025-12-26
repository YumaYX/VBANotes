---
layout: post
category: tools
title: "write_column.txt"
---

<small>この文章はAIで生成しています。誤りが含まれることがあります。</small>

このプログラムは、選択した列のデータをテキストファイルに書き出すVBAコードです。以下にその詳細な説明を日本語でMarkdown形式で記載します。

## プログラムの説明

### 機能
- 選択した列のデータをテキストファイルに書き出す機能を提供します。
- 書き出すファイルの名前は、元のシート名と現在の時刻を基に生成されます。

### 構成
- **列ファイル書き出し(W)**: `OutputColumn` という名前のサブルーチンです。

### 詳細な説明

1. **変数の宣言**:
   - `column`: アクティブセルの列番号を格納します。
   - `ws`: アクティブシートを格納します。
   - `max_row`: アクティブセルの列の最終行番号を格納します。
   - `output_filename`: 出力ファイルの名前を格納します。
   - `cell_val`: セルの値を格納します。

2. **ファイル名の生成**:
   - `output_filename` は、シート名と現在の時刻を基に生成されます。
   - 空白文字は削除されます。

3. **ファイルの書き出し**:
   - `output_filename` で指定された名前のテキストファイルを開きます。
   - アクティブセルの列の各セルの値を読み込み、空白でない場合はファイルに書き込みます。
   - ファイルを閉じます。

### 使用方法
- このコードは、ExcelのVBA環境で実行されます。
- 選択した列のデータをテキストファイルに書き出すことができます。

### 注意点
- ファイル名の生成方法や書き込みの処理は、特定の状況下で変更が必要な場合があります。
- エラー処理は追加することで、より堅牢なコードにすることができます。

```vb
' 列ファイル書き出し(W)
Sub OutputColumn()
  Dim column As Long: column = ActiveCell.column
  Dim ws As Worksheet: Set ws = ActiveSheet
  Dim max_row As Long: max_row = ws.Cells(Rows.Count, column).End(xlUp).Row

  Dim output_filename As String
  output_filename = ws.Name + "_" + Format(Time, "hhmmss") + ".txt"
  output_filename = Replace(output_filename, " ", "")

  Open output_filename For Output As #1
  Dim cell_val As String
  For i = 1 To max_row
    cell_val = ws.Cells(i, column).Value
    If Len(cell_val) <> 0 Then
      Print #1, cell_val
    End If
  Next
  Close #1
End Sub

```
