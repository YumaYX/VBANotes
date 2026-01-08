---
layout: post
category: tools
title: "write_column.txt"
---

<small>この文章はAIで生成しています。誤りが含まれることがあります。</small>

このプログラムは、選択した列のデータをテキストファイルに書き出すVBAコードです。以下に日本語で詳細な説明をMarkdown形式で記載します。

## VBAコードの説明

### プログラムの目的
選択したExcelの列のデータをテキストファイルに書き出す機能を提供します。

### 主な処理の流れ
1. **選択した列の情報を取得**: 現在アクティブなセルの列番号を取得し、その列のデータを処理対象とする。
2. **出力ファイル名の生成**: 現在の時刻を基にファイル名を生成し、選択したシート名と結合して出力ファイル名とする。
3. **データの書き出し**: 選択した列のデータを1行ずつ読み込み、空でないデータのみをテキストファイルに書き出す。

### 主要変数の説明
- `column`: 現在選択されている列の番号を保持する変数。
- `ws`: 現在アクティブなシートを保持するオブジェクト。
- `max_row`: 選択した列の最終行番号を保持する変数。
- `output_filename`: 出力するテキストファイルのファイル名を保持する変数。
- `cell_val`: 選択した列のセルの値を保持する変数。

### 注意点
- ファイル名の生成時に空白文字は削除されます。
- テキストファイルは現在のフォルダに保存されます。必要に応じてファイルパスを変更してください。

### 使用方法
- VBAエディタでこのコードをコピーし、必要なシートで実行してください。
- 選択した列のデータがテキストファイルとして出力されます。

このVBAコードは、選択した列のデータを効率的にテキストファイルに書き出すための強力なツールです。

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
