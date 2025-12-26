---
layout: post
category: tools
title: "get_all_sheet_name.txt"
---

<small>この文章はAIで生成しています。誤りが含まれることがあります。</small>

このプログラムは、ExcelのVBA（Visual Basic for Applications）を使用して、ワークシートの名前を順番に表示するプログラムです。以下にその詳細な説明を日本語でMarkdown形式で記載します。

### プログラムの説明

#### 目的
- Excelのワークシートの名前を順番に表示する。

#### コードの詳細
```vba
Dim i As Long
For i = 1 To Worksheets.Count
    Debug.Print Worksheets(i).Name
Next
```

#### コードの解説

1. **変数の宣言**:
   - `Dim i As Long`: `i`という名前の変数を宣言し、`Long`型として使用します。これは、ループカウンタとして使用されます。

2. **Forループ**:
   - `For i = 1 To Worksheets.Count`: `Worksheets.Count`は現在開いているワークシートの数を返します。このループは1からその数まで繰り返されます。

3. **ワークシート名の表示**:
   - `Debug.Print Worksheets(i).Name`: 現在のループカウンタ`i`に対応するワークシートの名前を表示します。`Debug.Print`は、デバッグウィンドウにテキストを出力するためのVBAの関数です。

4. **ループの終了**:
   - `Next`: ループの終了を示します。

#### 使用例
- このプログラムを実行すると、Excelのデバッグウィンドウに、現在開いているすべてのワークシートの名前が順番に表示されます。

このプログラムは、VBAの基本的なループ構造とデバッグ出力の使用方法を示すシンプルな例です。

```vb
Dim i As Long
For i = 1 To Worksheets.Count
    Debug.Print Worksheets(i).Name
Next
```
