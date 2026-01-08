---
layout: post
category: tools
title: "get_all_sheet_name.txt"
---

<small>この文章はAIで生成しています。誤りが含まれることがあります。</small>

このプログラムは、ExcelのVBA（Visual Basic for Applications）を使用して、ワークシートの名前を順番に表示するプログラムです。以下にその詳細な説明をMarkdown形式で記載します。

このプログラムは、ExcelのVBA環境で実行されます。`Dim i As Long`は、変数`i`を長整数型として宣言しています。この変数は、ループカウンタとして使用されます。`For i = 1 To Worksheets.Count`は、`Worksheets`コレクション内のワークシートの数だけループを繰り返します。`Debug.Print Worksheets(i).Name`は、現在のワークシートの名前をデバッグウィンドウに表示します。`Next`は、ループを終了し、次のワークシートに進みます。このプログラムは、Excelのワークシートの名前を順番に表示するための基本的なVBAプログラムです。

```vb
Dim i As Long
For i = 1 To Worksheets.Count
    Debug.Print Worksheets(i).Name
Next
```
