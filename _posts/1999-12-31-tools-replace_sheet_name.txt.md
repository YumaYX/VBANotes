---
layout: post
category: tools
title: "replace_sheet_name.txt"
---

<small>この文章はAIで生成しています。誤りが含まれることがあります。</small>

このプログラムは、ExcelのVBAを使用して、ワークシートの名前を一括で変更するものです。具体的には、すべてのワークシートに対して、名前のうち「before」という文字列を「after」に置き換える処理を行います。

- **Dim i As Long**: 変数iを長整数型として宣言します。これはループカウンタとして使用されます。
- **For i = 1 To Worksheets.Count**: ワークシートの数だけループします。
- **sn = Worksheets(i).Name**: 現在のワークシートの名前をsnという変数に格納します。
- **Worksheets(i).Name = WorksheetFunction.Substitute(sn, "before", "after")**: snの中の「before」を「after」に置き換えた文字列で、ワークシートの名前を変更します。
- **Next**: ループを終了します。

このコードは、ワークシートの名前を一括で変更する際に非常に便利です。特に、複数のワークシートで同じような名前の変更が必要な場合に、手動で変更するよりも効率的に処理できます。

```vb
Dim i As Long
For i = 1 To Worksheets.Count
  sn = Worksheets(i).Name
  Worksheets(i).Name = WorksheetFunction.Substitute(sn, "before", "after")
Next
```
