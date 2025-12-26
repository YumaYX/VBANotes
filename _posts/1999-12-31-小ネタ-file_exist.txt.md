---
layout: post
category: 小ネタ
title: "file_exist.txt"
---

<small>この文章はAIで生成しています。誤りが含まれることがあります。</small>

このプログラムは、指定されたファイル名が存在するかどうかをチェックし、その結果をメッセージボックスで表示するVBAコードです。具体的には、変数strにファイル名を設定し、Dir関数を使ってそのファイルが存在するかを確認します。存在する場合は「有り」、存在しない場合は「無し」というメッセージボックスが表示されます。

```vb
Dim str As String: str = "filename"
If Dir(str) <> "" Then
    MsgBox "有り"
Else
    MsgBox "無し"
End If
```
