---
layout: post
category: 小ネタ
title: "file_exist.txt"
---

<small>この文章はAIで生成しています。誤りが含まれることがあります。</small>

このプログラムは、指定されたファイル名がシステム内に存在するかどうかをチェックし、その結果をメッセージボックスで表示するVBAコードです。具体的には、変数`str`にチェックしたいファイル名を代入し、`Dir`関数を使用してそのファイルの存在を確認します。存在する場合は「有り」、存在しない場合は「無し」というメッセージボックスが表示されます。

```vb
Dim str As String: str = "filename"
If Dir(str) <> "" Then
    MsgBox "有り"
Else
    MsgBox "無し"
End If
```
