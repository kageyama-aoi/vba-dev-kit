# PLAN - やりたいこと

<!-- ここに思ったことを自由に書いてください。箇条書きでも口語でもOK -->
<!-- Claude がこの内容を読んでヒアリングし、SPEC.md を作成します -->

# 別のAIで質問して対応できてない件

あー！ボタン（フォームコントロール or ActiveXボタン）も ws.Shapes に含まれるので、RenameShapesSequentially で Flow_001 などの名前が付いてしまっていたんですね。
ボタンは TextFrame2 の操作が正常にできないのでそこで落ちていた、ということですね。
RenameShapesSequentially と InjectTextFromList の両方でボタンを除外するのが根本対策です：
vba' RenameShapesSequentially のソート前の収集部分
For i = 1 To ws.Shapes.Count
    ' ボタン系（フォームコントロール・ActiveX）を除外
    If ws.Shapes(i).Type <> msoFormControl And _
       ws.Shapes(i).Type <> msoOLEControlObject Then
        n = n + 1
        ReDim Preserve arr(1 To n)
        Set arr(n) = ws.Shapes(i)
    End If
Next i
InjectTextFromList 側は図形名で引いているので、ボタンに Flow_001 などの名前さえ付いていなければ自然に回避できます。なので RenameShapesSequentially だけ直せばOKです。