Attribute VB_Name = "Module1"
Sub スペース削除()

Dim Parag As Paragraph
Dim C_Txt(1) As String
C_Txt(0) = " "
C_Txt(1) = "　"

MsgBox "段落先頭の半角、全角スペースを削除します。"

For Each Parag In ActiveDocument.Paragraphs
    For i = 0 To UBound(C_Txt)
        If Parag.Range.Characters.Item(1).Text = C_Txt(i) Then
            Parag.Range.Characters.Item(1).Text = ""
        End If
    Next i
Next Parag

MsgBox "処理終了"
End Sub
