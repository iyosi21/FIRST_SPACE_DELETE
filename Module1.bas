Attribute VB_Name = "Module1"
Sub �X�y�[�X�폜()

Dim Parag As Paragraph
Dim C_Txt(1) As String
C_Txt(0) = " "
C_Txt(1) = "�@"

MsgBox "�i���擪�̔��p�A�S�p�X�y�[�X���폜���܂��B"

For Each Parag In ActiveDocument.Paragraphs
    For i = 0 To UBound(C_Txt)
        If Parag.Range.Characters.Item(1).Text = C_Txt(i) Then
            Parag.Range.Characters.Item(1).Text = ""
        End If
    Next i
Next Parag

MsgBox "�����I��"
End Sub
