Attribute VB_Name = "�V�[�g�̕\����\��"
Sub �V�[�g�̕\����\�����f()

'�Ώۃu�b�N���擾
bookName = Range("C2")

endRow = Range("B5").End(xlDown).Row
inputData = Range("B5:C" & endRow)

Dim sheet As sheet1

'�\����\���̔��f
For i = 1 To UBound(inputData, 1)
    If inputData(i, 2) = "�\��" Then
        Workbooks(bookName).Sheets(inputData(i, 1)).Visible = True
    ElseIf inputData(i, 2) = "��\��" Then
        Workbooks(bookName).Sheets(inputData(i, 1)).Visible = False
    Else
        MsgBox ("C��ɕ\��/��\���ȊO�̕����������Ă��܂�")
        Stop
    End If
Next i

Windows(bookName).Activate
   
End Sub
