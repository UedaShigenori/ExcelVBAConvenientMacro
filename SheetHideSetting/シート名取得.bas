Attribute VB_Name = "�V�[�g���擾"
Option Base 1
Sub �V�[�g���擾()

'�����f�[�^�̍폜
Range("B5:B1000").ClearContents

'�Ώۃu�b�N���擾
bookName = Range("C2")

Dim sheetName() As Variant
Dim sheetNum As Integer
sheetNum = 1

'�V�[�g���擾
For Each i In Workbooks(bookName).Sheets
    ReDim Preserve sheetName(sheetNum)
    sheetName(sheetNum) = i.Name
    sheetNum = sheetNum + 1
Next i

'�V�[�g����������
endRow = 4 + UBound(sheetName)
Range("B5:B" & endRow) = WorksheetFunction.Transpose(sheetName)

End Sub

