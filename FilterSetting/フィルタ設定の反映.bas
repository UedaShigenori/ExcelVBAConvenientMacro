Attribute VB_Name = "�t�B���^�ݒ�̔��f"
'Option Base 1
Sub �t�B���^�ݒ�̔��f()

'�Ώۃu�b�N���擾
bookName = Range("C2")
sheetName = Range("C3")
headLineCell = Range("C4")
firstCell = Range("C5")

'�t�B���^�������擾
endRow = Range("B11").End(xlDown).Row
inputData = Range("B11:C" & endRow)

Dim dataNum As Integer
dataNum = UBound(inputData, 1)

Dim search1() As String
ReDim Preserve search1(dataNum)

For i = 1 To dataNum
    search1(i) = inputData(i, 1)
Next

'�^�[�Q�b�g�ƂȂ�t�B���^�̑I��
firstCol = Range(firstCell).Column
Set r = Workbooks(bookName).Worksheets(sheetName).Range(firstCell).CurrentRegion
headLineCol = r.Find(what:=headLineCell, after:=r(1)).Column
targetField = headLineCol - firstCol + 1

'�t�B���^�����O
Workbooks(bookName).Worksheets(sheetName).Range(firstCell).AutoFilter _
    Field:=targetField, _
    Criteria1:=search1, _
    Operator:=xlFilterValues

Windows(bookName).Activate
Range(firstCell).Select

End Sub
