Attribute VB_Name = "�t�B���^�ƃE�B���h�E�g�̌Œ�"
'Option Base 1
Sub �t�B���^�ƃE�B���h�E�g�̌Œ�()

'�Ώۃu�b�N���擾
bookName = Range("C2")
sheetName = Range("C3")

Windows(bookName).Activate
Sheets(sheetName).Select
Range("A4").Select


Selection.End(xlToRight).Select
ActiveCell.Offset(1, 1).Range("A1").Select
Selection.Copy
Application.CutCopyMode = False
Selection.AutoFill Destination:=ActiveCell.Range("A1:A1004"), Type:= _
    xlFillDefault
ActiveCell.Range("A1:A1004").Select
ActiveCell.Offset(-1, 0).Range("A1").Select
Range(Selection, Selection.End(xlToLeft)).Select
Range(Selection, Selection.End(xlToLeft)).Select
Selection.AutoFilter
ActiveSheet.Range("$A$4:$L$19").AutoFilter Field:=12, Criteria1:="1"

Range("A1").Select
Range("E5").Select
ActiveWindow.FreezePanes = True


Windows(bookName).Activate
Range("A1").Select


End Sub
