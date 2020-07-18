Attribute VB_Name = "シートの表示非表示"
Sub シートの表示非表示反映()

'対象ブック名取得
bookName = Range("C2")

endRow = Range("B5").End(xlDown).Row
inputData = Range("B5:C" & endRow)

Dim sheet As sheet1

'表示非表示の反映
For i = 1 To UBound(inputData, 1)
    If inputData(i, 2) = "表示" Then
        Workbooks(bookName).Sheets(inputData(i, 1)).Visible = True
    ElseIf inputData(i, 2) = "非表示" Then
        Workbooks(bookName).Sheets(inputData(i, 1)).Visible = False
    Else
        MsgBox ("C列に表示/非表示以外の文字が入っています")
        Stop
    End If
Next i

Windows(bookName).Activate
   
End Sub
