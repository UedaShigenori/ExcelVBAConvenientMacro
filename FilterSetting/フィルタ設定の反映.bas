Attribute VB_Name = "フィルタ設定の反映"
'Option Base 1
Sub フィルタ設定の反映()

'対象ブック名取得
bookName = Range("C2")
sheetName = Range("C3")
Dim headLineCell As String
headLineCell = Range("C4").Value
firstCell = Range("C5")

'フィルタ条件を取得
endRow = Range("B11").End(xlDown).Row
inputData = Range("B11:C" & endRow)

Dim dataNum As Integer
dataNum = UBound(inputData, 1)

Dim search1() As String
ReDim Preserve search1(dataNum)

For i = 1 To dataNum
    search1(i) = inputData(i, 1)
Next

'ターゲットとなるフィルタの選定
firstCol = Range(firstCell).Column
Set r = Workbooks(bookName).Worksheets(sheetName).Range(firstCell).CurrentRegion

Dim firstData As String
firstData = Workbooks(bookName).Worksheets(sheetName).Range(firstCell).Value

If firstData = headLineCell Then
    headLineCol = 1
Else
    
    headLineCol = r.Find(what:=headLineCell, after:=r(1)).Column
End If
targetField = headLineCol - firstCol + 1

'フィルタリング
Workbooks(bookName).Worksheets(sheetName).Range(firstCell).AutoFilter _
    Field:=targetField, _
    Criteria1:=search1, _
    Operator:=xlFilterValues

Windows(bookName).Activate
Range(firstCell).Select

End Sub
