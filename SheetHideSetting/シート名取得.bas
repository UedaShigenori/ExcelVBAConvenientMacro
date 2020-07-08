Attribute VB_Name = "シート名取得"
Option Base 1
Sub シート名取得()

'既存データの削除
Range("B5:B1000").ClearContents

'対象ブック名取得
bookName = Range("C2")

Dim sheetName() As Variant
Dim sheetNum As Integer
sheetNum = 1

'シート名取得
For Each i In Workbooks(bookName).Sheets
    ReDim Preserve sheetName(sheetNum)
    sheetName(sheetNum) = i.Name
    sheetNum = sheetNum + 1
Next i

'シート名書き込み
endRow = 4 + UBound(sheetName)
Range("B5:B" & endRow) = WorksheetFunction.Transpose(sheetName)

End Sub

