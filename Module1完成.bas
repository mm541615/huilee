Attribute VB_Name = "Module1"
Sub 更新庫存()
    
    
    更新借歸還庫存
    更新庫存日期
End Sub

Sub 庫存計算()

Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
Dim lastRow1 As Long, lastRow2 As Long, lastRow3 As Long
Dim i As Long
Dim itemName As String
Dim stockQty As Integer, borrowQty As Integer, returnQty As Integer
Dim salesQty As Integer, purchaseQty As Integer, actualStockQty As Integer

 ' 設置表單
    Set ws1 = ThisWorkbook.Sheets("庫存")
    Set ws2 = ThisWorkbook.Sheets("借歸還")
    Set ws3 = ThisWorkbook.Sheets("進出貨")

    lastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    lastRow3 = ws3.Cells(ws3.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow1
            itemName = ws1.Cells(i, "A").Value ' 取得商品名稱
            stockQty = CDbl(ws1.Cells(i, "E").Value) ' 取得庫存數量並轉換為整數型態

    ws1.Range("E2:E" & lastRow1).NumberFormat = "0" ' 庫存表庫存數量的格式設定為數值格式
    ws2.Range("D2:E" & lastRow2).NumberFormat = "0" ' 借歸還表借用數量和歸還數量的格式設定為數值格式
    ws3.Range("E2:F" & lastRow3).NumberFormat = "0" ' 進出貨表出貨數量和進貨數量的格式設定為數值格式
  
        ' 初始化數量變數
        borrowQty = 0
        returnQty = 0
        salesQty = 0
        purchaseQty = 0

    ' 借歸還表，找對應的商品名稱並計算借用數量和歸還數量
    For j = 2 To lastRow2
    
            If ws2.Cells(j, "A").Value = itemName Then
                borrowQty = borrowQty + ws2.Cells(j, "D").Value ' 借用數量累計
                returnQty = returnQty + ws2.Cells(j, "E").Value ' 歸還數量累計

            End If
            
    Next j
        
    ' 進出貨表，找對應的商品名稱並計算出貨和進貨數量
   For k = 2 To lastRow3
   
            If ws3.Cells(k, "A").Value = itemName Then
                salesQty = salesQty + ws3.Cells(k, "E").Value ' 出貨數量累計
                purchaseQty = purchaseQty + ws3.Cells(k, "F").Value ' 進貨數量累計
            End If
            
    Next k
        
        ' 計算實際庫存量
        actualStockQty = stockQty + purchaseQty - borrowQty + returnQty - salesQty
        
        ' 實際庫存量寫入庫存表的相應欄位
        ws1.Cells(i, "H").Value = actualStockQty
    Next i
End Sub

Sub 更新進銷貨庫存()
    ' 更新進出貨資料
    Call 庫存計算
End Sub

Sub 更新借歸還庫存()
    ' 更新借歸還資料
    Call 庫存計算
End Sub

Sub 更新庫存日期()
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long, lastRow3 As Long
    Dim i As Long
    Dim itemName As String
    Dim currentDate As Date
    
    ' 設置表單
    Set ws1 = ThisWorkbook.Sheets("庫存")
    Set ws2 = ThisWorkbook.Sheets("借歸還")
    Set ws3 = ThisWorkbook.Sheets("進出貨")
    
    lastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    lastRow3 = ws3.Cells(ws3.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow1
        itemName = ws1.Cells(i, "A").Value ' 取得商品名稱
        
        ' 尋找相應的進出貨資料
        For j = 2 To lastRow3
            If ws3.Cells(j, "A").Value = itemName Then
                currentDate = ws3.Cells(j, "H").Value ' 取得進貨日期
                ws1.Cells(i, "I").Value = currentDate ' 更新庫存日期
                Exit For ' 找到相符的進貨資料後結束迴圈
            End If
        Next j
        
        ' 尋找相應的借歸還資料
        For k = 2 To lastRow2
            If ws2.Cells(k, "A").Value = itemName Then
                currentDate = ws2.Cells(k, "G").Value ' 取得借用日期
                ws1.Cells(i, "I").Value = currentDate ' 更新庫存日期
                Exit For ' 找到相符的借歸還資料後結束迴圈
            End If
        Next k
    Next i
End Sub
