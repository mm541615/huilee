Attribute VB_Name = "Module1"
Sub ��s�w�s()
    
    
    ��s���k�ٮw�s
    ��s�w�s���
End Sub

Sub �w�s�p��()

Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
Dim lastRow1 As Long, lastRow2 As Long, lastRow3 As Long
Dim i As Long
Dim itemName As String
Dim stockQty As Integer, borrowQty As Integer, returnQty As Integer
Dim salesQty As Integer, purchaseQty As Integer, actualStockQty As Integer

 ' �]�m���
    Set ws1 = ThisWorkbook.Sheets("�w�s")
    Set ws2 = ThisWorkbook.Sheets("���k��")
    Set ws3 = ThisWorkbook.Sheets("�i�X�f")

    lastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    lastRow3 = ws3.Cells(ws3.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow1
            itemName = ws1.Cells(i, "A").Value ' ���o�ӫ~�W��
            stockQty = CDbl(ws1.Cells(i, "E").Value) ' ���o�w�s�ƶq���ഫ����ƫ��A

    ws1.Range("E2:E" & lastRow1).NumberFormat = "0" ' �w�s��w�s�ƶq���榡�]�w���ƭȮ榡
    ws2.Range("D2:E" & lastRow2).NumberFormat = "0" ' ���k�٪�ɥμƶq�M�k�ټƶq���榡�]�w���ƭȮ榡
    ws3.Range("E2:F" & lastRow3).NumberFormat = "0" ' �i�X�f��X�f�ƶq�M�i�f�ƶq���榡�]�w���ƭȮ榡
  
        ' ��l�Ƽƶq�ܼ�
        borrowQty = 0
        returnQty = 0
        salesQty = 0
        purchaseQty = 0

    ' ���k�٪�A��������ӫ~�W�٨íp��ɥμƶq�M�k�ټƶq
    For j = 2 To lastRow2
    
            If ws2.Cells(j, "A").Value = itemName Then
                borrowQty = borrowQty + ws2.Cells(j, "D").Value ' �ɥμƶq�֭p
                returnQty = returnQty + ws2.Cells(j, "E").Value ' �k�ټƶq�֭p

            End If
            
    Next j
        
    ' �i�X�f��A��������ӫ~�W�٨íp��X�f�M�i�f�ƶq
   For k = 2 To lastRow3
   
            If ws3.Cells(k, "A").Value = itemName Then
                salesQty = salesQty + ws3.Cells(k, "E").Value ' �X�f�ƶq�֭p
                purchaseQty = purchaseQty + ws3.Cells(k, "F").Value ' �i�f�ƶq�֭p
            End If
            
    Next k
        
        ' �p���ڮw�s�q
        actualStockQty = stockQty + purchaseQty - borrowQty + returnQty - salesQty
        
        ' ��ڮw�s�q�g�J�w�s���������
        ws1.Cells(i, "H").Value = actualStockQty
    Next i
End Sub

Sub ��s�i�P�f�w�s()
    ' ��s�i�X�f���
    Call �w�s�p��
End Sub

Sub ��s���k�ٮw�s()
    ' ��s���k�ٸ��
    Call �w�s�p��
End Sub

Sub ��s�w�s���()
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long, lastRow3 As Long
    Dim i As Long
    Dim itemName As String
    Dim currentDate As Date
    
    ' �]�m���
    Set ws1 = ThisWorkbook.Sheets("�w�s")
    Set ws2 = ThisWorkbook.Sheets("���k��")
    Set ws3 = ThisWorkbook.Sheets("�i�X�f")
    
    lastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    lastRow3 = ws3.Cells(ws3.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow1
        itemName = ws1.Cells(i, "A").Value ' ���o�ӫ~�W��
        
        ' �M��������i�X�f���
        For j = 2 To lastRow3
            If ws3.Cells(j, "A").Value = itemName Then
                currentDate = ws3.Cells(j, "H").Value ' ���o�i�f���
                ws1.Cells(i, "I").Value = currentDate ' ��s�w�s���
                Exit For ' ���۲Ū��i�f��ƫᵲ���j��
            End If
        Next j
        
        ' �M����������k�ٸ��
        For k = 2 To lastRow2
            If ws2.Cells(k, "A").Value = itemName Then
                currentDate = ws2.Cells(k, "G").Value ' ���o�ɥΤ��
                ws1.Cells(i, "I").Value = currentDate ' ��s�w�s���
                Exit For ' ���۲Ū����k�ٸ�ƫᵲ���j��
            End If
        Next k
    Next i
End Sub
