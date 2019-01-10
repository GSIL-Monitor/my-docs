```
Sub yy()
    Dim sheetNum As Integer
    Dim value As String
    Dim rownum As Integer
    Dim columnnum As Integer
    Dim columnChar As String
    Dim srccolumnnum As Integer
    Dim srccolumnChar As String
    Dim i, j, a, b As Integer
    Dim flag As Boolean
    
    sheetNum = 5
    
    Sheets("合并").Cells.Clear
    Sheets("sheet1").Range("A1").CurrentRegion.Copy Sheets("合并").Range("A1")
    rownum = Sheets("合并").UsedRange.Rows.Count
    
    Rem 以下查找记录并复制到合并表中
    For i = 2 To sheetNum Step 1
        columnnum = Sheets("合并").UsedRange.Columns.Count + 1
        columnChar = Replace(Cells(1, columnnum).Address(False, False), "1", "")
        Rem 复制内容
        For j = 2 To rownum Step 1
            Sheets("合并").Activate
            value = Cells(j, 1)
            Sheets(i).Activate
            srccolumnnum = Sheets(i).UsedRange.Columns.Count + 1
            srccolumnChar = Replace(Cells(1, srccolumnnum).Address(False, False), "1", "")
            Rem 1.复制标题
            Sheets(i).Range("b1:" & srccolumnChar & "1").Copy Sheets("合并").Range(columnChar & 1)
            Rem 2.找到符合的记录粘贴
            flag = False
            For Each Rng In Range("a2:a" & rownum)
                If Rng = value Then
                    a = Rng.Row
                    b = Rng.Column
                    flag = True
                End If
                If flag = True Then
                    Exit For
                End If
            Next
            Sheets(i).Range("b" & a & ":" & srccolumnChar & a).Copy Sheets("合并").Range(columnChar & j)
        Next
    Next
    Rem 跳转到合并页
    Sheets("合并").Activate
    
    Rem 测试代码
    Rem Sheets(2).Activate
    Rem Sheets("sheet2").Range("b1:" & srccolumnChar & "1").EntireColumn.Copy Sheets("合并").Range(columnChar & "1")
    
    Rem Sheets("sheet3").Range("A1").CurrentRegion.Copy Sheets("合并").Range("C1")
    Rem Sheets("sheet4").Range("A1").CurrentRegion.Copy Sheets("合并").Range("D1")
    Rem Sheets("sheet5").Range("A1").CurrentRegion.Copy Sheets("合并").Range("E1")
End Sub

```