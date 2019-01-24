#### 多sheet合并的VBA

##### 地方数据汇总
```vb
Sub DownGather()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Sheets("地方数据汇总").Range("A3:S500").Clear
'        Sheets(1).Select
'        Sheets.Add '新建一个汇总表
'        Sheets(1).Name = "数据汇总"
        Dim i, j, k As Integer
            For k = 8 To 26
                If k = 8 Then
                    j = Sheets("地方数据汇总").[b65536].End(xlUp).Row + 2 '汇总所有工作表到第一个工作表，向下排列。表中间如果空一行把1改成2
                Else
                    j = Sheets("地方数据汇总").[b65536].End(xlUp).Row + 1 '汇总所有工作表到第一个工作表，向下排列。表中间如果空一行把1改成2
                End If
                Sheets(k).Select
                i = 18
                Sheets(k).Rows("4" & ":" & i).Copy
                Sheets("地方数据汇总").Cells(j, 1).PasteSpecial Paste:=xlPasteValues '如果有标题列，把前面一个1改成2
                Sheets("地方数据汇总").Cells(j, 1).PasteSpecial Paste:=xlPasteFormats
            Next
            Sheets("地方数据汇总").Select
'            Rows(1).Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
```
##### 行业数据汇总
```vb
Sub DownGather()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Sheets("行业数据汇总").Range("A3:S3000").Clear
'        Sheets(1).Select
'        Sheets.Add '新建一个汇总表
'        Sheets(1).Name = "数据汇总"
        Dim i, j, k As Integer
            For k = 8 To 26
                If k = 8 Then
                    j = Sheets("行业数据汇总").[b65536].End(xlUp).Row + 2 '汇总所有工作表到第一个工作表，向下排列。表中间如果空一行把1改成2
                Else
                    j = Sheets("行业数据汇总").[b65536].End(xlUp).Row + 1 '汇总所有工作表到第一个工作表，向下排列。表中间如果空一行把1改成2
                End If
                Sheets(k).Select
                Sheets(k).Rows("4" & ":" & 4).Copy
                Sheets("行业数据汇总").Cells(j, 1).PasteSpecial Paste:=xlPasteValues '如果有标题列，把前面一个1改成2
                Sheets("行业数据汇总").Cells(j, 1).PasteSpecial Paste:=xlPasteFormats
                If k = 8 Then
                    j = Sheets("行业数据汇总").[b65536].End(xlUp).Row + 3 '汇总所有工作表到第一个工作表，向下排列。表中间如果空一行把1改成2
                Else
                    j = Sheets("行业数据汇总").[b65536].End(xlUp).Row + 2 '汇总所有工作表到第一个工作表，向下排列。表中间如果空一行把1改成2
                End If
                i = 209
                Sheets(k).Rows("118" & ":" & i).Copy
                Sheets("行业数据汇总").Cells(j, 1).PasteSpecial Paste:=xlPasteValues '如果有标题列，把前面一个1改成2
                Sheets("行业数据汇总").Cells(j, 1).PasteSpecial Paste:=xlPasteFormats
            Next
            Sheets("行业数据汇总").Select
'            Rows(1).Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
```
##### 行业筛选
```vb
Sub 筛选行业()
'
' 筛选行业 宏
'

'
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Sheets("按行业分组").Range("A1:S3000").Clear
    Dim i, j, k, m As Integer
    Dim crit As String
    For i = 2 To 98
        Sheets("行业分类").Select
        k = 0
        For j = 2 To i
            If i <> j And ActiveSheet.Cells(j, 5) = ActiveSheet.Cells(i, 5) Then
                k = 1
                Exit For
            End If
        Next
        If k = 1 Then
            GoTo fff
        End If
        crit = ActiveSheet.Cells(i, 5)
        Sheets("行业数据汇总").Range("$C$1:$C$1769").AutoFilter Field:=1
        Sheets("行业数据汇总").Select
        ActiveSheet.Range("$C$1:$C$1769").AutoFilter Field:=1, Criteria1:=crit
        ActiveSheet.Rows("5" & ":" & "1769").EntireRow.SpecialCells(xlCellTypeVisible).Copy
        Sheets("按行业分组").Select
        m = Sheets("按行业分组").[b65536].End(xlUp).Row + 1
        Range("A" & m).select
        ActiveSheet.Paste
        Sheets("行业数据汇总").Range("$C$1:$C$1769").AutoFilter Field:=1
        
fff:
        Next
        Sheets("按行业分组").Select
End Sub
```

##### 拷贝分地区分行业企业数统计数据
```vb
Sub 分地区分行业企业数统计()
'
' 分地区分行业企业数统计 宏
'

'
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Sheets("分地区分行业企业数统计").Range("A1:S3000").Clear
    Dim j As Integer
    j = 1
    Sheets("企业数").Rows("211" & ":" & 1409).Copy
    Sheets("分地区分行业企业数统计").Cells(j, 1).PasteSpecial Paste:=xlPasteValues '如果有标题列，把前面一个1改成2
    Sheets("分地区分行业企业数统计").Cells(j, 1).PasteSpecial Paste:=xlPasteFormats
    Sheets("分地区分行业企业数统计").Select
    Sheets("分地区分行业企业数统计").Range("A1").Select
End Sub
```
```vb
Sub 编辑标题()
'
' 编辑标题 宏
'

'
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
'   设置样式准备工作
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
'   清空数据
    Sheets("按行业分组").Range("A3:S3000").Clear
    Dim i, j, k, m As Integer
    Dim companyCount, varc As Integer
    Dim crit As String
    For i = 2 To 98
        Sheets("行业分类").Select
        k = 0
        For j = 2 To i
            If i <> j And ActiveSheet.Cells(j, 5) = ActiveSheet.Cells(i, 5) Then
                k = 1
                Exit For
            End If
        Next
        If k = 1 Then
            GoTo fff
        End If
        crit = ActiveSheet.Cells(i, 5)
'       设置每个行业的标题行数据
'       去空行的开始行数
        m = Sheets("按行业分组").[h65536].End(xlUp).Row + 1
        Sheets("按行业分组").Select
        Rows(m & ":" & m).RowHeight = 31.5
'       设置样式
        Range("A" & m & ":" & "P" & m).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.399975585192419
            .PatternTintAndShade = 0
        End With
        Range("A" & m).Select
        ActiveCell.FormulaR1C1 = crit
        Range("A" & m).Select
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        Selection.Font.Bold = True
        With Selection.Font
            .Name = "等线"
            .Size = 14
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
            .ThemeFont = xlThemeFontMinor
        End With
        Selection.Font.Italic = True
'       获取行业的企业数
        companyCount = 0
        For varc = 118 To 209
            If Sheets("企业数").Cells(varc, 3) = crit Then
               companyCount = Sheets("企业数").Cells(varc, 4)
               Exit For
            End If
        Next
        Range("D" & m).Select
        ActiveCell.FormulaR1C1 = "(企业数: " & companyCount & ")"
        Selection.Font.Bold = True
'       筛选数据并复制
        Sheets("行业数据汇总").Range("$C$1:$C$1769").AutoFilter Field:=1
        Sheets("行业数据汇总").Select
        ActiveSheet.Range("$C$1:$C$1769").AutoFilter Field:=1, Criteria1:=crit
        ActiveSheet.Rows("5" & ":" & "1769").EntireRow.SpecialCells(xlCellTypeVisible).Copy
        Sheets("按行业分组").Select
        m = m + 1
        Range("A" & m).Select
        ActiveSheet.Paste
        Sheets("行业数据汇总").Range("$C$1:$C$1769").AutoFilter Field:=1
        
fff:
        Next
        Sheets("按行业分组").Select
'       删除多余的两列
        Columns("B:C").Select
        Range("C1").Activate
        Selection.Delete Shift:=xlToLeft
        Sheets("按行业分组").Range("A1").Select
End Sub

```