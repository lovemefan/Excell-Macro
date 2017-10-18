Sub 自动分类()

' 前提条件:总表已经根据某个要分类的标识字段排好了顺序
' 此方法是将一张数据量非常大的sheet表按某个字段进行分类并分割成不同的sheet表
' 同时将该sheet的名称命名为sheet内标识该sheet的字段的名称
    Dim sheetPage As Integer
    Dim length As Integer '总表的最后一列的列数
    Dim i As Integer
    Dim Name As String
    Dim topLeft As Integer '所选区域的左上角
    Dim bottomRight As Integer '所选区域的右下角
    
 '  参数初始化
    Name = "2013" '总表的名称
    length = 1510 '总表的最后一列的列数
    sheetPage = 4 '从第几张sheet开始生成
    topLeft = 2   '初始从第几列开始
    bottomRight = 2 '初始从第几列结束
 '
    '先建表
        For i = 2 To length
            Sheets(Name).Select
            bottomRight = i
            If Range("A" & topLeft).Value <> Range("A" & bottomRight).Value Then
                Sheets.Add After:=ActiveSheet
                topLeft = bottomRight
            End If
        Next i
    '割分
    
        topLeft = 2
        For i = 2 To length
           Sheets(Name).Select
           
           bottomRight = i
           If Range("A" & topLeft).Value <> Range("A" & bottomRight).Value Then
                Sheets(Name).Select
                Range("A" & topLeft & ":" & "G" & (bottomRight - 1)).Select '选择从AtopLeft到Gbottom的区域等待复制
                Selection.Copy
                
                Sheets(sheetPage).Select
                Range("A1").Select
                ActiveSheet.Paste
                Sheets(sheetPage).Name = Range("A1").Value
                sheetPage = sheetPage + 1
                topLeft = bottomRight
           End If
            
        Next i
        
         Sheets("2013").Select
         Range("A" & topLeft & ":" & "G" & (bottomRight - 1)).Select
         Selection.Copy
         
         Sheets(sheetPage - 1).Select
         Sheets.Add After:=ActiveSheet
         Sheets(sheetPage).Select
         
         Range("A1").Select
         ActiveSheet.Paste
         sheetName = Range("A1").Value
         Sheets(sheetPage).Name = sheetName
         
         topLeft = bottomRight
    
    
    
End Sub
Sub 排序()
    For Each Sh In Worksheets
        If Sh.Index > 4 Then  '限定工作表范围
            Sh.Select
            Columns("B:B").Select
            Sh.Sort.SortFields.Clear
            Sh.Sort.SortFields.Add Key:=Range("B1"), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            With Sh.Sort
                .SetRange Range("A1:G69")
                .Header = xlNo
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            End If
        Next




End Sub


