Sub �Զ�����()

' ǰ������:�ܱ��Ѿ�����ĳ��Ҫ����ı�ʶ�ֶ��ź���˳��
' �˷����ǽ�һ���������ǳ����sheet��ĳ���ֶν��з��ಢ�ָ�ɲ�ͬ��sheet��
' ͬʱ����sheet����������Ϊsheet�ڱ�ʶ��sheet���ֶε�����
    Dim sheetPage As Integer
    Dim length As Integer '�ܱ�����һ�е�����
    Dim i As Integer
    Dim Name As String
    Dim topLeft As Integer '��ѡ��������Ͻ�
    Dim bottomRight As Integer '��ѡ��������½�
    
 '  ������ʼ��
    Name = "2013" '�ܱ������
    length = 1510 '�ܱ�����һ�е�����
    sheetPage = 4 '�ӵڼ���sheet��ʼ����
    topLeft = 2   '��ʼ�ӵڼ��п�ʼ
    bottomRight = 2 '��ʼ�ӵڼ��н���
 '
    '�Ƚ���
        For i = 2 To length
            Sheets(Name).Select
            bottomRight = i
            If Range("A" & topLeft).Value <> Range("A" & bottomRight).Value Then
                Sheets.Add After:=ActiveSheet
                topLeft = bottomRight
            End If
        Next i
    '���
    
        topLeft = 2
        For i = 2 To length
           Sheets(Name).Select
           
           bottomRight = i
           If Range("A" & topLeft).Value <> Range("A" & bottomRight).Value Then
                Sheets(Name).Select
                Range("A" & topLeft & ":" & "G" & (bottomRight - 1)).Select 'ѡ���AtopLeft��Gbottom������ȴ�����
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
Sub ����()
    For Each Sh In Worksheets
        If Sh.Index > 4 Then  '�޶�������Χ
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


