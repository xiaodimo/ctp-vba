Rem �汾���У������ xiangbin_xu@yahoo.com

Rem ��������
Sub declaration()
    Dim funcNo As String
    funcNo = "809402"
    
    Dim rowSelected As Integer
    
    For row = 1 To 2634
        If Cells(row, 2) = "���ܺ�" And Cells(row, 3) = funcNo Then
            rowSelected = row
            Exit For
        End If
    Next row
    
    Dim rowStart As Integer, rowEnd As Integer
    For row = rowSelected To rowSelected + 50
        If (Cells(row, 2) = "�������" Or Cells(row, 2) = "D") And Cells(row + 1, 2) = "" Then
            rowStart = row + 1
        End If
        
        If Cells(row, 2) = "�������" And Cells(row - 1, 2) = "" Then
            rowEnd = row - 1
            Exit For
        End If
    Next row
    
    Debug.Print "row:["; rowStart; " - "; rowEnd; "]"
    
    Rem private String crdt_status;     //����״̬
    For row = rowStart To rowEnd Step 1
        Dim field As String
        Dim shuoMing As String
            
        
        field = Cells(row, 3)
        shuoMing = Cells(row, 5)
        
        Debug.Print "private String "; field; "; //"; shuoMing
        
    Next row
    
    

End Sub

Rem ���ɳ�ʼ������
Sub initParams()
    Dim funcNo As String
    funcNo = "809402"
    
    Dim rowSelected As Integer
    
    For row = 1 To 2634
        If Cells(row, 2) = "���ܺ�" And Cells(row, 3) = funcNo Then
            rowSelected = row
            Exit For
        End If
    Next row
    
    Dim rowStart As Integer, rowEnd As Integer
    For row = rowSelected To rowSelected + 50
        If (Cells(row, 2) = "�������" Or Cells(row, 2) = "D") And Cells(row + 1, 2) = "" Then
            rowStart = row + 1
        End If
        
        If Cells(row, 2) = "�������" And Cells(row - 1, 2) = "" Then
            rowEnd = row - 1
            Exit For
        End If
    Next row
    
    Debug.Print "row:["; rowStart; " - "; rowEnd; "]"
    
    Rem params.put(Fields.MONEY_TYPE, getMoney_type())
    For row = rowStart To rowEnd Step 1
        Dim field As String
        field = Cells(row, 3)
        
        Dim strLen As Integer
        strLen = Len(field)
        
        Dim firstChar As String, secondStr As String
        firstChar = Left(field, 1)
        secondStr = Mid(field, 2, strLen - 1)

        Debug.Print "params.put(Fields."; UCase(field); ", get"; UCase(firstChar); secondStr; "());"
        
    Next row
    
    
    

End Sub

Rem �����ļ����ݣ���д������������
Sub generateFields()
    Dim text As String
    
    Close #1
    Open "D:\fields.txt" For Input As #1
    
    Rem public static final String ACPT_ID = "acpt_id"
    Do While Not EOF(1)
        Line Input #1, text
        
        Dim shuoMing As String
        For row = 1 To 2636
            If Cells(row, 3) = LCase(text) Then
                shuoMing = Cells(row, 5).Value
                Exit For
            End If
        Next row
        
        Debug.Print "public static final String "; UCase(text); " = "; Chr(34); LCase(text); Chr(34); ";"; " //"; shuoMing
    Loop
    
    Close #1

End Sub
