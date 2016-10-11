Rem 版本所有：徐祥斌 xiangbin_xu@yahoo.com

Rem 生成声明
Sub declaration()
    Dim funcNo As String
    funcNo = "809402"
    
    Dim rowSelected As Integer
    
    For row = 1 To 2634
        If Cells(row, 2) = "功能号" And Cells(row, 3) = funcNo Then
            rowSelected = row
            Exit For
        End If
    Next row
    
    Dim rowStart As Integer, rowEnd As Integer
    For row = rowSelected To rowSelected + 50
        If (Cells(row, 2) = "输入参数" Or Cells(row, 2) = "D") And Cells(row + 1, 2) = "" Then
            rowStart = row + 1
        End If
        
        If Cells(row, 2) = "输出参数" And Cells(row - 1, 2) = "" Then
            rowEnd = row - 1
            Exit For
        End If
    Next row
    
    Debug.Print "row:["; rowStart; " - "; rowEnd; "]"
    
    Rem private String crdt_status;     //征信状态
    For row = rowStart To rowEnd Step 1
        Dim field As String
        Dim shuoMing As String
            
        
        field = Cells(row, 3)
        shuoMing = Cells(row, 5)
        
        Debug.Print "private String "; field; "; //"; shuoMing
        
    Next row
    
    

End Sub

Rem 生成初始化参数
Sub initParams()
    Dim funcNo As String
    funcNo = "809402"
    
    Dim rowSelected As Integer
    
    For row = 1 To 2634
        If Cells(row, 2) = "功能号" And Cells(row, 3) = funcNo Then
            rowSelected = row
            Exit For
        End If
    Next row
    
    Dim rowStart As Integer, rowEnd As Integer
    For row = rowSelected To rowSelected + 50
        If (Cells(row, 2) = "输入参数" Or Cells(row, 2) = "D") And Cells(row + 1, 2) = "" Then
            rowStart = row + 1
        End If
        
        If Cells(row, 2) = "输出参数" And Cells(row - 1, 2) = "" Then
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

Rem 根据文件内容（大写），生成域定义
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
