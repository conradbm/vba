'
' Author: Blake Conrad
' Purpose: Clear all textboxes in an Access Form
' Resource: http://www.fontstuff.com/access/acctut18.htm
'
Private Sub QueryBtn_Click()
    
    ' STRING BUILDERS
    Dim A_outputString As String
    Dim B_outputString As String
    Dim strCondition As String
    
    ' BUILD COMBOBOX A STRING
    Dim varItm As Variant
    For Each varItm In Forms!Form1!A_LB.ItemsSelected
        A_outputString = A_outputString & Chr(34) & Forms!Form1!A_LB.ItemData(varItm) & Chr(34) & ","
    Next varItm
    
    ' BUILD COMBOBOX B STRING
    For Each varItm In Forms!Form1!B_LB.ItemsSelected
        B_outputString = B_outputString & Chr(34) & Forms!Form1!B_LB.ItemData(varItm) & Chr(34) & ","
    Next varItm
    
    ' COMBOBOX A - IF NONE SELECTED ASSUME ALL
    If Len(A_outputString) = 0 Then
        A_outputString = "Like '*'"
    Else
        A_outputString = Left(A_outputString, Len(A_outputString) - 1)
        'A_outputString = Right(A_outputString, Len(A_outputString) - 1)
        A_outputString = "IN(" & A_outputString & ")"
    End If
    'Me.A_TB A_outputString
    
    ' COMBOBOX B - IF NONE SELECTED ASSUME ALL
    If Len(B_outputString) = 0 Then
        B_outputString = "Like '*'"
    Else
        B_outputString = Left(B_outputString, Len(B_outputString) - 1)
        'B_outputString = Right(B_outputString, Len(B_outputString) - 1)
        B_outputString = "IN(" & B_outputString & ")"
    End If
    'Me.B_TB = B_outputString
    
    If Me.optAnd.Value = True Then
        strCondition = " AND "
    Else
        strCondition = " OR "
    End If

    ' BUILD THE SQL STRING
    strSQL = "SELECT Table1.* FROM Table1 " & _
                 "WHERE Table1.A " & A_outputString & _
                 strCondition & "Table1.B " & _
                 B_outputString & ";"
    '
    '
    '
    '
    'MsgBox (strSQL)
    '             SELECT Table1.* FROM Table1
    '             WHERE Table1.A  IN(1,2,3) AND
    '                   Table1.B IN(2,3,4);
    '
    '
    
    ' EXECUTE SQL STRING IN ANOTHER WINDOW
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Set db = CurrentDb
    Set qdf = db.QueryDefs("queryQ")
    qdf.SQL = strSQL
    DoCmd.OpenQuery "queryQ"
    DoCmd.Close Form1, Me.Name
    Set qdf = Nothing
    Set db = Nothing
    
End Sub
