'
' Author: Blake Conrad
' Purpose: Clear all textboxes in an Access Form
' Resource: http://www.fontstuff.com/access/acctut18.htm
'
Private Sub ClearBtn_Click()
    ' Centry Variable
    Dim ctrl As Control
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" Then
            ctrl = ""
        End If

     Next ctrl
End Sub
