Private Sub optAnd_Click()
    If Me.optAnd.Value = True Then
        Me.optOr.Value = False
    Else
        Me.optOr.Value = True
    End If
End Sub

Private Sub optOr_Click()
    If Me.optOr.Value = True Then
        Me.optAnd.Value = False
    Else
        Me.optAnd.Value = True
    End If
End Sub
