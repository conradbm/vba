'
'
' Author: Blake Conrad
' Purpose: Upon selecting an item in a listbox, update other things dynamically from it
'
'
Private Sub scenario_lb_AfterUpdate()
     
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' This is literally how you get a selected item from a listbox
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim selected_scenario As String
    For Each ctrl In Me.folder1.Pages("tab1").Controls
        If ctrl.Name = "scenario_lb" Then
            Dim rowIndex As Integer
            Dim rowValue As String
            Dim rowIsSelected As Integer
            Dim result As String
            rowIndex = ctrl.ListIndex
            rowValue = ctrl.Column(0)
            rowIsSelected = ctrl.Selected(rowIndex)
                    
              If (rowIsSelected = -1) Then
                result = rowValue
              Else
                
                result = rowValue
                selected_scenario = result
                'MsgBox (result)
              End If
    
        End If
    Next ctrl
    
    MsgBox (selected_scenario)

   ' Determine what to update dynamically
   Select Case selected_scenario

   Case "1"
      Me.A_LB.RowSource = "SELECT T.[X] FROM Table2 T GROUP BY T.[X]"

   Case "3"
      Me.A_LB.RowSource = "SELECT T.[Y] FROM Table2 T GROUP BY T.[Y]"

   Case "7"
      Me.A_LB.RowSourceType = "Value List"
      Me.A_LB.RowSource = "SELECT T.[Z] FROM Table2 T GROUP BY T.[Z]"

   Case Else
      MsgBox ("You're really dumb")

End Select


End Sub
