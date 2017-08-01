Private Sub getComboBoxInsideTabControl()
    
    Dim selectedPage As Page
    Dim pageIter As Page
    Dim ctrl As Control
    Dim varItm As Variant
    Dim str As String
    Set selectedPage = Me.folder1.Pages(1)
    
    For Each ctrl In selectedPage.Controls
        If ctrl.Name = "fields_lb" Then
            MsgBox ("ok...")
            For Each varItm In ctrl.ItemsSelected
            str = str & ctrl.ItemData(varItm) & ","
            Next varItm
        End If
    Next ctrl
    MsgBox (str)
    
End Sub
