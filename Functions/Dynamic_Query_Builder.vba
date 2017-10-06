Option Compare Database

'Author: Blake Conrad

'////////////////////////////////////////////////////////////////'
' setLabelOnPage SUBROUTINE
'
'
'
' USAGE: setLabelOnPage(somePage, someLabelsName, someTextToChange)
'
'
'
'
' GOALS:
'       1. UPDATE THE LABEL OF A TEXTBOX
'
'
'
'
'////////////////////////////////////////////////////////////////'
Private Sub setLabelOnPage(p As Page, labelName As String, labelString As String)
    Dim c As Control
    For Each c In p.Controls
        If c.name = labelName Then
            c.Caption = labelString
        End If
    Next c
End Sub


'////////////////////////////////////////////////////////////////'
' getLabelOnPage SUBROUTINE
'
'
'
' USAGE: getLabelOnPage(somePage, someLabelsName)
'
'
'
'
' GOALS:
'       1. GET THE NAME OF A TEXTBOX ON A PAGE TAB
'
'
'
'
'////////////////////////////////////////////////////////////////'
Private Function getLabelOnPage(p As Page, labelName As String) As String
    Dim c As Control
    For Each c In p.Controls
        If c.name = labelName Then
            getLabelOnPage = c.Caption
        End If
    Next c
End Function


'////////////////////////////////////////////////////////////////'
' Make_All_ListBoxes_Visible SUBROUTINE
'
'
'
' USAGE: Make_All_ListBoxes_Visible
'
'
'
'
' GOALS:
'       1. LIGHTS UP ALL COMBOBOXS TO 'VISIBLE'
'
'
'
'
'////////////////////////////////////////////////////////////////'
Private Sub Make_All_ListBoxes_Visible()
    For Each ctrl In Me.Controls!Folder1.Pages("Interests").Controls
        If ctrl.name = "fields_lb" Then
                For i = 0 To ctrl.ListCount - 1
                     
                    ' TAB 1 VISIBILITY CONTROLS
                    Dim c1 As Control
                    For Each c1 In Me.Controls!Folder2.Pages("Unit").Controls
                        If c1.name = ctrl.Column(1, i) Then
                            c1.Visible = True
                        End If
                    Next c1
                
                    ' TAB 2 VISIBILITY CONTROLS
                    Dim c2 As Control
                    For Each c2 In Me.Controls!Folder2.Pages("Communication").Controls
                        If c2.name = ctrl.Column(1, i) Then
                            c2.Visible = True
                        End If
                    Next c2
                    
                    ' TAB 3 VISIBILITY CONTROLS
                    Dim c3 As Control
                    For Each c3 In Me.Controls!Folder2.Pages("Misc.").Controls
                        If c3.name = ctrl.Column(1, i) Then
                            c3.Visible = True
                        End If
                    Next c3
                Next i
            Exit For
        End If
    Next ctrl
End Sub
    

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////'
' Determine_Visibility_For_Each_ListBox SUBROUTINE
'
'
'
' USAGE: Call Determine_Visibility_For_Each_ListBox(someControl, someIndex)
'
'
'
'
' GOALS:
'       1. FLICKER ON OR OFF VISIBILITY BASED ON IF THE CONTROL INDEX VALUE IS SELECTED OR NOT
'
'
'
'
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////'
Private Sub Determine_Visibility_For_Each_ListBox(ctrl As Control, i)
    ' TAB 1 VISIBILITY CONTROLS
    Dim c1 As Control
    For Each c1 In Me.Controls!Folder2.Pages("Unit").Controls
        If c1.name = ctrl.Column(1, i) Then
            ' IF SELECTED, MAKE VISIBLE
            If ctrl.Selected(i) Then
                If c1.Visible = False Then
                    c1.Visible = True
                End If
            ' IF NOT SELECTED, MAKE INVISIBLE
            Else
                If c1.Visible = True Then
                '
                '
                '
                ' For each item in c1 listbox
                    'if item.Selected Then
                        'item.Selected=False
                    'End If
                'Next
                    c1.Visible = False
                End If
            End If
        End If
    Next c1

    ' TAB 2 VISIBILITY CONTROLS
    Dim c2 As Control
    For Each c2 In Me.Controls!Folder2.Pages("Communication").Controls
        If c2.name = ctrl.Column(1, i) Then
            ' IF SELECTED, MAKE VISIBLE
            If ctrl.Selected(i) Then
                If c2.Visible = False Then
                    c2.Visible = True
                End If
            ' IF NOT SELECTED, MAKE INVISIBLE
            Else
                If c2.Visible = True Then
                    c2.Visible = False
                End If
            End If
        End If
    Next c2
    
    ' TAB 3 VISIBILITY CONTROLS
    Dim c3 As Control
    For Each c3 In Me.Controls!Folder2.Pages("Misc.").Controls
        If c3.name = ctrl.Column(1, i) Then
            ' IF SELECTED, MAKE VISIBLE
            If ctrl.Selected(i) Then
                If c3.Visible = False Then
                    c3.Visible = True
                End If
            ' IF NOT SELECTED, MAKE INVISIBLE
            Else
                If c3.Visible = True Then
                    c3.Visible = False
                End If
            End If
        End If
    Next c3

End Sub







'////////////////////////////////////////////////////////////////'
' FIELDS_LB_MouseUp INSIDE FOLDER 1 "Interests" Tab
'
'
'
' GOALS:
'       1. GET SCENARIO(S) DESIRED, IF NONE THEN ASSUME ALL
'       2. GET TIMEFRAME, IF NONE THEN ASSUME ALL
'       3. GET FIELDS DESIRED, IF NONE, THEN ASSUME ALL
'           3a. MAKE COMBOBOXES VISIBLE FOR FIELDS DESIRED
'
'
'
'
'////////////////////////////////////////////////////////////////'
Private Sub FIELDS_LB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

     ' GET THE PAGE
    Dim selectedPage As Page
    Dim pageIter As Page
    Dim ctrl As Control
    Dim varItm As Variant
    Set selectedPage = Me.Controls!Folder1.Pages("Interests")
    
    ' STRINGS
    Dim scenarioString As String
    Dim timeframeString As String
    Dim fieldsString As String
    
    ' INTERESTS TAB -- TURN ON APPROPRIATE
    For Each ctrl In Me.Controls!Folder1.Pages("Interests").Controls
        If ctrl.name = "fields_lb" Then
    
                
                For i = 0 To ctrl.ListCount - 1
                    ' ////////////////////////////////////////////////////////////////////////////////////////////////
                    ' *** THIS CHECKS WHICH ONES ARE ON/OFF, AND FLICKERS VISIBILITY AS NEEDED
                    ' ////////////////////////////////////////////////////////////////////////////////////////////////
                    Call Determine_Visibility_For_Each_ListBox(ctrl, i)
                    
                    ' ///////////////////////////////////////////////////////////////////////////////////////////////////
                    ' *** NOTE TO SELF: I MAY NEED TO APPEND A 'T.' IN FRONT OF EACH FOR FIELD SPECIFIC QUERIES LATER ON
                    ' ///////////////////////////////////////////////////////////////////////////////////////////////////
                    If ctrl.Selected(i) Then
                        'Me.Controls!Folder1.Pages("Interests").Controls("testLabel").Caption
                        fieldsString = fieldsString & "T.[" & ctrl.Column(1, i) & "], "
                    End If
                Next i
            Exit For
        End If
    Next ctrl
    
    '////////////////////////////////////////////////////////////////////////////////////////////////
    ' FIELDS SELECTED STRING -- IF NONE ASSUME ALL -- MAKE EVERYONE VISIBLE AS WELL
    '////////////////////////////////////////////////////////////////////////////////////////////////
    
    ' If the len of the caption is = 0 then the user did not did not select satellite service
    ' Then ACT AS USUAL
    ' Else there already exists the satellite contribution from the previous tab
    '
    'End If
    '
    
    If Me.Controls!Folder1.Pages("Service").Controls("SERVICE_CHECKBOX").Value = -1 Then
            'MsgBox ("CHECKED")
            fieldsString = fieldsString & "T.[Satellite Service], "
            
        Else
            'MsgBox "Not Checked"
    End If
    
    If Len(fieldsString) = 0 Then
        
        
        fieldsString = "T.*"
        Make_All_ListBoxes_Visible
        
    Else
        ' PURGE THE LAST COMMA
        fieldsString = Left(fieldsString, Len(fieldsString) - 2)
    End If
    
    ' UPDATE A LABEL TO STORE THE FIELD INFORMATION FROM THE USER
    'Call setLabelOnPage(Me.Controls!Folder1.Pages("Interests"), "testLabel", fieldsString)
    Me.Controls!Folder1.Pages("Interests").Controls("testLabel").Caption = fieldsString
    ' -----------------------------------------
    ' FOR DEBUGGING
    ' -----------------------------------------
    'MsgBox (fieldsString)
    'Set varItm = Nothing
    'Dim i As Variant
    'For Each i In Me.Folder1.Pages(2).Controls
    '    If i.Name = "testLabel" Then
    '        MsgBox ("found it")
    '    End If
    'Next i
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
' LOAD BUTTON INSIDE FOLDER 1 "Interests" Tab
'
'
'
' GOALS:
'       1. GET SCENARIO(S) DESIRED, IF NONE THEN ASSUME ALL
'       2. GET TIMEFRAME, IF NONE THEN ASSUME ALL
'       3. GET FIELDS DESIRED, IF NONE, THEN ASSUME ALL
'           3a. MAKE COMBOBOXES VISIBLE FOR FIELDS DESIRED
'
'
'
'
'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub loadBtn_Click()
    
    ' GET THE PAGE
    Dim selectedPage As Page
    Dim pageIter As Page
    Dim ctrl As Control
    Dim varItm As Variant
    Set selectedPage = Me.Controls!Folder1.Pages("Interests")
    
    ' STRINGS
    Dim scenarioString As String
    Dim timeframeString As String
    Dim fieldsString As String
    
    ' INTERESTS TAB -- TURN ON APPROPRIATE
    For Each ctrl In Me.Controls!Folder1.Pages("Interests").Controls
        If ctrl.name = "fields_lb" Then
    
                
                For i = 0 To ctrl.ListCount - 1
                    ' *** ************  ******************************************************************************
                    ' *** THIS CHECKS WHICH ONES ARE ON/OFF, AND FLICKERS VISIBILITY AS NEEDED
                    ' *** ************  ******************************************************************************
                    Call Determine_Visibility_For_Each_ListBox(ctrl, i)
                    
                    ' *** ************  *********************************************************************************
                    ' *** NOTE TO SELF: I MAY NEED TO APPEND A 'T.' IN FRONT OF EACH FOR FIELD SPECIFIC QUERIES LATER ON
                    ' *** ************  *********************************************************************************
                    If ctrl.Selected(i) Then
                        fieldsString = fieldsString & "T.[" & ctrl.Column(1, i) & "], "
                    End If
                Next i
            Exit For
        End If
    Next ctrl
    
    ' FIELDS SELECTED STRING -- IF NONE ASSUME ALL -- MAKE EVERYONE VISIBLE AS WELL
    If Len(fieldsString) = 0 Then
        fieldsString = "T.*"
        Make_All_ListBoxes_Visible
    Else
        ' PURGE THE LAST COMMA
        fieldsString = Left(fieldsString, Len(fieldsString) - 2)
    End If
    
    ' UPDATE A LABEL TO STORE THE FIELD INFORMATION FROM THE USER
    Call setLabelOnPage(Me.Controls!Folder1.Pages("Interests"), "testLabel", fieldsString)
    'Me.Controls!Folder1.Pages("Interests").Controls("testLabel").Caption = fieldsString
    ' -----------------------------------------
    ' FOR DEBUGGING
    ' -----------------------------------------
    'MsgBox (fieldsString)
    'Set varItm = Nothing
    'Dim i As Variant
    'For Each i In Me.Folder1.Pages(2).Controls
    '    If i.Name = "testLabel" Then
    '        MsgBox ("found it")
    '    End If
    'Next i
    
    

End Sub

'////////////////////////////////////////////////////////////////////////////////////////////////
' DetermineRowSourceByScenario()
'
'
'
' GOALS:
'       1. UPDATE THE DISTINCT ITEMS IN EACH FIELD BOX IN
'          FOLDER2 BASED ON INITIAL SCENARIO CHOSEN.
'
'
'
'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub DetermineRowSourceByScenario(scenarioTableString As String)
    
    Dim sourceString As String

    ' TAB 1 VISIBILITY CONTROLS
    Dim c1 As Control
    For Each c1 In Me.Controls!Folder2.Pages("Unit").Controls
        If TypeName(c1) = "ListBox" Then
            sourceString = "SELECT T.[" & c1.name & "] FROM " & scenarioTableString & " AS T GROUP BY T.[" & c1.name & "]"
            c1.RowSource = sourceString
            'MsgBox (sourceString)
            sourceString = ""
        End If
    Next c1

    ' TAB 2 VISIBILITY CONTROLS
    Dim c2 As Control
    For Each c2 In Me.Controls!Folder2.Pages("Communication").Controls
        If TypeName(c2) = "ListBox" Then
            sourceString = "SELECT T.[" & c2.name & "] FROM " & scenarioTableString & " AS T GROUP BY T.[" & c2.name & "]"
            c2.RowSource = sourceString
            'MsgBox (sourceString)
            sourceString = ""
        End If
    Next c2
    
    ' TAB 3 VISIBILITY CONTROLS
    Dim c3 As Control
    For Each c3 In Me.Controls!Folder2.Pages("Misc.").Controls
        If TypeName(c3) = "ListBox" Then
            sourceString = "SELECT T.[" & c3.name & "] FROM " & scenarioTableString & " AS T GROUP BY T.[" & c3.name & "]"
            c3.RowSource = sourceString
            'MsgBox (sourceString)
            sourceString = ""
        End If
    Next c3
End Sub

'////////////////////////////////////////////////////////////////////////////////////////////////
' SCENARIO_LB_AfterUpdate() INSIDE FOLDER 1 "Scenario" Tab
'
'
'
' GOALS:
'       1. Update the entire form dynamically based on the scenario chosen.
'
'
'
'
'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub SCENARIO_LB_AfterUpdate()
    '///////////////////////////////////////////////////////////////
    ' This is literally how you get a selected item from a listbox
    '///////////////////////////////////////////////////////////////
    Dim selected_scenario As String
    For Each ctrl In Me.Folder1.Pages("Scenario").Controls
        If ctrl.name = "scenario_lb" Then
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
    
    'MsgBox (selected_scenario)

   '///////////////////////////////////////////////////////////////
   ' Determine what to update dynamically
   '///////////////////////////////////////////////////////////////
   Select Case selected_scenario

   Case "1"
   
        '//////////////////////////////////
        ' UPDATE THE TIMEFRAMES_LB ITEMS
        '//////////////////////////////////
        Dim c As Control
        For Each c In Me.Folder1.Pages("Timeframe").Controls
            If c.name = "TIMEFRAME_LB" Then
                c.RowSourceType = "Value List"
                c.RowSource = "1;"
            End If
        Next c
        
        
        '//////////////////////////////////
        ' UPDATE THE TIMEFRAMES_LB ITEMS
        '//////////////////////////////////
        Dim c11 As Control
        For Each c11 In Me.Folder1.Pages("Service").Controls
            If c11.name = "SERVICE_LB" Then
                'c.RowSourceType = "Table/Query"
                c11.RowSource = "SELECT DISTINCT (T.[Satellite Service]) FROM SC1_INPUT_DATA AS T; "
            End If
        Next c11
        
        
        '//////////////////////////////////
        ' UPDATE THE UNIT/COMMUNICATION/MISC
        '//////////////////////////////////
        DetermineRowSourceByScenario ("SC1_INPUT_DATA")
        
   Case "3"
      '//////////////////////////////////
      ' UPDATE THE TIMEFRAMES_LB ITEMS
      '//////////////////////////////////
       Dim c2 As Control
        For Each c2 In Me.Folder1.Pages("Timeframe").Controls
            If c2.name = "TIMEFRAME_LB" Then
                c2.RowSourceType = "Value List"
                c2.RowSource = "1;2;3;4;"
            End If
        Next c2
        
        
        
        '//////////////////////////////////
        ' UPDATE THE TIMEFRAMES_LB ITEMS
        '//////////////////////////////////
        Dim c22 As Control
        For Each c22 In Me.Folder1.Pages("Service").Controls
            If c22.name = "SERVICE_LB" Then
                'c.RowSourceType = "Table/Query"
                c22.RowSource = "SELECT DISTINCT (T.[Satellite Service]) FROM SC3_INPUT_DATA AS T; "
            End If
        Next c22
        
        
        '//////////////////////////////////
        ' UPDATE THE UNIT/COMMUNICATION/MISC
        '//////////////////////////////////
        DetermineRowSourceByScenario ("SC3_INPUT_DATA")
   Case "7"
        '//////////////////////////////////
        ' UPDATE THE TIMEFRAMES_LB ITEMS
        '//////////////////////////////////
        Dim c3 As Control
        For Each c3 In Me.Folder1.Pages("Timeframe").Controls
            If c3.name = "TIMEFRAME_LB" Then
                c3.RowSourceType = "Value List"
                c3.RowSource = "1;2;3;4;5;"
            End If
        Next c3
        
        
        '//////////////////////////////////
        ' UPDATE THE TIMEFRAMES_LB ITEMS
        '//////////////////////////////////
        Dim c33 As Control
        For Each c33 In Me.Folder1.Pages("Service").Controls
            If c33.name = "SERVICE_LB" Then
                'c.RowSourceType = "Table/Query"
                c33.RowSource = "SELECT DISTINCT (T.[Satellite Service]) FROM SC7_INPUT_DATA AS T; "
            End If
        Next c33
        
        '//////////////////////////////////
        ' UPDATE THE UNIT/COMMUNICATION/MISC
        '//////////////////////////////////
        DetermineRowSourceByScenario ("SC7_INPUT_DATA")
   Case Else
      MsgBox ("Error: You selected an invalid option. Please try again.")
    End Select
End Sub

'////////////////////////////////////////////////////////////////////
' SCENARIO_LB_MouseUp SUBROUTINE
'
'
'
' GOALS:
'       1. SHOULD UPDATE THE TIMEFRAME DROPDOWN MENU
'       2. SHOULD UPDATE EACH OF THE `WHERE` SPECS in the Unit/Communication/Misc. Tabs in Folder2 with that Scenario's specific fields.
'
'
'
'
'////////////////////////////////////////////////////////////////////


'////////////////////////////////////////////////////////////////////
' searchBtn_Click SUBROUTINE
'
'
'
' GOALS:
'       1. GET THE USERS FIELDS DESIRED
'       2. GET ALL VALUES THE USER WANTS PER FIELD
'       3. CONSTRUCT A SQL QUERY FOR THAT SPECIFIC
'       4. EXECUTE THE SQL INTO ANOTHER WINDOW / INTO THE SPLIT VIEW
'
'
'
'
'////////////////////////////////////////////////////////////////////
Private Sub searchBtn_Click()


    strSQL = Build_SQL_String
    
    '/////////////////////////////////////////////////////////////
    ' EXECUTE SQL STRING IN ANOTHER WINDOW
    '/////////////////////////////////////////////////////////////
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Set db = CurrentDb
    Set qdf = db.QueryDefs("queryQ")
    qdf.SQL = strSQL                     ' BUILDS THE EMPTY 'queryQ' QUERY INTO WHAT WE WANT
    'DoCmd.OpenQuery "queryQ"
    'DoCmd.Close Form1, Me.name

    'DoCmd.Close acForm, "EXAMPLE_FORM"
    
    DoCmd.OpenForm "formQ"
    Forms!formQ!graphQ.RowSource = "select * from queryQ"

    Set qdf = Nothing                    ' CLEARS THE NEWLY BUILT 'queryQ' QUERY SO WE WANT DO THIS PROCESS AGAIN WITH FRESH SELECTIONS
    Set db = Nothing

    
    ' ////////////////////////////////////////////////////////////
    ' BUILD A DYNAMIC REPORT FOR THE USERS SELECTION
    ' ////////////////////////////////////////////////////////////
    'CreateDynamicReport (strSQL)
    
    
    ' This works btw!
    'DoCmd.OpenReport "reportQ", acViewPreview, "queryQ"
    
    '////////////////////////////////////////////////////////////////////
    ' R&D
    '////////////////////////////////////////////////////////////////////
    
    
    'Me.RecordSource = "queryQ"
    'https://social.msdn.microsoft.com/Forums/en-US/ac0a7102-86bc-48e0-b94b-eebc035ba0b3/access-vba-export-to-specific-excel-worksheets?forum=accessdev
    'DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, "queryQ", "C:\Users\1517766115.CIV\Desktop\QUERY_RESULTS.XLSB"
    
    'Dim dbs As Database

    'Set dbs = CurrentDb
    
    
    'Set rsQuery = dbs.OpenRecordset("Table1 Query")
    
    'Set excelApp = CreateObject("Excel.application", "")
    'excelApp.Visible = True
    'Set targetWorkbook = excelApp.workbooks.Open("C:\Book1.xlsx")
    'targetWorkbook.Worksheets("tab1").Range("A5").CopyFromRecordset rsQuery
End Sub



'///////////////////////////////////////////////////////////////////////////////////
'
'
'
'
'///////////////////////////////////////////////////////////////////////////////////
Private Sub exportBtn_Click()
    
    strSQL = Build_SQL_String
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Set db = CurrentDb
    Set qdf = db.QueryDefs("queryQ")
    qdf.SQL = strSQL                     ' BUILDS THE EMPTY 'queryQ' QUERY INTO WHAT WE WANT
    
    currentPath = CurrentProject.Path
    currentPath = currentPath & "\QUERY_RESULTS.XLSB"
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, "queryQ", currentPath
    'insert_statement = "insert into Table1(SQL_STRING) values( " & """" & strSQL & """" & ")"
    'DoCmd.RunSQL insert_statement
    MsgBox ("Exported Successfully: " & currentPath)
End Sub

Function Build_SQL_String()
        ' ////////////////////////////////////////
    ' DECLARE: SCENARIO_CHOICE
    '   FOR OBVIOUS REASONS, WE NEED TO
    '   KNOW WHAT DATA TO QUERY
    ' ////////////////////////////////////////
    Dim SCENARIO_CHOICE As String
    
    ' ////////////////////////////////////////
    ' DECLARE: SCENARIO_CHOICE
    '   WE ALSO NEED TO ERROR HANDLE
    '   TO SEE IF THEY DID INDEED SELECT
    '   SOMETHING.
    ' ////////////////////////////////////////
    Dim SCENARIO_FLAG As Boolean
    
    
    ' ////////////////////////////////////////
    ' INSTANTIATE OUR DATA MEMBERS TO NULL
    ' SO WE KNOW THEIR INITIAL STATES
    ' ////////////////////////////////////////
    SCENARIO_CHOICE = ""
    SCENARIO_FLAG = True
    
    ' ////////////////////////////////////////
    ' DETERMINE WHICH SCENARIO THEY CHOSE
    ' ////////////////////////////////////////
    Dim scenarioCtrl As Control
    For Each scenarioCtrl In Me.Controls!Folder1.Pages("Scenario").Controls
        If scenarioCtrl.name = "SCENARIO_LB" Then
            Dim varItm As Variant
            For Each varItm In scenarioCtrl.ItemsSelected
                SCENARIO_CHOICE = scenarioCtrl.ItemData(varItm)
            Next varItm
        End If
    Next scenarioCtrl
    
    '///////////////////////////////////
    ' FOR DEBUGGING
    '///////////////////////////////////
    'MsgBox(SCENARIO_CHOICE)
    
    ' ////////////////////////////////////////
    ' DECLARE: TIME_CHOICE
    '   WE CAN ONLY ALLOW 1 DAY AT A TIME
    '   BUT WE DO NEED TO KNOW WHICH DAY
    '   THE USER IS INTERESTED IN AT LEAST.
    ' ////////////////////////////////////////
    Dim TIME_CHOICE As String
    
    ' ////////////////////////////////////////
    ' DECLARE: TIME_FLAG
    '   WE ALSO NEED TO ERROR HANDLE
    '   TO SEE IF THEY DID INDEED SELECT
    '   SOMETHING.
    ' ////////////////////////////////////////
    Dim TIME_FLAG As Boolean
    
    
    ' ////////////////////////////////////////
    ' INSTANTIATE OUR DATA MEMBERS TO NULL
    ' SO WE KNOW THEIR INITIAL STATES
    ' ////////////////////////////////////////
    TIME_CHOICE = ""
    TIME_FLAG = True
    
    ' //////////////////////////////////////////////////
    ' DETERMINE WHICH TIME SLICE (DAY) THEY CHOSE
    ' /////////////////////////////////////////////////
    Dim timeCtrl As Control
    For Each varItm2 In Me.Controls!Folder1.Pages("Timeframe").Controls("TIMEFRAME_LB").ItemsSelected
        'MsgBox (Me.Controls!Folder1.Pages("Timeframe").Controls("TIMEFRAME_LB").ItemData(varItm2))
        TIME_CHOICE = Me.Controls!Folder1.Pages("Timeframe").Controls("TIMEFRAME_LB").ItemData(varItm2)
    Next varItm2

    '///////////////////////////////////
    ' FOR DEBUGGING
    '///////////////////////////////////
    'MsgBox (TIME_CHOICE)
    
    
    '/////////////////////////////////////////////////
    '///////// ERROR HANDLING SCENARIO FLAGS /////////
    '/////////////////////////////////////////////////
    If SCENARIO_CHOICE = "" Then
        SCENARIO_FLAG = False
        MsgBox ("Please select a scenario.")
        Exit Function
    End If
    

    '/////////////////////////////////////////////////
    '///////// ERROR HANDLING TIME FLAGS /////////
    '/////////////////////////////////////////////////
    If TIME_CHOICE = "" Then
        TIME_FLAG = False
        MsgBox ("Please select a day you are interested in.")
        Exit Function
    End If



    '//////////////////////////////////////////////
    ' SC1 TIME OPTIONS
    '
    '
    ' BUILD EVERY JOIN STATEMENT WE WILL NEED TO
    ' LINK TOGETHER DAYS FOR SC1 DAYS.
    '
    ' THIS IS ALSO THE OPPORTUNITY TO DETERMINE
    ' THE TABLES THE QUERY WILL NEED TO
    ' ACQUIRE APPROPRIATE FIELDS.
    '
    '//////////////////////////////////////////////

    'SC1 DAY 1
    Dim SC1_FROM_STATEMENT_DAY_1 As String
    Dim SC1_ADDED_TABLES_DAY_1 As String
    SC1_FROM_STATEMENT_DAY_1 = " FROM (SC1_INPUT_DATA AS T INNER JOIN SC1_DAY1_FROM_0_TO_1145 ON T.Order = SC1_DAY1_FROM_0_TO_1145.[Order Number]) INNER JOIN SC1_DAY1_FROM_1215_TO_2345 ON SC1_DAY1_FROM_0_TO_1145.[Order Number] = SC1_DAY1_FROM_1215_TO_2345.[Order Number]"
    'SC1_ADDED_TABLES_DAY_1 = " SC1_DAY1_FROM_0_TO_1145.* , SC1_DAY1_FROM_1215_TO_2345.* "
    SC1_ADDED_TABLES_DAY_1 = "sum(SC1_DAY1_FROM_0_TO_1145.[1/1/30 0:15]) AS 015 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 0:30]) AS 030 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 0:45]) AS 045 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 1:00]) AS 100 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 1:15]) AS 115 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 1:30]) AS 130 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 1:45]) AS 145 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 2:00]) AS 200 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 2:15]) AS 215 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 2:30]) AS 230 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 2:45]) AS 245 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 3:00]) AS 300 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 3:15]) AS 315 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 3:30]) AS 330 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 3:45]) AS 345 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 4:00]) AS 400 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 4:15]) AS 415 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 4:30]) AS 430 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 4:45]) AS 445 ,"
    SC1_ADDED_TABLES_DAY_1 = SC1_ADDED_TABLES_DAY_1 & "sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 5:00]) AS 500 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 5:15]) AS 515 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 5:30]) AS 530 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 5:45]) AS 545 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 6:00]) AS 600 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 6:15]) AS 615 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 6:30]) AS 630 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 6:45]) AS 645 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 7:00]) AS 700 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 7:15]) AS 715 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 7:30]) AS 730 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 7:45]) AS 745 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 8:00]) AS 800 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 8:15]) AS 815 ,"
    SC1_ADDED_TABLES_DAY_1 = SC1_ADDED_TABLES_DAY_1 & "sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 8:30]) AS 830 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 8:45]) AS 845 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 9:00]) AS 900 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 9:15]) AS 915 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 9:30]) AS 930 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 9:45]) AS 945 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 10:00]) AS 1000 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 10:15]) AS 1015 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 10:30]) AS 1030 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 10:45]) AS 1045 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 11:00]) AS 1100 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 11:15]) AS 1115 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 11:30]) AS 1130 ,sum( SC1_DAY1_FROM_0_TO_1145.[1/1/30 11:45]) AS 1145 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 12:15]) AS 1215 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 12:30]) AS 1230 ,"
    SC1_ADDED_TABLES_DAY_1 = SC1_ADDED_TABLES_DAY_1 & "sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 12:45]) AS 1245 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 13:00]) AS 1300 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 13:15]) AS 1315 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 13:30]) AS 1330 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 13:45]) AS 1345 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 14:00]) AS 1400 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 14:15]) AS 1415 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 14:30]) AS 1430 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 14:45]) AS 1445 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 15:00]) AS 1500 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 15:15]) AS 1515 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 15:30]) AS 1530 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 15:45]) AS 1545 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 16:00]) AS 1600 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 16:15]) AS 1615 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 16:30]) AS 1630 ,"
    SC1_ADDED_TABLES_DAY_1 = SC1_ADDED_TABLES_DAY_1 & "sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 16:45]) AS 1645 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 17:00]) AS 1700 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 17:15]) AS 1715 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 17:30]) AS 1730 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 17:45]) AS 1745 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 18:00]) AS 1800 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 18:15]) AS 1815 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 18:30]) AS 1830 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 18:45]) AS 1845 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 19:00]) AS 1900 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 19:15]) AS 1915 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 19:30]) AS 1930 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 19:45]) AS 1945 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 20:00]) AS 2000 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 20:15]) AS 2015 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 20:30]) AS 2030 ,"
    SC1_ADDED_TABLES_DAY_1 = SC1_ADDED_TABLES_DAY_1 & "sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 20:45]) AS 2045 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 21:00]) AS 2100 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 21:15]) AS 2115 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 21:30]) AS 2130 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 21:45]) AS 2145 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 22:00]) AS 2200 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 22:15]) AS 2215 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 22:30]) AS 2230 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 22:45]) AS 2245 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 23:00]) AS 2300 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 23:15]) AS 2315 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 23:30]) AS 2330 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/1/30 23:45]) AS 2345 ,sum( SC1_DAY1_FROM_1215_TO_2345.[1/2/30 0:00]) AS 000 "
    
    
    
    
    
    '//////////////////////////////////////////////
    ' SC3 TIME OPTIONS
    '
    '
    ' BUILD EVERY JOIN STATEMENT WE WILL NEED TO
    ' LINK TOGETHER DAYS FOR SC3 DAYS.
    '
    ' THIS IS ALSO THE OPPORTUNITY TO DETERMINE
    ' THE TABLES THE QUERY WILL NEED TO
    ' ACQUIRE APPROPRIATE FIELDS.
    '
    '//////////////////////////////////////////////
    
    ' SC3 DAY 1
    Dim SC3_FROM_STATEMENT_DAY_1 As String
    Dim SC3_ADDED_TABLES_DAY_1 As String
    SC3_FROM_STATEMENT_DAY_1 = " FROM (SC3_INPUT_DATA T INNER JOIN SC3_DAY1_FROM_0_TO_1145 ON T.Order = SC3_DAY1_FROM_0_TO_1145.[Order Number]) INNER JOIN SC3_DAY1_FROM_12_TO_2345 ON SC3_DAY1_FROM_0_TO_1145.[Order Number] = SC3_DAY1_FROM_12_TO_2345.[Order Number] "
    'SC3_ADDED_TABLES_DAY_1 = " SC3_DAY1_FROM_0_TO_1145.*, SC3_DAY1_FROM_12_TO_2345.* "
    SC3_ADDED_TABLES_DAY_1 = "sum(SC3_DAY1_FROM_0_TO_1145.[3/7/30 0:00]) AS 000 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 0:15]) AS 015 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 0:30]) AS 030 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 0:45]) AS 045 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 1:00]) AS 100 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 1:15]) AS 115 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 1:30]) AS 130 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 1:45]) AS 145 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 2:00]) AS 200 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 2:15]) AS 215 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 2:30]) AS 230 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 2:45]) AS 245 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 3:00]) AS 300 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 3:15]) AS 315 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 3:30]) AS 330 ,"
    SC3_ADDED_TABLES_DAY_1 = SC3_ADDED_TABLES_DAY_1 & "sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 3:45]) AS 345 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 4:00]) AS 400 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 4:15]) AS 415 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 4:30]) AS 430 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 4:45]) AS 445 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 5:00]) AS 500 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 5:15]) AS 515 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 5:30]) AS 530 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 5:45]) AS 545 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 6:00]) AS 600 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 6:15]) AS 615 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 6:30]) AS 630 ,"
    SC3_ADDED_TABLES_DAY_1 = SC3_ADDED_TABLES_DAY_1 & "sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 6:45]) AS 645 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 7:00]) AS 700 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 7:15]) AS 715 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 7:30]) AS 730 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 7:45]) AS 745 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 8:00]) AS 800 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 8:15]) AS 815 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 8:30]) AS 830 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 8:45]) AS 845 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 9:00]) AS 900 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 9:15]) AS 915 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 9:30]) AS 930 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 9:45]) AS 945 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 10:00]) AS 1000 ,"
    SC3_ADDED_TABLES_DAY_1 = SC3_ADDED_TABLES_DAY_1 & "sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 10:15]) AS 1015 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 10:30]) AS 1030 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 10:45]) AS 1045 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 11:00]) AS 1100 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 11:15]) AS 1115 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 11:30]) AS 1130 ,sum( SC3_DAY1_FROM_0_TO_1145.[3/7/30 11:45]) AS 1145 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 12:00]) AS 1200 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 12:15]) AS 1215 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 12:30]) AS 1230 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 12:45]) AS 1245 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 13:00]) AS 1300 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 13:15]) AS 1315 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 13:30]) AS 1330 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 13:45]) AS 1345 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 14:00]) AS 1400 ,"
    SC3_ADDED_TABLES_DAY_1 = SC3_ADDED_TABLES_DAY_1 & "sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 14:15]) AS 1415 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 14:30]) AS 1430 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 14:45]) AS 1445 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 15:00]) AS 1500 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 15:15]) AS 1515 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 15:30]) AS 1530 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 15:45]) AS 1545 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 16:00]) AS 1600 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 16:15]) AS 1615 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 16:30]) AS 1630 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 16:45]) AS 1645 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 17:00]) AS 1700 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 17:15]) AS 1715 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 17:30]) AS 1730 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 17:45]) AS 1745 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 18:00]) AS 1800 ,"
    SC3_ADDED_TABLES_DAY_1 = SC3_ADDED_TABLES_DAY_1 & "sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 18:15]) AS 1815 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 18:30]) AS 1830 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 18:45]) AS 1845 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 19:00]) AS 1900 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 19:15]) AS 1915 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 19:30]) AS 1930 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 19:45]) AS 1945 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 20:00]) AS 2000 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 20:15]) AS 2015 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 20:30]) AS 2030 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 20:45]) AS 2045 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 21:00]) AS 2100 ,"
    SC3_ADDED_TABLES_DAY_1 = SC3_ADDED_TABLES_DAY_1 & "sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 21:15]) AS 2115 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 21:30]) AS 2130 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 21:45]) AS 2145 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 22:00]) AS 2200 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 22:15]) AS 2215 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 22:30]) AS 2230 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 22:45]) AS 2245 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 23:00]) AS 2300 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 23:15]) AS 2315 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 23:30]) AS 2330 ,sum( SC3_DAY1_FROM_12_TO_2345.[3/7/30 23:45]) AS 2345 "
    
    
    ' SC3 DAY 2
    Dim SC3_FROM_STATEMENT_DAY_2 As String
    Dim SC3_ADDED_TABLES_DAY_2 As String
    SC3_FROM_STATEMENT_DAY_2 = " FROM SC3_INPUT_DATA AS T INNER JOIN SC3_DAY2_FROM_9_TO_2345 ON T.Order = SC3_DAY2_FROM_9_TO_2345.[Order Number] "
    'SC3_ADDED_TABLES_DAY_2 = " SC3_DAY2_FROM_9_TO_2345.* "
    SC3_ADDED_TABLES_DAY_2 = "sum(SC3_DAY2_FROM_9_TO_2345.[3/8/30 9:00])  AS 900,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 9:15])  AS 915,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 9:30])  AS 930,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 9:45])  AS 945,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 10:00])  AS 1000,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 10:15])  AS 1015,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 10:30])  AS 1030,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 10:45])  AS 1045,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 11:00])  AS 1100,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 11:15])  AS 1115,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 11:30])  AS 1130,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 11:45])  AS 1145,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 12:00])  AS 1200,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 12:15])  AS 1215,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 12:30])  AS 1230,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 12:45])  AS 1245,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 13:00])  AS 1300,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 13:15])  AS 1315,"
    SC3_ADDED_TABLES_DAY_2 = SC3_ADDED_TABLES_DAY_2 & "sum(SC3_DAY2_FROM_9_TO_2345.[3/8/30 13:30])  AS 1330,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 13:45])  AS 1345,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 14:00])  AS 1400,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 14:15])  AS 1415,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 14:30])  AS 1430,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 14:45])  AS 1445,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 15:00])  AS 1500,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 15:15])  AS 1515,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 15:30])  AS 1530,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 15:45])  AS 1545,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 16:00])  AS 1600,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 16:15])  AS 1615,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 16:30])  AS 1630,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 16:45])  AS 1645,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 17:00])  AS 1700,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 17:15])  AS 1715,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 17:30])  AS 1730,"
    SC3_ADDED_TABLES_DAY_2 = SC3_ADDED_TABLES_DAY_2 & "sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 17:45])  AS 1745,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 18:00])  AS 1800,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 18:15])  AS 1815,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 18:30])  AS 1830,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 18:45])  AS 1845,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 19:00])  AS 1900,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 19:15])  AS 1915,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 19:30])  AS 1930,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 19:45])  AS 1945,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 20:00])  AS 2000,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 20:15])  AS 2015,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 20:30])  AS 2030,"
    SC3_ADDED_TABLES_DAY_2 = SC3_ADDED_TABLES_DAY_2 & "sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 20:45])  AS 2045,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 21:00])  AS 2100,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 21:15])  AS 2115,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 21:30])  AS 2130,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 21:45])  AS 2145,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 22:00])  AS 2200,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 22:15])  AS 2215,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 22:30])  AS 2230,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 22:45])  AS 2245,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 23:00])  AS 2300,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 23:15])  AS 2315,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 23:30])  AS 2330,sum( SC3_DAY2_FROM_9_TO_2345.[3/8/30 23:45])  AS 2345"
    

    
    'SC3 DAY 3
    Dim SC3_FROM_STATEMENT_DAY_3 As String
    Dim SC3_ADDED_TABLES_DAY_3 As String
    SC3_FROM_STATEMENT_DAY_3 = " FROM SC3_INPUT_DATA T INNER JOIN (SC3_DAY3_FROM_0_TO_1145 INNER JOIN SC3_DAY3_FROM_12_2345 ON SC3_DAY3_FROM_0_TO_1145.[Order Number] = SC3_DAY3_FROM_12_2345.[Order Number]) ON T.Order = SC3_DAY3_FROM_0_TO_1145.[Order Number]"
    'SC3_ADDED_TABLES_DAY_3 = "  SC3_DAY3_FROM_0_TO_1145.*, SC3_DAY3_FROM_12_2345.*"
    SC3_ADDED_TABLES_DAY_3 = "sum(SC3_DAY3_FROM_0_TO_1145.[3/9/30 0:00]) AS 000 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 0:15]) AS 015 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 0:30]) AS 030 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 0:45]) AS 045 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 1:00]) AS 100 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 1:15]) AS 115 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 1:30]) AS 130 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 1:45]) AS 145 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 2:00]) AS 200 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 2:15]) AS 215 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 2:30]) AS 230 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 2:45]) AS 245 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 3:00]) AS 300 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 3:15]) AS 315 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 3:30]) AS 330 ,"
    SC3_ADDED_TABLES_DAY_3 = SC3_ADDED_TABLES_DAY_3 & "sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 3:45]) AS 345 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 4:00]) AS 400 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 4:15]) AS 415 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 4:30]) AS 430 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 4:45]) AS 445 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 5:00]) AS 500 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 5:15]) AS 515 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 5:30]) AS 530 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 5:45]) AS 545 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 6:00]) AS 600 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 6:15]) AS 615 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 6:30]) AS 630 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 6:45]) AS 645 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 7:00]) AS 700 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 7:15]) AS 715 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 7:30]) AS 730 ,"
    SC3_ADDED_TABLES_DAY_3 = SC3_ADDED_TABLES_DAY_3 & "sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 7:45]) AS 745 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 8:00]) AS 800 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 8:15]) AS 815 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 8:30]) AS 830 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 8:45]) AS 845 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 9:00]) AS 900 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 9:15]) AS 915 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 9:30]) AS 930 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 9:45]) AS 945 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 10:00]) AS 1000 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 10:15]) AS 1015 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 10:30]) AS 1030 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 10:45]) AS 1045 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 11:00]) AS 1100 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 11:15]) AS 1115 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 11:30]) AS 1130 ,sum( SC3_DAY3_FROM_0_TO_1145.[3/9/30 11:45]) AS 1145 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 12:00]) AS 1200, "
    SC3_ADDED_TABLES_DAY_3 = SC3_ADDED_TABLES_DAY_3 & "sum( SC3_DAY3_FROM_12_2345.[3/9/30 12:15]) AS 1215 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 12:30]) AS 1230 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 12:45]) AS 1245 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 13:00]) AS 1300 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 13:15]) AS 1315 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 13:30]) AS 1330 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 13:45]) AS 1345 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 14:00]) AS 1400 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 14:15]) AS 1415 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 14:30]) AS 1430 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 14:45]) AS 1445 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 15:00]) AS 1500 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 15:15]) AS 1515 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 15:30]) AS 1530 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 15:45]) AS 1545 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 16:00]) AS 1600 "
    'SC3_ADDED_TABLES_DAY_3 = SC3_ADDED_TABLES_DAY_3 & "sum( SC3_DAY3_FROM_12_2345.[3/9/30 16:15]) AS 1615 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 16:30]) AS 1630 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 16:45]) AS 1645 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 21:00]) AS 2100 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 21:15]) AS 2115 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 21:30]) AS 2130 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 21:45]) AS 2145 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 22:00]) AS 2200 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 22:15]) AS 2215 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 22:30]) AS 2230 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 22:45]) AS 2245 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 23:00]) AS 2300 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 23:15]) AS 2315 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 23:30]) AS 2330 ,sum( SC3_DAY3_FROM_12_2345.[3/9/30 23:45]) AS 2345 "
    
    'SC3 DAY 4
    Dim SC3_FROM_STATEMENT_DAY_4 As String
    Dim SC3_ADDED_TABLES_DAY_4 As String
    SC3_FROM_STATEMENT_DAY_4 = " FROM SC3_INPUT_DATA AS T INNER JOIN SC3_DAY4_FROM_0_TO_11 ON T.Order = SC3_DAY4_FROM_0_TO_11.[Order Number]"
    'SC3_ADDED_TABLES_DAY_4 = " SC3_DAY4_FROM_0_TO_11.* "
    SC3_ADDED_TABLES_DAY_4 = "sum(SC3_DAY4_FROM_0_TO_11.[3/10/30 0:00]) AS 000 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 0:15]) AS 015 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 0:30]) AS 030 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 0:45]) AS 045 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 1:00]) AS 100 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 1:15]) AS 115 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 1:30]) AS 130 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 1:45]) AS 145 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 2:00]) AS 200 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 2:15]) AS 215 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 2:30]) AS 230 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 2:45]) AS 245 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 3:00]) AS 300 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 3:15]) AS 315 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 3:30]) AS 330 ,"
    SC3_ADDED_TABLES_DAY_4 = SC3_ADDED_TABLES_DAY_4 & "sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 3:45]) AS 345 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 4:00]) AS 400 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 4:15]) AS 415 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 4:30]) AS 430 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 4:45]) AS 445 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 5:00]) AS 500 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 5:15]) AS 515 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 5:30]) AS 530 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 5:45]) AS 545 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 6:00]) AS 600 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 6:15]) AS 615 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 6:30]) AS 630 ,"
    SC3_ADDED_TABLES_DAY_4 = SC3_ADDED_TABLES_DAY_4 & "sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 6:45]) AS 645 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 7:00]) AS 700 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 7:15]) AS 715 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 7:30]) AS 730 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 7:45]) AS 745 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 8:00]) AS 800 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 8:15]) AS 815 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 8:30]) AS 830 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 8:45]) AS 845 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 9:00]) AS 900 ,"
    SC3_ADDED_TABLES_DAY_4 = SC3_ADDED_TABLES_DAY_4 & "sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 9:15]) AS 915 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 9:30]) AS 930 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 9:45]) AS 945 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 10:00]) AS 1000 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 10:15]) AS 1015 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 10:30]) AS 1030 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 10:45]) AS 1045 ,sum( SC3_DAY4_FROM_0_TO_11.[3/10/30 11:00]) AS 1100"
    
    

    
    
    '//////////////////////////////////////////////
    ' SC7 TIME OPTIONS
    '
    '
    ' BUILD EVERY JOIN STATEMENT WE WILL NEED TO
    ' LINK TOGETHER DAYS FOR SC7 DAYS.
    '
    ' THIS IS ALSO THE OPPORTUNITY TO DETERMINE
    ' THE TABLES THE QUERY WILL NEED TO
    ' ACQUIRE APPROPRIATE FIELDS.
    '
    '//////////////////////////////////////////////
    
    'SC7 DAY 1
    Dim SC7_FROM_STATEMENT_DAY_1 As String
    Dim SC7_ADDED_TABLES_DAY_1 As String
    SC7_FROM_STATEMENT_DAY_1 = " FROM SC7_INPUT_DATA T INNER JOIN SC7_DAY1_FROM_0615_TO_2345 ON T.Order = SC7_DAY1_FROM_0615_TO_2345.[Order Number] "
    'SC7_ADDED_TABLES_DAY_1 = " SC7_DAY1_FROM_0615_TO_2345.* "
    SC7_ADDED_TABLES_DAY_1 = "sum(SC7_DAY1_FROM_0615_TO_2345.[5/1/30 6:15]) AS 615 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 6:30]) AS 630 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 6:45]) AS 645 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 7:00]) AS 700 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 7:15]) AS 715 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 7:30]) AS 730 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 7:45]) AS 745 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 8:00]) AS 800 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 8:15]) AS 815 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 8:30]) AS 830 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 8:45]) AS 845 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 9:00]) AS 900 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 9:15]) AS 915 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 9:30]) AS 930 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 9:45]) AS 945 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 10:00]) AS 1000 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 10:15]) AS 1015 ,"
    SC7_ADDED_TABLES_DAY_1 = SC7_ADDED_TABLES_DAY_1 & "sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 10:30]) AS 1030 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 10:45]) AS 1045 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 11:00]) AS 1100 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 11:15]) AS 1115 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 11:30]) AS 1130 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 11:45]) AS 1145 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 12:00]) AS 1200 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 12:15]) AS 1215 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 12:30]) AS 1230 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 12:45]) AS 1245 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 13:00]) AS 1300 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 13:15]) AS 1315 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 13:30]) AS 1330 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 13:45]) AS 1345 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 14:00]) AS 1400 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 14:15]) AS 1415 ,"
    SC7_ADDED_TABLES_DAY_1 = SC7_ADDED_TABLES_DAY_1 & "sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 14:30]) AS 1430 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 14:45]) AS 1445 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 15:00]) AS 1500 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 15:15]) AS 1515 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 15:30]) AS 1530 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 15:45]) AS 1545 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 16:00]) AS 1600 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 16:15]) AS 1615 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 16:30]) AS 1630 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 16:45]) AS 1645 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 17:00]) AS 1700 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 17:15]) AS 1715 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 17:30]) AS 1730 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 17:45]) AS 1745 ,"
    SC7_ADDED_TABLES_DAY_1 = SC7_ADDED_TABLES_DAY_1 & "sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 18:00]) AS 1800 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 18:15]) AS 1815 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 18:30]) AS 1830 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 18:45]) AS 1845 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 19:00]) AS 1900 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 19:15]) AS 1915 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 19:30]) AS 1930 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 19:45]) AS 1945 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 20:00]) AS 2000 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 20:15]) AS 2015 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 20:30]) AS 2030 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 20:45]) AS 2045 ,"
    SC7_ADDED_TABLES_DAY_1 = SC7_ADDED_TABLES_DAY_1 & "sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 21:00]) AS 2100 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 21:15]) AS 2115 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 21:30]) AS 2130 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 21:45]) AS 2145 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 22:00]) AS 2200 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 22:15]) AS 2215 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 22:30]) AS 2230 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 22:45]) AS 2245 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 23:00]) AS 2300 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 23:15]) AS 2315 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 23:30]) AS 2330 ,sum( SC7_DAY1_FROM_0615_TO_2345.[5/1/30 23:45]) AS 2345 "
    
    
    
    
    
    'SC7 DAY 2
    Dim SC7_FROM_STATEMENT_DAY_2 As String
    Dim SC7_ADDED_TABLES_DAY_2 As String
    SC7_FROM_STATEMENT_DAY_2 = " FROM (SC7_INPUT_DATA AS T INNER JOIN SC7_DAY2_FROM_0015_TO_1145 ON T.Order = SC7_DAY2_FROM_0015_TO_1145.[Order Number]) INNER JOIN SC7_DAY2_FROM_1215_TO_2345 ON SC7_DAY2_FROM_0015_TO_1145.[Order Number] = SC7_DAY2_FROM_1215_TO_2345.[Order Number]"
    'SC7_ADDED_TABLES_DAY_2 = " SC7_DAY2_FROM_0015_TO_1145.*, SC7_DAY2_FROM_1215_TO_2345.* "
    SC7_ADDED_TABLES_DAY_2 = "sum(SC7_DAY2_FROM_0015_TO_1145.[5/2/30 0:15]) AS 015 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 0:30]) AS 030 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 0:45]) AS 045 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 1:00]) AS 100 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 1:15]) AS 115 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 1:30]) AS 130 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 1:45]) AS 145 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 2:00]) AS 200 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 2:15]) AS 215 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 2:30]) AS 230 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 2:45]) AS 245 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 3:00]) AS 300 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 3:15]) AS 315 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 3:30]) AS 330 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 3:45]) AS 345 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 4:00]) AS 400 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 4:15]) AS 415, "
    SC7_ADDED_TABLES_DAY_2 = SC7_ADDED_TABLES_DAY_2 & "sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 4:30]) AS 430 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 4:45]) AS 445 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 5:00]) AS 500 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 5:15]) AS 515 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 5:30]) AS 530 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 5:45]) AS 545 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 6:00]) AS 600 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 6:15]) AS 615 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 6:30]) AS 630 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 6:45]) AS 645 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 7:00]) AS 700 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 7:15]) AS 715 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 7:30]) AS 730 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 7:45]) AS 745 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 8:00]) AS 800 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 8:15]) AS 815 ,"
    SC7_ADDED_TABLES_DAY_2 = SC7_ADDED_TABLES_DAY_2 & "sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 8:30]) AS 830 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 8:45]) AS 845 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 9:00]) AS 900 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 9:15]) AS 915 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 9:30]) AS 930 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 9:45]) AS 945 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 10:00]) AS 1000 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 10:15]) AS 1015 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 10:30]) AS 1030 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 10:45]) AS 1045 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 11:00]) AS 1100 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 11:15]) AS 1115 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 11:30]) AS 1130 ,sum( SC7_DAY2_FROM_0015_TO_1145.[5/2/30 11:45]) AS 1145 ,"
    SC7_ADDED_TABLES_DAY_2 = SC7_ADDED_TABLES_DAY_2 & "sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 12:15]) AS 1215 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 12:30]) AS 1230 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 12:45]) AS 1245 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 13:00]) AS 1300 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 13:15]) AS 1315 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 13:30]) AS 1330 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 13:45]) AS 1345 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 14:00]) AS 1400 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 14:15]) AS 1415 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 14:30]) AS 1430 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 14:45]) AS 1445 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 15:00]) AS 1500 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 15:15]) AS 1515 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 15:30]) AS 1530 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 15:45]) AS 1545 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 16:00]) AS 1600 ,"
    SC7_ADDED_TABLES_DAY_2 = SC7_ADDED_TABLES_DAY_2 & "sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 16:15]) AS 1615 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 16:30]) AS 1630 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 16:45]) AS 1645 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 17:00]) AS 1700 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 17:15]) AS 1715 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 17:30]) AS 1730 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 17:45]) AS 1745 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 18:00]) AS 1800 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 18:15]) AS 1815 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 18:30]) AS 1830 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 18:45]) AS 1845 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 19:00]) AS 1900 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 19:15]) AS 1915 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 19:30]) AS 1930 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 19:45]) AS 1945 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 20:00]) AS 2000 ,"
    SC7_ADDED_TABLES_DAY_2 = SC7_ADDED_TABLES_DAY_2 & "sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 20:15]) AS 2015 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 20:30]) AS 2030 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 20:45]) AS 2045 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 21:00]) AS 2100 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 21:15]) AS 2115 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 21:30]) AS 2130 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 21:45]) AS 2145 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 22:00]) AS 2200 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 22:15]) AS 2215 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 22:30]) AS 2230 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 22:45]) AS 2245 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 23:00]) AS 2300 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 23:15]) AS 2315 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 23:30]) AS 2330 ,sum( SC7_DAY2_FROM_1215_TO_2345.[5/2/30 23:45]) AS 2345"
    
    
    'SC7 DAY 3
    Dim SC7_FROM_STATEMENT_DAY_3 As String
    Dim SC7_ADDED_TABLES_DAY_3 As String
    SC7_FROM_STATEMENT_DAY_3 = " FROM (SC7_INPUT_DATA AS T INNER JOIN SC7_DAY3_FROM_0_TO_1145 ON T.Order = SC7_DAY3_FROM_0_TO_1145.[Order Number]) INNER JOIN SC7_DAY3_FROM_1215_TO_2345 ON SC7_DAY3_FROM_0_TO_1145.[Order Number] = SC7_DAY3_FROM_1215_TO_2345.[Order Number]"
    'SC7_ADDED_TABLES_DAY_3 = " SC7_DAY3_FROM_0_TO_1145.*, SC7_DAY3_FROM_1215_TO_2345.* "
    SC7_ADDED_TABLES_DAY_3 = "sum(SC7_DAY3_FROM_0_TO_1145.[5/3/30 0:00]) AS 000 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 0:15]) AS 015 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 0:30]) AS 030 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 0:45]) AS 045 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 1:00]) AS 100 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 1:15]) AS 115 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 1:30]) AS 130 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 1:45]) AS 145 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 2:00]) AS 200 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 2:15]) AS 215 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 2:30]) AS 230 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 2:45]) AS 245 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 3:00]) AS 300 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 3:15]) AS 315 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 3:30]) AS 330 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 3:45]) AS 345 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 4:00]) AS 400 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 4:15]) AS 415 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 4:30]) AS 430 ,"
    SC7_ADDED_TABLES_DAY_3 = SC7_ADDED_TABLES_DAY_3 & "sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 4:45]) AS 445 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 5:00]) AS 500 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 5:15]) AS 515 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 5:30]) AS 530 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 5:45]) AS 545 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 6:00]) AS 600 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 6:15]) AS 615 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 6:30]) AS 630 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 6:45]) AS 645 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 7:00]) AS 700 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 7:15]) AS 715 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 7:30]) AS 730 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 7:45]) AS 745 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 8:00]) AS 800 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 8:15]) AS 815 ,"
    SC7_ADDED_TABLES_DAY_3 = SC7_ADDED_TABLES_DAY_3 & "sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 8:30]) AS 830 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 8:45]) AS 845 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 9:00]) AS 900 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 9:15]) AS 915 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 9:30]) AS 930 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 9:45]) AS 945 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 10:00]) AS 1000 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 10:15]) AS 1015 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 10:30]) AS 1030 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 10:45]) AS 1045 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 11:00]) AS 1100 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 11:15]) AS 1115 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 11:30]) AS 1130 ,sum( SC7_DAY3_FROM_0_TO_1145.[5/3/30 11:45]) AS 1145 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 12:15]) AS 1215 ,"
    SC7_ADDED_TABLES_DAY_3 = SC7_ADDED_TABLES_DAY_3 & "sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 12:30]) AS 1230 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 12:45]) AS 1245 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 13:00]) AS 1300 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 13:15]) AS 1315 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 13:30]) AS 1330 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 13:45]) AS 1345 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 14:00]) AS 1400 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 14:15]) AS 1415 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 14:30]) AS 1430 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 14:45]) AS 1445 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 15:00]) AS 1500 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 15:15]) AS 1515 ,"
    SC7_ADDED_TABLES_DAY_3 = SC7_ADDED_TABLES_DAY_3 & "sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 15:30]) AS 1530 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 15:45]) AS 1545 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 16:00]) AS 1600 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 16:15]) AS 1615 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 16:30]) AS 1630 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 16:45]) AS 1645 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 17:00]) AS 1700 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 17:15]) AS 1715 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 17:30]) AS 1730 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 17:45]) AS 1745 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 18:00]) AS 1800 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 18:15]) AS 1815 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 18:30]) AS 1830 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 18:45]) AS 1845 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 19:00]) AS 1900 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 19:15]) AS 1915 ,"
    SC7_ADDED_TABLES_DAY_3 = SC7_ADDED_TABLES_DAY_3 & "sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 19:30]) AS 1930 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 19:45]) AS 1945 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 20:00]) AS 2000 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 20:15]) AS 2015 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 20:30]) AS 2030 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 20:45]) AS 2045 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 21:00]) AS 2100 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 21:15]) AS 2115 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 21:30]) AS 2130 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 21:45]) AS 2145 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 22:00]) AS 2200 ,"
    SC7_ADDED_TABLES_DAY_3 = SC7_ADDED_TABLES_DAY_3 & "sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 22:15]) AS 2215 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 22:30]) AS 2230 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 22:45]) AS 2245 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 23:00]) AS 2300 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 23:15]) AS 2315 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 23:30]) AS 2330 ,sum( SC7_DAY3_FROM_1215_TO_2345.[5/3/30 23:45]) AS 2345 "
    
    
    
    
    
    'SC7 DAY 4
    Dim SC7_FROM_STATEMENT_DAY_4 As String
    Dim SC7_ADDED_TABLES_DAY_4 As String
    SC7_FROM_STATEMENT_DAY_4 = " FROM (SC7_INPUT_DATA AS T INNER JOIN SC7_DAY4_FROM_0015_TO_1145 ON T.Order = SC7_DAY4_FROM_0015_TO_1145.[Order Number]) INNER JOIN SC7_DAY4_FROM_1215_TO_2345 ON SC7_DAY4_FROM_0015_TO_1145.[Order Number] = SC7_DAY4_FROM_1215_TO_2345.[Order Number]"
    'SC7_ADDED_TABLES_DAY_4 = "  SC7_DAY4_FROM_0015_TO_1145.*, SC7_DAY4_FROM_1215_TO_2345.* "
    SC7_ADDED_TABLES_DAY_4 = "sum(SC7_DAY4_FROM_0015_TO_1145.[5/4/30 0:15]) AS 015 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 0:30]) AS 030 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 0:45]) AS 045 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 1:00]) AS 100 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 1:15]) AS 115 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 1:30]) AS 130 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 1:45]) AS 145 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 2:00]) AS 200 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 2:15]) AS 215 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 2:30]) AS 230 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 2:45]) AS 245 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 3:00]) AS 300 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 3:15]) AS 315 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 3:30]) AS 330 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 3:45]) AS 345 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 4:00]) AS 400 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 4:15]) AS 415 ,"
    SC7_ADDED_TABLES_DAY_4 = SC7_ADDED_TABLES_DAY_4 & "sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 4:30]) AS 430 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 4:45]) AS 445 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 5:00]) AS 500 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 5:15]) AS 515 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 5:30]) AS 530 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 5:45]) AS 545 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 6:00]) AS 600 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 6:15]) AS 615 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 6:30]) AS 630 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 6:45]) AS 645 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 7:00]) AS 700 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 7:15]) AS 715 ,"
    SC7_ADDED_TABLES_DAY_4 = SC7_ADDED_TABLES_DAY_4 & "sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 7:30]) AS 730 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 7:45]) AS 745 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 8:00]) AS 800 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 8:15]) AS 815 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 8:30]) AS 830 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 8:45]) AS 845 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 9:00]) AS 900 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 9:15]) AS 915 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 9:30]) AS 930 ,"
    SC7_ADDED_TABLES_DAY_4 = SC7_ADDED_TABLES_DAY_4 & "sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 9:45]) AS 945 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 10:00]) AS 1000 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 10:15]) AS 1015 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 10:30]) AS 1030 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 10:45]) AS 1045 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 11:00]) AS 1100 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 11:15]) AS 1115 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 11:30]) AS 1130 ,sum( SC7_DAY4_FROM_0015_TO_1145.[5/4/30 11:45]) AS 1145 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 12:15]) AS 1215 ,"
    SC7_ADDED_TABLES_DAY_4 = SC7_ADDED_TABLES_DAY_4 & "sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 12:30]) AS 1230 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 12:45]) AS 1245 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 13:00]) AS 1300 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 13:15]) AS 1315 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 13:30]) AS 1330 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 13:45]) AS 1345 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 14:00]) AS 1400 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 14:15]) AS 1415 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 14:30]) AS 1430 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 14:45]) AS 1445 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 15:00]) AS 1500 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 15:15]) AS 1515 ,"
    SC7_ADDED_TABLES_DAY_4 = SC7_ADDED_TABLES_DAY_4 & "sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 15:30]) AS 1530 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 15:45]) AS 1545 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 16:00]) AS 1600 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 16:15]) AS 1615 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 16:30]) AS 1630 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 16:45]) AS 1645 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 17:00]) AS 1700 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 17:15]) AS 1715 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 17:30]) AS 1730 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 17:45]) AS 1745 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 18:00]) AS 1800 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 18:15]) AS 1815 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 18:30]) AS 1830 ,"
    SC7_ADDED_TABLES_DAY_4 = SC7_ADDED_TABLES_DAY_4 & "sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 18:45]) AS 1845 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 19:00]) AS 1900 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 19:15]) AS 1915 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 19:30]) AS 1930 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 19:45]) AS 1945 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 20:00]) AS 2000 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 20:15]) AS 2015 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 20:30]) AS 2030 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 20:45]) AS 2045 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 21:00]) AS 2100 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 21:15]) AS 2115 ,"
    SC7_ADDED_TABLES_DAY_4 = SC7_ADDED_TABLES_DAY_4 & "sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 21:30]) AS 2130 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 21:45]) AS 2145 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 22:00]) AS 2200 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 22:15]) AS 2215 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 22:30]) AS 2230 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 22:45]) AS 2245 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 23:00]) AS 2300 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 23:15]) AS 2315 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 23:30]) AS 2330 ,sum( SC7_DAY4_FROM_1215_TO_2345.[5/4/30 23:45]) AS 2345 "
    
    'SC7 DAY 5
    Dim SC7_FROM_STATEMENT_DAY_5 As String
    Dim SC7_ADDED_TABLES_DAY_5 As String
    SC7_FROM_STATEMENT_DAY_5 = " FROM SC7_INPUT_DATA AS T INNER JOIN SC7_DAY5_FROM_0_TO_0600 ON T.Order = SC7_DAY5_FROM_0_TO_0600.[Order Number]"
    'SC7_ADDED_TABLES_DAY_5 = " SC7_DAY5_FROM_0_TO_0600.* "
    SC7_ADDED_TABLES_DAY_5 = "sum(SC7_DAY5_FROM_0_TO_0600.[5/5/30 0:15]) AS 015 ,sum( SC7_DAY5_FROM_0_TO_0600.[5/5/30 0:30]) AS 030 ,sum( SC7_DAY5_FROM_0_TO_0600.[5/5/30 0:45]) AS 045 ,sum( SC7_DAY5_FROM_0_TO_0600.[5/5/30 1:00]) AS 100 ,sum( SC7_DAY5_FROM_0_TO_0600.[5/5/30 1:15]) AS 115 ,sum( SC7_DAY5_FROM_0_TO_0600.[5/5/30 1:30]) AS 130 ,sum( SC7_DAY5_FROM_0_TO_0600.[5/5/30 1:45]) AS 145 ,sum( SC7_DAY5_FROM_0_TO_0600.[5/5/30 2:00]) AS 200 ,sum( SC7_DAY5_FROM_0_TO_0600.[5/5/30 2:15]) AS 215 ,sum( SC7_DAY5_FROM_0_TO_0600.[5/5/30 2:30]) AS 230 ,sum( SC7_DAY5_FROM_0_TO_0600.[5/5/30 2:45]) AS 245 ,sum( SC7_DAY5_FROM_0_TO_0600.[5/5/30 3:00]) AS 300 ,sum( SC7_DAY5_FROM_0_TO_0600.[5/5/30 3:15]) AS 315 ,"
    SC7_ADDED_TABLES_DAY_5 = SC7_ADDED_TABLES_DAY_5 & "sum( SC7_DAY5_FROM_0_TO_0600.[5/5/30 3:30]) AS 330 ,sum( SC7_DAY5_FROM_0_TO_0600.[5/5/30 3:45]) AS 345 ,sum( SC7_DAY5_FROM_0_TO_0600.[5/5/30 4:00]) AS 400 ,sum( SC7_DAY5_FROM_0_TO_0600.[5/5/30 4:15]) AS 415 ,sum( SC7_DAY5_FROM_0_TO_0600.[5/5/30 4:30]) AS 430 ,sum( SC7_DAY5_FROM_0_TO_0600.[5/5/30 4:45]) AS 445 ,sum( SC7_DAY5_FROM_0_TO_0600.[5/5/30 5:00]) AS 500 ,sum( SC7_DAY5_FROM_0_TO_0600.[5/5/30 5:15]) AS 515 ,sum( SC7_DAY5_FROM_0_TO_0600.[5/5/30 5:30]) AS 530 ,sum( SC7_DAY5_FROM_0_TO_0600.[5/5/30 5:45]) AS 545 ,sum( SC7_DAY5_FROM_0_TO_0600.[5/5/30 6:00]) AS 600 "

    
    ' /////////////////////////////////////////////////////////////////////////
    ' DECLARE: selectedFromStatement
    '   This allows for a single string to contain where the day selection
    '   will be coming from
    '
    ' /////////////////////////////////////////////////////////////////////////
    Dim selectedFromStatement As String
    
    ' /////////////////////////////////////////////////////////////////////////
    ' DECLARE: selectedAddedTables
    '   This allows for a single string to contain all the tables we will
    '   need to appropriately address the day selection above
    '
    ' /////////////////////////////////////////////////////////////////////////
    Dim selectedAddedTables As String
    
    ' /////////////////////////////////////////////////////////////////////////
    ' BASED ON THE PREVIOUS OPTIONS AVAIALBLE, DETERMINE WHICH ONE THE USER
    ' ACTUALLY WANTS.
    ' /////////////////////////////////////////////////////////////////////////
    Select Case SCENARIO_CHOICE
        Case "1"
            'MsgBox ("CHOICE1")
            Select Case TIME_CHOICE
                Case "1"
                    selectedFromStatement = SC1_FROM_STATEMENT_DAY_1
                    selectedAddedTables = SC1_ADDED_TABLES_DAY_1
                    'MsgBox ("DAY1")
                Case Else
                    MsgBox ("Error: Invalid selection. Try picking a valid day. Note: Scenario 1 contains only day 1.")
            End Select
        Case "3"
            'MsgBox ("CHOICE3")
            Select Case TIME_CHOICE
                Case "1"
                    'MsgBox ("DAY1")
                    selectedFromStatement = SC3_FROM_STATEMENT_DAY_1
                    selectedAddedTables = SC3_ADDED_TABLES_DAY_1
                Case "2"
                    'MsgBox ("DAY2")
                    selectedFromStatement = SC3_FROM_STATEMENT_DAY_2
                    selectedAddedTables = SC3_ADDED_TABLES_DAY_2
                Case "3"
                    'MsgBox ("DAY3")
                    selectedFromStatement = SC3_FROM_STATEMENT_DAY_3
                    selectedAddedTables = SC3_ADDED_TABLES_DAY_3
                Case "4"
                    'MsgBox ("DAY4")
                    selectedFromStatement = SC3_FROM_STATEMENT_DAY_4
                    selectedAddedTables = SC3_ADDED_TABLES_DAY_4
                Case Else
                    MsgBox ("Error: Invalid selection. Try picking a valid day. Note: Scenario 3 contains only days 1,2,3, and 4.")
            End Select
        Case "7"
            'MsgBox ("CHOICE7")
            Select Case TIME_CHOICE
                Case "1"
                    'MsgBox ("DAY1")
                    selectedFromStatement = SC7_FROM_STATEMENT_DAY_1
                    selectedAddedTables = SC7_ADDED_TABLES_DAY_1
                Case "2"
                    'MsgBox ("DAY2")
                    selectedFromStatement = SC7_FROM_STATEMENT_DAY_2
                    selectedAddedTables = SC7_ADDED_TABLES_DAY_2
                Case "3"
                    'MsgBox ("DAY3")
                    selectedFromStatement = SC7_FROM_STATEMENT_DAY_3
                    selectedAddedTables = SC7_ADDED_TABLES_DAY_3
                Case "4"
                    'MsgBox ("DAY4")
                    selectedFromStatement = SC7_FROM_STATEMENT_DAY_4
                    selectedAddedTables = SC7_ADDED_TABLES_DAY_4
                Case "5"
                    'MsgBox ("DAY5")
                    selectedFromStatement = SC7_FROM_STATEMENT_DAY_5
                    selectedAddedTables = SC7_ADDED_TABLES_DAY_5
                Case Else
                    MsgBox ("Error: Invalid selection. Try picking a valid day. Note: Scenario 1 contains only day 1,2,3,4, and 5.")
            End Select
        Case Else
            MsgBox ("Error: Invalid selection. Try picking a valid day.")
    End Select


    '/////////////////////////////////////////////////////
    ' DECLARE: IN_COLLECTION CONTAINER OBJECT
    '
    ' THE GOAL OF THIS CONTAINER IS TO APPEND INSIDE
    ' IT ALL OF THE CHOICES THE USER WANTS FROM THEIR
    ' SELECTIONS IN THE 'FOLDER2' TABS. EACH OF THESE
    ' WILL BE STRUCTURED AS STRINGS, THEN WRAPPED UP
    ' INSIDE OF THE `IN()` CLAUSE SEPARATED BY COMMAS.
    '
    ' THIS WILL BE APPENED TO THE END OF OUR SQL STRING
    ' PREFIXED BY THE `WHERE` CLAUSE AND SEPARATED BY
    ' `AND` STATEMENTS WITH EACH ITEMS TABLE IN PREFIXING
    ' THEM INDIVIDUALLY.
    '
    '
    ' HERE IS AN EXAMPLE
    '             SELECT T.* FROM SC3_INPUT_DATA T
    '             WHERE Table1.A  IN(1,2,3) AND
    '                   Table1.B IN(2,3,4);
    '
    '
    '/////////////////////////////////////////////////////
    Dim IN_COLLECTION As New Collection
    
    '//////////////////////////////////////////////////////////////////
    ' DECLARE: FIELDS_COLLECTION CONTAINER OBJECT
    '
    ' THE GOAL OF THIS OBJECT IS TO JUST COLLECT THE
    ' RELEVANT FIELDS THE USER IS INTERESTED IN.
    '
    ' ONE HORENDOUS LITTLE DETAIL THAT MUST BE NOTED
    ' IS THAT IF THE USER SELECTS NONE OF THE OPTIONS FROM
    ' THE FOLDER1 `INTERESTS` TAB, WE ASSUME THEY WANT ALL.
    ' THEN IN ADDITION TO THAT, WE MUST ALSO ALLOW FOR THEIR
    ' OWN CUSTOM SPECIFIS TO THAT FIELD, WHICH IS WHERE THE
    ' `IN_COLLECTION` COMES INTO PLAY.
    '
    '
    '////////////////////////////////////////////////////////////////
    Dim FIELDS_COLLECTION As New Collection
    
   
    '//////////////////////////////////////////////////////////////////////////////////////////////////
    '
    ' COLLECT VISIBILITY AND INFORMATION ON TABS
    '
    ' BY LOOPING THROUGH EACH OF FOLDER2`S TABS (UNIT {looped as c1},
    ' COMMUNICATION {looped as c2}, AND MISC. {looped as c3}) WE CAN SEE WHICH ITEMS ARE
    ' LISTBOX`S AND DETERMINE WHETHER OR NOT WE SHOULD COLLECT
    ' INFORMATION FROM THEM BASED ON IT'S VISIBILITY
    '
    ' ASSUMING THE ITEM HAS BEEN MADE VISIBLE (I.E., THE USER WANTS
    ' SOMETHING FROM IT.) GO THROUGH EACH OF IT'S ITEMS SELECTED (v1,v2,v3)
    ' AND STORE EACH OF THOSE AS COMMA SEPARATED VALUES IN A LIST, THEN OF
    ' COURSE TAKE CARE OF THE SUFFIX `,`. THIS WILL BE USED FOR CREATING EACH
    ' `IN` STATEMENT LATER
    '
    ' THIS IS ALSO A GOOD TIME TO GATHER THE FIELD NAMES THEY'RE INTERESTED IN BASED ON WHAT WAS
    ' MADE VISIBLE. SO STORE THOSE INTO THE `FIELDS_COLLECTION` CONTAINER AS WELL WHILE WE
    ' ARE LOOPING THROUGH EACH OF THE TABS.
    '
    '
    '
    '////////////////////////////////////////////////////////////////////////////////////////////////////
    
    
    Dim c0 As Control
    For Each c0 In Me.Controls!Folder1.Pages("Service").Controls
        If TypeName(c0) = "ListBox" Then
            If c0.Visible = True Then
                'MsgBox (c1.name)
                Dim tmp0 As String

                ' GET EACH ITEM SELECTED
                Dim v0 As Variant
                    For Each v0 In c0.ItemsSelected
                    
                        ' FORMAT IT
                        tmp0 = tmp0 & Chr(34) & c0.ItemData(v0) & Chr(34) & ","
                    Next v0
                    
                    ' STORE THE NAMES OF EACH FIELD
                    FIELDS_COLLECTION.Add "T.[" & c0.name & "]"
                    
                    ' IF NONE, ASSUME ALL
                    If Len(tmp0) = 0 Then
                        'tmp1 = "T.[" & c1.name & "] Like '*'"
                        
                    ' IF FIELD(s) SELECTED, TRIM THE LAST COMMA & PUSH INTO AN IN(..) STATEMENT FOR EACH FIELD WE WANT
                    Else
                        tmp0 = Left(tmp0, Len(tmp0) - 1)
                        tmp0 = "T.[" & c0.name & "] IN(" & tmp0 & ")"
                        'MsgBox (tmp0)
                        IN_COLLECTION.Add tmp0
                    End If
                    
                    'IN_COLLECTION.Add tmp1
                    'MsgBox (tmp1)
                    tmp0 = ""
            End If
        End If
    Next c0
    
    '///////////////////////////////////////////////////////////////
    '
    ' COLLECT VISIBILITY AND INFORMATION ON THE UNIT TAB
    '
    '///////////////////////////////////////////////////////////////
    Dim c1 As Control
    For Each c1 In Me.Controls!Folder2.Pages("Unit").Controls
        If TypeName(c1) = "ListBox" Then
            If c1.Visible = True Then
                'MsgBox (c1.name)
                Dim tmp1 As String

                ' GET EACH ITEM SELECTED
                Dim v1 As Variant
                    For Each v1 In c1.ItemsSelected
                    
                        ' FORMAT IT
                        tmp1 = tmp1 & Chr(34) & c1.ItemData(v1) & Chr(34) & ","
                    Next v1
                    
                    ' STORE THE NAMES OF EACH FIELD
                    FIELDS_COLLECTION.Add "T.[" & c1.name & "]"
                    
                    ' IF NONE, ASSUME ALL
                    If Len(tmp1) = 0 Then
                        'tmp1 = "T.[" & c1.name & "] Like '*'"
                        
                    ' IF FIELD(s) SELECTED, TRIM THE LAST COMMA & PUSH INTO AN IN(..) STATEMENT FOR EACH FIELD WE WANT
                    Else
                        tmp1 = Left(tmp1, Len(tmp1) - 1)
                        tmp1 = "T.[" & c1.name & "] IN(" & tmp1 & ")"
                        IN_COLLECTION.Add tmp1
                    End If
                    
                    'IN_COLLECTION.Add tmp1
                    'MsgBox (tmp1)
                    tmp1 = ""
            End If
        End If
    Next c1
    
    
    '///////////////////////////////////////////////////////////////
    '
    ' COLLECT VISIBILITY AND INFORMATION ON THE COMMUNICATION TAB
    '
    '///////////////////////////////////////////////////////////////
    Dim c2 As Control
    For Each c2 In Me.Controls!Folder2.Pages("Communication").Controls
        If TypeName(c2) = "ListBox" Then
            If c2.Visible = True Then
                'MsgBox (c2.name)
                Dim tmp2 As String

                Dim v2 As Variant
                    ' GET EACH ITEM SELECTED
                    For Each v2 In c2.ItemsSelected
                        ' FORMAT IT
                        tmp2 = tmp2 & Chr(34) & c2.ItemData(v2) & Chr(34) & ","
                    Next v2
                    
                    ' STORE THE NAMES OF EACH FIELD
                    FIELDS_COLLECTION.Add "T.[" & c2.name & "]"
                    
                    ' IF NONE, ASSUME ALL
                    If Len(tmp2) = 0 Then
                        'tmp2 = "T.[" & c2.name & "] Like '*'"
                        
                    ' IF FIELD(s) SELECTED, TRIM THE LAST COMMA & PUSH INTO AN IN(..) STATEMENT FOR EACH FIELD WE WANT
                    Else
                        tmp2 = Left(tmp2, Len(tmp2) - 1)
                        tmp2 = "T.[" & c2.name & "] IN(" & tmp2 & ")"
                        IN_COLLECTION.Add tmp2
                    End If
                    
                    'IN_COLLECTION.Add tmp2
                    'MsgBox (tmp2)
                    tmp2 = ""
            End If
        End If
    Next c2
    
    '///////////////////////////////////////////////////////
    '
    ' COLLECT VISIBILITY AND INFORMATION ON THE MISC TAB
    '
    '///////////////////////////////////////////////////////
    Dim c3 As Control
    For Each c3 In Me.Controls!Folder2.Pages("Misc.").Controls
        If TypeName(c3) = "ListBox" Then
            If c3.Visible = True Then
                
                'MsgBox (c3.name)
                Dim tmp3 As String
                
                Dim v3 As Variant
                    ' GET EACH ITEM SELECTED
                    For Each v3 In c3.ItemsSelected
                        'FORMAT IT
                        tmp3 = tmp3 & Chr(34) & c3.ItemData(v3) & Chr(34) & ","
                    Next v3
                    
                    ' IF VISIBLE ... STORE THE NAMES OF EACH FIELD
                    FIELDS_COLLECTION.Add "T.[" & c3.name & "]"
                    
                    ' IF NONE, ASSUME ALL
                    If Len(tmp3) = 0 Then
                        'tmp3 = "T.[" & c3.name & "] Like '*'"
                        
                    ' IF FIELD(s) SELECTED, TRIM THE LAST COMMA & PUSH INTO AN IN(..) STATEMENT FOR EACH FIELD WE WANT
                    Else
                        'tmp3 = Left(tmp3, Len(tmp3) - 1)
                        tmp3 = "T.[" & c3.name & "] IN(" & tmp3 & ")"
                        IN_COLLECTION.Add tmp3
                    End If
                    
                    'IN_COLLECTION.Add tmp3
                    'MsgBox (tmp3)
                    tmp3 = ""

            End If
        End If
    Next c3

    
    
    
    '/////////////////////////////////////////////////////////////
    ' THIS FEATURE WAS NEVER IMPLEMENTED DUE TO TIME CONSTRAINTS
    ' HOWEVER IT COULD BE SUPPORTED, WE JUST ASSUME THE 'AND'
    ' LOGIC AND APPEND AS FOLLOWS:  WHERE COND1 AND COND2 AND ...
    ' THIS COULD EASILY BE ALTERED TO WHERE COND1 OR COND2 OR ...
    ' A LOGIC CHECKBOX WOULD HAVE BEEN MY APPROACH.
    '/////////////////////////////////////////////////////////////
    Dim optAnd As Boolean
    Dim optOR As Boolean
    optAnd = True
    optOR = False
    
    
    '/////////////////////////////////////////////////////////////
    ' FOR DEBUGGING
    ' NOTE: GO TO 'View' -> 'Immediate Window' to view the
    ' actual console for the debugger.
    '/////////////////////////////////////////////////////////////
    ' MsgBox (selectedFromStatement)
    ' MsgBox (selectedAddedTables)
    
    
    '////////////////////////////////
    ' INSTANTIATE SQL STRING
    '////////////////////////////////
    Dim strSQL As String
    strSQL = ""
    
    '////////////////////////////////
    ' INSTANTIATE SQL STRING
    '////////////////////////////////
    Dim FIELDS_STRING As String
    FIELDS_STRING = ""
    
    '////////////////////////////////
    ' BUILD THE FIELDS STRING
    '////////////////////////////////
    For i = 1 To FIELDS_COLLECTION.Count
        FIELDS_STRING = FIELDS_STRING & FIELDS_COLLECTION.Item(i) & ", "
    Next i
    FIELDS_STRING = Left(FIELDS_STRING, Len(FIELDS_STRING) - 2)             ' CLEAN UP THE LAST ", " <-- 2 characters in length
    
    
    '/////////////////////////////////////////////////////////////////////////////////////////////
    '
    '  1. CONCATENTATE THE UNIQUE FIELDS FOR THE GRAPH
    '  2. UPDATE THE GROUP BY STATEMENT WITH THE CORRECT FIELDS (I.E., THE INITIAL SET)
    '  3. UPDATE THE FIELDS_STRING WITH ALL THE TIME FIELDS (I.E., THE SECOND SET)
    '
    '  ***APPEND TIME TABLES TO THE FIELDS STRING***
    '/////////////////////////////////////////////////////////////////////////////////////////////
    
    ' ////////////////////////////////////////////////////////////////////////////////////////////
    '
    ' 1. T.[CoE], T.[Sending Unit] --> (T.[CoE]+ " "+ T.[Sending Unit]) AS UniqueSelection
    '
    ' ////////////////////////////////////////////////////////////////////////////////////////////
    '
    ' Take what our current fields string has and turn it into the concatenation.
    '
    ' ////////////////////////////////////////////////////////////////////////////////////////////
    replacement = " + " & """" & " | " & """" & " + "
    UPDATED_FIELDS = "("
    UPDATED_FIELDS = UPDATED_FIELDS & Replace(FIELDS_STRING, ",", replacement)
    UPDATED_FIELDS = UPDATED_FIELDS & ") AS UNIQUE_SELECTION"
    
    '//////
    '2.
    '//////
    GROUP_BY_STATEMENT = "GROUP BY " & FIELDS_STRING & ";"
    
    '//////
    '3.
    '//////
    FIELDS_STRING = UPDATED_FIELDS & ", " & selectedAddedTables
    
    
    '////////////////////////////////
    ' BUILD THE SQL STRING
    '////////////////////////////////
    
    
    '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    ' IF THE USER SELECTED SPECIFICS, ACCOUNT FOR THOSE HERE
    ' BY APPENDING THEM TO A STRING WITH THE WHERE CLAUSE
    ' Ex) SELECT T.[CoE], T.[Sending Unit], ..., FROM ... WHERE T.[CoE] IN("Cyber","Fires",...) AND ...T.[Sending Unit] IN ("Air1").
    '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    If IN_COLLECTION.Count > 0 Then
        strSQL = "SELECT " & FIELDS_STRING & selectedFromStatement & " WHERE "
        
        For i = 1 To IN_COLLECTION.Count
            If optAnd Then
                strSQL = strSQL & IN_COLLECTION.Item(i) & " AND "
            Else
                strSQL = strSQL & IN_COLLECTION.Item(i) & " OR "
            End If
        Next i
        ' CLEAN UP THE LAST ' AND' <- 4 characters
        strSQL = Left(strSQL, Len(strSQL) - 4)
        strSQL = strSQL & GROUP_BY_STATEMENT
    '//////////////////////////////////////////////////
    ' ELSE THE USER DIDN'T SPECIFY ANYTHINGS, SO TAKE
    ' THE SQLSTRING 'AS IS' AND APPEND OUR FIELDS AND
    ' TABLES REQUESTED TO IT.
    '//////////////////////////////////////////////////
    Else
        strSQL = "SELECT " & FIELDS_STRING & selectedFromStatement
        strSQL = strSQL & GROUP_BY_STATEMENT
    End If
    
        
    
    
    '/////////////////////////////////////////////////////////////
    ' FOR DEBUGGING
    ' NOTE: GO TO 'View' -> 'Immediate Window' to view the
    ' actual console for the debugger.
    '/////////////////////////////////////////////////////////////
    MsgBox (strSQL)
    Debug.Print ("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
    Debug.Print (strSQL)
    Build_SQL_String = strSQL
End Function
' //////////////////////////////////////////////////////////////////////////////////
' VERY INSPIRING TUTORIAL TO GET THE DYNAMIC REPORT BUILT
'
'
'
'
' https://bytes.com/topic/access/insights/696050-create-dynamic-report-using-vba
'
'
'
'
' //////////////////////////////////////////////////////////////////////////////////
Function CreateDynamicReport(strSQL As String)

    Dim db As DAO.Database ' database object
    Dim rs As DAO.Recordset ' recordset object
    Dim fld As DAO.Field ' recordset field
    Dim txtNew As Access.TextBox ' textbox control
    Dim lblNew As Access.Label ' label control
    Dim rpt As Report ' hold report object
    Dim lngTop As Long ' holds top value of control position
    Dim lngLeft As Long ' holds left value of controls position
    Dim title As String 'holds title of report
 
     'set the title
     title = "Satcom Database Custom Query Report"
 
     ' initialise position variables
     lngLeft = 0
     lngTop = 0
 
     'Create the report
     Set rpt = CreateReport
 
     ' set properties of the Report
     With rpt
         .Width = 8500
         .RecordSource = strSQL
         .Caption = title
     End With
 
     ' Open SQL query as a recordset
     Set db = CurrentDb
     Set rs = db.OpenRecordset(strSQL)
 
     ' Create Label Title
     Set lblNew = CreateReportControl(rpt.name, acLabel, _
     acPageHeader, , "SATCOM Bandwidth Study | Report", 0, 0)
     lblNew.FontBold = True
     lblNew.FontSize = 18
     lblNew.SizeToFit
 
     ' Create corresponding label and text box controls for each field.
     For Each fld In rs.Fields
 
         ' Create new text box control and size to fit data.
         Set txtNew = CreateReportControl(rpt.name, acTextBox, _
         acDetail, , fld.name, lngLeft + 5000, lngTop)
         txtNew.SizeToFit
         txtNew.Width = txtNew.Width * 2
 
         ' Create new label control and size to fit data.
         Set lblNew = CreateReportControl(rpt.name, acLabel, acDetail, _
         txtNew.name, fld.name, lngLeft, lngTop, 5000, txtNew.Height)
         lblNew.SizeToFit
         lblNew.Width = lblNew.Width * 2
 
         ' Increment top value for next control
         lngTop = lngTop + txtNew.Height + 25
     Next

 
     ' Create datestamp in Footer
     Set lblNew = CreateReportControl(rpt.name, acLabel, _
     acPageFooter, , Now(), 0, 0)
 
     ' Create page numbering on footer
     Set txtNew = CreateReportControl(rpt.name, acTextBox, _
     acPageFooter, , "='Page ' & [Page] & ' of ' & [Pages]", rpt.Width - 1000, 0)
     txtNew.SizeToFit
 
     ' Open new report.
     DoCmd.OpenReport rpt.name, acViewPreview
 
     'reset all objects
     rs.Close
     Set rs = Nothing
     Set rpt = Nothing
     Set db = Nothing
 
End Function




















'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////
' <----------------------------- R&D ----------------------------->
'////////////////////////////////////////////////////////////////////
' <------------- EXPERIMENTAL BELOW THIS POINT -------------->
'////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////

Public Sub DisplayRecordset()
    Dim myrs As DAO.Recordset ' Create a recordset to hold the data
    Dim myExcel As New Excel.Application ' Create Excel with Early binding
    Dim mySheet As Excel.Worksheet

    Set mySheet = myExcel.workbooks.Add(1).Worksheets(1) ' Create Workbook
    Set myrs = CurrentDb.OpenRecordset(strQuery) ' Define recordset

    'Add header names
    For i = 0 To myrs.Fields.Count - 1
        mySheet.Cells(1, i + 1).Value = myrs.Fields(i).name
    Next

    'Add data to excel and make Excel visible
    mySheet.Range("A1").CopyFromRecordset myrs
    myExcel.Visible = True
End Sub


'***************************************************************'
' optAnd SUBROUTINE
'
'
'
' GOALS:
'       1. Execute AND logic
'       2. Disable OR logic
'
'
'
'
'***************************************************************'
Private Sub optAnd_Click()
    If Me.optAnd.Value = True Then
        Me.optOR.Value = False
    Else
         Me.optOR.Value = True
     End If
End Sub


'***************************************************************'
' optOR SUBROUTINE
'
'
'
' GOALS:
'       1. Execute OR logic
'       2. Disable AND logic
'
'
'
'
'***************************************************************'
 
Private Sub optOr_Click()
     If Me.optOR.Value = True Then
         Me.optAnd.Value = False
     Else
         Me.optAnd.Value = True
     End If
End Sub



Private Sub SERVICE_CHECKBOX_Click()
    
    If Me.Controls!Folder1.Pages("Service").Controls("SERVICE_CHECKBOX").Value = -1 Then
        'MsgBox ("checked")
        Me.Controls!Folder1.Pages("Service").Controls("Satellite Service").Visible = True
        
    Else
        'MsgBox ("unchecked")
        Me.Controls!Folder1.Pages("Service").Controls("Satellite Service").Visible = False
    End If
    
    
    
     ' GET THE PAGE
    Dim selectedPage As Page
    Dim pageIter As Page
    Dim ctrl As Control
    Dim varItm As Variant
    Set selectedPage = Me.Controls!Folder1.Pages("Interests")
    
    ' STRINGS
    Dim scenarioString As String
    Dim timeframeString As String
    Dim fieldsString As String
    
    ' INTERESTS TAB -- TURN ON APPROPRIATE
    For Each ctrl In Me.Controls!Folder1.Pages("Interests").Controls
        If ctrl.name = "fields_lb" Then
    
                
                For i = 0 To ctrl.ListCount - 1
                    ' ////////////////////////////////////////////////////////////////////////////////////////////////
                    ' *** THIS CHECKS WHICH ONES ARE ON/OFF, AND FLICKERS VISIBILITY AS NEEDED
                    ' ////////////////////////////////////////////////////////////////////////////////////////////////
                    Call Determine_Visibility_For_Each_ListBox(ctrl, i)
                    
                    ' ///////////////////////////////////////////////////////////////////////////////////////////////////
                    ' *** NOTE TO SELF: I MAY NEED TO APPEND A 'T.' IN FRONT OF EACH FOR FIELD SPECIFIC QUERIES LATER ON
                    ' ///////////////////////////////////////////////////////////////////////////////////////////////////
                    If ctrl.Selected(i) Then
                        'Me.Controls!Folder1.Pages("Interests").Controls("testLabel").Caption
                        fieldsString = fieldsString & "T.[" & ctrl.Column(1, i) & "], "
                    End If
                Next i
            Exit For
        End If
    Next ctrl
    
    '////////////////////////////////////////////////////////////////////////////////////////////////
    ' FIELDS SELECTED STRING -- IF NONE ASSUME ALL -- MAKE EVERYONE VISIBLE AS WELL
    '////////////////////////////////////////////////////////////////////////////////////////////////
    
    ' If the len of the caption is = 0 then the user did not did not select satellite service
    ' Then ACT AS USUAL
    ' Else there already exists the satellite contribution from the previous tab
    '
    'End If
    '
    
    If Me.Controls!Folder1.Pages("Service").Controls("SERVICE_CHECKBOX").Value = -1 Then
            'MsgBox ("CHECKED")
            fieldsString = fieldsString & "T.[Satellite Service], "
            
        Else
            'MsgBox "Not Checked"
    End If
    
    If Len(fieldsString) = 0 Then
        
        
        fieldsString = "T.*"
        Make_All_ListBoxes_Visible
        
    Else
        ' PURGE THE LAST COMMA
        fieldsString = Left(fieldsString, Len(fieldsString) - 2)
    End If
    
    ' UPDATE A LABEL TO STORE THE FIELD INFORMATION FROM THE USER
    'Call setLabelOnPage(Me.Controls!Folder1.Pages("Interests"), "testLabel", fieldsString)
    Me.Controls!Folder1.Pages("Interests").Controls("testLabel").Caption = fieldsString
    ' -----------------------------------------
    ' FOR DEBUGGING
    ' -----------------------------------------
    'MsgBox (fieldsString)
    'Set varItm = Nothing
    'Dim i As Variant
    'For Each i In Me.Folder1.Pages(2).Controls
    '    If i.Name = "testLabel" Then
    '        MsgBox ("found it")
    '    End If
    'Next i
End Sub
    

