'This script sorts and stylizes a monthly report. This report requires a redacted top sheet to function properly.
Private Sub CommandButton1_Click()
Dim Monthstr As String
Dim Yearstr As String
Dim busDays As String
Dim defaultMonth As String
Dim didWork As Boolean

didWork = True

If month(Date) = (1) Then
    defaultMonth = MonthName(12)
Else
    defaultMonth = MonthName(month(Date) - 1)
End If



'Get Month and Year
Yearstr = Application.InputBox("Enter the year this for this report", "Data Needed", year(Date), , , , , 2)
Monthstr = Application.InputBox("Enter the month for this report", "Data Needed", defaultMonth, , , , , 2)
busDays = Application.InputBox("Enter the number of business days in this month", "Data Needed", 20, , , , , 2)

If (Yearstr = "False" Or Monthstr = "False" Or busDays = "False") Then
    MsgBox "Canceled Formatting :("
    didWork = False
    GoTo NoWork
End If

Sheets(1).Cells(1, 1) = "Badge Access Report " & Monthstr & " " & Yearstr & " (" & busDays & " Business Days)"

'Optimize Macro Speed.
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

'Order sheet names
For i = 2 To ThisWorkbook.Sheets.Count - 1
    For j = i + 1 To ThisWorkbook.Sheets.Count
        If UCase(Sheets(j).Name) < UCase(Sheets(i).Name) Then
            Sheets(j).Move before:=Sheets(i)
        End If
    Next j
Next i

'Set Headers for locations sheets.
Dim headers(5) As String
    headers(0) = "Last Name"
    headers(1) = "First name"
    headers(2) = "Personnel Subarea"
    headers(3) = "Days in Office"
    headers(4) = "% of time in Office"
    
'Create conistent sheet style
For numSheet = 2 To ThisWorkbook.Sheets.Count
    
    'Activate sheet for formating
    Worksheets(numSheet).Activate
    
    'Last Row Finder
    lRow = ThisWorkbook.Worksheets(numSheet).Cells(Rows.Count, 5).End(xlUp).Row
    
    'Paste Headers at the top and center/bold them.
    Sheets(numSheet).Range("A1:E1") = headers
    Sheets(numSheet).Range("A1:E1").Font.Bold = True
    Sheets(numSheet).Range("A:E").ColumnWidth = 20
    Sheets(numSheet).Rows("1:1").RowHeight = 30
    Sheets(numSheet).Rows("1:1").VerticalAlignment = xlCenter
    Sheets(numSheet).Rows("1:1").HorizontalAlignment = xlCenter
    
    'Scroll to start of sheet
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
    'Freeze headers to top
    Sheets(numSheet).Rows("1:1").Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
            

    
    
    'Set Column Data type
    Sheets(numSheet).Range("A2:C" & lRow).NumberFormat = "general"
    Sheets(numSheet).Range("D2:D" & lRow).NumberFormat = "0"
    Sheets(numSheet).Range("E2:E" & lRow).NumberFormat = "0.00%"
    
    'zoom for easy reading
    ActiveWindow.Zoom = 130
    
    'Sort sheet conents
    ThisWorkbook.Sheets(numSheet).Range("A:E").Sort Key1:=ThisWorkbook.Sheets(numSheet).Range("D:D"), Order1:=xlDescending, Header:=xlYes, _
                                                    Key2:=ThisWorkbook.Sheets(numSheet).Range("A:A"), Order1:=xlAscending, Header:=xlYes
                                                        
    'Colorize sheet
    For i = 1 To lRow
        If (i Mod 2 = 0) Then
            ActiveSheet.Range("A:E").Rows(i).Borders.Weight = xlThin
            ActiveSheet.Range("A:E").Rows(i).BorderAround Weight:=xlThick
            ActiveSheet.Range("A:E").Rows(i).Interior.Color = xlNone 'RGB(217, 217, 217)
        Else
            ActiveSheet.Range("A:E").Rows(i).Borders.Weight = xlThin
            ActiveSheet.Range("A:E").Rows(i).BorderAround Weight:=xlThick
            ActiveSheet.Range("A:E").Rows(i).Interior.Color = xlNone
        End If
    Next
    
    
    'Fix Sheet Names
    Select Case Sheets(numSheet).Name
        Case "NY"
             Sheets(numSheet).Name = "New York"
        Case "CincinnatiCarver"
            Sheets(numSheet).Name = "Cincinnati Carver"
        Case "ElSegundo"
            Sheets(numSheet).Name = "El Segundo"
        Case "LewisCenter"
            Sheets(numSheet).Name = "Lewis Center"
    End Select

    'Update SAP tags of Topsheet
    Sheets(1).Cells(numSheet + 1, 1) = Sheets(numSheet).Name & "-" & (lRow - 1)
    
Next

Sheets(1).Name = "Teammate Percante " & Monthstr

NoWork:
'Reset settings
Worksheets(1).Activate


lRow = ThisWorkbook.Worksheets(1).Cells(Rows.Count, 6).End(xlUp).Row
ActiveSheet.Range("A3:A" & lRow).HorizontalAlignment = xlLeft
ActiveSheet.Range("A3:F" & lRow).VerticalAlignment = xlCenter
ActiveSheet.Range("B3:F" & lRow).HorizontalAlignment = xlCenter


ActiveWindow.Zoom = 100
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

If (didWork) Then
    MsgBox "Sheets Formated :) Saving"
    'Save So template doesn't get overwritten
    ThisWorkbook.SaveAs Filename:=ThisWorkbook.Path & Application.PathSeparator & Yearstr & " " & Monthstr & " Badge Access Report", FileFormat:=51
    'Deleting button
    ActiveSheet.Shapes.Range(Array("CommandButton1")).Select
    Selection.Delete
    ActiveWorkbook.Save
End If

End Sub
