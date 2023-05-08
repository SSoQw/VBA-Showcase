Public MadetoEnd As Boolean    'Used to check if update runs
Public Save      As Boolean    'Determines if Sheet should be saved automatically
Public numDays   As Integer    'Holds calculated number of days
Public Yearstr   As String     'Holds year of report


Sub UpdateYTD()
    
    Dim wb          As Workbook   'Shorthand
    Dim lRow        As Integer    'Holds the last row of the curremt month.
    Dim myPath      As String     'Holds folder path
    Dim myFile      As String     'Holds working file
    Dim parentFile  As String     'Holds master  file
    Dim myExtension As String     'Holds extension .xls* in our case
    Dim FldrPicker  As FileDialog 'Shorthand invokation
    Dim firstError  As Boolean    'Prevents message box spam.
    firstError = True             'Set to true becasue no errors have occured

    
    Yearstr = Application.InputBox("Enter the year this for this report", "Data Needed", Year(Date), , , , , 2)
    'Retrieve Target Folder For Monthly Reports.
    Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)
    
    With FldrPicker
        .Title = "Select A Target Folder"
        .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NextCode
        myPath = .SelectedItems(1) & "\"
    End With
    
    'In Case of Canceled Folder Dialogue.
NextCode:
    myPath = myPath
    If myPath = "" Then GoTo NoFolder
    myExtension = "*.xls*"
    myFile = Dir(myPath & myExtension)
    parentFile = myFile
    
    'Set Headers for monthly sheets.
    Dim headers(5) As String
    headers(0) = "Last Name"
    headers(1) = "First name"
    headers(2) = "Personnel Subarea"
    headers(3) = "Days in Office"
    headers(4) = "% of time in Office"
    
    'Clear contents of sheets to prevent bad or mismatched data from lingering.
     For numSheet = 2 To ThisWorkbook.Sheets.Count
        lRow = ThisWorkbook.Worksheets(numSheet).Cells(Rows.Count, 5).End(xlUp).Row
        'Clear cell contens
        Sheets(numSheet).Range("A2:E" & lRow).ClearContents
        'Paste Headers at the top.
        Sheets(numSheet).Range("A1:E1") = headers
     Next
    
    
    
    'Itterate through all Sheets in YTD workbook. For each location, we will call every month's worksheet for that location and create a single consolidated version.
    For i = 1 To ThisWorkbook.Worksheets.Count
        'MsgBox "Working on: " & ThisWorkbook.Worksheets(i).Name
        
        'We only need to update the number of the days for the report on the first page.
        If i = 1 Then
            
            Dim countDays As Integer
            Dim arrSplitStrings() As String
            
            Do While myFile <> ""
                'Set variable equal to opened workbook.
                Set wb = Workbooks.Open(Filename:=myPath & myFile)
                
                'Ensure Workbook has opened before moving on to next line of code.
                DoEvents
                wb.Activate
                
                'Updates the Number of Days Variable.
                arrSplitStrings() = Split(ActiveWorkbook.Sheets(1).Range("A1"), " ")
                arrSplitStrings(5) = Right(arrSplitStrings(5), Len(arrSplitStrings(5)) - 1)
                countDays = countDays + CInt(arrSplitStrings(5))
                
                'Save and Close Workbook.
                wb.Close SaveChanges:=True
                
                'Ensure Workbook has closed before moving on to next line of code.
                DoEvents
                
                'Get next file name.
                myFile = Dir
            Loop
            
            'Updates the A1 cell of the YTD so that it reflect the current number of office days recroded.
            numDays = countDays
            ThisWorkbook.Sheets(1).Cells(1, 1) = "Badge Access Report Year To Date 2023 (" & numDays & " Business Days)"
            
        'Otherwise we must consolidate the total number of days per person and add anyone who is new to the YTD sheet per location
        Else
            Dim employeeData()   As Variant   'Master Array.
            Dim sheetData()      As Variant   'Array for active month.
            Dim firstMonth       As Boolean   'Determines if the location is being examined for the first time this script run.
            Dim curSheet         As String    'Hold name of current sheet, used for error checking.
            Dim compSheet        As String    'Hold name of the first sheet data comes from, used for error checking.
            Dim compBook         As String    'Hold the name of the parent book that the master data came from, used for error message.
            firstMonth = True                 'Set to true for every location the first time.
            
            
            
            'Same Loop Strcuture, however its function is different.
            Do While myFile <> ""
                Set wb = Workbooks.Open(Filename:=myPath & myFile)
                DoEvents
                wb.Activate
                lRow = Worksheets(i).Cells(Rows.Count, 4).End(xlUp).Row
                'Loop through all rows for current location's month files.
                sheetData = Worksheets(i).Range("A2:D" & lRow).Value
                    If firstMonth Then
                        employeeData = sheetData
                        firstMonth = False
                        compSheet = Worksheets(i).Name
                        compBook = wb.Name
                    Else
                        Dim iRw As Integer
                        Dim jRw As Integer
                        Dim exist As Boolean
                        Dim curRow(3) As Variant
                        Dim checkRow(3) As Variant
                        curSheet = Worksheets(i).Name
                        
                        'Print error message if sheet names are not the same for a given index. Only Occurs once to prevent spam.
                        If Not (compSheet = curSheet) And firstError Then
                            firstError = False
                            MsgBox "Error in Sheet names: It seems a location has been added or removed in one of the monthly reports. The location mismatch occured for: " & vbCrLf & vbCrLf & _
                                    compBook & "!" & compSheet & vbCrLf & vbCrLf & " and " & vbCrLf & vbCrLf & wb.Name & "!" & curSheet & vbCrLf & vbCrLf & _
                                    "Please remove the inconsistency and run again.", vbExclamation, "Sheet Names Inconsistent: The generated YTD will NOT be correct."
                        End If
                        
                        
                        For iRw = 1 To UBound(sheetData, 1)
                            curRow(0) = sheetData(iRw, 1)
                            curRow(1) = sheetData(iRw, 2)
                            curRow(2) = sheetData(iRw, 3)
                            curRow(3) = sheetData(iRw, 4)
                            exist = False
                            
                            For jRw = 1 To UBound(employeeData, 1)
                                checkRow(0) = employeeData(jRw, 1)
                                checkRow(1) = employeeData(jRw, 2)
                                checkRow(2) = employeeData(jRw, 3)
                                checkRow(3) = employeeData(jRw, 4)
                                
                                'If true (the person is  already in the array) add know person's days in office to their total in the array.
                                If curRow(0) = checkRow(0) And curRow(1) = checkRow(1) Then
                                    employeeData(jRw, 4) = CInt(employeeData(jRw, 4)) + CInt(curRow(3))
                                    exist = True
                                    Exit For
                                End If
                            Next

                            'Otherwise add them to the array.
                            If Not exist Then
                                'Add new row.
                                employeeData = ReDimPreserve(employeeData, (UBound(employeeData, 1) + 1), 4)
                                'Insert data for unknown employee into the new row.
                                employeeData((UBound(employeeData, 1)), 1) = curRow(0)
                                employeeData((UBound(employeeData, 1)), 2) = curRow(1)
                                employeeData((UBound(employeeData, 1)), 3) = curRow(2)
                                employeeData((UBound(employeeData, 1)), 4) = curRow(3)
                                
                                
                            End If
                        Next
                    End If
                    
                    wb.Close SaveChanges:=True
                    
                    DoEvents
                    
                    myFile = Dir
                Loop
                
            'Add row to hold percentage of time in office to locations YTD data.
            employeeData = ReDimPreserve(employeeData, UBound(employeeData, 1), 5)
            
            'Add percentage of time in office to locations YTD data.
            For kRow = 1 To UBound(employeeData, 1)
                employeeData(kRow, 5) = employeeData(kRow, 4) / numDays
            Next
            
            'Paste updated ytd data into the locations sheet.
            Worksheets(i).Range("A2").Resize(UBound(employeeData, 1), UBound(employeeData, 2)) = employeeData
        End If
        
        'Reset to top of directory for next location.
        myFile = Dir(myPath & myExtension)
        
    Next
    
    'All Done!
    If firstError Then
        MadetoEnd = True
        Save = True
        MsgBox "YTD Updated with no errors! Generating Uniform formatting.", vbInformation, "Operation Completed"
    Else
        MsgBox "YTD had an issue when updating sheets :( Generating Uniform formatting.", vbCritical, "Operation Completed"
        MadetoEnd = True
    End If
    
NoFolder:
    If Not MadetoEnd Then
        MsgBox "No Folder selected, so no operations were performed", vbInformation, "Operation Completed"
    End If
End Sub

Private Function ReDimPreserve(MyArray As Variant, nNewFirstUBound As Long, nNewLastUBound As Long) As Variant
    Dim i, j As Long
    Dim nOldFirstUBound, nOldLastUBound, nOldFirstLBound, nOldLastLBound As Long
    Dim TempArray() As Variant

    ReDimPreserve = False
    'check if its in array first.
    If Not IsArray(MyArray) Then MsgBox "You didn't pass the function an array.", vbCritical, "No Array Detected": End
    
    'get old lBound/uBound.
    nOldFirstUBound = UBound(MyArray, 1): nOldLastUBound = UBound(MyArray, 2)
    nOldFirstLBound = LBound(MyArray, 1): nOldLastLBound = LBound(MyArray, 2)
    'create new array.
    ReDim TempArray(nOldFirstLBound To nNewFirstUBound, nOldLastLBound To nNewLastUBound)
    'loop through first.
    For i = LBound(MyArray, 1) To nNewFirstUBound
        For j = LBound(MyArray, 2) To nNewLastUBound
            'if its in range, then append to new array the same way.
            If nOldFirstUBound >= i And nOldLastUBound >= j Then
                TempArray(i, j) = MyArray(i, j)
            End If
        Next
    Next
    'return the array redimmed.
    If IsArray(TempArray) Then ReDimPreserve = TempArray
End Function

Private Sub CommandButton1_Click()

    MadetoEnd = False
    'Optimize Macro Speed.
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    'Update sheets.
    UpdateYTD
    
    'Check if anything happened
    If MadetoEnd Then
    
        'Sort the results and formatl.
        For numSheet = 2 To ThisWorkbook.Sheets.Count
        
            lRow = ThisWorkbook.Worksheets(numSheet).Cells(Rows.Count, 5).End(xlUp).Row
            
            'Format Headers
            ThisWorkbook.Worksheets(numSheet).Activate
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
    
            ActiveWindow.Zoom = 130
            
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
        Next
        
        MsgBox "All Done!", vbInformation, "Operation Completed"
        
        If Save Then
            Worksheets(1).Activate
            'Save So template doesn't get overwritten
            ThisWorkbook.SaveAs Filename:=ThisWorkbook.Path & Application.PathSeparator & Yearstr & " Year to Date Badge Access Report (" & numDays & " Business Days)", FileFormat:=51
            'Deleting button
            ActiveSheet.Shapes.Range(Array("CommandButton1")).Select
            Selection.Delete
            ActiveWorkbook.Save
        End If
        
    End If
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
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub
