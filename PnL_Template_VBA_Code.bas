Attribute VB_Name = "Module1"
Private Sub Run_PnLs()
    ThisWorkbook.Unprotect Password:="Lock2021SUN"
    ThisWorkbook.Worksheets("Locations").Unprotect Password:="Lock2021SUN"
    Application.ScreenUpdating = False
    
    ActiveWorkbook.Model.ModelTables("Dates_").Refresh
    
    Application.Calculation = xlManual
    
    DoEvents
    
    Dim ws As Worksheet
    Dim excludedSheets As Variant
    Dim copyRange As Range
    Dim pasteRange As Range
    Dim locationsSheet As Worksheet
    Dim x As Integer
    Dim Last_Cell As Integer
    Dim myButton As Button
    Dim St_Sheets As Range
    Dim cell As Range
    Dim duplicateFound As Boolean
    Dim rangesToCopy As Variant
    Dim copyRangeStr As Variant

      On Error Resume Next

 Last_Cell = ThisWorkbook.Worksheets("Locations").Cells(Rows.Count, 5).End(xlUp).Row
    
    If Last_Cell = 2 Then
        Set St_Sheets = ThisWorkbook.Worksheets("Locations").Range("E2", "E2")
    Else
        Set St_Sheets = ThisWorkbook.Worksheets("Locations").Range("E2", ThisWorkbook.Worksheets("Locations").Cells(Last_Cell, 5))
    End If
    
    duplicateFound = False
    
    ' Check for duplicates in St_Sheets (column E)
    For Each cell In St_Sheets
        If Application.WorksheetFunction.CountIf(St_Sheets, cell.Value) > 1 Then
            duplicateFound = True
            Exit For
        End If
    Next cell
    
    ' Display message box and stop code execution if duplicates are found
    If duplicateFound Then
        MsgBox "Duplicates found in Sheet Name column E, fix it and re-run", vbCritical
        End
    End If

' Delete any button
    Set locationsSheet = ThisWorkbook.Worksheets("Locations")
        For Each myButtons In locationsSheet.Shapes
        If myButtons.Type = msoFormControl Then
            myButtons.Delete
        End If
    Next myButtons

    
    P = ThisWorkbook.Worksheets("Locations").Range("B4")
    
    ThisWorkbook.Worksheets("Locations").Activate
    ActiveSheet.Calculate
    
'Unhide sheets
    Worksheets("Start").Visible = True
    Worksheets("End").Visible = True
    Worksheets("Template P").Visible = True
    Worksheets("Template DM").Visible = True
    Worksheets("Template GLs").Visible = True
    Worksheets("Store Summary").Visible = True
    Worksheets("Expense Analysis").Visible = True
    
ThisWorkbook.Worksheets("Template P").Activate

' Define template
    If ThisWorkbook.Worksheets("Locations").Range("B7") = "Papa John's Controllable Profit" Then
    Sheets("Template P").Select
    Range("A:F,J:K").Select
    Selection.Delete Shift:=xlToLeft
    Sheets("Template DM").Select
    Range("A:F,J:K").Select
    Selection.Delete Shift:=xlToLeft
    Sheets("Store Summary").Select
    Range("H:EW,HG:LC").Select
    Selection.Delete Shift:=xlToLeft
    Sheets("Template GLs").Select
    Range("A:D,G:H").Select
    Selection.Delete Shift:=xlToLeft
    Else
    If ThisWorkbook.Worksheets("Locations").Range("B7") = "EBITDA Report" Then
    Sheets("Template P").Select
    Range("A:C,G:K").Select
    Selection.Delete Shift:=xlToLeft
    Sheets("Template DM").Select
    Range("A:C,G:K").Select
    Selection.Delete Shift:=xlToLeft
    Sheets("Store Summary").Select
    Range("H:BI,EW:LC").Select
    Selection.Delete Shift:=xlToLeft
    Sheets("Template GLs").Select
    Range("A:B,E:H").Select
    Selection.Delete Shift:=xlToLeft
    Else
    If ThisWorkbook.Worksheets("Locations").Range("B7") = "Papa John's EBITDA Report" Then
    Sheets("Template P").Select
    Range("A:I").Select
    Selection.Delete Shift:=xlToLeft
    Sheets("Template DM").Select
    Range("A:I").Select
    Selection.Delete Shift:=xlToLeft
    Sheets("Store Summary").Select
    Range("H:HG").Select
    Selection.Delete Shift:=xlToLeft
    Sheets("Template GLs").Select
    Range("A:F").Select
    Selection.Delete Shift:=xlToLeft
    Else
    Sheets("Template P").Select
    Range("D:K").Select
    Selection.Delete Shift:=xlToLeft
    Sheets("Template DM").Select
    Range("D:K").Select
    Selection.Delete Shift:=xlToLeft
    Sheets("Store Summary").Select
    Range("BI:LC").Select
    Selection.Delete Shift:=xlToLeft
    Sheets("Template GLs").Select
    Range("C:H").Select
    Selection.Delete Shift:=xlToLeft
    End If
    End If
    End If
    
' Formulas
    ThisWorkbook.Worksheets("Template P").Activate
    ActiveSheet.Range("B8:B300").Copy
    ActiveSheet.Range(Cells(8, 3), Cells(300, P + 1)).PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    ThisWorkbook.Worksheets("Template DM").Activate
    ActiveSheet.Range("B8:B300").Copy
    ActiveSheet.Range(Cells(8, 3), Cells(300, P + 1)).PasteSpecial xlPasteAll
    Application.CutCopyMode = False

' Copy Fiscal year and period
    ThisWorkbook.Worksheets("Template P").Activate
    Worksheets("Locations").Range("B1").Copy
    ActiveSheet.Cells(8, P + 1).PasteSpecial Paste:=xlPasteValues
    Worksheets("Locations").Range("B3").Copy
    ActiveSheet.Cells(9, P + 1).PasteSpecial Paste:=xlPasteValues
    ThisWorkbook.Worksheets("Template DM").Activate
    Worksheets("Locations").Range("B1").Copy
    ActiveSheet.Cells(8, P + 1).PasteSpecial Paste:=xlPasteValues
    Worksheets("Locations").Range("B3").Copy
    ActiveSheet.Cells(9, P + 1).PasteSpecial Paste:=xlPasteValues

' Create tabs based on Locations table
    Last_Cell = ThisWorkbook.Worksheets("Locations").Cells(Rows.Count, 5).End(xlUp).Row

    For x = 2 To Last_Cell

    With ThisWorkbook

    New_Name = ThisWorkbook.Worksheets("Locations").Cells(x, 5).Value
    .Sheets.Add After:=.Sheets("Start")
    ActiveSheet.Name = New_Name
    
    End With
    
    Next
    
' Select all sheets except the specified as excluded
' Change the sheet names below to exclude specific worksheets
    excludedSheets = Array("Locations", "Start", "End", "Template P", "Template DM", "Store Summary", "Template GLs", "Expense Analysis")
    
    For Each ws In ThisWorkbook.Worksheets
        If IsError(Application.Match(ws.Name, excludedSheets, 0)) Then
            ws.Select False
        End If
    
    Next ws
    
' Copy and Paste the data of Template P
    Set locationsSheet = Worksheets("Template P")
    
    ' Set the range to copy
    Set copyRange = locationsSheet.Range(locationsSheet.Cells(1, 1), locationsSheet.Cells(ActiveSheet.Rows.Count, P + 1))
    
    ' Loop through selected tabs and paste the data
    For Each ws In ActiveWindow.SelectedSheets
        Set pasteRange = ws.Cells(1, 1)
        copyRange.Copy
        pasteRange.PasteSpecial xlPasteAll
        Exit For
    Next ws
    
' Create Consolidated tab
    ThisWorkbook.Worksheets("Template DM").Activate
    Sheets(ActiveSheet.Name).Select
    Sheets(ActiveSheet.Name).Name = "Consolidated"
    Application.CutCopyMode = False
    Range("B10").Select
    ActiveWindow.FreezePanes = True

' Freeze panes and remove gridlines
    For Each ws In ThisWorkbook.Worksheets
        If IsError(Application.Match(ws.Name, excludedSheets, 0)) Then
            ws.Activate
            ws.Range("B10").Select
            ActiveWindow.FreezePanes = True
            ActiveWindow.DisplayGridlines = False
        End If
    Next ws
    
    
' Fill out Store Summary tab
   Set St_Sheets = ThisWorkbook.Worksheets("Locations").Range("E2", ThisWorkbook.Worksheets("Locations").Range("E2").End(xlDown))
   Last_Cell = ThisWorkbook.Worksheets("Locations").Cells(Rows.Count, 7).End(xlUp).Row
   St_Sheets.Copy
   Sheets("Store Summary").Select
   Cells(3, 4).Select
   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
   Range("E3", "F3").Copy
   Range(Cells(3, 5), Cells(1 + Last_Cell, 6)).Select
   Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
   Range("G3").Copy
   
    If ThisWorkbook.Worksheets("Locations").Range("B7") = "Papa John's Controllable Profit" Then
        Range(Cells(3, 7), Cells(1 + Last_Cell, 68)).Select
    Else
    If ThisWorkbook.Worksheets("Locations").Range("B7") = "EBITDA Report" Then
        Range(Cells(3, 7), Cells(1 + Last_Cell, 98)).Select
    Else
    If ThisWorkbook.Worksheets("Locations").Range("B7") = "Papa John's EBITDA Report" Then
        Range(Cells(3, 7), Cells(1 + Last_Cell, 107)).Select
    Else
        Range(Cells(3, 7), Cells(1 + Last_Cell, 58)).Select
    End If
    End If
    End If
  
   Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
   Range("D2", Range("D2").End(xlToRight)).Select
   Selection.AutoFilter
   Range("G3").Select
   
' Fill out Expense Analysis tab
   St_Sheets.Copy
   Sheets("Expense Analysis").Select
   Cells(6, 1).Select
   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
   
   Range("B6", "C6").Copy
   Range(Cells(6, 2), Cells(Last_Cell + 4, 3)).Select
   Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
   Range("D6").Copy
   Range(Cells(6, 4), Cells(Last_Cell + 4, P + 3)).Select
   Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("B2").Select
    With Selection.Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
     xlBetween, Formula1:="='Store Summary'!$G$1:" & ThisWorkbook.Worksheets("Store Summary").Range("$G$1").End(xlToRight).Address
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = ""
    .ErrorTitle = ""
    .InputMessage = ""
    .ErrorMessage = ""
    .ShowInput = True
    .ShowError = True
End With

    
' Hide all sheets and create static copy button
   ThisWorkbook.Worksheets("Locations").Activate
    Set locationsSheet = ThisWorkbook.Worksheets("Locations")
    Set myButton = ActiveSheet.Buttons.Add(102, 102, 200, 25)
    
    With myButton
        .Name = "Button" 'Esto asi dejalo
        .Caption = "Create Static Copy" 'Esto es lo que va a decir el boton
        .OnAction = "order" 'Este es el nombre de la macro que va a llamar el boton
        .Left = locationsSheet.Range("A8").Left 'Esta es la celda en la que se va a poner el boton
        .Top = locationsSheet.Range("A8").Top 'Esto igual es la celda donde se pone
    End With
    
' Hide all tabs
    For Each ws In ActiveWorkbook.Worksheets
    If ws.Name <> "Locations" Then
        ws.Visible = xlSheetHidden
    End If
    Next ws

' Return to main tab and calculate
    ThisWorkbook.Worksheets("Locations").Activate
        ThisWorkbook.Protect Password:="Lock2021SUN", Structure:=True, Windows:=False
        ThisWorkbook.Worksheets("Locations").Columns("B:E").Locked = False
        ThisWorkbook.Worksheets("Locations").Protect Password:="Lock2021SUN"
        Application.Calculation = xlAutomatic
        Calculate
        DoEvents
End Sub



