VBA-Code
========

Various Macros for Work
Sub rename_cell_j1()

    Dim rng As Range
    Dim sh As Worksheet
    Dim Cell As Object
    
    Set sh = Worksheets("Sheet1")
    
    Set rng = sh.Range("A1:Z1")
    
    '--> I've put the second line in here in order to allow for entities using the Group Currency Column
    '--> If this is the case, you will have to delete a column so that the Group Currency is in column J
    With sh
        For Each Cell In rng
            Cell = Trim(Cell)
            rng.Replace What:="Amt CumGlCoCur", Replacement:="Amount Cumulated Global Company Currency"
            rng.Replace What:="Amt CumGrpCurr", Replacement:="Amount Cumulated Global Company Currency"
        Next
    End With

End Sub


Sub Text_to_columns()

    Dim rng As Range
    Dim sh As Worksheet
    Dim Cell As Object
    
    Set sh = Worksheets("Sheet1")
    
    With sh
        Set rng = .[C2]
        Set rng = Range(rng, Cells(Rows.Count, rng.Column).End(xlUp))
    
        rng.TextToColumns Destination:=rng, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
        Tab:=False, Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), _
        Array(2, 1), Array(3, 1)), TrailingMinusNumbers:=True
        
    End With
    
    Range("C2").CurrentRegion.Sort _
        Key1:=Range("C2"), _
        Order1:=xlAscending, _
        Header:=True
    
End Sub

Sub Highlight_Cells()

    Dim rng As Range
    Dim sh As Worksheet
    Dim Cell As Object
    
    Set sh = Worksheets("Sheet1")
    
    With sh
        Set rng = .[C2]
        Set rng = Range(rng, Cells(Rows.Count, rng.Column).End(xlUp))
        For Each Cell In rng
            If Cell.Value >= 1000000 And Cell.Value <= 1099999 Then
                Cell.EntireRow.Interior.Color = 65535
            End If
        Next
    End With
    
End Sub

Sub Select_Bank_Charges()

    Dim rng As Range
    Dim sh As Worksheet
    Dim Cell As Object
    
    str1 = "BANK CHARGES"
    
    Set sh = Worksheets("Sheet1")
    
    With sh
        Set rng = .[D2]
        Set rng = Range(rng, Cells(Rows.Count, rng.Column).End(xlUp))
        For Each Cell In rng
            If Cell.Value = str1 Then
                Cell.EntireRow.Interior.Color = 65535
            End If
        Next
    End With
    
End Sub

Sub copy_and_paste_highlighted_cells()

    Dim wsI As Worksheet, wsO As Worksheet
    Dim lRow As Long, wsOlRow As Long, OutputRow As Long
    Dim copyfrom As Range
    Dim Cell As Object
    
    Set wsI = Worksheets("Sheet1")
    Set wsO = Worksheets("Sheet2")
    
    With wsI
        Set copyfrom = .[F2]
        Set copyfrom = Range(copyfrom, Cells(Rows.Count, copyfrom.Column).End(xlUp))
        For Each Cell In copyfrom
            If Cell.Interior.Color = 65535 Then
                Cell.Value = "Yellow"
            End If
        Next
    End With
    

    '~~> This is the row where the data will be written
    OutputRow = wsO.Range("A" & wsI.Rows.Count).End(xlUp).Row + 1

    With wsI
        wsOlRow = .Range("G" & .Rows.Count).End(xlUp).Row

        '~~> Remove any filters
        .AutoFilterMode = False

        '~~> Filter G on "Sick Off"

        With .Range("F1:F" & wsOlRow)
            .AutoFilter Field:=1, Criteria1:="=Yellow"
            Set copyfrom = .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow
        End With

        '~~> Remove any filters
        .AutoFilterMode = False
    End With

    If Not copyfrom Is Nothing Then
        copyfrom.Copy wsO.Rows(OutputRow)

        
    End If


End Sub

Sub copy_headers_across()
    
    Sheets("Sheet2").Range("1:1").Value = Sheets("Sheet1").Range("1:1").Value

End Sub

Sub delete_columns_tab_2()

    Dim sh As Worksheet
    Set sh = Worksheets("Sheet2")
    
    sh.Columns(1).EntireColumn.Delete
    sh.Columns(1).EntireColumn.Delete
    sh.Columns(3).EntireColumn.Delete
    sh.Columns(3).EntireColumn.Delete
    sh.Columns(3).EntireColumn.Delete
    sh.Columns(3).EntireColumn.Delete
    sh.Columns(3).EntireColumn.Delete
    sh.Columns(4).EntireColumn.Delete
    sh.Columns(4).EntireColumn.Delete

End Sub

Sub copy_tab2_to_tab3()

    Dim sh1 As Worksheet
    ActiveWorkbook.Sheets("Sheet2").Copy _
        after:=ActiveWorkbook.Sheets("Sheet2")

End Sub

Sub remove_duplicates()

    Dim sh3 As Worksheet
    Set sh3 = Worksheets("Sheet2 (2)")
    
    sh3.Columns(3).EntireColumn.Delete
    
    sh3.Range(Range("A2"), Range("B2").End(xlDown)).RemoveDuplicates Columns:=1, Header:=xlNo
    
    

End Sub

Sub add_pivot_table()

    Dim wsNew As Worksheet
    Set wsNew = Sheets.Add
    
        ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
            "Sheet2!A1:C1000", Version:=xlPivotTableVersion14). _
            CreatePivotTable TableDestination:=wsNew.Name & "!R1C1", TableName:= _
            "PivotTableName", DefaultVersion:=xlPivotTableVersion14
        
        With ActiveSheet.PivotTables("PivotTableName").PivotFields("Account Number")
            .Orientation = xlRowField
            .Position = 1
        End With
        
        ActiveSheet.PivotTables("PivotTableName").AddDataField ActiveSheet.PivotTables( _
            "PivotTableName").PivotFields("Amount Cumulated Global Company Currency"), "Sum of Currency", xlSum
            
        
End Sub

Sub copy_pivot_table_values_to_breakdown()

    Sheets("Sheet2 (2)").Range("C2:C100").Value = Sheets("Sheet5").Range("B2:B100").Value

End Sub

Sub rename_sheets()

    Sheets("Sheet1").Name = "TB 2014"
    Sheets("Sheet2").Name = "Bank Codes"
    Sheets("Sheet5").Name = "Pivot Table"
    Sheets("Sheet2 (2)").Name = "Code Breakout"
    
    Application.DisplayAlerts = False
    Sheets("Sheet3").Delete
    

End Sub

Sub format_code_breakout_tab()

    Dim rng As Range
    Dim sh As Worksheet
    Dim Cell As Object
    
    Set sh = Worksheets("Code Breakout")
    
    With sh
        .Cells.ClearFormats
    End With

End Sub

Sub Format_Whole_TB()

'--> In order for this macro to work, a/c no must be in "C", a/c name must be in "D", currency must be in "J"
'--> Make sure that the name search finds all bank charges - may need to add text to "Select_Bank_Charges" if not
'--> This macro will only work for SAP entities since the GL numbers are different for Oracle Entities
'--> The column headings are different for Oracle Entities too, I'm sure it's possible to rewrite this to accommodate Oracle though


    Call rename_cell_j1
    Call Text_to_columns
    Call Highlight_Cells
    Call Select_Bank_Charges
    Call copy_and_paste_highlighted_cells
    Call copy_headers_across
    Call delete_columns_tab_2
    Call copy_tab2_to_tab3
    Call remove_duplicates
    Call add_pivot_table
    Call copy_pivot_table_values_to_breakdown
    Call rename_sheets
    Call format_code_breakout_tab
    
End Sub


Sub Standard_Number_Format()
'
' Standard_Number_Format Macro
' Recorded 10 Nov 2014 - apply number formatting to cell
'
' Keyboard Shortcut: Ctrl+m
'
    Selection.NumberFormat = "#,##0;[Red](#,##0);-"
    With Selection
        .Font.Name = "Calibri"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub

Sub copy_and_paste_numbers()

'--> This doesn't work at the moment and I think the culprit is the .Worksheets(1) instead of .Sheets("Sheet1")

    Dim vFile As Variant
    Dim wbCopyTo As Workbook
    Dim wsCopyTo As Worksheet
    Dim wbCopyFrom As Workbook
    Dim wsCopyFrom As Worksheet

    Set wbCopyTo = ActiveWorkbook
    Set wsCopyTo = ActiveSheet
    
    '----------------------------------------------------------------------------------------
    ' Open file with data to be copied
    
        vFile = Application.GetOpenFilename("Excel Files (*.xl*)," & _
        "*.xl*", 1, "Select Excel File", "Open", False)
    
    ' If Cancel then Exit
        If TypeName(vFile) = "Boolean" Then
            Exit Sub
        Else
        Set wbCopyFrom = Workbooks.Open(vFile)
        Set wsCopyFrom = wbCopyFrom.Worksheets(1)
        End If
    
    '-----------------------------------------------------------------------------------------
    With wsCopyFrom
    ' Copy From this Place in Source Workbook:
        .Range("Z2:Z200").Copy
    End With
    
    With wsCopyTo
    ' Paste to this Place in Destination Workbook:
        wsCopyTo.Range("D5").PasteSpecial _
        Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=True, Transpose:=False
    End With
    
    
    ' Close File that was Opened
        wbCopyFrom.Close SaveChanges:=False

End Sub

Sub copy_and_paste_numbers_effort2()

    Dim wbI As Workbook, wbO As Workbook
    Dim wsI As Worksheet, wsO As Worksheet

    '~~> Source/Input Workbook
    Set wbI = ActiveWorkbook
    '~~> Set the relevant sheet from where you want to copy
    Set wsI = wbI.Sheets("Sheet1")

    '----------------------------------------------------------------------------------------
    ' Open file with data to be copied
    
    vFile = Application.GetOpenFilename("Excel Files (*.xl*)," & _
    "*.xl*", 1, "Select Excel File", "Open", False)
    
    ' If Cancel then Exit
    If TypeName(vFile) = "Boolean" Then
        Exit Sub
    Else
    Set wbO = Workbooks.Open(vFile)
    Set wsO = wbO.Sheets("Display")
    End If
    
    '-----------------------------------------------------------------------------------------


    With wbO
        '~~> Set the relevant sheet to where you want to paste
        Set wsO = wbO.Sheets("Display")

        '~~> Copy the range
        wsO.Range("Z2:Z200").Copy

        '~~> Paste it in say Cell D5. Change as applicable
        wsI.Range("D5").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
        SkipBlanks:=True, Transpose:=False
        
        '~~> Copy the range
        wsO.Range("T2:T200").Copy
        
        '~~> Paste it in say Cell D5. Change as applicable
        wsI.Range("B5").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
        SkipBlanks:=True, Transpose:=False
    End With

End Sub


Sub attempt_simple_copy_paste_multiple_columns()

'--> No need
Dim ws1 As Worksheet

'--> No need
Dim ws2 As Worksheet

'--> I'll need to set my range again, might be better just to do it this way
Dim rng1 As Range

'--> No need
Set ws1 = Sheets("Input")

'--> No need
Set ws2 = Sheets("Output")

'--> I should probably include this part
On Error Resume Next

'--> Setting my range - should probably do it this way
Set rng1 = ws1.Columns("B").SpecialCells(xlConstants)

'--> Should probably include this part
On Error GoTo 0

'--> I'll include this part too
If rng1 Is Nothing Then Exit Sub

'--> Include (why not??)
Application.ScreenUpdating = False

'--> Set your range like this, change rng2 obviously
Set rng2 = ws2.[b2]

'--> This is how he's actually copying from one sheet to another
rng1.Copy rng2
'copy column I to Output C2
rng1.Offset(0, 7).Copy rng2.Offset(0, 1)
'copy column N to Output d2
rng1.Offset(0, 12).Copy rng2.Offset(0, 2)
rng2.Offset(0, 3).Resize(rng1.Cells.Count, 1) = "Scheduled Site"

'-->Include this bit too
Application.ScreenUpdating = True

End Sub


Sub combined_copy_paste()

    Dim wbI As Workbook, wbO As Workbook
    Dim wsI As Worksheet, wsO As Worksheet

    '~~> Source/Input Workbook
    Set wbI = ActiveWorkbook
    '~~> Set the relevant sheet from where you want to copy
    Set wsI = wbI.Sheets("Sheet1")

    '----------------------------------------------------------------------------------------
    ' Open file with data to be copied
    
    vFile = Application.GetOpenFilename("Excel Files (*.xl*)," & _
    "*.xl*", 1, "Select Excel File", "Open", False)
    
    ' If Cancel then Exit
    If TypeName(vFile) = "Boolean" Then
        Exit Sub
    Else
    Set wbO = Workbooks.Open(vFile)
    Set wsO = wbO.Sheets("Display")
    End If
    
    '-----------------------------------------------------------------------------------------


'--> I'll need to set my range again, might be better just to do it this way
Dim rng1 As Range

'--> I should probably include this part
On Error Resume Next


'--> Setting my range - should probably do it this way
Set rng1 = wsO.Range("Z2:Z200").SpecialCells(xlConstants)
'--> wsO.Range("Z2:Z200")

'--> Should probably include this part
On Error GoTo 0

'--> I'll include this part too
If rng1 Is Nothing Then Exit Sub

'--> Include (why not??)
Application.ScreenUpdating = False

'--> Set your range like this, change rng2 obviously
Set rng2 = wsI.[D5]

'--> This is how he's actually copying from one sheet to another
rng1.Copy rng2

'copy column I to Output C2
'--> I need to offset input range by -6, offset output range by -2 (I'll take in the dates)
rng1.Offset(0, -6).Copy rng2.Offset(0, -2)

'copy column N to Output d2
'--> now take in the description!
rng1.Offset(0, -3).Copy rng2.Offset(0, -1)

'-->Include this bit too
Application.ScreenUpdating = True


End Sub

Sub misc()

    ' Paste to this Place in Destination Workbook:
    wsCopyTo.Range("D5").PasteSpecial Paste:=xlPasteValues, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False

End Sub
