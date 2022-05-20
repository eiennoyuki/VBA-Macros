Attribute VB_Name = "Module1"
Sub Format_FIDS()
Attribute Format_FIDS.VB_ProcData.VB_Invoke_Func = "j\n14"
'
' Format Departure FIDS
'
' Copy Departure FIDS with Headers
    ActiveSheet.PasteSpecial Format:="HTML", Link:=False, DisplayAsIcon:= _
        False, NoHTMLFormatting:=True
    Range("a1,f1,g1,h1").EntireColumn.Select
    Selection.Delete Shift:=xlToLeft
' Center all Columns
    With Columns("A:G")
       .HorizontalAlignment = xlCenter
        .AutoFit
    End With
' Sort all columns by Time
    Range("D2").CurrentRegion.Sort Range("D2"), xlAscending, Header:=xlYes
' Format all to general
    Columns("B:C").NumberFormat = "General"
' Filter non-colored flights
    Range("A1").AutoFilter Field:=6, Criteria1:=RGB(255, 0, 0), Operator:=xlFilterNoFill
' Delete No Fill Color
    Range("A2:G2000").EntireRow.Delete
' Show all
    Worksheets("Sheet1").ShowAllData
' Hide Filter
    Selection.AutoFilter
' Bold Header
    Range("A1:G1").Font.Bold = True
' Add column for assigned owner
    [H1].Value = "Assigned Owner"
    [H1].Font.Bold = True
' add rows at top
    Rows("1:7").EntireRow.Insert
' insert informational text
    [A1].Value = "List of flights to own for " & Format(Now(), "ddmmmyy")
    [A3].Value = " *RED is a critical flight - tough market & equipment"
    [A4].Value = " *YELLOW is an international flight"
    [A5].Value = " *GREEN is a flight with equipment trending high on delays"
    [A6].Value = " *BLUE is a flight with high delay trend"
' adjust column width of DEP time & Assigned Owner
    Columns("D").ColumnWidth = 10
    Columns("H").ColumnWidth = 20
' align columns A1-A6 to left
    Range("A1:A6").HorizontalAlignment = xlLeft
' Select Cell A1
    Range("H9").Select
    
End Sub
Sub Mail_Selection_Range_Outlook_Body()
Attribute Mail_Selection_Range_Outlook_Body.VB_ProcData.VB_Invoke_Func = "m\n14"

    Dim rng As Range
    Dim OutApp As Object
    Dim OutMail As Object

    Set rng = Nothing
    On Error Resume Next
    'Only the visible cells in the selection
    Set rng = Selection.SpecialCells(xlCellTypeVisible)
    'You can also use a fixed range if you want
    'Set rng = Sheets("YourSheet").Range("D4:D12").SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If rng Is Nothing Then
        MsgBox "The selection is not a range or the sheet is protected" & _
               vbNewLine & "please correct and try again.", vbOKOnly
        Exit Sub
    End If

    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    On Error Resume Next
    With OutMail
        .To = "liv.mark@hotmail.com"
        .CC = ""
        .BCC = ""
        .Subject = "List Of Flights To Own - " & Format(Now(), "ddmmmyy")
        .HTMLBody = RangetoHTML(rng)
        .Display   'or use .Display
    End With
    On Error GoTo 0

    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub
Function RangetoHTML(rng As Range)

    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    ActiveSheet.UsedRange.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function

