Attribute VB_Name = "Module1"
Public DocType As String
Public MyDocTitle As String
Public FAName As String
Public FAPPR As String
Public description As String
Public Base As String
Public TemplateURL As String
Public SaveURL As String
Public DesktopURL As String
Public Sub TestForDir()
Dim strDir As String
    strDir = Environ$("userprofile") & "\Desktop\Sharepoint Documents\"
    If Dir(strDir, vbDirectory) = "" Then
        MkDir strDir
    End If
End Sub
Public Sub SingleSpaceFormat()
Selection.WholeStory
    With Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
End Sub
'Prompt for base, name, ppr to be used for saveas file name and email subject
Public Sub Info()

'AUTO BASE TO LAX - FOR LAX USE ONLY
    Base = "LAX"

'Specify what document is to be opened (DDOpen/DDClose/ShiftReport/DIF)
    DocType = InputBox("Please select among the following document types (Case Sensative):" _
    & vbCrLf & "(o) Opening Roster" & vbCrLf & "(c) Closing Roster" & vbCrLf & "(s) Shift Report" _
    & vbCrLf & "(d) DIF Report" & vbCrLf & "(i) Incident Report" & vbCrLf & "(x) Close and exit" _
    , "Input Desired Document Type")
    
        Do Until DocType = "o" _
        Or DocType = "c" Or DocType = "s" Or DocType = "d" Or DocType = "i" Or DocType = "x"
            DocType = InputBox("ERROR: Please select among the following document types (Case Sensative):" _
            & vbCrLf & "(o) Opening Roster" & vbCrLf & "(c) Closing Roster" & vbCrLf & "(s) Shift Report" _
            & vbCrLf & "(d) DIF Report" & vbCrLf & "(i) Incident Report" & vbCrLf & "(x) Close and exit" _
            , "Input Desired Document Type")
        Loop

End Sub
'Only to be ran upon creating a new document from template
Sub AutoNew()
    Module1.Info
    TemplateURL = "https://SP_Website/Forms/"
    SaveURL = "https://SP_Website/Shared Documents/"
    
' FILE INSERT - Template files need to be in SP_Website/Forms/ folder
    If DocType = "x" Then
        Application.Quit
    ElseIf DocType = "o" Then
        Selection.InsertFile (TemplateURL & "DDOpen.docx")
'        ActiveDocument.Tables(1).AutoFitBehavior (wdAutoFitWindow)
        MyDocTitle = "LAX Opening Roster " & Format(Date, "ddmmmyy")
        ActiveDocument.Fields.Update
        Module1.SingleSpaceFormat
        
    ElseIf DocType = "c" Then
        Selection.InsertFile (TemplateURL & "DDClose.docx")
'        ActiveDocument.Tables(1).AutoFitBehavior (wdAutoFitWindow)
        MyDocTitle = "LAX Closing Roster " & Format(Date, "ddmmmyy")
        With ActiveDocument.Bookmarks("MyDate").Range
            .InsertBefore Format(Date + 1, "dddd, mmmm dd, yyyy")
        End With
        ActiveDocument.Fields.Update
        Module1.SingleSpaceFormat
        
    ElseIf DocType = "s" Then
        With ActiveDocument.PageSetup
            .Orientation = wdOrientLandscape
            .TopMargin = Application.InchesToPoints(0.5)
            .BottomMargin = Application.InchesToPoints(0.5)
            .LeftMargin = Application.InchesToPoints(0.5)
            .RightMargin = Application.InchesToPoints(0.5)
        End With
        Selection.InsertFile (TemplateURL & "ShiftReport.docx")
        ActiveDocument.Tables(1).AutoFitBehavior (wdAutoFitWindow)
        ActiveDocument.Tables(1).AllowAutoFit = True
        MyDocTitle = "LAX Shift Report " & Format(Date, "ddmmmyy")
        ActiveDocument.Fields.Update
        
    ElseIf DocType = "d" Then
        Selection.InsertFile (TemplateURL & "DIF.docx")
        FAName = InputBox("Enter FA first & last name:", "FA Name")
        FAPPR = InputBox("Enter FA 6# PPR:", "FA 6 Digit PPR")
            Do Until Len(FAPPR) = 6 And IsNumeric(FAPPR) Or FAPPR = ""
                FAPPR = InputBox("ERROR:" & vbNewLine & "Please enter a SIX(6)digit FA PPR:", "FA 6 Digit PPR")
            Loop
        MyDocTitle = Base & " DIF " & Format(Date, "ddmmmyy") & " " & FAPPR & " " & FAName
        
    ElseIf DocType = "i" Then
        Selection.InsertFile (TemplateURL & "IncidentReport.docx")
        description = InputBox("Please enter a brief description of the incident." & vbNewLine & _
                    "I.E. 'Near Miss', 'Crew Conflict', ETC.")
    MyDocTitle = Base & " Incident Report " & Format(Date, "ddmmmyy") & " " & description
    
    Else
        Selection.TypeText ("There is nothing here right now, but this would be tied to its respective file.")
    End If
    
    Selection.HomeKey unit:=wdStory
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "^pDear "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute

    If Selection.Find.Found Then
        Selection.MoveRight unit:=wdCharacter, Count:=1
        Selection.EndKey unit:=wdLine, Extend:=wdExtend
        Selection.MoveLeft unit:=wdCharacter, Count:=2, Extend:=wdExtend
        If Len(Selection.Text) > 1 Then
            MyDocTitle = MyDocTitle + " to " + Selection.Text
        End If
    End If

    With Dialogs(wdDialogFileSummaryInfo)
        .Title = MyDocTitle
        .Execute
    End With

'Disable save function for all except Shift Report
    If DocType = "s" Then
        With Dialogs(wdDialogFileSaveAs)
            .Name = SaveURL & "Shift Report/" & MyDocTitle
            .Show
        End With
    ElseIf DocType = "d" Then
        With Dialogs(wdDialogFileSaveAs)
            .Name = SaveURL & "DIF/" & MyDocTitle
            .Show
        End With
    ElseIf DocType = "i" Then
        With Dialogs(wdDialogFileSaveAs)
            .Name = SaveURL & "Incident Report/" & MyDocTitle
            .Show
        End With
'    Else
'        Module1.FileSaveAs
'        Module1.filesave
    End If
    
End Sub
'To be ran if template is no longer prompting autorun sub above
Sub ResetHasRun()

ActiveDocument.Variables("HasRun") = False

End Sub
'Copies and paste word document information, bound to CTRL + .
Sub SendDocAsMail()

'Create the directory folder for "sharepoint documents" if it doesn't exist
Module1.TestForDir

If MyDocTitle <> "" Then
    DesktopURL = Environ$("userprofile") & "\Desktop\Sharepoint Documents\" & MyDocTitle & ".docx"
ElseIf MyDocTitle = "" Then
    DesktopURL = Environ$("userprofile") & "\Desktop\Sharepoint Documents\" & ActiveDocument.Name
End If

'only save if the document is a shift report/DIF/Incident Report
    If DocType = "s" Or ActiveDocument.Name Like "*Shift*" _
    Or DocType = "d" Or ActiveDocument.Name Like "*DIF*" _
    Or DocType = "i" Or ActiveDocument.Name Like "*Incident*" Then
        Documents.Save noprompt:=True

    End If

'save a copy to the desktop so that the rest of code below can run
    ActiveDocument.SaveAs2 FileName:=DesktopURL

Dim oOutlookApp As Outlook.Application
Dim oItem As Outlook.MailItem


On Error Resume Next

'Start Outlook if it isn't running
Set oOutlookApp = GetObject(, "Outlook.Application")
    If Err <> 0 Then
        Set oOutlookApp = CreateObject("Outlook.Application")
    End If

'Create a new message
Set oItem = oOutlookApp.CreateItem(olMailItem)

'Copy the open document
    Selection.WholeStory
    Selection.Copy
    Selection.End = True

'To and subject fields
Dim DistListCon As String
Dim DistListVar As String
Dim BaseList As String

'List of all distribution lists and recipients
If DocType = "o" Or ActiveDocument.Name Like "*Opening*" Then
    DistListCon = "mark.liv@email.com"
                
ElseIf DocType = "c" Or ActiveDocument.Name Like "*Closing*" Then
    DistListCon = "mark.liv@email.com"
                               
ElseIf DocType = "s" Or ActiveDocument.Name Like "*Shift*" Then
    DistListCon = "mark.liv@email.com"
                               
ElseIf DocType = "d" Or ActiveDocument.Name Like "*DIF*" Then
    Urgent = InputBox("Is this the death of either:" & vbNewLine & _
        "1.  The FA themselves?" & vbNewLine & _
        "2.  A child of the respective FA?" & vbNewLine & _
        "3.  The spouse/partner of the respective FA?")
    Do Until Urgent = "Yes" Or Urgent = "Y" Or Urgent = "No" Or Urgent = "N" _
        Or Urgent = "yes" Or Urgent = "y" Or Urgent = "no" Or Urgent = "n"
        Urgent = InputBox("Please enter 'yes' or 'no'." & vbNewLine & _
        "Is this the death of either:" & vbNewLine & _
        "1.  The FA themselves?" & vbNewLine & _
        "2.  A child of the respective FA?" & vbNewLine & _
        "3.  The spouse/partner of the respective FA?")
    Loop
    If Urgent = "No" Or Urgent = "N" Or Urgent = "no" Or Urgent = "n" Then
        DistListCon = "mark.liv@email.com"
                        
    Else
    DistListCon = "mark.liv@email.com"
                
    oItem.Importance = 2
    End If

ElseIf DocType = "i" Or ActiveDocument.Name Like "*Incident*" Then
        DistListCon = "mark.liv@email.com"
                
End If

'Send to and CC designation
    oItem.TO = DistListCon

'    If Not Base = vbNullString Then
'    oItem.CC = BaseList
'    End If

'    oItem.BCC = ""
    If DocType = "s" Or ActiveDocument.Name Like "*Shift*" _
    Or DocType = "d" Or ActiveDocument.Name Like "*DIF*" _
    Or DocType = "i" Or ActiveDocument.Name Like "*Incident*" Then
        oItem.Subject = Mid(ActiveDocument.Name, 1, Len(ActiveDocument.Name) - 5)
    ElseIf MyDocTitle = "" Then
        oItem.Subject = Mid(ActiveDocument.Name, 1, Len(ActiveDocument.Name) - 5)
    Else
        oItem.Subject = MyDocTitle
    End If
    oItem.BodyFormat = olFormatHTML
'only send as attachment if document is shift/dif/incident report
    If DocType = "s" Or ActiveDocument.Name Like "*Shift*" _
    Or DocType = "d" Or ActiveDocument.Name Like "*DIF*" _
    Or DocType = "i" Or ActiveDocument.Name Like "*Incident*" Then
        Options.SendMailAttach = True
        oItem.Attachments.Add (ActiveDocument.FullName)
    End If

'''save a copy to the desktop so that the rest of outlook code can run
''    ActiveDocument.SaveAs2 FileName:=DesktopURL
'    oItem.Importance = 2

'Set the WordEditor
Dim objInsp As Outlook.Inspector
Dim wdEditor As Word.Document
Dim objsel As Word.Selection
Dim FYI As String
Set objInsp = oItem.GetInspector
Set wdEditor = objInsp.WordEditor
Set objsel = wdEditor.Windows(1).Selection

'set the font attributes
    'With objsel.Font
    '    .Bold = True
    '   .Size = "24"
    '    .ColorIndex = wdRed
    'End With

'FYI = ""

'Place comment "FYI" first then body then signature
'    objsel.HomeKey unit:=wdStory
'    objsel.TypeText (FYI) & vbCrLf
'    objsel.Move wdParagraph, 1
'    objsel.PasteAndFormat (wdFormatOriginalFormatting)

wdEditor.Characters(1).PasteAndFormat (wdFormatOriginalFormatting)

'Display the message
oItem.Display

'Clean up
Set oItem = Nothing
Set oOutlookApp = Nothing
Set objInsp = Nothing
Set wdEditor = Nothing


End Sub

