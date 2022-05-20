Attribute VB_Name = "Module1"
Sub rename_folder()
Attribute rename_folder.VB_ProcData.VB_Invoke_Func = "m\n14"

Dim old_name, new_name As String

For i = 2 To Sheets(1).Range("a1").End(xlDown).Row

new_name = Left(Sheets(1).Cells(i, 1).Value, Len(Sheets(1).Cells(i, 1).Value) - Len(Sheets(1).Cells(i, 2).Value))

new_name = new_name & Sheets(1).Cells(i, 3).Value

old_name = Sheets(1).Cells(i, 1).Value
Name old_name As new_name

Next i

End Sub
