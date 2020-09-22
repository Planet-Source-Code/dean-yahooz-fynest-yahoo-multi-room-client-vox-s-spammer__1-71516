Attribute VB_Name = "listboxMOD"
'Not Coded By Yahooz Fynest these are Borrowed, Thanks whoever coded them!
'Just the Common List Functions Every Person seems to use
'They work so no need to bother to Code these Fresh.
'Also i have not used all of this function within the rest of the source yet, only some.

Public Function SaveList(dialogCommon As CommonDialog, list45 As listbox)
On Error GoTo Error_Killer
With dialogCommon
.DialogTitle = "Save List"
.Filter = "*.txt"
.ShowSave
Dim Nbr As Long
On Error Resume Next
Open .filename For Output As #1
For Nbr = 0 To list45.ListCount - 1
Print #1, list45.list(Nbr)
Next Nbr
Close #1
End With
Exit Function
Error_Killer:
Exit Function
End Function

Public Function LoadList(dialogCommon As CommonDialog, list45 As listbox)
On Error GoTo Error_Killer
With dialogCommon
.DialogTitle = "Load List"
.Filter = "All Supported Types|*.txt"
.ShowOpen
LoadListFromFile list45, .filename
End With
Error_Killer:
Exit Function
End Function

Sub LoadListFromFile(cListBox As listbox, sCurrentFile As String)
Dim sLineIn As String
On Error GoTo ErrLoadListFromFile
Open sCurrentFile For Input As #1
While Not EOF(1)
Line Input #1, sLineIn
If Trim$(sLineIn) <> "" Then cListBox.AddItem LCase(Replace(sLineIn, "", ""))
Wend
Close #1
sCurrentFile = ""
AfterLoadListFromFile:
Exit Sub
ErrLoadListFromFile:
Resume AfterLoadListFromFile
End Sub

Public Sub RemoveSelected(listbox As listbox)
Dim ListCount As Long
ListCount& = listbox.ListCount
Do While ListCount& > 0&
ListCount& = ListCount& - 1
If listbox.Selected(ListCount&) = True Then
listbox.RemoveItem (ListCount&)
End If
Loop
End Sub

Public Sub KillDupes(listbox As listbox)
On Error Resume Next
Dim Search1 As Long
Dim Search2 As Long
Dim KillDupe As Long
KillDupe = 0
For Search1& = 0 To listbox.ListCount - 1
For Search2& = Search1& + 1 To listbox.ListCount - 1
KillDupe = KillDupe + 1
If listbox.list(Search1&) = listbox.list(Search2&) Then
listbox.RemoveItem Search2&
Search2& = Search2& - 1
End If
Next Search2&
Next Search1&
End Sub

Public Function OpenFile(listbox As listbox, filename As String)
Dim i As Integer, str As String
i = FreeFile
If Not filename = "" Then
Open filename For Input As #i
Do While Not EOF(i)
Line Input #i, str
If Not str = "" Or Left(str, 1) = "#" Or str = " " Then listbox.AddItem str
DoEvents
Loop
Close #i
End If
End Function

Public Function SaveFile(listbox As listbox, filename As String)
Dim i As Integer
Dim X As Integer
i = FreeFile
If Not filename = "" Then
Open filename For Output As #i
For X = 0 To listbox.ListCount - 1
If Not listbox.list(X) = "" Then Print #i, listbox.list(X)
DoEvents
Next
Close #i
End If
End Function

