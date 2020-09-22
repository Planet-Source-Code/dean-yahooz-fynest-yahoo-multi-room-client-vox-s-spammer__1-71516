VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Rooms To Rooms To Join List From Here!"
   ClientHeight    =   3495
   ClientLeft      =   915
   ClientTop       =   6195
   ClientWidth     =   6255
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Hide"
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add This Room To Rooms List!!"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   4575
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4560
      TabIndex        =   7
      Text            =   "Yahooz Fynest"
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      TabIndex        =   6
      Text            =   "ROOM NAMES"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "CATERGORIES"
      Top             =   120
      Width           =   2655
   End
   Begin VB.ComboBox cmCatagory 
      Height          =   315
      ItemData        =   "Form4.frx":0CCE
      Left            =   120
      List            =   "Form4.frx":0D17
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   480
      Width           =   1755
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   8760
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":0DDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":16B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":19D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":22AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":24A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView treCats 
      Height          =   1860
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   3281
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   3
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin MSComctlLib.TreeView treRooms 
      Height          =   2175
      Left            =   2880
      TabIndex        =   2
      Top             =   480
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   3836
      _Version        =   393217
      Indentation     =   353
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.TextBox txtRoom 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Build Your Room To Join List Using This!"
      Top             =   2760
      Width           =   6015
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Special Thanks to Whoever Done This, And To DSP-FSG whom hooked me up with it in December of 2007
'Changes i did take the time to do, is stop it Showing the RoomKey
'And the 2 Command Buttons!
Option Explicit

Dim i As Integer, ii As Integer
Dim Spt() As String, Spt2() As String
Dim R As String, k As String, L As String

Sub GetCats(Cats As String)
On Error Resume Next
treCats.Nodes.Clear
Spt = Split(Cats, "<category")
DoEvents
For i = 1 To UBound(Spt)
k = "chatroom_" & GetVal(Spt(i), "id")
Spt(i) = GetVal(Spt(i), "name")
Spt(i) = Replace(Spt(i), "&amp;", "&", , , vbTextCompare)
Spt(i) = Replace(Spt(i), "&apos;", "'", , , vbTextCompare)
treCats.Nodes.Add , , Trim(k), Trim(Spt(i)), 5
Next
End Sub

Sub GetRooms(Rooms As String)
On Error Resume Next
Spt = Split(Rooms, "<room")
treRooms.Nodes.Clear
DoEvents
For i = 1 To UBound(Spt)
L = GetVal(Spt(i), "id")
R = GetVal(Spt(i), "name")
R = Replace(R, "&apos;", "'", , , vbTextCompare)
R = Replace(R, "&amp;", "&", , , vbTextCompare)
Spt2 = Split(Spt(i), "<lobby")
For ii = 1 To UBound(Spt2)
k = GetVal(Spt2(ii), "count")
treRooms.Nodes.Add , , R, R, 1
treRooms.Nodes.Add R, tvwChild, (R & ":" & k), R & ":" & k & _
" (" & GetVal(Spt2(ii), "users") & ")" & _
" (v" & GetVal(Spt2(ii), "voices") & ")" & _
" (w" & GetVal(Spt2(ii), "webcams") & ")", 4
Next ii
Next i
End Sub

Function GetVal(ByVal StrAll As String, str As String)
On Error GoTo Error
Dim i As Integer
    i = InStr(1, StrAll, str & "=" & Chr(34), vbTextCompare)
    If i < 2 Then GoTo Error
    StrAll = Mid(StrAll, i + Len(str) + 2)
    i = InStr(StrAll, Chr(34))
    If i < 2 Then GoTo Error
    StrAll = Left(StrAll, i - 1)
Error:
    GetVal = StrAll
End Function

Sub LoadRooms(Cat As String)
    If Cat = "" Then Exit Sub
    treRooms.Nodes.Clear
    treRooms.Nodes.Add , , , "Loading...", 3

    Dim data As String
        If cmCatagory.Text = "USA" Then
            data = "http://insider.msg.yahoo.com/ycontent/?" & Cat
        ElseIf cmCatagory.Text = "Argentina" Then
            data = "http://ar.insider.msg.yahoo.com/ycontent/?" & Cat
        ElseIf cmCatagory.Text = "Asia" Then
            data = "http://aa.insider.msg.yahoo.com/ycontent/?" & Cat
        ElseIf cmCatagory.Text = "Australia" Then
            data = "http://au.insider.msg.yahoo.com/ycontent/?" & Cat
        ElseIf cmCatagory.Text = "Brazil" Then
            data = "http://br.insider.msg.yahoo.com/ycontent/?" & Cat
        ElseIf cmCatagory.Text = "Canada" Then
            data = "http://ca.insider.msg.yahoo.com/ycontent/?" & Cat
        ElseIf cmCatagory.Text = "Canada French" Then
            data = "http://cf.insider.msg.yahoo.com/ycontent/?" & Cat
        ElseIf cmCatagory.Text = "China" Then
            data = "http://cn.insider.msg.yahoo.com/ycontent/?" & Cat
        ElseIf cmCatagory.Text = "Denmark" Then
            data = "http://de.insider.msg.yahoo.com/ycontent/?" & Cat
        ElseIf cmCatagory.Text = "France" Then
            data = "http://fr.insider.msg.yahoo.com/ycontent/?" & Cat
        ElseIf cmCatagory.Text = "Germany" Then
            data = "http://insider.msg.yahoo.com/ycontent/?" & Cat
        ElseIf cmCatagory.Text = "Hong Kong" Then
            data = "http://hk.insider.msg.yahoo.com/ycontent/?" & Cat
        ElseIf cmCatagory.Text = "India" Then
            data = "http://in.insider.msg.yahoo.com/ycontent/?" & Cat
        ElseIf cmCatagory.Text = "Italy" Then
            data = "http://it.insider.msg.yahoo.com/ycontent/?" & Cat
        ElseIf cmCatagory.Text = "Japan" Then
            data = "http://insider.msg.yahoo.co.jp/ycontent/?" & Cat
        ElseIf cmCatagory.Text = "Korea" Then
            data = "http://kr.insider.msg.yahoo.com/ycontent/?" & Cat
        ElseIf cmCatagory.Text = "Mexico" Then
            data = "http://mx.insider.msg.yahoo.com/ycontent/?" & Cat
        ElseIf cmCatagory.Text = "Norway" Then
            data = "http://no.insider.msg.yahoo.com/ycontent/?" & Cat
        ElseIf cmCatagory.Text = "Singapore" Then
            data = "http://sg.insider.msg.yahoo.com/ycontent/?" & Cat
        ElseIf cmCatagory.Text = "Spain" Then
            data = "http://se.insider.msg.yahoo.com/ycontent/?" & Cat
        ElseIf cmCatagory.Text = "Sweden" Then
            data = "http://es.insider.msg.yahoo.com/ycontent/?" & Cat
        ElseIf cmCatagory.Text = "Taiwan" Then
            data = "http://tw.insider.msg.yahoo.com/ycontent/?" & Cat
        ElseIf cmCatagory.Text = "Ukraine" Then
            data = "http://uk.insider.msg.yahoo.com/ycontent/?" & Cat
        End If
    DoEvents
    data = Inet1.OpenURL(data)
    DoEvents
    GetRooms data
    
End Sub


Private Sub cmCatagory_Click()
    treCats.Nodes.Clear
    treRooms.Nodes.Clear
    Call cmdRefresh_Click
End Sub



Public Sub cmdRefresh_Click()
Dim data As String
    treRooms.Nodes.Clear
    treCats.Nodes.Clear
    treCats.Nodes.Add , , , "Loading..."
        If cmCatagory.Text = "USA" Then
            data = "http://insider.msg.yahoo.com/ycontent/?chatcat"
        ElseIf cmCatagory.Text = "Asia" Then
            data = "http://aa.insider.msg.yahoo.com/ycontent/?chatcat"
        ElseIf cmCatagory.Text = "Argentina" Then
            data = "http://ar.insider.msg.yahoo.com/ycontent/?chatcat"
        ElseIf cmCatagory.Text = "Australia" Then
            data = "http://au.insider.msg.yahoo.com/ycontent/?chatcat"
        ElseIf cmCatagory.Text = "Brazil" Then
            data = "http://br.insider.msg.yahoo.com/ycontent/?chatcat"
        ElseIf cmCatagory.Text = "Canada" Then
            data = "http://ca.insider.msg.yahoo.com/ycontent/?chatcat"
        ElseIf cmCatagory.Text = "Canada French" Then
            data = "http://cf.insider.msg.yahoo.com/ycontent/?chatcat"
        ElseIf cmCatagory.Text = "China" Then
            data = "http://cn.insider.msg.yahoo.com/ycontent/?chatcat"
        ElseIf cmCatagory.Text = "Denmark" Then
            data = "http://dk.insider.msg.yahoo.com/ycontent/?chatcat"
        ElseIf cmCatagory.Text = "Germany" Then
            data = "http://insider.msg.yahoo.com/ycontent/?chatcat"
        ElseIf cmCatagory.Text = "France" Then
            data = "http://fr.insider.msg.yahoo.com/ycontent/?chatcat"
        ElseIf cmCatagory.Text = "Hong Kong" Then
            data = "http://hk.insider.msg.yahoo.com/ycontent/?chatcat"
        ElseIf cmCatagory.Text = "India" Then
            data = "http://in.insider.msg.yahoo.com/ycontent/?chatcat"
        ElseIf cmCatagory.Text = "Italy" Then
            data = "http://it.insider.msg.yahoo.com/ycontent/?chatcat"
        ElseIf cmCatagory.Text = "Japan" Then
            data = "http://insider.msg.yahoo.co.jp/ycontent/?chatcat"
        ElseIf cmCatagory.Text = "Korea" Then
            data = "http://kr.insider.msg.yahoo.com/ycontent/?chatcat"
        ElseIf cmCatagory.Text = "Mexico" Then
            data = "http://mx.insider.msg.yahoo.com/ycontent/?chatcat"
        ElseIf cmCatagory.Text = "Norway" Then
            data = "http://no.insider.msg.yahoo.com/ycontent/?chatcat"
        ElseIf cmCatagory.Text = "Singapore" Then
            data = "http://sg.insider.msg.yahoo.com/ycontent/?chatcat"
        ElseIf cmCatagory.Text = "Spain" Then
            data = "http://es.insider.msg.yahoo.com/ycontent/?chatcat"
        ElseIf cmCatagory.Text = "Sweden" Then
            data = "http://se.insider.msg.yahoo.com/ycontent/?chatcat"
        ElseIf cmCatagory.Text = "Taiwan" Then
            data = "http://tw.insider.msg.yahoo.com/ycontent/?chatcat"
        ElseIf cmCatagory.Text = "Ukraine" Then
            data = "http://uk.insider.msg.yahoo.com/ycontent/?chatcat"
        End If

data = Inet1.OpenURL(data)
GetCats data

End Sub

Private Sub Command1_Click()
Dim Room As String
Room = Me.txtRoom.Text
LoopList Room, Form1.List3
Form1.List3.AddItem Room
End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

Private Sub Form_Load()
cmCatagory.ListIndex = 22
Call cmdRefresh_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub treCats_Click()
On Error Resume Next
LoadRooms treCats.SelectedItem.Key
End Sub

Private Sub TreRooms_Click()
On Error Resume Next
If InStr(treRooms.SelectedItem.Key, ":") Then
txtRoom = treRooms.SelectedItem.Key
End If
End Sub

Private Sub treRooms_Collapse(ByVal Node As MSComctlLib.Node)
Node.Image = 1
End Sub

Private Sub treRooms_Expand(ByVal Node As MSComctlLib.Node)
Node.Image = 2
End Sub




