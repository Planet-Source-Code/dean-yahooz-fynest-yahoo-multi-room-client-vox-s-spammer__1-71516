VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{2B323CCC-50E3-11D3-9466-00A0C9700498}#1.0#0"; "yacscom.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vox Master"
   ClientHeight    =   4590
   ClientLeft      =   1230
   ClientTop       =   8025
   ClientWidth     =   3135
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   306
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   209
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Showem 
      Caption         =   "More >>"
      Height          =   255
      Left            =   2160
      TabIndex        =   18
      Top             =   3840
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "More Voice options"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3240
      TabIndex        =   14
      Top             =   1800
      Width           =   3615
      Begin VB.ComboBox VoiceServer 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "Form3.frx":0000
         Left            =   960
         List            =   "Form3.frx":0016
         TabIndex        =   17
         Text            =   "vc.yahoo.com"
         Top             =   960
         Width           =   2580
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Auto Voice Ignore New Arrivals"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   3315
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "Form3.frx":0093
         TabIndex        =   15
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Voice Key"
      Top             =   840
      Width           =   3615
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Room Key"
      Top             =   480
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Room"
      Top             =   120
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Last User To Join/Leave"
      Top             =   4200
      Width           =   2895
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form3.frx":0103
      TabIndex        =   9
      Top             =   3240
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "VoX M4$T3R"
      Top             =   120
      Width           =   3615
   End
   Begin ACTIVESKINLibCtl.SkinLabel VoiceNum 
      Height          =   255
      Left            =   2160
      OleObjectBlob   =   "Form3.frx":0161
      TabIndex        =   7
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton Iggy 
      Caption         =   "^ Ignore ^"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Enable 
      Caption         =   "Voice On"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel Talker 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form3.frx":01C9
      TabIndex        =   4
      Top             =   3480
      Width           =   2895
   End
   Begin MSComctlLib.Slider SLIDER 
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1320
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   661
      _Version        =   393216
      Max             =   20
      SelStart        =   20
      TickStyle       =   3
      Value           =   20
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   2205
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   3889
      _Version        =   393217
      Sorted          =   -1  'True
      Style           =   1
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox HandFree 
      Caption         =   "Lock Mic"
      Enabled         =   0   'False
      Height          =   330
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   1035
   End
   Begin VB.CheckBox Mute 
      Caption         =   "Mute"
      Height          =   330
      Left            =   1320
      TabIndex        =   0
      ToolTipText     =   "Mute"
      Top             =   2760
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Proudly Programmed By Yahooz Fynest"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3240
      TabIndex        =   19
      Top             =   3480
      Width           =   3615
   End
   Begin YACSCOMLibCtl.YAcs voice 
      Left            =   120
      OleObjectBlob   =   "Form3.frx":0227
      Top             =   4680
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Coded By Yahooz Fynest in 24 hours!!
'Dont Ripped My Prog All will know!
'Welcome to Study it, and learn how it works, then try make your own..

'Remember Yahooz Fynest was First to make a Program such as this Even Exist!
'Multi Room Joining, Chat support, INF Tag Processing, Names Collecting, Illy sorting with Names Collecting..
'Debugging of Incomming Packet, Stats for All the User comming an going from rooms, specifying the Room an Username's.
'Multi Vox's, Can Load a Vox for any Bot in any Room at Any Time! And be Tuned into, 5 ChatRooms at the Same time!
'Can DJ to Multiple Room, Chat by Text, and View Chat for 1 Bot and its Room at a Time!
'Each Bot has individual Boot protection, And Auto Gawd Mode/YF Mode. And the Program alerts you which Bot from which room got attacked.
'It gets a new smiley, and at anytime you can double click that gawd moded bot smiley to set back to normal!
'Built in Spammer, with custom delay, Can have 50 bots in 50 Room, Spamming a Custom Spam list you are able to Load!!
'Chat Block option to protect you if someone room lagg the Chat Bot for room your Viewing.
'Easy one click another bots smiley thats in room to change chat bot, an view a different room instantly!
'No one has made a Tool such as this before, Multi Room Data Handling, Protection, Multi Vox's
'Hope this Released Example by myself is helpful for people to prehaps make some top line tools aswell in future, Enjoy!


Dim VNode As MSComctlLib.Node

Public Sub VoiceOn()
On Error Resume Next
With voice
    .HostName = VoiceServer.Text
    .Username = Text2.Text
    .appInfo = "mc(9.0.0.234)&u=" & .Username & "&ia=us"
    .LoadSound .Username
    .confKey = Text6
    .confName = "ch/" & Text3 & "::" & Text5
    .inputGain = 99
    .inputAGC = 99
    .outputGain = 99
    .inputSource = 99
    .createAndJoinConference
    .joinConference
End With
End Sub

Private Sub Mute_Click()
On Error Resume Next
voice.outputMute = Mute.Value
End Sub

Private Sub Showem_Click()
On Error Resume Next
If Me.Width = "7065" Then
Me.Width = "3225"
Showem.Caption = "More >>"
Else
Me.Width = "7065"
Showem.Caption = "<< Hide"
End If
End Sub

Private Sub HandFree_Click()
On Error Resume Next
If HandFree.Value = 1 Then
    voice.startTransmit
Else
    voice.stopTransmit
    SkinLabel5.Caption = ""
End If
End Sub

Public Sub Enable_Click()
On Error Resume Next
If Enable.Caption = "Voice On" Then
    voice.leaveConference
    DoEvents
    TreeView1.Nodes.Clear
    VoiceOn
    Enable.Caption = "Voice Off"
Else
    voice.leaveConference
    TreeView1.Nodes.Clear
    Enable.Caption = "Voice On"
    HandFree.Value = 0
    HandFree.Enabled = False
End If
End Sub

Private Sub Iggy_Click()
On Error Resume Next
IggyLamer IDENTIFY, IDENTITY, TreeView1
End Sub

Private Sub IggyLamer(VcID As String, VcName As String, VcList As TreeView)
On Error Resume Next
Dim Igg As Integer
voice.muteSource VcID, VcName
For Igg = 1 To VcList.Nodes.Count
If VcName = VcList.Nodes.item(Igg).Tag Then
VcList.Nodes.item(Igg).Image = 6
End If
DoEvents
Next Igg
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Icon = Form1.Icon
TreeView1.ImageList = Form1.ImageList1
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Unload Me
End Sub

Private Sub TreeView1_DblClick()
On Error Resume Next
If VNode.Image = 5 Then
    VNode.Image = 6
    voice.muteSource 0, VNode.Text
Else
    VNode.Image = 5
    voice.unmuteSource 0, VNode.Text
End If
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next
Set UserNode = Node
Set VNode = Node
End Sub

Private Sub Voice_onTransmitReport(ByVal Listners As Long, ByVal All As Long)
On Error Resume Next
VoiceNum.Caption = Listners & "/" & All
End Sub

Private Sub SLIDER_Scroll()
On Error Resume Next
voice.outputGain = SLIDER.Value * 5
End Sub

Private Sub Voice_onLocalOffAir()
On Error Resume Next
SkinLabel5.Caption = ""
VoiceNum.Caption = "0/0"
End Sub

Private Sub Voice_onLocalOnAir()
On Error Resume Next
SkinLabel5.Caption = Text2.Text
End Sub

Private Sub Voice_onRemoteSourceOffAir(ByVal Id As Long, ByVal name As String)
On Error Resume Next
Talker = ""
End Sub

Private Sub Voice_onRemoteSourceOnAir(ByVal Id As Long, ByVal name As String)
On Error Resume Next
Talker = name
IDENTIFY = Id
IDENTITY = name
End Sub

Private Sub Voice_onSourceEntry(ByVal Id As Long, ByVal name As String)
On Error Resume Next
Dim i As Integer
For i = 1 To TreeView1.Nodes.Count
If UCase(TreeView1.Nodes(i).Text) = UCase(name) Then Exit Sub
DoEvents
Next
TreeView1.Nodes.Add , , name, name, 5
Text4.Text = name & " Joined"
If Check2.Value = 1 Then
If name = Form3.Text2.Text Then GoTo Lol
voice.muteSource Id, name
For i = 1 To TreeView1.Nodes.Count
If name = TreeView1.Nodes.item(i).Text Then
TreeView1.Nodes.item(i).Image = 6
End If
DoEvents
Next
Lol:
End If
End Sub

Private Sub Voice_onSourceExit(ByVal Id As Long, ByVal name As String)
Dim i As Integer
On Error Resume Next
If TreeView1.Nodes.Count < 1 Then Exit Sub
For i = 1 To TreeView1.Nodes.Count
If UCase(TreeView1.Nodes(i).Text) = UCase(name) Then
TreeView1.Nodes.Remove i
Text4.Text = name & " Leaft"
End If
DoEvents
Next
End Sub

Private Sub Voice_onSystemConnect()
On Error Resume Next
Talker = "Voice Connected!"
HandFree.Enabled = True
Enable.Caption = "Voice Off"
End Sub

Private Sub Voice_onSystemConnectFailure(ByVal codes As Long, ByVal messages As String)
On Error Resume Next
Talker = "Error!"
Enable.Caption = "Voice On"
HandFree.Enabled = False
MsgBox codes & " - " & messages, , "Voice Error Report!"
End Sub

Private Sub Voice_onSystemDisconnect()
On Error Resume Next
Talker = "Disconnected!"
HandFree.Enabled = False
Enable.Caption = "Voice On"
End Sub

Private Sub Voice_onAudioError(ByVal codes As Long, ByVal messages As String)
On Error Resume Next
Talker = "Error!"
HandFree.Enabled = False
Enable.Caption = "Voice On"
MsgBox codes & " - " & messages, , "Voice Error Report!"
End Sub

Private Sub Voice_onConferenceNotReady()
On Error Resume Next
Talker = "Error!"
HandFree.Enabled = False
Enable.Caption = "Voice On"
End Sub

Private Sub Voice_onConferenceReady()
On Error Resume Next
Talker = "Voice Ready!"
HandFree.Enabled = True
Enable.Caption = "Voice Off"
voice.outputMute = Mute.Value
End Sub



