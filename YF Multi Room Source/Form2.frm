VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Captcha"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5055
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   4815
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form2.frx":0CCE
      TabIndex        =   3
      Top             =   3000
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Enter Captcha"
      Top             =   360
      Width           =   4815
   End
   Begin VB.TextBox TxtVerify 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Captcha Here"
      Top             =   2640
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Join Next Room On List (Use This If Room Is Full)"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   4815
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1935
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Right Click Once Grabbed Desired Pic, and select Save Picture As...."
      Top             =   720
      Width           =   4815
      ExtentX         =   8493
      ExtentY         =   3413
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DONT CLOSE TILL FINISHED JOINING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   4815
   End
End
Attribute VB_Name = "Form2"
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


Private Const ENTER_KEY = 13

Private Sub Check1_Click()
TxtVeri.SetFocus
End Sub

Private Sub Command1_Click()
Dim i As Integer
On Error GoTo Dodge
For i = 1 To Form1.Combo1.Text
If Form1.ListView1.ListItems(i).SmallIcon = 1 Then
Form1.List3.ListIndex = Form1.List3.ListIndex + 1
Form1.ListView1.ListItems(i).SmallIcon = 7
Form1.Status.Panels(1).Text = "Status:Joining Room"
Form1.socket(i).SendData JoinRoom(YahooID(i))
RoomJoinedd(i) = Form1.Text1.Text
Form2.Text1.Text = "Joining Next Bot To Next Room!"
DoEvents
Exit Sub
End If
If Form1.ListView1.ListItems(i).SmallIcon = 7 Then
Form1.List3.ListIndex = Form1.List3.ListIndex + 1
Form1.ListView1.ListItems(i).SmallIcon = 7
Form1.Status.Panels(1).Text = "Status:Joining Room"
Form1.socket(i).SendData JoinRoom(YahooID(i))
RoomJoinedd(i) = Form1.Text1.Text
Form2.Text1.Text = "Joining Next Bot To Next Room!"
DoEvents
Exit Sub
End If
Next i
Exit Sub
Dodge:
Form2.Text1.Text = "Finished (No More Rooms In List To Join)"
Pause (3)
Form2.Hide
Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim X As Integer
Form1.Text28.Text = "1"
On Error Resume Next
For i = 1 To Form1.Text25.Text
If Form1.ListView1.ListItems(i).SmallIcon = 7 Then
Form1.ListView1.ListItems(i).SmallIcon = 1
DoEvents
End If
Next i
End Sub

Sub TxtVerify_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If (KeyAscii = 13) Then

Form1.CapSck.Close
Form1.CapSck.Connect "captcha.chat.yahoo.com", 80

    End If
End Sub

Private Sub TxtVerify_DblClick()
TxtVerify.Text = ""
End Sub

