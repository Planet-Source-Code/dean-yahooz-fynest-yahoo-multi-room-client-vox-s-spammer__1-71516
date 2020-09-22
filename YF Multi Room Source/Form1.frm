VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "YF Multi Room Client/Spammer          ..::Version1.0::..  (50bot)            made by Yahooz Fynest aka Dean"
   ClientHeight    =   5910
   ClientLeft      =   1230
   ClientTop       =   1620
   ClientWidth     =   11280
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   11280
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text13 
      Height          =   315
      Left            =   840
      TabIndex        =   107
      Text            =   "<font INF ID:YF-MRC Ver:1.0 By:Yahooz-Fynest About:Client/Spammer>"
      Top             =   6720
      Width           =   7095
   End
   Begin VB.Timer Timer4 
      Interval        =   2000
      Left            =   6480
      Top             =   7080
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6000
      Top             =   7080
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Room"
      Top             =   7920
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog c 
      Left            =   8040
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox SpamText 
      Height          =   315
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Spam"
      Top             =   7560
      Width           =   2175
   End
   Begin VB.TextBox Text10 
      Height          =   255
      Left            =   2400
      TabIndex        =   15
      Text            =   "Text10"
      Top             =   7440
      Width           =   1575
   End
   Begin VB.TextBox Text31 
      Height          =   315
      Left            =   840
      TabIndex        =   14
      Text            =   "Text31"
      Top             =   7440
      Width           =   855
   End
   Begin VB.TextBox Text30 
      Height          =   315
      Left            =   2880
      TabIndex        =   13
      Text            =   "2"
      Top             =   8160
      Width           =   495
   End
   Begin VB.TextBox Text28 
      Height          =   315
      Left            =   7320
      TabIndex        =   12
      Text            =   "1"
      Top             =   7560
      Width           =   615
   End
   Begin VB.TextBox Text26 
      Height          =   315
      Left            =   1560
      TabIndex        =   11
      Text            =   "1"
      Top             =   8160
      Width           =   615
   End
   Begin VB.TextBox Text25 
      Height          =   315
      Left            =   7320
      TabIndex        =   10
      Text            =   "0"
      Top             =   7920
      Width           =   615
   End
   Begin VB.TextBox Text19 
      Height          =   315
      Left            =   4080
      TabIndex        =   9
      Text            =   """>"
      Top             =   7440
      Width           =   855
   End
   Begin VB.TextBox Text18 
      Height          =   315
      Left            =   840
      TabIndex        =   8
      Text            =   "<font face=""Comic Sans MS"" size=""12"">"
      Top             =   7080
      Width           =   4095
   End
   Begin VB.TextBox Text16 
      Height          =   315
      Left            =   4200
      TabIndex        =   7
      Text            =   "1"
      Top             =   7800
      Width           =   735
   End
   Begin VB.TextBox Text15 
      Height          =   315
      Left            =   2040
      TabIndex        =   6
      Text            =   "1"
      Top             =   7800
      Width           =   2055
   End
   Begin VB.TextBox Text9 
      Height          =   315
      Left            =   840
      TabIndex        =   5
      Text            =   "1"
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5040
      Top             =   7080
   End
   Begin VB.TextBox txtnumber 
      Height          =   315
      Left            =   1800
      TabIndex        =   4
      Text            =   "546"
      Top             =   7440
      Width           =   495
   End
   Begin VB.TextBox Text11 
      Height          =   315
      Left            =   840
      TabIndex        =   3
      Text            =   "0"
      Top             =   8160
      Width           =   615
   End
   Begin VB.TextBox Text12 
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Text            =   "0"
      Top             =   8160
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   8040
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3480
      TabIndex        =   1
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   5520
      Top             =   7080
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5655
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   11
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5733
            MinWidth        =   5733
            Text            =   "Status:"
            TextSave        =   "Status:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1446
            MinWidth        =   1446
            Text            =   "Logged In:"
            TextSave        =   "Logged In:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            Text            =   "0"
            TextSave        =   "0"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1270
            MinWidth        =   1270
            Text            =   "In Room:"
            TextSave        =   "In Room:"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            Text            =   "0"
            TextSave        =   "0"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "Spam Sent:"
            TextSave        =   "Spam Sent:"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1147
            MinWidth        =   1147
            Text            =   "0"
            TextSave        =   "0"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2293
            MinWidth        =   2293
            Text            =   "Total Incomming:"
            TextSave        =   "Total Incomming:"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "0"
            TextSave        =   "0"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1235
            MinWidth        =   1235
            Text            =   "Blocked:"
            TextSave        =   "Blocked:"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "0"
            TextSave        =   "0"
         EndProperty
      EndProperty
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8640
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0CCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1ECE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2BAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2CAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2DB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3A8E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock socket 
      Index           =   0
      Left            =   7440
      Top             =   7080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8040
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock CapSck 
      Left            =   6960
      Top             =   7080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab Debugging 
      Height          =   5655
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   9975
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Bots / Controls / Stats"
      TabPicture(0)   =   "Form1.frx":4048
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label15"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label14"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Combo1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command5"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command11"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command14"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text22"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Combo2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "List4"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Command6"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Command7"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Combo3"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Chat Bot / Boot"
      TabPicture(1)   =   "Form1.frx":4064
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label8"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame6"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command24"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Text7"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Text5"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "SpamMessages / Rooms"
      TabPicture(2)   =   "Form1.frx":4080
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "FrameSpam"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame1"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame3"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Names Grabbed / PMs"
      TabPicture(3)   =   "Form1.frx":409C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame7"
      Tab(3).Control(1)=   "Frame5"
      Tab(3).Control(2)=   "Frame4"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Incomming Debug"
      TabPicture(4)   =   "Form1.frx":40B8
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label16"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Text555"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Check4"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Command26"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).ControlCount=   4
      Begin VB.ComboBox Combo3 
         Height          =   330
         ItemData        =   "Form1.frx":40D4
         Left            =   3000
         List            =   "Form1.frx":40E7
         Style           =   2  'Dropdown List
         TabIndex        =   109
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Clear Debug"
         Height          =   255
         Left            =   -74880
         TabIndex        =   100
         Top             =   5280
         Width           =   1215
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Debug??"
         Height          =   255
         Left            =   -64800
         TabIndex        =   99
         Top             =   5280
         Width           =   975
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Clear List"
         Height          =   255
         Left            =   1200
         TabIndex        =   97
         Top             =   5280
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Save"
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   5280
         Width           =   1095
      End
      Begin VB.ListBox List4 
         Height          =   900
         Left            =   120
         TabIndex        =   95
         Top             =   4320
         Width           =   2175
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":40FF
         Left            =   120
         List            =   "Form1.frx":414E
         Style           =   2  'Dropdown List
         TabIndex        =   94
         Top             =   3600
         Width           =   3615
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   82
         Text            =   "Not Processing Chat"
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox Text7 
         Height          =   315
         Left            =   -70080
         Locked          =   -1  'True
         TabIndex        =   81
         Text            =   "Not Processing Chat"
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox Text22 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   76
         Text            =   "Login Delay:"
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton Command14 
         Caption         =   "All Leave All DC"
         Height          =   315
         Left            =   2280
         TabIndex        =   75
         ToolTipText     =   "Make Bots In Rooms Leave and All Bots DC from Yahoo!"
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Reset Failed Join Bots"
         Height          =   495
         Left            =   2280
         TabIndex        =   74
         ToolTipText     =   "Reset Failed Joins Back to just Logged In Smiley!"
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3120
         TabIndex        =   73
         Text            =   "0.75"
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Join Bot/s"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         TabIndex        =   72
         Top             =   2640
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   330
         Left            =   2280
         TabIndex        =   71
         Text            =   "#Bots"
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "LoG In Bots"
         Height          =   330
         Left            =   2280
         TabIndex        =   70
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Load Bots"
         Height          =   375
         Left            =   2280
         MouseIcon       =   "Form1.frx":43B8
         Picture         =   "Form1.frx":5282
         TabIndex        =   69
         Top             =   480
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Caption         =   "If anyone PMs any Bot, Auto PM back this Message..."
         Height          =   855
         Left            =   -74880
         TabIndex        =   66
         Top             =   5820
         Visible         =   0   'False
         Width           =   11055
         Begin VB.TextBox Text21 
            Height          =   315
            Left            =   120
            MaxLength       =   150
            TabIndex        =   68
            Text            =   "Your PM Spam Message To Auto Send to anyone that PMs any of your Bots Goes Here!!"
            Top             =   360
            Width           =   10815
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Auto PM Spam Response On/Off"
            Height          =   255
            Left            =   8160
            TabIndex        =   67
            Top             =   0
            Width           =   2775
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Stats.... and Listing Of What Bot is In What Room!!"
         Height          =   5055
         Left            =   3840
         TabIndex        =   64
         Top             =   480
         Width           =   7335
         Begin VB.CommandButton Command3 
            Caption         =   "Clear Stats Screen"
            Height          =   255
            Left            =   3720
            TabIndex        =   79
            Top             =   4680
            Width           =   3495
         End
         Begin VB.ListBox List2 
            Height          =   4680
            ItemData        =   "Form1.frx":614C
            Left            =   120
            List            =   "Form1.frx":614E
            TabIndex        =   65
            Top             =   240
            Width           =   3495
         End
         Begin RichTextLib.RichTextBox Text17 
            Height          =   4335
            Left            =   3720
            TabIndex        =   80
            ToolTipText     =   "Double Click To Clear!"
            Top             =   240
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   7646
            _Version        =   393217
            ScrollBars      =   2
            TextRTF         =   $"Form1.frx":6150
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
      End
      Begin VB.Frame Frame7 
         Caption         =   "illy and @ , ID User's Found in each Room!"
         Height          =   5055
         Left            =   -67560
         TabIndex        =   56
         Top             =   480
         Width           =   3735
         Begin VB.ListBox List10 
            Height          =   1950
            Left            =   120
            TabIndex        =   61
            ToolTipText     =   "Double Click to Remove 1"
            Top             =   2400
            Width           =   3495
         End
         Begin VB.CommandButton Command18 
            Caption         =   "Save @ ID's"
            Height          =   255
            Left            =   1440
            TabIndex        =   60
            Top             =   4440
            Width           =   1095
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Clear Both"
            Height          =   495
            Left            =   2640
            TabIndex        =   59
            Top             =   4440
            Width           =   975
         End
         Begin VB.CommandButton Command15 
            Caption         =   "Save illy's"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   4440
            Width           =   1215
         End
         Begin VB.ListBox List9 
            Height          =   2160
            Left            =   120
            TabIndex        =   57
            ToolTipText     =   "Double Click to Remove 1"
            Top             =   240
            Width           =   3495
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   255
            Left            =   1440
            TabIndex        =   63
            Top             =   4680
            Width           =   1095
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   4680
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Users That Sent a PM To 1 Of Your Bots!!"
         Height          =   5055
         Left            =   -71280
         TabIndex        =   50
         Top             =   480
         Width           =   3615
         Begin VB.CommandButton Command23 
            Caption         =   "Kill Duplicates"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   4680
            Width           =   1575
         End
         Begin VB.CommandButton Command31 
            Caption         =   "Clear All"
            Height          =   495
            Left            =   2520
            TabIndex        =   53
            Top             =   4440
            Width           =   975
         End
         Begin VB.CommandButton Command30 
            Caption         =   "Save PM User ID's"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   4440
            Width           =   1575
         End
         Begin VB.ListBox List6 
            Height          =   4050
            Left            =   120
            TabIndex        =   51
            ToolTipText     =   "Double Click to Remove 1"
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   255
            Left            =   1680
            TabIndex        =   55
            Top             =   4680
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "User's ID's From Every Room Bots Are In!!"
         Height          =   5055
         Left            =   -74880
         TabIndex        =   44
         Top             =   480
         Width           =   3495
         Begin VB.CommandButton Command4 
            Caption         =   "Kill Duplicates"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   4680
            Width           =   1695
         End
         Begin VB.CommandButton Command29 
            Caption         =   "Clear All"
            Height          =   495
            Left            =   2520
            TabIndex        =   47
            Top             =   4440
            Width           =   855
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Save Room/s Users"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   4440
            Width           =   1695
         End
         Begin VB.ListBox List5 
            Height          =   4050
            Left            =   120
            TabIndex        =   45
            ToolTipText     =   "Double Click to Remove 1"
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   255
            Left            =   1800
            TabIndex        =   49
            Top             =   4680
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Rooms For Bots To Join : 0"
         Height          =   4575
         Left            =   -74880
         TabIndex        =   36
         Top             =   480
         Width           =   3735
         Begin VB.CommandButton Command28 
            Caption         =   "ROOMS"
            Height          =   375
            Left            =   2520
            TabIndex        =   43
            Top             =   3600
            Width           =   1095
         End
         Begin VB.CommandButton Command27 
            Caption         =   "Load List"
            Height          =   375
            Left            =   120
            TabIndex        =   42
            Top             =   4080
            Width           =   1095
         End
         Begin VB.CommandButton Command25 
            Caption         =   "Clear All"
            Height          =   375
            Left            =   2520
            TabIndex        =   41
            Top             =   4080
            Width           =   1095
         End
         Begin VB.CommandButton Command22 
            Caption         =   "^^ Add To RoomsList^^"
            Height          =   375
            Left            =   120
            TabIndex        =   40
            Top             =   3600
            Width           =   2295
         End
         Begin VB.CommandButton Command21 
            Caption         =   "Save List"
            Height          =   375
            Left            =   1320
            TabIndex        =   39
            Top             =   4080
            Width           =   1095
         End
         Begin VB.TextBox Text20 
            Height          =   315
            Left            =   120
            MaxLength       =   50
            TabIndex        =   38
            Text            =   "RoomName2Join:12"
            Top             =   3240
            Width           =   3495
         End
         Begin VB.ListBox List3 
            Height          =   3000
            ItemData        =   "Form1.frx":61C7
            Left            =   120
            List            =   "Form1.frx":61C9
            TabIndex        =   37
            ToolTipText     =   "Double Click to Remove 1"
            Top             =   240
            Width           =   3495
         End
      End
      Begin VB.Frame FrameSpam 
         Caption         =   "Spam Messages List : 0"
         Height          =   4575
         Left            =   -71040
         TabIndex        =   29
         Top             =   480
         Width           =   7215
         Begin VB.CommandButton Command33 
            Caption         =   "^^ Load Spam Messages ^^"
            Height          =   375
            Left            =   2760
            TabIndex        =   35
            Top             =   4080
            Width           =   2295
         End
         Begin VB.CommandButton Command32 
            Caption         =   "Save List"
            Height          =   375
            Left            =   5160
            TabIndex        =   34
            Top             =   4080
            Width           =   975
         End
         Begin VB.CommandButton Command19 
            Caption         =   "^Add To Spam Messages List^"
            Height          =   375
            Left            =   120
            TabIndex        =   33
            Top             =   4080
            Width           =   2535
         End
         Begin VB.CommandButton Command20 
            Caption         =   "Clear All"
            Height          =   375
            Left            =   6240
            TabIndex        =   32
            Top             =   4080
            Width           =   855
         End
         Begin VB.TextBox Text14 
            Height          =   285
            Left            =   120
            MaxLength       =   150
            TabIndex        =   31
            Text            =   "Spam Message To Add Here"
            Top             =   3720
            Width           =   6975
         End
         Begin VB.ListBox List1 
            Height          =   3420
            ItemData        =   "Form1.frx":61CB
            Left            =   120
            List            =   "Form1.frx":61CD
            TabIndex        =   30
            ToolTipText     =   "DOUBLE CLICK: Remove From Spam Filter List"
            Top             =   240
            Width           =   6975
         End
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Load Vox For Current Chat Bot (Multi Vox's)"
         Height          =   330
         Left            =   -67680
         TabIndex        =   28
         Top             =   480
         Width           =   3855
      End
      Begin VB.Frame Frame6 
         Caption         =   "CHAT  (Here you can see chat, and chat back, to the chatroom of a Bot you make the Chatbot)"
         Height          =   4695
         Left            =   -74880
         TabIndex        =   19
         Top             =   840
         Width           =   11055
         Begin VB.CheckBox Check6 
            Caption         =   "No INF Tag with Chat Messages"
            Height          =   495
            Left            =   120
            TabIndex        =   105
            Top             =   4080
            Width           =   1695
         End
         Begin VB.CheckBox Check3 
            Caption         =   "YF Mode All Other Bots in Rooms"
            Height          =   375
            Left            =   120
            TabIndex        =   93
            ToolTipText     =   "Blocks all Incommign Data to Other Bots in Room"
            Top             =   1080
            Width           =   1695
         End
         Begin VB.CheckBox Check2 
            Caption         =   "YF Mode Chat-Bot"
            Height          =   255
            Left            =   120
            TabIndex        =   92
            ToolTipText     =   "Blocks All Data to Set Chat Bot"
            Top             =   720
            Width           =   1695
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Chat Text Block !"
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   360
            Width           =   1575
         End
         Begin VB.Frame Frame8 
            Caption         =   "Spam Controls!!"
            Height          =   2535
            Left            =   120
            TabIndex        =   85
            Top             =   1560
            Width           =   1695
            Begin VB.CommandButton Command12 
               Caption         =   "Spamming ON"
               Height          =   375
               Left            =   120
               TabIndex        =   88
               Top             =   360
               Width           =   1335
            End
            Begin VB.CommandButton Command13 
               Caption         =   "Spamming OFF"
               Height          =   375
               Left            =   120
               TabIndex        =   87
               Top             =   960
               Width           =   1335
            End
            Begin VB.TextBox Text4 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   1080
               TabIndex        =   86
               Text            =   "60"
               ToolTipText     =   "Amount of Seconds to Pause Between Spamming Messages to Multiple Rooms!"
               Top             =   1440
               Width           =   375
            End
            Begin VB.Label Label13 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Dont Run Below 30 To Avoid Chat Ban"
               Height          =   495
               Left            =   120
               TabIndex        =   104
               Top             =   1920
               Width           =   1455
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Spam is OFF"
               Height          =   255
               Left            =   120
               TabIndex        =   90
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Spam Delay: (Secs)"
               Height          =   495
               Left            =   120
               TabIndex        =   89
               Top             =   1440
               Width           =   975
            End
         End
         Begin VB.TextBox Text8 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   8520
            TabIndex        =   24
            Top             =   4200
            Width           =   2415
         End
         Begin VB.ListBox List8 
            Height          =   1950
            Left            =   8520
            TabIndex        =   23
            ToolTipText     =   "Single Click To Select Victim, Double to Remove Name From List"
            Top             =   2280
            Width           =   2415
         End
         Begin VB.ListBox List7 
            Height          =   1530
            Left            =   8520
            TabIndex        =   22
            Top             =   480
            Width           =   2415
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Send"
            Height          =   375
            Left            =   7680
            TabIndex        =   21
            Top             =   4200
            Width           =   735
         End
         Begin VB.TextBox Text6 
            Height          =   315
            Left            =   1920
            TabIndex        =   20
            Text            =   "DoubleClick to Clear! Push Enter To Send Faster!"
            Top             =   4200
            Width           =   5655
         End
         Begin RichTextLib.RichTextBox Text177 
            Height          =   3855
            Left            =   1920
            TabIndex        =   25
            ToolTipText     =   "Double Click To Clear!"
            Top             =   240
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   6800
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"Form1.frx":61CF
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Users On Yahoo In Current Room"
            Height          =   255
            Left            =   8400
            TabIndex        =   27
            Top             =   2040
            Width           =   2655
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Users On Client In Current Room"
            Height          =   255
            Left            =   8520
            TabIndex        =   26
            Top             =   240
            Width           =   2415
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3015
         Left            =   120
         TabIndex        =   77
         ToolTipText     =   $"Form1.frx":6246
         Top             =   480
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   5318
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin RichTextLib.RichTextBox Text555 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   102
         ToolTipText     =   "Double Click To Clear!"
         Top             =   480
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   8281
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Form1.frx":6376
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
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "YMSG: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   108
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "YF"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   50.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   2400
         TabIndex        =   103
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Debugging will not be Processed For Chat Bot if in Gawd Mode, and Same for if Gawd moding all other bots in Rooms!!"
         Height          =   255
         Left            =   -73560
         TabIndex        =   101
         Top             =   5280
         Width           =   8655
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "Working Bots List"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   98
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Chat Bot Selected:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   84
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Bot Is In Room:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -71160
         TabIndex        =   83
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"Form1.frx":63ED
         Height          =   495
         Left            =   -74880
         TabIndex        =   78
         Top             =   5160
         Width           =   11055
      End
   End
   Begin VB.Label YahoozFynestLabel 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   0
      TabIndex        =   106
      Top             =   10200
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
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

Option Explicit
Dim i As Integer
Dim Jj As Integer
Dim dD As Integer
Dim PAckets As Integer
Dim Header As ColumnHeader, item As ListItem, X As Integer
Dim Vic As String
Private Const ENTER_KEY = 13
Dim IncomePacks(0 To 50) As String
Dim LoggedInBot(0 To 50) As Boolean

Private Sub Command10_Click()
List8.ListIndex = -1
Timer3 = True
End Sub

Private Sub Command11_Click()
On Error Resume Next
For i = 1 To Text25.Text
If ListView1.ListItems(i).SmallIcon = 7 Then
ListView1.ListItems(i).SmallIcon = 1
VoiceKey(i) = ""
RoomKey(i) = ""
RoomJoinedd(i) = ""
Pause (0.1)
DoEvents
End If
Next i
End Sub

Private Sub Command1_Click()
On Error Resume Next
If Command1.Caption = "Load Bots" Then
Dim LTD As String, F As Variant, X As Integer
Set Header = ListView1.ColumnHeaders.Add(, , "", 300)
Set Header = ListView1.ColumnHeaders.Add(, , "YBots", 3000)
Set Header = ListView1.ColumnHeaders.Add(, , "Y!PW")
F = FreeFile
With CommonDialog1
.filename = ""
.DialogTitle = "Load Bots"
.Filter = "All Supported Types|*.txt"
.ShowOpen
If .filename = "" Then Exit Sub
Open .filename For Input As #F
While Not EOF(1)
Input #1, LTD
F = Split(LTD, ":")
           
If X < 50 Then
Set item = ListView1.ListItems.Add(, , , , 2)
item.SubItems(1) = F(0)
item.SubItems(2) = F(1)
X = X + 1
Combo1.AddItem X
Status.Panels(1).Text = "Bots Loaded : " & X
DoEvents
End If
Wend
Close #1
ListView1.View = lvwReport
End With
Combo1.Text = "1"
Text25.Text = X
Command1.Caption = "Clear"
Else
Combo1.RemoveItem (X)
Combo1.Clear
Combo1.Text = "#Bots"
ListView1.ListItems.Clear
Command1.Caption = "Load Bots"
End If
End Sub

Private Sub Command12_Click()
If Status.Panels(7) = "0" Then List1.ListIndex = -1
Timer2 = True
Label3.Caption = "Spam is ON"
End Sub

Private Sub Command13_Click()
Timer2 = False
Label3.Caption = "Spam is OFF"
End Sub

Private Sub Command14_Click()
On Error Resume Next
'Only need to Close the Sockets to DC ID's an Make them Leave Rooms!
Command5.Enabled = False
List4.Clear
Text15.Text = "1"
Command2.Caption = "LoG In Bots"
Status.Panels(1) = "Status: All DC'd, All LeaftRooms"
For i = 1 To Text25.Text
ListView1.ListItems(i).SmallIcon = 2
socket(i).Close
Unload socket(i)
LoggedInBot(i) = False
VoiceKey(i) = ""
RoomKey(i) = ""
RoomJoinedd(i) = ""
Status.Panels(3).Text = Status.Panels(3).Text - 1
Next i
Pause (0.2)
List2.Clear
Status.Panels(3) = "0"
Status.Panels(5) = "0"
Status.Panels(7) = "0"
End Sub

Private Sub Command15_Click()
Call SaveList(c, List9)
End Sub

Private Sub Command16_Click()
Call SaveList(c, List5)
End Sub

Private Sub Command17_Click()
List9.Clear
List10.Clear
End Sub

Private Sub Command18_Click()
Call SaveList(c, List10)
End Sub

Private Sub Command19_Click()
List1.AddItem Text14.Text
Text14.Text = ""
End Sub

Private Sub Command2_Click()
On Error Resume Next
If Combo1.Text = "Bots" Then
MsgBox "Select amount to Login!", , "YF"
Exit Sub
End If
Command2.Caption = "LoGin OffLine"
Dim i As Integer
For i = 1 To Combo1.Text
If ListView1.ListItems(i).SmallIcon = 2 Then
Load socket(i)
YahooID(i) = ListView1.ListItems(i).SubItems(1)
Password(i) = ListView1.ListItems(i).SubItems(2)
DoEvents
LoggedInBot(i) = False
socket(i).Close
socket(i).Connect Combo2.Text, 5050
DoEvents
If Combo2.ListIndex = Combo2.ListCount - 1 Then
Combo2.ListIndex = 0
Else
Combo2.ListIndex = Combo2.ListIndex + 1
End If
DoEvents
Pause (1)
Pause (Text3.Text)
DoEvents
End If
DoEvents
Next i
DoEvents
Pause (2)
Status.Panels(1).Text = "Bots Logged In!!"
Status.Panels(3).Text = List4.ListCount
Command5.Enabled = True
End Sub

Private Sub Command20_Click()
List1.Clear
End Sub

Private Sub Command21_Click()
Call SaveList(c, List3)
End Sub

Private Sub Command22_Click()
List3.AddItem Text20.Text
End Sub

Private Sub Command23_Click()
Call KillDupes(List6)
End Sub

Private Sub Command24_Click()
On Error Resume Next
Dim Rky As String
Dim vcky As String
Dim VcID As String
Dim vcrm As String
Dim i As Integer
For i = 1 To Form1.Text25.Text
If Form1.ListView1.ListItems(i).SmallIcon = 4 Then
Rky = RoomKey(i)
vcky = VoiceKey(i)
GoTo oke
End If
Next i
MsgBox "You have no Room Chat Bot Selected!", , "YF"
Exit Sub
oke:
VcID = Text5.Text
vcrm = Text7.Text
If Rky = "" Then Exit Sub
If vcky = "" Then Exit Sub
If VcID = "" Then Exit Sub
If vcrm = "" Then Exit Sub
Dim Vox As Integer
Vox = FindVox(VcID)
If Vox <= 0 Then
Set VCForm = Forms.Add("Form3")
Load VCForm
VCForm.Tag = VcID
VCForm.Text2.Text = VcID
VCForm.Text3.Text = vcrm
VCForm.Text5.Text = Rky
VCForm.Text6.Text = vcky
VCForm.Caption = "Vox: " & VcID
VCForm.Show
Set VCForm = Nothing
DoEvents
Exit Sub
End If
Forms(Vox).Show
End Sub

Private Sub Command25_Click()
List3.Clear
End Sub

Private Sub Command26_Click()
Text555.Text = ""
End Sub

Private Sub Command27_Click()
Call LoadList(c, List3)
End Sub

Private Sub Command28_Click()
Form4.Show
End Sub

Private Sub Command29_Click()
List5.Clear
End Sub

Private Sub Command3_Click()
Text17.Text = ""
End Sub

Private Sub Command30_Click()
Call SaveList(c, List6)
End Sub

Private Sub Command31_Click()
List6.Clear
End Sub

Private Sub Command32_Click()
Call SaveList(c, List1)
End Sub

Private Sub Command33_Click()
Call LoadList(c, List1)
End Sub

Private Sub Command4_Click()
Call KillDupes(List5)
End Sub

Private Sub Command5_Click()
On Error GoTo Error
If Status.Panels(5) = "0" Then List3.ListIndex = -1
For i = 1 To Combo1.Text
If ListView1.ListItems(i).SmallIcon = 1 Then
List3.ListIndex = List3.ListIndex + 1
ListView1.ListItems(i).SmallIcon = 7
Form1.Text28.Text = "2"
Status.Panels(1).Text = "Status:Joining Room"
VoiceKey(i) = ""
RoomKey(i) = ""
RoomJoinedd(i) = Form1.Text1.Text
socket(i).SendData JoinRoom(YahooID(i))
DoEvents
Exit Sub
End If
If ListView1.ListItems(i).SmallIcon = 7 Then
List3.ListIndex = List3.ListIndex + 1
Form1.Text28.Text = "2"
Status.Panels(1).Text = "Status:Joining Room"
VoiceKey(i) = ""
RoomKey(i) = ""
RoomJoinedd(i) = Form1.Text1.Text
socket(i).SendData JoinRoom(YahooID(i))
DoEvents
Exit Sub
End If
Next i
MsgBox "There No More Bots Logged In To Join The Rooms!!", , "No Bots"
Exit Sub
Error:
MsgBox "There No More Rooms On Rooms to Join List to Join!", , "No Rooms"
Exit Sub
End Sub

Private Sub Command6_Click()
Call SaveList(c, List4)
End Sub

Private Sub Command7_Click()
List4.Clear
End Sub

Private Sub Command8_Click()
On Error Resume Next
For i = 1 To Text25.Text
If ListView1.ListItems(i).SmallIcon = 4 Then
socket(i).SendData SendChat(YahooID(i), RoomJoinedd(i), Text18.Text, Text6.Text)
Text177.SelLength = 0
Text177.SelStart = Len(Text177)
Text177.SelBold = True: Text177.SelColor = vbBlue
Text177.SelText = YahooID(i) & ": "
Text177.SelBold = False: Text177.SelColor = vbBlack
Text177.SelText = Text6.Text & vbNewLine
Text177.SelStart = Len(Text177)
DoEvents
Text6.Text = ""
Text6.SetFocus
Exit Sub
End If
Next i
DoEvents
End Sub

Private Sub Form_Load()
Combo2.Text = "scsa.msg.yahoo.com"
Combo3.Text = "15"
MsgBox "Made in Tribute to Dazza aka Stooge the maker of Y-Hook Client and Owner of the Yhook forum, who Claims that i Yahooz Fynest am a Ripper, And cannot Code Ymsg parsing or a Client for MYSELF From Scratch... Well this right here Mr Dazza is, multiple rooms, multiple voices, amoungst the rest it can do. Is a Creative Idea i had, and has NOT been done before, so impossible to have been ripped! I made this V1.0 in roughly 24 hours... Over 2 days. Special Thanks to Yah-shit's DSP, The Bro Roomy, My Wigga EliteKris, And my boy Crazy owner of www.yahoo-load.com for there time testing when asked of them, to help me make this quickly and not spend all my time finding the lil things to make work right..!!! ..::MULTI ROOM/MULTI VOX's/MULTIPLE SPAMMING/NAMES COLLECTING::.. ", , "Dazza the maker of Y-Hook is a Lame Hater!"
End Sub

Private Sub List1_DblClick()
Call RemoveSelected(List1)
End Sub

Private Sub List1_Click()
SpamText.Text = List1
Text14.Text = List1
End Sub

Private Sub List10_DblClick()
Call RemoveSelected(List9)
End Sub

Private Sub List3_DblClick()
Call RemoveSelected(List3)
End Sub

Private Sub List3_Click()
Text1.Text = List3
Text20.Text = List3
End Sub

Private Sub List5_DblClick()
Call RemoveSelected(List5)
End Sub

Private Sub List6_DblClick()
Call RemoveSelected(List6)
End Sub

Private Sub List7_Click()
Text8.Text = List7
End Sub

Private Sub List8_Click()
Text8.Text = List8.Text
End Sub

Private Sub List8_DblClick()
Call RemoveSelected(List8)
End Sub

Private Sub List9_DblClick()
Call RemoveSelected(List9)
End Sub

Private Sub ListView1_DblClick()
On Error Resume Next
If ListView1.SelectedItem.SmallIcon = 4 Then
ListView1.SelectedItem.SmallIcon = 3
Form1.Text5.Text = "Not Processing Chat"
Form1.Text7.Text = "Not Processing Chat"
Form1.Text177.SelLength = 0
Form1.Text177.SelStart = Len(Form1.Text177)
Form1.Text177.SelBold = True: Form1.Text177.SelColor = &H404040
Form1.Text177.SelText = "Chat Recieving Ended! :"
Form1.Text177.SelBold = False: Form1.Text177.SelColor = vbBlack
Form1.Text177.SelText = "No Chat Bot Selected" & vbNewLine
DoEvents
Exit Sub
End If
If ListView1.SelectedItem.SmallIcon = 8 Then
ListView1.SelectedItem.SmallIcon = 3
Form1.Text17.SelLength = 0
Form1.Text17.SelStart = Len(Form1.Text17)
Form1.Text17.SelBold = True: Form1.Text17.SelColor = vbGreen
Form1.Text17.SelText = "Gawd Mode Disabled :"
Form1.Text17.SelBold = False: Form1.Text17.SelColor = vbBlack
Form1.Text17.SelText = "Bot Now Processing Packets Again!" & vbNewLine
DoEvents
Exit Sub
End If
End Sub

Private Sub ListView1_Click()
On Error Resume Next
Dim Q As Integer
For Q = 1 To Text25.Text
If ListView1.ListItems(Q).SmallIcon = 8 Then
Exit Sub
End If
If ListView1.ListItems(Q).SmallIcon = 4 Then
ListView1.ListItems(Q).SmallIcon = 3
End If
Next Q
DoEvents
If ListView1.SelectedItem.SmallIcon = 3 Then
ListView1.SelectedItem.SmallIcon = 4
Form1.Text5.Text = ListView1.SelectedItem.SubItems(1)
For Q = 1 To Text25.Text
If ListView1.ListItems(Q).SmallIcon = 4 Then
List7.Clear
List8.Clear
Form1.Text7.Text = RoomJoinedd(Q)
Form1.Text177.SelLength = 0
Form1.Text177.SelStart = Len(Form1.Text177)
Form1.Text177.SelBold = True: Form1.Text177.SelColor = &H404040
Form1.Text177.SelText = "Now Chating with " & YahooID(Q) & " : "
Form1.Text177.SelBold = False: Form1.Text177.SelColor = vbBlack
Form1.Text177.SelText = "In " & RoomJoinedd(Q) & vbNewLine
DoEvents
End If
Next Q
Exit Sub
Else
MsgBox "You Must Select a Bot In a Room To Chat With!", , "YF"
Exit Sub
End If
End Sub

Private Sub BlockedString()
On Error Resume Next
Status.Panels(11) = Status.Panels(11) + 1
End Sub

Private Sub socket_Connect(Index As Integer)
socket(Index).SendData Get_Key(YahooID(Index))
End Sub

Private Sub bugs(data As String)
On Error Resume Next
If Len(Text555.Text) > 150000 Then Text555.Text = ""
Text555.Text = Text555.Text & vbCrLf & "[IN] - " & data & " " & vbCrLf
Text555.SelStart = Len(Text555.Text)
DoEvents
End Sub

Private Sub socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
If ListView1.ListItems(Index).SmallIcon = 8 Then BlockedString: Exit Sub
If ListView1.ListItems(Index).SmallIcon = 3 And Check3.Value = 1 Then BlockedString: Exit Sub
If ListView1.ListItems(Index).SmallIcon = 4 And Check2.Value = 1 Then BlockedString: Exit Sub
socket(Index).GetData Buffer(Index)
Debug.Print Asc(Mid(Buffer(Index), 12, 1)) & " - " & Buffer(Index)
If Check4.Value = 1 Then
bugs Asc(Mid(Buffer(Index), 12, 1)) & " - " & Buffer(Index)
End If
FirstData Index, Buffer(Index)
Buffer(Index) = Empty
Exit Sub
End Sub

Public Sub FirstData(Index As Integer, BufferTuf As String)
On Error GoTo Buds
If ListView1.ListItems(Index).SmallIcon = 8 Then BlockedString: Exit Sub
If ListView1.ListItems(Index).SmallIcon = 3 And Check3.Value = 1 Then BlockedString: Exit Sub
If ListView1.ListItems(Index).SmallIcon = 4 And Check2.Value = 1 Then BlockedString: Exit Sub
Dim cstring As String
Dim H As Integer
Dim p As Integer
Status.Panels(9) = Status.Panels(9) + 1
IncomePacks(Index) = IncomePacks(Index) + 1
If ListView1.ListItems(Index).SmallIcon = 8 Then BlockedString: Exit Sub
If IncomePacks(Index) > 10 Then
If ListView1.ListItems(Index).SmallIcon = 3 Or ListView1.ListItems(Index).SmallIcon = 4 Then
If ListView1.ListItems(Index).SmallIcon = 8 Then BlockedString: Exit Sub
ListView1.ListItems(Index).SmallIcon = 8
Form1.Text17.SelLength = 0
Form1.Text17.SelStart = Len(Form1.Text17)
Form1.Text17.SelBold = True: Form1.Text17.SelColor = &H8000&
Form1.Text17.SelText = YahooID(Index) & " in " & RoomJoinedd(Index) & " Is Getting Booted: "
Form1.Text17.SelBold = False: Form1.Text17.SelColor = vbGreen
Form1.Text17.SelText = "Bot Now in Gawd Mode(Not Processing Packets)! You can Disable the Gawd Mode on the Bot by Double Clicking the Bot on Bot List!" & vbNewLine
DoEvents
Form1.Text177.SelLength = 0
Form1.Text177.SelStart = Len(Form1.Text177)
Form1.Text177.SelBold = True: Form1.Text177.SelColor = &H8000&
Form1.Text177.SelText = YahooID(Index) & " in " & RoomJoinedd(Index) & " Is Getting Booted: "
Form1.Text177.SelBold = False: Form1.Text177.SelColor = vbGreen
Form1.Text177.SelText = "Bot Now in Gawd Mode(Not Processing Packets)! You can Disable the Gawd Mode on the Bot by Double Clicking the Bot on Bot List!" & vbNewLine
DoEvents
BlockedString
BufferTuf = Empty
Exit Sub
End If
End If

Select Case Asc(Mid(BufferTuf, 12, 1))

Case Is = 87
cstring = Split(BufferTuf, "94À€")(1)
cstring = Split(cstring, "À€")(0)
Encrypt YahooID(Index), Password(Index), cstring, 1
socket(Index).SendData Log(YahooID(Index))
SessionKey(Index) = Mid(BufferTuf, 17, 4)
DoEvents
BufferTuf = Empty
Exit Sub

Case Is = 85
LoggedInBot(Index) = True
ListView1.ListItems(Index).SmallIcon = 1
Status.Panels(3) = Status.Panels(3) + 1
Status.Panels(1).Text = "Logged In:" & YahooID(Index)
LoopList YahooID(Index) & ":" & Password(Index), List4
List4.AddItem YahooID(Index) & ":" & Password(Index)
DoEvents
BufferTuf = Empty
Exit Sub

Case Is = 152
RoomJoining Index, BufferTuf
DoEvents
BufferTuf = Empty
Exit Sub

Case Is = 155
If ListView1.ListItems(Index).SmallIcon = 4 Then 'Chat bot only will doe this
RemoveUser BufferTuf, List8, RoomJoinedd(Index)
End If
DoEvents
BufferTuf = Empty
Exit Sub

End Select

IncommingData Index, BufferTuf
DoEvents
Buds:
BufferTuf = Empty
Exit Sub

End Sub

Public Sub RoomJoining(Index As Integer, BufferTuf As String)
On Error GoTo Buds
Dim Link As String, Cases As String
          
GetUsers Index, BufferTuf, List5, YahooID(Index), RoomJoinedd(Index)
       
If InStr(1, BufferTuf, "To help prevent spam") Then
Link = Split(BufferTuf, "http://ab.login")(1)
Link = Split(Link, ".jpg")(0)
Text2.Text = "http://ab.login" & Link & ".jpg"
Form2.WebBrowser1.Navigate Text2.Text
DoEvents
Form2.Show
Form2.Text1.Text = "Enter Captcha"
Form2.Text2.Text = "Captcha for " & YahooID(Index) & " to join " & RoomJoinedd(Index)
Form2.Caption = "Captcha for " & YahooID(Index)
Form2.TxtVerify.Text = ""
Form2.TxtVerify.SetFocus
BufferTuf = Empty
Exit Sub
End If

If Not InStr(1, BufferTuf, "ÿÿÿÿ", vbTextCompare) = 0 Then
If InStr(BufferTuf, "-35") Then
Form2.Text1.Text = "Room Full, Resend Captcha or Send Bot to Next Room!"
ListView1.ListItems(Index).SmallIcon = 1
End If
BufferTuf = Empty
Exit Sub

Else

If ListView1.ListItems(Index).SmallIcon = 7 Then
If InStr(BufferTuf, "129À€") And InStr(BufferTuf, "130À€") Then
If VoiceKey(Index) = "" Then VoiceKey(Index) = Parsing("130À€", "À€109À€", BufferTuf)
If RoomKey(Index) = "" Then RoomKey(Index) = Parsing("129À€", "À€130", BufferTuf)
End If
Status.Panels(1).Text = "Status:Joined Room"
Status.Panels(5) = Status.Panels(5) + 1
Form2.Text1.Text = "Joined Room!"
If RoomJoinedd(Index) = "" Then RoomJoinedd(Index) = Form1.Text1.Text
ListView1.ListItems(Index).SmallIcon = 3
Form1.Text17.SelLength = 0
Form1.Text17.SelStart = Len(Form1.Text17)
Form1.Text17.SelBold = True: Form1.Text17.SelColor = vbGreen
Form1.Text17.SelText = YahooID(Index)
Form1.Text17.SelBold = False: Form1.Text17.SelColor = vbBlack
Form1.Text17.SelText = " Joined " & RoomJoinedd(Index) & " (RKEY:" & RoomKey(Index) & " - Vkey:" & VoiceKey(Index) & ")" & vbNewLine
DoEvents
LoopList YahooID(Index) & " - " & RoomJoinedd(Index), List2
List2.AddItem YahooID(Index) & " - " & RoomJoinedd(Index)

Pause (1)

For i = 1 To Combo1.Text
If ListView1.ListItems(i).SmallIcon = 1 Then
On Error GoTo Dodge
List3.ListIndex = List3.ListIndex + 1
ListView1.ListItems(i).SmallIcon = 7
Status.Panels(1).Text = "Status:Joining Room"
socket(i).SendData JoinRoom(YahooID(i))
Pause 0.2
socket(i).SendData GotoRoom(YahooID(i), Form1.Text1.Text)
RoomJoinedd(i) = Form1.Text1.Text
Form2.Text1.Text = "Joining Next Bot To Next Room!"
DoEvents
BufferTuf = Empty
Exit Sub
End If
Next i
Form2.Text1.Text = "Finished Joining Bots!"

Pause (3)

Form2.Hide
BufferTuf = Empty
Exit Sub
Dodge:
Form2.Text1.Text = "Finished (No More Rooms In List To Join)"

Pause (3)

Form2.Hide
BufferTuf = Empty
Exit Sub
End If

BufferTuf = Empty
Exit Sub
End If

Buds:
BufferTuf = Empty
Exit Sub

End Sub

Public Sub IncommingData(Index As Integer, ICANTPARSE As String)
If Len(ICANTPARSE) > 900 Then
BlockedString
ICANTPARSE = Empty
Exit Sub
Else
SecondData Index, ICANTPARSE
ICANTPARSE = Empty
Exit Sub
End If
End Sub

Public Sub SecondData(Index As Integer, BufferTuf As String)
On Error Resume Next
Dim H As Integer
Dim p As Integer
Dim struser As String

Select Case Asc(Mid(BufferTuf, 12, 1))

Case Is = 6
On Error GoTo ErrorMurder
struser = Parsing("4À€", "À€", BufferTuf)
If struser = "" Then BlockedString: Exit Sub
If struser = " " Then BlockedString: Exit Sub
If Len(struser) < 4 Then BlockedString: Exit Sub
If InStr(struser, " ") Then BlockedString: Exit Sub
If InStr(struser, "À€") Then BlockedString: Exit Sub
LoopList struser, List6
List6.AddItem struser
Status.Panels(1).Text = "PM From:" & struser
Form1.Text17.SelLength = 0
Form1.Text17.SelStart = Len(Form1.Text17)
Form1.Text17.SelBold = True: Form1.Text17.SelColor = &HC0C0C0
Form1.Text17.SelText = struser
Form1.Text17.SelBold = False: Form1.Text17.SelColor = vbBlack
Form1.Text17.SelText = " Attempted to PM " & YahooID(Index) & " In Room " & RoomJoinedd(Index) & vbNewLine
Form1.Text17.SelStart = Len(Form1.Text17)
DoEvents
BufferTuf = Empty
Exit Sub
ErrorMurder:
BlockedString
BufferTuf = Empty
Exit Sub

Case Is = 150
If ListView1.ListItems(Index).SmallIcon = 7 Then
socket(Index).SendData GotoRoom(YahooID(Index), RoomJoinedd(Index))
End If
BlockedString
BufferTuf = Empty
Exit Sub


Case Is = 168
On Error GoTo Bum
If Check1.Value = 1 Then GoTo Bum
Dim RN As String
RN = Parsing("4À€", "À€", BufferTuf)
RoomJoinedd(Index) = RN
ParseChat Index, BufferTuf, RoomJoinedd(Index)
DoEvents
BufferTuf = Empty
Exit Sub
Bum:
BlockedString
BufferTuf = Empty
Exit Sub

Case Is = 84
If LoggedInBot(Index) = True Then Exit Sub
On Error GoTo figz
ListView1.ListItems(Index).SmallIcon = 2
LoggedInBot(Index) = False
Status.Panels(1).Text = "Failed:" & YahooID(Index)
socket(Index).Close
DoEvents
BufferTuf = Empty
Exit Sub
figz:
BlockedString
BufferTuf = Empty
Exit Sub

End Select

Buds:
BlockedString
BufferTuf = Empty
Exit Sub
End Sub

Private Sub socket_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Status.Panels(1).Text = "Socket Error: " & YahooID(Index)
ListView1.ListItems(Index).SmallIcon = 2
socket(Index).Close
LoggedInBot(Index) = False
End Sub

Private Sub Text14_DblClick()
Text14 = ""
End Sub

Private Sub Text17_DblClick()
Text17.Text = ""
End Sub

Private Sub Text177_DblClick()
Text177 = ""
End Sub

Private Sub Text20_DblClick()
Text20 = ""
End Sub

Private Sub Text21_DblClick()
Text21 = ""
End Sub

Private Sub Text6_DblClick()
Text6.Text = ""
End Sub

Sub Text6_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
On Error Resume Next
For i = 1 To Text25.Text
If ListView1.ListItems(i).SmallIcon = 4 Then
If Check6.Value = 1 Then
socket(i).SendData SendChat(Text5, Text7, Text18.Text, Text6.Text)
Text177.SelLength = 0
Text177.SelStart = Len(Text177)
Text177.SelBold = True: Text177.SelColor = vbBlue
Text177.SelText = YahooID(i) & ": "
Text177.SelBold = False: Text177.SelColor = vbBlack
Text177.SelText = Text6.Text & vbNewLine
Text177.SelStart = Len(Text177)
Else
socket(i).SendData SendChat(Text5, Text7, Text13 & Text18, Text6)
Text177.SelLength = 0
Text177.SelStart = Len(Text177)
Text177.SelBold = True: Text177.SelColor = vbBlue
Text177.SelText = YahooID(i) & "(YF-MRC): "
Text177.SelBold = False: Text177.SelColor = vbBlack
Text177.SelText = Text6.Text & vbNewLine
Text177.SelStart = Len(Text177)
End If
DoEvents
Text6.Text = ""
Text6.SetFocus
Exit Sub
End If
Next i
DoEvents

End If
End Sub

Private Sub Timer1_Timer()
FrameSpam.Caption = "Spam Messages List : " & List1.ListCount
Frame1.Caption = "Rooms For Bots To Join : " & List3.ListCount
Label1.Caption = List5.ListCount
Label2.Caption = List6.ListCount
Label10.Caption = List9.ListCount
Label11.Caption = List10.ListCount
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If Timer2 = False Then Exit Sub
Status.Panels(1).Text = "Changing Spam Message"
If List1.ListIndex = List1.ListCount - 1 Then
List1.ListIndex = 0
Else
List1.ListIndex = List1.ListIndex + 1
End If
If Timer2 = False Then Exit Sub
Form1.Text17.SelLength = 0
Form1.Text17.SelStart = Len(Form1.Text17)
Form1.Text17.SelBold = True: Form1.Text17.SelColor = vbBlue
Form1.Text17.SelText = "Message Bots Are Spamming:"
Form1.Text17.SelBold = False: Form1.Text17.SelColor = vbBlack
Form1.Text17.SelText = SpamText.Text & vbNewLine
DoEvents
If Timer2 = False Then Exit Sub
For Jj = 1 To Combo1.Text
If Timer2 = False Then Exit Sub
If ListView1.ListItems(Jj).SmallIcon = 3 Then
If Check6.Value = 1 Then
socket(Jj).SendData SendChat(YahooID(Jj), RoomJoinedd(Jj), Text18.Text, SpamText.Text)
Else
socket(Jj).SendData SendChat(YahooID(Jj), RoomJoinedd(Jj), Text13 & Text18, SpamText.Text)
End If
Status.Panels(7) = Status.Panels(7) + 1
Pause (0.5)
GoTo Done
End If
If ListView1.ListItems(Jj).SmallIcon = 4 Then
If Check6.Value = 1 Then
socket(Jj).SendData SendChat(YahooID(Jj), RoomJoinedd(Jj), Text18.Text, SpamText.Text)
Text177.SelLength = 0
Text177.SelStart = Len(Text177)
Text177.SelBold = True: Text177.SelColor = vbBlue
Text177.SelText = YahooID(Jj) & ": "
Text177.SelBold = False: Text177.SelColor = vbBlack
Text177.SelText = SpamText.Text & vbNewLine
Text177.SelStart = Len(Text177)
Else
socket(Jj).SendData SendChat(YahooID(Jj), RoomJoinedd(Jj), Text13 & Text18, SpamText.Text)
Text177.SelLength = 0
Text177.SelStart = Len(Text177)
Text177.SelBold = True: Text177.SelColor = vbBlue
Text177.SelText = YahooID(Jj) & "(YF-MRC): "
Text177.SelBold = False: Text177.SelColor = vbBlack
Text177.SelText = SpamText.Text & vbNewLine
Text177.SelStart = Len(Text177)
End If
Status.Panels(7) = Status.Panels(7) + 1
DoEvents
Pause (0.5)
GoTo Done
End If
If ListView1.ListItems(Jj).SmallIcon = 8 Then
socket(Jj).SendData SendChat(YahooID(Jj), RoomJoinedd(Jj), Text18.Text, SpamText.Text)
Status.Panels(7) = Status.Panels(7) + 1
Pause (0.5)
End If
Done:
Next Jj
DoEvents
If Timer2 = False Then Exit Sub
Status.Panels(1).Text = "Pausing for" & Text4.Text & " Secs..."
Pause (Text4.Text)
If Timer2 = False Then Exit Sub
End Sub

Private Sub Timer4_Timer()
On Error Resume Next
Dim Logic As Integer
For Logic = 0 To 50
IncomePacks(Logic) = 0
Next Logic
End Sub

Private Sub CapSck_Connect()
CapSck.SendData Heading(Text2.Text, Form2.TxtVerify.Text)
End Sub

Function Heading(Question As String, Answer As String) As String
Heading = "POST /captcha1 HTTP/1.1" & vbCrLf & _
"Accept: */*" & vbCrLf & _
"Referer: YF" & vbCrLf & _
"Accept-Language: en-us" & vbCrLf & _
"Content-Type: application/x-www-form-urlencoded" & vbCrLf & _
"Accept-Encoding: gzip, deflate" & vbCrLf & _
"User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; .NET CLR 1.1.4322; .NET CLR 2.0.50727)" & vbCrLf & _
"Host: captcha.chat.yahoo.com" & vbCrLf & _
"Content-Length: " & Len("question=" & Text2.Text & "&.intl=us&answer=" & Form2.TxtVerify.Text) & vbCrLf & _
"Connection: Keep-Alive" & vbCrLf & _
"Cache-Control: no-cache" & vbCrLf & _
"Cookie: " & vbCrLf & vbCrLf & "question=" & Text2.Text & "&.intl=us&answer=" & Form2.TxtVerify.Text
End Function

Private Sub CapSck_DataArrival(ByVal bytesTotal As Long)
Dim Dat As String
CapSck.GetData Dat
Debug.Print Dat
If InStr(Dat, "close?.intl") Then
Form2.Text1.Text = "Correct!"
CapSck.Close
Exit Sub
ElseIf InStr(Dat, "?tryagain") Then
Form2.Text1.Text = "Try Again"
CapSck.Close
Exit Sub
ElseIf InStr(Dat, "?exceeded") Then
Form2.Text1.Text = "Exceeded Amount Of Trys, Try Next Room!"
CapSck.Close
Exit Sub
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End
End Sub
