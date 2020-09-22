Attribute VB_Name = "modroom"
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


Private str As String
Private p As Integer


'"9À€y4h00z.fyn3st_d34n0À€1À€y4h00z.fyn3st_d34n0À€6À€abcdeÀ€98À€usÀ€135À€ym9.0.0.2034À€"
Public Function JoinRoom(User As String) As String
Dim Packet As String
Packet = "109À€" & User & "À€1À€" & User & "À€6À€abcdeÀ€98À€usÀ€135À€ym9.0.0.2034À€"
JoinRoom = YFHeader("96", Packet)
End Function
'"1À€y4h00z.fyn3st_d34n0À€104À€Auckland Global Chat:2À€129À€1168À€62À€2À€"
Public Function GotoRoom(User As String, Room As String) As String
Dim Packet As String
Packet = "1À€" & User & "À€104À€" & Room & "À€129À€0À€62À€2À€"
'The change i made to "0" Makes NO RoomKey Needed! This Program get the RoomKeys for all the Differents Rooms its self as joining each..
'Packet = "1À€" & User & "À€62À€À€2À€À€104À€" & Room & "À€129À€0À€"
GotoRoom = YFHeader("98", Packet)
End Function
'"1À€y4h00z.fyn3st_d34n0À€1005À€77619688À€"
Public Function LeaveRoom(ByVal Username As String) As String
Dim Packet As String
Pck = "1À€" & Username & "À€1005À€77619688À€"
LeaveRoom = YFHeader("A0", Packet)
End Function

'"1À€y4h00z.fyn3st_d34n0À€104À€Auckland Global Chat:2À€124À€1À€117À€[1m[31m<font face=""Comic Sans MS"" size=""12"">hello Room Test SniffÀ€"
Public Function SendChat(WhoFrom As String, Room As String, Size As String, message As String) As String
Dim Pck As String
Pck = "1À€" & WhoFrom & "À€104À€" & Room & "À€124À€1À€117À€" & Size & message & "À€"
'Pck = "1À€" & WhoFrom & "À€104À€" & Room & "À€117À€" & Size & message & "À€124À€1À€"
SendChat = YFHeader("A8", Pck)
Debug.Print SendChat
End Function

'"13À€1À€104À€Auckland Global Chat:2À€105À€Discuss the area with visitors, residents, and expats Visit http://au.yahoo.com/À€108À€13À€126À€328704À€128À€1037À€129À€1168À€130À€kNIQJBeMml5Y7EzpaLzQmqiGIHczB_K.M-À€109À€abbasachillesÀ€110À€0À€111À€neuterÀ€113À€1024À€109À€gurlie_serena_184À€110À€22À€111À€femaleÀ€113À€66576À€109À€girl_deanna_114À€110À€23À€111À€femaleÀ€113À€66576À€109À€ladyblue_0489À€110À€0À€111À€neuterÀ€113À€1024À€109À€halfpipedisimhÀ€110À€0À€111À€neuterÀ€113À€1040À€109À€cutie_ann_351À€110À€0À€111À€neuterÀ€113À€1040À€109À€allen_raph21À€110À€0À€111À€neuterÀ€113À€1024À€109À€abbasgudawalaÀ€110À€0À€111À€neuterÀ€113À€1024À€109À€castillo_extremeÀ€110À€0À€111À€neuterÀ€113À€1024À€109À€jillianm37xouÀ€110À€0À€111À€neuterÀ€113À€1024À€109À€scoutteeÀ€110À€0À€141À€scoutteeÀ€142À€new zealandÀ€111À€femaleÀ€113À€66560À€109À€y4h00z.fyn3st_d34n0À€110À€23À€141À€                    À€142À€NewzealandÀ€111À€maleÀ€113À€33792À€109À€YahooÀ€141À€Messenger Chat AdminÀ€"
'"4À€Auckland Global Chat:2À€105À€Discuss the area with visitors, residents, and expats Visit http://au.yahoo.com/À€108À€1À€109À€zia007zohaibÀ€113À€1024À€"
Public Function GetUsers(Index As Integer, data As String, lst As listbox, Id As String, Room As String)
On Error Resume Next
Dim p As Integer
Dim ChatUsers() As String, Users As Integer
If InStr(data, "129À€") And InStr(data, "130À€") Then
If VoiceKey(Index) = "" Then VoiceKey(Index) = Parsing("130À€", "À€109À€", data)
If RoomKey(Index) = "" Then RoomKey(Index) = Parsing("129À€", "À€130", data)
End If
ChatUsers = Split(data, "À€109À€")
For Users = 0 To UBound(ChatUsers)
ChatUsers(Users) = Split(ChatUsers(Users), "À€")(0)
LoopList ChatUsers(Users), lst
If ChatUsers(Users) = Id Then GoTo NoAd
If ChatUsers(Users) = "" Then GoTo NoAd
If Len(ChatUsers(Users)) < 4 Then GoTo NoAd
If InStr(LCase(ChatUsers(Users)), "yahoo") Then GoTo NoAd
If InStr(LCase(ChatUsers(Users)), "ymsg") Then GoTo NoAd
If InStr(ChatUsers(Users), " ") Then GoTo NoAd
If InStr(LCase(ChatUsers(Users)), LCase(YahooID(Index))) Then GoTo NoAd
lst.AddItem ChatUsers(Users)
If Len(Text17.Text) > 150000 Then Text17.Text = ""
If Len(Text177.Text) > 150000 Then Text177.Text = ""
If Form1.ListView1.ListItems(Index).SmallIcon = 4 Then
Form1.Text177.SelLength = 0
Form1.Text177.SelStart = Len(Form1.Text177)
Form1.Text177.SelBold = True: Form1.Text177.SelColor = vbRed
Form1.Text177.SelText = ChatUsers(Users)
Form1.Text177.SelBold = False: Form1.Text177.SelColor = vbBlack
Form1.Text177.SelText = " Joined " & Room & vbNewLine
Form1.Text177.SelStart = Len(Form1.Text177)

Else

Form1.Text17.SelLength = 0
Form1.Text17.SelStart = Len(Form1.Text17)
Form1.Text17.SelBold = True: Form1.Text17.SelColor = vbRed
Form1.Text17.SelText = ChatUsers(Users)
Form1.Text17.SelBold = False: Form1.Text17.SelColor = vbBlack
Form1.Text17.SelText = " Joined " & Room & vbNewLine
Form1.Text17.SelStart = Len(Form1.Text17)
End If
DoEvents
If InStr(Users, "@") Then
ATcheck ChatUsers(Users)
Else
illycheck ChatUsers(Users)
End If
DoEvents
NoAd:
  Next Users
  
End Function

Public Function ATcheck(Id As String)
On Error Resume Next
If InStr(Id, "@") Then
LoopList Id, Form1.List10
Form1.List10.AddItem Id
Exit Function
End If
End Function

Public Function illycheck(Id As String)
On Error Resume Next
If InStr(Id, "@") Then
LoopList Id, Form1.List10
Form1.List10.AddItem Id
Exit Function
End If
If InStr(Id, "+") Then
LoopList Id, Form1.List9
Form1.List9.AddItem Id
Exit Function
End If
If InStr(Id, "--") Then
LoopList Id, Form1.List9
Form1.List9.AddItem Id
Exit Function
End If
If InStr(Id, "__") Then
LoopList Id, Form1.List9
Form1.List9.AddItem Id
Exit Function
End If
If InStr(Id, "..") Then
LoopList Id, Form1.List9
Form1.List9.AddItem Id
Exit Function
End If
If InStr(Id, "-") Then
LoopList Id, Form1.List9
Form1.List9.AddItem Id
Exit Function
End If
If LCase(Left(Id, 1)) = "_" Then
LoopList Id, Form1.List9
Form1.List9.AddItem Id
Exit Function
End If
End Function

Public Function LoopList(data, list As listbox)
Dim die As Integer
For die = 0 To list.ListCount - 1
If data = list.list(die) Then
list.RemoveItem die
End If
Next die
End Function

'"4À€Auckland Global Chat:2À€108À€1À€109À€aya0466À€113À€1024À€"
Public Function RemoveUser(data As String, lst As listbox, Room As String)
On Error Resume Next
Dim struser As String
struser = Parsing("109À€", "À€", data)
LoopList struser, lst
DoEvents
Form1.Text177.SelLength = 0
Form1.Text177.SelStart = Len(Form1.Text177)
Form1.Text177.SelBold = True: Form1.Text177.SelColor = vbRed
Form1.Text177.SelText = struser
Form1.Text177.SelBold = False: Form1.Text177.SelColor = vbBlack
Form1.Text177.SelText = " Leaft " & Room & vbNewLine
Form1.Text177.SelStart = Len(Form1.Text177)
DoEvents
End Function

Public Function Parsing(LL As String, RR As String, data As String) As String
p = InStr(data, LL)
If p > 0 Then
str = Mid(data, p + Len(LL))
End If
p = InStr(str, RR)
If p > 0 Then
Parsing = Left(str, p - 1)
Else
Parsing = str
End If
End Function
            
'Coded from watching the Immediate from VB6 and via Debug.Print
'"4À€Auckland Global Chat:2À€109À€gurlie_serena_184À€117À€[2m[#8207e1m<font size=""18"" face=""Lucida Sans"">hey guyzÀ€124À€1À€"
Public Function ParseChat(Index As Integer, strData As String, Room As String)
On Error Resume Next
Dim struser As String
Dim strMessage As String
Dim INF As String

struser = Parsing("109À€", "À€", strData)
If struser = "" Then GoTo Sweet2
    
If struser = "" Then GoTo sweet
If struser = YahooID(Index) Then GoTo Sweet2
If InStr(struser, " ") Then GoTo Sweet2
If InStr(struser, "Yahoo") Then GoTo Sweet2
If InStr(struser, "YMSG") Then GoTo Sweet2
Dim H As Integer
For H = 0 To Form1.List5.ListCount - 1
If InStr(struser, Form1.List5.list(H)) Then GoTo sweet
Next H
LoopList struser, Form1.List5
DoEvents
Form1.List5.AddItem struser
Form1.Text17.SelLength = 0
Form1.Text17.SelStart = Len(Form1.Text17)
Form1.Text17.SelBold = True: Form1.Text17.SelColor = vbRed
Form1.Text17.SelText = struser
Form1.Text17.SelBold = False: Form1.Text17.SelColor = vbBlack
Form1.Text17.SelText = " Joined " & Room & vbNewLine
Form1.Text17.SelStart = Len(Form1.Text17)
DoEvents
illycheck struser
DoEvents
ATcheck struser
sweet:
DoEvents

If Form1.ListView1.ListItems(Index).SmallIcon = 4 Then
Dim RN As String
RN = Parsing("4À€", "À€", strData)
If Not Roomy = RN Then
Form1.Text7.Text = RN
RoomJoinedd(Index) = RN
End If

strMessage = Parsing("117À€", "À€", strData)

If InStr(strMessage, "</font") Or InStr(strMessage, "</Font") Or InStr(strMessage, "</FONT") Then
If InStr(strMessage, Form1.Text19.Text) Then
strMessage = Parsing(Form1.Text19.Text, "</", strMessage)
Else
strMessage = Parsing(">", "", strMessage)
End If
End If

'1 Hour into coding Text Processing, And this is what i got so Far...
'Yes its Basic Text Stripping right now, and is Breaking it Down to Raw.. !
'but i will work on it bit by bit..
'Strictly going for Raw effect, but stripped of everthing to just there chat text.
'Stil got Color Codes to Strip, but otherwise works reasonably well.
'Remember i have made this program over a Time Frame of Just 24 hours!!
If InStr(strMessage, "<font") Or InStr(strMessage, "<Font") Or InStr(strMessage, "<FONT") Then
If InStr(strMessage, "</") Then
strMessage = Parsing(Form1.Text19.Text, "</", strMessage)
ElseIf InStr(strMessage, Form1.Text19.Text) Then
strMessage = Parsing(Form1.Text19.Text, "<", strMessage)
Else
strMessage = Parsing(Form1.Text19.Text, "", strMessage)
End If
End If
If InStr(strMessage, "<fade") Or InStr(strMessage, "<Fade") Or InStr(strMessage, "<FADE") Then
strMessage = Parsing(">", "</", strMessage)
End If
If InStr(strMessage, "<alt") Or InStr(strMessage, "<Alt") Or InStr(strMessage, "<ALT") Then
Dim Op As String
If InStr(strMessage, "<ALT") Then
Op = Parsing("<ALT ", ">", strMessage)
strMessage = Parsing(Op & ">", "</ALT", strMessage)
End If
If InStr(strMessage, "<Alt") Then
Op = Parsing("<ALT ", ">", strMessage)
strMessage = Parsing(Op & ">", "</Alt", strMessage)
End If
If InStr(strMessage, "<alt") Then
Op = Parsing("<ALT ", ">", strMessage)
strMessage = Parsing(Op & ">", "</alt", strMessage)
End If
End If
If InStr(strMessage, "<font") Or InStr(strMessage, "<Font") Or InStr(strMessage, "<FONT") Then
strMessage = Parsing(">", "<f", strMessage)
End If
If InStr(strMessage, "</") Then
strMessage = Parsing("", "</", strMessage)
End If
If strMessage = "" Then
strMessage = Parsing("117À€", "À€", strData)
End If
If strMessage = "" Then
Exit Function
End If

'YF Design for INF Tag Reading.. Client name only..
'Change the " " Space in parsing to a ">" and it will get full INF Data's..
If InStr(strData, "ID:") Then
INF = Parsing("ID:", " ", strData)
Form1.Text177.SelLength = 0
Form1.Text177.SelStart = Len(Form1.Text177)
Form1.Text177.SelBold = True: Form1.Text177.SelColor = &H800000
Form1.Text177.SelText = struser & "(" & INF & "): "
Form1.Text177.SelBold = False: Form1.Text177.SelColor = vbBlack
Form1.Text177.SelText = strMessage & vbNewLine
LoopList struser, Form1.List7
Form1.List7.AddItem struser
ElseIf InStr(strData, "Client:") Then
INF = Parsing("Client:", " ", strData)
Form1.Text177.SelLength = 0
Form1.Text177.SelStart = Len(Form1.Text177)
Form1.Text177.SelBold = True: Form1.Text177.SelColor = &H800000
Form1.Text177.SelText = struser & "(" & INF & "): "
Form1.Text177.SelBold = False: Form1.Text177.SelColor = vbBlack
Form1.Text177.SelText = strMessage & vbNewLine
LoopList struser, Form1.List7
Form1.List7.AddItem struser
ElseIf InStr(strData, "<font INF ") Then
INF = Parsing("INF ", " ", strData)
Form1.Text177.SelLength = 0
Form1.Text177.SelStart = Len(Form1.Text177)
Form1.Text177.SelBold = True: Form1.Text177.SelColor = &H800000
Form1.Text177.SelText = struser & "(" & INF & "): "
Form1.Text177.SelBold = False: Form1.Text177.SelColor = vbBlack
Form1.Text177.SelText = strMessage & vbNewLine
LoopList struser, Form1.List7
Form1.List7.AddItem struser
ElseIf InStr(strData, "yimg") Then
INF = "Y-Tunnel"
Form1.Text177.SelLength = 0
Form1.Text177.SelStart = Len(Form1.Text177)
Form1.Text177.SelBold = True: Form1.Text177.SelColor = &H800000
Form1.Text177.SelText = struser & "(" & INF & "): "
Form1.Text177.SelBold = False: Form1.Text177.SelColor = vbBlack
Form1.Text177.SelText = strMessage & vbNewLine
LoopList struser, Form1.List7
Form1.List7.AddItem struser
Else
Form1.Text177.SelLength = 0
Form1.Text177.SelStart = Len(Form1.Text177)
Form1.Text177.SelBold = True: Form1.Text177.SelColor = &H800000
Form1.Text177.SelText = struser & ": "
Form1.Text177.SelBold = False: Form1.Text177.SelColor = vbBlack
Form1.Text177.SelText = strMessage & vbNewLine
LoopList struser, Form1.List8
Form1.List8.AddItem struser
End If
Form1.Text177.SelStart = Len(Form1.Text177)
Exit Function
DoEvents
End If
Exit Function
   
Sweet2:
Form1.Status.Panels(11) = Form1.Status.Panels(11) + 1
   Exit Function

End Function
