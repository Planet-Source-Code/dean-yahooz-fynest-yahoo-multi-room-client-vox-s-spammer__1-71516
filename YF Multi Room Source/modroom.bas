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


'"9¿Äy4h00z.fyn3st_d34n0¿Ä1¿Äy4h00z.fyn3st_d34n0¿Ä6¿Äabcde¿Ä98¿Äus¿Ä135¿Äym9.0.0.2034¿Ä"
Public Function JoinRoom(User As String) As String
Dim Packet As String
Packet = "109¿Ä" & User & "¿Ä1¿Ä" & User & "¿Ä6¿Äabcde¿Ä98¿Äus¿Ä135¿Äym9.0.0.2034¿Ä"
JoinRoom = YFHeader("96", Packet)
End Function
'"1¿Äy4h00z.fyn3st_d34n0¿Ä104¿ÄAuckland Global Chat:2¿Ä129¿Ä1168¿Ä62¿Ä2¿Ä"
Public Function GotoRoom(User As String, Room As String) As String
Dim Packet As String
Packet = "1¿Ä" & User & "¿Ä104¿Ä" & Room & "¿Ä129¿Ä0¿Ä62¿Ä2¿Ä"
'The change i made to "0" Makes NO RoomKey Needed! This Program get the RoomKeys for all the Differents Rooms its self as joining each..
'Packet = "1¿Ä" & User & "¿Ä62¿Ä¿Ä2¿Ä¿Ä104¿Ä" & Room & "¿Ä129¿Ä0¿Ä"
GotoRoom = YFHeader("98", Packet)
End Function
'"1¿Äy4h00z.fyn3st_d34n0¿Ä1005¿Ä77619688¿Ä"
Public Function LeaveRoom(ByVal Username As String) As String
Dim Packet As String
Pck = "1¿Ä" & Username & "¿Ä1005¿Ä77619688¿Ä"
LeaveRoom = YFHeader("A0", Packet)
End Function

'"1¿Äy4h00z.fyn3st_d34n0¿Ä104¿ÄAuckland Global Chat:2¿Ä124¿Ä1¿Ä117¿Ä[1m[31m<font face=""Comic Sans MS"" size=""12"">hello Room Test Sniff¿Ä"
Public Function SendChat(WhoFrom As String, Room As String, Size As String, message As String) As String
Dim Pck As String
Pck = "1¿Ä" & WhoFrom & "¿Ä104¿Ä" & Room & "¿Ä124¿Ä1¿Ä117¿Ä" & Size & message & "¿Ä"
'Pck = "1¿Ä" & WhoFrom & "¿Ä104¿Ä" & Room & "¿Ä117¿Ä" & Size & message & "¿Ä124¿Ä1¿Ä"
SendChat = YFHeader("A8", Pck)
Debug.Print SendChat
End Function

'"13¿Ä1¿Ä104¿ÄAuckland Global Chat:2¿Ä105¿ÄDiscuss the area with visitors, residents, and expats Visit http://au.yahoo.com/¿Ä108¿Ä13¿Ä126¿Ä328704¿Ä128¿Ä1037¿Ä129¿Ä1168¿Ä130¿ÄkNIQJBeMml5Y7EzpaLzQmqiGIHczB_K.M-¿Ä109¿Äabbasachilles¿Ä110¿Ä0¿Ä111¿Äneuter¿Ä113¿Ä1024¿Ä109¿Ägurlie_serena_184¿Ä110¿Ä22¿Ä111¿Äfemale¿Ä113¿Ä66576¿Ä109¿Ägirl_deanna_114¿Ä110¿Ä23¿Ä111¿Äfemale¿Ä113¿Ä66576¿Ä109¿Äladyblue_0489¿Ä110¿Ä0¿Ä111¿Äneuter¿Ä113¿Ä1024¿Ä109¿Ähalfpipedisimh¿Ä110¿Ä0¿Ä111¿Äneuter¿Ä113¿Ä1040¿Ä109¿Äcutie_ann_351¿Ä110¿Ä0¿Ä111¿Äneuter¿Ä113¿Ä1040¿Ä109¿Äallen_raph21¿Ä110¿Ä0¿Ä111¿Äneuter¿Ä113¿Ä1024¿Ä109¿Äabbasgudawala¿Ä110¿Ä0¿Ä111¿Äneuter¿Ä113¿Ä1024¿Ä109¿Äcastillo_extreme¿Ä110¿Ä0¿Ä111¿Äneuter¿Ä113¿Ä1024¿Ä109¿Äjillianm37xou¿Ä110¿Ä0¿Ä111¿Äneuter¿Ä113¿Ä1024¿Ä109¿Äscouttee¿Ä110¿Ä0¿Ä141¿Äscouttee¿Ä142¿Änew zealand¿Ä111¿Äfemale¿Ä113¿Ä66560¿Ä109¿Äy4h00z.fyn3st_d34n0¿Ä110¿Ä23¿Ä141¿Ä                    ¿Ä142¿ÄNewzealand¿Ä111¿Ämale¿Ä113¿Ä33792¿Ä109¿ÄYahoo¿Ä141¿ÄMessenger Chat Admin¿Ä"
'"4¿ÄAuckland Global Chat:2¿Ä105¿ÄDiscuss the area with visitors, residents, and expats Visit http://au.yahoo.com/¿Ä108¿Ä1¿Ä109¿Äzia007zohaib¿Ä113¿Ä1024¿Ä"
Public Function GetUsers(Index As Integer, data As String, lst As listbox, Id As String, Room As String)
On Error Resume Next
Dim p As Integer
Dim ChatUsers() As String, Users As Integer
If InStr(data, "129¿Ä") And InStr(data, "130¿Ä") Then
If VoiceKey(Index) = "" Then VoiceKey(Index) = Parsing("130¿Ä", "¿Ä109¿Ä", data)
If RoomKey(Index) = "" Then RoomKey(Index) = Parsing("129¿Ä", "¿Ä130", data)
End If
ChatUsers = Split(data, "¿Ä109¿Ä")
For Users = 0 To UBound(ChatUsers)
ChatUsers(Users) = Split(ChatUsers(Users), "¿Ä")(0)
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

'"4¿ÄAuckland Global Chat:2¿Ä108¿Ä1¿Ä109¿Äaya0466¿Ä113¿Ä1024¿Ä"
Public Function RemoveUser(data As String, lst As listbox, Room As String)
On Error Resume Next
Dim struser As String
struser = Parsing("109¿Ä", "¿Ä", data)
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
'"4¿ÄAuckland Global Chat:2¿Ä109¿Ägurlie_serena_184¿Ä117¿Ä[2m[#8207e1m<font size=""18"" face=""Lucida Sans"">hey guyz¿Ä124¿Ä1¿Ä"
Public Function ParseChat(Index As Integer, strData As String, Room As String)
On Error Resume Next
Dim struser As String
Dim strMessage As String
Dim INF As String

struser = Parsing("109¿Ä", "¿Ä", strData)
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
RN = Parsing("4¿Ä", "¿Ä", strData)
If Not Roomy = RN Then
Form1.Text7.Text = RN
RoomJoinedd(Index) = RN
End If

strMessage = Parsing("117¿Ä", "¿Ä", strData)

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
strMessage = Parsing("117¿Ä", "¿Ä", strData)
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
