Attribute VB_Name = "Login"
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


Public Crip(1) As String, SessionKey(0 To 50) As String, RoomJoinedd(0 To 50) As String, YahooID(0 To 50) As String, Password(0 To 50) As String, VoiceKey(0 To 50) As String, RoomKey(0 To 50) As String, Buffer(0 To 50) As String, Encryption As String
Private Declare Function YMSG12_ScriptedMind_Encrypt Lib "YMSG12ENCRYPT.dll" (ByVal Username As String, ByVal Password As String, ByVal Seed As String, ByVal result_6 As String, ByVal result_96 As String, intt As Long) As Boolean

'Latest Yahoo 9 Login - "1À€y4h00z.fyn3st_d34n0À€0À€y4h00z.fyn3st_d34n0À€277À€v=1&n=9pn8ngvpjrkd2&l=ou7qqp.5odtij_3tudq/o&p=m2k0c1p012000000&r=if&lg=en-NZ&intl=nz&np=1; path=/; domain=.yahoo.comÀ€278À€z=7yIQJB74dQJB8eAKFz1SFf0MzI1BjZPM08zMzQzMjI-&a=QAE&sk=DAA4CN2XgqIX3w&ks=EAA_TZ0NnZGXHPff5bJXEwUHg--~C&d=c2wBTkRVeUFURTRORGcwTkRNME5UVS0BYQFRQUUBZwFITlhETjZGUFlRVkZMR1RGWVhQSFNUTlpXNAF6egE3eUlRSkJnV0EBdGlwAUtUTFlPRA--; path=/; domain=.yahoo.comÀ€307À€1yEFsoy.R432klvq7mkHbA--À€244À€4194239À€2À€y4h00z.fyn3st_d34n0À€2À€1À€59À€B    bv79nmt4b44jr&b=4&d=AYaBpJppYEKS4UTsP2_zgFdd5sPXVfC3YfUmiQ--&s=tmÀ€59À€F    a=365sMRoMvTDwtTTJRz9YNsi5mL0Z87AbRRl.hNZmQ.br4P3q3c6Z4_42UV_63JfAcQl5LVk-&b=9OWr
'To Long For 1 Line!  -  À€59À€T    z=ADIQJBAJdQJBUF5we1ox4yZMzI1BjZPM08zMzQzMjI-&a=QAE&sk=DAAIdP8MTcAHfT&ks=EAA2ZOkndSpiK2p1OWpHuqzkg--~C&d=c2wBTkRVeUFURTRORGcwTkRNME5UVS0BYQFRQUUBZwFITlhETjZGUFlRVkZMR1RGWVhQSFNUTlpXNAF6egFBRElRSkJnV0EBdGlwAUtUTFlPRA--À€59À€Y    v=1&n=9pn8ngvpjrkd2&l=ou7qqp.5odtij_3tudq/o&p=m2k0c1p012000000&r=if&lg=en-NZ&intl=nz&np=1À€98À€usÀ€135À€9.0.0.234 À€"
'^^^
'The ONLY PART I HAVE taken to Use Because im Not Using SSL - À€98À€usÀ€135À€9.0.0.234 À€  << AND >> À€2À€1À€
'9.0.0.234  << this part i also will use for Voice! They Dont have to Match, but just to be up to date!
'Header("54", Packet)
'ORIGINAL LOGIN PACKET BELOW IS MOSTLY SAME AS MY OLD YAHOO 8 LOGIN JUST UPDATED FOR THIS PROGRAM!
Public Function Log(YahooID As String) As String
Dim Packet As String
Packet = "6À€" & Crip(0) & "À€96À€" & Crip(1) & "À€0À€" & YahooID & "À€2À€" & YahooID & "À€2À€1À€1À€" & YahooID & "À€98À€usÀ€135À€9.0.0.234À€"
'Packet = "6À€" & Crip(0) & "À€96À€" & Crip(1) & "À€0À€" & YahooID & "À€2À€" & YahooID & "À€192À€-1À€1À€" & YahooID & "À€135À€8.1.0.249À€148À€480À€"
'^^ THE FIRST PACKET I WAS USING TO BUILD THIS BEFORE UPDATING ITS PACKET TO THE LATEST
Log = YFHeader("54", Packet)
Debug.Print Log
End Function

'Packet = "1À€y4h00z.fyn3st_d34n0À€"
'Header("57", Packet)
Public Function Get_Key(YahooID As String) As String
Dim Packet As String
Packet = "1À€" & YahooID & "À€"
Get_Key = YFHeader("57", Packet)
Debug.Print Get_Key
End Function

Public Function Encrypt(YahooID As String, Passy As String, Seed As String, YF As Long)
On Error GoTo DazSmells
Dim Ace(1) As String
Ace(0) = String(80, vbNullChar)
Ace(1) = String(80, vbNullChar)
Encryption = YMSG12_ScriptedMind_Encrypt(YahooID, Passy, Seed, Ace(0), Ace(1), YF)
Dim Bee As String
Bee = InStr(1, Ace(0), vbNullChar)
Crip(0) = Left$(Ace(0), Bee - 1)
Bee = InStr(1, Ace(1), vbNullChar)
Crip(1) = Left$(Ace(1), Bee - 1)
DazSmells:
End Function

Public Function YFHeader(ByVal Head As String, ByVal data As String) As String
Dim Y, F
F = 0
Y = Len(data)
Do While Y > 299
Y = Y - 300
F = F + 1
Loop
YFHeader = "YMSG" & Chr(0) & Chr(Form1.Combo3.Text) & String(2, 0) & Chr(F) & Chr(Y) & Chr(0) & Chr("&H" & Head) & String(8, 0) & data
End Function

Public Function Pause(howlong)
Dim YF
YF = Timer
Do While Timer - YF < Val(howlong)
DoEvents
Loop
End Function



