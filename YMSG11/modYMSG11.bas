Attribute VB_Name = "modYMSG11"
' modYMSG11 by: Matthew Robertson.
' most packets i sniffed off yahoo messenger 5.6.
' i tryed to make this mod so u could add it to
' any project and make it login ymsg11.
' plz give cridit.

' YM11AUTH.DLL by: Deep

' mailto:uphome@nbnet.nb.ca

Option Explicit
Private Declare Function GetYahooStrings Lib "YM11AUTH.DLL" (ByVal UserName As String, ByVal password As String, ByVal seed As String, ByVal result_6 As String, ByVal result_96 As String, intt As Long) As Boolean

Type typYMSG
    Server      As String
    ID          As String
    PW          As String
    Profiles(6) As String
    Room        As String
    RoomID      As String
    Key         As String
End Type
Global YMSG As typYMSG

Function AddBuddie(ID As String, Who As String, Optional Grp As String = "Buddies", Optional Msg As String)
AddBuddie = Packet(83, "1¿Ä" & ID & "¿Ä7¿Ä" & Who & "¿Ä14¿Ä" & Msg & "¿Ä65¿Ä" & Grp & "¿Ä")
End Function

Function AwayMessage(Msg As String)
If LCase(Msg) = "invisible" Then
    AwayMessage = Packet(3, "10¿Ä12¿Ä")
Else
    AwayMessage = Packet(3, "10¿Ä99¿Ä19¿Ä" & Msg & "¿Ä47¿Ä0¿Ä")
End If
End Function


Function DelBuddie(ID As String, Who As String, Optional Grp As String = "Buddies")
DelBuddie = Packet(84, "1¿Ä" & ID & "¿Ä7¿Ä" & Who & "¿Ä65¿Ä" & Grp & "¿Ä")
End Function

Sub GetProfiles(MainID As String, Profiles As String, Optional Cbo As ComboBox)
'ymsg.profiles(num) will return that profile, but if there is no profiles it will return the main name
'not the best coding ever but it was the fastest way i could think to do it
On Error Resume Next
Dim Spt() As String, i As Integer
Spt = Split(Profiles & ",", ",")
i = UBound(Spt)
If i > 6 Then i = 6
With YMSG
 For i = 0 To i
    .Profiles(i) = Spt(i)
    If Not Cbo Is Nothing Then Cbo.AddItem Spt(i) ' adds to a combo box if present
 Next
 For i = UBound(Spt) To 6 ' if u have all profiles this will do nothing
    .Profiles(i) = MainID
 Next
End With
If Not Cbo Is Nothing Then Cbo.Text = MainID
End Sub
Public Function GetEncrStrings(ID As String, PW As String, SD As String, C1 As String, C2 As String, MD As Long) As Boolean
'YM11AUTH.DLL
On Error GoTo Error
Dim TS As String, TS2 As String, N As Long
 TS = String(80, vbNullChar)
 TS2 = String(80, vbNullChar)
 GetEncrStrings = GetYahooStrings(ID, PW, SD, TS, TS2, MD)
 N = InStr(1, TS, vbNullChar)
 C1 = Left(TS, N - 1)
 N = InStr(1, TS2, vbNullChar)
 C2 = Left(TS2, N - 1)
 GetEncrStrings = True
Exit Function
Error:
 GetEncrStrings = False
End Function


Function InviteConfrence(From As String, Who As String, Room As String, Msg As String)
InviteConfrence = Packet(18, "1¿Ä" & From & "¿Ä50¿Ä" & From & "¿Ä57¿Ä" & Room & "¿Ä58¿Ä" & Msg & "¿Ä52¿Ä" & Replace(Who, ",", "¿Ä52¿Ä") & "¿Ä13¿Ä0¿Ä")
End Function

Function JoinChat(ID As String, Room As String, Key As String)
JoinChat = Packet(98, "1¿Ä" & ID & "¿Ä62¿Ä¿Ä2¿Ä¿Ä104¿Ä" & Room & "¿Ä", Key)
End Function

Function Prejoin(ID As String)
Prejoin = Packet(96, "109¿Ä" & ID & "¿Ä1¿Ä" & ID & "¿Ä6¿Äabcde¿Ä", YMSG.Key)
End Function




Function SendChat(From As String, Room As String, Msg As String)
SendChat = Packet("A8", "1¿Ä" & From & "¿Ä104¿Ä" & Room & "¿Ä117¿Ä" & Msg & "¿Ä124¿Ä1¿Ä")
End Function

Function SendFile(From As String, Who As String, URL As String, Optional Size As String = "Undefined", Optional Msg As String = "")
'sends a url as if it where a file transfer (the size can b a string)
Dim FileName As String
FileName = Right(URL, Len(URL) - InStrRev(URL, "/"))
SendFile = Packet("4D", "5¿Ä" & Who & "¿Ä49¿ÄFILEXFER¿Ä1¿Ä" & From & "¿Ä14¿Ä" & Msg & "¿Ä13¿Ä1¿Ä27¿Ä" & FileName & "¿Ä28¿Ä" & Size & "¿Ä20¿Ä" & URL & "¿Ä")
End Function

Function SendIMV(From As String, WhoTo As String, IMV As String)
SendIMV = Packet("4D", "49¿ÄIMVIRONMENT¿Ä1¿Ä" & From & "¿Ä14¿Ä¿Ä13¿Ä0¿Ä5¿Ä" & WhoTo & "¿Ä63¿Ä" & IMV & "¿Ä64¿Ä0¿Ä")
End Function

Function UserKey(ID As String, Optional Key As String)
'prelogin
UserKey = Packet(57, "1¿Ä" & ID & "¿Ä", Key)
End Function
Function Packet(PackType As String, Pack As String, Optional ByVal Key As String)
'adds header to packet
' i seen a lot of other codes where this was coded usng a 'calc size' function
' wich looped till the packlen was under 256 and counted the times it had to loop
' wich was simple dividing, and then the remaindure, wich can b done simply w/ 'mod'
If Key = "" Then Key = String(4, 0)
Packet = "YMSG" & Chr(0) & Chr(11) & String(2, 0) & Chr(Fix(Len(Pack) / 256)) & _
Chr(Len(Pack) Mod 256) & Chr(0) & Chr("&H" & PackType) & _
String(4, 0) & Key & Pack
End Function
Public Function Conference(From As String, WhoTo As String, Message As String)
Conference = Packet("1A", "1¿Ä" + From + "¿Ä57¿Ä" + WhoTo + "¿Ä14¿Ä" + Message + "¿Ä97¿Ä1¿Ä")
End Function

Function SendPM(From As String, Who As String, Msg As String)
SendPM = Packet(6, "1¿Ä" & From & "¿Ä5¿Ä" & Who & "¿Ä14¿Ä" & Msg & "¿Ä97¿Ä1¿Ä")
End Function
Function Login(ByVal C1 As String, ByVal C2 As String, ID As String)
'login info
Login = Packet(54, "6¿Ä" & C1 & "¿Ä96¿Ä" & C2 & "¿Ä0¿Ä" & ID & "¿Ä2¿Ä1¿Ä1¿Ä" & ID & "¿Ä99¿Äbeta¿Ä135¿Ä6,0,0,0000¿Ä148¿Ä300¿Ä59¿ÄB" & Chr(&H9) & "284f5sh08s788&b=2¿Ä")
End Function

Function ViewShareFiles(From As String, Who As String)
ViewShareFiles = Packet("4D", "5¿Ä" & Who & "¿Ä49¿ÄFILEXFER¿Ä1¿Ä" & From & "¿Ä13¿Ä5¿Ä54¿ÄMSG1.0¿Ä")
End Function


