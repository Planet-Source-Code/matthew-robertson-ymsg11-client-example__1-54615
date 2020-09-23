VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "YMSG11 [Client Example By: Matthew Robertson]"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtIn 
      Height          =   1095
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   3480
      Width           =   4815
   End
   Begin VB.TextBox txtChatMsg 
      Height          =   285
      Left            =   2640
      TabIndex        =   16
      Text            =   "Chat Msg"
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton cmdChat 
      Caption         =   "Join"
      Height          =   288
      Left            =   4080
      TabIndex        =   15
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox txtRoom 
      Height          =   285
      Left            =   2640
      TabIndex        =   14
      Top             =   2160
      Width           =   2295
   End
   Begin VB.ComboBox cboProfiles 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1440
      Width           =   2415
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   288
      Left            =   1560
      TabIndex        =   9
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox txtMsg 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Text            =   "Private Msg"
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox txtTo 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox txtServ 
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Text            =   "alphacs2.msg.dcn.yahoo.com"
      Top             =   240
      Width           =   2295
   End
   Begin MSWinsockLib.Winsock wskYMSG 
      Left            =   3480
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Default         =   -1  'True
      Height          =   288
      Left            =   2520
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtPW 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Incomming PM/Chat::"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   18
      Top             =   3240
      Width           =   1545
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Join Room::"
      Height          =   195
      Index           =   4
      Left            =   2640
      TabIndex        =   13
      Top             =   1920
      Width           =   840
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Profile:"
      Height          =   195
      Index           =   3
      Left            =   1440
      TabIndex        =   12
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Send PM:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   705
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server:"
      Height          =   195
      Index           =   1
      Left            =   2520
      TabIndex        =   6
      Top             =   0
      Width           =   510
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User/Pass:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   795
   End
   Begin VB.Label lblStat 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "By: Matthew Robertson"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   3375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub AddText(Txt As String)
With txtIn
    If Len(.Text) > 5000 Then .Text = Left(.Text, 700) ' dont let it get to long
    .SelStart = Len(.Text)
    If Not .Text = "" Then .SelText = vbCrLf ' like break
    .SelText = Txt
    .SelStart = Len(.Text) ' scroll down
End With
End Sub


Sub LoadInfo()
txtID = GetSetting("YMSG11", "Login", "ID", "")
txtPW = GetSetting("YMSG11", "Login", "PW", "")
txtTo = GetSetting("YMSG11", "Other", "PM", "")
txtRoom = GetSetting("YMSG11", "Other", "RM", "")
End Sub

Sub SaveInfo()
If Not YMSG.ID = "" Then SaveSetting "YMSG11", "Login", "ID", YMSG.ID
If Not YMSG.PW = "" Then SaveSetting "YMSG11", "Login", "PW", YMSG.PW
If Not txtTo.Text = "" Then SaveSetting "YMSG11", "Other", "PM", txtTo.Text
If Not txtRoom.Text = "" Then SaveSetting "YMSG11", "Other", "RM", txtRoom.Text
End Sub

Function SendPack(Pack As String) As Boolean
On Error GoTo Error
    wskYMSG.SendData Pack
    Debug.Print Pack
    SendPack = True
Exit Function
Error:
    SendPack = False
End Function



Private Sub cmdChat_Click()
If cmdChat.Caption = "Join" Then
 YMSG.Room = txtRoom
 YMSG.RoomID = cboProfiles.Text ' this is so all chatsends and such use the right profile
 SendPack Prejoin(YMSG.RoomID)
ElseIf cmdChat.Caption = "Send" Then
 SendPack SendChat(YMSG.RoomID, YMSG.Room, txtChatMsg)
End If
End Sub


Private Sub cmdLogin_Click()
' i seen some ppl use 2 buttons for this and thats gay...
If cmdLogin.Caption = "Login" Then
 YMSG.ID = txtID
 YMSG.PW = txtPW
 If txtServ = "" Then txtServ = "alphacs2.msg.dcn.yahoo.com"
 YMSG.Server = txtServ
 wskYMSG.Close
 wskYMSG.Connect YMSG.Server, 5050
 lblStat = "Connecting..."
ElseIf cmdLogin.Caption = "Logout" Then
 wskYMSG.Close
 lblStat = "Logged Out!"
 cmdLogin.Caption = "Login"
End If
End Sub



Private Sub cmdSend_Click()
SendPack SendPM(cboProfiles.Text, txtTo, txtMsg)
End Sub





Private Sub Form_Load()
LoadInfo
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveInfo
End Sub


Private Sub txtChatMsg_Change()
cmdChat.Default = True
End Sub

Private Sub txtMsg_Change()
cmdSend.Default = True
End Sub


Private Sub txtPW_Change()
cmdSend.Default = True
End Sub


Private Sub txtRoom_Change()
cmdChat.Default = True
End Sub

Private Sub wskYMSG_Connect()
SendPack UserKey(YMSG.ID)
End Sub

Private Sub wskYMSG_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Data As String, SptData() As String
Dim C1   As String, C2        As String
wskYMSG.GetData Data
SptData = Split(Data, "À€")
Select Case Asc(Mid(Data, 12, 1))
 Case 87 'send pass
  If GetEncrStrings(YMSG.ID, YMSG.PW, SptData(3), C1, C2, 1) = True Then
      YMSG.Key = Mid(Data, 17, 4)
    SendPack Login(C1, C2, YMSG.ID)
  Else
    wskYMSG.Close
    lblStat = "Error!"
    MsgBox "There was an error on 'YM11AUTH.DLL'" & vbCrLf & "Be sure you extraced the hole zip!", vbCritical, "YM11AUTH.DLL"
  End If
 Case 85 ' password correct, logged in
    GetProfiles YMSG.ID, SptData(5), cboProfiles
    lblStat = "Logged in!"
    cmdLogin.Caption = "Logout"
 Case 84 ' password thing
  If Mid(Data, 13, 4) = "ÿÿÿÿ" Then ' error
    lblStat = "Wrong password!"
    txtPW = ""
    wskYMSG.Close
  End If
 Case 150 ' room
    SendPack JoinChat(YMSG.RoomID, YMSG.Room, YMSG.Key) ' join room
    cmdChat.Caption = "Send"
 Case 168  ' incomming chat
    AddText SptData(1) & " - " & SptData(3) & ": " & SptData(5) ' room - sender: message
 Case 6 ' incomming pm
    AddText SptData(3) & ": " & SptData(7) ' sender: message
End Select

'i though it would b easyer to learn the packets viewed like this:
Dim PacketView As String
PacketView = Asc(Mid(Data, 12, 1)) & "- " & "(" & Mid(SptData(0), 17) & ")"
For i = 1 To UBound(SptData)
    PacketView = PacketView & Chr(32) & i & "(" & SptData(i) & ")"
Next
Debug.Print PacketView & vbCrLf
End Sub


