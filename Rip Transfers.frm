VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Transfers 
   BorderStyle     =   0  'None
   ClientHeight    =   3300
   ClientLeft      =   1170
   ClientTop       =   285
   ClientWidth     =   8880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   Begin VB.Timer OHstat 
      Interval        =   10
      Left            =   5880
      Top             =   1800
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1200
      TabIndex        =   21
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox port0 
      Height          =   285
      Index           =   0
      Left            =   6360
      TabIndex        =   20
      Top             =   840
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Index           =   0
      Interval        =   2000
      Left            =   360
      Top             =   0
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   6480
      TabIndex        =   6
      Text            =   "127.0.0.1"
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox HandleWhat 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   5
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   10
      Left            =   2400
      TabIndex        =   4
      Top             =   9000
      Width           =   1335
   End
   Begin VB.TextBox port0 
      Height          =   285
      Index           =   10
      Left            =   3720
      TabIndex        =   3
      Top             =   8640
      Width           =   495
   End
   Begin VB.TextBox HandleWhat 
      Height          =   285
      Index           =   10
      Left            =   4200
      TabIndex        =   2
      Top             =   8640
      Width           =   375
   End
   Begin VB.Timer Timer2 
      Index           =   0
      Interval        =   2000
      Left            =   4800
      Top             =   120
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   0
      Left            =   5280
      TabIndex        =   1
      Text            =   "c:\"
      Top             =   240
      Width           =   1335
   End
   Begin VB.Timer OStat 
      Interval        =   10
      Left            =   4800
      Top             =   1680
   End
   Begin VB.TextBox HandleWhater 
      Height          =   285
      Index           =   0
      Left            =   6840
      TabIndex        =   0
      Top             =   840
      Width           =   375
   End
   Begin MSWinsockLib.Winsock Get0 
      Left            =   5280
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Send0 
      Left            =   5280
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   7473
   End
   Begin VB.Label Label7 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   19
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Not Connected"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   18
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   17
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   16
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Size:"
      Height          =   255
      Index           =   10
      Left            =   2400
      TabIndex        =   15
      Top             =   9360
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Not Connected"
      Height          =   255
      Index           =   10
      Left            =   2400
      TabIndex        =   14
      Top             =   8760
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Done:"
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   13
      Top             =   9600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Speed:"
      Height          =   255
      Index           =   10
      Left            =   3360
      TabIndex        =   12
      Top             =   9360
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   11
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "Not Connected"
      Height          =   255
      Index           =   0
      Left            =   5280
      TabIndex        =   10
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   9
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   8
      Top             =   840
      Width           =   1095
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   4680
      X2              =   4680
      Y1              =   120
      Y2              =   8280
   End
   Begin VB.Label Label9 
      Caption         =   "Done:                       Speed:             Size:"
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   1080
      Visible         =   0   'False
      Width           =   3735
   End
End
Attribute VB_Name = "Transfers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim givemore(7) As String
Dim currentint(7) As Long
Dim currentin(7) As Long
Dim ratee(7) As Integer
Dim ratey(7) As Integer
Dim filesizeu(7) As Long
Dim FileNumber1(7)
Dim FileNumber2(7)
Dim GotBytes(7)
Sub Sleep(ByVal MillaSec As Long, Optional ByVal DeepSleep As Boolean = False)
    Dim tStart#, Tmr#
    tStart = Timer

    While Tmr < (MillaSec / 1000)
        Tmr = Timer - tStart
        If DeepSleep = False Then DoEvents
    Wend
End Sub
Sub SetFileNumbers()
ajax = 1
ajox = 10
For i = 0 To 7
FileNumber1(i) = ajax
FileNumber2(i) = ajox
ajax = ajax + 1
ajox = ajox + 1
Next
End Sub

Private Sub Command1_Click()
Timer1(0).Enabled = True
Timer1(1).Enabled = True
Text3(0).Text = 0
Text3(1).Text = 1
End Sub

Private Sub Form_Load()
SetFileNumbers
End Sub



'==================Start Connect Code================
Sub Connectex0()
Get0.Close
Get0.Connect Text1(0).Text, port0(0).Text
Label6(0).Caption = "Connecting"

End Sub
Sub Connectex1()
Get1.Connect Text1(1).Text, port0(1).Text
Label6(1).Caption = "Connecting"
End Sub
Sub Connectex2()
Send2.Connect Text1(2).Text, port0(2).Text
Label5(2).Caption = "Connecting"
Label7(2).Caption = FileLen(Text3(2).Text)
Text4.Text = 1
End Sub
Sub Connectex3()
Send3.Connect Text1(3).Text, port0(3).Text
Label5(3).Caption = "Connecting"
Label7(3).Caption = FileLen(Text3(3).Text)
Text4.Text = 1
End Sub
Sub Connectex4()
Send4.Connect Text1(4).Text, port0(4).Text
Label5(4).Caption = "Connecting"
Label7(4).Caption = FileLen(Text3(4).Text)
Text4.Text = 1
End Sub
Sub Connectex5()
Send5.Connect Text1(5).Text, port0(5).Text
Label5(5).Caption = "Connecting"
Label7(5).Caption = FileLen(Text3(5).Text)
Text4.Text = 1
End Sub
Sub Connectex6()
Send6.Connect Text1(6).Text, port0(6).Text
Label5(6).Caption = "Connecting"
Label7(6).Caption = FileLen(Text3(6).Text)
Text4.Text = 1
End Sub
Sub Connectex7()
Send7.Connect Text1(7).Text, port0(7).Text
Label5(7).Caption = "Connecting"
Label7(7).Caption = FileLen(Text3(7).Text)
Text4.Text = 1
End Sub
'==================End Connect Code================



'==================Start Listen Code===============
Sub Listenex0()
Send0.Close
Send0.Listen
Label5(0).Caption = "Listening"


End Sub
Sub Listenex1()
Send1.Close
Send1.Listen
Label5(1).Caption = "Listening"
End Sub
Sub Listenex2()
Get2.Close
Get2.Listen
Label6(2).Caption = "Listening"
End Sub
Sub Listenex3()
Get3.Close
Get3.Listen
Label6(3).Caption = "Listening"
End Sub
Sub Listenex4()
Get4.Close
Get4.Listen
Label6(4).Caption = "Listening"
End Sub
Sub Listenex5()
Get5.Close
Get5.Listen
Label6(5).Caption = "Listening"
End Sub
Sub Listenex6()
Get6.Close
Get6.Listen
Label6(6).Caption = "Listening"
End Sub
Sub Listenex7()
Get7.Close
Get7.Listen
Label6(7).Caption = "Listening"
End Sub
'==================End Listen Code===============




Private Sub Get0_Connect()
Text4.Text = 0
End Sub

Private Sub Get1_Connect()
Text4.Text = 0
End Sub

Private Sub Get2_Connect()
Text4.Text = 0
End Sub

Private Sub Get3_Connect()
Text4.Text = 0
End Sub

Private Sub Get4_Connect()
Text4.Text = 0
End Sub

Private Sub Get5_Connect()
Text4.Text = 0
End Sub

Private Sub Get6_Connect()
Text4.Text = 0
End Sub

Private Sub Get7_Connect()
Text4.Text = 0
End Sub



Private Sub OHstat_Timer()
If HandleWhater(0).Text = "" Then Exit Sub

Form1.kbs(HandleWhater(0)).Text = Label3(0).Caption
Form1.done(HandleWhater(0)).Text = Label4(0).Caption

If Form1.done(HandleWhater(0)).Text > 98 Then
Form1.done(HandleWhater(0)).Text = 100
Form1.status(HandleWhater(0)).Text = "Done!"
End If
End Sub

Private Sub OStat_Timer()
If HandleWhat(0).Text = "" Then Exit Sub

Form1.kbs(HandleWhat(0)).Text = Label1(0).Caption
Form1.done(HandleWhat(0)).Text = Label2(0).Caption

If Form1.done(HandleWhat(0)).Text > 98 Then
Form1.done(HandleWhat(0)).Text = 100
Form1.status(HandleWhat(0)).Text = "Done!"
End If
End Sub

'===============Start SendSocket_Close Code======
Private Sub Send0_Close()
If Label2(0).Caption > 99 Then
Label5(0).Caption = "Done!"
Else
Label5(0).Caption = "Canceled"
End If
HandleWhat(0).Text = ""

End Sub

Private Sub Send0_ConnectionRequest(ByVal requestID As Long)
Send0.Close
Send0.Accept requestID
Sleep 1000
Label7(0).Caption = FileLen(Text3(0).Text)
Label5(0).Caption = "Connected"
Send0.SendData "FILSZ " & Label7(0).Caption
Sleep 2000
Send0.SendData "FILNM " & Form1.fname(HandleWhat(0)).Text
End Sub

Private Sub Send1_Close()
If Label2(1).Caption > 99 Then
Label5(1).Caption = "Done!"
Else
Label5(1).Caption = "Canceled"
End If
End Sub

Private Sub Send1_ConnectionRequest(ByVal requestID As Long)
Send1.Close
Send1.Accept requestID
Sleep 1000
Label7(1).Caption = FileLen(Text3(1).Text)
Label5(1).Caption = "Connected"
Send1.SendData "FILSZ " & Label7(1).Caption
Sleep 2000
Send1.SendData "FILNM " & Form1.fname(HandleWhat(1)).Text
End Sub

Private Sub Send2_Close()
If Label2(2).Caption > 99 Then
Label5(2).Caption = "Done!"
Else
Label5(2).Caption = "Canceled"
End If
End Sub

Private Sub Send2_ConnectionRequest(ByVal requestID As Long)
Send2.Close
Send2.Accept requestID
Sleep 1000
Label7(2).Caption = FileLen(Text3(2).Text)
Label5(2).Caption = "Connected"
Send2.SendData "FILSZ " & Label7(2).Caption
Sleep 2000
Send2.SendData "FILNM " & Form1.fname(HandleWhat(2)).Text
End Sub

Private Sub Send3_Close()
If Label2(3).Caption > 99 Then
Label5(3).Caption = "Done!"
Else
Label5(3).Caption = "Canceled"
End If
End Sub
Private Sub Send4_Close()
If Label2(4).Caption > 99 Then
Label5(4).Caption = "Done!"
Else
Label5(4).Caption = "Canceled"
End If
End Sub
Private Sub Send5_Close()
If Label2(5).Caption > 99 Then
Label5(5).Caption = "Done!"
Else
Label5(5).Caption = "Canceled"
End If
End Sub
Private Sub Send6_Close()
If Label2(6).Caption > 99 Then
Label5(6).Caption = "Done!"
Else
Label5(6).Caption = "Canceled"
End If
End Sub
Private Sub Send7_Close()
If Label2(7).Caption > 99 Then
Label5(7).Caption = "Done!"
Else
Label5(7).Caption = "Canceled"
End If
End Sub
'===============End SendSocket_Close Code========



'============Start Speed Calc Code===============
Private Sub Timer1_Timer(Index As Integer)
Label1(Index).Caption = (ratee(Index) / 2)
ratee(Index) = 0
End Sub

Private Sub Timer2_Timer(Index As Integer)
Label3(Index).Caption = (ratey(Index) / 2)
ratey(Index) = 0
End Sub
'============End Speed Calc Code=================



'========Start SendSocket_Connect Code===========
Private Sub Send0_Connect()

End Sub
Private Sub Send2_Connect()
Label5(2).Caption = "Connected"
Send2.SendData "FILSZ " & Label7(2).Caption
Sleep 2000
Send2.SendData "FILNM " & Form1.fname(HandleWhat(2)).Text
End Sub
Private Sub Send3_Connect()
Label5(3).Caption = "Connected"
Send3.SendData "FILSZ " & Label7(3).Caption
Sleep 2000
Send3.SendData "FILNM " & Form1.fname(HandleWhat(3)).Text
End Sub
Private Sub Send4_Connect()
Label5(4).Caption = "Connected"
Send4.SendData "FILSZ " & Label7(4).Caption
Sleep 2000
Send4.SendData "FILNM " & Form1.fname(HandleWhat(4)).Text
End Sub
Private Sub Send5_Connect()
Label5(5).Caption = "Connected"
Send5.SendData "FILSZ " & Label7(5).Caption
Sleep 2000
Send5.SendData "FILNM " & Form1.fname(HandleWhat(5)).Text
End Sub
Private Sub Send6_Connect()
Label5(6).Caption = "Connected"
Send6.SendData "FILSZ " & Label7(6).Caption
Sleep 2000
Send6.SendData "FILNM " & Form1.fname(HandleWhat(6)).Text
End Sub
Private Sub Send7_Connect()
Label5(7).Caption = "Connected"
Send7.SendData "FILSZ " & Label7(7).Caption
Sleep 2000
Send7.SendData "FILNM " & Form1.fname(HandleWhat(7)).Text
End Sub
'=========End SendSocket_Connect Code===========



'==========Start SendSocket_DataArrival Code====
Private Sub Send0_DataArrival(ByVal bytesTotal As Long)
Dim dart As String
Send0.GetData dart

If InStr(1, dart, "GTB", vbTextCompare) <> 0 Then
Label5(0).Caption = "Sending File"
givemore(0) = 1
SendTheFile0
End If

If InStr(1, dart, "RSM", vbTextCompare) <> 0 Then
a = Mid(dart, 4, Len(dart) - 3)
currentint(0) = a
Label5(0).Caption = "Sending File"
givemore(0) = 1
SendRSMFile0
End If

If dart = "GIVEMORE" Then givemore(0) = 1
End Sub
Private Sub Send1_DataArrival(ByVal bytesTotal As Long)
Dim dart As String
Send1.GetData dart

If InStr(1, dart, "GTB", vbTextCompare) <> 0 Then
Label5(1).Caption = "Sending File"
givemore(1) = 1
SendTheFile1
End If

If InStr(1, dart, "RSM", vbTextCompare) <> 0 Then
a = Mid(dart, 4, Len(dart) - 3)
currentint(1) = a
Label5(1).Caption = "Sending File"
givemore(1) = 1
SendRSMFile1
End If

If dart = "GIVEMORE" Then givemore(1) = 1
End Sub
Private Sub Send2_DataArrival(ByVal bytesTotal As Long)
Dim dart As String
Send2.GetData dart

If InStr(1, dart, "GTB", vbTextCompare) <> 0 Then
Label5(2).Caption = "Sending File"
givemore(2) = 1
SendTheFile2
End If

If InStr(1, dart, "RSM", vbTextCompare) <> 0 Then
a = Mid(dart, 4, Len(dart) - 3)
currentint(2) = a
Label5(2).Caption = "Sending File"
givemore(2) = 1
SendRSMFile2
End If

If dart = "GIVEMORE" Then givemore(2) = 1
End Sub
Private Sub Send3_DataArrival(ByVal bytesTotal As Long)
Dim dart As String
Send3.GetData dart

If InStr(1, dart, "GTB", vbTextCompare) <> 0 Then
Label5(3).Caption = "Sending File"
givemore(3) = 1
SendTheFile3
End If

If InStr(1, dart, "RSM", vbTextCompare) <> 0 Then
a = Mid(dart, 4, Len(dart) - 3)
currentint(3) = a
Label5(3).Caption = "Sending File"
givemore(3) = 1
SendRSMFile3
End If

If dart = "GIVEMORE" Then givemore(3) = 1
End Sub
Private Sub Send4_DataArrival(ByVal bytesTotal As Long)
Dim dart As String
Send4.GetData dart

If InStr(1, dart, "GTB", vbTextCompare) <> 0 Then
Label5(4).Caption = "Sending File"
givemore(4) = 1
SendTheFile4
End If

If InStr(1, dart, "RSM", vbTextCompare) <> 0 Then
a = Mid(dart, 4, Len(dart) - 3)
currentint(4) = a
Label5(4).Caption = "Sending File"
givemore(4) = 1
SendRSMFile4
End If

If dart = "GIVEMORE" Then givemore(4) = 1
End Sub
Private Sub Send5_DataArrival(ByVal bytesTotal As Long)
Dim dart As String
Send5.GetData dart

If InStr(1, dart, "GTB", vbTextCompare) <> 0 Then
Label5(5).Caption = "Sending File"
givemore(5) = 1
SendTheFile5
End If

If InStr(1, dart, "RSM", vbTextCompare) <> 0 Then
a = Mid(dart, 4, Len(dart) - 3)
currentint(5) = a
Label5(5).Caption = "Sending File"
givemore(5) = 1
SendRSMFile5
End If

If dart = "GIVEMORE" Then givemore(5) = 1
End Sub
Private Sub Send6_DataArrival(ByVal bytesTotal As Long)
Dim dart As String
Send6.GetData dart

If InStr(1, dart, "GTB", vbTextCompare) <> 0 Then
Label5(6).Caption = "Sending File"
givemore(6) = 1
SendTheFile6
End If

If InStr(1, dart, "RSM", vbTextCompare) <> 0 Then
a = Mid(dart, 4, Len(dart) - 3)
currentint(6) = a
Label5(6).Caption = "Sending File"
givemore(6) = 1
SendRSMFile6
End If

If dart = "GIVEMORE" Then givemore(6) = 1
End Sub
Private Sub Send7_DataArrival(ByVal bytesTotal As Long)
Dim dart As String
Send7.GetData dart

If InStr(1, dart, "GTB", vbTextCompare) <> 0 Then
Label5(7).Caption = "Sending File"
givemore(7) = 1
SendTheFile7
End If

If InStr(1, dart, "RSM", vbTextCompare) <> 0 Then
a = Mid(dart, 4, Len(dart) - 3)
currentint(7) = a
Label5(7).Caption = "Sending File"
givemore(7) = 1
SendRSMFile7
End If

If dart = "GIVEMORE" Then givemore(7) = 1
End Sub
'==========End SendSocket_DataArrival Code=======




'==================Start Resume Code=============
Sub SendRSMFile0()

Dim tempbuffer As String

Open Text3(0).Text For Binary Access Read As #FileNumber1(0)
Label2(0).Caption = 0

Do Until EOF(FileNumber1(0))
Do Until givemore(0) = 1
DoEvents
Loop
tempbuffer = Space$(1024)

Get #FileNumber1(0), currentint(0), tempbuffer


currentint(0) = currentint(0) + 1024
filesizeu(0) = Label7(0).Caption
Label2(0).Caption = Int((currentint(0) / filesizeu(0)) * 100)
ratee(0) = ratee(0) + 1

Send0.SendData "NMDATA" & tempbuffer
givemore(0) = 0
Loop
Sleep 4000
Close #FileNumber1(0)
Send0.Close
Label5(0).Caption = "Done!"
If HandleWhat(0) <> "" Then Form1.status(HandleWhat(0)).Text = "Done!"
Exit Sub
End Sub
Sub SendRSMFile1()
Dim tempbuffer As String

Open Text3(1).Text For Binary Access Read As #FileNumber1(1)
Label2(1).Caption = 0

Do Until EOF(FileNumber1(1))
Do Until givemore(1) = 1
DoEvents
Loop
tempbuffer = Space$(1024)

Get #FileNumber1(1), currentint(1), tempbuffer


currentint(1) = currentint(1) + 1024
filesizeu(1) = Label7(1).Caption
Label2(1).Caption = Int((currentint(1) / filesizeu(1)) * 100)
ratee(1) = ratee(1) + 1

Send1.SendData "NMDATA" & tempbuffer
givemore(1) = 0
Loop
Sleep 4000
Close #FileNumber1(1)
Send1.Close
Label5(1).Caption = "Done!"
Exit Sub
End Sub
Sub SendRSMFile2()
Dim tempbuffer As String

Open Text3(2).Text For Binary Access Read As #FileNumber1(2)
Label2(2).Caption = 0

Do Until EOF(FileNumber1(2))
Do Until givemore(2) = 1
DoEvents
Loop
tempbuffer = Space$(1024)

Get #FileNumber1(2), currentint(2), tempbuffer


currentint(2) = currentint(2) + 1024
filesizeu(2) = Label7(2).Caption
Label2(2).Caption = Int((currentint(2) / filesizeu(2)) * 100)
ratee(2) = ratee(2) + 1

Send2.SendData "NMDATA" & tempbuffer
givemore(1) = 0
Loop
Sleep 4000
Close #FileNumber1(1)
Send1.Close
Label5(1).Caption = "Done!"
Exit Sub
End Sub
Sub SendRSMFile3()
Dim tempbuffer As String

Open Text3(3).Text For Binary Access Read As #FileNumber1(3)
Label2(3).Caption = 0

Do Until EOF(FileNumber1(3))
Do Until givemore(3) = 1
DoEvents
Loop
tempbuffer = Space$(1024)

Get #FileNumber1(3), currentint(3), tempbuffer


currentint(3) = currentint(3) + 1024
filesizeu(3) = Label7(3).Caption
Label2(3).Caption = Int((currentint(3) / filesizeu(3)) * 100)
ratee(3) = ratee(3) + 1

Send3.SendData "NMDATA" & tempbuffer
givemore(3) = 0
Loop
Sleep 4000
Close #FileNumber1(3)
Send3.Close
Label5(3).Caption = "Done!"
Exit Sub
End Sub
Sub SendRSMFile4()
Dim tempbuffer As String

Open Text3(4).Text For Binary Access Read As #FileNumber1(4)
Label2(4).Caption = 0

Do Until EOF(FileNumber1(4))
Do Until givemore(4) = 1
DoEvents
Loop
tempbuffer = Space$(1024)

Get #FileNumber1(4), currentint(4), tempbuffer


currentint(4) = currentint(4) + 1024
filesizeu(4) = Label7(4).Caption
Label2(4).Caption = Int((currentint(4) / filesizeu(4)) * 100)
ratee(4) = ratee(4) + 1

Send4.SendData "NMDATA" & tempbuffer
givemore(4) = 0
Loop
Sleep 4000
Close #FileNumber1(4)
Send4.Close
Label5(4).Caption = "Done!"
Exit Sub
End Sub

Sub SendRSMFile5()
Dim tempbuffer As String

Open Text3(5).Text For Binary Access Read As #FileNumber1(5)
Label2(5).Caption = 0

Do Until EOF(FileNumber1(5))
Do Until givemore(5) = 1
DoEvents
Loop
tempbuffer = Space$(1024)

Get #FileNumber1(5), currentint(5), tempbuffer


currentint(5) = currentint(p) + 1024
filesizeu(5) = Label7(5).Caption
Label2(5).Caption = Int((currentint(5) / filesizeu(5)) * 100)
ratee(5) = ratee(5) + 1

Send5.SendData "NMDATA" & tempbuffer
givemore(5) = 0
Loop
Sleep 5000
Close #FileNumber1(5)
Send5.Close
Label5(5).Caption = "Done!"
Exit Sub
End Sub
Sub SendRSMFile6()
Dim tempbuffer As String

Open Text3(6).Text For Binary Access Read As #FileNumber1(6)
Label2(6).Caption = 0

Do Until EOF(FileNumber1(6))
Do Until givemore(6) = 1
DoEvents
Loop
tempbuffer = Space$(1024)

Get #FileNumber1(6), currentint(6), tempbuffer


currentint(6) = currentint(6) + 1024
filesizeu(6) = Label7(6).Caption
Label2(6).Caption = Int((currentint(6) / filesizeu(6)) * 100)
ratee(6) = ratee(6) + 1

Send6.SendData "NMDATA" & tempbuffer
givemore(6) = 0
Loop
Sleep 6000
Close #FileNumber1(6)
Send6.Close
Label5(6).Caption = "Done!"

Exit Sub
End Sub
Sub SendRSMFile7()
Dim tempbuffer As String

Open Text3(7).Text For Binary Access Read As #FileNumber1(7)
Label2(7).Caption = 0

Do Until EOF(FileNumber1(7))
Do Until givemore(7) = 1
DoEvents
Loop
tempbuffer = Space$(1024)

Get #FileNumber1(7), currentint(7), tempbuffer


currentint(7) = currentint(7) + 1024
filesizeu(7) = Label7(7).Caption
Label2(7).Caption = Int((currentint(7) / filesizeu(7)) * 100)
ratee(7) = ratee(7) + 1

Send7.SendData "NMDATA" & tempbuffer
givemore(7) = 0
Loop
Sleep 7000
Close #FileNumber1(7)
Send7.Close
Label5(7).Caption = "Done!"
Exit Sub
End Sub
'===============End ResumeFile Code==========




'=================Start SendFile Code==========
Sub SendTheFile0()
Dim tempbuffer As String
'On Error GoTo errhand
FileNumber3 = FreeFile
Open Text3(0).Text For Binary Access Read As #FileNumber1(0)

Label2(0).Caption = 0
currentint(0) = 0

Do Until EOF(FileNumber1(0))
Do Until givemore(0) = 1
DoEvents
Loop
tempbuffer = Space$(1024)

Get #FileNumber1(0), , tempbuffer
currentint(0) = currentint(0) + 1024
filesizeu(0) = Label7(0).Caption
Label2(0).Caption = Int((currentint(0) / filesizeu(0)) * 100)
ratee(0) = ratee(0) + 1

Send0.SendData "NMDATA" & tempbuffer
givemore(0) = 0
Loop
Sleep 3000
Close #FileNumber1(0)
Send0.Close
Label5(0).Caption = "Done!"
If HandleWhat(0) <> "" Then Form1.status(HandleWhat(0)).Text = "Done!"
errhand:
Send0.Close
Exit Sub
End Sub
Sub SendTheFile1()
Dim tempbuffer As String
'On Error GoTo errhand
FileNumber3 = FreeFile
Open Text3(1).Text For Binary Access Read As #FileNumber1(1)

Label2(1).Caption = 0
currentint(1) = 0

Do Until EOF(FileNumber1(1))
Do Until givemore(1) = 1
DoEvents
Loop
tempbuffer = Space$(1024)

Get #FileNumber1(1), , tempbuffer
currentint(1) = currentint(1) + 1024
filesizeu(1) = Label7(1).Caption
Label2(1).Caption = Int((currentint(1) / filesizeu(1)) * 100)
ratee(1) = ratee(1) + 1

Send1.SendData "NMDATA" & tempbuffer
givemore(1) = 0
Loop
Sleep 4000
Close #FileNumber1(1)
Send1.Close
Label5(1).Caption = "Done!"
errhand:
Send1.Close
Exit Sub
End Sub
Sub SendTheFile2()
Dim tempbuffer As String
'On Error GoTo errhand
FileNumber3 = FreeFile
Open Text3(2).Text For Binary Access Read As #FileNumber1(2)

Label2(2).Caption = 0
currentint(2) = 0

Do Until EOF(FileNumber1(2))
Do Until givemore(2) = 1
DoEvents
Loop
tempbuffer = Space$(1024)

Get #FileNumber1(2), , tempbuffer
currentint(2) = currentint(2) + 1024
filesizeu(2) = Label7(2).Caption
Label2(2).Caption = Int((currentint(2) / filesizeu(2)) * 100)
ratee(2) = ratee(2) + 1

Send2.SendData "NMDATA" & tempbuffer
givemore(2) = 0
Loop
Sleep 4000
Close #FileNumber1(2)
Send2.Close
Label5(2).Caption = "Done!"
errhand:
Send2.Close
Exit Sub
End Sub
Sub SendTheFile3()
Dim tempbuffer As String
'On Error GoTo errhand
FileNumber3 = FreeFile
Open Text3(3).Text For Binary Access Read As #FileNumber1(3)

Label2(3).Caption = 0
currentint(3) = 0

Do Until EOF(FileNumber1(3))
Do Until givemore(3) = 1
DoEvents
Loop
tempbuffer = Space$(1024)

Get #FileNumber1(3), , tempbuffer
currentint(3) = currentint(3) + 1024
filesizeu(3) = Label7(3).Caption
Label2(3).Caption = Int((currentint(3) / filesizeu(3)) * 100)
ratee(3) = ratee(3) + 1

Send3.SendData "NMDATA" & tempbuffer
givemore(3) = 0
Loop
Sleep 3000
Close #FileNumber1(3)
Send3.Close
Label5(3).Caption = "Done!"
errhand:
Send3.Close
Exit Sub
End Sub
Sub SendTheFile4()
Dim tempbuffer As String
'On Error GoTo errhand
FileNumber3 = FreeFile
Open Text3(4).Text For Binary Access Read As #FileNumber1(4)

Label2(4).Caption = 0
currentint(4) = 0

Do Until EOF(FileNumber1(4))
Do Until givemore(4) = 1
DoEvents
Loop
tempbuffer = Space$(1024)

Get #FileNumber1(4), , tempbuffer
currentint(4) = currentint(4) + 1024
filesizeu(4) = Label7(4).Caption
Label2(4).Caption = Int((currentint(4) / filesizeu(4)) * 100)
ratee(4) = ratee(4) + 1

Send4.SendData "NMDATA" & tempbuffer
givemore(4) = 0
Loop
Sleep 4000
Close #FileNumber1(4)
Send4.Close
Label5(4).Caption = "Done!"
errhand:
Send4.Close
Exit Sub
End Sub
Sub SendTheFile5()
Dim tempbuffer As String
'On Error GoTo errhand
FileNumber3 = FreeFile
Open Text3(5).Text For Binary Access Read As #FileNumber1(5)

Label2(5).Caption = 0
currentint(5) = 0

Do Until EOF(FileNumber1(5))
Do Until givemore(5) = 1
DoEvents
Loop
tempbuffer = Space$(1024)

Get #FileNumber1(5), , tempbuffer
currentint(5) = currentint(5) + 1024
filesizeu(5) = Label7(5).Caption
Label2(5).Caption = Int((currentint(5) / filesizeu(5)) * 100)
ratee(5) = ratee(5) + 1

Send5.SendData "NMDATA" & tempbuffer
givemore(5) = 0
Loop
Sleep 4000
Close #FileNumber1(5)
Send5.Close
Label5(5).Caption = "Done!"
errhand:
Send5.Close
Exit Sub
End Sub
Sub SendTheFile6()
Dim tempbuffer As String
'On Error GoTo errhand
FileNumber3 = FreeFile
Open Text3(6).Text For Binary Access Read As #FileNumber1(6)

Label2(6).Caption = 0
currentint(6) = 0

Do Until EOF(FileNumber1(6))
Do Until givemore(6) = 1
DoEvents
Loop
tempbuffer = Space$(1024)

Get #FileNumber1(6), , tempbuffer
currentint(6) = currentint(6) + 1024
filesizeu(6) = Label7(6).Caption
Label2(6).Caption = Int((currentint(6) / filesizeu(6)) * 100)
ratee(6) = ratee(6) + 1

Send6.SendData "NMDATA" & tempbuffer
givemore(6) = 0
Loop
Sleep 4000
Close #FileNumber1(6)
Send6.Close
Label5(6).Caption = "Done!"
errhand:
Send6.Close
Exit Sub
End Sub
Sub SendTheFile7()
Dim tempbuffer As String
'On Error GoTo errhand
FileNumber3 = FreeFile
Open Text3(7).Text For Binary Access Read As #FileNumber1(7)

Label2(7).Caption = 0
currentint(7) = 0

Do Until EOF(FileNumber1(7))
Do Until givemore(7) = 1
DoEvents
Loop
tempbuffer = Space$(1024)

Get #FileNumber1(7), , tempbuffer
currentint(7) = currentint(7) + 1024
filesizeu(7) = Label7(7).Caption
Label2(7).Caption = Int((currentint(7) / filesizeu(7)) * 100)
ratee(7) = ratee(7) + 1

Send7.SendData "NMDATA" & tempbuffer
givemore(7) = 0
Loop
Sleep 4000
Close #FileNumber1(7)
Send7.Close
Label5(7).Caption = "Done!"
errhand:
Send7.Close
Exit Sub
End Sub
'=============End SendtheFile Code========




'=============Start Send_Error Code=======
Private Sub Send0_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Label5(0).Caption = "Canceled"
Send0.Close
End Sub

Private Sub Send1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Label5(1).Caption = "Canceled"
Send1.Close
End Sub

Private Sub Send2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Label5(2).Caption = "Canceled"
Send2.Close
End Sub
Private Sub Send3_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Label5(3).Caption = "Canceled"
Send3.Close
End Sub
Private Sub Send4_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Label5(4).Caption = "Canceled"
Send4.Close
End Sub
Private Sub Send5_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Label5(5).Caption = "Canceled"
Send5.Close
End Sub
Private Sub Send6_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Label5(6).Caption = "Canceled"
Send6.Close
End Sub
Private Sub Send7_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Label5(7).Caption = "Canceled"
Send7.Close
End Sub
'=================End Send_Error Code=====



'==============Start Get_DataArrival Code====
Private Sub Get0_DataArrival(ByVal bytesTotal As Long)
Dim dart As String
Dim jki As Long
Dim jku As Long
Get0.GetData dart
If InStr(1, dart, "FILSZ", vbTextCompare) <> 0 Then
a = Mid(dart, 7, Len(dart) - 6)
Label8(0).Caption = a
End If

If InStr(1, dart, "FILNM", vbTextCompare) <> 0 Then
a = Mid(dart, 7, Len(dart) - 6)
Text2(0).Text = Text2(0).Text & a
currentin(0) = 0
On Error GoTo GetFileFstTime
b = FileLen(Text2(0).Text) 'resume
currentin(0) = b
If b < 1 Then b = 1
Get0.SendData "RSM" & b
Open Text2(0).Text For Binary Access Write As #FileNumber2(0)
GotBytes(0) = currentin(0)
Exit Sub

GetFileFstTime: 'the first time getting the file
Get0.SendData "GTB"
Label4(0).Caption = 0
Open Text2(0).Text For Binary Access Write As #FileNumber2(0)
Label6(0).Caption = "Getting File"
GotBytes(0) = "0"
Exit Sub
End If
'---------normal download
If InStr(1, dart, "NMDATA") <> 0 Then
sendsze = 1024
dart = Mid(dart, 7, sendsze + 1)
If currentin(0) = 0 Then currentin(0) = 1
GotBytes(0) = GotBytes(0) + Len(dart)
jki = GotBytes(0)
jku = Label8(0).Caption
If jki > jku Then
jop = jki - jku
joh = 1024 = jop
hobo = Mid(dart, 1, joh)
dart = hobo
End If


Put #FileNumber2(0), currentin(0), dart
currentin(0) = currentin(0) + sendsze
filesizeo = Label8(0).Caption
hg = Int((currentin(0) / filesizeo) * 100)
Label4(0).Caption = hg
ratey(0) = ratey(0) + 1
If hg > 99 Then
Label6(0).Caption = "Done!"
Form1.status(HandleWhater(0)).Text = "Done!"
HandleWhater(0).Text = ""

Close FileNumber2(0)
Get0.Close
Exit Sub
End If
Get0.SendData "GIVEMORE"
End If

'---------resume----
If InStr(1, dart, "RSDATA") <> 0 Then
sendsze = 1024
dart = Mid(dart, 7, sendsze + 1)

Label4(0).Caption = hg
GotBytes(0) = GotBytes(0) + Len(dart)
jki = GotBytes(0)
jku = Label8(0).Caption
If jki > jku Then
jop = jki - jku
joh = 1024 = jop
hobo = Mid(dart, 1, joh)
dart = hobo
End If

Put #FileNumber2(0), currentin(0), dart
Get0.SendData "GIVEMORE"
ratey(0) = ratey(0) + 1
currentin(0) = currentin(0) + sendsze
dart = Mid(dart, 7, sendsze + 1)
filesizeo = Label8(0).Caption
hg = Int((currentin(0) / filesizeo) * 100)
If hg > 99 Then
Label6(0).Caption = "Done!"
Form1.status(HandleWhater(0)).Text = "Done!"
HandleWhater(0).Text = ""
Close #FileNumber2(0)
Get0.Close
End If
End If
End Sub

Private Sub Get1_DataArrival(ByVal bytesTotal As Long)
Dim dart As String
Dim jki As Long
Dim jku As Long
Get1.GetData dart
If InStr(1, dart, "FILSZ", vbTextCompare) <> 0 Then
a = Mid(dart, 7, Len(dart) - 6)
Label8(1).Caption = a
End If

If InStr(1, dart, "FILNM", vbTextCompare) <> 0 Then
a = Mid(dart, 7, Len(dart) - 6)
Text2(1).Text = Text2(1).Text & a
currentin(1) = 0
On Error GoTo GetFileFstTime
b = FileLen(Text2(1).Text) 'resume
currentin(1) = b
If b < 1 Then b = 1
Get1.SendData "RSM" & b
Open Text2(1).Text For Binary Access Write As #FileNumber2(1)
GotBytes(1) = currentin(1)
Exit Sub

GetFileFstTime: 'the first time getting the file
Get1.SendData "GTB"
Label4(1).Caption = 0
Open Text2(1).Text For Binary Access Write As #FileNumber2(1)
Label6(1).Caption = "Getting File"
GotBytes(1) = "0"
Exit Sub
End If
'---------normal download
If InStr(1, dart, "NMDATA") <> 0 Then
sendsze = 1024
dart = Mid(dart, 7, sendsze + 1)
If currentin(1) = 0 Then currentin(1) = 1
GotBytes(1) = GotBytes(1) + Len(dart)
jki = GotBytes(1)
jku = Label8(1).Caption
If jki > jku Then
jop = jki - jku
joh = 1024 = jop
hobo = Mid(dart, 1, joh)
dart = hobo
End If


Put #FileNumber2(1), currentin(1), dart
currentin(1) = currentin(1) + sendsze
filesizeo = Label8(1).Caption
hg = Int((currentin(1) / filesizeo) * 100)
Label4(1).Caption = hg
ratey(1) = ratey(1) + 1
If hg > 99 Then
Label6(1).Caption = "Done!"
Get1.Close
Exit Sub
End If
Get1.SendData "GIVEMORE"
End If

'---------resume----
If InStr(1, dart, "RSDATA") <> 0 Then
sendsze = 1024
dart = Mid(dart, 7, sendsze + 1)

Label4(1).Caption = hg
GotBytes(1) = GotBytes(1) + Len(dart)
jki = GotBytes(1)
jku = Label8(1).Caption
If jki > jku Then
jop = jki - jku
joh = 1024 = jop
hobo = Mid(dart, 1, joh)
dart = hobo
End If

Put #FileNumber2(1), currentin(1), dart
Get1.SendData "GIVEMORE"
ratey(1) = ratey(1) + 1
currentin(1) = currentin(1) + sendsze
dart = Mid(dart, 7, sendsze + 1)
filesizeo = Label8(1).Caption
hg = Int((currentin(1) / filesizeo) * 100)
If hg > 99 Then
Label6(1).Caption = "Done!"
Close #FileNumber2(1)
Get1.Close
End If
End If
End Sub

Private Sub Get2_DataArrival(ByVal bytesTotal As Long)
Dim dart As String
Dim jki As Long
Dim jku As Long
Get2.GetData dart
If InStr(1, dart, "FILSZ", vbTextCompare) <> 0 Then
a = Mid(dart, 7, Len(dart) - 6)
Label8(2).Caption = a
End If

If InStr(1, dart, "FILNM", vbTextCompare) <> 0 Then
a = Mid(dart, 7, Len(dart) - 6)
Text2(2).Text = Text2(2).Text & a
currentin(2) = 0
On Error GoTo GetFileFstTime
b = FileLen(Text2(2).Text) 'resume
currentin(2) = b
If b < 1 Then b = 1
Get2.SendData "RSM" & b
Open Text2(2).Text For Binary Access Write As #FileNumber2(2)
GotBytes(2) = currentin(2)
Exit Sub

GetFileFstTime: 'the first time getting the file
Get2.SendData "GTB"
Label4(2).Caption = 0
Open Text2(2).Text For Binary Access Write As #FileNumber2(2)
Label6(2).Caption = "Getting File"
GotBytes(2) = "0"
Exit Sub
End If
'---------normal download
If InStr(1, dart, "NMDATA") <> 0 Then
sendsze = 1024
dart = Mid(dart, 7, sendsze + 1)
If currentin(2) = 0 Then currentin(2) = 1
GotBytes(2) = GotBytes(2) + Len(dart)
jki = GotBytes(2)
jku = Label8(2).Caption
If jki > jku Then
jop = jki - jku
joh = 1024 = jop
hobo = Mid(dart, 1, joh)
dart = hobo
End If


Put #FileNumber2(2), currentin(2), dart
currentin(2) = currentin(2) + sendsze
filesizeo = Label8(2).Caption
hg = Int((currentin(2) / filesizeo) * 100)
Label4(2).Caption = hg
ratey(2) = ratey(2) + 1
If hg > 99 Then
Label6(2).Caption = "Done!"
Get2.Close
Exit Sub
End If
Get2.SendData "GIVEMORE"
End If

'---------resume----
If InStr(1, dart, "RSDATA") <> 0 Then
sendsze = 1024
dart = Mid(dart, 7, sendsze + 1)

Label4(2).Caption = hg
GotBytes(2) = GotBytes(2) + Len(dart)
jki = GotBytes(2)
jku = Label8(2).Caption
If jki > jku Then
jop = jki - jku
joh = 1024 = jop
hobo = Mid(dart, 1, joh)
dart = hobo
End If

Put #FileNumber2(2), currentin(2), dart
Get2.SendData "GIVEMORE"
ratey(2) = ratey(2) + 1
currentin(2) = currentin(2) + sendsze
dart = Mid(dart, 7, sendsze + 1)
filesizeo = Label8(2).Caption
hg = Int((currentin(2) / filesizeo) * 100)
If hg > 99 Then
Label6(2).Caption = "Done!"
Close #FileNumber2(2)
Get2.Close
End If
End If
End Sub
Private Sub Get3_DataArrival(ByVal bytesTotal As Long)
Dim dart As String
Dim jki As Long
Dim jku As Long
Get3.GetData dart
If InStr(1, dart, "FILSZ", vbTextCompare) <> 0 Then
a = Mid(dart, 7, Len(dart) - 6)
Label8(3).Caption = a
End If

If InStr(1, dart, "FILNM", vbTextCompare) <> 0 Then
a = Mid(dart, 7, Len(dart) - 6)
Text2(3).Text = Text2(3).Text & a
currentin(3) = 0
On Error GoTo GetFileFstTime
b = FileLen(Text2(3).Text) 'resume
currentin(3) = b
If b < 1 Then b = 1
Get3.SendData "RSM" & b
Open Text2(3).Text For Binary Access Write As #FileNumber2(3)
GotBytes(3) = currentin(3)
Exit Sub

GetFileFstTime: 'the first time getting the file
Get3.SendData "GTB"
Label4(3).Caption = 0
Open Text2(3).Text For Binary Access Write As #FileNumber2(3)
Label6(3).Caption = "Getting File"
GotBytes(3) = "0"
Exit Sub
End If
'---------normal download
If InStr(1, dart, "NMDATA") <> 0 Then
sendsze = 1024
dart = Mid(dart, 7, sendsze + 1)
If currentin(3) = 0 Then currentin(3) = 1
GotBytes(3) = GotBytes(3) + Len(dart)
jki = GotBytes(3)
jku = Label8(3).Caption
If jki > jku Then
jop = jki - jku
joh = 1024 = jop
hobo = Mid(dart, 1, joh)
dart = hobo
End If


Put #FileNumber2(3), currentin(3), dart
currentin(3) = currentin(3) + sendsze
filesizeo = Label8(3).Caption
hg = Int((currentin(3) / filesizeo) * 100)
Label4(3).Caption = hg
ratey(3) = ratey(3) + 1
If hg > 99 Then
Label6(3).Caption = "Done!"
Get3.Close
Exit Sub
End If
Get3.SendData "GIVEMORE"
End If

'---------resume----
If InStr(1, dart, "RSDATA") <> 0 Then
sendsze = 1024
dart = Mid(dart, 7, sendsze + 1)

Label4(3).Caption = hg
GotBytes(3) = GotBytes(3) + Len(dart)
jki = GotBytes(3)
jku = Label8(3).Caption
If jki > jku Then
jop = jki - jku
joh = 1024 = jop
hobo = Mid(dart, 1, joh)
dart = hobo
End If

Put #FileNumber2(3), currentin(3), dart
Get3.SendData "GIVEMORE"
ratey(3) = ratey(3) + 1
currentin(3) = currentin(3) + sendsze
dart = Mid(dart, 7, sendsze + 1)
filesizeo = Label8(3).Caption
hg = Int((currentin(3) / filesizeo) * 100)
If hg > 99 Then
Label6(3).Caption = "Done!"
Close #FileNumber2(3)
Get3.Close
End If
End If
End Sub
Private Sub Get4_DataArrival(ByVal bytesTotal As Long)
Dim dart As String
Dim jki As Long
Dim jku As Long
Get4.GetData dart
If InStr(1, dart, "FILSZ", vbTextCompare) <> 0 Then
a = Mid(dart, 7, Len(dart) - 6)
Label8(4).Caption = a
End If

If InStr(1, dart, "FILNM", vbTextCompare) <> 0 Then
a = Mid(dart, 7, Len(dart) - 6)
Text2(4).Text = Text2(4).Text & a
currentin(4) = 0
On Error GoTo GetFileFstTime
b = FileLen(Text2(4).Text) 'resume
currentin(4) = b
If b < 1 Then b = 1
Get4.SendData "RSM" & b
Open Text2(4).Text For Binary Access Write As #FileNumber2(4)
GotBytes(4) = currentin(4)
Exit Sub

GetFileFstTime: 'the first time getting the file
Get4.SendData "GTB"
Label4(4).Caption = 0
Open Text2(4).Text For Binary Access Write As #FileNumber2(4)
Label6(4).Caption = "Getting File"
GotBytes(4) = "0"
Exit Sub
End If
'---------normal download
If InStr(1, dart, "NMDATA") <> 0 Then
sendsze = 1024
dart = Mid(dart, 7, sendsze + 1)
If currentin(4) = 0 Then currentin(4) = 1
GotBytes(4) = GotBytes(4) + Len(dart)
jki = GotBytes(4)
jku = Label8(4).Caption
If jki > jku Then
jop = jki - jku
joh = 1024 = jop
hobo = Mid(dart, 1, joh)
dart = hobo
End If


Put #FileNumber2(4), currentin(4), dart
currentin(4) = currentin(4) + sendsze
filesizeo = Label8(4).Caption
hg = Int((currentin(4) / filesizeo) * 100)
Label4(4).Caption = hg
ratey(4) = ratey(4) + 1
If hg > 99 Then
Label6(4).Caption = "Done!"
Get4.Close
Exit Sub
End If
Get4.SendData "GIVEMORE"
End If

'---------resume----
If InStr(1, dart, "RSDATA") <> 0 Then
sendsze = 1024
dart = Mid(dart, 7, sendsze + 1)

Label4(4).Caption = hg
GotBytes(4) = GotBytes(4) + Len(dart)
jki = GotBytes(4)
jku = Label8(4).Caption
If jki > jku Then
jop = jki - jku
joh = 1024 = jop
hobo = Mid(dart, 1, joh)
dart = hobo
End If

Put #FileNumber2(4), currentin(4), dart
Get4.SendData "GIVEMORE"
ratey(4) = ratey(4) + 1
currentin(4) = currentin(4) + sendsze
dart = Mid(dart, 7, sendsze + 1)
filesizeo = Label8(4).Caption
hg = Int((currentin(4) / filesizeo) * 100)
If hg > 99 Then
Label6(4).Caption = "Done!"
Close #FileNumber2(4)
Get4.Close
End If
End If
End Sub
Private Sub Get5_DataArrival(ByVal bytesTotal As Long)
Dim dart As String
Dim jki As Long
Dim jku As Long
Get5.GetData dart
If InStr(1, dart, "FILSZ", vbTextCompare) <> 0 Then
a = Mid(dart, 7, Len(dart) - 6)
Label8(5).Caption = a
End If

If InStr(1, dart, "FILNM", vbTextCompare) <> 0 Then
a = Mid(dart, 7, Len(dart) - 6)
Text2(5).Text = Text2(5).Text & a
currentin(5) = 0
On Error GoTo GetFileFstTime
b = FileLen(Text2(5).Text) 'resume
currentin(5) = b
If b < 1 Then b = 1
Get5.SendData "RSM" & b
Open Text2(5).Text For Binary Access Write As #FileNumber2(5)
GotBytes(5) = currentin(5)
Exit Sub

GetFileFstTime: 'the first time getting the file
Get5.SendData "GTB"
Label4(5).Caption = 0
Open Text2(5).Text For Binary Access Write As #FileNumber2(5)
Label6(5).Caption = "Getting File"
GotBytes(5) = "0"
Exit Sub
End If
'---------normal download
If InStr(1, dart, "NMDATA") <> 0 Then
sendsze = 1024
dart = Mid(dart, 7, sendsze + 1)
If currentin(5) = 0 Then currentin(5) = 1
GotBytes(5) = GotBytes(5) + Len(dart)
jki = GotBytes(5)
jku = Label8(5).Caption
If jki > jku Then
jop = jki - jku
joh = 1024 = jop
hobo = Mid(dart, 1, joh)
dart = hobo
End If


Put #FileNumber2(5), currentin(5), dart
currentin(5) = currentin(5) + sendsze
filesizeo = Label8(5).Caption
hg = Int((currentin(5) / filesizeo) * 100)
Label4(5).Caption = hg
ratey(5) = ratey(5) + 1
If hg > 99 Then
Label6(5).Caption = "Done!"
Get5.Close
Exit Sub
End If
Get5.SendData "GIVEMORE"
End If

'---------resume----
If InStr(1, dart, "RSDATA") <> 0 Then
sendsze = 1024
dart = Mid(dart, 7, sendsze + 1)

Label4(5).Caption = hg
GotBytes(5) = GotBytes(5) + Len(dart)
jki = GotBytes(5)
jku = Label8(5).Caption
If jki > jku Then
jop = jki - jku
joh = 1024 = jop
hobo = Mid(dart, 1, joh)
dart = hobo
End If

Put #FileNumber2(5), currentin(5), dart
Get5.SendData "GIVEMORE"
ratey(5) = ratey(5) + 1
currentin(5) = currentin(5) + sendsze
dart = Mid(dart, 7, sendsze + 1)
filesizeo = Label8(5).Caption
hg = Int((currentin(5) / filesizeo) * 100)
If hg > 99 Then
Label6(5).Caption = "Done!"
Close #FileNumber2(5)
Get5.Close
End If
End If
End Sub
Private Sub Get6_DataArrival(ByVal bytesTotal As Long)
Dim dart As String
Dim jki As Long
Dim jku As Long
Get6.GetData dart
If InStr(1, dart, "FILSZ", vbTextCompare) <> 0 Then
a = Mid(dart, 7, Len(dart) - 6)
Label8(6).Caption = a
End If

If InStr(1, dart, "FILNM", vbTextCompare) <> 0 Then
a = Mid(dart, 7, Len(dart) - 6)
Text2(6).Text = Text2(6).Text & a
currentin(6) = 0
On Error GoTo GetFileFstTime
b = FileLen(Text2(6).Text) 'resume
currentin(6) = b
If b < 1 Then b = 1
Get6.SendData "RSM" & b
Open Text2(6).Text For Binary Access Write As #FileNumber2(6)
GotBytes(6) = currentin(6)
Exit Sub

GetFileFstTime: 'the first time getting the file
Get6.SendData "GTB"
Label4(6).Caption = 0
Open Text2(6).Text For Binary Access Write As #FileNumber2(6)
Label6(6).Caption = "Getting File"
GotBytes(6) = "0"
Exit Sub
End If
'---------normal download
If InStr(1, dart, "NMDATA") <> 0 Then
sendsze = 1024
dart = Mid(dart, 7, sendsze + 1)
If currentin(6) = 0 Then currentin(6) = 1
GotBytes(6) = GotBytes(6) + Len(dart)
jki = GotBytes(6)
jku = Label8(6).Caption
If jki > jku Then
jop = jki - jku
joh = 1024 = jop
hobo = Mid(dart, 1, joh)
dart = hobo
End If


Put #FileNumber2(6), currentin(6), dart
currentin(6) = currentin(6) + sendsze
filesizeo = Label8(6).Caption
hg = Int((currentin(6) / filesizeo) * 100)
Label4(6).Caption = hg
ratey(6) = ratey(6) + 1
If hg > 99 Then
Label6(6).Caption = "Done!"
Get6.Close
Exit Sub
End If
Get6.SendData "GIVEMORE"
End If

'---------resume----
If InStr(1, dart, "RSDATA") <> 0 Then
sendsze = 1024
dart = Mid(dart, 7, sendsze + 1)

Label4(6).Caption = hg
GotBytes(6) = GotBytes(6) + Len(dart)
jki = GotBytes(6)
jku = Label8(6).Caption
If jki > jku Then
jop = jki - jku
joh = 1024 = jop
hobo = Mid(dart, 1, joh)
dart = hobo
End If

Put #FileNumber2(6), currentin(6), dart
Get6.SendData "GIVEMORE"
ratey(6) = ratey(6) + 1
currentin(6) = currentin(6) + sendsze
dart = Mid(dart, 7, sendsze + 1)
filesizeo = Label8(6).Caption
hg = Int((currentin(6) / filesizeo) * 100)
If hg > 99 Then
Label6(6).Caption = "Done!"
Close #FileNumber2(6)
Get6.Close
End If
End If
End Sub
Private Sub Get7_DataArrival(ByVal bytesTotal As Long)
Dim dart As String
Dim jki As Long
Dim jku As Long
Get7.GetData dart
If InStr(1, dart, "FILSZ", vbTextCompare) <> 0 Then
a = Mid(dart, 7, Len(dart) - 6)
Label8(7).Caption = a
End If

If InStr(1, dart, "FILNM", vbTextCompare) <> 0 Then
a = Mid(dart, 7, Len(dart) - 6)
Text2(7).Text = Text2(7).Text & a
currentin(7) = 0
On Error GoTo GetFileFstTime
b = FileLen(Text2(7).Text) 'resume
currentin(7) = b
If b < 1 Then b = 1
Get7.SendData "RSM" & b
Open Text2(7).Text For Binary Access Write As #FileNumber2(7)
GotBytes(7) = currentin(7)
Exit Sub

GetFileFstTime: 'the first time getting the file
Get7.SendData "GTB"
Label4(7).Caption = 0
Open Text2(7).Text For Binary Access Write As #FileNumber2(7)
Label6(7).Caption = "Getting File"
GotBytes(7) = "0"
Exit Sub
End If
'---------normal download
If InStr(1, dart, "NMDATA") <> 0 Then
sendsze = 1024
dart = Mid(dart, 7, sendsze + 1)
If currentin(7) = 0 Then currentin(7) = 1
GotBytes(7) = GotBytes(7) + Len(dart)
jki = GotBytes(7)
jku = Label8(7).Caption
If jki > jku Then
jop = jki - jku
joh = 1024 = jop
hobo = Mid(dart, 1, joh)
dart = hobo
End If


Put #FileNumber2(7), currentin(7), dart
currentin(7) = currentin(7) + sendsze
filesizeo = Label8(7).Caption
hg = Int((currentin(7) / filesizeo) * 100)
Label4(7).Caption = hg
ratey(7) = ratey(7) + 1
If hg > 99 Then
Label6(7).Caption = "Done!"
Get7.Close
Exit Sub
End If
Get7.SendData "GIVEMORE"
End If

'---------resume----
If InStr(1, dart, "RSDATA") <> 0 Then
sendsze = 1024
dart = Mid(dart, 7, sendsze + 1)

Label4(7).Caption = hg
GotBytes(7) = GotBytes(7) + Len(dart)
jki = GotBytes(7)
jku = Label8(7).Caption
If jki > jku Then
jop = jki - jku
joh = 1024 = jop
hobo = Mid(dart, 1, joh)
dart = hobo
End If

Put #FileNumber2(7), currentin(7), dart
Get7.SendData "GIVEMORE"
ratey(7) = ratey(7) + 1
currentin(7) = currentin(7) + sendsze
dart = Mid(dart, 7, sendsze + 1)
filesizeo = Label8(7).Caption
hg = Int((currentin(7) / filesizeo) * 100)
If hg > 99 Then
Label6(7).Caption = "Done!"
Close #FileNumber2(7)
Get7.Close
End If
End If
End Sub
'=============End Get_Dataarrival Code=======



'=============Start Get_Close code===========
Private Sub Get0_Close()
a = Label4(0).Caption
If a > 99 Then
Label6(0).Caption = "Done!"
Form1.status(HandleWhater(0)).Text = "Done!"
Close #FileNumber2(0)
Else
Label6(0).Caption = "Canceled"
Close #FileNumber2(0)
End If
HandleWhater(0).Text = ""
Form1.status(HandleWhater(0)).Text = "Done!"
End Sub
Private Sub Get1_Close()
a = Label4(1).Caption
If a > 99 Then
Label6(1).Caption = "Done!"
Close #FileNumber2(1)
Else
Label6(1).Caption = "Canceled"
Close #FileNumber2(1)
End If
End Sub
Private Sub Get2_Close()
a = Label4(2).Caption
If a > 99 Then
Label6(2).Caption = "Done!"
Close #FileNumber2(2)
Else
Label6(2).Caption = "Canceled"
Close #FileNumber2(2)
End If
End Sub
Private Sub Get3_Close()
a = Label4(3).Caption
If a > 99 Then
Label6(3).Caption = "Done!"
Close #FileNumber2(3)
Else
Label6(3).Caption = "Canceled"
Close #FileNumber2(3)
End If
End Sub
Private Sub Get4_Close()
a = Label4(4).Caption
If a > 99 Then
Label6(4).Caption = "Done!"
Close #FileNumber2(4)
Else
Label6(4).Caption = "Canceled"
Close #FileNumber2(4)
End If
End Sub
Private Sub Get5_Close()
a = Label4(5).Caption
If a > 99 Then
Label6(5).Caption = "Done!"
Close #FileNumber2(5)
Else
Label6(5).Caption = "Canceled"
Close #FileNumber2(5)
End If
End Sub
Private Sub Get6_Close()
a = Label4(6).Caption
If a > 99 Then
Label6(6).Caption = "Done!"
Close #FileNumber2(6)
Else
Label6(6).Caption = "Canceled"
Close #FileNumber2(6)
End If
End Sub
Private Sub Get7_Close()
a = Label4(7).Caption
If a > 99 Then
Label6(7).Caption = "Done!"
Close #FileNumber2(7)
Else
Label6(7).Caption = "Canceled"
Close #FileNumber2(7)
End If
End Sub
'================End Get_Close Code=======



'=============Start Get_Error Code=====
Private Sub Get0_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Label6(0).Caption = "Canceled"
Get0.Close
End Sub
Private Sub Get1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Label6(1).Caption = "Canceled"
Get1.Close
End Sub
Private Sub Get2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Label6(2).Caption = "Canceled"
Get2.Close
End Sub
Private Sub Get3_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Label6(3).Caption = "Canceled"
Get3.Close
End Sub
Private Sub Get4_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Label6(4).Caption = "Canceled"
Get4.Close
End Sub
Private Sub Get5_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Label6(5).Caption = "Canceled"
Get5.Close
End Sub
Private Sub Get6_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Label6(6).Caption = "Canceled"
Get6.Close
End Sub
Private Sub Get7_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Label6(7).Caption = "Canceled"
Get7.Close
End Sub
'===========End Get_Error Code=========



'===========Start Get_ConnectionRequest code=====
Private Sub Get0_ConnectionRequest(ByVal requestID As Long)
Get0.Close
Get0.Accept requestID
Label6(0).Caption = "Connected"
End Sub
Private Sub Get1_ConnectionRequest(ByVal requestID As Long)
Get1.Close
Get1.Accept requestID
Label6(1).Caption = "Connected"
End Sub
Private Sub Get2_ConnectionRequest(ByVal requestID As Long)
Get2.Close
Get2.Accept requestID
Label6(2).Caption = "Connected"
End Sub
Private Sub Get3_ConnectionRequest(ByVal requestID As Long)
Get3.Close
Get3.Accept requestID
Label6(3).Caption = "Connected"
End Sub
Private Sub Get4_ConnectionRequest(ByVal requestID As Long)
Get4.Close
Get4.Accept requestID
Label6(4).Caption = "Connected"
End Sub
Private Sub Get5_ConnectionRequest(ByVal requestID As Long)
Get5.Close
Get5.Accept requestID
Label6(5).Caption = "Connected"
End Sub
Private Sub Get6_ConnectionRequest(ByVal requestID As Long)
Get6.Close
Get6.Accept requestID
Label6(6).Caption = "Connected"
End Sub
Private Sub Get7_ConnectionRequest(ByVal requestID As Long)
Get7.Close
Get7.Accept requestID
Label6(7).Caption = "Connected"
End Sub
'===================End Get_Connectionrequest code====
