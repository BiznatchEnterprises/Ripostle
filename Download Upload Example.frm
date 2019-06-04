VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   2640
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   0
      Top             =   2160
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2520
      TabIndex        =   12
      Text            =   "c:\"
      Top             =   840
      Width           =   1575
   End
   Begin VB.FileListBox File1 
      Height          =   870
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   1575
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   3000
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   360
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "add path"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "listen"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label SizeIHave 
      Height          =   255
      Left            =   3600
      TabIndex        =   15
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Size:"
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "Size:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Not Connected"
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Not Connected"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Done:"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Done:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Speed:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Line Line1 
      X1              =   2280
      X2              =   2280
      Y1              =   120
      Y2              =   3000
   End
   Begin VB.Label Label3 
      Caption         =   "Speed:"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   1440
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim thename
Dim givemore As String
Dim currentint As Long
Dim currentin As Long
Dim ratee As Integer
Dim ratey As Integer
Sub Sleep(ByVal MillaSec As Long, Optional ByVal DeepSleep As Boolean = False)
    Dim tStart#, Tmr#
    tStart = Timer

    While Tmr < (MillaSec / 1000)
        Tmr = Timer - tStart
        If DeepSleep = False Then DoEvents
    Wend
End Sub
Private Sub Command1_Click()
Winsock1.Connect Text1.Text, 7473
Label5.Caption = "Connecting"
End Sub

Private Sub Command2_Click()
Winsock2.Close
Winsock2.LocalPort = 7473
Winsock2.Listen
Label6.Caption = "Listening"
End Sub

Private Sub Command3_Click()
Dim siz As String
Text3.Text = File1.Path & "\" & File1.FileName
thename = File1.FileName
siz = FileLen(Text3.Text)
Label7.Caption = "Size: " & siz

End Sub

Private Sub Form_Load()
File1.Path = "d:\mp3z"
End Sub

Private Sub Timer1_Timer()
'2 second timer that calculates the rate
'NOTE: If you change the sendsize you will have to change this too!
Label1.Caption = "Speed: " & (ratee / 2)
ratee = 0
End Sub

Private Sub Timer2_Timer()
Label3.Caption = "Speed: " & (ratey / 2)
ratey = 0
End Sub

Private Sub Winsock1_Close()
a = Mid(Label2.Caption, 6, 4)
If a = 100 Then
Label5.Caption = "Done!"
Else
Label5.Caption = "Canceled"
End If
End Sub

Private Sub Winsock1_Connect()
Label5.Caption = "Connected!"
Winsock1.SendData "FILSZ " & Mid(Label7.Caption, 6, Len(Label7.Caption) - 5)
Sleep 2000
Winsock1.SendData "FILNM " & thename
End Sub
Sub sendfile(ByVal Position)
currentint = currentint + Position
On Error Resume Next
Dim tempbuffer As String
Open Text3.Text For Binary Access Read As #2

tempbuffer = Space$(1024)
Get #2, Position, tempbuffer
Winsock1.SendData "SXDAT " & tempbuffer
givemore = 0


Do Until EOF(2)

Do Until givemore = 1
DoEvents
Loop
tempbuffer = Space$(1024)

Get #2, , tempbuffer


currentint = currentint + 1024
filesizeu = Mid(Label7.Caption, 6, Len(Label7.Caption) - 5)

Label2.Caption = "Done: " & Int((currentint / filesizeu) * 100)
ratee = ratee + 1

Winsock1.SendData "STDAT " & tempbuffer
givemore = 0
Loop
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim dart As String
Winsock1.GetData dart

If InStr(1, dart, "GTB", vbTextCompare) <> 0 Then
Label5.Caption = "Sending File"
givemore = 1
SendTheFile
End If

If InStr(1, dart, "RSM", vbTextCompare) <> 0 Then
a = Mid(dart, 4, Len(dart) - 3)
currentint = a
Label5.Caption = "Sending File"
givemore = 1
SendRSMFile
End If

If dart = "GIVEMORE" Then givemore = 1

End Sub
Sub SendRSMFile()
Dim tempbuffer As String
On Error GoTo errhand
Open Text3.Text For Binary Access Read As #2
Label2.Caption = "Done: " & 0

Do Until EOF(2)
Do Until givemore = 1
DoEvents
Loop
tempbuffer = Space$(1024)

Get #2, currentint, tempbuffer


currentint = currentint + 1024
filesizeu = Mid(Label7.Caption, 6, Len(Label7.Caption) - 5)

Label2.Caption = "Done: " & Int((currentint / filesizeu) * 100)
ratee = ratee + 1

Winsock1.SendData "NMDATA" & tempbuffer
givemore = 0
Loop

Close #2
Winsock1.Close
Label5.Caption = "Done!"
errhand:
Winsock1.Close
Exit Sub
End Sub

Sub SendTheFile()
Dim tempbuffer As String
On Error GoTo errhand
Open Text3.Text For Binary Access Read As #2
Label2.Caption = "Done: " & 0
currentint = 0

Do Until EOF(2)
Do Until givemore = 1
DoEvents
Loop
tempbuffer = Space$(1024)

Get #2, , tempbuffer


currentint = currentint + 1024
filesizeu = Mid(Label7.Caption, 6, Len(Label7.Caption) - 5)

Label2.Caption = "Done: " & Int((currentint / filesizeu) * 100)
ratee = ratee + 1

Winsock1.SendData "NMDATA" & tempbuffer
givemore = 0
Loop

Close #2
Winsock1.Close
Label5.Caption = "Done!"
errhand:
Winsock1.Close
Exit Sub
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Label5.Caption = "Canceled"
Winsock1.Close
End Sub

Private Sub Winsock2_Close()
a = Mid(Label4.Caption, 6, 4)
If a = 100 Then
Label6.Caption = "Done!"
Close #5
Else
Label6.Caption = "Canceled"
End If

End Sub

Private Sub Winsock2_ConnectionRequest(ByVal requestID As Long)
Winsock2.Close
Winsock2.Accept requestID
Label6.Caption = "Connected"
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
Dim dart As String
Winsock2.GetData dart
If InStr(1, dart, "FILSZ", vbTextCompare) <> 0 Then
a = Mid(dart, 7, Len(dart) - 6)
Label8.Caption = "Size: " & a
End If

If InStr(1, dart, "FILNM", vbTextCompare) <> 0 Then
a = Mid(dart, 7, Len(dart) - 6)
Text2.Text = Text2.Text & a
currentin = 0
On Error GoTo GetFileFstTime
b = FileLen(Text2.Text) 'resume
currentin = b
If b < 1 Then b = 1
Winsock2.SendData "RSM" & b
Open Text2.Text For Binary Access Write As #5
Exit Sub

GetFileFstTime: 'the firts time getting the file
Winsock2.SendData "GTB"
Label4.Caption = "Done: " & 0
Open Text2.Text For Binary Access Write As #5
Label6.Caption = "Getting File"
Exit Sub
End If
'---------normal download
If InStr(1, dart, "NMDATA") <> 0 Then
sendsze = 1024
currentin = currentin + sendsze
dart = Mid(dart, 7, sendsze + 1)
filesizeo = Mid(Label8.Caption, 6, Len(Label8.Caption) - 5)
hg = Int((currentin / filesizeo) * 100)
If hg = 99 Then
hg = 100
Label6.Caption = "Done!"

End If
Label4.Caption = "Done: " & hg
Put #5, currentin, dart
Winsock2.SendData "GIVEMORE"
ratey = ratey + 1
End If
'---------resume----
If InStr(1, dart, "RSDATA") <> 0 Then
sendsze = 1024
Label4.Caption = "Done: " & hg
Put #5, currentin, dart
Winsock2.SendData "GIVEMORE"
ratey = ratey + 1

currentin = currentin + sendsze
dart = Mid(dart, 7, sendsze + 1)
filesizeo = Mid(Label8.Caption, 6, Len(Label8.Caption) - 5)
hg = Int((currentin / filesizeo) * 100)
If hg = 99 Then
hg = 100
Label6.Caption = "Done!"
End If
End If
End Sub

Private Sub Winsock2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Label5.Caption = "Canceled"
Winsock1.Close
End Sub
