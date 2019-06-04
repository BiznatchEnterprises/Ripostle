VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form getfile1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nestor's File Trany"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox HandleWhatFile 
      Height          =   285
      Left            =   720
      TabIndex        =   8
      Top             =   2400
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   2400
      Top             =   1680
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3360
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar recieved 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.TextBox txtfile 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "C:\Windows\Desktop\"
      Top             =   840
      Width           =   3495
   End
   Begin VB.CommandButton cmdlisten 
      Caption         =   "Listen for File"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lbltime 
      Caption         =   "Time Left"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   4455
   End
   Begin VB.Label lblstatus 
      Alignment       =   2  'Center
      Caption         =   "No Connection"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Save to folder/Source file:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label thespeed 
      Caption         =   "Speed"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Label byteslabel 
      Caption         =   "Bytes Label"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   4455
   End
End
Attribute VB_Name = "getfile1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'created by David Halter
'you may use this code as you wish
'if you do use it, just email me


'for pauses in between
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Dim mainbuffer As String
Dim sendsize As Integer
Dim sendmore As Integer
Dim thename As String
Dim filesize As Long
Dim currentint As Long
Dim rate As Integer
Dim filestart As Long




Private Sub Command4_Click()
'exit
End

End Sub
Private Sub cmdlisten_Click()
'start listening
Winsock1.Close
Winsock1.LocalPort = 7473
Winsock1.Listen
AddStat "Listening"
End Sub

Public Sub AddStat(message As String)
'changes stat message
lblstatus.Caption = message


End Sub

Private Sub Form_Load()
'sets size for each package
sendsize = 1024
cmdlisten_Click

End Sub

Private Sub Form_Unload(Cancel As Integer)
'makes sure it will end
End

End Sub

Private Sub Timer1_Timer()
On Error GoTo timeerror
'2 second timer that calculates the rate
'NOTE: If you change the sendsize you will have to change this too!

thespeed.Caption = "Speed: " & (rate / 2) & " KB/second (" & ((rate / 2) * 8) & " KBits/second)"

If ((filesize - currentint) / ((rate / 2) * 1024)) <= 60 Then
    lbltime.Caption = "Time left: " & Int((filesize - currentint) / ((rate / 2) * 1024)) & " seconds"
ElseIf ((filesize - currentint) / ((rate / 2) * 1024)) > 60 And ((filesize - currentint) / ((rate / 2) * 1024)) <= 120 Then
    lbltime.Caption = "Time left: 1 minute"
ElseIf ((filesize - currentint) / ((rate / 2) * 1024)) >= 120 And Int(Int((filesize - currentint) / ((rate / 2) * 1024)) / 60) < 60 Then
    lbltime.Caption = "Time left: " & Int(Int((filesize - currentint) / ((rate / 2) * 1024)) / 60) & " minutes"
ElseIf ((filesize - currentint) / ((rate / 2) * 1024)) > 0 Then
    lbltime.Caption = "Time left: " & Int(Int(Int((filesize - currentint) / ((rate / 2) * 1024)) / 60) / 60) & " hours"
End If
rate = 0

Exit Sub
timeerror:
    lbltime.Caption = "Time left: Infinity"
End Sub

Private Sub Winsock1_Close()
'make sure it closes
Winsock1.Close

AddStat "Connection Closed"
End Sub

Private Sub Winsock1_Connect()
'when connected start sending
AddStat "Sending Information"
Winsock1.SendData "FILESIZE " & FileLen(txtfile.Text)
DoEvents
Sleep 1000
Winsock1.SendData "SENDNAME " & thename
DoEvents
Sleep 1000
AddStat "Sending File"
SendFile
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
'accept the request
AddStat "Connected"
Winsock1.Close
Winsock1.Accept requestID
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'get information
On Error Resume Next
Dim temprecieve As String
DoEvents
Winsock1.GetData temprecieve

If InStr(1, temprecieve, "FILESIZE ") <> 0 Then
    filesize = Mid(temprecieve, 10, sendsize)
    'sets file size for progressbar etc.
ElseIf InStr(1, temprecieve, "SENDNAME ") <> 0 Then
    'set filename of file to be saved
    thename = Mid(temprecieve, 10, sendsize)
    currentint = 0
    currentint = FileLen(txtfile.Text & thename)
    'see if already done
    If currentint >= filesize Then
        Winsock1.SendData "ALLDONE"
        DoEvents
        Close #2
        Winsock1.Close
        AddStat "DONE!"
        Beep
        Exit Sub
    Else
        'resume!
        currentint = currentint + 1
        Winsock1.SendData "RESUME " & currentint
    End If
    Close #2
    'open file for editing
    Open txtfile.Text & thename For Binary Access Write As #2
    
ElseIf InStr(1, temprecieve, "THEDATA ") <> 0 Then
    'get the actual data
    temprecieve = Mid(temprecieve, 9, sendsize + 1)
    currentint = currentint + sendsize
    recieved.Value = Int(currentint / filesize * 100)
    AddStat "Recieved " & Int((currentint / filesize) * 100) & "%"
    Me.Caption = "Nestor's File Trany - " & Int((currentint / filesize) * 100) & "%" 'title bar
    'place in file
    Put #2, , temprecieve
    Winsock1.SendData "SENDMORE"
    'sends a 'sendmore' to get more data
    rate = rate + 1
    byteslabel = "Recieved " & currentint & " of " & filesize & " bytes"
ElseIf InStr(1, temprecieve, "SRTDATA ") <> 0 Then
    'get the actual data started
    temprecieve = Mid(temprecieve, 9, sendsize + 1)
    recieved.Value = Int(currentint / filesize * 100)
    AddStat "Recieved " & Int((currentint / filesize) * 100) & "%"
    Me.Caption = "Nestor's File Trany - " & Int((currentint / filesize) * 100) & "%" 'title bar
    'place in file
    Put #2, currentint, temprecieve
    Winsock1.SendData "SENDMORE"
    'sends a 'sendmore' to get more data
    rate = rate + 1
    byteslabel = "Recieved " & currentint & " of " & filesize & " bytes"
    currentint = currentint + sendsize
ElseIf InStr(1, temprecieve, "RESUME ") <> 0 Then
    'set resume position
    currentint = Mid(temprecieve, 8, 20)
    filestart = currentint
    sendmore = 1
ElseIf temprecieve = "SENDMORE" Then
    'sendmore data to reciever
    sendmore = 1
ElseIf temprecieve = "ALLDONE" Then
    'close file and say 'DONE'
    Close #2
    Winsock1.Close
    AddStat "DONE!"
    Beep
End If

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'error
Winsock1.Close
AddStat "Error Transfering"
End Sub

Public Sub SendFile()
'send the puppy
On Error GoTo tranyerror
Dim tempbuffer As String

sendmore = 0
currentint = 0
filestart = 0

Do Until sendmore = 1
DoEvents
Loop

filesize = FileLen(txtfile.Text)
'open file to get info
Close #1
Open txtfile.Text For Binary Access Read As #1

tempbuffer = Space$(sendsize)
'start where needed to resume... if not it will start at 1 anyway
Get #1, filestart, tempbuffer

Winsock1.SendData "SRTDATA " & tempbuffer
sendmore = 0


Do Until EOF(1)

Do Until sendmore = 1
DoEvents
Loop

'set the buffer size

tempbuffer = Space$(sendsize)

Get #1, , tempbuffer


'stat stuff
currentint = currentint + sendsize
recieved.Value = Int(currentint / filesize * 100) 'progressbar
AddStat "Sent " & Int((currentint / filesize) * 100) & "%" 'percent sent
Me.Caption = "Nestor's File Trany - " & Int((currentint / filesize) * 100) & "%" 'title bar
byteslabel = "Sent " & currentint & " of " & filesize & " bytes" 'bytes sent
rate = rate + 1 'used for rate

'send data
Winsock1.SendData "THEDATA " & tempbuffer
sendmore = 0
'dont do anything until the other machine says you can




Loop

On Error Resume Next
Sleep 500
'ALL DONE!
Close #1
DoEvents
Winsock1.SendData "ALLDONE"
DoEvents
Sleep 500
DoEvents
Winsock1.Close
DoEvents
AddStat "DONE!"
Exit Sub

tranyerror:
'error
MsgBox "Transmission Error"
Winsock1.Close
AddStat "Error Transfering"
Exit Sub

End Sub
