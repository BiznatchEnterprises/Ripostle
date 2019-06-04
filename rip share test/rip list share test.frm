VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Ripostle List share method test...."
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Load List from file"
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   4560
      Width           =   855
   End
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   2160
      TabIndex        =   4
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "enable listening"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   1815
      ItemData        =   "rip list share test.frx":0000
      Left            =   0
      List            =   "rip list share test.frx":000D
      TabIndex        =   1
      Top             =   2640
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "connect"
      Height          =   735
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock sock 
      Left            =   1200
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   28144
   End
   Begin MSWinsockLib.Winsock foot 
      Left            =   720
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   28144
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1800
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
'***********
'remember to delete this line and enable the one under it
'***********
iptocont = List1.List(0)
'iptocont = List1.List(Int(Rnd * List1.ListCount))
foot.Close
foot.Connect iptocont, 28144

End Sub
Private Sub RemoveDupes(lst As ListBox)
    Dim iPos As Integer
    iPos = 0
    '-- if listbox empty then exit..
    If lst.ListCount < 1 Then Exit Sub


    Do While iPos < lst.ListCount
        lst.Text = lst.List(iPos)
        '-- check if text already exists..


        If lst.ListIndex <> iPos Then
            '-- if so, remove it and keep iPos..
            lst.RemoveItem iPos
        Else
            '-- if not, increase iPos..
            iPos = iPos + 1
        End If
    Loop
    '-- used to unselect the last selected l
    '     ine..
    lst.Text = "~~~^^~~~"
End Sub


Sub rmdupandmyip()
For i = 0 To List1.ListCount
If List1.List(i) = sock.LocalIP Then List1.RemoveItem (i)
Next

For i = 0 To List1.ListCount
If List1.List(i) = foot.LocalIP Then List1.RemoveItem (i)
Next

RemoveDupes List1
End Sub

Sub slcaller()
'sends the person who diald their list
For i = 0 To List1.ListCount
foot.SendData List1.List(i)
Sleep 200, False
Next
foot.SendData sock.LocalIP
Sleep 200, False
foot.SendData "xx"
End Sub
Sub slcallie()
For i = 0 To List1.ListCount
sock.SendData List1.List(i)
Sleep 200, False
Next

sock.SendData "x"
End Sub

Private Sub Command2_Click()
Open "ipcfg" For Input As #1
List1.Clear
For i = 1 To LOF(1)
If Not EOF(1) Then
Line Input #1, hedgehog
List1.AddItem hedgehog
End If
Next







End Sub

Private Sub Command4_Click()
sock.Close
sock.Listen
End Sub
Sub rmdupesandmyipandmerge()
For i = 0 To List2.ListCount
List1.AddItem List2.List(i)
Next
List2.Clear

For i = 0 To List1.ListCount
If List1.List(i) = sock.LocalIP Then List1.RemoveItem (i)
Next

For i = 0 To List1.ListCount
If List1.List(i) = foot.LocalIP Then List1.RemoveItem (i)
Next

RemoveDupes List1
End Sub

Private Sub foot_DataArrival(ByVal bytesTotal As Long)
Dim ag As String
foot.GetData ag
If ag = "x" Then
slcaller
rmdupesandmyipandmerge
Else
List2.AddItem ag
Label1.Caption = List1.ListCount
End If

End Sub

Private Sub sock_ConnectionRequest(ByVal requestID As Long)
sock.Close
sock.Accept requestID
Sleep 300
slcallie
End Sub

Private Sub sock_DataArrival(ByVal bytesTotal As Long)
Dim ag As String
sock.GetData ag
If ag = "xx" Then
rmdupandmyip
Else
List1.AddItem ag
Label1.Caption = List1.ListCount
End If
End Sub

Sub Sleep(ByVal MillaSec As Long, Optional ByVal DeepSleep As Boolean = False)
    Dim tStart#, Tmr#
    tStart = Timer


    While Tmr < (MillaSec / 1000)
        Tmr = Timer - tStart
        If DeepSleep = False Then DoEvents
    Wend
End Sub
