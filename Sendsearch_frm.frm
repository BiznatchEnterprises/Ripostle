VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Sendsearch_frm 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2175
   ClientLeft      =   10635
   ClientTop       =   6750
   ClientWidth     =   1635
   LinkTopic       =   "Form2"
   ScaleHeight     =   2175
   ScaleWidth      =   1635
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   -120
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Timer searchtimout 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox mart 
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Text            =   "0"
      Top             =   480
      Width           =   375
   End
   Begin MSWinsockLib.Winsock sendsearch 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   7471
   End
End
Attribute VB_Name = "Sendsearch_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim agx As String
Dim folt As String
Dim word(6)
Dim ahm As String
Dim posiptocontt
Sub search4song()
Dim hds(2)

If mart = 0 Then Exit Sub

If Form1.zort.Text = Form1.List1.ListCount Then
Form1.Label36.Caption = "Idle (Done)"
Form1.stopsearch_button.Enabled = False
Sendsearch_frm.mart.Text = "0"

sendsearch.Close
Exit Sub
Else


If sendsearch.State = 0 Then
posiptocontt = Form1.zort.Text
iptocontt = Form1.List1.List(Form1.zort.Text)
mactoc = 1
For i = 1 To Len(iptocontt)
a = Mid(iptocontt, i, 1)
If a = ";" Then
a = ""
mactoc = mactoc + 1
End If
hds(mactoc) = hds(mactoc) & a
Next
If Form1.List1.List(Form1.zort.Text) = "" Then Exit Sub
sendsearch.Connect hds(1), 7471
searchtimout.Enabled = True
End If
End If
End Sub
Public Function DateTagIP(ByVal IPPositionOnList1 As String)
Dim hds(2)
mactoc = 1
ipz = Form1.List1.List(IPPositionOnList1)
For i = 1 To Len(Form1.List1.List(IPPositionOnList1))
a = Mid(ipz, i, 1)
If a = ";" Then
a = ""
mactoc = mactoc + 1
End If
hds(mactoc) = hds(mactoc) & a
Next

a = DatePart("d", Now)
b = DatePart("m", Now)
c = DatePart("yyyy", Now)

If b < 10 Then b = "0" & b

hds(2) = a & "/" & b & "/" & c


Form1.List1.List(IPPositionOnList1) = hds(1) & ";" & hds(2)
End Function

Private Sub mart_Change()
If mart.Text = "0" Then
Form1.Search_Button.Enabled = True
Form1.stopsearch_button.Enabled = False
If Form1.Label6.Text < "0" Then
Form1.search_results.TextMatrix(1, 0) = "File Not Found! Check your"
Form1.search_results.Rows = 4
Form1.search_results.TextMatrix(2, 0) = "Spelling, Or Try Again In"
Form1.search_results.TextMatrix(3, 0) = "A While... Thank You!"
End If

Form1.Label36.Caption = "Idle (Done)"
Conv_listtosearchresults
On Error GoTo gf
sendsearch.SendData "cs"
Sleep 1000
End If

gf:
Exit Sub
End Sub

Sub Conv_listtosearchresults()
For h = 0 To List1.ListCount - 1
koh = 0
    
  For g = 0 To 6
  word(g) = ""
  Next
  
  For w = 1 To Len(List1.List(h))
  ahm = Mid(List1.List(h), w, 1)
  If ahm = ";" Then
  ahm = ""
  koh = koh + 1
  End If
  word(koh) = word(koh) & ahm
  Next w

word(1) = Format(word(1), "#.##")
If Form1.file_type.Text = "MP3" Then word(3) = Format(word(3), "#.##")

a = Form1.search_results.Rows
Form1.search_results.Rows = Form1.search_results.Rows + 1

a = a - 2
If Form1.search_results.TextMatrix(a - 1, 0) = "Searching... Please wait!" Then a = a - 1
If Form1.search_results.TextMatrix(a - 1, 0) = "" Then a = a - 1
For y = 0 To 6

If word(y) <> "le" Then Form1.search_results.TextMatrix(a, y) = word(y)


Next

Next h
Form1.Label6.Text = Form1.Label6.Text + List1.ListCount - 1
'Form1.draw_Search_Results_Window
List1.Clear



End Sub

Private Sub searchtimout_Timer()
sendsearch.Close
Form1.zort.Text = Form1.zort.Text + 1
search4song
End Sub

Private Sub sendsearch_Connect()
DateTagIP posiptocontt
End Sub

Private Sub sendsearch_DataArrival(ByVal bytesTotal As Long)
sendsearch.GetData agx

If agx = "le" Then
Conv_listtosearchresults
sendsearch.SendData "cs"
Sleep 1000
Sendsearch_frm.sendsearch.Close
Form1.zort = Form1.zort + 1
If Form1.search_results.Rows > 200 Then mart = 0
If mart <> 0 Then search4song
End If

If agx = "nf" Then
Conv_listtosearchresults
sendsearch.SendData "cs"
Sleep 1000
Sendsearch_frm.sendsearch.Close
Form1.zort = Form1.zort + 1
If Form1.search_results.Rows > 200 Then mart = 0
If mart <> 0 Then search4song
End If

If agx = "cntd" Then
searchtimout.Enabled = False
If Form1.file_type = "MP3" Then jax = "!"
If Form1.file_type = "Lyric" Then jax = "@"
If Form1.file_type = "Drumb Tab" Then jax = "#"
If Form1.file_type = "Guitar Tab" Then jax = "$"
If Form1.file_type = "Midi" Then jax = "%"
If Form1.file_type = "Sheet Music" Then jax = "^"

xaj = jax & Form1.artistname_txt.Text & "*" & Form1.songname_txt.Text
sendsearch.SendData xaj
Exit Sub
End If

If agx <> "nf" Then List1.AddItem agx
End Sub

Sub Sleep(ByVal MillaSec As Long, Optional ByVal DeepSleep As Boolean = False)
    Dim tStart#, Tmr#
    tStart = Timer


    While Tmr < (MillaSec / 1000)
        Tmr = Timer - tStart
        If DeepSleep = False Then DoEvents
    Wend
End Sub
