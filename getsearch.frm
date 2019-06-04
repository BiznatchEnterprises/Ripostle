VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form getsearch_frm 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   ClientHeight    =   2340
   ClientLeft      =   10635
   ClientTop       =   3825
   ClientWidth     =   2085
   LinkTopic       =   "Form2"
   ScaleHeight     =   2340
   ScaleWidth      =   2085
   ShowInTaskbar   =   0   'False
   Begin MSWinsockLib.Winsock getsearch 
      Left            =   960
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   7471
   End
   Begin VB.FileListBox search_results 
      Height          =   1650
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "getsearch_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim agx As String
Dim folt As String
Dim word(5)
Dim ahm As String
'change zort to a txt box on form1
Private Sub Form_Load()

getsearch.Listen
End Sub

Private Sub getsearch_ConnectionRequest(ByVal requestID As Long)
getsearch.Close
getsearch.Accept requestID
getsearch.SendData "cntd"
End Sub

Private Sub getsearch_DataArrival(ByVal bytesTotal As Long)
getsearch.GetData folt

If folt = "cs" Then 'close connection
Sleep 2000
getsearch.Close
Form1.reload_getsearch_frm
Exit Sub
End If


asd = Mid(folt, 1, 1)
If asd = "!" Then kom = ".mp3"
If asd = "@" Then kom = "[Lyric].txt"
If asd = "#" Then kom = "[Dtab].txt"
If asd = "$" Then kom = "[Gtab].txt"
If asd = "^" Then kom = "[Sheet Music].txt"
If asd = "%" Then kom = ".mid"

asdf = Mid(folt, 2, Len(folt))
asdfg = asdf & "*" & kom
asf = Replace(asdfg, " ", "*", , , vbTextCompare)
asf = "*" & asf

search_results.Pattern = asf


Sleep 10

If search_results.List(0) = "" Then
getsearch.SendData "nf"
Sleep 2000
Form1.reload_getsearch_frm
Exit Sub
End If

For b = 0 To search_results.ListCount - 1
On Error Resume Next
filsizea = FileLen(search_results.Path & "\" & search_results.List(b))
filsizea = filsizea / 1000024
filsizea = Format(filsizea, "#.##")


  Dim accMP3Info As MP3Info
  agoh = search_results.Path & "\" & search_results.List(b)
  getMP3Info agoh, accMP3Info

If accMP3Info.LENGTH <> "N/F" Then lengtha = accMP3Info.LENGTH / 60
Bitratea = accMP3Info.BITRATE

IpGetter.Visible = True
ipo = IpGetter.IPLIST.List(0)
Unload IpGetter
If search_results.List(b) <> "" Then getsearch.SendData search_results.List(b) & ";" & lengtha & ";" & Form1.username.Text & ";" & filsizea & ";" & Bitratea & ";" & Form1.ModemSpeed.Text & ";" & ipo
Sleep 200
Next
getsearch.SendData "le"






End Sub




Sub Sleep(ByVal MillaSec As Long, Optional ByVal DeepSleep As Boolean = False)
    Dim tStart#, Tmr#
    tStart = Timer


    While Tmr < (MillaSec / 1000)
        Tmr = Timer - tStart
        If DeepSleep = False Then DoEvents
    Wend
End Sub

