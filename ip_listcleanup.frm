VERSION 5.00
Begin VB.Form ip_listcleanup 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   2385
   ClientLeft      =   9960
   ClientTop       =   1230
   ClientWidth     =   3765
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List4 
      Height          =   450
      Left            =   1680
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.ListBox List3 
      Height          =   450
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "ip_listcleanup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" _
         (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
          ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const HWND_TOPMOST = -1
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Private Sub RemoveDupesz(lst As ListBox)
On Error Resume Next
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
            List4.RemoveItem iPos
          
        Else
            '-- if not, increase iPos..
            iPos = iPos + 1
        End If
    Loop
    '-- used to unselect the last selected l
    '     ine..
    lst.Text = "~~~^^~~~"
End Sub
Sub RMDUPES()
For i = 0 To Form1.List1.ListCount - 1
z = InStr(1, Form1.List1.List(i), ";")
If z <> 0 Then
x = Mid(Form1.List1.List(i), 1, z - 1)
y = Mid(Form1.List1.List(i), z + 1, Len(Form1.List1.List(i)) - z)
List3.AddItem x
List4.AddItem y
Else
List3.AddItem Form1.List1.List(i)
List4.AddItem "Null"
End If
Next
RemoveDupesz List3

Form1.List1.Clear
For b = 0 To List3.ListCount - 1
j = List3.List(b)
k = List4.List(b)
If k = "Null" Then k = ""
If k = "" Then Form1.List1.AddItem j Else Form1.List1.AddItem j & ";" & k
Next

List3.Clear
List4.Clear
End Sub

Private Sub Form_Load()
Call PutWindowOnTop(Me)
If a = 109 Then
For i = 0 To Form1.List1.ListCount - 1
c = Form1.List1.List(i)
a = InStr(1, Form1.List1.List(i), ".", vbTextCompare)
b = Mid(Form1.List1.List(i), 1, a - 1)
h = InStr(1, Form1.List1.List(i), "d2g.com", vbTextCompare)
If h <> 0 Then b = 0
If b < 127 Then List1.AddItem c
Next


For i = 0 To List1.ListCount - 1
a = InStr(1, List1.List(i), ";", vbTextCompare)
If a = "0" Then List2.AddItem List1.List(i)
d = InStr(1, List1.List(i), "d2g.com", vbTextCompare)

If d <> 0 Then
List2.AddItem List1.List(i)
a = 0
End If
If a <> "0" Then b = Mid(List1.List(i), a + 1, 10)
If a <> "0" Then c = DateDiff("d", b, Now)
If c < 10 Then List2.AddItem List1.List(i)
Next

Form1.List1.Clear
For i = 0 To List2.ListCount - 1
Form1.List1.AddItem List1.List(i)
Next
End If

RMDUPES

End Sub

Public Function PutWindowOnTop(pFrm As Form)
  Dim lngWindowPosition As Long
  lngWindowPosition = SetWindowPos(pFrm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Function
