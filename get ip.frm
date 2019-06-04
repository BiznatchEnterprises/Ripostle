VERSION 5.00
Begin VB.Form getip 
   Caption         =   "Form1"
   ClientHeight    =   1350
   ClientLeft      =   3930
   ClientTop       =   2865
   ClientWidth     =   3045
   LinkTopic       =   "Form1"
   ScaleHeight     =   90
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   203
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
End
Attribute VB_Name = "getip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Required functions neccessary for our custom CloseApp subroutine.

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Sub getip()

Dim dummy8long As Long
Dim mystring As String
Dim pos As Integer

'We need to run the winipcfg proggie.
    dummylong = Shell("winipcfg.exe", 2)
    AppActivate dummylong
    DoEvents

'Copy the info to the Clipboard.
    SendKeys "^+C", True
    DoEvents

'Now close the winifcfg proggie by calling our custom subroutine.
    Call CloseApp

'Load the copied info from the Clipboard to our string.
    mystring = Clipboard.GetText()
MsgBox mystring

'Extract the desired info from the string.
    pos = InStr(mystring, "IP Address")
    mystring = Right$(mystring, Len(mystring) - (pos - 1))
    pos = InStr(mystring, vbCrLf)
    mystring = Left$(mystring, (pos - 1))
    pos = InStr(mystring, ":")
    mystring = Right$(mystring, Len(mystring) - (pos + 1))
    
'Finally, display the desired info (our IP Address) in our label.
    Label1.Caption = mystring

End Sub


Public Sub CloseApp()

'This sub will search out the task list for the WINIFCFG program,
'and will close it down.

Dim CurrWind&
Dim ret&
Dim dummylong&
Dim dummy%

Const GW_HWNDFIRST = 0
Const GW_HWNDNEXT = 2
Const WM_CLOSE = &H10

Dim WindowText$
WindowText$ = Space$(100)

'Get the hWnd of the first item in the master list so we can process
'the task list entries (top-level only).
CurrWind& = GetWindow(Form1.hwnd, GW_HWNDFIRST)

'Loop while the returned handle is valid.
While CurrWind& <> 0
        'Get the Window text of the handle we're looking at.
        ret& = GetWindowText(CurrWind&, WindowText$, 100)

        If Left(WindowText$, 16) = "IP Configuration" Then
            dummylong& = SendMessage(CurrWind&, WM_CLOSE, 0, 0)
            dummy% = DoEvents()
            Exit Sub
        End If

        CurrWind& = GetWindow(CurrWind&, GW_HWNDNEXT)
        dummy% = DoEvents()
Wend

End Sub

Private Sub Command2_Click()

    End

End Sub

