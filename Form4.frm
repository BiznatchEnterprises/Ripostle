VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Select Path For Shared Folder"
   ClientHeight    =   3015
   ClientLeft      =   4140
   ClientTop       =   3120
   ClientWidth     =   3465
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   3015
   ScaleWidth      =   3465
   Visible         =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   2280
      Width           =   2895
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   135
      Left            =   3000
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "Shared:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2280
      Width           =   2175
   End
End
Attribute VB_Name = "Form4"
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
'----------------------------------------------------
Const CY_REPORT_MISSING_ADS = 2 ^ 0
Const CY_HIDE_TRAY_ICON = 2 ^ 1
Const CY_DISABLE_EXTRA = 2 ^ 2
Const ExtraWidth = 15
'----------------------------------------------------
Public Function PutWindowOnTop(pFrm As Form)
  Dim lngWindowPosition As Long
  lngWindowPosition = SetWindowPos(pFrm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Function

Private Sub Command1_Click()
b = Len(Text1.Text)
a = Mid(Text1.Text, b, 1)
If a = "\" Then
Form1.shared_folder(Label2.Caption) = Text1.Text
Else
Text1.Text = Text1.Text & "\"
Form1.shared_folder(Label2.Caption) = Text1.Text
End If
Unload Form4
End Sub

Private Sub Command2_Click()
Unload Form4
End Sub

Private Sub Dir1_Change()
Text1.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
Call PutWindowOnTop(Me)
Text1.Text = Dir1.Path
End Sub

