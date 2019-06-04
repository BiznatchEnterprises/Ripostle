VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   2190
   ClientLeft      =   4995
   ClientTop       =   3495
   ClientWidth     =   1815
   LinkTopic       =   "Form3"
   ScaleHeight     =   2190
   ScaleWidth      =   1815
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command9 
      Caption         =   "Send User a MSG"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Move Transfer >"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      ToolTipText     =   "Move file up or down to set the order of Downloading/Uploading"
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Delete Transfer"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      ToolTipText     =   "Delete the selected file from ur Download/Upload box"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cancel Transfer"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      ToolTipText     =   "Cancel so this file can't transfer, Manual Start is required."
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Queue Transfer"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      ToolTipText     =   "Set to Queue, so when no files are currently transfering, this one will automaticly!"
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find Another Source"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      ToolTipText     =   "Find the EXACT same file from Someone Else so you can resume!"
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   195
      Left            =   2280
      TabIndex        =   0
      Top             =   960
      Width           =   75
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop Transfer"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "Stop the selected File from Downloading/Uploading, "
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start Transfer"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Start the selected File to Download/Upload"
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "=Options="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   1560
      X2              =   1560
      Y1              =   0
      Y2              =   720
   End
   Begin VB.Label Label1 
      Caption         =   "X"
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "Form3"
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

Private Sub Form_Load()
PutWindowOnTop Me
End Sub

Private Sub Label1_Click()
Unload Form3
End Sub
