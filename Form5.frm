VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   495
   ClientLeft      =   7875
   ClientTop       =   6000
   ClientWidth     =   1815
   LinkTopic       =   "Form5"
   ScaleHeight     =   495
   ScaleWidth      =   1815
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Clear Uploads List"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Clear all the Uploads in the Uploads Box"
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear Downloads List"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Clear all Downloads in the Download's Box"
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
For i = 0 To 23
Form1.fname(i).Text = ""
Form1.status(i).Text = ""
Form1.usr(i).Text = ""
Form1.done(i).Text = ""
Form1.kbs(i).Text = ""
Next
Unload Form5
End Sub

Private Sub Command2_Click()
For i = 24 To 47
Form1.fname(i).Text = ""
Form1.status(i).Text = ""
Form1.usr(i).Text = ""
Form1.done(i).Text = ""
Form1.kbs(i).Text = ""
Next
Unload Form5
End Sub
