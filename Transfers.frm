VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Ripostle Transfer Control"
   ClientHeight    =   8220
   ClientLeft      =   495
   ClientTop       =   0
   ClientWidth     =   9900
   LinkTopic       =   "Form2"
   ScaleHeight     =   8220
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   Begin MSWinsockLib.Winsock Get7 
      Left            =   8640
      Top             =   7440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   7480
   End
   Begin MSWinsockLib.Winsock Get6 
      Left            =   8160
      Top             =   7440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   7479
   End
   Begin MSWinsockLib.Winsock Get5 
      Left            =   7680
      Top             =   7440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   7478
   End
   Begin MSWinsockLib.Winsock Get4 
      Left            =   7200
      Top             =   7440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   7477
   End
   Begin MSWinsockLib.Winsock Get3 
      Left            =   6720
      Top             =   7440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   7476
   End
   Begin MSWinsockLib.Winsock Get2 
      Left            =   6240
      Top             =   7440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   7475
   End
   Begin MSWinsockLib.Winsock Get1 
      Left            =   5760
      Top             =   7440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   7474
   End
   Begin MSWinsockLib.Winsock Get0 
      Left            =   5280
      Top             =   7440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   7473
   End
   Begin MSWinsockLib.Winsock Send7 
      Left            =   8640
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Send6 
      Left            =   8160
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Send5 
      Left            =   7680
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Send4 
      Left            =   7200
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Send3 
      Left            =   6720
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Send2 
      Left            =   6240
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Send1 
      Left            =   5760
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Send0 
      Left            =   5280
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox HandleWhater 
      Height          =   285
      Index           =   7
      Left            =   9480
      TabIndex        =   134
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox HandleWhater 
      Height          =   285
      Index           =   6
      Left            =   9480
      TabIndex        =   133
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox HandleWhater 
      Height          =   285
      Index           =   5
      Left            =   9480
      TabIndex        =   132
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox HandleWhater 
      Height          =   285
      Index           =   4
      Left            =   9480
      TabIndex        =   131
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox HandleWhater 
      Height          =   285
      Index           =   3
      Left            =   6840
      TabIndex        =   130
      Top             =   5640
      Width           =   375
   End
   Begin VB.TextBox HandleWhater 
      Height          =   285
      Index           =   2
      Left            =   6840
      TabIndex        =   129
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox HandleWhater 
      Height          =   285
      Index           =   1
      Left            =   6960
      TabIndex        =   128
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox HandleWhater 
      Height          =   285
      Index           =   0
      Left            =   6840
      TabIndex        =   127
      Top             =   840
      Width           =   375
   End
   Begin VB.Timer OStat 
      Enabled         =   0   'False
      Index           =   14
      Interval        =   10
      Left            =   9480
      Top             =   6720
   End
   Begin VB.Timer OStat 
      Enabled         =   0   'False
      Index           =   13
      Interval        =   10
      Left            =   9480
      Top             =   6360
   End
   Begin VB.Timer OStat 
      Enabled         =   0   'False
      Index           =   12
      Interval        =   10
      Left            =   9120
      Top             =   6360
   End
   Begin VB.Timer OStat 
      Enabled         =   0   'False
      Index           =   11
      Interval        =   10
      Left            =   8760
      Top             =   6360
   End
   Begin VB.Timer OStat 
      Enabled         =   0   'False
      Index           =   10
      Interval        =   10
      Left            =   8400
      Top             =   6360
   End
   Begin VB.Timer OStat 
      Enabled         =   0   'False
      Index           =   9
      Interval        =   10
      Left            =   8040
      Top             =   6360
   End
   Begin VB.Timer OStat 
      Enabled         =   0   'False
      Index           =   8
      Interval        =   10
      Left            =   7680
      Top             =   6360
   End
   Begin VB.Timer OStat 
      Enabled         =   0   'False
      Index           =   7
      Interval        =   10
      Left            =   7320
      Top             =   6360
   End
   Begin VB.Timer OStat 
      Enabled         =   0   'False
      Index           =   6
      Interval        =   10
      Left            =   6960
      Top             =   6360
   End
   Begin VB.Timer OStat 
      Enabled         =   0   'False
      Index           =   5
      Interval        =   10
      Left            =   6600
      Top             =   6360
   End
   Begin VB.Timer OStat 
      Enabled         =   0   'False
      Index           =   4
      Interval        =   10
      Left            =   6240
      Top             =   6360
   End
   Begin VB.Timer OStat 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   10
      Left            =   5880
      Top             =   6360
   End
   Begin VB.Timer OStat 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   10
      Left            =   5520
      Top             =   6360
   End
   Begin VB.Timer OStat 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   10
      Left            =   5160
      Top             =   6360
   End
   Begin VB.Timer OStat 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   4800
      Top             =   6360
   End
   Begin VB.TextBox HandWhat 
      Height          =   285
      Index           =   7
      Left            =   8760
      TabIndex        =   126
      Top             =   5520
      Width           =   735
   End
   Begin VB.TextBox HandWhat 
      Height          =   285
      Index           =   6
      Left            =   8760
      TabIndex        =   125
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox HandWhat 
      Height          =   285
      Index           =   5
      Left            =   8760
      TabIndex        =   124
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox HandWhat 
      Height          =   285
      Index           =   4
      Left            =   8760
      TabIndex        =   123
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox HandWhat 
      Height          =   285
      Index           =   3
      Left            =   6120
      TabIndex        =   122
      Top             =   5640
      Width           =   735
   End
   Begin VB.TextBox HandWhat 
      Height          =   285
      Index           =   2
      Left            =   6120
      TabIndex        =   121
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox HandWhat 
      Height          =   285
      Index           =   1
      Left            =   6120
      TabIndex        =   120
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox HandWhat 
      Height          =   285
      Index           =   0
      Left            =   6120
      TabIndex        =   119
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   7
      Left            =   7920
      TabIndex        =   114
      Text            =   "c:\"
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Index           =   7
      Interval        =   2000
      Left            =   7440
      Top             =   4800
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   6
      Left            =   7920
      TabIndex        =   109
      Text            =   "c:\"
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Index           =   6
      Interval        =   2000
      Left            =   7440
      Top             =   3360
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   5
      Left            =   7920
      TabIndex        =   104
      Text            =   "c:\"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Index           =   5
      Interval        =   2000
      Left            =   7440
      Top             =   1800
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   4
      Left            =   7920
      TabIndex        =   99
      Text            =   "c:\"
      Top             =   240
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Index           =   4
      Interval        =   2000
      Left            =   7440
      Top             =   120
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   3
      Left            =   5280
      TabIndex        =   94
      Text            =   "c:\"
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Index           =   3
      Interval        =   2000
      Left            =   4800
      Top             =   4920
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   2
      Left            =   5280
      TabIndex        =   89
      Text            =   "c:\"
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Index           =   2
      Interval        =   2000
      Left            =   4800
      Top             =   3360
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   1
      Left            =   5280
      TabIndex        =   84
      Text            =   "c:\"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Index           =   1
      Interval        =   2000
      Left            =   4800
      Top             =   1800
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   0
      Left            =   5280
      TabIndex        =   79
      Text            =   "c:\"
      Top             =   240
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Index           =   0
      Interval        =   2000
      Left            =   4800
      Top             =   120
   End
   Begin VB.TextBox HandleWhat 
      Height          =   285
      Index           =   10
      Left            =   4200
      TabIndex        =   78
      Top             =   8640
      Width           =   375
   End
   Begin VB.TextBox port0 
      Height          =   285
      Index           =   10
      Left            =   3720
      TabIndex        =   77
      Top             =   8640
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   10
      Left            =   2400
      TabIndex        =   72
      Top             =   9000
      Width           =   1335
   End
   Begin VB.TextBox HandleWhat 
      Height          =   285
      Index           =   8
      Left            =   4200
      TabIndex        =   71
      Top             =   5280
      Width           =   375
   End
   Begin VB.TextBox port0 
      Height          =   285
      Index           =   8
      Left            =   3720
      TabIndex        =   70
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   3240
      TabIndex        =   65
      Text            =   "127.0.0.1"
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   8
      Left            =   2400
      TabIndex        =   64
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   8
      Interval        =   2000
      Left            =   2400
      Top             =   4920
   End
   Begin VB.TextBox HandleWhat 
      Height          =   285
      Index           =   7
      Left            =   4200
      TabIndex        =   63
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox port0 
      Height          =   285
      Index           =   7
      Left            =   3720
      TabIndex        =   62
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   3240
      TabIndex        =   57
      Text            =   "127.0.0.1"
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   7
      Left            =   2400
      TabIndex        =   56
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Index           =   7
      Interval        =   2000
      Left            =   2400
      Top             =   3360
   End
   Begin VB.TextBox HandleWhat 
      Height          =   285
      Index           =   6
      Left            =   4200
      TabIndex        =   55
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox port0 
      Height          =   285
      Index           =   6
      Left            =   3720
      TabIndex        =   54
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   3240
      TabIndex        =   49
      Text            =   "127.0.0.1"
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   6
      Left            =   2400
      TabIndex        =   48
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Index           =   6
      Interval        =   2000
      Left            =   2400
      Top             =   1680
   End
   Begin VB.TextBox HandleWhat 
      Height          =   285
      Index           =   5
      Left            =   4200
      TabIndex        =   47
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox port0 
      Height          =   285
      Index           =   5
      Left            =   3720
      TabIndex        =   46
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   3240
      TabIndex        =   41
      Text            =   "127.0.0.1"
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   5
      Left            =   2400
      TabIndex        =   40
      Top             =   720
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Index           =   5
      Interval        =   2000
      Left            =   2400
      Top             =   0
   End
   Begin VB.TextBox HandleWhat 
      Height          =   285
      Index           =   4
      Left            =   1800
      TabIndex        =   39
      Top             =   6960
      Width           =   375
   End
   Begin VB.TextBox port0 
      Height          =   285
      Index           =   4
      Left            =   1320
      TabIndex        =   38
      Top             =   6960
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   840
      TabIndex        =   33
      Text            =   "127.0.0.1"
      Top             =   6600
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   4
      Left            =   0
      TabIndex        =   32
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Index           =   4
      Interval        =   2000
      Left            =   0
      Top             =   6600
   End
   Begin VB.TextBox HandleWhat 
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   31
      Top             =   5280
      Width           =   375
   End
   Begin VB.TextBox port0 
      Height          =   285
      Index           =   3
      Left            =   1320
      TabIndex        =   30
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   840
      TabIndex        =   25
      Text            =   "127.0.0.1"
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   3
      Left            =   0
      TabIndex        =   24
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Index           =   3
      Interval        =   2000
      Left            =   0
      Top             =   4920
   End
   Begin VB.TextBox HandleWhat 
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   23
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox port0 
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   22
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   840
      TabIndex        =   17
      Text            =   "127.0.0.1"
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   2
      Left            =   0
      TabIndex        =   16
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Index           =   2
      Interval        =   2000
      Left            =   0
      Top             =   3360
   End
   Begin VB.TextBox HandleWhat 
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   15
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox port0 
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   14
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   840
      TabIndex        =   9
      Text            =   "127.0.0.1"
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Index           =   1
      Interval        =   2000
      Left            =   0
      Top             =   1680
   End
   Begin VB.TextBox HandleWhat 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   7
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox port0 
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   6
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Index           =   0
      Interval        =   2000
      Left            =   0
      Top             =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   4680
      X2              =   4680
      Y1              =   120
      Y2              =   8280
   End
   Begin VB.Label Label3 
      Caption         =   "Speed:"
      Height          =   255
      Index           =   7
      Left            =   7560
      TabIndex        =   118
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Done:"
      Height          =   255
      Index           =   7
      Left            =   7560
      TabIndex        =   117
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Not Connected"
      Height          =   255
      Index           =   7
      Left            =   7920
      TabIndex        =   116
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Size:"
      Height          =   255
      Index           =   7
      Left            =   7560
      TabIndex        =   115
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Speed:"
      Height          =   255
      Index           =   6
      Left            =   7560
      TabIndex        =   113
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Done:"
      Height          =   255
      Index           =   6
      Left            =   7560
      TabIndex        =   112
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Not Connected"
      Height          =   255
      Index           =   6
      Left            =   7920
      TabIndex        =   111
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Size:"
      Height          =   255
      Index           =   6
      Left            =   7560
      TabIndex        =   110
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Speed:"
      Height          =   255
      Index           =   5
      Left            =   7560
      TabIndex        =   108
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Done:"
      Height          =   255
      Index           =   5
      Left            =   7560
      TabIndex        =   107
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Not Connected"
      Height          =   255
      Index           =   5
      Left            =   7920
      TabIndex        =   106
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Size:"
      Height          =   255
      Index           =   5
      Left            =   7560
      TabIndex        =   105
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Speed:"
      Height          =   255
      Index           =   4
      Left            =   7560
      TabIndex        =   103
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Done:"
      Height          =   255
      Index           =   4
      Left            =   7560
      TabIndex        =   102
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Not Connected"
      Height          =   255
      Index           =   4
      Left            =   7920
      TabIndex        =   101
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Size:"
      Height          =   255
      Index           =   4
      Left            =   7560
      TabIndex        =   100
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Speed:"
      Height          =   255
      Index           =   3
      Left            =   4920
      TabIndex        =   98
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Done:"
      Height          =   255
      Index           =   3
      Left            =   4920
      TabIndex        =   97
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Not Connected"
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   96
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Size:"
      Height          =   255
      Index           =   3
      Left            =   4920
      TabIndex        =   95
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Speed:"
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   93
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Done:"
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   92
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Not Connected"
      Height          =   255
      Index           =   2
      Left            =   5280
      TabIndex        =   91
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Size:"
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   90
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Speed:"
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   88
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Done:"
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   87
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Not Connected"
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   86
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Size:"
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   85
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Speed:"
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   83
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Done:"
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   82
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Not Connected"
      Height          =   255
      Index           =   0
      Left            =   5280
      TabIndex        =   81
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Size:"
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   80
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Speed:"
      Height          =   255
      Index           =   10
      Left            =   3360
      TabIndex        =   76
      Top             =   9360
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Done:"
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   75
      Top             =   9600
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Not Connected"
      Height          =   255
      Index           =   10
      Left            =   2400
      TabIndex        =   74
      Top             =   8760
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Size:"
      Height          =   255
      Index           =   10
      Left            =   2400
      TabIndex        =   73
      Top             =   9360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Speed:"
      Height          =   255
      Index           =   8
      Left            =   3360
      TabIndex        =   69
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Done:"
      Height          =   255
      Index           =   8
      Left            =   2760
      TabIndex        =   68
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Not Connected"
      Height          =   255
      Index           =   8
      Left            =   2400
      TabIndex        =   67
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Size:"
      Height          =   255
      Index           =   8
      Left            =   2400
      TabIndex        =   66
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Speed:"
      Height          =   255
      Index           =   7
      Left            =   3360
      TabIndex        =   61
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Done:"
      Height          =   255
      Index           =   7
      Left            =   2760
      TabIndex        =   60
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Not Connected"
      Height          =   255
      Index           =   7
      Left            =   2400
      TabIndex        =   59
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Size:"
      Height          =   255
      Index           =   7
      Left            =   2400
      TabIndex        =   58
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Speed:"
      Height          =   255
      Index           =   6
      Left            =   3360
      TabIndex        =   53
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Done:"
      Height          =   255
      Index           =   6
      Left            =   2760
      TabIndex        =   52
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Not Connected"
      Height          =   255
      Index           =   6
      Left            =   2400
      TabIndex        =   51
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Size:"
      Height          =   255
      Index           =   6
      Left            =   2400
      TabIndex        =   50
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Speed:"
      Height          =   255
      Index           =   5
      Left            =   3360
      TabIndex        =   45
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Done:"
      Height          =   255
      Index           =   5
      Left            =   2760
      TabIndex        =   44
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Not Connected"
      Height          =   255
      Index           =   5
      Left            =   2400
      TabIndex        =   43
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Size:"
      Height          =   255
      Index           =   5
      Left            =   2400
      TabIndex        =   42
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Speed:"
      Height          =   255
      Index           =   4
      Left            =   960
      TabIndex        =   37
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Done:"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   36
      Top             =   7920
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Not Connected"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   35
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Size:"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   34
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Speed:"
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   29
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Done:"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   28
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Not Connected"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   27
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Size:"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   26
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Speed:"
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   21
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Done:"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   20
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Not Connected"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   19
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Size:"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   18
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Speed:"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   13
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Done:"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   12
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Not Connected"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Size:"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   10
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Speed:"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Done:"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Not Connected"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Size:"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim givemore(7) As String
Dim currentint(7) As Long
Dim currentin(7) As Long
Dim ratee(7) As Integer
Dim ratey(7) As Integer
Dim filesizeu(7) As Long
Dim FileNumber1(7)
Dim FileNumber2(7)
Dim GotBytes(7)


Sub Connecter(ByVal SocketNumber As String)
Winsock1(SocketNumber).Connect Text1(SocketNumber).Text, port0(SocketNumber).Text
Label5(SocketNumber).Caption = "Connecting"
Label7(SocketNumber).Caption = FileLen(Text3(SocketNumber).Text)
End Sub

Sub Listener(ByVal SocketNumber As String)
Winsock2(SocketNumber).Close
Winsock2(SocketNumber).Listen
Label6(SocketNumber).Caption = "Listening"
End Sub

Sub Sleep(ByVal MillaSec As Long, Optional ByVal DeepSleep As Boolean = False)
    Dim tStart#, Tmr#
    tStart = Timer

    While Tmr < (MillaSec / 1000)
        Tmr = Timer - tStart
        If DeepSleep = False Then DoEvents
    Wend
End Sub

Private Sub Command1_Click()
port0(0).Text = "7473"
Listener 0
Sleep 3000
Connecter 0
End Sub

Private Sub Form_Load()
ajax = 1
ajox = 10
For i = 0 To 7
FileNumber1(i) = ajax
FileNumber2(i) = ajox
ajax = ajax + 1
ajox = ajox + 1
Next
End Sub

Private Sub Timer1_Timer(Index As Integer)
Label1(Index).Caption = (ratee(Index) / 2)
ratee(Index) = 0
End Sub

Private Sub Timer2_Timer(Index As Integer)
Label3(Index).Caption = (ratey(Index) / 2)
ratey(Index) = 0
End Sub

Private Sub Winsock1_Close(Index As Integer)

If Label2(Index).Caption > 99 Then
Label5(Index).Caption = "Done!"
Else
Label5(Index).Caption = "Canceled"
End If
End Sub

Private Sub Winsock1_Connect(Index As Integer)
Label5(Index).Caption = "Connected"
Winsock1(Index).SendData "FILSZ " & Label7(Index).Caption
Sleep 2000
Winsock1(Index).SendData "FILNM " & Form1.fname(HandleWhat(Index)).Text
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim dart As String
Winsock1(Index).GetData dart

If InStr(1, dart, "GTB", vbTextCompare) <> 0 Then
Label5(Index).Caption = "Sending File"
givemore(Index) = 1
SendTheFile Index
End If

If InStr(1, dart, "RSM", vbTextCompare) <> 0 Then
a = Mid(dart, 4, Len(dart) - 3)
currentint(Index) = a
Label5(Index).Caption = "Sending File"
givemore(Index) = 1
SendRSMFile (Index)
End If

If dart = "GIVEMORE" Then givemore(Index) = 1
End Sub
Sub SendRSMFile(ByVal SocketNumbr As Integer)
Dim tempbuffer As String

Open Text3(SocketNumbr).Text For Binary Access Read As #FileNumber1(SocketNumbr)
Label2(SocketNumbr).Caption = 0

Do Until EOF(FileNumber1(SocketNumbr))
Do Until givemore(SocketNumbr) = 1
DoEvents
Loop
tempbuffer = Space$(1024)

Get #FileNumber1(SocketNumbr), currentint(SocketNumbr), tempbuffer


currentint(SocketNumbr) = currentint(SocketNumbr) + 1024
filesizeu(Index) = Label7(SocketNumbr).Caption
Label2(SocketNumbr).Caption = Int((currentint(SocketNumbr) / filesizeu(SocketNumbr)) * 100)
ratee(SocketNumbr) = ratee(SocketNumbr) + 1

Winsock1(SocketNumbr).SendData "NMDATA" & tempbuffer
givemore(SocketNumbr) = 0
Loop
Sleep 4000
Close #FileNumber1(SocketNumbr)
Winsock1(SocketNumbr).Close
Label5(SocketNumbr).Caption = "Done!"
'Winsock1(SocketNumbr).Close
Exit Sub
End Sub

Sub SendTheFile(ByVal SocketNumbr As Long)
Dim tempbuffer As String
'On Error GoTo errhand
FileNumber3 = FreeFile
Open Text3(SocketNumbr).Text For Binary Access Read As #FileNumber1(SocketNumbr)

Label2(SocketNumbr).Caption = 0
currentint(SocketNumbr) = 0

Do Until EOF(FileNumber1(SocketNumbr))
Do Until givemore(SocketNumbr) = 1
DoEvents
Loop
tempbuffer = Space$(1024)

Get #FileNumber1(SocketNumbr), , tempbuffer
currentint(SocketNumbr) = currentint(SocketNumbr) + 1024
filesizeu(SocketNumbr) = Label7(SocketNumbr).Caption
Label2(SocketNumbr).Caption = Int((currentint(SocketNumbr) / filesizeu(SocketNumbr)) * 100)
ratee(SocketNumbr) = ratee(SocketNumbr) + 1

Winsock1(SocketNumbr).SendData "NMDATA" & tempbuffer
givemore(SocketNumbr) = 0
Loop
Sleep 4000
Close #FileNumber1(SocketNumbr)
Winsock1(SocketNumbr).Close
Label5(SocketNumbr).Caption = "Done!"
errhand:
Winsock1(SocketNumbr).Close
Exit Sub
End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Label5(Index).Caption = "Canceled"
Winsock1(Index).Close
End Sub

Private Sub Winsock2_Close(Index As Integer)
a = Label4(Index).Caption
If a > 99 Then
Label6(Index).Caption = "Done!"
Close #FileNumber2(Index)
Else
Label6(Index).Caption = "Canceled"
Close #FileNumber2(Index)
End If
End Sub

Private Sub Winsock2_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Winsock2(Index).Close
Winsock2(Index).Accept requestID
Label6(Index).Caption = "Connected"
End Sub

Private Sub Winsock2_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim dart As String
Dim jki As Long
Dim jku As Long
Winsock2(Index).GetData dart
If InStr(1, dart, "FILSZ", vbTextCompare) <> 0 Then
a = Mid(dart, 7, Len(dart) - 6)
Label8(Index).Caption = a
End If

If InStr(1, dart, "FILNM", vbTextCompare) <> 0 Then
a = Mid(dart, 7, Len(dart) - 6)
Text2(Index).Text = Text2(Index).Text & a
currentin(Index) = 0
On Error GoTo GetFileFstTime
b = FileLen(Text2(Index).Text) 'resume
currentin(Index) = b
If b < 1 Then b = 1
Winsock2(Index).SendData "RSM" & b
Open Text2(Index).Text For Binary Access Write As #FileNumber2(Index)
GotBytes(Index) = currentin(Index)
Exit Sub

GetFileFstTime: 'the first time getting the file
Winsock2(Index).SendData "GTB"
Label4(Index).Caption = 0
Open Text2(Index).Text For Binary Access Write As #FileNumber2(Index)
Label6(Index).Caption = "Getting File"
GotBytes(Index) = "0"
Exit Sub
End If
'---------normal download
If InStr(1, dart, "NMDATA") <> 0 Then
sendsze = 1024
dart = Mid(dart, 7, sendsze + 1)
If currentin(Index) = 0 Then currentin(Index) = 1
GotBytes(Index) = GotBytes(Index) + Len(dart)
jki = GotBytes(Index)
jku = Label8(Index).Caption
If jki > jku Then
jop = jki - jku
joh = 1024 = jop
hobo = Mid(dart, 1, joh)
dart = hobo
End If


Put #FileNumber2(Index), currentin(Index), dart
currentin(Index) = currentin(Index) + sendsze
filesizeo = Label8(Index).Caption
hg = Int((currentin(Index) / filesizeo) * 100)
Label4(Index).Caption = hg
ratey(Index) = ratey(Index) + 1
If hg > 99 Then
Label6(Index).Caption = "Done!"
Winsock2(Index).Close
Exit Sub
End If
Winsock2(Index).SendData "GIVEMORE"
End If

'---------resume----
If InStr(1, dart, "RSDATA") <> 0 Then
sendsze = 1024
dart = Mid(dart, 7, sendsze + 1)

Label4(Index).Caption = hg
GotBytes(Index) = GotBytes(Index) + Len(dart)
jki = GotBytes(Index)
jku = Label8(Index).Caption
If jki > jku Then
jop = jki - jku
joh = 1024 = jop
hobo = Mid(dart, 1, joh)
dart = hobo
End If

Put #FileNumber2(Index), currentin(Index), dart
Winsock2(Index).SendData "GIVEMORE"
ratey(Index) = ratey(Index) + 1
currentin(Index) = currentin(Index) + sendsze
dart = Mid(dart, 7, sendsze + 1)
filesizeo = Label8(Index).Caption
hg = Int((currentin(Index) / filesizeo) * 100)
If hg > 99 Then
Label6(Index).Caption = "Done!"
Close #FileNumber2(Index)
Winsock2(Index).Close
End If
End If
End Sub

Private Sub Winsock2_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Label6(Index).Caption = "Canceled"
Winsock2(Index).Close
End Sub
