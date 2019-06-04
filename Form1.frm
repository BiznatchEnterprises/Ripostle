VERSION 5.00
Object = "{60CC5D62-2D08-11D0-BDBE-00AA00575603}#1.0#0"; "dcsystray.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{B79097AE-F549-11D3-AFEB-8F2913C35F06}#2.0#0"; "FT.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00A3B6BE&
   BorderStyle     =   0  'None
   Caption         =   "Ripostle Beta 1.0"
   ClientHeight    =   9000
   ClientLeft      =   1875
   ClientTop       =   0
   ClientWidth     =   7755
   DrawStyle       =   5  'Transparent
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":08CA
   MousePointer    =   99  'Custom
   Picture         =   "Form1.frx":0A1C
   ScaleHeight     =   9000
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   Begin VB.Timer UL_KBS_TMR 
      Enabled         =   0   'False
      Index           =   24
      Interval        =   1000
      Left            =   2040
      Top             =   840
   End
   Begin VB.Timer DL_KBS_TMR 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1000
      Left            =   1560
      Top             =   840
   End
   Begin VB.Timer iplistcleanup 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   360
      Top             =   3960
   End
   Begin VB.Timer Downloader_Timout 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   5000
      Left            =   7080
      Top             =   3240
   End
   Begin VB.Timer UPL_Transfer_Bot 
      Interval        =   3000
      Left            =   240
      Top             =   3240
   End
   Begin MSWinsockLib.Winsock Uploader 
      Index           =   0
      Left            =   6960
      Tag             =   "vb"
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   7472
   End
   Begin MSWinsockLib.Winsock Downloader 
      Index           =   0
      Left            =   6960
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   7472
   End
   Begin VB.Timer DL_Transfer_Bot 
      Interval        =   2000
      Left            =   240
      Top             =   2400
   End
   Begin VB.Frame Frame6 
      Caption         =   "Frame6"
      Height          =   375
      Left            =   360
      TabIndex        =   355
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
      Begin SHDocVwCtl.WebBrowser ads2 
         Height          =   135
         Left            =   120
         TabIndex        =   356
         Top             =   120
         Width           =   30
         ExtentX         =   53
         ExtentY         =   238
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00A3B6BE&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   930
      Left            =   360
      TabIndex        =   353
      Top             =   1200
      Width           =   6975
      Begin SHDocVwCtl.WebBrowser ads 
         Height          =   2055
         Left            =   0
         TabIndex        =   354
         Top             =   0
         Width           =   7815
         ExtentX         =   13785
         ExtentY         =   3625
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.TextBox zort 
      Height          =   285
      Left            =   240
      TabIndex        =   94
      Text            =   "0"
      Top             =   840
      Width           =   615
   End
   Begin VB.Timer search_tooltip_tmr 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   480
      Top             =   4440
   End
   Begin VB.Timer Update_MFlist 
      Interval        =   61000
      Left            =   5040
      Top             =   120
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   4920
      TabIndex        =   26
      Top             =   480
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Timer cnt4listsharetimout 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   480
      Top             =   1320
   End
   Begin VB.TextBox sharelistcnt 
      Height          =   285
      Left            =   120
      TabIndex        =   25
      Text            =   "3"
      Top             =   7320
      Width           =   495
   End
   Begin VB.Timer sharelist 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   0
      Top             =   6840
   End
   Begin VB.Timer usercount 
      Interval        =   100
      Left            =   840
      Top             =   8400
   End
   Begin VB.ListBox List1 
      Height          =   450
      ItemData        =   "Form1.frx":E5ECE
      Left            =   360
      List            =   "Form1.frx":E5ED8
      TabIndex        =   24
      Top             =   240
      Width           =   2175
   End
   Begin MSWinsockLib.Winsock foot 
      Left            =   360
      Top             =   5760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   7470
   End
   Begin MSWinsockLib.Winsock sock 
      Left            =   600
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   7470
   End
   Begin FormTracer.FT FT1 
      Left            =   360
      Top             =   1920
      _ExtentX        =   423
      _ExtentY        =   397
      RegKey          =   "125962351"
   End
   Begin SysTrayCtl.cSysTray cSysTray1 
      Left            =   360
      Top             =   2760
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   -1  'True
      TrayIcon        =   "Form1.frx":E5EF9
      TrayTip         =   "Ripostle Beta 1.0"
   End
   Begin VB.Frame Myfiles_window 
      BackColor       =   &H00A3B6BE&
      Height          =   4815
      Left            =   1080
      TabIndex        =   11
      Top             =   2040
      Visible         =   0   'False
      Width           =   5655
      Begin VB.Frame myfiles_frame3 
         BackColor       =   &H00A3B6BE&
         Height          =   6855
         Left            =   0
         TabIndex        =   59
         Top             =   5400
         Width           =   5415
         Begin VB.CommandButton Command14 
            Caption         =   "Open"
            Height          =   255
            Left            =   4320
            TabIndex        =   90
            Top             =   6120
            Width           =   615
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Open"
            Height          =   255
            Left            =   4320
            TabIndex        =   89
            Top             =   3960
            Width           =   615
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Play"
            Height          =   255
            Left            =   4320
            TabIndex        =   84
            Top             =   1920
            Width           =   615
         End
         Begin VB.FileListBox My_files_Shm 
            Height          =   1455
            Left            =   120
            Pattern         =   "*[Sheet Music].txt"
            TabIndex        =   63
            Top             =   4920
            Width           =   5175
         End
         Begin VB.FileListBox My_files_Gtabs 
            Height          =   1455
            Left            =   120
            Pattern         =   "*[Gtab].txt"
            TabIndex        =   62
            Top             =   2760
            Width           =   5175
         End
         Begin VB.FileListBox My_files_Midi 
            Height          =   1455
            Left            =   120
            Pattern         =   "*.mid"
            TabIndex        =   61
            Top             =   720
            Width           =   5175
         End
         Begin VB.Frame Frame4 
            Caption         =   "Frame1"
            Height          =   15
            Left            =   0
            TabIndex        =   60
            Top             =   8760
            Width           =   5415
         End
         Begin VB.Label total_shm 
            BackColor       =   &H00A3B6BE&
            Caption         =   "Total:"
            Height          =   255
            Left            =   4200
            TabIndex        =   81
            Top             =   4680
            Width           =   855
         End
         Begin VB.Label total_gtab 
            BackColor       =   &H00A3B6BE&
            Caption         =   "Total:"
            Height          =   255
            Left            =   4200
            TabIndex        =   80
            Top             =   2520
            Width           =   855
         End
         Begin VB.Label total_midi 
            BackColor       =   &H00A3B6BE&
            Caption         =   "Total:"
            Height          =   255
            Left            =   4320
            TabIndex        =   79
            Top             =   360
            Width           =   855
         End
         Begin VB.Line Line8 
            X1              =   0
            X2              =   5400
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Label Label61 
            BackColor       =   &H00A3B6BE&
            Height          =   255
            Left            =   1200
            TabIndex        =   73
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label58 
            BackColor       =   &H00A3B6BE&
            Caption         =   "File Size MB:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Line Line9 
            X1              =   0
            X2              =   5400
            Y1              =   4560
            Y2              =   4560
         End
         Begin VB.Label Label62 
            BackColor       =   &H00A3B6BE&
            Height          =   255
            Left            =   3360
            TabIndex        =   71
            Top             =   4320
            Width           =   1335
         End
         Begin VB.Label Label60 
            BackColor       =   &H00A3B6BE&
            Height          =   255
            Left            =   3240
            TabIndex        =   70
            Top             =   6480
            Width           =   1335
         End
         Begin VB.Label Label59 
            BackColor       =   &H00A3B6BE&
            Caption         =   "File Size MB:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   69
            Top             =   4320
            Width           =   1575
         End
         Begin VB.Label Label57 
            Alignment       =   2  'Center
            BackColor       =   &H00A3B6BE&
            Caption         =   "Guitar Tabs: [Gtab]"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1920
            TabIndex        =   68
            Top             =   2520
            Width           =   1575
         End
         Begin VB.Line Line10 
            X1              =   0
            X2              =   5400
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Label Label56 
            Alignment       =   2  'Center
            BackColor       =   &H00A3B6BE&
            Caption         =   "Sheet Music [SheetMusic]"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1800
            TabIndex        =   67
            Top             =   4680
            Width           =   1935
         End
         Begin VB.Label Label55 
            BackColor       =   &H00A3B6BE&
            Caption         =   "File Size MB:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   66
            Top             =   6480
            Width           =   1575
         End
         Begin VB.Label Label54 
            BackColor       =   &H00A3B6BE&
            Height          =   255
            Left            =   3000
            TabIndex        =   65
            Top             =   4560
            Width           =   735
         End
         Begin VB.Label Label53 
            Alignment       =   2  'Center
            BackColor       =   &H00A3B6BE&
            Caption         =   "Midi: [.Mid]"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1920
            TabIndex        =   64
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label51 
            BackColor       =   &H00A3B6BE&
            Caption         =   "Length (Mins):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   74
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label Label52 
            BackColor       =   &H00A3B6BE&
            Height          =   255
            Left            =   4080
            TabIndex        =   75
            Top             =   2160
            Width           =   330
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   4575
         LargeChange     =   50
         Left            =   5400
         Max             =   9999
         SmallChange     =   30
         TabIndex        =   49
         Top             =   120
         Width           =   255
      End
      Begin VB.Frame myfiles_frame2 
         BackColor       =   &H00A3B6BE&
         Height          =   7575
         Left            =   0
         TabIndex        =   39
         Top             =   -120
         Width           =   5415
         Begin VB.CommandButton Command12 
            Caption         =   "Play"
            Height          =   255
            Left            =   4320
            TabIndex        =   87
            Top             =   2640
            Width           =   615
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Open"
            Height          =   255
            Left            =   4320
            TabIndex        =   86
            Top             =   4800
            Width           =   615
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Open"
            Height          =   255
            Left            =   4320
            TabIndex        =   85
            Top             =   6840
            Width           =   615
         End
         Begin VB.Frame Frame1 
            Caption         =   "Frame1"
            Height          =   15
            Left            =   0
            TabIndex        =   58
            Top             =   8760
            Width           =   5415
         End
         Begin VB.FileListBox My_files_Dtabs 
            Height          =   1455
            Left            =   120
            Pattern         =   "*[Dtab].txt"
            TabIndex        =   55
            Top             =   5640
            Width           =   5175
         End
         Begin VB.FileListBox My_files_Mp3z 
            Height          =   2040
            Left            =   120
            Pattern         =   "*.mp3"
            TabIndex        =   40
            Top             =   840
            Width           =   5175
         End
         Begin VB.FileListBox My_files_Lyrics 
            Height          =   1455
            Left            =   120
            Pattern         =   "*[lyric].txt"
            TabIndex        =   51
            Top             =   3600
            Width           =   5175
         End
         Begin VB.Label Label42 
            BackColor       =   &H00A3B6BE&
            Height          =   255
            Left            =   3120
            TabIndex        =   88
            Top             =   7200
            Width           =   1455
         End
         Begin VB.Line Line5 
            X1              =   0
            X2              =   5400
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label total_mp3z 
            BackColor       =   &H00A3B6BE&
            Caption         =   "Total:"
            Height          =   255
            Left            =   4320
            TabIndex        =   76
            Top             =   600
            Width           =   855
         End
         Begin VB.Line Line7 
            X1              =   0
            X2              =   5400
            Y1              =   5400
            Y2              =   5400
         End
         Begin VB.Label Label49 
            BackColor       =   &H00A3B6BE&
            Caption         =   "File Size MB:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   56
            Top             =   7200
            Width           =   1215
         End
         Begin VB.Label Label50 
            BackColor       =   &H00A3B6BE&
            Height          =   255
            Left            =   3120
            TabIndex        =   57
            Top             =   7920
            Width           =   615
         End
         Begin VB.Label Label48 
            Alignment       =   2  'Center
            BackColor       =   &H00A3B6BE&
            Caption         =   "Dumb Tabs: [Dtab]"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1800
            TabIndex        =   54
            Top             =   5400
            Width           =   1575
         End
         Begin VB.Label Label47 
            BackColor       =   &H00A3B6BE&
            Height          =   255
            Left            =   3000
            TabIndex        =   53
            Top             =   5160
            Width           =   1575
         End
         Begin VB.Label Label46 
            BackColor       =   &H00A3B6BE&
            Caption         =   "File Size MB:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   52
            Top             =   5160
            Width           =   1575
         End
         Begin VB.Line Line6 
            X1              =   0
            X2              =   5400
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Label Label45 
            Alignment       =   2  'Center
            BackColor       =   &H00A3B6BE&
            Caption         =   "Lyrics [Lyric]:"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1920
            TabIndex        =   50
            Top             =   3360
            Width           =   1575
         End
         Begin VB.Label Label44 
            Alignment       =   2  'Center
            BackColor       =   &H00A3B6BE&
            Caption         =   "MP3z:"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1920
            TabIndex        =   48
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label25 
            BackColor       =   &H00A3B6BE&
            Caption         =   "My Shared Files"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   47
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label28 
            BackColor       =   &H00A3B6BE&
            Caption         =   "File Size MB:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   3000
            Width           =   1215
         End
         Begin VB.Label length_mf 
            BackColor       =   &H00A3B6BE&
            Height          =   255
            Left            =   3480
            TabIndex        =   43
            Top             =   3000
            Width           =   615
         End
         Begin VB.Label Label30 
            BackColor       =   &H00A3B6BE&
            Caption         =   "Bitrate:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4320
            TabIndex        =   42
            Top             =   3000
            Width           =   615
         End
         Begin VB.Label bitrate_mf 
            BackColor       =   &H00A3B6BE&
            Height          =   255
            Left            =   5040
            TabIndex        =   41
            Top             =   3000
            Width           =   270
         End
         Begin VB.Label total_lyrics 
            BackColor       =   &H00A3B6BE&
            Caption         =   "Total:"
            Height          =   255
            Left            =   4320
            TabIndex        =   77
            Top             =   3360
            Width           =   855
         End
         Begin VB.Label total_drumtabs 
            BackColor       =   &H00A3B6BE&
            Caption         =   "Total:"
            Height          =   255
            Left            =   4320
            TabIndex        =   78
            Top             =   5400
            Width           =   855
         End
         Begin VB.Label Label29 
            BackColor       =   &H00A3B6BE&
            Caption         =   "Length (Mins):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   44
            Top             =   3000
            Width           =   1335
         End
         Begin VB.Label filsize_mf 
            BackColor       =   &H00A3B6BE&
            Height          =   255
            Left            =   1320
            TabIndex        =   45
            Top             =   3000
            Width           =   1455
         End
      End
   End
   Begin VB.Frame information_window 
      BackColor       =   &H00A3B6BE&
      Height          =   4815
      Left            =   1080
      TabIndex        =   14
      Top             =   2040
      Width           =   5655
      Begin VB.Timer scrolltext1 
         Interval        =   100
         Left            =   120
         Top             =   240
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00A3B6BE&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   5295
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00A3B6BE&
            Caption         =   $"Form1.frx":E67D3
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   5280
            TabIndex        =   17
            Top             =   0
            Width           =   12120
         End
      End
      Begin VB.Label Label9 
         BackColor       =   &H00A3B6BE&
         Caption         =   "Features in this version:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   365
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label31 
         BackColor       =   &H00A3B6BE&
         Caption         =   $"Form1.frx":E688E
         Height          =   1815
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   5415
      End
      Begin VB.Label Label26 
         BackColor       =   &H00A3B6BE&
         Caption         =   "Information:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label32 
         BackColor       =   &H00A3B6BE&
         Caption         =   $"Form1.frx":E6B31
         ForeColor       =   &H000000FF&
         Height          =   1575
         Left            =   120
         TabIndex        =   19
         Top             =   2880
         Width           =   5415
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackColor       =   &H00A3B6BE&
         Caption         =   "To Get Started Click ""Pref"" to Define your settings, or click ""Etc"" For Help"
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   4440
         Width           =   5415
      End
   End
   Begin VB.Frame pref_window 
      BackColor       =   &H00A3B6BE&
      Height          =   4935
      Left            =   1080
      TabIndex        =   27
      Top             =   2040
      Visible         =   0   'False
      Width           =   5655
      Begin VB.OptionButton Option7 
         BackColor       =   &H00A3B6BE&
         Caption         =   "T1"
         Height          =   255
         Left            =   5040
         TabIndex        =   374
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox ModemSpeed 
         Height          =   285
         Left            =   4920
         TabIndex        =   373
         Text            =   "Unknown"
         Top             =   2760
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00A3B6BE&
         Caption         =   "Cable"
         Height          =   255
         Left            =   4320
         TabIndex        =   372
         Top             =   3120
         Width           =   735
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00A3B6BE&
         Caption         =   "DSL"
         Height          =   255
         Left            =   3720
         TabIndex        =   371
         Top             =   3120
         Width           =   735
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00A3B6BE&
         Caption         =   "ISDN"
         Height          =   255
         Left            =   3000
         TabIndex        =   370
         Top             =   3120
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00A3B6BE&
         Caption         =   "56k"
         Height          =   255
         Left            =   2400
         TabIndex        =   369
         Top             =   3120
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00A3B6BE&
         Caption         =   "36.6"
         Height          =   255
         Left            =   1800
         TabIndex        =   368
         Top             =   3120
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00A3B6BE&
         Caption         =   "28.8"
         Height          =   255
         Left            =   1200
         TabIndex        =   367
         ToolTipText     =   "Select Your Modem speed"
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox max_ups 
         Height          =   285
         Left            =   4800
         TabIndex        =   361
         Text            =   "1"
         Top             =   3360
         Width           =   495
      End
      Begin VB.TextBox max_dls 
         Height          =   315
         Left            =   2160
         TabIndex        =   360
         Text            =   "1"
         Top             =   3360
         Width           =   495
      End
      Begin VB.TextBox username 
         Height          =   285
         Left            =   2520
         TabIndex        =   358
         Text            =   "Guest"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox shared_folder 
         Height          =   285
         Index           =   0
         Left            =   1440
         Locked          =   -1  'True
         MousePointer    =   3  'I-Beam
         TabIndex        =   32
         ToolTipText     =   "This is the folder that contains all the files you want to share with other Ripostle users!"
         Top             =   720
         Width           =   2895
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Browse"
         Height          =   255
         Index           =   0
         Left            =   4440
         TabIndex        =   31
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H0006D411&
         Caption         =   "Update Settings, And Continue using Ripostle!"
         Height          =   495
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   4320
         Width           =   2055
      End
      Begin VB.TextBox shared_folder 
         Height          =   285
         Index           =   1
         Left            =   1440
         Locked          =   -1  'True
         MousePointer    =   3  'I-Beam
         TabIndex        =   29
         ToolTipText     =   "This is the folder that all your downloads go to!"
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Browse"
         Height          =   255
         Index           =   1
         Left            =   4440
         TabIndex        =   28
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox up_on_start 
         BackColor       =   &H00A3B6BE&
         Caption         =   "Update Settings, and Set to ""Ready"" On start up"
         Height          =   375
         Left            =   1080
         TabIndex        =   93
         Top             =   3960
         Width           =   4095
      End
      Begin VB.Label Label20 
         BackColor       =   &H00A3B6BE&
         Caption         =   "(1-8)                                                  (1-8)"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2280
         TabIndex        =   366
         Top             =   3720
         Width           =   3015
      End
      Begin VB.Label Label19 
         BackColor       =   &H00A3B6BE&
         Caption         =   "Username:"
         Height          =   255
         Left            =   1320
         TabIndex        =   357
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label40 
         BackColor       =   &H00A3B6BE&
         Caption         =   "Download Folder:"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H00A3B6BE&
         Caption         =   "Preferences And Settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   37
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00A3B6BE&
         Caption         =   "Shared Folder:"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Width           =   1695
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   5640
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label34 
         BackColor       =   &H00A3B6BE&
         Caption         =   "Modem Speed:"
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label38 
         BackColor       =   &H00A3B6BE&
         Caption         =   "Max. Downloads At A Time:"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label Label39 
         BackColor       =   &H00A3B6BE&
         Caption         =   "Max. Uploads At a Time:"
         Height          =   255
         Left            =   3000
         TabIndex        =   33
         Top             =   3480
         Width           =   1935
      End
   End
   Begin VB.Frame search_window 
      BackColor       =   &H00A3B6BE&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   1080
      TabIndex        =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   5655
      Begin VB.ComboBox file_type 
         Height          =   315
         ItemData        =   "Form1.frx":E6DC0
         Left            =   3600
         List            =   "Form1.frx":E6DD6
         TabIndex        =   82
         Text            =   "MP3"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton stopsearch_button 
         Caption         =   "Cancel Search"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4080
         TabIndex        =   10
         ToolTipText     =   "End Searching"
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Search_Button 
         Caption         =   "Start Search"
         Height          =   255
         Left            =   2880
         TabIndex        =   5
         ToolTipText     =   "Start Searching"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox songname_txt 
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         ToolTipText     =   "Enter the Song name here:"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox artistname_txt 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         ToolTipText     =   "Enter the Artist name here:"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00A3B6BE&
         Height          =   3855
         Left            =   0
         TabIndex        =   91
         Top             =   960
         Width           =   5655
         Begin VB.TextBox Srch_tooltip 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   0
            TabIndex        =   359
            Top             =   3480
            Visible         =   0   'False
            Width           =   5655
         End
         Begin VB.TextBox label6 
            Height          =   285
            Left            =   3000
            TabIndex        =   96
            Top             =   3480
            Width           =   495
         End
         Begin MSFlexGridLib.MSFlexGrid Search_results 
            Height          =   3375
            Left            =   0
            TabIndex        =   95
            Top             =   0
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   5953
            _Version        =   393216
            Rows            =   3
            Cols            =   7
            FixedCols       =   0
            FocusRect       =   2
            HighLight       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
         End
         Begin VB.Label Label3 
            BackColor       =   &H00A3B6BE&
            Caption         =   "Found:"
            Height          =   255
            Left            =   2400
            TabIndex        =   92
            Top             =   3480
            Width           =   1455
         End
      End
      Begin VB.Label Label41 
         BackColor       =   &H00A3B6BE&
         Caption         =   "File Type:"
         Height          =   255
         Left            =   2880
         TabIndex        =   83
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00A3B6BE&
         Caption         =   "Artist Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00A3B6BE&
         Caption         =   "Title Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Frame transfers_window 
      BackColor       =   &H00A3B6BE&
      Height          =   4815
      Left            =   1080
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   5655
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00A3B6BE&
         Height          =   1815
         Left            =   0
         ScaleHeight     =   1755
         ScaleWidth      =   5595
         TabIndex        =   102
         Top             =   2640
         Width           =   5655
         Begin VB.VScrollBar VScroll3 
            Height          =   1815
            LargeChange     =   80
            Left            =   5400
            Max             =   4560
            Min             =   120
            SmallChange     =   30
            TabIndex        =   103
            Top             =   0
            Value           =   120
            Width           =   255
         End
         Begin VB.Frame uploads_frme 
            BackColor       =   &H00A3B6BE&
            Height          =   6375
            Left            =   0
            TabIndex        =   104
            Top             =   -120
            Width           =   5535
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   24
               Left            =   0
               Locked          =   -1  'True
               ScrollBars      =   1  'Horizontal
               TabIndex        =   364
               Top             =   480
               Width           =   2535
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   25
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   363
               Top             =   720
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   24
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   222
               Top             =   480
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   24
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   221
               Top             =   480
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   24
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   220
               Top             =   480
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   24
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   219
               Top             =   480
               Width           =   1095
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   25
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   218
               Top             =   720
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   25
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   217
               Top             =   720
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   25
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   216
               Top             =   720
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   25
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   215
               Top             =   720
               Width           =   1095
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   26
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   214
               Top             =   960
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   26
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   213
               Top             =   960
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   26
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   212
               Top             =   960
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   26
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   211
               Top             =   960
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   26
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   210
               Top             =   960
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   27
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   209
               Top             =   1200
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   27
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   208
               Top             =   1200
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   27
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   207
               Top             =   1200
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   27
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   206
               Top             =   1200
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   27
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   205
               Top             =   1200
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   28
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   204
               Top             =   1440
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   28
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   203
               Top             =   1440
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   28
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   202
               Top             =   1440
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   28
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   201
               Top             =   1440
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   28
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   200
               Top             =   1440
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   29
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   199
               Top             =   1680
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   29
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   198
               Top             =   1680
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   29
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   197
               Top             =   1680
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   29
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   196
               Top             =   1680
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   29
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   195
               Top             =   1680
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   30
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   194
               Top             =   1920
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   30
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   193
               Top             =   1920
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   30
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   192
               Top             =   1920
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   30
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   191
               Top             =   1920
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   30
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   190
               Top             =   1920
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   31
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   189
               Top             =   2160
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   31
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   188
               Top             =   2160
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   31
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   187
               Top             =   2160
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   31
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   186
               Top             =   2160
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   31
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   185
               Top             =   2160
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   32
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   184
               Top             =   2400
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   32
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   183
               Top             =   2400
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   32
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   182
               Top             =   2400
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   32
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   181
               Top             =   2400
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   32
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   180
               Top             =   2400
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   33
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   179
               Top             =   2640
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   33
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   178
               Top             =   2640
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   33
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   177
               Top             =   2640
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   33
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   176
               Top             =   2640
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   33
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   175
               Top             =   2640
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   34
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   174
               Top             =   2880
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   34
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   173
               Top             =   2880
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   34
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   172
               Top             =   2880
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   34
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   171
               Top             =   2880
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   34
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   170
               Top             =   2880
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   35
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   169
               Top             =   3120
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   35
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   168
               Top             =   3120
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   35
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   167
               Top             =   3120
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   35
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   166
               Top             =   3120
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   35
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   165
               Top             =   3120
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   36
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   164
               Top             =   3360
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   36
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   163
               Top             =   3360
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   36
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   162
               Top             =   3360
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               HideSelection   =   0   'False
               Index           =   36
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   161
               Top             =   3360
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   36
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   160
               Top             =   3360
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   37
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   159
               Top             =   3600
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   37
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   158
               Top             =   3600
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   37
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   157
               Top             =   3600
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   37
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   156
               Top             =   3600
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   37
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   155
               Top             =   3600
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   38
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   154
               Top             =   3840
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   38
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   153
               Top             =   3840
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   38
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   152
               Top             =   3840
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   38
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   151
               Top             =   3840
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   38
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   150
               Top             =   3840
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   39
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   149
               Top             =   4080
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   405
               Index           =   39
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   148
               Top             =   3960
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   39
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   147
               Top             =   4080
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   39
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   146
               Top             =   4080
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   39
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   145
               Top             =   4080
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   40
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   144
               Top             =   4320
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   40
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   143
               Top             =   4320
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   40
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   142
               Top             =   4320
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   40
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   141
               Top             =   4320
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   40
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   140
               Top             =   4320
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   41
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   139
               Top             =   4560
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   41
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   138
               Top             =   4560
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   41
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   137
               Top             =   4560
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   41
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   136
               Top             =   4560
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   41
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   135
               Top             =   4560
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   42
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   134
               Top             =   4800
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   42
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   133
               Top             =   4800
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   42
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   132
               Top             =   4800
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   42
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   131
               Top             =   4800
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   42
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   130
               Top             =   4800
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   43
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   129
               Top             =   5040
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   43
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   128
               Top             =   5040
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   43
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   127
               Top             =   5040
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   43
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   126
               Top             =   5040
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   43
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   125
               Top             =   5040
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   44
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   124
               Top             =   5280
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   405
               Index           =   44
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   123
               Top             =   5160
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   44
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   122
               Top             =   5280
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   44
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   121
               Top             =   5280
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   44
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   120
               Top             =   5280
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   45
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   119
               Top             =   5520
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   45
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   118
               Top             =   5520
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   45
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   117
               Top             =   5520
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   45
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   116
               Top             =   5520
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   45
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   115
               Top             =   5520
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   46
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   114
               Top             =   5760
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   46
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   113
               Top             =   5760
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   46
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   112
               Top             =   5760
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   46
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   111
               Top             =   5760
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   46
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   110
               Top             =   5760
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   47
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   109
               Top             =   6000
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   47
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   108
               Top             =   6000
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   47
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   107
               Top             =   6000
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   47
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   106
               Top             =   6000
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   47
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   105
               Top             =   6000
               Width           =   2535
            End
            Begin VB.Line Line3 
               X1              =   0
               X2              =   5520
               Y1              =   360
               Y2              =   360
            End
            Begin VB.Label Label16 
               BackColor       =   &H00A3B6BE&
               Caption         =   "%Done:"
               Height          =   255
               Left            =   4800
               TabIndex        =   227
               Top             =   120
               Width           =   615
            End
            Begin VB.Label Label15 
               BackColor       =   &H00A3B6BE&
               Caption         =   "KB/S:"
               Height          =   255
               Left            =   3720
               TabIndex        =   226
               Top             =   120
               Width           =   495
            End
            Begin VB.Label Label14 
               BackColor       =   &H00A3B6BE&
               Caption         =   "User:"
               Height          =   255
               Left            =   4320
               TabIndex        =   225
               Top             =   120
               Width           =   495
            End
            Begin VB.Label Label13 
               BackColor       =   &H00A3B6BE&
               Caption         =   "Status:"
               Height          =   255
               Left            =   2760
               TabIndex        =   224
               Top             =   120
               Width           =   615
            End
            Begin VB.Label Label8 
               BackColor       =   &H00A3B6BE&
               Caption         =   "Filename:"
               Height          =   255
               Left            =   120
               TabIndex        =   223
               Top             =   120
               Width           =   735
            End
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Stop Transfer"
         Height          =   255
         Left            =   1320
         TabIndex        =   101
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Recent Finished Downloads"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2400
         TabIndex        =   100
         Top             =   4560
         Width           =   2175
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Clear List"
         Height          =   255
         Left            =   4560
         TabIndex        =   99
         Top             =   4560
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00A3B6BE&
         Height          =   1815
         Left            =   0
         ScaleHeight     =   1755
         ScaleWidth      =   5595
         TabIndex        =   97
         Top             =   480
         Width           =   5655
         Begin VB.VScrollBar VScroll2 
            Height          =   1815
            LargeChange     =   80
            Left            =   5400
            Max             =   4560
            Min             =   120
            SmallChange     =   30
            TabIndex        =   98
            Top             =   0
            Value           =   120
            Width           =   255
         End
         Begin VB.Frame downloads_frame 
            BackColor       =   &H00A3B6BE&
            Height          =   6375
            Left            =   0
            TabIndex        =   228
            Top             =   -120
            Width           =   5535
            Begin VB.TextBox status 
               Height          =   285
               Index           =   0
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   348
               Top             =   480
               Width           =   1095
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   0
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   347
               Top             =   480
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   0
               Left            =   3600
               Locked          =   -1  'True
               TabIndex        =   346
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   4200
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   345
               Top             =   480
               Width           =   855
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   0
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   344
               Top             =   480
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   1
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   343
               Top             =   720
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   1
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   342
               Top             =   720
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   1
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   341
               Top             =   720
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   1
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   340
               Top             =   720
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   1
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   339
               Top             =   720
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   2
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   338
               Top             =   960
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   2
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   337
               Top             =   960
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   2
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   336
               Top             =   960
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   2
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   335
               Top             =   960
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   2
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   334
               Top             =   960
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   3
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   333
               Top             =   1200
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   3
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   332
               Top             =   1200
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   3
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   331
               Top             =   1200
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   3
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   330
               Top             =   1200
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   3
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   329
               Top             =   1200
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   4
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   328
               Top             =   1440
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   4
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   327
               Top             =   1440
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   4
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   326
               Top             =   1440
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   4
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   325
               Top             =   1440
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   4
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   324
               Top             =   1440
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   5
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   323
               Top             =   1680
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   5
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   322
               Top             =   1680
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   5
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   321
               Top             =   1680
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   5
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   320
               Top             =   1680
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   5
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   319
               Top             =   1680
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   6
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   318
               Top             =   1920
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   6
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   317
               Top             =   1920
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   6
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   316
               Top             =   1920
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   6
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   315
               Top             =   1920
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   6
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   314
               Top             =   1920
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   7
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   313
               Top             =   2160
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   7
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   312
               Top             =   2160
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   7
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   311
               Top             =   2160
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   7
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   310
               Top             =   2160
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   7
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   309
               Top             =   2160
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   8
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   308
               Top             =   2400
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   8
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   307
               Top             =   2400
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   8
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   306
               Top             =   2400
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   8
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   305
               Top             =   2400
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   8
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   304
               Top             =   2400
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   9
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   303
               Top             =   2640
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   9
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   302
               Top             =   2640
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   9
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   301
               Top             =   2640
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   9
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   300
               Top             =   2640
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   9
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   299
               Top             =   2640
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   10
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   298
               Top             =   2880
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   10
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   297
               Top             =   2880
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   10
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   296
               Top             =   2880
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   10
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   295
               Top             =   2880
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   10
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   294
               Top             =   2880
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   11
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   293
               Top             =   3120
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   11
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   292
               Top             =   3120
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   11
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   291
               Top             =   3120
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   11
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   290
               Top             =   3120
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   11
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   289
               Top             =   3120
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   12
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   288
               Top             =   3360
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   12
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   287
               Top             =   3360
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   12
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   286
               Top             =   3360
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   12
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   285
               Top             =   3360
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   12
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   284
               Top             =   3360
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   13
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   283
               Top             =   3600
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   13
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   282
               Top             =   3600
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   13
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   281
               Top             =   3600
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   13
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   280
               Top             =   3600
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   13
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   279
               Top             =   3600
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   14
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   278
               Top             =   3840
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   14
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   277
               Top             =   3840
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   14
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   276
               Top             =   3840
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   14
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   275
               Top             =   3840
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   14
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   274
               Top             =   3840
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   15
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   273
               Top             =   4080
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   15
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   272
               Top             =   4080
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   15
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   271
               Top             =   4080
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   15
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   270
               Top             =   4080
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   15
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   269
               Top             =   4080
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   16
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   268
               Top             =   4320
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   16
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   267
               Top             =   4320
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   16
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   266
               Top             =   4320
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   16
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   265
               Top             =   4320
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   16
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   264
               Top             =   4320
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   17
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   263
               Top             =   4560
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   17
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   262
               Top             =   4560
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   17
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   261
               Top             =   4560
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   17
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   260
               Top             =   4560
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   17
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   259
               Top             =   4560
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   18
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   258
               Top             =   4800
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   18
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   257
               Top             =   4800
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   18
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   256
               Top             =   4800
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   18
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   255
               Top             =   4800
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   18
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   254
               Top             =   4800
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   19
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   253
               Top             =   5040
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   19
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   252
               Top             =   5040
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   19
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   251
               Top             =   5040
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   19
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   250
               Top             =   5040
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   19
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   249
               Top             =   5040
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   20
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   248
               Top             =   5280
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   20
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   247
               Top             =   5280
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   20
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   246
               Top             =   5280
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   20
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   245
               Top             =   5280
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   20
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   244
               Top             =   5280
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   21
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   243
               Top             =   5520
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   21
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   242
               Top             =   5520
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   21
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   241
               Top             =   5520
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   21
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   240
               Top             =   5520
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   21
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   239
               Top             =   5520
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   22
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   238
               Top             =   5760
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   22
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   237
               Top             =   5760
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   22
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   236
               Top             =   5760
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   22
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   235
               Top             =   5760
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   22
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   234
               Top             =   5760
               Width           =   2535
            End
            Begin VB.TextBox status 
               Height          =   285
               Index           =   23
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   233
               Top             =   6000
               Width           =   855
            End
            Begin VB.TextBox done 
               Height          =   285
               Index           =   23
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   232
               Top             =   6000
               Width           =   375
            End
            Begin VB.TextBox kbs 
               Height          =   285
               Index           =   23
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   231
               Top             =   6000
               Width           =   375
            End
            Begin VB.TextBox usr 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   23
               Left            =   3840
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   230
               Top             =   6000
               Width           =   1095
            End
            Begin VB.TextBox fname 
               Height          =   285
               Index           =   23
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   229
               Top             =   6000
               Width           =   2535
            End
            Begin VB.Line Line2 
               X1              =   0
               X2              =   5520
               Y1              =   360
               Y2              =   360
            End
            Begin VB.Label Label12 
               BackColor       =   &H00A3B6BE&
               Caption         =   "%Done:"
               Height          =   255
               Left            =   4800
               TabIndex        =   352
               Top             =   120
               Width           =   615
            End
            Begin VB.Label Label11 
               BackColor       =   &H00A3B6BE&
               Caption         =   "KB/S:"
               Height          =   255
               Left            =   3720
               TabIndex        =   351
               Top             =   120
               Width           =   495
            End
            Begin VB.Label Label10 
               BackColor       =   &H00A3B6BE&
               Caption         =   "User:"
               Height          =   255
               Left            =   4320
               TabIndex        =   350
               Top             =   120
               Width           =   495
            End
            Begin VB.Label Label7 
               BackColor       =   &H00A3B6BE&
               Caption         =   "Filename:"
               Height          =   255
               Left            =   120
               TabIndex        =   349
               Top             =   120
               Width           =   735
            End
            Begin VB.Label statlab 
               BackColor       =   &H00A3B6BE&
               Caption         =   "Status:"
               Height          =   255
               Left            =   2760
               TabIndex        =   362
               Top             =   120
               Width           =   615
            End
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Start Transfer"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Label18 
         BackColor       =   &H00A3B6BE&
         Caption         =   "-Uploads-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label17 
         BackColor       =   &H00A3B6BE&
         Caption         =   "-Downloads-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Idle (Pref Not Set)"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   4920
      TabIndex        =   23
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Label Label35 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4320
      TabIndex        =   22
      Top             =   6960
      Width           =   615
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   975
      Left            =   -1560
      TabIndex        =   20
      Top             =   1920
      Width           =   1575
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Label usersval 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2520
      TabIndex        =   13
      Top             =   6960
      Width           =   975
   End
   Begin VB.Label onlineusers 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Users On Your List:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Image Image11 
      Height          =   495
      Left            =   6480
      ToolTipText     =   "Other Stuff (Help, About, Credits, etc)"
      Top             =   7680
      Width           =   495
   End
   Begin VB.Image Image10 
      Height          =   615
      Left            =   5520
      ToolTipText     =   "Preferences And Settings"
      Top             =   7800
      Width           =   495
   End
   Begin VB.Image Image9 
      Height          =   615
      Left            =   4680
      ToolTipText     =   "Chat"
      Top             =   7800
      Width           =   495
   End
   Begin VB.Image Image8 
      Height          =   375
      Left            =   3600
      ToolTipText     =   "Information"
      Top             =   8280
      Width           =   615
   End
   Begin VB.Image Image7 
      Height          =   495
      Left            =   3600
      ToolTipText     =   "My Files"
      Top             =   7680
      Width           =   615
   End
   Begin VB.Image Image6 
      Height          =   495
      Left            =   2640
      ToolTipText     =   "Play/Rip Your Songs"
      Top             =   7920
      Width           =   495
   End
   Begin VB.Image Image5 
      Height          =   615
      Left            =   1680
      ToolTipText     =   "View your transfers"
      Top             =   7800
      Width           =   495
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   720
      ToolTipText     =   "Search For Songs"
      Top             =   7680
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   6600
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   6120
      ToolTipText     =   "Minimize"
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   1800
      MousePointer    =   5  'Size
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim selected_row As Integer
Dim ahg
Dim Mp3KHZ
Dim Mp3MODE
Dim Mp3Bitrate
Private Declare Function SetWindowPos Lib "user32" _
         (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
          ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const HWND_TOPMOST = -1
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Public strFilePath As String
Dim dline
Dim GotBytes(24)
Dim SentBytes(47)
Dim Tnp1(24)
Dim Tnp2(24)
Dim Tnp3(24)
Dim Tnp4(24)
Dim Tnp5(24)
Dim Tnp6(24)
Dim Tnp7(24)
Dim SelRow_SearchRes
Dim ipcfg(47)
Dim curntdling
Dim PosIpToCont As String
Dim macdl
Dim wordx(9999) As String
Dim intConnection As Integer
Dim upline
Dim wordz(9999) As String
Dim curntupling
Dim FILESIZEY(47)
Dim FILESIZEU(24)
Dim RateUP(47)
Dim RateDL(24)
Dim filenumberg(47)
Dim FileWritePosi(47)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'----------------------------------------------------
Public Function PutWindowOnTop(pFrm As Form)
  Dim lngWindowPosition As Long
  lngWindowPosition = SetWindowPos(pFrm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Function
Private Sub ads_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
If URL = "http://bannerpower.com/cgi-bin/redirect.cgi?2562&1&1" Then
Cancel = True
Dim lReturn As Long
lReturn = ShellExecute(hWnd, "open", _
URL, _
vbNull, vbNull, SW_SHOWNORMAL)
End If
End Sub
Private Sub ads_NewWindow2(ppDisp As Object, Cancel As Boolean)
Set ppDisp = ads2.Object
End Sub

Private Sub ads2_NewWindow2(ppDisp As Object, Cancel As Boolean)
Set ppDisp = ads2.Object
End Sub

Private Sub cnt4listsharetimout_Timer()
cnt4listsharetimout.Enabled = False
con4listshare
End Sub

Private Sub Command10_Click()
Shell "notepad.exe " & My_files_Dtabs.Path & "\" & My_files_Dtabs.FileName, vbNormalFocus
End Sub

Private Sub Command11_Click()
Shell "notepad.exe " & My_files_Lyrics.Path & "\" & My_files_Lyrics.FileName, vbNormalFocus
End Sub

Private Sub Command12_Click()
If Command12.Caption = "Play" Then
MediaPlayer1.AutoStart = True
Command12.Caption = "Stop"
MediaPlayer1.Open My_files_Mp3z.Path & "\" & My_files_Mp3z.FileName
Else
MediaPlayer1.Stop
Command12.Caption = "Play"
End If
End Sub

Private Sub Command13_Click()
Shell "notepad.exe " & My_files_Gtabs.Path & "\" & My_files_Gtabs.FileName, vbNormalFocus
End Sub

Private Sub Command14_Click()
Shell "notepad.exe " & My_files_Shm.Path & "\" & My_files_Shm.FileName, vbNormalFocus
End Sub

Private Sub Command2_Click()
Downloader(0).Close
End Sub

Private Sub Command6_Click()
Form5.Visible = True
PutWindowOnTop Form5
End Sub

Private Sub Command7_Click(Index As Integer)
Form4.Visible = True
Form4.Label2.Caption = Index
If Index = 0 Then Form4.Caption = "Select Path For Shared Folder" Else Form4.Caption = "Select Path For Download Folder"
End Sub

Private Sub Command8_Click()
sharelist.Enabled = True
Label36.Caption = "Idle (Ready!)"
sock.Close
sock.Listen
savepref
Form1.Label36.ForeColor = &HFF00&
getsearch_frm.Search_results.Path = Form1.My_files_Midi.Path
If max_ups.Text = "" Then max_ups.Text = "1"
If max_ups.Text < "1" Then max_ups.Text = "1"
getsearch_frm.Visible = True

If Uploader(0).State = 0 Then Uploader(0).Listen

End Sub

Private Sub Command9_Click()
If Command9.Caption = "Play" Then
MediaPlayer1.AutoStart = True
Command9.Caption = "Stop"
MediaPlayer1.Open My_files_Midi.Path & My_files_Midi.FileName
Else
MediaPlayer1.Stop
Command9.Caption = "Play"
End If
End Sub

Private Sub cSysTray1_MouseDblClick(Button As Integer, Id As Long)
Form1.Visible = True
Form1.Show
End Sub
Private Sub filesval_Change()
If filesval.Caption < 0 Then filesval.Caption = 0
End Sub

Private Sub filename_list_Click()
tooltip1.Text = filename_list.List(filename_list.ListIndex)
tooltip1.Visible = True
End Sub

Private Sub DL_Transfer_Bot_Timer()
Dim Lineer
If curntdling = max_dls.Text Then Exit Sub

Lineer = ""
For i = 0 To 23
If status(i) = "Queued" Then
Lineer = i
GoTo jpac
End If
Next

jpac:
If Lineer = "" Then Exit Sub
status(Lineer).Text = "Downloading"
Open fname(Lineer) & ".dli" For Input As #90
Input #90, Tnp1(Lineer)
Input #90, Tnp2(Lineer)
Input #90, Tnp3(Lineer)
Input #90, Tnp4(Lineer)
Input #90, Tnp5(Lineer)
Input #90, Tnp6(Lineer)
Input #90, Tnp7(Lineer)
Close #90

Downloader(Lineer).Connect Tnp7(Lineer), 7472
Downloader_Timout(Lineer).Enabled = True
curntdling = curntdling + 1
End Sub

Private Sub Downloader_Connect(Index As Integer)
Downloader(Index).SendData "IWT" & Tnp1(Index) & ";" & Tnp3(Index)
DL_KBS_TMR(i).Enabled = True
End Sub

Private Sub Downloader_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim Packet As String

Downloader(Index).GetData Packet

If Mid(Packet, 1, 3) = "SAF" Then
status(Index) = "Queued"
Exit Sub
End If

If Mid(Packet, 1, 3) = "TFS" Then
FILESIZEU(Index) = Mid(Packet, 3, Len(Packet))
On Error GoTo Get4FirstTime
agf = shared_folder(1) & fname(Index).Text
a = FileLen(agf)
Open agf For Binary Access Write As #filenumberg(Index)
a = a - 2 'roll back a bit to make sure u get the file
FileWritePosi(Index) = a
Downloader(Index).SendData "RSM" & a
Exit Sub
Get4FirstTime:
Open agf For Binary Access Write As #filenumberg(Index)
Downloader(Index).SendData "FST"
FileWritePosi(Index) = 1
Exit Sub
End If



If Mid(Packet, 1, 3) = "RAW" Then
dart = Mid(Packet, 4, Len(Packet))

GotBytes(Index) = GotBytes(Index) + Len(dart)
jki = GotBytes(Index)  '__
jku = FILESIZEU(Index)   '|
If jki > jku Then        '|
jop = jki - jku          '|  take unwanted bytes off end of file
joh = 1024 = jop         '|
hobo = Mid(dart, 1, joh) '|
dart = hobo            '__|
RateDL(Index) = RateDL(Index) + Len(dart)
Put #filenumberg(Index), FileWritePosi(Index), dart
End If
End If


End Sub

Private Sub Downloader_Timout_Timer(Index As Integer)
If Downloader(Index).State <> 7 Then
status(Index).Text = "Failed"
Downloader(Index).Close
End If
End Sub

Private Sub fname_Change(Index As Integer)
fname(Index).ToolTipText = fname(Index).Text
End Sub

Private Sub fname_Click(Index As Integer)
For i = 0 To 47
If i = Index Then fname(i).BackColor = &H8000000D Else fname(i).BackColor = &H80000005
If i = Index Then status(i).BackColor = &H8000000D Else status(i).BackColor = &H80000005
If i = Index Then usr(i).BackColor = &H8000000D Else usr(i).BackColor = &H80000005
If i = Index Then done(i).BackColor = &H8000000D Else done(i).BackColor = &H80000005
If i = Index Then kbs(i).BackColor = &H8000000D Else kbs(i).BackColor = &H80000005
Next
End Sub

Private Sub fname_DblClick(Index As Integer)
Form3.Visible = True
End Sub

Private Sub foot_Connect()
DateTagIP PosIpToCont
Label36.Caption = "Building a list of users"
cnt4listsharetimout.Enabled = False
foot.SendData "mo"
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
Private Sub foot_DataArrival(ByVal bytesTotal As Long)
Dim ag As String
foot.GetData ag
If ag = "x" Then
'sends the person who diald their list
For i = 0 To List1.ListCount
jebg = List1.List(i)
jebg = Replace(jebg, "0", "à")
jebg = Replace(jebg, "1", "é")
jebg = Replace(jebg, "2", "ä")
jebg = Replace(jebg, "3", "â")
jebg = Replace(jebg, "4", "ê")
jebg = Replace(jebg, "5", "ç")
jebg = Replace(jebg, "6", "å")
jebg = Replace(jebg, "7", "ï")
jebg = Replace(jebg, "8", "è")
jebg = Replace(jebg, "9", "ë")
jebg = Replace(jebg, ".", "@")
foot.SendData jebg
Sleep 200, False
Next
IpGetter.Visible = True
ipo = IpGetter.IPLIST.List(0)
Unload IpGetter
jebg = ipo

jebg = Replace(jebg, "0", "à")
jebg = Replace(jebg, "1", "é")
jebg = Replace(jebg, "2", "ä")
jebg = Replace(jebg, "3", "â")
jebg = Replace(jebg, "4", "ê")
jebg = Replace(jebg, "5", "ç")
jebg = Replace(jebg, "6", "å")
jebg = Replace(jebg, "7", "ï")
jebg = Replace(jebg, "8", "è")
jebg = Replace(jebg, "9", "ë")
jebg = Replace(jebg, ".", "@")
foot.SendData jebg
Sleep 300, False
foot.SendData "xx"
cleanlist1
Label36.Caption = "Done Building List"
Sleep 30000, False
Label36.Caption = "Idle"
Else

jebg = ag
jebg = Replace(jebg, "à", "0")
jebg = Replace(jebg, "é", "1")
jebg = Replace(jebg, "ä", "2")
jebg = Replace(jebg, "â", "3")
jebg = Replace(jebg, "ê", "4")
jebg = Replace(jebg, "ç", "5")
jebg = Replace(jebg, "å", "6")
jebg = Replace(jebg, "ï", "7")
jebg = Replace(jebg, "è", "8")
jebg = Replace(jebg, "ë", "9")
jebg = Replace(jebg, "@", ".")

List2.AddItem jebg
End If
End Sub
Sub cleanlist1()
For i = 0 To List2.ListCount
List1.AddItem List2.List(i)
Next
List2.Clear

ip_listcleanup.RMDUPES
End Sub

Private Sub RemoveDupe(lst As ListBox)
For i = 0 To List1.ListCount - 1
   For b = i + 1 To List1.ListCount - 1
   fd1 = List1.List(i)
   asw1 = Right(fd1, 11)
   fd1 = Replace(fd1, asw1, "")
   
   fd2 = List1.List(b)
   asw2 = Right(fd2, 11)
   fd2 = Replace(fd2, asw2, "")
   If fd1 = fd2 Then List1.List(b) = ""
   Next
Next

h = List1.ListCount - 1
For i = 1 To List1.ListCount
If List1.List(h) = "" Then List1.RemoveItem h
h = h - 1
Next
End Sub

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
        
          
        Else
            '-- if not, increase iPos..
            iPos = iPos + 1
        End If
    Loop
    '-- used to unselect the last selected l
    '     ine..
    lst.Text = "~~~^^~~~"
End Sub

Sub cleanlist2()

End Sub

Private Sub Form_Load()
draw_Search_Results_Window
 'draw_upload_dl_window

FT1.NoShowColor = RGB(255, 0, 0)

'loadiplist
loadpref
ip_listcleanup.Visible = True
ads.Navigate "http://ripostle.tripod.com/ads.html"
Sendsearch_frm.Visible = True
If up_on_start.Value = 1 Then Command8_Click

For i = 0 To 47
status(i).Width = 1095
kbs(i).Left = 3600
kbs(i).Width = 615
usr(i).Left = 4200
usr(i).Width = 855
Next
curntupling = 0
curntdling = "0"
ip_listcleanup.Visible = True

For i = 0 To 23
On Error Resume Next
Load Downloader(i)
Load Downloader_Timout(i)
Load DL_KBS_TMR(i)
Next

For i = 24 To 47
On Error Resume Next
Load Uploader(i)
Next

For i = 25 To 47
Load UL_KBS_TMR(i)
Next

kop = 100
For i = 0 To 47
filenumberg(i) = kop
kop = kop + 1
Next
End Sub

Private Sub Form_Terminate()
Image3_Click
saveiplist
savepref
End Sub

Private Sub Form_Unload(Cancel As Integer)
Image3_Click
saveiplist
savepref
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'this is the title
If Button = 0 Then
Y1 = y
X1 = x
End If
If Button = 1 Then
Form1.Left = Form1.Left - (X1 - x)
Form1.Top = Form1.Top - (Y1 - y)
End If
End Sub

Private Sub Image10_Click()
pref_window.Visible = True
transfers_window.Visible = False
search_window.Visible = False
Myfiles_window.Visible = False
information_window.Visible = False
End Sub

Private Sub Image2_Click()
'this is the green _ box
Form1.Visible = False
End Sub

Private Sub Image3_Click()
Form1.MousePointer = 11
savepref
saveiplist
End
End Sub

Private Sub Image4_Click()
search_window.Visible = True
transfers_window.Visible = False
pref_window.Visible = False
Myfiles_window.Visible = False
information_window.Visible = False
End Sub

Private Sub Image5_Click()
transfers_window.Visible = True
search_window.Visible = False
pref_window.Visible = False
Myfiles_window.Visible = False
information_window.Visible = False
End Sub

Private Sub Image7_Click()
Myfiles_window.Visible = True
search_window.Visible = False
pref_window.Visible = False
transfers_window.Visible = False
information_window.Visible = False
End Sub

Private Sub Image8_Click()
pref_window.Visible = False
transfers_window.Visible = False
search_window.Visible = False
Myfiles_window.Visible = False
information_window.Visible = True
End Sub

Private Sub My_files_Click()
Form1.MousePointer = 11
Form1.MousePointer = 99
Form1.MouseIcon = theicon.Picture
End Sub

Private Sub iplistcleanup_Timer()
iplistcleanup.Enabled = False
ip_listcleanup.Visible = True
End Sub

Private Sub max_dls_Change()
If IsNumeric(max_dls.Text) = False Then MsgBox "This Box Must be A Number from 1-24"
End Sub

Private Sub max_ups_Change()
If IsNumeric(max_ups.Text) = False Then MsgBox "This Box Must be A Number from 1-24"
End Sub

Private Sub MediaPlayer1_EndOfStream(ByVal Result As Long)
Command9.Caption = "Play"
Command12.Caption = "Play"
End Sub

Private Sub ModemSpeed_Change()
If ModemSpeed.Text = "28.8" Then Option1.Value = True
If ModemSpeed.Text = "36.6" Then Option2.Value = True
If ModemSpeed.Text = "56k" Then Option3.Value = True
If ModemSpeed.Text = "ISDN" Then Option4.Value = True
If ModemSpeed.Text = "DSL" Then Option5.Value = True
If ModemSpeed.Text = "Cable" Then Option6.Value = True
If ModemSpeed.Text = "T1" Then Option7.Value = True
End Sub

Private Sub My_files_Dtabs_Click()
Open My_files_Dtabs.Path & My_files_Dtabs.FileName For Random As #1
Label42.Caption = LOF(1) / 1000000
Close #1
Label42.Caption = Format(Label42.Caption, "#.##")
End Sub

Private Sub My_files_Gtabs_Click()
Open My_files_Gtabs.Path & My_files_Gtabs.FileName For Random As #1
Label62.Caption = LOF(1) / 1000000
Close #1
Label62.Caption = Format(Label42.Caption, "#.##")
End Sub

Private Sub My_files_Lyrics_Click()
Open My_files_Lyrics.Path & My_files_Lyrics.FileName For Random As #1
Label47.Caption = LOF(1) / 1000000
Close #1
Label47.Caption = Format(Label47.Caption, "#.##")
End Sub

Private Sub My_files_Midi_Click()
Form1.MousePointer = 11
MediaPlayer1.AutoStart = False

MediaPlayer1.Open My_files_Midi.Path & My_files_Midi.FileName
Sleep 2000, False


Label52.Caption = MediaPlayer1.Duration / 60
Label52.Caption = Format(Label52.Caption, "#.##")
Open MediaPlayer1.FileName For Random As #1
Label61.Caption = LOF(1) / 1000000
Close #1
Label61.Caption = Format(Label61.Caption, "#.##")
Form1.MousePointer = 99
End Sub

Private Sub My_files_Mp3z_Click()
On Error Resume Next
  Dim accMP3Info As MP3Info
  agoh = My_files_Mp3z.Path & My_files_Mp3z.FileName
  getMP3Info agoh, accMP3Info
  
filsize_mf.Caption = accMP3Info.SIZE / 1000024
length_mf = accMP3Info.LENGTH / 60
bitrate_mf = accMP3Info.BITRATE

filsize_mf.Caption = Format(filsize_mf.Caption, "#.##")
length_mf.Caption = Format(length_mf.Caption, "#.##")
End Sub

Private Sub My_files_Shm_Click()
Open My_files_Shm.Path & My_files_Shm.FileName For Random As #1
Label60.Caption = LOF(1) / 1000024
Close #1
Label60.Caption = Format(Label60.Caption, "#.##")
End Sub

Private Sub my_modem_speed_ItemCheck(Item As Integer)
MsgBox my_modem_speed.List(Item)
End Sub

Private Sub Option1_Click()
ModemSpeed.Text = "28.8"
End Sub

Private Sub Option2_Click()
ModemSpeed.Text = "36.6"
End Sub

Private Sub Option3_Click()
ModemSpeed.Text = "56k"
End Sub

Private Sub Option4_Click()
ModemSpeed.Text = "ISDN"
End Sub

Private Sub Option5_Click()
ModemSpeed.Text = "DSL"
End Sub

Private Sub Option6_Click()
ModemSpeed.Text = "Cable"
End Sub

Private Sub Option7_Click()
ModemSpeed.Text = "T1"
End Sub

Private Sub scrolltext1_Timer()
Label27.Left = Label27.Left - 30
If Label27.Left = -11880 Then Label27.Left = 5500
End Sub

Private Sub Search_Button_Click()
If Form1.Search_results.Rows <> 3 Then
a = Form1.Search_results.Rows
b = a - 3
For i = 1 To b
Form1.Search_results.RemoveItem (a)
a = a - 1
Next
End If
Form1.Search_results.Rows = 3
Sendsearch_frm.mart.Text = "1"
zort.Text = 0
Search_results.Clear
Form1.draw_Search_Results_Window
label6.Text = 0
Search_results.TextMatrix(1, 0) = "Searching... Please wait!"
artistname_txt.Text = Replace(artistname_txt.Text, "-", " ", 1, , vbTextCompare)
songname_txt.Text = Replace(songname_txt.Text, "-", " ", 1, , vbTextCompare)


Sendsearch_frm.sendsearch.Close
Sendsearch_frm.search4song

stopsearch_button.Enabled = True
Search_Button.Enabled = False

Label36.Caption = "Searching..."
End Sub
Sub reload_getsearch_frm()
Unload getsearch_frm
getsearch_frm.Visible = True
getsearch_frm.Search_results.Path = Form1.My_files_Midi.Path
getsearch_frm.Search_results.Pattern = "*.*"
getsearch_frm.getsearch.Close
getsearch_frm.getsearch.Listen
End Sub

Private Sub search_results_Click()
Srch_tooltip.Text = Search_results.TextMatrix(Search_results.MouseRow, 0)
Srch_tooltip.Visible = True
search_tooltip_tmr.Enabled = True
SelRow_SearchRes = Search_results.MouseRow
End Sub

Private Sub Search_results_DblClick()
dline = ""
For i = 0 To 23
If fname(i).Text = "" Then dline = i
If dline = i Then GoTo hgd
Next

hgd:
If dline = "" Then
MsgBox "No Room Left, Wait for Some Downloads to finish... Then Try later!"
Exit Sub
End If

fname(dline).Text = Search_results.TextMatrix(SelRow_SearchRes, 0)
status(dline).Text = "Queued"
kbs(dline).Text = "0"
done(dline).Text = "0"
usr(dline).Text = Search_results.TextMatrix(SelRow_SearchRes, 2)
ipcfg(dline) = Search_results.TextMatrix(SelRow_SearchRes, 6)


Srch_tooltip.BackColor = &HFF00&

Close #19
Open fname(dline) & ".dli" For Output As #19
Write #19, Search_results.TextMatrix(SelRow_SearchRes, 0)
Write #19, Search_results.TextMatrix(SelRow_SearchRes, 1)
Write #19, Search_results.TextMatrix(SelRow_SearchRes, 2)
Write #19, Search_results.TextMatrix(SelRow_SearchRes, 3)
Write #19, Search_results.TextMatrix(SelRow_SearchRes, 4)
Write #19, Search_results.TextMatrix(SelRow_SearchRes, 5)
Write #19, Search_results.TextMatrix(SelRow_SearchRes, 6)
Close #19
End Sub

Private Sub search_tooltip_tmr_Timer()
Srch_tooltip.Visible = False
End Sub

Private Sub shared_folder_Change(Index As Integer)
On Error GoTo err

If Index = 0 Then
My_files_Mp3z.Path = shared_folder(Index).Text
My_files_Lyrics.Path = shared_folder(Index).Text
My_files_Dtabs.Path = shared_folder(Index).Text
My_files_Midi.Path = shared_folder(Index).Text
My_files_Gtabs.Path = shared_folder(Index).Text
My_files_Shm.Path = shared_folder(Index).Text
Update_MFlist_Timer
End If
err:
Exit Sub
End Sub

Private Sub sharelist_Timer()
sharelistcnt.Text = sharelistcnt.Text + 1
End Sub

Private Sub sharelistcnt_Change()
If sharelistcnt.Text = "5" Then con4listshare
End Sub

Sub con4listshare()
Dim iptocont As String
Dim hds(2)
'==try's to make a connection for list share===
sharelistcnt = "0"
Randomize Timer
PosIpToCont = Int(Rnd * List1.ListCount)
iptocont = List1.List(PosIpToCont)
mactoc = 1
For i = 1 To Len(iptocont)
a = Mid(iptocont, i, 1)
If a = ";" Then
a = ""
mactoc = mactoc + 1
End If
hds(mactoc) = hds(mactoc) & a
Next
iptocont = hds(1)

foot.Close
foot.Connect iptocont, 7470
Label36.Caption = "Scanning for Users"
cnt4listsharetimout.Enabled = True
End Sub

Private Sub sock_ConnectionRequest(ByVal requestID As Long)
sock.Close
sock.Accept requestID
End Sub

Private Sub sock_DataArrival(ByVal bytesTotal As Long)
Dim ag As String
sock.GetData ag
If ag = "mo" Then
'=======B list to A list====
'This sends the person who called, their ip list

For i = 0 To List1.ListCount
jebg = List1.List(i)
jebg = Replace(jebg, "0", "à")
jebg = Replace(jebg, "1", "é")
jebg = Replace(jebg, "2", "ä")
jebg = Replace(jebg, "3", "â")
jebg = Replace(jebg, "4", "ê")
jebg = Replace(jebg, "5", "ç")
jebg = Replace(jebg, "6", "å")
jebg = Replace(jebg, "7", "ï")
jebg = Replace(jebg, "8", "è")
jebg = Replace(jebg, "9", "ë")
jebg = Replace(jebg, ".", "@")
sock.SendData jebg
Sleep 200, False
Next
Sleep 200, False
sock.SendData "x"
Exit Sub
End If

If ag = "xx" Then
cleanlist1

sock.Close
sock.Listen

Else
jebg = ag
jebg = Replace(jebg, "à", "0")
jebg = Replace(jebg, "é", "1")
jebg = Replace(jebg, "ä", "2")
jebg = Replace(jebg, "â", "3")
jebg = Replace(jebg, "ê", "4")
jebg = Replace(jebg, "ç", "5")
jebg = Replace(jebg, "å", "6")
jebg = Replace(jebg, "ï", "7")
jebg = Replace(jebg, "è", "8")
jebg = Replace(jebg, "ë", "9")
jebg = Replace(jebg, "@", ".")

List1.AddItem jebg
End If
End Sub

Private Sub Socket1_Accept(SocketId As Integer)

End Sub

Private Sub status_Change(Index As Integer)
If Index < 24 Then
If status(Index).Text = "Done!" Then
Downloader(Index).Close
curntdling = curntdling - 1
MsgBox curntdling
End If
Else
If status(Index).Text = "Done!" Then
Uploader(Index).Close
MsgBox curntupling
End If
End If
End Sub

Private Sub stopsearch_button_Click()
stopsearch_button.Enabled = False
Search_Button.Enabled = True
Sendsearch_frm.mart.Text = "0"
Label36.Caption = "Idle"
If Form1.label6.Text < "0" Then
Form1.Search_results.TextMatrix(1, 0) = "File Not Found! Check your"
Form1.Search_results.Rows = 4
Form1.Search_results.TextMatrix(2, 0) = "Spelling, Or Try Again In"
Form1.Search_results.TextMatrix(3, 0) = "A While... Thank You!"
End If
Sendsearch_frm.sendsearch.Close
End Sub

Private Sub tooltip1_Change()
tooltip1_tmr.Enabled = True
End Sub

Private Sub tooltip1_tmr_Timer()
tooltip1_tmr.Enabled = False
tooltip1.Visible = False
End Sub


Private Sub Timer1_Timer()

End Sub

Private Sub Timer2_Timer()

End Sub

Private Sub Update_MFlist_Timer()
My_files_Mp3z.Refresh
My_files_Lyrics.Refresh
My_files_Dtabs.Refresh
My_files_Midi.Refresh
My_files_Gtabs.Refresh
My_files_Shm.Refresh

total_mp3z.Caption = "Total: " & My_files_Mp3z.ListCount
total_lyrics.Caption = "Total: " & My_files_Lyrics.ListCount
total_drumtabs.Caption = "Total: " & My_files_Dtabs.ListCount
total_midi.Caption = "Total: " & My_files_Midi.ListCount
total_gtab.Caption = "Total: " & My_files_Gtabs.ListCount
total_shm.Caption = "Total: " & My_files_Shm.ListCount
End Sub

Sub draw_Search_Results_Window()
Search_results.TextMatrix(0, 0) = "FileName:"
Search_results.TextMatrix(0, 1) = "Length:"
Search_results.TextMatrix(0, 2) = "User:"
Search_results.TextMatrix(0, 3) = "Size:"
Search_results.TextMatrix(0, 4) = "Bitrate:"
Search_results.TextMatrix(0, 5) = "Speed:"
Search_results.TextMatrix(0, 6) = "IP#:"
Search_results.ColWidth(1) = Len(Search_results.TextMatrix(0, 1)) * 100
Search_results.ColWidth(2) = "750"
Search_results.ColWidth(3) = "600"
Search_results.ColWidth(4) = "570"
Search_results.ColWidth(5) = Len(Search_results.TextMatrix(0, 5)) * 100
Search_results.ColWidth(0) = "2280"
End Sub

Private Sub UPL_Transfer_Bot_Timer()
For i = 24 To 47
If status(i).Text = "Queued" Then
 If curntdling = max_ups.Text Then
 Uploader(i).SendData "SAF"
 Exit Sub
 End If
agf = shared_folder(0).Text & fname(i).Text
FILESIZEY(i) = FileLen(agf)
Uploader(i).SendData "TFS" & FILESIZEY(i)
status(i).Text = "Uploading"
Exit Sub
End If
Next
End Sub


Private Sub Uploader_ConnectionRequest(Index As Integer, ByVal requestID As Long)
For i = 24 To 47
If fname(i).Text = "" Then
dube = i
GoTo desb
End If
Next

desb:
Uploader(dube).Accept requestID
End Sub

Private Sub Uploader_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim Packet As String
Uploader(Index).GetData Packet

'===Remote user sends request for file===
If Mid(Packet, 1, 3) = "IWT" Then
k = Mid(Packet, 4, Len(Packet) - 3)
j = InStr(1, k, ";")
filnamer = Mid(k, 1, j - 1)
usernamer = Mid(k, j + 1, Len(Packet) - j)

fname(Index).Text = filnamer
usr(Index).Text = usernamer
status(Index).Text = "Queued"
kbs(Index).Text = "0"
done(Index).Text = "0"
End If
'=========================================

If Mid(Packet, 1, 3) = "RSM" Then
af = Mid(Packet, 4, Len(Packet))
MsgBox af
End If
End Sub

Private Sub usercount_Timer()
usersval.Caption = List1.ListCount
End Sub
Sub Sleep(ByVal MillaSec As Long, Optional ByVal DeepSleep As Boolean = False)
    Dim tStart#, Tmr#
    tStart = Timer

    While Tmr < (MillaSec / 1000)
        Tmr = Timer - tStart
        If DeepSleep = False Then DoEvents
    Wend
End Sub
Sub loadiplist()
On Error GoTo errhand

Open "ipcfg.cfg" For Input As #1
List1.Clear
For i = 1 To LOF(1)
If Not EOF(1) Then
Line Input #1, hedgehog
If hedgehog <> "" Then List1.AddItem hedgehog
End If
Next
Close #1

errhand:
Exit Sub
saveiplist
End Sub
Sub saveiplist()
Open "ipcfg.cfg" For Output As #1
For i = 0 To List1.ListCount
Print #1, List1.List(i)
Next
Close #1
End Sub
Sub savepref()
Close #1
Open "ripset" For Output As #1
Write #1, shared_folder(0).Text
Write #1, shared_folder(1).Text
Write #1, ModemSpeed.Text
Write #1, username.Text
Write #1, max_dls.Text
Write #1, max_ups.Text
Write #1, up_on_start.Value
Close #1
End Sub
Sub loadpref()
On Error GoTo errhand

Open "ripset" For Input As #1
If Not EOF(1) Then Input #1, kot1
If Not EOF(1) Then Input #1, kot2
If Not EOF(1) Then Input #1, kot3
If Not EOF(1) Then Input #1, kot4
If Not EOF(1) Then Input #1, kot5
If Not EOF(1) Then Input #1, kot6
If Not EOF(1) Then Input #1, kot7
Close #1
 shared_folder(0).Text = kot1
 shared_folder(1).Text = kot2
 ModemSpeed.Text = kot3
 username.Text = kot4
 max_dls.Text = kot5
 max_ups.Text = kot6
up_on_start.Value = kot7
errhand:
Exit Sub
savepref
End Sub

Private Sub VScroll1_Change()
myfiles_frame2.Top = 0 - VScroll1.Value
myfiles_frame3.Top = 0 - VScroll1.Value + myfiles_frame2.Height
If myfiles_frame2.Top = -3960 Then myfiles_frame3.Top = 4680
End Sub

Private Sub VScroll2_Change()
downloads_frame.Top = 0 - VScroll2.Value
End Sub

Private Sub VScroll3_Change()
uploads_frme.Top = 0 - VScroll3.Value
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

End Sub

Private Sub Winsock2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

End Sub

Public Function DownloadFile(ByVal SockNumber As Integer)

End Function

Public Function DelMoveUpTrans(ByVal PosToDelandMove)
If PosToDelandMove < 24 Then
Downloader(PosToDelandMove).Close
For otr = PosToDelandMove To 22
fname(otr).Text = fname(otr - 1).Text
status(otr).Text = status(otr - 1).Text
usr(otr).Text = usr(otr - 1).Text
kbs(otr).Text = kbs(otr - 1).Text
done(otr).Text = done(otr - 1).Text
Next

fname(23).Text = ""
status(23).Text = ""
usr(23).Text = ""
kbs(23).Text = ""
done(23).Text = ""

End If

If PosToDelandMove > 23 Then
For otr = PosToDelandMove To 47
fname(otr).Text = fname(otr - 1).Text
status(otr).Text = status(otr - 1).Text
usr(otr).Text = usr(otr - 1).Text
kbs(otr).Text = kbs(otr - 1).Text
done(otr).Text = done(otr - 1).Text
Next
fname(47).Text = ""
status(47).Text = ""
usr(47).Text = ""
kbs(47).Text = ""
done(47).Text = ""

End If

End Function

Private Sub zort_Change()
If zort.Text > Form1.List1.ListCount - 1 Then
stopsearch_button_Click
zort = 0
End If
End Sub
Public Function ResumeFileSend(ByVal socketnumber, ByVal AtWhatByte)

End Function


Public Function StartFileSend(ByVal SockNumber)
'send the file for the first time
'NO RESUME!
Dim tmpbuff As String
sn = socketnumber
agf = shared_folder(0).Text & fname(sn).Text

Open agf For Binary Access Read As #filenumberg(sn)
tmpbuff = Space$(1024)
Get #filenumberg(sn), 1, tmpbuff
RateUP(sn) = RateUP(sn) + 1024
Uploader(sn).SendData "RAW" & tmpbuff

If 1024 > FILESIZEY(sn) Then
Sleep 2000
Uploader(sn).SendData "EOF"
End If
End Function
