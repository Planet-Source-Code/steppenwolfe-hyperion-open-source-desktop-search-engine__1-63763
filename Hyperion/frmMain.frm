VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   Caption         =   "Hyperion 1.0"
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13185
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9690
   ScaleWidth      =   13185
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picPanel 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3A282&
      ForeColor       =   &H80000008&
      Height          =   9435
      Index           =   2
      Left            =   30
      ScaleHeight     =   9405
      ScaleWidth      =   3345
      TabIndex        =   49
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
      Begin VB.PictureBox picOpt 
         Appearance      =   0  'Flat
         BackColor       =   &H00F7DFD6&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   60
         ScaleHeight     =   285
         ScaleWidth      =   3195
         TabIndex        =   101
         Top             =   6990
         Width           =   3225
         Begin VB.OptionButton optDisplay 
            BackColor       =   &H00F7DFD6&
            Caption         =   "Search"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   6
            Left            =   120
            TabIndex        =   104
            Top             =   30
            Value           =   -1  'True
            Width           =   825
         End
         Begin VB.OptionButton optDisplay 
            BackColor       =   &H00F7DFD6&
            Caption         =   "Filters"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   7
            Left            =   1080
            TabIndex        =   103
            Top             =   30
            Width           =   825
         End
         Begin VB.OptionButton optDisplay 
            BackColor       =   &H00F7DFD6&
            Caption         =   "Advanced"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   8
            Left            =   2010
            TabIndex        =   102
            Top             =   30
            Width           =   1125
         End
      End
      Begin VB.CheckBox chkTFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Windows\security"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   67
         Top             =   5355
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkTFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Windows\system32\CatRoot2"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   66
         Top             =   5040
         Value           =   1  'Checked
         Width           =   2475
      End
      Begin VB.CheckBox chkTFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Windows\system32\usmt"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   65
         Top             =   4725
         Value           =   1  'Checked
         Width           =   2355
      End
      Begin VB.CheckBox chkTFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Windows\system32"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   64
         Top             =   6615
         Width           =   1755
      End
      Begin VB.CheckBox chkTFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "System Volume Information"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   63
         Top             =   4410
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkTFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Windows\ServicePackFiles"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   62
         Top             =   5670
         Value           =   1  'Checked
         Width           =   2385
      End
      Begin VB.CheckBox chkTFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Windows\SoftwareDistribution"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   61
         Top             =   5985
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox chkTFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Application Data\Microsoft"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   60
         Top             =   6300
         Value           =   1  'Checked
         Width           =   2235
      End
      Begin VB.CheckBox chkStart 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Start With Windows"
         ForeColor       =   &H00404040&
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   59
         Top             =   750
         Width           =   1965
      End
      Begin VB.CheckBox chkStart 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Start Indexing in the Background"
         ForeColor       =   &H00404040&
         Height          =   225
         Index           =   1
         Left            =   210
         TabIndex        =   58
         Top             =   1050
         Width           =   2805
      End
      Begin VB.CheckBox chkStart 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Run From System Tray"
         ForeColor       =   &H00404040&
         Height          =   225
         Index           =   2
         Left            =   210
         TabIndex        =   57
         Top             =   1350
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkIndex 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Enable Automatic Reindexing"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   210
         TabIndex        =   56
         Top             =   2220
         Width           =   2415
      End
      Begin VB.PictureBox picOptions 
         Appearance      =   0  'Flat
         BackColor       =   &H00F7DFD6&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   150
         ScaleHeight     =   1215
         ScaleWidth      =   2925
         TabIndex        =   50
         Top             =   2550
         Width           =   2925
         Begin VB.ComboBox cbDays 
            Height          =   315
            ItemData        =   "frmMain.frx":1982
            Left            =   840
            List            =   "frmMain.frx":1984
            TabIndex        =   51
            Text            =   "Combo1"
            Top             =   810
            Width           =   1485
         End
         Begin VB.OptionButton optIndex 
            BackColor       =   &H00F7DFD6&
            Caption         =   "24 Hours"
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   1
            Left            =   30
            TabIndex        =   54
            Top             =   570
            Value           =   -1  'True
            Width           =   1005
         End
         Begin VB.OptionButton optIndex 
            BackColor       =   &H00F7DFD6&
            Caption         =   "12 Hours"
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   0
            Left            =   30
            TabIndex        =   53
            Top             =   240
            Width           =   1005
         End
         Begin VB.OptionButton optIndex 
            BackColor       =   &H00F7DFD6&
            Caption         =   "Every:"
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   2
            Left            =   30
            TabIndex        =   52
            Top             =   900
            Width           =   885
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule:"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   55
            Top             =   0
            Width           =   705
         End
      End
      Begin VB.Label lblHeader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exclude Paths in Search"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   68
         Top             =   4110
         Width           =   1995
      End
      Begin VB.Label lblCompression 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Indexing Options"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   69
         Top             =   1950
         Width           =   1440
      End
      Begin VB.Label lblRePnl 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Startup Options"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   70
         Top             =   450
         Width           =   1335
      End
      Begin VB.Label lblControls 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Advanced Options"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   71
         Top             =   90
         Width           =   1530
      End
      Begin VB.Shape spBorder 
         BackColor       =   &H00F7DFD6&
         BackStyle       =   1  'Opaque
         Height          =   2955
         Index           =   0
         Left            =   60
         Top             =   3990
         Width           =   3225
      End
      Begin VB.Shape spBorder 
         BackColor       =   &H00F7DFD6&
         BackStyle       =   1  'Opaque
         Height          =   1425
         Index           =   1
         Left            =   60
         Top             =   360
         Width           =   3225
      End
      Begin VB.Shape spBorder 
         BackColor       =   &H00F7DFD6&
         BackStyle       =   1  'Opaque
         Height          =   2115
         Index           =   2
         Left            =   60
         Top             =   1830
         Width           =   3225
      End
   End
   Begin VB.PictureBox picPanel 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3A282&
      ForeColor       =   &H80000008&
      Height          =   9435
      Index           =   1
      Left            =   0
      ScaleHeight     =   9405
      ScaleWidth      =   3345
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
      Begin VB.PictureBox picOpt 
         Appearance      =   0  'Flat
         BackColor       =   &H00F7DFD6&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   60
         ScaleHeight     =   285
         ScaleWidth      =   3195
         TabIndex        =   97
         Top             =   9030
         Width           =   3225
         Begin VB.OptionButton optDisplay 
            BackColor       =   &H00F7DFD6&
            Caption         =   "Search"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   3
            Left            =   120
            TabIndex        =   100
            Top             =   30
            Value           =   -1  'True
            Width           =   825
         End
         Begin VB.OptionButton optDisplay 
            BackColor       =   &H00F7DFD6&
            Caption         =   "Filters"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   4
            Left            =   1080
            TabIndex        =   99
            Top             =   30
            Width           =   825
         End
         Begin VB.OptionButton optDisplay 
            BackColor       =   &H00F7DFD6&
            Caption         =   "Advanced"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   5
            Left            =   2010
            TabIndex        =   98
            Top             =   30
            Width           =   1125
         End
      End
      Begin VB.CheckBox chkIFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Movies: mpeg;wmv;avi"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   12
         Left            =   210
         TabIndex        =   91
         Top             =   8670
         Value           =   1  'Checked
         Width           =   2025
      End
      Begin VB.CheckBox chkIFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "WBMP: wbmp;wbm"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   11
         Left            =   210
         TabIndex        =   90
         Top             =   8430
         Width           =   1695
      End
      Begin VB.CheckBox chkIFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "TIFF: tif;tiff"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   10
         Left            =   210
         TabIndex        =   89
         Top             =   8190
         Width           =   1215
      End
      Begin VB.CheckBox chkIFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Targa: tga"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   9
         Left            =   210
         TabIndex        =   88
         Top             =   7950
         Width           =   1125
      End
      Begin VB.CheckBox chkIFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Photoshop: psd"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   8
         Left            =   210
         TabIndex        =   87
         Top             =   7710
         Width           =   1515
      End
      Begin VB.CheckBox chkIFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "JPEG: jpg;jpeg"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   7
         Left            =   210
         TabIndex        =   86
         Top             =   7470
         Value           =   1  'Checked
         Width           =   1425
      End
      Begin VB.CheckBox chkIFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "GIF: gif"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   6
         Left            =   210
         TabIndex        =   85
         Top             =   7230
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.CheckBox chkIFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Freehand: fh9;fh10;fh11;ft7;ft8"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   5
         Left            =   210
         TabIndex        =   84
         Top             =   6990
         Width           =   2655
      End
      Begin VB.CheckBox chkIFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Fireworks: png"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   83
         Top             =   6750
         Width           =   1425
      End
      Begin VB.CheckBox chkIFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "EPS: eps"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   82
         Top             =   6510
         Width           =   975
      End
      Begin VB.CheckBox chkIFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Corel Draw: cdr"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   81
         Top             =   6270
         Width           =   1455
      End
      Begin VB.CheckBox chkIFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "BMP: bmp;dib;rle"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   80
         Top             =   6030
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkIFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Adobe Illustrator: ai;art"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   79
         Top             =   5310
         Width           =   2055
      End
      Begin VB.CheckBox chkMFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "MP3: mp3; mp2"
         ForeColor       =   &H00404040&
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   77
         Top             =   3750
         Value           =   1  'Checked
         Width           =   2025
      End
      Begin VB.CheckBox chkMFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Wav: wav; snd; au; aif"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   76
         Top             =   4020
         Value           =   1  'Checked
         Width           =   2115
      End
      Begin VB.CheckBox chkMFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Media Playlist: asx; wax; m3u; wvl"
         ForeColor       =   &H00404040&
         Height          =   225
         Index           =   2
         Left            =   210
         TabIndex        =   75
         Top             =   4260
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox chkMFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Microsoft Recorded TV: dvr-ms"
         ForeColor       =   &H00404040&
         Height          =   225
         Index           =   3
         Left            =   210
         TabIndex        =   74
         Top             =   4800
         Width           =   2745
      End
      Begin VB.CheckBox chkMFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Midi: mid; rmi; midi"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   73
         Top             =   4530
         Width           =   2085
      End
      Begin VB.CheckBox chkMFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Windows Media: asf; wm; wma; wmv"
         ForeColor       =   &H00404040&
         Height          =   225
         Index           =   5
         Left            =   210
         TabIndex        =   72
         Top             =   5040
         Width           =   2955
      End
      Begin VB.CheckBox chkDocFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Adobe Document: pdf"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   46
         Top             =   390
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkDFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "ASCII: txt"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   45
         Top             =   660
         Value           =   1  'Checked
         Width           =   1125
      End
      Begin VB.CheckBox chkDFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Excel: xls;xlw"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   44
         Top             =   930
         Width           =   1335
      End
      Begin VB.CheckBox chkDFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "HTML: htm;html;xhtm;xhtml;cfm"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   43
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CheckBox chkDFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Lotus 1-2-3: wk1;wk3;wk4"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   42
         Top             =   1470
         Width           =   2265
      End
      Begin VB.CheckBox chkDFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Rich Text: rtf"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   5
         Left            =   210
         TabIndex        =   41
         Top             =   1740
         Value           =   1  'Checked
         Width           =   1395
      End
      Begin VB.CheckBox chkDFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Extended Text: ans"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   6
         Left            =   210
         TabIndex        =   40
         Top             =   2010
         Width           =   1845
      End
      Begin VB.CheckBox chkDFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Word: doc;dot"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   7
         Left            =   210
         TabIndex        =   39
         Top             =   2280
         Value           =   1  'Checked
         Width           =   1395
      End
      Begin VB.CheckBox chkDFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Word for Mac: mcw"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   8
         Left            =   210
         TabIndex        =   38
         Top             =   2550
         Width           =   1785
      End
      Begin VB.CheckBox chkDFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Microsoft Works: wps"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   9
         Left            =   210
         TabIndex        =   37
         Top             =   2820
         Width           =   1905
      End
      Begin VB.CheckBox chkDFilter 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Write: wri"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   10
         Left            =   210
         TabIndex        =   36
         Top             =   3090
         Width           =   1095
      End
      Begin VB.Label lblHeader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Image Filters"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   9
         Left            =   210
         TabIndex        =   92
         Top             =   5760
         Width           =   1125
      End
      Begin VB.Shape spBorder 
         BackColor       =   &H00F7DFD6&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         Height          =   3285
         Index           =   4
         Left            =   60
         Top             =   5700
         Width           =   3225
      End
      Begin VB.Label lblHeader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Music Filters"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   8
         Left            =   210
         TabIndex        =   78
         Top             =   3510
         Width           =   1050
      End
      Begin VB.Shape spBorder 
         BackColor       =   &H00F7DFD6&
         BackStyle       =   1  'Opaque
         Height          =   2175
         Index           =   5
         Left            =   60
         Top             =   3450
         Width           =   3225
      End
      Begin VB.Label lblControls 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Document Filters"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   48
         Top             =   120
         Width           =   1440
      End
      Begin VB.Shape spBorder 
         BackColor       =   &H00F7DFD6&
         BackStyle       =   1  'Opaque
         Height          =   3345
         Index           =   3
         Left            =   60
         Top             =   60
         Width           =   3225
      End
      Begin VB.Label lblHeader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Filters"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   270
         TabIndex        =   47
         Top             =   630
         Width           =   1155
      End
   End
   Begin VB.PictureBox picPanel 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3A282&
      ForeColor       =   &H80000008&
      Height          =   9435
      Index           =   0
      Left            =   30
      ScaleHeight     =   9405
      ScaleWidth      =   3345
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.PictureBox picOpt 
         Appearance      =   0  'Flat
         BackColor       =   &H00F7DFD6&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   60
         ScaleHeight     =   285
         ScaleWidth      =   3195
         TabIndex        =   93
         Top             =   8520
         Width           =   3225
         Begin VB.OptionButton optDisplay 
            BackColor       =   &H00F7DFD6&
            Caption         =   "Advanced"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   2
            Left            =   2010
            TabIndex        =   96
            Top             =   30
            Width           =   1125
         End
         Begin VB.OptionButton optDisplay 
            BackColor       =   &H00F7DFD6&
            Caption         =   "Filters"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   1
            Left            =   1080
            TabIndex        =   95
            Top             =   30
            Width           =   825
         End
         Begin VB.OptionButton optDisplay 
            BackColor       =   &H00F7DFD6&
            Caption         =   "Search"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   94
            Top             =   30
            Value           =   -1  'True
            Width           =   825
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F7DFD6&
         ForeColor       =   &H80000008&
         Height          =   3165
         Left            =   60
         ScaleHeight     =   3135
         ScaleWidth      =   3195
         TabIndex        =   22
         Top             =   60
         Width           =   3225
         Begin VB.CommandButton cmdControls 
            BackColor       =   &H00F7DFD6&
            Caption         =   "Search >"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2040
            TabIndex        =   34
            Top             =   2670
            Width           =   1035
         End
         Begin VB.DriveListBox Drive1 
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   120
            TabIndex        =   29
            Top             =   270
            Width           =   2325
         End
         Begin VB.CommandButton cmdPath 
            BackColor       =   &H00FFFFFF&
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2790
            TabIndex        =   28
            Top             =   1410
            Width           =   315
         End
         Begin VB.TextBox txtPath 
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   120
            TabIndex        =   27
            Top             =   1380
            Width           =   2595
         End
         Begin VB.ComboBox cboFileType 
            ForeColor       =   &H00404040&
            Height          =   315
            ItemData        =   "frmMain.frx":1986
            Left            =   120
            List            =   "frmMain.frx":1999
            TabIndex        =   26
            Text            =   "cboFileType"
            Top             =   840
            Width           =   2325
         End
         Begin VB.TextBox txtString 
            Height          =   315
            Left            =   120
            TabIndex        =   25
            Top             =   1950
            Width           =   2955
         End
         Begin VB.OptionButton optSearchType 
            BackColor       =   &H00F7DFD6&
            Caption         =   "Exact Name"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   24
            Top             =   2310
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton optSearchType 
            BackColor       =   &H00F7DFD6&
            Caption         =   "Sounds Like"
            Height          =   225
            Index           =   1
            Left            =   1380
            TabIndex        =   23
            Top             =   2310
            Width           =   1245
         End
         Begin VB.Label lblHeader 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "All Files of Type:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   4
            Left            =   150
            TabIndex        =   33
            Top             =   660
            Width           =   1350
         End
         Begin VB.Label lblHeader 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Path:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   3
            Left            =   150
            TabIndex        =   32
            Top             =   1200
            Width           =   435
         End
         Begin VB.Label lblHeader 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Drive:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   31
            Top             =   90
            Width           =   495
         End
         Begin VB.Label lblHeader 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Search For:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   5
            Left            =   150
            TabIndex        =   30
            Top             =   1770
            Width           =   945
         End
      End
      Begin VB.PictureBox pnlModified 
         Appearance      =   0  'Flat
         BackColor       =   &H00F7DFD6&
         ForeColor       =   &H80000008&
         Height          =   2985
         Left            =   60
         ScaleHeight     =   2955
         ScaleWidth      =   3195
         TabIndex        =   10
         Top             =   3270
         Width           =   3225
         Begin VB.TextBox txtTo 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   840
            TabIndex        =   18
            Top             =   2490
            Width           =   1905
         End
         Begin VB.TextBox txtFrom 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   840
            TabIndex        =   17
            Top             =   2100
            Width           =   1905
         End
         Begin VB.OptionButton optModified 
            BackColor       =   &H00F7DFD6&
            Caption         =   "Specify Dates"
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   16
            Top             =   1320
            Width           =   2025
         End
         Begin VB.OptionButton optModified 
            BackColor       =   &H00F7DFD6&
            Caption         =   "Within the last year"
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   15
            Top             =   1080
            Width           =   2265
         End
         Begin VB.OptionButton optModified 
            BackColor       =   &H00F7DFD6&
            Caption         =   "Past month"
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   1785
         End
         Begin VB.OptionButton optModified 
            BackColor       =   &H00F7DFD6&
            Caption         =   "Within the last week"
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Width           =   2325
         End
         Begin VB.OptionButton optModified 
            BackColor       =   &H00F7DFD6&
            Caption         =   "Don't Remember"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Value           =   -1  'True
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker dtDateType 
            Height          =   345
            Left            =   150
            TabIndex        =   11
            Top             =   1590
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   22937601
            CurrentDate     =   38699
         End
         Begin VB.Label lblHeader 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "File Last Accessed"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   90
            Width           =   1530
         End
         Begin VB.Label lblTo 
            BackStyle       =   0  'Transparent
            Caption         =   "To:"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   420
            TabIndex        =   20
            Top             =   2490
            Width           =   285
         End
         Begin VB.Label lblFrom 
            BackStyle       =   0  'Transparent
            Caption         =   "From:"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   270
            TabIndex        =   19
            Top             =   2190
            Width           =   465
         End
      End
      Begin VB.PictureBox pnlSize 
         Appearance      =   0  'Flat
         BackColor       =   &H00F7DFD6&
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   60
         ScaleHeight     =   2145
         ScaleWidth      =   3195
         TabIndex        =   1
         Top             =   6300
         Width           =   3225
         Begin VB.OptionButton optSize 
            BackColor       =   &H00F7DFD6&
            Caption         =   "Don't Remember"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Value           =   -1  'True
            Width           =   2025
         End
         Begin VB.OptionButton optSize 
            BackColor       =   &H00F7DFD6&
            Caption         =   "Small (less than 100KB)"
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   2445
         End
         Begin VB.OptionButton optSize 
            BackColor       =   &H00F7DFD6&
            Caption         =   "Medium (less than 1MB)"
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   6
            Top             =   840
            Width           =   2535
         End
         Begin VB.OptionButton optSize 
            BackColor       =   &H00F7DFD6&
            Caption         =   "Large (more than 1MB)"
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   5
            Top             =   1080
            Width           =   2505
         End
         Begin VB.OptionButton optSize 
            BackColor       =   &H00F7DFD6&
            Caption         =   "Specify size (in MB)"
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   4
            Top             =   1320
            Width           =   2325
         End
         Begin VB.ComboBox cboSize 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            ForeColor       =   &H00404040&
            Height          =   315
            ItemData        =   "frmMain.frx":19D2
            Left            =   150
            List            =   "frmMain.frx":19DD
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1620
            Width           =   1155
         End
         Begin VB.TextBox txtSize 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   1410
            TabIndex        =   2
            Text            =   "0"
            Top             =   1620
            Width           =   1005
         End
         Begin VB.Label lblHeader 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "File Size"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   90
            Width           =   675
         End
      End
   End
   Begin MSComctlLib.ImageList ilsList 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":71E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D480
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FC32
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10084
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10B4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13300
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13752
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13BA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13FF6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstResults 
      Height          =   9435
      Left            =   3390
      TabIndex        =   105
      Top             =   0
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   16642
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ilsList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   106
      Tag             =   "No"
      Top             =   9435
      Width           =   13185
      _ExtentX        =   23257
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin VB.Menu mnuMain 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuMain1 
         Caption         =   "Jump to Location"
         Index           =   0
      End
      Begin VB.Menu mnuMain1 
         Caption         =   "File Properties"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*~ Hey people.. Steppenwolfe gives you an early Christmas present.. (hooray!)
'*~ Hope you like this latest, (it was a lot of work!)
'*~ This is (of course), is an Alpha of a freeware search engine I am working on.
'*~ It will probably take me another week or so to get it all straightened out
'*~ so bear with me here.. (ps: if you do find a bug, please, just email me at
'*~ steppenwolfe_2000@yahoo.com and be as descriptive as possible ie where/when/what..
'*~ error numbers/routine name would be nice too..)
'*~ It has been tested on XP SP2, and works fine on this platform, ME/98 users, let
'*~ me know if you have issues.. (Note: always end app with main form close button first to unload engine..)
'*~ Anyhow, that's it for now, have a good holiday, and enjoy..
'*~ John

'/* search engine
Private WithEvents cEngine                  As clsEngine
Attribute cEngine.VB_VarHelpID = -1

'/* file time structs
Private Type FT
    lLD                                     As Long
    lHD                                     As Long
End Type

Private Type BHFI
    lFA                                     As Long
    fCT                                     As FT
    fLA                                     As FT
    fLWT                                    As FT
    lVSN                                    As Long
    lFSH                                    As Long
    lFSL                                    As Long
    lNOL                                    As Long
    lFIH                                    As Long
    lFIL                                    As Long
End Type

'/* time conversion struct
Private Type TIME
    wYR                                     As Integer
    wMNT                                    As Integer
    wDOW                                    As Integer
    wDAY                                    As Integer
    wHR                                     As Integer
    wMIN                                    As Integer
    wSEC                                    As Integer
    wMSC                                    As Integer
End Type

'/* timer all input
Private Const QS_ALLINPUT As Double = _
(&H1 Or &H2 Or &H4 Or &H8 Or &H10 Or _
&H20 Or &H40 Or &H80)

'/* collections/list
Private cFilter                             As New Collection
Private cLocal                              As New Collection
Private m_lItem                             As ListItem

'/* browse for folder constants
Private Const BIF_RETURNONLYFSDIRS          As Integer = 1
Private Const BIF_DONTGOBELOWDOMAIN         As Integer = 2
Private Const MAX_PATH                      As Integer = 260

'/* directory
Private Type BrowseInfo
    hwndOwner                               As Long
    pIDLRoot                                As Long
    pszDisplayName                          As Long
    lpszTitle                               As Long
    ulFlags                                 As Long
    lpfnCallback                            As Long
    lParam                                  As Long
    iImage                                  As Long
End Type

'/* directory api
'/* lstrcat is very slow in vb - avoid it
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, _
                                                                  ByVal lpString2 As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, _
                                                            ByVal lpBuffer As String) As Long
'/* file info api
Private Declare Function GetFileInformationByHandle Lib "kernel32" (ByVal lHD As Long, _
                                                                    lFI As BHFI) As Long

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lFN As String, _
                                                                        ByVal lDA As Long, _
                                                                        ByVal lSM As Long, _
                                                                        lSA As Any, _
                                                                        ByVal lCD As Long, _
                                                                        ByVal lFA As Long, _
                                                                        ByVal lHD As Long) As Long

Private Declare Function FileTimeToSystemTime Lib "kernel32.dll" (lpFileTime As FT, _
                                                                  lST As TIME) As Long

Private Declare Function FileTimeToLocalFileTime Lib "kernel32.dll" (lFT As FT, _
                                                                     lLFT As FT) As Long


Private Declare Function CloseHandle Lib "kernel32" (ByVal lHD As Long) As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, _
                                                                          ByVal lpClassName As String, _
                                                                          ByVal nMaxCount As Long) As Long

Private Declare Function MsgWaitForMultipleObjects Lib "user32" (ByVal nCount As Long, _
                                                                 pHandles As Long, _
                                                                 ByVal fWaitAll As Long, _
                                                                 ByVal dwMilliseconds As Long, _
                                                                 ByVal dwWakeMask As Long) As Long
                                                                 
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function LockWindowUpdate Lib "user32" (ByVal lhwnd As Long) As Long

Private m_lLoopCounter      As Long
Private m_bUpdateStatus     As Boolean
Private m_bQuickExit        As Boolean
Private FileInfo            As BHFI
Private m_tTime             As Double
Private m_lCounter          As Long
Private m_bChange           As Boolean
Private m_bSaveIndex        As Boolean


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                           ENGINE EVENTS
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


'/* search completed
Private Sub cEngine_eEComplete()

Debug.Print "processing complete"

    If frmQuery.Visible = True Then
        With frmQuery
            .lblInfo(1).Caption = "Scan Complete!"
            .cmdSet(1).Visible = True
            .cmdSet(0).Visible = False
            .cmdSet(1).Visible = True
            .lblInfo(1).Visible = True
            .lblInfo(1).Caption = "Index for Drive: " & left$(Drive1.Drive, 2) & Chr$(92) & " loaded."
        End With
        Unload frmStatus
    End If
    
    frmSysTray.mnuMain.Enabled = True
    cmdControls.Enabled = True

End Sub

'/* progress counter tick
Private Sub cEngine_eECount()

On Error Resume Next

    With frmStatus
        .pbStatus.Value = .pbStatus.Value + 1
        .lblStatus(1).Caption = Format$((.pbStatus.Value / .pbStatus.Max) * 100, "#0.0") & "%"
    End With

On Error GoTo 0

End Sub

'/* progress counter max
Private Sub cEngine_eECountMax(lMax As Long)

On Error Resume Next

    frmStatus.pbStatus.Max = lMax

On Error GoTo 0

End Sub

'/* index has been saved
Private Sub cEngine_eEDump()

Debug.Print "index saved"

End Sub

'/* engine processing state
Private Sub cEngine_eEEngaged()

Debug.Print "index engine engaged"
    frmSysTray.mnuMain.Enabled = False

End Sub

'/* index loaded
Private Sub cEngine_eERestore()

Debug.Print "index loaded"

    stBar.Panels(1).Text = "Index for Drive: " & left$(Drive1.Drive, 2) & Chr$(92) & " loaded."

End Sub

'/* index has loaded
Private Sub cEngine_eEBuild()

Debug.Print "index has been built"
    
    stBar.Panels(1).Text = "Index for Drive: " & left$(Drive1.Drive, 2) & Chr$(92) & " loaded."
    
End Sub

'/* returns index status (not used)
Private Sub cEngine_eEIndStatus(bState As Boolean)

    Debug.Print "index status is: " & CStr(bState)
    
End Sub

'/* update loop completed
Private Sub cEngine_eEMultiTask()

    Debug.Print "multi task operations complete"
    
    m_bUpdateStatus = False
    
End Sub

'/* pattern search complete
Private Sub cEngine_eEPatternComplete()

Dim vItem As Variant

On Error Resume Next
    Debug.Print "pattern search complete"
    '/* transfer to local list
    Set cLocal = cEngine.p_CReturn
    '/* search for a file type
    For Each vItem In cLocal
        Filter_Check CStr(vItem)
    Next vItem
    
    '/* status update
    stBar.Panels(1).Text = "Scan Complete! Scan Time: " & _
    Format$(Timer - m_tTime, "#0.0000") & " Seconds.. Found: " & lstResults.ListItems.Count
    LockWindowUpdate &H0
    m_lCounter = 0
    cmdControls.Enabled = True

On Error GoTo 0

End Sub

'/* match search complete
Private Sub cEngine_eEProcessComplete()

Dim vItem As Variant

On Error Resume Next

    Debug.Print "process search complete"
    '/* transfer to local list
    Set cLocal = cEngine.p_CReturn
    '/* user queue
    stBar.Panels(1).Text = "Scanning for files.."
    '/* send to filter for evaluation
    For Each vItem In cLocal
        Filter_Check CStr(vItem)
    Next vItem
    
    '/* status bar
    stBar.Panels(1).Text = "Scan Complete! Scan Time: " & _
    Format$(Timer - m_tTime, "#0.0000") & " Seconds.. Found: " & lstResults.ListItems.Count
    LockWindowUpdate &H0
    m_lCounter = 0
    cmdControls.Enabled = True

On Error GoTo 0

End Sub


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                           ENGINE WORKER CALLS
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Private Sub cmdControls_Click()
'/* values check

Dim vItem   As Variant

On Error Resume Next

    '/* timers/counters
    m_tTime = 0
    m_tTime = Timer
    m_lCounter = 0
    
    '/* path test
    If Len(txtPath.Text) = 0 Then
        txtPath.Text = left$(Drive1.Drive, 2) & Chr$(92)
    End If
    If Len(txtString.Text) = 0 Then Exit Sub
    cmdControls.Enabled = False
    
    '/* get search item(s)
    Set cLocal = Nothing
    Set cLocal = New Collection
    '/* lock updates
    LockWindowUpdate lstResults.hwnd
    '/* set status queue
    stBar.Panels(1).Text = "Scanning for files.."
    '/* clear list
    lstResults.ListItems.Clear
        
    '/* search for a phrase
    Search_Process txtString.Text
    '/* transfer search list
    With cEngine
        '/* clear engine
        .Engine_Reload
        '/* scan path
        .p_BuildPath = Trim$(txtPath.Text)
        '/* add search items
        Set .p_CForward = cLocal
            '/* search type
            If optSearchType(0).Value Then
                .p_EngineTask = Search_Exact
            Else
                .p_EngineTask = Search_Pattern
            End If
            .Start
    End With

On Error GoTo 0

End Sub

Private Sub Search_Process(ByVal sSearch As String)
'/* add search item(s) to collection
'/* tests option for multiple search
'/* items seperated by semicolon

Dim aResults() As String
Dim i          As Integer

    '/* test for wildcard and add to collection
    If Not InStr(1, sSearch, Chr$(59)) Then
        cLocal.Add sSearch
    Else
        aResults = Split(sSearch, Chr$(59))
        For i = 0 To UBound(aResults)
            If Len(aResults(i)) > 0 Then
                cLocal.Add Trim$(aResults(i))
            End If
        Next i
    End If

End Sub

Public Sub Index_Start(ByVal sPath As String)
'/* engine index builder

Dim sIndPath      As String

    sPath = left$(Drive1.Drive, 2) & Chr$(92)
    frmSysTray.mnuMain1(0).Enabled = False
    frmSysTray.mnuMain1(1).Enabled = False
    frmSysTray.mnuMain1(3).Enabled = False
    sIndPath = App.Path & "\index" & CInt(Drive1.ListIndex) & ".dat"
    
    '/* reset the engine
    Set cEngine = Nothing
    Set cEngine = New clsEngine
    
    '/* flag wait timer
    m_bUpdateStatus = True
    With cEngine
        '/* set scan path
        .p_BuildPath = sPath
        '/* set index path
        .p_IndexPath = sIndPath
        '/* build index flag
        .p_CMultiTask.Add 1
        '/* save index flag
        .p_CMultiTask.Add 2
        '/* set to multi task mode
        .p_EngineTask = Engine_MultiTask
        '/* start engine
        .Start
        '/* wait for callback
        Status_Loop
        '/* just in case..
        m_bUpdateStatus = False
    End With
    
    Me.SetFocus

End Sub


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                           FORM CONTROLS
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Private Sub Form_Load()

On Error Resume Next

    '/* preload and instantiate objects
    Set_Options
    
    '*/ load engine
    Set cEngine = New clsEngine
    '*/ search objects
    Set cFilter = New Collection

    '/* loaded switch - running
    With New clsLightning
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "loaded", 2
    End With
    
    '/* restore user settings
    Get_Settings
    
    '/* load systray
    If chkStart(2).Value = 1 Then
        Load frmSysTray
    End If

On Error GoTo 0

End Sub

Private Sub Form_Resize()

Dim pic As PictureBox

On Error Resume Next
    
    For Each pic In picPanel
        pic.Height = (Me.ScaleHeight - stBar.Height)
        pic.left = 0
        pic.top = 0
    Next

    With lstResults
        .Height = (Me.ScaleHeight - stBar.Height)
        .top = 0
        .left = (picPanel(0).Width)
        .Width = Me.ScaleWidth - picPanel(0).Width
        .ColumnHeaders.Clear
        .AllowColumnReorder = True
        .ColumnHeaders.Add , , "Name", .Width / 8
        .ColumnHeaders.Add , , "Path", (.Width / 8) * 3
        .ColumnHeaders.Add , , "Accessed", .Width / 8
        .ColumnHeaders.Add , , "Created", .Width / 8
        .ColumnHeaders.Add , , "Modified", (.Width / 8)
        .ColumnHeaders.Add , , "Size", (.Width / 8) - 70
    End With

On Error GoTo 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim sMessage    As String

    '/* save settings
    Save_Settings
    
    '/* save loaded state
    If chkStart(2).Value Then
        '/* in tray
        With New clsLightning
            .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "loaded", 3
        End With
    Else
        '/* ending session
        With New clsLightning
            .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "loaded", 1
        End With
    End If
    
    '/* if reindex was aborted, notify
    '/* of impending index
    With New clsLightning
        If .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "schreset") = 1 Then
            Me.Hide
            '/* send start update notification
            sNotify 1
        End If
    End With
    
    '/* unload engine
    Set cEngine = Nothing

End Sub

Public Sub Progress_Load()
'/* load progress form

    With frmStatus
        .lblStatus(0).Caption = "Indexing Drive: " & left$(Drive1.Drive, 2) & Chr$(92)
        .Show vbModeless, frmQuery
    End With
    
End Sub

Private Sub chkIndex_Click()
'/* enable auto indexing

    If chkIndex.Value Then
        If chkIndex.Value Then
            Select Case True
            Case optIndex(0).Value
                Schedule_Add 1
            Case optIndex(1).Value
                Schedule_Add 2
            Case optIndex(2).Value
                Schedule_Add 3, cbDays.ListIndex
            End Select
        End If
    Else
        Schedule_Remove
    End If
    
End Sub

Private Sub chkStart_Click(Index As Integer)

    Select Case Index
    Case 0
        '/* run at startup
        If chkStart(0).Value Then
            With New clsLightning
                .Write_String HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", _
                "hyperion", App.Path & Chr$(92) & App.EXEName & ".exe -r"
            End With
        Else
            '/* remove switch
            With New clsLightning
                .Delete_Value HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", _
                "hyperion"
            End With
        End If
        
    Case 1
        '/* silent mode
        If chkStart(1).Value Then
            With New clsLightning
                .Write_String HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", _
                "hyperion", App.Path & Chr$(92) & App.EXEName & ".exe -s"
                
            End With
        Else
            '/* remove switch
            With New clsLightning
                .Delete_Value HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", _
                "hyperion"
            End With
        End If
    
    Case 2
        '/* load/unload systray
        frmQuery.Drive_Reflect
        If chkStart(2).Value Then
            Load frmSysTray
        Else
            Unload frmSysTray
        End If

    End Select
    
End Sub

Public Sub Drive_Reflect()

On Error Resume Next

    m_bChange = True
    Me.Drive1.ListIndex = frmQuery.Drive1.ListIndex
    Me.chkStart(2).Value = frmQuery.chkStart.Value
    m_bChange = False

On Error GoTo 0

End Sub

Private Sub Set_Options()
'/* set control options

    cboSize.ListIndex = 0
    '/* panel caption
    With stBar
        .Panels.Add Text:="Idle.."
        .Panels(1).Width = Me.Width
    End With
    '/* sched days
    With cbDays
        .AddItem "Monday", 0
        .AddItem "Tuesday", 1
        .AddItem "Wednesday", 2
        .AddItem "Thursday", 3
        .AddItem "Friday", 4
        .AddItem "Saturday", 5
        .AddItem "Sunday", 6
        .ListIndex = 0
    End With
    
    '/* search path
    txtPath.Text = left$(Drive1.Drive, 2) & Chr$(92)
    cboFileType.ListIndex = 0
    '/* setup listview
    List_Init
    
End Sub

Private Sub cmdIndex_Click()
    frmQuery.Show
End Sub

Private Sub dtDateType_CloseUp()
'/* set search time params

    If Len(txtFrom.Text) = 0 Then
        txtFrom.Text = dtDateType.Value
    Else
        txtTo.Text = dtDateType.Value
    End If
    
End Sub

Private Sub cmdPath_Click()
'/* folder selection routine

Dim lList       As Long
Dim sTitle      As String
Dim tBrowseInfo As BrowseInfo
Dim sBuffer     As String

On Error Resume Next
    
    '/* title
    sTitle = "Select a Directory to Scan: "
    '/* fill struct
    With tBrowseInfo
        .hwndOwner = Me.hwnd
        .lpszTitle = lstrcat(sTitle, vbNullString)
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With

    lList = SHBrowseForFolder(tBrowseInfo)
    '/* call dialog
    If lList Then
        sBuffer = Space$(MAX_PATH)
        SHGetPathFromIDList lList, sBuffer
        sBuffer = left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        txtPath.Text = sBuffer & Chr$(92)
        txtPath.SetFocus
    End If

On Error GoTo 0

End Sub

Private Sub Drive1_Change()
'/* test index and ask if new
'/* drive selection should be
'/* indexed

Dim sIndPath As String

On Error Resume Next

    sIndPath = App.Path & "\index" & CInt(Drive1.ListIndex) & ".dat"
    txtPath.Text = left$(Drive1.Drive, 2) & Chr$(92)
    '/* loader switch
    If m_bChange Then Exit Sub
    
    cmdControls.Enabled = False
    '/* test for index, if exists, load
    '/* if not, invoke indexing dialog
    If File_Exists(sIndPath) Then
        stBar.Panels(1).Text = "Loading Index for Drive: " & left$(Drive1.Drive, 2) & Chr$(92)
        cmdControls.Enabled = False
        
        '/* fetch the index
        With cEngine
            .p_IndexPath = sIndPath
            .p_EngineTask = Index_Restore
            .Start
        End With
    Else
        '/* hand it off to frmQuery
        With frmQuery
            .Drive_Reflect
            .cmdSet(0).Visible = True
            .cmdSet(0).Enabled = True
            .cmdSet(1).Visible = False
            .lblInfo(1).Caption = "To start Indexing this Drive now, choose 'Start Indexing'"
            .Show vbModeless, Me
        End With
    End If

On Error GoTo 0

End Sub

Private Sub List_Init()
'/* set up listview

    With lstResults
        '.HeaderButtons = False
        .ColumnHeaders.Clear
        .AllowColumnReorder = True
        .View = lvwReport
        .ColumnHeaders.Add , , "Name", .Width / 8
        .ColumnHeaders.Add , , "Path", (.Width / 8) * 3
        .ColumnHeaders.Add , , "Accessed", .Width / 8
        .ColumnHeaders.Add , , "Created", .Width / 8
        .ColumnHeaders.Add , , "Modified", (.Width / 8)
        .ColumnHeaders.Add , , "Size", (.Width / 8) - 70
        .SmallIcons = Me.ilsList
        .FullRowSelect = True
    End With

End Sub

Private Sub optDisplay_Click(Index As Integer)
'/* option panels show/hide

Dim pic As PictureBox

    For Each pic In picPanel
        pic.Visible = False
    Next
    
    Select Case Index
    Case 0, 3, 6
        picPanel(0).Visible = True
        optDisplay(0).Value = True
    Case 1, 4, 7
        picPanel(1).Visible = True
        optDisplay(4).Value = True
    Case 2, 5, 8
        picPanel(2).Visible = True
        optDisplay(8).Value = True
    End Select
    
End Sub

Private Sub lstResults_MouseUp(Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               y As Single)
'/* show list menu

    If Button = vbRightButton Then
        Me.PopupMenu Me.mnuMain
        
    End If
    
End Sub

Private Sub mnuMain1_Click(Index As Integer)
'/* menu calls

Dim sPath As String

On Error GoTo Handler

    If vbRightButton Then
        Select Case Index
        '/* jump to location
        Case 0
            ShellExecute Me.hwnd, "open", lstResults.SelectedItem.SubItems(1) & Chr$(92), "", "", 1
            
        '/* file properties
        Case 1
            With lstResults
                sPath = .SelectedItem.SubItems(1) + Chr$(92) + .SelectedItem.Text
            End With
            With frmProperties
                .File_Data sPath, lstResults.SelectedItem.SubItems(5)
                .Show vbModeless, Me
            End With
        End Select
    End If
    
Handler:
    
End Sub

Private Sub optIndex_Click(Index As Integer)
'/* reflect to frmQuery

    frmQuery.optIndex(Index).Value = True

End Sub

Private Sub optModified_Click(Index As Integer)
'/* enable/disable user options

Dim mOpt As OptionButton

    For Each mOpt In optModified
        mOpt.FontBold = False
    Next mOpt

    dtDateType.Enabled = False
    txtFrom.Enabled = False
    txtTo.Enabled = False

    Select Case Index
    Case 0
        optModified(0).FontBold = True
    Case 1
        optModified(1).FontBold = True
    Case 2
        optModified(2).FontBold = True
    Case 3
        optModified(3).FontBold = True
    Case 4
        optModified(4).FontBold = True
        dtDateType.Enabled = True
        txtFrom.Enabled = True
        txtTo.Enabled = True
    End Select

End Sub

Private Sub optSize_Click(Index As Integer)
'/* enable/disable user options

Dim mOpt As OptionButton

    For Each mOpt In optSize
        mOpt.FontBold = False
    Next mOpt

    cboSize.Enabled = False
    txtSize.Enabled = False

    Select Case Index
    Case 0
        optSize(0).FontBold = True
    Case 1
        optSize(1).FontBold = True
    Case 2
        optSize(2).FontBold = True
    Case 3
        optSize(3).FontBold = True
    Case 4
        optSize(4).FontBold = True
        cboSize.Enabled = True
        txtSize.Enabled = True
    End Select

End Sub

Private Sub List_Reset()
'/* clear listview

    lstResults.ListItems.Clear

End Sub


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                           FILTER ROUTINES
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


Private Sub Filter_Check(sMatch As String)
'/* search filter hub

Dim lDays   As Long
Dim sDate   As String
Dim bDate   As Boolean
Dim bSize   As Boolean
Dim bType   As Boolean
Dim dSize   As Double
Dim mDate   As Date

On Error Resume Next

    '/* minimize lstview flicker
    m_lCounter = m_lCounter + 1
    If m_lCounter = 10 Then
        LockWindowUpdate &H0
        DoEvents
        LockWindowUpdate lstResults.hwnd
        m_lCounter = 0
    End If
    
    '/* file date
    If Not optModified(0).Value Then
        bDate = True
        '/* file file data structure
        Get_Structure sMatch
        'sDate = Get_Created(sMatch)
        mDate = CDate(Get_Created(sMatch))
        lDays = DateDiff("d", mDate, Format(Now, "dd/mm/yyyy"))
        Select Case True
        '/* less then 7 days
        Case optModified(1).Value
            If lDays < 8 Then bDate = False
        '/* less then a month
        Case optModified(2).Value
            If lDays < 32 Then bDate = False
        '/* less then a year
        Case optModified(3).Value
            If lDays < 366 Then bDate = False
        '/* user select
        Case optModified(4).Value
            If (mDate > CDate(txtFrom.Text)) And (mDate < CDate(txtTo.Text)) Then bDate = False
        End Select
    End If
    
    '/* file size
    If Not optSize(0).Value Then
        bSize = True
        '/* file file data structure
        Get_Structure sMatch
        '/* get file size in kb
        dSize = Get_Size(sMatch) / 1000
        Select Case True
        '/* less then 100kb
        Case optSize(1).Value
            If dSize < 100 Then bSize = False
        '/* less then 1mb
        Case optSize(2).Value
            If dSize < 1000 Then bSize = False
        '/* more then 1mb
        Case optSize(3).Value
            If dSize > 1000 Then bSize = False
        '/* user specfied
        Case optSize(4).Value
            If dSize > (CLng(txtSize.Text) * 1000) Then bSize = False
        End Select
    End If
    
    '/* type filters
    If Not cboFileType.ListIndex = 0 Then
        Select Case cboFileType.ListIndex
        '/* picture
        Case 1
            bType = Image_Filter(sMatch)
        '/* document
        Case 2
            bType = Document_Filter(sMatch)
        '/* music
        Case 3
            bType = Music_Filter(sMatch)
        '/* executable
        Case 4
            If right$(sMatch, 4) = ".exe" Then bType = False
        End Select
    End If
    
    '/* path exclusion filter
    If Path_Filter(sMatch) Then
        bType = True
    End If
    
    '/* if still a match, add it
    If Not (bDate Or bSize Or bType) = True Then
        Get_Structure sMatch
        Attributes sMatch
    End If
    DoEvents
    
On Error GoTo 0

End Sub

Private Sub Get_Structure(ByVal sMatch As String)
'/* get the file data and load into type struct

Dim hFile As Long

    hFile = CreateFile(sMatch, 0, 0, ByVal 0&, &H3, 0, ByVal 0&)
    GetFileInformationByHandle hFile, FileInfo
    CloseHandle hFile
    
End Sub

Private Function Get_Size(ByVal sMatch As String) As Long
'/* get file size from structure

    Get_Size = (FileInfo.lFSH * (&HFFFF + 1) + FileInfo.lFSL)
    
End Function

Private Function Get_Created(ByVal sMatch As String) As String
'/* get time created from structure

    Get_Created = Date_Convert(CTime(FileInfo.fCT))
    
End Function

Private Sub Attributes(sMatch As String)
'/* fetch and return file attributes for match

Dim hFile    As Long
Dim lIcon As Integer

On Error GoTo Handler

    lIcon = Get_Icon(Mid$(sMatch, InStrRev(sMatch, Chr$(46)) + 1)) - 1
    Set m_lItem = lstResults.ListItems.Add(Text:=Mid$(sMatch, InStrRev(sMatch, Chr$(92)) + 1))
    m_lItem.SmallIcon = lIcon + 1

    With FileInfo
        '/* full path
        m_lItem.SubItems(1) = left$(sMatch, InStrRev(sMatch, Chr$(92)) - 1)
        '/* accessed
        m_lItem.SubItems(2) = FDate(CTime(.fLA))
        '/* created
        m_lItem.SubItems(3) = FDate(CTime(.fCT))
        '/* written
        m_lItem.SubItems(4) = FDate(CTime(.fLWT))
        '/* size
        m_lItem.SubItems(5) = .lFSH * (&HFFFF + 1) + .lFSL & " Bytes"
    End With

Handler:

End Sub

Private Function CTime(ByRef tFILETIME As FT) As Double
'/* format api time to readable

Dim tFile As FT
Dim m_tTime As TIME

    FileTimeToLocalFileTime tFILETIME, tFile
    FileTimeToSystemTime tFile, m_tTime
    With m_tTime
        CTime = DateSerial(.wYR, .wMNT, .wDAY) + TimeSerial(.wHR, .wMIN, .wSEC)
    End With

End Function

Private Function Date_Convert(ByVal dDate As Double) As String
'/* format date

    Date_Convert = Format$(dDate, "dd/mm/yyyy")

End Function

Private Function FDate(ByVal dDate As Double) As String
'/* format date

    FDate = Format$(dDate, "dd/mm/yyyy") & " " & Format$(dDate, "dd/mm/yyyy")

End Function

Private Function Get_Icon(ByVal sType As String) As Integer
'/* custom icons.. I will expand
'/* this in a future revision

    Select Case sType
    '/* music
    Case "mp3", "mp2", "asf", "wm", "wma", "wmv", _
    "asx", "wax", "m3u", "wvl", "wav", "snd", _
    "au", "aif", "mid", "rmi", "dvr-ms"
    Get_Icon = 1
    '/* images
    Case "jpeg", "jpg", "mpeg", "wmv", "avi", _
    "bmp", "dib", "rle", "gif", "psd", "ai", _
    "art", "cdr", "eps", "png", "fh9", "fh10", _
    "fh11", "ft7", "ft8", "tga", "tif", "tif", "wbm"
    Get_Icon = 2
    '/* web page
    Case "htm", "xhtm", "xhtml", "cfm"
    Get_Icon = 3
    '/* document
    Case "txt", "pdf", "doc", "dot", "rtf", "wri", "xls", _
    "xlw", "ans", "wps", "wk1", "wk3", "wk4", "mcw", "log"
    Get_Icon = 4
    '/* executable
    Case "exe", "lnk", "msi", "scr", "msc", "com", "cmd", "bat", _
    "mst", "bin"
    Get_Icon = 5
    '/* library
    Case "dll", "src", "sys", "cpl", "cfg", "wsc", "ocx", "inf", _
    "acm", "srg", "tlb", "nls"
    Get_Icon = 6
    '/* compressed
    Case "zip", "ace", "000", "rar", "pak", "gz", "gzip"
    Get_Icon = 7
    '/* help
    Case "hlp", "chm", "cnt", "hta"
    Get_Icon = 8
    '/* project files - just for you guys ;o)
    Case "vba", "vbg", "cls", "bas", "scc", "cpp", "h", "c", _
    "rcc", "res", "plg", "dsp", "dsw", "rul", "vb", "vbe", "js", "jse"
    Get_Icon = 9
    Case Else
    Get_Icon = 10
    
    End Select
    
End Function

Private Function Music_Filter(ByVal sMatch As String) As Boolean
'/* music file filter list
'/* in all filters, placed
'/* most likely matches first

On Error Resume Next

    Music_Filter = True
    Select Case Mid$(sMatch, InStrRev(sMatch, Chr$(46)))
    '/* mp3
    Case ".mp3", ".mp2"
    If chkMFilter(0).Value Then Music_Filter = False
    '/* windows media
    Case ".asf", ".wm", ".wma", ".wmv"
        If chkMFilter(5).Value Then Music_Filter = False
    '/* playlist
    Case ".asx", ".wax", ".m3u", ".wvl"
        If chkMFilter(2).Value Then Music_Filter = False
    '/* wav
    Case ".wav", ".snd", ".au", ".aif"
        If chkMFilter(1).Value Then Music_Filter = False
    '/* midi
    Case ".mid", ".rmi"
        If chkMFilter(4).Value Then Music_Filter = False
    '/* ms tv
    Case ".dvr-ms"
        If chkMFilter(3).Value Then Music_Filter = False
    End Select

On Error GoTo 0

End Function

Private Function Image_Filter(ByVal sMatch As String) As Boolean
'/* image file filter list

    Image_Filter = True
    Select Case Mid$(sMatch, InStrRev(sMatch, Chr$(46)))
    '/* jpeg
    Case ".jpeg", ".jpg"
    If chkIFilter(7).Value Then Image_Filter = False
    '/* movie
    Case ".mpeg", ".wmv", ".avi"
        If chkIFilter(12).Value Then Image_Filter = False
    '/* bitmap
    Case ".bmp", ".dib", ".rle"
        If chkIFilter(1).Value Then Image_Filter = False
    '/* gif
    Case ".gif"
        If chkIFilter(6).Value Then Image_Filter = False
    '/* photoshop
    Case ".psd"
        If chkIFilter(8).Value Then Image_Filter = False
    '/* illustrator
    Case ".ai", ".art"
        If chkIFilter(0).Value Then Image_Filter = False
    '/* corel draw
    Case ".cdr"
        If chkIFilter(2).Value Then Image_Filter = False
    '/* eps
    Case ".eps"
        If chkIFilter(3).Value Then Image_Filter = False
    '/* fireworks
    Case ".png"
        If chkIFilter(4).Value Then Image_Filter = False
    '/* freehand
    Case ".fh9", ".fh10", ".fh11", ".ft7", ".ft8"
        If chkIFilter(5).Value Then Image_Filter = False
    '/* targa
    Case ".tga"
        If chkIFilter(9).Value Then Image_Filter = False
    '/* tiff
    Case ".tif"
        If chkIFilter(10).Value Then Image_Filter = False
    '/* wbmp
    Case ".wbm"
        If chkIFilter(11).Value Then Image_Filter = False
    End Select

On Error GoTo 0

End Function

Private Function Document_Filter(ByVal sMatch As String) As Boolean
'/* document file filter list

On Error Resume Next

    Document_Filter = True
    Select Case Mid$(sMatch, InStrRev(sMatch, Chr$(46)))
    '/* ascii
    Case ".txt"
        If chkDocFilter(1).Value Then Document_Filter = False
    '/* html
    Case ".htm", ".xhtm", ".xhtml", ".cfm"
        If chkDocFilter(3).Value Then Document_Filter = False
    '/* adobe
    Case ".pdf"
        If chkDocFilter(0).Value Then Document_Filter = False
    '/* word
    Case ".doc", ".dot"
        If chkDocFilter(7).Value Then Document_Filter = False
    '/* rich text
    Case ".rtf"
        If chkDocFilter(5).Value Then Document_Filter = False
    '/* write
    Case ".wri"
        If chkDocFilter(10).Value Then Document_Filter = False
    '/* excel
    Case ".xls", ".xlw"
        If chkDocFilter(2).Value Then Document_Filter = False
    '/* extended text
    Case ".ans"
        If chkDocFilter(6).Value Then Document_Filter = False
    '/* works
    Case ".wps"
        If chkDocFilter(9).Value Then Document_Filter = False
    '/* lotus 1-2-3
    Case ".wk1", ".wk3", ".wk4"
        If chkDocFilter(4).Value Then Document_Filter = False
    '/* word for mac
    Case ".mcw"
        If chkDocFilter(8).Value Then Document_Filter = False
    End Select

On Error GoTo 0

End Function

Private Function Path_Filter(ByVal sPath As String) As Boolean
'/* filter on critical paths

On Error Resume Next

    Select Case True
    '/* libraries
    Case InStr(sPath, "system32") > 0
        If chkTFilter(7).Value Then
            Path_Filter = True
        End If
    '/* security data
    Case InStr(sPath, "Windows\system32\usmt") > 0
        If chkTFilter(1).Value Then
            Path_Filter = True
        End If
    '/* updates
    Case InStr(sPath, "Windows\system32\CatRoot2") > 0
        If chkTFilter(2).Value Then
            Path_Filter = True
        End If
    '/* sys volume
    Case InStr(sPath, "System Volume Information") > 0
        If chkTFilter(0).Value Then
            Path_Filter = True
        End If
    '/* security
    Case InStr(sPath, "Windows\security") > 0
        If chkTFilter(3).Value Then
            Path_Filter = True
        End If
    '/* sp files
    Case InStr(sPath, "Windows\ServicePackFiles") > 0
        If chkTFilter(4).Value Then
            Path_Filter = True
        End If
    '/* windows files
    Case InStr(sPath, "Windows\SoftwareDistribution") > 0
        If chkTFilter(5).Value Then
            Path_Filter = True
        End If
    '/* application data
    Case InStr(sPath, "Application Data\Microsoft") > 0
        If chkTFilter(6).Value Then
            Path_Filter = True
        End If
    End Select

On Error GoTo 0

End Function


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                           REGISTRY ROUTINES
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Private Sub Get_Settings()
'/* restore option settings

On Error Resume Next

    '/* lock the index update trigger
    m_bChange = True
    
    With New clsLightning
        '/* test for first run flag
        If .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "fstrun") = 0 Then Exit Sub
        
        '/* service state test result:
        '/* if wmi failed on last run, revert
        '/* to 98/ME display mode
        If .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "svfl") = 1 Then Revert_Mode
        
        '/* image filters
        chkIFilter(0).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "imf0")
        chkIFilter(1).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "imf1")
        chkIFilter(2).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "imf2")
        chkIFilter(3).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "imf3")
        chkIFilter(4).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "imf4")
        chkIFilter(5).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "imf5")
        chkIFilter(6).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "imf6")
        chkIFilter(7).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "imf7")
        chkIFilter(8).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "imf8")
        chkIFilter(9).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "imf9")
        chkIFilter(10).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "imf10")
        chkIFilter(11).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "imf11")
        
        '/* music filters
        chkMFilter(0).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "mcf0")
        chkMFilter(1).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "mcf1")
        chkMFilter(2).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "mcf2")
        chkMFilter(3).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "mcf3")
        chkMFilter(4).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "mcf4")
        chkMFilter(5).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "mcf5")
        
        '/* document filters
        chkDFilter(0).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "dtf0")
        chkDFilter(1).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "dtf1")
        chkDFilter(2).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "dtf2")
        chkDFilter(3).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "dtf3")
        chkDFilter(4).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "dtf4")
        chkDFilter(5).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "dtf5")
        chkDFilter(6).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "dtf6")
        chkDFilter(7).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "dtf7")
        chkDFilter(8).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "dtf8")
        chkDFilter(9).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "dtf9")
        chkDFilter(10).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "dtf10")
        
        '/* startup options
        chkStart(0).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "suf0")
        chkStart(1).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "suf1")
        chkStart(2).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "suf2")
        
        '/* indexing options
        chkIndex.Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "ckid")
        optIndex(0).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "ixo0")
        optIndex(1).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "ixo1")
        optIndex(2).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "ixo2")
        cbDays.ListIndex = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "dyc0")
        
        '/* exclusion filters
        chkTFilter(0).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "exf0")
        chkTFilter(1).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "exf1")
        chkTFilter(2).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "exf2")
        chkTFilter(3).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "exf3")
        chkTFilter(4).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "exf4")
        chkTFilter(5).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "exf5")
        chkTFilter(6).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "exf6")
        chkTFilter(7).Value = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "exf7")
        
        '/* last drive displayed
        Drive1.ListIndex = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "lstidx")
    End With
    
    '/* enable trigger
    m_bChange = False
    
On Error GoTo 0

End Sub

Private Sub Save_Settings()
'/* save option settings

On Error Resume Next

    With New clsLightning
        '/* first run switch
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "fstrun", 1
        
        '/* image filters
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "imf0", chkIFilter(0).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "imf1", chkIFilter(1).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "imf2", chkIFilter(2).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "imf3", chkIFilter(3).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "imf4", chkIFilter(4).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "imf5", chkIFilter(5).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "imf6", chkIFilter(6).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "imf7", chkIFilter(7).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "imf8", chkIFilter(8).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "imf9", chkIFilter(9).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "imf10", chkIFilter(10).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "imf11", chkIFilter(11).Value
        
        '/* music filters
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "mcf0", chkMFilter(0).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "mcf1", chkMFilter(1).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "mcf2", chkMFilter(2).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "mcf3", chkMFilter(3).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "mcf4", chkMFilter(4).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "mcf5", chkMFilter(5).Value
        
        '/* document filters
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "dtf0", chkDFilter(0).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "dtf1", chkDFilter(1).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "dtf2", chkDFilter(2).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "dtf3", chkDFilter(3).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "dtf4", chkDFilter(4).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "dtf5", chkDFilter(5).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "dtf6", chkDFilter(6).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "dtf7", chkDFilter(7).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "dtf8", chkDFilter(8).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "dtf9", chkDFilter(9).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "dtf10", chkDFilter(10).Value
        
        '/* startup options
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "suf0", chkStart(0).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "suf1", chkStart(1).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "suf2", chkStart(2).Value
        
        '/* indexing options
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "ckid", chkIndex.Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "ixo0", CBool(optIndex(0).Value)
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "ixo1", CBool(optIndex(1).Value)
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "ixo2", CBool(optIndex(2).Value)
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "ixo3", CBool(chkIndex.Value)
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "dyc0", cbDays.ListIndex
        
        '/* exclusion filters
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "exf0", chkTFilter(0).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "exf1", chkTFilter(1).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "exf2", chkTFilter(2).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "exf3", chkTFilter(3).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "exf4", chkTFilter(4).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "exf5", chkTFilter(5).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "exf6", chkTFilter(6).Value
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "exf7", chkTFilter(7).Value
        
        '/* search properties
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "lstidx", Drive1.ListIndex
        
        '/* set loaded switch
        If Forms = 1 Then
            '/* unloaded
            .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "loaded", 1
        Else
            '/* running minimized
            .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "loaded", 3
        End If
        
    End With
    
On Error GoTo 0

End Sub


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                           INDEX ROUTINES
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Public Function Index_Control(ByVal iOperation As Integer, _
                              Optional ByVal iIndex As Long) As Boolean

'/* index update hub

Dim sIndPath    As String
Dim iLast       As Integer
Dim sScnTm      As String
Dim sUsed       As String

On Error Resume Next

    Select Case iOperation
    '/* load last index
    Case 1
        '/* get last index loaded
        With New clsLightning
            iLast = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "lstidx")
        End With
        '/* if it Exists, load It
        sIndPath = App.Path & "\index" & iLast & ".dat"
        If File_Exists(sIndPath) Then
            With cEngine
                .p_IndexPath = sIndPath
                .p_EngineTask = Index_Restore
                .Start
            End With
            Index_Control = True
            stBar.Panels(1).Text = "Loading Index for Drive: " & left$(Drive1.Drive, 2) & Chr$(92)
            cmdControls.Enabled = False
            m_bChange = True
            Drive1.ListIndex = CInt(iLast)
            m_bChange = False
        End If
        
    '/* save current index
    Case 2
        '/* put last index loaded
        With New clsLightning
            .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "lstidx", iIndex
        End With
        sIndPath = App.Path & "\index" & iIndex & ".dat"
        '/* dump the index
        With cEngine
            .p_IndexPath = sIndPath
            .p_EngineTask = Index_Save
            .Start
        End With
        
    '/* enumerate and update indeces
    Case 3
        Dim vItem   As Variant
        Dim lLoaded As Long
        Dim sDrive  As String
        '/* loaded: /1=unloaded/2=running/3=minimized
        
        '/* loaded mode
        With New clsLightning
            '/* get the display mode
            lLoaded = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "loaded")
        End With
        
        '/* update method minimized/running
        Select Case lLoaded

        '/* running
        Case 2
            '/* ask user if they want to defer
            If Not Custom_Message("Pending Update", "Hyperion has scheduled to update the drive indeces.", _
                "Would you like to update now, or start the update the next time the application is minimized?", frmMain) = 1 Then
                '/* reschedule
                With New clsLightning
                    .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "schreset", 1
                End With
            Else
                '/* pass to update control
                Update_Controller 2
                '/* get last index loaded
                With New clsLightning
                    iLast = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "lstidx")
                End With
                '/* restore the last used index
                sIndPath = App.Path & "\index" & iLast & ".dat"
                With cEngine
                    .p_IndexPath = sIndPath
                    .p_EngineTask = Index_Restore
                    .Start
                End With
            End If
            
        '/* minimized
        Case 3
            'Load frmStatus
            '/* notify user of updates
            Update_Controller 1
        End Select
    End Select

On Error GoTo 0

End Function

Public Sub Update_Controller(Optional iMethod As Integer)
'/* NOTE: Update routines (may) not work properly running from
'/* linked projects. Engine may have to be compiled, and run as a
'/* seperate object for timer seperation to occur..

Dim sDrive      As String
Dim sUsed       As String
Dim sScnTm      As String
Dim vItem       As Variant
Dim lItem       As Long
Dim sIndPath    As String

On Error Resume Next

    With New clsLightning
        '/* lightning returns a collection of index values
        For Each vItem In .List_Values(HKEY_CURRENT_USER, m_sRegPath & "\Index")
            '/* get the drive letter
            sDrive = Get_Drive(CLng(vItem))
            '/* test for proper path
            If Len(sDrive) = 3 Then
                Scan_Time sDrive, sUsed, sScnTm
                '/* notify methods
                Select Case iMethod
                '/* main form notifier
                Case 1
                    '/* notify user
                    With frmMain
                        .sNotify 2, sDrive, sUsed, sScnTm
                        .stBar.Panels(1).Text = "Index for Drive: " & _
                        left$(.Drive1.Drive, 2) & Chr$(92) & " loaded."
                    End With
                Case 2
                    '/* notify user
                    With frmMain
                        .sNotify 3, sDrive, sUsed, sScnTm
                        .stBar.Panels(1).Text = "Index for Drive: " & _
                        left$(.Drive1.Drive, 2) & Chr$(92) & " loaded."
                    End With
                End Select
                '/* start scan
                sIndPath = App.Path & "\index" & CInt(vItem) & ".dat"
                '/* disable menu during update
                frmSysTray.mnuMain.Enabled = False
                '/* flag wait timer
                m_bUpdateStatus = True
                With cEngine
                    '/* reset our container
                    .Storage_Reset
                    '/* set scan path
                    .p_BuildPath = sDrive
                    '/* set index path
                    .p_IndexPath = sIndPath
                    '/* build index flag
                    .p_CMultiTask.Add 1
                    '/* save index flag
                    .p_CMultiTask.Add 2
                    '/* set to multi task mode
                    .p_EngineTask = Engine_MultiTask
                    '/* start engine
                    .Start
                    '/* wait for callback
                    Status_Loop
                    '/* just in case..
                    m_bUpdateStatus = False
                    '/* reload
                    .Engine_Reload
                End With
            End If
        Next
    End With
    
    '/* done - reset scheduler flag
    With New clsLightning
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "schreset", 2
    End With
    
On Error GoTo 0

End Sub

Public Sub sNotify(iMessage As Integer, _
                   Optional ByVal sDrv As String, _
                   Optional sUsed As String, _
                   Optional sTime As String)
'/* taskbar alert notification

Dim Alert           As frmAlert
Dim sMessage        As String

    Set Alert = New frmAlert
    
    Select Case iMessage
    '/* user notify of operation
    Case 1
        '/* get the ball rolling..
        Index_Control 3
        
    '/* pending reindex - minimized
    Case 2
        sMessage = "Hyperion is creating an Index" & vbNewLine & _
        "Scanning Drive: " & sDrv & vbNewLine & _
        sUsed & vbNewLine & sTime
        With Alert
            .DisplayMessage sMessage, 2, False, True, 1, 15526111, 1
        End With
        Msg_Timer 3000
        
    '/* pending reindex - running
    Case 3
        sMessage = "Hyperion is creating an Index" & vbNewLine & _
        "Scanning Drive: " & sDrv & vbNewLine & _
        sUsed & vbNewLine & sTime
        With Alert
            .DisplayMessage sMessage, 2, False, True, 1, 15526111, 1
        End With
        Msg_Timer 3000
        
    '/* pending reindex - unload all
    Case 4
        sMessage = "Hyperion will now Resume updating Drives"
        Me.Hide
        With Alert
            '.Show vbModeless, frmSysTray
            .DisplayMessage sMessage, 2, False, True, 1, 15526111, 1
        End With
        Msg_Timer 3000
        Unload_All
        
    '/* reserved
    Case 5
    
    End Select

End Sub

Private Sub Msg_Timer(ByVal lInterval As Long)
'/* rough timer

Dim lTick   As Long
Dim lCount  As Long

On Error Resume Next

    If lInterval = -1 Then Exit Sub
        lTick = GetTickCount()
        lTick = lTick + lInterval
        If lTick > 0& Then
            lTick = ((lTick + &H80000000) + lInterval) + &H80000000
        Else
            lTick = ((lTick - &H80000000) + lInterval) - &H80000000
        End If

    Do
        If Err.Number = 0 Then Exit Sub
        lCount = GetTickCount()
        lCount = lTick - lCount
        If lTick > 0& Then
            lCount = ((lTick + &H80000000) - (lCount - &H80000000))
        Else
            lCount = ((lTick - &H80000000) - (lCount + &H80000000))
        End If
            
        If IIf((lCount Xor lInterval) > 0&, lCount > lInterval, lCount < 0&) Then
            Exit Sub
        End If
        
        MsgWaitForMultipleObjects 0&, 0&, 0&, lCount, QS_ALLINPUT
        DoEvents
    Loop

On Error GoTo 0
    
End Sub

Private Sub Status_Loop()
'/* wait for message flag

    Do While m_bUpdateStatus
        m_lLoopCounter = m_lLoopCounter + 1
        Msg_Timer 200
        '/* if time exceeds 5 minutes, bail
        If m_lLoopCounter > 1000 Then
            m_bQuickExit = True
            m_bUpdateStatus = False
            Exit Do
        End If
        DoEvents
    Loop

    m_lLoopCounter = 0

End Sub










'*~ December 21, 2005 10pm~*

'*~ I've had this song stuck in my head all day..
'*~ As I get ready now to zip this project, and enter this
'*~ last note, it is running full tilt on the stereo..

'*~ AC/DC - Are You Ready For a Good Time? ~*

'*~ Are you ready for a good time?
'*~ Then get ready for the night line..
'*~ Are you ready for a good time?
'*~ Then get ready for the night line..
'*~ Are you ready?!?
'*~ Are you Ready?????



'                          STEPPENWOLFE
'        STEPPENWOLFE  STEPPENWOLFESTEPPENWOLFEST
'        STEPPENWOLFESTEPPENWOLFE    STEPPENWOLFESTEPPE
'        STEPPENW      STEPPENWOLFESTEPPE    STEPPENWOLFESTEPPENW
'          STEPPE            NWOLFESTEPPENWOLFE  STEP     PENWOLFEST
'          STEPPE                NWOLFESTEPPENWOL                  FEST
'            STEPPE                ST  EPPENWOLFE              STEP  ENWO
'            STEPPE                STEPPENWOLFE              STEPPENWOLFEST
'            STEPPE                  STEP  PENW            STEPPENWOLFESTEPPENW
'            STEPPE                    NWOLFE            STEPPENWOLFESTEPPENWOLFESTEPPENWOLFESTE
'            STEPPE    STEPPENWOLFESTEPPE         STEPPENWOLFE                            STEPPENWO
'            STEPPENWOLFESTEP                    PENWOLFEST                                      EP
'            STEPPENWOLFESTEPPENWOLFE        STEPPENWOLFE                                      STEP
'            STEPPENWOL    ESTEPPENWOLFESTEPPENWOLF                    ESTE                    PPE
'            STEP      STEPPENW    OLFESTEPPENWOLFEST                EPPENWOLFEST           EPPEN
'            STEP      PENWOL                  FESTEPPE          NWOLFESTEPPENWOLFE    STEPPEN
'          STEPPENWOLFESTEPPENWOLFESTEPPENWOLFEST                EPPENWOLFESTEPPENWOLFEST
'          STEPPENWOLFESTEP  PENWOLFWS  TEPPEN                WOLFESTEPP            ENW
'          STEPPENWOL    FESTEP        STEPPENWOLFE        STEPPENWOLFEST          EP
'          STEP      PENWOL      FESTEPPENWOLFESTEPPENWOLFEST      EPPENWOLFE    STEP
'        STEPPE    NWOL      FESTEPPENWOL                FE          STEPPENWOLFEST
'        STEP  PENWOL    FESTEPPEN                    WO
'


'~ update: 3am: - wow, am I ever wasted.. ~*
