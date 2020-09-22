VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form RipperProgramMainForm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Advanced Ripper Program"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6390
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFF00&
   Icon            =   "RipperProgramMainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "RipperProgramMainForm.frx":0BC2
   MousePointer    =   99  'Custom
   ScaleHeight     =   6555
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command10 
      BackColor       =   &H00FF8080&
      Caption         =   "Expand Viv"
      Height          =   315
      Left            =   4830
      MouseIcon       =   "RipperProgramMainForm.frx":0D14
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   3660
      Width           =   1515
   End
   Begin VB.CheckBox Fsh 
      Caption         =   "Check1"
      Height          =   195
      Left            =   2250
      MouseIcon       =   "RipperProgramMainForm.frx":0E66
      MousePointer    =   99  'Custom
      TabIndex        =   80
      Top             =   2580
      Width           =   195
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FF8080&
      Caption         =   "Replace text in exe"
      Height          =   315
      Left            =   3270
      MouseIcon       =   "RipperProgramMainForm.frx":0FB8
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   4740
      Width           =   1515
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FF8080&
      Caption         =   "Exit"
      Height          =   315
      Left            =   4830
      MouseIcon       =   "RipperProgramMainForm.frx":110A
      MousePointer    =   99  'Custom
      OLEDropMode     =   1  'Manual
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   4320
      Width           =   1515
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FF8080&
      Caption         =   "Decrypt File"
      Height          =   315
      Left            =   1560
      MouseIcon       =   "RipperProgramMainForm.frx":125C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   6150
      Width           =   1515
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FF8080&
      Caption         =   "Encrypt File"
      Height          =   315
      Left            =   30
      MouseIcon       =   "RipperProgramMainForm.frx":13AE
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   6150
      Width           =   1515
   End
   Begin VB.CheckBox Tiff 
      Caption         =   "Check1"
      Height          =   195
      Left            =   2250
      MouseIcon       =   "RipperProgramMainForm.frx":1500
      MousePointer    =   99  'Custom
      TabIndex        =   72
      Top             =   2340
      Width           =   195
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FF8080&
      Caption         =   "Replace File"
      Height          =   315
      Left            =   4830
      MouseIcon       =   "RipperProgramMainForm.frx":1652
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   390
      Width           =   1515
   End
   Begin VB.CheckBox Swf 
      Caption         =   "Check1"
      Height          =   195
      Left            =   2250
      MouseIcon       =   "RipperProgramMainForm.frx":17A4
      MousePointer    =   99  'Custom
      TabIndex        =   68
      Top             =   2100
      Width           =   195
   End
   Begin VB.CheckBox AllOptions 
      Caption         =   "Check1"
      Height          =   195
      Left            =   1920
      MouseIcon       =   "RipperProgramMainForm.frx":18F6
      MousePointer    =   99  'Custom
      TabIndex        =   66
      Top             =   3810
      Width           =   195
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3780
      Top             =   2400
   End
   Begin VB.CheckBox Junk 
      Caption         =   "Check1"
      Height          =   195
      Left            =   690
      MouseIcon       =   "RipperProgramMainForm.frx":1A48
      MousePointer    =   99  'Custom
      TabIndex        =   64
      Top             =   4860
      Width           =   195
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF8080&
      Caption         =   "About/Help"
      Height          =   315
      Left            =   4830
      MouseIcon       =   "RipperProgramMainForm.frx":1B9A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   3990
      Width           =   1515
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   510
      Top             =   420
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF8080&
      Caption         =   "Search Exe 4 Text"
      Height          =   315
      Left            =   4830
      MouseIcon       =   "RipperProgramMainForm.frx":1CEC
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   4740
      Width           =   1515
   End
   Begin VB.CheckBox Avi 
      Caption         =   "Check1"
      Height          =   195
      Left            =   2250
      MouseIcon       =   "RipperProgramMainForm.frx":1E3E
      MousePointer    =   99  'Custom
      TabIndex        =   58
      Top             =   1860
      Width           =   195
   End
   Begin VB.CheckBox Bink 
      Caption         =   "Check1"
      Height          =   195
      Left            =   2250
      MouseIcon       =   "RipperProgramMainForm.frx":1F90
      MousePointer    =   99  'Custom
      TabIndex        =   55
      Top             =   1620
      Width           =   195
   End
   Begin VB.CheckBox Html 
      Caption         =   "Check1"
      Height          =   195
      Left            =   2250
      MouseIcon       =   "RipperProgramMainForm.frx":20E2
      MousePointer    =   99  'Custom
      TabIndex        =   38
      Top             =   1380
      Width           =   195
   End
   Begin VB.CheckBox Acrobat 
      Caption         =   "Check1"
      Height          =   195
      Left            =   2250
      MouseIcon       =   "RipperProgramMainForm.frx":2234
      MousePointer    =   99  'Custom
      TabIndex        =   36
      Top             =   1140
      Width           =   195
   End
   Begin VB.CheckBox Wave 
      Caption         =   "Check1"
      Height          =   195
      Left            =   2250
      MouseIcon       =   "RipperProgramMainForm.frx":2386
      MousePointer    =   99  'Custom
      TabIndex        =   34
      Top             =   900
      Width           =   195
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   5340
      Top             =   840
   End
   Begin VB.CheckBox Selectall 
      Caption         =   "Check1"
      Height          =   195
      Left            =   120
      MouseIcon       =   "RipperProgramMainForm.frx":24D8
      MousePointer    =   99  'Custom
      TabIndex        =   31
      Top             =   3810
      Width           =   195
   End
   Begin VB.CheckBox Tga 
      Caption         =   "Check1"
      Height          =   195
      Left            =   540
      MouseIcon       =   "RipperProgramMainForm.frx":262A
      MousePointer    =   99  'Custom
      TabIndex        =   29
      Top             =   3540
      Width           =   195
   End
   Begin VB.CheckBox Deletthem 
      Caption         =   "Check1"
      Height          =   195
      Left            =   120
      MouseIcon       =   "RipperProgramMainForm.frx":277C
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   4110
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Bmp16 
      Caption         =   "Check1"
      Height          =   195
      Left            =   540
      MouseIcon       =   "RipperProgramMainForm.frx":28CE
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   3300
      Width           =   195
   End
   Begin VB.CheckBox Bmp256 
      Caption         =   "Check1"
      Height          =   195
      Left            =   540
      MouseIcon       =   "RipperProgramMainForm.frx":2A20
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   3060
      Width           =   195
   End
   Begin VB.CheckBox Mus 
      Caption         =   "Midi"
      Height          =   195
      Left            =   540
      MouseIcon       =   "RipperProgramMainForm.frx":2B72
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   2820
      Width           =   195
   End
   Begin VB.CheckBox Xmi 
      Caption         =   "Midi"
      Height          =   195
      Left            =   540
      MouseIcon       =   "RipperProgramMainForm.frx":2CC4
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   2580
      Width           =   195
   End
   Begin VB.CheckBox Paint 
      Caption         =   "Midi"
      Height          =   195
      Left            =   540
      MouseIcon       =   "RipperProgramMainForm.frx":2E16
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   2340
      Width           =   195
   End
   Begin VB.CheckBox Hmp 
      Caption         =   "Midi"
      Height          =   195
      Left            =   540
      MouseIcon       =   "RipperProgramMainForm.frx":2F68
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   2100
      Width           =   195
   End
   Begin VB.CheckBox Zip 
      Caption         =   "Midi"
      Height          =   195
      Left            =   540
      MouseIcon       =   "RipperProgramMainForm.frx":30BA
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   1860
      Width           =   195
   End
   Begin VB.CheckBox Midi 
      Caption         =   "Midi"
      Height          =   195
      Left            =   540
      MouseIcon       =   "RipperProgramMainForm.frx":320C
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   1620
      Width           =   195
   End
   Begin VB.CheckBox Jpeg 
      Caption         =   "Check1"
      Height          =   195
      Left            =   540
      MouseIcon       =   "RipperProgramMainForm.frx":335E
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   1380
      Width           =   195
   End
   Begin VB.CheckBox Gif 
      Caption         =   "Check1"
      Height          =   195
      Left            =   540
      MouseIcon       =   "RipperProgramMainForm.frx":34B0
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1140
      Width           =   195
   End
   Begin VB.CheckBox Messages 
      Caption         =   "Check1"
      Height          =   195
      Left            =   120
      MouseIcon       =   "RipperProgramMainForm.frx":3602
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   4410
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Bitmap 
      Caption         =   "Check1"
      Height          =   195
      Left            =   540
      MouseIcon       =   "RipperProgramMainForm.frx":3754
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   900
      Width           =   195
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      Caption         =   "Exit"
      Height          =   315
      Left            =   4830
      MouseIcon       =   "RipperProgramMainForm.frx":38A6
      MousePointer    =   99  'Custom
      OLEDropMode     =   1  'Manual
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6150
      Width           =   1515
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5850
      Top             =   390
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open File"
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "Search File"
      Height          =   315
      Left            =   4830
      MouseIcon       =   "RipperProgramMainForm.frx":39F8
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   60
      Width           =   1515
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5010
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      Picture         =   "RipperProgramMainForm.frx":3B4A
      ScaleHeight     =   375
      ScaleWidth      =   1185
      TabIndex        =   33
      Top             =   1950
      Width           =   1185
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Ea Fsh"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2490
      TabIndex        =   82
      Top             =   2580
      Width           =   1395
   End
   Begin VB.Label Result1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Index           =   19
      Left            =   2010
      TabIndex        =   81
      Top             =   2580
      Width           =   405
   End
   Begin VB.Label TimeLabel 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   810
      TabIndex        =   78
      Top             =   6180
      Width           =   4005
   End
   Begin VB.Label Result1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Index           =   18
      Left            =   2010
      TabIndex        =   74
      Top             =   2340
      Width           =   405
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Tif"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2490
      TabIndex        =   73
      Top             =   2340
      Width           =   1395
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Swf"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2490
      TabIndex        =   70
      Top             =   2100
      Width           =   1395
   End
   Begin VB.Label Result1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Index           =   17
      Left            =   2010
      TabIndex        =   69
      Top             =   2100
      Width           =   405
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Search using all different options"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2160
      TabIndex        =   67
      Top             =   3810
      Width           =   2565
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Filter text (Remove junk)"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   930
      TabIndex        =   65
      Top             =   4860
      Width           =   3645
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"RipperProgramMainForm.frx":4072
      ForeColor       =   &H00FFFF00&
      Height          =   975
      Left            =   390
      TabIndex        =   62
      Top             =   5100
      Width           =   5655
   End
   Begin VB.Label Result1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Index           =   16
      Left            =   2010
      TabIndex        =   60
      Top             =   1860
      Width           =   405
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Avi"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2490
      TabIndex        =   59
      Top             =   1860
      Width           =   1395
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Bink"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2490
      TabIndex        =   57
      Top             =   1620
      Width           =   1395
   End
   Begin VB.Label Result1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Index           =   15
      Left            =   2010
      TabIndex        =   56
      Top             =   1620
      Width           =   405
   End
   Begin VB.Label Result1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Index           =   14
      Left            =   2010
      TabIndex        =   54
      Top             =   1380
      Width           =   405
   End
   Begin VB.Label Result1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Index           =   13
      Left            =   2010
      TabIndex        =   53
      Top             =   1140
      Width           =   405
   End
   Begin VB.Label Result1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Index           =   12
      Left            =   2010
      TabIndex        =   52
      Top             =   900
      Width           =   405
   End
   Begin VB.Label Result1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Index           =   11
      Left            =   300
      TabIndex        =   51
      Top             =   3540
      Width           =   405
   End
   Begin VB.Label Result1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Index           =   10
      Left            =   300
      TabIndex        =   50
      Top             =   3300
      Width           =   405
   End
   Begin VB.Label Result1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Index           =   9
      Left            =   300
      TabIndex        =   49
      Top             =   3060
      Width           =   405
   End
   Begin VB.Label Result1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Index           =   8
      Left            =   300
      TabIndex        =   48
      Top             =   2820
      Width           =   405
   End
   Begin VB.Label Result1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Index           =   7
      Left            =   300
      TabIndex        =   47
      Top             =   2580
      Width           =   405
   End
   Begin VB.Label Result1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Index           =   6
      Left            =   300
      TabIndex        =   46
      Top             =   2340
      Width           =   405
   End
   Begin VB.Label Result1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Index           =   5
      Left            =   300
      TabIndex        =   45
      Top             =   2100
      Width           =   405
   End
   Begin VB.Label Result1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Index           =   4
      Left            =   300
      TabIndex        =   44
      Top             =   1860
      Width           =   405
   End
   Begin VB.Label Result1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Index           =   3
      Left            =   300
      TabIndex        =   43
      Top             =   1620
      Width           =   405
   End
   Begin VB.Label Result1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Index           =   2
      Left            =   300
      TabIndex        =   42
      Top             =   1380
      Width           =   405
   End
   Begin VB.Label Result1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Index           =   1
      Left            =   300
      TabIndex        =   41
      Top             =   1140
      Width           =   405
   End
   Begin VB.Label Result1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Index           =   0
      Left            =   300
      TabIndex        =   40
      Top             =   900
      Width           =   405
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Html"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2490
      TabIndex        =   39
      Top             =   1380
      Width           =   1395
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Acrobat (Pdf)"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2490
      TabIndex        =   37
      Top             =   1140
      Width           =   1395
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Wave"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2490
      TabIndex        =   35
      Top             =   900
      Width           =   1395
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Select all"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   360
      TabIndex        =   32
      Top             =   3810
      Width           =   4125
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Truevision Targa"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   780
      TabIndex        =   30
      Top             =   3540
      Width           =   1395
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Delete corrupt or incomplete Bmp and Gif files"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   360
      TabIndex        =   28
      Top             =   4110
      Width           =   4125
   End
   Begin VB.Image Image2 
      Height          =   1500
      Left            =   4830
      Picture         =   "RipperProgramMainForm.frx":4188
      Top             =   2790
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Bitmap 16"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   780
      TabIndex        =   26
      Top             =   3300
      Width           =   1065
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Bitmap 256"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   780
      TabIndex        =   24
      Top             =   3060
      Width           =   1065
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Mus"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   780
      TabIndex        =   22
      Top             =   2820
      Width           =   825
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Xmi"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   780
      TabIndex        =   20
      Top             =   2580
      Width           =   825
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Psp"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   780
      TabIndex        =   18
      Top             =   2340
      Width           =   825
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Hmp - Hmi"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   780
      TabIndex        =   16
      Top             =   2100
      Width           =   825
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Zip"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   780
      TabIndex        =   14
      Top             =   1860
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Midi"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   780
      TabIndex        =   12
      Top             =   1620
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Jpeg"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   780
      TabIndex        =   10
      Top             =   1380
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Gif"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   780
      TabIndex        =   8
      Top             =   1140
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Suppress Messages"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   4410
      Width           =   2085
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Bitmap 64K"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   780
      TabIndex        =   3
      Top             =   900
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Search For:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   630
      TabIndex        =   2
      Top             =   570
      Width           =   4005
   End
   Begin VB.Image Image1 
      Height          =   4560
      Left            =   90
      Picture         =   "RipperProgramMainForm.frx":5961
      Top             =   120
      Width           =   4635
   End
   Begin VB.Image Image3 
      Height          =   1440
      Left            =   0
      Picture         =   "RipperProgramMainForm.frx":BF88
      Top             =   4680
      Width           =   6465
   End
End
Attribute VB_Name = "RipperProgramMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------
'File Ripper - By Alex M (Not Fully Finished Comments)
'---------------------------
'This is varibles that store information about what you
'have selected (i.e. to add twice the height value to the
'bitmaps file length)
Dim AddHeight As Integer
Dim AddHeight2 As Integer
'Used mainly for loops and used to keep track of what
'position you are within a file when scanning it
Dim I As Long
Dim W As Long
Dim X As Long
Dim Y As Long
Dim Z As Long
'Used to keep track of what methods or extraction you are
'up to (i.e. whether your extracting .bmp or .hmi files)
Dim P%
Dim R%
'used to position picture boxes on the form
Dim X1%
Dim X2%
Dim X3%
Dim Y2%
'Store Whether or not you want to proceed with replacing
'the difference in file size with null charaters when
'replacing a file
Dim Doit%
'used to figure out how many times to find %%eof within a
'.pdf file before saying it is the enf of the file
Dim Times%
'used to store the ascii value of a character while
'searching a file
Dim Atemp%
'How many times you replaced a file within another file
Dim Patch%
'Store Whether or not you want to proceed with replacing
'the difference in file size with null charaters when
'replacing a file
Dim FitFile%
'if you want to add to extra bytes to a gif file or not
Dim ExtraGif%
'The extension of the file you are searching for or
'replacing
Dim Ext1$
'As it suggests, it is just temporary storage for values
Dim Atemp2$
'used to store the selected file to scan
Dim InFile$
Dim InFile2$
Dim InFile4$
'I think (it was a long time ago I made this program) that
'this is used when you drag and drop a viv file onto this
'program, it automatically extracts the viv file
Dim Comamd$
'the name of the file to open (I don't know why I called
'it extension!!
Dim Extension$
'The password to use for the encryption
Dim Password1$
'used to extimate the time for encryption
Dim TheOldTime$
'just temp storage
Dim USer&
'add extra bytes or not to gif files
Dim Extra&
'used for mus files, it is explained why at the mus
'subroutine
Dim Chance&
'height and width of a picture as specified in the header
Dim Width1&
Dim Height1&
'Well, You can probably guess what I use this for
'The length of the file to be loaded
Dim TheFileLength&

'This subroutine opens the file that the user has selected
Public Sub Do_Open()
'Record filelength (also for later use)
TheFileLength& = FileLen(CommonDialog1.FileName)
'set the string to the size of the file (needed for binary
'open)
InFile$ = Space(TheFileLength&)
InFile2$ = ""
'Open the selected file (binary open is the fastest method
'I know to access files) The Access Read Lock Read is just
'(hopefully) locking the file that you open so no other
'programs can access it
Open CommonDialog1.FileName For Binary Access Read Lock Read As #1
    Get #1, 1, InFile$
Close
End Sub

'When scanning Adobe .pdf files, sometimes they contain
'%%eof halfway through the file without any apparent
'reason. The user has a choice of whether to compensate
'for it or not when extracting the files
Private Sub Acrobat_Click()
If Acrobat.Value = 1 Then
    'ask the user in a YesNo dialog
    Temp$ = MsgBox("Some acrobat files have two '%%EOF' in them. I think they do this to stop people from modifying their work. Do you want to stop at the first '%%EOF'? (Default is YES) ", vbYesNo, "Double EOF's")
    Times% = 0
    If Temp$ = vbYes Then Times% = 1
End If
End Sub

'Scan a file with all possible combinations of options
'selected. This is recommended if you want the best
'possible chance of finding a file
Private Sub AllOptions_Click()
If AllOptions.Value = 1 Then
    Mess ("All possible options will be tested with the selected file formats regardless of the options you have currently selected, But the mus converter success rate is still the same as what you have selected.")
End If
End Sub

'For some reason, Some bitmaps seem to be bigger than
'others. I have had a few suggestions, but the best
'solution so far is to ask the user what to do
Private Sub Bmp16_Click()
If Bmp16.Value = 1 Then
    'ask the user in a YesNo dialog
    Temp$ = MsgBox("Sometimes Bitmaps have extra data on the end. I still haven't figured what the pattern is but I do know that some have tripple the height added on the them. It doesn't affect the bitmap's appearence, but it is essential that you guess the correct length or the bitmap will be corrupted. Do you want extra data?", vbYesNo, "Extra Data")
    AddHeight2 = 0
    If Temp$ = vbYes Then AddHeight2 = 1
End If
End Sub

'For some reason, Some bitmaps seem to be bigger than
'others. I have had a few suggestions, but the best
'solution so far is to ask the user what to do
Private Sub Bmp256_Click()
If Bmp256.Value = 1 Then
    'ask the user in a YesNo dialog
    Temp$ = MsgBox("Sometimes Bitmaps have extra data on the end. I still haven't figured what the pattern is but I do know that some have double the height added on the them. It doesn't affect the bitmap's appearence, but it is essential that you guess the correct length or the bitmap will be corrupted. Do you want extra data?", vbYesNo, "Extra Data")
    AddHeight = 0
    If Temp$ = vbYes Then AddHeight = 1
End If
End Sub

'This generally updates the display and sets some
'important variables to their default value
Private Sub SetNullS()
'Send data to the user telling them what the program
'has found for each file format
Do_Result
'Update number of files found on the form
Result1(P%).Caption = Mid$(Str$(Z), 2, 100)
'set a few varibles to zero
W = 0
'I think this keeps track of which file format
'searching method it is up to
P% = P% + 1
Z = 0
Label1.Caption = "Scanning File..."
'redraws the form
Me.Refresh
End Sub

'This is when you click search file, it will search the
'file for other files within it
Private Sub Command1_Click()
'on error display 'ripping complete but with errors'
On Error GoTo 10
'if there is no specified file, then open the common
'dialog (this is neccesary so do not delete this line
'and put just "CommonDialog1.ShowOpen")
If CommonDialog1.FileName = "" Then CommonDialog1.ShowOpen
'set the file 'format up to' to zero
P% = 0
'stop the animation of the ripper image
Timer1.Enabled = False
Label1.Caption = "Reading File..."
'set the label's captions to null, they will be updated
'with number of files found when each format is scanned
For T = 0 To (Result1.Count - 1)
    Result1(T).Caption = "--"
Next T
'Refresh form
Me.Refresh
'call the Do_open subroutine, opens the selected file
'to scan
Do_Open
'if you have selected all options, then first scan these
'options
If AllOptions.Value = 1 Then
    ExtraGif% = 0
    Times% = 1
    AddHeight = 1
    AddHeight2 = 1
    Extra& = 20
End If
'Call each subroutine to scan for these files
If Bitmap.Value = 1 Then RipBitmap
'sets the default values of important variables
SetNullS
If Gif.Value = 1 Then RipGif
SetNullS
If Jpeg.Value = 1 Then RipJpeg
SetNullS
If Jpeg.Value = 1 Then
    RipJpeg2
    P% = P% - 1
End If
SetNullS
GoTo 7889
If Jpeg.Value = 1 Then
    RipJpeg3
    P% = P% - 1
End If
SetNullS
7889
If Midi.Value = 1 Then ripMidi
SetNullS
If Zip.Value = 1 Then RipZip
SetNullS
If Hmp.Value = 1 Then RipHmp
SetNullS
If Paint.Value = 1 Then RipPsp
SetNullS
If Xmi.Value = 1 Then RipXmi
SetNullS
If Mus.Value = 1 Then RipMus
SetNullS
If Bmp256.Value = 1 Then RipBitmap256
SetNullS
If Bmp16.Value = 1 Then RipBitmap16
SetNullS
If Tga.Value = 1 Then RipTga
SetNullS
If Wave.Value = 1 Then RipWave
SetNullS
If Acrobat.Value = 1 Then RipPdf
SetNullS
If Html.Value = 1 Then RipHtml
SetNullS
If Bink.Value = 1 Then RipBink
SetNullS
If Avi.Value = 1 Then RipAvi
SetNullS
If Swf.Value = 1 Then RipSWF
SetNullS
If Tiff.Value = 1 Then RipTif
SetNullS
If Fsh.Value = 1 Then RipFsh
SetNullS
'If you have selected scan with all options, then scan
'the other options that the program has not scanned yet
If AllOptions.Value = 1 Then
    ExtraGif% = 10
    Times% = 0
    AddHeight = 0
    AddHeight2 = 0
    Extra& = 100000
    
    Result1(13).Caption = "--"
    For H1 = 8 To 10
    Result1(H1).Caption = "--"
    Next H1
    Result1(1).Caption = "--"
    Me.Refresh
    
    If Acrobat.Value = 1 Then RipPdf
    P% = 13
    SetNullS
    P% = 8
    If Mus.Value = 1 Then RipMus
    SetNullS
    If Bmp256.Value = 1 Then RipBitmap256
    SetNullS
    If Bmp16.Value = 1 Then RipBitmap16
    SetNullS
    P% = 1
    If Gif.Value = 1 Then RipGif
    SetNullS
End If
'Display ripping complete
Mess ("Ripping Complete!!")
GoTo 20
10
'display ripping complete but with some errors
Mess ("Ripping Complete!! But there was some errors.")
20
'set the commondialog filename to null and resume the
'ripper image animation
CommonDialog1.FileName = ""
Timer1.Enabled = True
Label1.Caption = "Search For:-"
End Sub

'This subroutine extracts file from a viv file
Private Sub Command10_Click()
On Error Resume Next
'set some default values and get the file to be opened
CommonDialog1.ShowOpen
P% = 0
Timer1.Enabled = False
Label1.Caption = "Reading File..."
Me.Refresh
'Open the selected file to a string "Infile1$"
Do_Open
'call the subroutine to extract the files from the viv
DoStuffViv
'set values back to normal
CommonDialog1.FileName = ""
Label1.Caption = "Search For:-"
End Sub

Public Sub DoStuffViv()
On Error Resume Next
Randomize Timer
Tempstring11$ = Str$(Int((Rnd * 1000)))
MkDir ("C:\RippViv")
MkDir ("C:\RippViv\" + Tempstring11$)
'if the file does not appear to be a viv then ask user
'what to do
If Mid$(InFile$, 1, 4) <> "BIGF" Then Call MsgBox("The File Doesn't appear to be a viv, But Continuing anyway", vbInformation, "Error")

'Read length of file
ATremp2$ = Hex(Asc(Mid(InFile$, 5, 1))) + Hex(Asc(Mid(InFile$, 6, 1))) + Hex(Asc(Mid(InFile$, 7, 1))) + Hex(Asc(Mid(InFile$, 8, 1)))
ConvHexToDec (ATremp2$)
lengthoffile& = Height1&

'Read Amount of Files in viv
ATremp2$ = Hex(Asc(Mid(InFile$, 14, 1))) + Hex(Asc(Mid(InFile$, 13, 1))) + Hex(Asc(Mid(InFile$, 12, 1)))
ConvHexToDec (ATremp2$)
amountoffile& = Height1&

'Read End of header
ATremp2$ = Hex(Asc(Mid(InFile$, 18, 1))) + Hex(Asc(Mid(InFile$, 17, 1))) + Hex(Asc(Mid(InFile$, 16, 1))) + Hex(Asc(Mid(InFile$, 15, 1)))
ConvHexToDec (ATremp2$)
endofheader& = Height1&

'-------------------------------------------
'Obsolete data
'-------------------------------------------
'Atewmp = MsgBox("Is this a Nfs5/6 Viv File", vbYesNo, "nfs5/6")
'MsgBox (lengthoffile&)
'MsgBox (amountoffile&)
'MsgBox (endofheader&)
'-------------------------------------------
Label1.Caption = "Extracting" + Str$(amountoffile&) + " Files"
'-------------------------------------------
'Obsolete data
'-------------------------------------------
'19 start of file
'-------------------------------------------
'When Found File
I = 17
For J = 1 To amountoffile&
    Atemp2$ = ""
    'Get Place In File in Bytes
    Atemp2$ = Space(2 - Len(Hex$(Asc(Mid$(InFile$, I, 1))))) + Hex$(Asc(Mid$(InFile$, I, 1)))
    Atemp2$ = Atemp2$ + Space(2 - Len(Hex$(Asc(Mid$(InFile$, I + 1, 1))))) + Hex$(Asc(Mid$(InFile$, I + 1, 1)))
    Atemp2$ = Atemp2$ + Space(2 - Len(Hex$(Asc(Mid$(InFile$, I + 2, 1))))) + Hex$(Asc(Mid$(InFile$, I + 2, 1)))
    Atemp2$ = Atemp2$ + Space(2 - Len(Hex$(Asc(Mid$(InFile$, I + 3, 1))))) + Hex$(Asc(Mid$(InFile$, I + 3, 1)))
    ConvHexToDec2 (UCase$(Atemp2$))
    placeinfile& = Height1&
    
    Atemp2$ = ""
    'Get Size of File in Bytes
    Atemp2$ = Space(2 - Len(Hex$(Asc(Mid$(InFile$, I + 4, 1))))) + Hex$(Asc(Mid$(InFile$, I + 4, 1)))
    Atemp2$ = Atemp2$ + Space(2 - Len(Hex$(Asc(Mid$(InFile$, I + 5, 1))))) + Hex$(Asc(Mid$(InFile$, I + 5, 1)))
    Atemp2$ = Atemp2$ + Space(2 - Len(Hex$(Asc(Mid$(InFile$, I + 6, 1))))) + Hex$(Asc(Mid$(InFile$, I + 6, 1)))
    Atemp2$ = Atemp2$ + Space(2 - Len(Hex$(Asc(Mid$(InFile$, I + 7, 1))))) + Hex$(Asc(Mid$(InFile$, I + 7, 1)))
    ConvHexToDec2 (UCase$(Atemp2$))
    sizeoffile& = Height1&
    I = I + 8
    K = 0
1
    K = K + 1
    If Asc(Mid$(InFile$, I, 1)) = "0" Then GoTo 2
    I = I + 1
    GoTo 1
2
    'Get Name of File
    Extension$ = "RippViv\" + Tempstring11$ + "\" + Mid$(InFile$, I - (K - 1), K)
    '-------------------------------------------
    'MsgBox (Extension$)
    'MsgBox (placeinfile&)
    'MsgBox (sizeoffile&)
    '-------------------------------------------
    I = I + 1
    X = placeinfile& + 1
    Y = sizeoffile&
    'Extract/save File
    RippItNo3
Next J
'Display done and reset commondialog filename
MsgBox ("Done")
MsgBox (Str$(amountoffile&) + " Files Extracted")
CommonDialog1.FileName = ""
End Sub

'Used for extracting file from a viv file
Public Sub RippItNo3()
'used to determine whether you are searching for the start
'of a file instead of the end
W = 0
'Z is used to count the amout of files found
Z = Z + 1
'Get the file to be extracted, saved to a string
a$ = Mid$(InFile$, X, Y)
'Save the data to a file
Open Extension$ For Output As #1
Print #1, a$;
Close
End Sub

'Exit the program
Private Sub Command2_Click()
End
End Sub

'Ripps a bitmap with a depth of 256 colours
Public Sub RipBitmap256()
On Error Resume Next
'Sets the extension to use
Extension$ = "Bmp"
'Got through the file byte by byte
For I = 1 To Len(InFile$)
    'search for characters 'BM' in the file, it is set
    'out like this because It is faster (think about
    'it, it searches for a B and if it finds a B then
    'it sees if the next charater is an M, but
    'it was originally set to search for BM both at
    'the same time, much slower!!)
    If Mid$(InFile$, I, 1) = "B" Then
        If Mid$(InFile$, I + 1, 1) = "M" Then
            X = I
            'finds the width (gets the bytes from
            'position 18-21) in Hex then converts it
            'to decimal
            Atemp2$ = Hex(Asc(Mid$(InFile$, I + 21, 1))) + Hex(Asc(Mid$(InFile$, I + 20, 1))) + Hex(Asc(Mid$(InFile$, I + 19, 1))) + Hex(Asc(Mid$(InFile$, I + 18, 1)))
            ConvHexToDec (UCase$(Atemp2$))
            'Because the ConvHexToDec saves the decimal
            'number to Height1&, Therefore it makes
            'Height1& equal to width1&
            Width1& = Height1&
            'finds the width (gets the bytes from
            'position 22-25) in Hex then converts it
            'to decimal
            Atemp2$ = Hex(Asc(Mid$(InFile$, I + 25, 1))) + Hex(Asc(Mid$(InFile$, I + 24, 1))) + Hex(Asc(Mid$(InFile$, I + 23, 1))) + Hex(Asc(Mid$(InFile$, I + 22, 1)))
            ConvHexToDec (UCase$(Atemp2$))
            'if the user selects that they want to add
            'the height, then add the height to the file
            'size
            If AddHeight = 1 Then
                f% = Height1& * 2
            Else
                f% = 0
            End If
            'Calculate the size of the file, the size
            'of the file will then be (Y-X)bytes
            Y = (Height1& * Width1&) + X + 1078 + f%
            'Saves the bitmap to file
            RippBmp
        End If
    End If
Next I
End Sub

'This sub is used to replace a file within another file
'and also used to save other images and determine whether
'they are corrupt or not
Private Sub RippBmp()
On Error GoTo 10
Randomize Timer
'used to determine whether you are searching for the start
'of a file instead of the end
W = 0
'Z is used to count the amout of files found
Z = Z + 1
'Get the file to be extracted, saved to a string
a$ = Mid$(InFile$, X, (Y - X))
'Used to randomize the name of the file
R2D2% = Rnd * (1000)
'if you select that you want to replace a file, then
'continue
If Doit% = 1 Then
    'If the file is the same size as the file that has been
    'found then
    If Y - X = Len(InFile4$) And UCase$(Extension$) = UCase$(Ext1$) Then
        'replace the file with the file that the user has
        'selected (If and only if it is a perfect match)
        InFile$ = Mid$(InFile$, 1, X - 1) + InFile4$ + Mid$(InFile$, Y, Len(InFile$))
        Patch% = Patch% + 1
        Me.Refresh
    Else
        'If the file the user has selected is smaller than
        'The one that they want replaced, then ask if they
        'want the difference in the file sizes replaced
        'by null characters (i.e. chr$(0))
        If UCase$(Extension$) = UCase$(Ext1$) And Y - X > Len(InFile4$) And FitFile% = 0 Then
            atemprrr$ = MsgBox("Do you want to replace the" + Str$(Y - X) + " Byte file with your" + Str$(Len(InFile4$)) + " Byte file, the missing bytes will be replaced by zeros but may corrupt the exe and/or the file that you are adding. Replace the file?", vbYesNo, "Replace?")
            If atemprrr$ = vbYes Then
                'replace the file with the file that the
                'user has selected
                TempstringYY$ = Mid$(SubStringZeros$, 1, (Y - X - Len(InFile4$)))
                InFile$ = Mid$(InFile$, 1, X - 1) + InFile4$ + TempstringYY$ + Mid$(InFile$, Y, Len(InFile$))
                Patch% = Patch% + 1
            End If
        End If
    End If
    Exit Sub
End If
'Write file to disk
Open "C:\" + Mid$(Str$(R2D2%), 2, 1000) + Extension$ + Mid$(Str$(Z), 2, 1000) + "." + Extension$ For Output As #1
Print #1, a$;
Close #1
'If an error occurs here, it will automatically assume that
'the picture is a corrupt image
If Deletthem.Value = 1 Then Image2.Picture = LoadPicture("C:\" + Mid$(Str$(R2D2%), 2, 1000) + Extension$ + Mid$(Str$(Z), 2, 1000) + "." + Extension$)
GoTo 20
10
Z = Z - 1
'delete the corrupt image (if the option is selected)
If Deletthem.Value = 1 Then
    Kill ("C:\" + Mid$(Str$(R2D2%), 2, 1000) + Extension$ + Mid$(Str$(Z + 1), 2, 1000) + "." + Extension$)
End If
20
End Sub

'Ripps a bitmap with a depth of 16 colours
Public Sub RipBitmap16()
On Error Resume Next
'Sets the extension to use
Extension$ = "Bmp"
'Got through the file byte by byte
For I = 1 To Len(InFile$)
    If Mid$(InFile$, I, 1) = "B" Then
        If Mid$(InFile$, I + 1, 1) = "M" Then
            'finds the width (gets the bytes from
            'position 18-21) in Hex then converts it
            'to decimal
            X = I
            Atemp2$ = Hex(Asc(Mid$(InFile$, I + 21, 1))) + Hex(Asc(Mid$(InFile$, I + 20, 1))) + Hex(Asc(Mid$(InFile$, I + 19, 1))) + Hex(Asc(Mid$(InFile$, I + 18, 1)))
            ConvHexToDec (UCase$(Atemp2$))
            'Because the ConvHexToDec saves the decimal
            'number to Height1&, Therefore it makes
            'Height1& equal to width1&
            Width1& = Height1&
            'finds the width (gets the bytes from
            'position 22-25) in Hex then converts it
            'to decimal
            Atemp2$ = Hex(Asc(Mid$(InFile$, I + 25, 1))) + Hex(Asc(Mid$(InFile$, I + 24, 1))) + Hex(Asc(Mid$(InFile$, I + 23, 1))) + Hex(Asc(Mid$(InFile$, I + 22, 1)))
            ConvHexToDec (UCase$(Atemp2$))
            'if the user selects that they want to add
            'the height, then add the 3*height to the
            'file size
            If AddHeight = 1 Then
                f% = Height1& * 3
            Else
                f% = 0
            End If
            'Calculate the size of the file, the size
            'of the file will then be (Y-X)bytes
            Y = ((Height1& * Width1&) / 2) + X + 118 + f%
            'Saves the bitmap to file
            RippBmp
        End If
    End If
Next I
End Sub

'Ripps a bitmap with 24-bit colours
Public Sub RipBitmap()
On Error Resume Next
'Sets the extension to use
Extension$ = "Bmp"
'Got through the file byte by byte
For I = 1 To Len(InFile$)
    If Mid$(InFile$, I, 1) = "B" Then
        If Mid$(InFile$, I + 1, 1) = "M" Then
            'finds the width (gets the bytes from
            'position 18-21) in Hex then converts it
            'to decimal
            Atemp2$ = Hex(Asc(Mid$(InFile$, I + 21, 1))) + Hex(Asc(Mid$(InFile$, I + 20, 1))) + Hex(Asc(Mid$(InFile$, I + 19, 1))) + Hex(Asc(Mid$(InFile$, I + 18, 1)))
            ConvHexToDec (UCase$(Atemp2$))
            'Because the ConvHexToDec saves the decimal
            'number to Height1&, Therefore it makes
            'Height1& equal to width1&
            Width1& = Height1&
            'finds the width (gets the bytes from
            'position 22-25) in Hex then converts it
            'to decimal
            Atemp2$ = Hex(Asc(Mid$(InFile$, I + 25, 1))) + Hex(Asc(Mid$(InFile$, I + 24, 1))) + Hex(Asc(Mid$(InFile$, I + 23, 1))) + Hex(Asc(Mid$(InFile$, I + 22, 1)))
            ConvHexToDec (UCase$(Atemp2$))
            X = I
            'Calculate the size of the file, the size
            'of the file will then be (Y-X)bytes
            Y = ((Height1& * Width1&) * 3 + 54) + X
            'Saves the bitmap to file
            RippBmp
        End If
    End If
Next I
End Sub

Public Sub RippIt()
Randomize Timer
W = 0

Z = Z + 1
a$ = Mid$(InFile$, X, (Y - X))
R2D2% = Rnd * (10000)

If Doit% = 1 Then
    If Y - X = Len(InFile4$) And UCase$(Extension$) = UCase$(Ext1$) Then
        InFile$ = Mid$(InFile$, 1, X - 1) + InFile4$ + Mid$(InFile$, Y, Len(InFile$))
        Patch% = Patch% + 1
    Else
        If UCase$(Extension$) = UCase$(Ext1$) And Y - X > Len(InFile4$) And FitFile% = 0 Then
            atemprrr$ = MsgBox("Do you want to replace the" + Str$(Y - X) + " Byte file with your" + Str$(Len(InFile4$)) + " Byte file, the missing bytes will be replaced by zeros but may corrupt the exe and/or the file that you are adding. Replace the file?", vbYesNo, "Replace?")
            If atemprrr$ = vbYes Then
                TempstringYY$ = Mid$(SubStringZeros$, 1, (Y - X - Len(InFile4$)))
                InFile$ = Mid$(InFile$, 1, X - 1) + InFile4$ + TempstringYY$ + Mid$(InFile$, Y, Len(InFile$))
                Patch% = Patch% + 1
            End If
        End If
    End If
    Exit Sub
End If

Open "C:\" + Mid$(Str$(R2D2%), 2, 1000) + Extension$ + Mid$(Str$(Z), 2, 1000) + "." + Extension$ For Output As #1
Print #1, a$;
Close

Result1(P%).Caption = Mid$(Str$(Z), 2, 100)
Me.Refresh
End Sub

Public Sub Do_Result()
If Messages <> 1 Then Mess ("There was" + Str$(Z) + " " + Extension$ + " file(s) Ripped from " + CommonDialog1.FileName)
End Sub

Public Sub Mess(Message As String)
Call MsgBox(Message, vbInformation, "Ripper")
End Sub

Public Sub RipGif()
On Error Resume Next

Extension$ = "Gif"

For I = 1 To Len(InFile$)
    If W = 1 Then
            If Mid$(InFile$, I, 2) = Chr$(0) + Chr$(&H3B) Then
                Y = I + 1 + ExtraGif%
                RippBmp
                W = 0
            End If
    Else
        If Mid$(InFile$, I, 1) = "G" Then
            If Mid$(InFile$, I + 1, 2) = "IF" Then
                X = I
                W = 1
            End If
        End If
    End If
Next I

End Sub

Private Sub Command3_Click()
Mess ("HELP - Click on the filetypes you wish to search and click search file. The 'Search Exe 4 Text' button is seperate to the ripper, It can extract text (or Passwords) from compiled Visual Basic programs (If the password is in the EXE). - ABOUT - Created by Alex Murray. This is a updated version of my original Ripper. A few bugs have been fixed and searching times are alot faster. Made in VB. If you have any questions or problems please email me at Alex_Murray1@Hotmail.com. To get an idea of how long a big file will take to scan, my 1.1Ghz took 1.5Mins to do 12 files types in a 10MB file but times will go up exponentially from there (30Mb file took 10.5mins).")
End Sub

Private Sub Command4_Click()
On Error GoTo 10
If CommonDialog1.FileName = "" Then CommonDialog1.ShowOpen

Do_Open
W& = 0

For Y = 1 To Len(InFile$)
    If W& = 0 Then
        If Mid$(InFile$, Y, 1) = Chr$(0) Then
            If Mid$(InFile$, Y + 1, 1) <> Chr$(0) Then
                InFile2$ = InFile2$ + Mid$(InFile$, Y + 1, 1)
                W& = 1
                Y = Y + 1
            End If
        End If
    Else
        If Mid$(InFile$, Y, 1) = Chr$(0) Then
            If Mid$(InFile$, Y + 1, 1) <> Chr$(0) Then
                InFile2$ = InFile2$ + Mid$(InFile$, Y + 1, 1)
                Y = Y + 1
            End If
        Else
            InFile2$ = InFile2$ + Chr$(13) + Chr$(10)
            W = 0
        End If
    End If
Next Y
Randomize Timer
Temp% = (Rnd * (1000))

If Junk.Value = 1 Then
    For R1 = 1 To Len(InFile2$)
        Atemp% = Asc(Mid$(InFile2$, R1, 1))
        If Atemp% < 32 Or Atemp% > 126 Then
            InFile2$ = Mid$(InFile2$, 1, (R1 - 1)) + "," + Mid$(InFile2$, R1 + 1, Len(InFile2$))
        End If
    Next R1
End If

Open "C:\WordOut" + Str$(Temp%) + ".txt" For Output As #3
Print #3, InFile2$;
Close
    
If Junk.Value = 1 Then
    Open "C:\WordOut" + Str$(Temp%) + ".txt" For Input As #3
    Open "C:\WordOut" + Str$(Temp%) + ".txt2" For Output As #2
    
    While Not EOF(3)
        Input #3, Atemp2$
        If Len(Atemp2$) < 3 Then
        Else
            Print #2, Atemp2$
        End If
    Wend
    Close
    
    FileCopy "C:\WordOut" + Str$(Temp%) + ".txt2", "C:\WordOut" + Str$(Temp%) + ".txt"
    Kill ("C:\WordOut" + Str$(Temp%) + ".txt2")
End If

Mess ("Text search is complete")
GoTo 20
10
Mess ("There was some errors")
20
CommonDialog1.FileName = ""
End Sub

Private Sub Command5_Click()
On Error GoTo 10
Mess ("Please select a the file that you want to try to insert")
If CommonDialog1.FileName = "" Then CommonDialog1.ShowOpen
Patch% = 0

Ext1$ = Mid$(CommonDialog1.FileName, Len(CommonDialog1.FileName) - 2, 3)

atemprrr$ = MsgBox("Do you want to try and fit a smaller file in than the one that is already there (E.g. replace a 10K Jpeg with a 5K Jpeg)", vbYesNo, "Prompt")
If atemprrr$ = vbYes Then
    FitFile% = 0
Else
    FitFile% = 1
End If

Doit% = 1

Mess ("Select the file that your going to patch")
Label1.Caption = "Gathering Data..."
Me.Refresh

Module1.TypeZeros

Do_Open
InFile4$ = InFile$
CommonDialog1.FileName = ""
Command1_Click
        
Open "c:\Patched" + CommonDialog1.FileTitle For Output As #5
    Print #5, InFile$;
Close #5

SubStringZeros$ = ""

Mess (Str$(Patch%) + " Replaces were Successful")
Mess ("Don't loose your old file. You may want to zip it up and copy it some where incase you want to undo the changes")

Doit% = 0
GoTo 20
10
Mess ("There was some errors")
20
CommonDialog1.FileName = ""
End Sub

Private Sub Command8_Click()
End
End Sub

Private Sub Command9_Click()
On Error GoTo 10
Dim Temp455$
Dim Temp456$
Dim Temp457$
Dim Temp458$
Dim WHen&

Temp455$ = InputBox("Type the text you wish to search for", "Search Text", "")
Temp456$ = InputBox("Type the text you wish to replace it with", "Replace Text", "")
If Len(Temp455$) <> Len(Temp456$) Then
    If Len(Temp455$) > Len(Temp456$) Then
        Mess ("Replacement string longer than Search string")
        Exit Sub
    End If
    
    temp468$ = MsgBox("Lengths of strings do not match!!!, Do you want to add spaces??", vbYesNo, "Stings do not match")
    
    If temp468$ = vbYes Then
        Temp456$ = Temp456$ + Space(Len(Temp455$) - Len(Temp456$))
    Else
        Exit Sub
    End If
End If

CommonDialog1.FileName = ""
If CommonDialog1.FileName = "" Then CommonDialog1.ShowOpen

Do_Open
W& = 0

For Ty = 1 To Len(Temp455$)
    Temp458$ = Temp458$ + Mid$(Temp455$, Ty, 1) + Chr$(0)
Next Ty
Temp458$ = Mid$(Temp458$, 1, Len(Temp458$) - 1)

For Ty = 1 To Len(Temp455$)
    Temp457$ = Temp457$ + Mid$(Temp456$, Ty, 1) + Chr$(0)
Next Ty
Temp457$ = Mid$(Temp457$, 1, Len(Temp457$) - 1)

For Y = 1 To Len(InFile$)
    If UCase$(Mid$(InFile$, Y, Len(Temp457$))) = UCase$(Temp458$) Then
        WHen& = WHen& + 1
        InFile$ = Mid$(InFile$, 1, Y - 1) + Temp457$ + Mid$(InFile$, Y + Len(Temp457$), Len(InFile$))
    End If
    If UCase$(Mid$(InFile$, Y, Len(Temp455$))) = UCase$(Temp455$) Then
        WHen& = WHen& + 1
        InFile$ = Mid$(InFile$, 1, Y - 1) + Temp456$ + Mid$(InFile$, Y + Len(Temp456$), Len(InFile$))
    End If
Next Y

Randomize Timer
Temp% = (Rnd * (1000))


Open "C:\Out.exe" + Str$(Temp%) + ".exe" For Output As #3
Print #3, InFile$;
Close

Mess ("Text replace is complete," + Str$(WHen&) + " Replacements where made!!")
GoTo 20
10
Mess ("There was some errors")
20
CommonDialog1.FileName = ""
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
If X > 5220 Then X = 5200
If Y > 3600 Then Y = 3600
If X < 4830 Then X = 4900
If Y < 720 Then Y = 720
Picture1.Left = X
Picture1.Top = Y
End Sub

Private Sub Form_Load()
On Error Resume Next
X1% = 25
X2% = 25
Me.Height = 6540
R% = 25

If Mid$(Command$, 1, 1) = Chr$(34) Then
    Comamd$ = Mid$(Command$, 2, Len(Command$) - 2)
Else
    Comamd$ = Command$
End If

TempTemp$ = LCase(Mid$(Comamd$, Len(Comamd$) - 2, 3))
If TempTemp$ = "viv" Then
        CommonDialog1.FileName = Comamd$
        P% = 0
        Timer1.Enabled = False
        Do_Open
        DoStuffViv2
        End
End If
End Sub


Public Sub DoStuffViv2()
On Error Resume Next
Tempstring11$ = Mid$(Comamd$, 1, Len(Comamd$) - 4) + "\"
MkDir (Mid$(Comamd$, 1, Len(Comamd$) - 4))

If Mid$(InFile$, 1, 4) <> "BIGF" Then Call MsgBox("The File Doesn't appear to be a viv, But Continuing anyway", vbInformation, "Error")

'Read length of file
ATremp2$ = Hex(Asc(Mid(InFile$, 5, 1))) + Hex(Asc(Mid(InFile$, 6, 1))) + Hex(Asc(Mid(InFile$, 7, 1))) + Hex(Asc(Mid(InFile$, 8, 1)))
ConvHexToDec (ATremp2$)
lengthoffile& = Height1&

'Read Amount of Files
ATremp2$ = Hex(Asc(Mid(InFile$, 14, 1))) + Hex(Asc(Mid(InFile$, 13, 1))) + Hex(Asc(Mid(InFile$, 12, 1)))
ConvHexToDec (ATremp2$)
amountoffile& = Height1&

'Read End of header
ATremp2$ = Hex(Asc(Mid(InFile$, 18, 1))) + Hex(Asc(Mid(InFile$, 17, 1))) + Hex(Asc(Mid(InFile$, 16, 1))) + Hex(Asc(Mid(InFile$, 15, 1)))
ConvHexToDec (ATremp2$)
endofheader& = Height1&

'Atewmp = MsgBox("Is this a Nfs5/6 Viv File", vbYesNo, "nfs5/6")
 
'MsgBox (lengthoffile&)
'MsgBox (amountoffile&)
'MsgBox (endofheader&)
Label1.Caption = "Extracting" + Str$(amountoffile&) + " Files"

'19 start of file

'When Found File
I = 17
For J = 1 To amountoffile&
    Atemp2$ = ""
    'Get Place In File in Bytes
    Atemp2$ = Space(2 - Len(Hex$(Asc(Mid$(InFile$, I, 1))))) + Hex$(Asc(Mid$(InFile$, I, 1)))
    Atemp2$ = Atemp2$ + Space(2 - Len(Hex$(Asc(Mid$(InFile$, I + 1, 1))))) + Hex$(Asc(Mid$(InFile$, I + 1, 1)))
    Atemp2$ = Atemp2$ + Space(2 - Len(Hex$(Asc(Mid$(InFile$, I + 2, 1))))) + Hex$(Asc(Mid$(InFile$, I + 2, 1)))
    Atemp2$ = Atemp2$ + Space(2 - Len(Hex$(Asc(Mid$(InFile$, I + 3, 1))))) + Hex$(Asc(Mid$(InFile$, I + 3, 1)))
    ConvHexToDec2 (UCase$(Atemp2$))
    placeinfile& = Height1&
    
    Atemp2$ = ""
    'Get Size of File in Bytes
    Atemp2$ = Space(2 - Len(Hex$(Asc(Mid$(InFile$, I + 4, 1))))) + Hex$(Asc(Mid$(InFile$, I + 4, 1)))
    Atemp2$ = Atemp2$ + Space(2 - Len(Hex$(Asc(Mid$(InFile$, I + 5, 1))))) + Hex$(Asc(Mid$(InFile$, I + 5, 1)))
    Atemp2$ = Atemp2$ + Space(2 - Len(Hex$(Asc(Mid$(InFile$, I + 6, 1))))) + Hex$(Asc(Mid$(InFile$, I + 6, 1)))
    Atemp2$ = Atemp2$ + Space(2 - Len(Hex$(Asc(Mid$(InFile$, I + 7, 1))))) + Hex$(Asc(Mid$(InFile$, I + 7, 1)))
    ConvHexToDec2 (UCase$(Atemp2$))
    sizeoffile& = Height1&
    I = I + 8
    K = 0
1
    K = K + 1
    If Asc(Mid$(InFile$, I, 1)) = "0" Then GoTo 2
    I = I + 1
    GoTo 1
2
    'Get Name of File
    Extension$ = Tempstring11$ + Mid$(InFile$, I - (K - 1), K)
    'MsgBox (Extension$)
    'MsgBox (placeinfile&)
    'MsgBox (sizeoffile&)
    I = I + 1
    X = placeinfile& + 1
    Y = sizeoffile&
    RippItNo3
Next J
CommonDialog1.FileName = ""

End Sub

Private Sub Gif_Click()
If Gif.Value = 1 Then
    Mess ("Microsoft's Gif's have 10 less bytes at the end. Although some formats still use this, most don't. Unfortunately, My program cannot determine whether the gif has this or not.")
    ase = MsgBox("What do you want to set the default 'Add-on' bytes to (10=Yes 0=No) (Default=Yes)", vbYesNo, "Extra Bytes")
        
        If ase = vbYes Then
            ExtraGif% = 10
        Else
            ExtraGif% = 0
        End If

End If
End Sub

Public Sub RipJpeg()

On Error Resume Next

Extension$ = "Jpg"

For I = 1 To Len(InFile$)
    If W = 1 Then
            If Mid$(InFile$, I, 2) = Chr$(&HFF) + Chr$(&HD9) Then
                Y = I + 1
                RippIt
                I = X + 7
            End If
    Else
        If Mid$(InFile$, I, 1) = "J" Then
            If Mid$(InFile$, I + 1, 3) = "FIF" Then
                X = I - 6
                W = 1
            End If
        End If
    End If
Next I

End Sub

Public Sub RipJpeg3()
On Error Resume Next

Extension$ = "Jpg"

For I = 1 To Len(InFile$)
        If Mid$(InFile$, I, 1) = Chr(&HFF) Then
            If Mid$(InFile$, I + 1, 2) = Chr(&HD8) + Chr(&HFF) Then
                X = I
                W = 1
                For TYUU = X To Len(InFile$)
                    If Mid$(InFile$, TYUU, 2) = Chr$(&HFF) + Chr$(&HD9) Then
                        Y = TYUU + 1
                        RippIt
                    End If
                Next TYUU
            End If
        End If
Next I
End Sub

Public Sub RipJpeg2()
On Error Resume Next

Extension$ = "Jpg"

For I = 1 To Len(InFile$)
    If W = 1 Then
            If Mid$(InFile$, I, 2) = Chr$(&HFF) + Chr$(&HD9) Then
                RTY = RTY + 1
                If RTY = 2 Then
                    Y = I + 1
                    RippIt
                    I = X + 7 - 4
                End If

            End If
    Else
    'FF:D8:FF:E1
        If Mid$(InFile$, I, 1) = Chr(&HFF) Then
            If Mid$(InFile$, I + 1, 2) = Chr(&HD8) + Chr(&HFF) Then
                X = I
                W = 1
                RTY = 0
            End If
        End If
    End If
Next I

End Sub

Public Sub ripMidi()
On Error Resume Next

Extension$ = "Mid"
 
For I = 1 To Len(InFile$)
    If W = 1 Then
        If Mid$(InFile$, I, 1) = Chr$(0) Then
            If Mid$(InFile$, I + 1, 3) = Chr$(255) + Chr$(&H2F) + Chr$(0) And Mid$(InFile$, I + 4, 1) <> Chr$(77) Then
                Y = I + 4
                RippIt
            End If
        End If
    Else
        If Mid$(InFile$, I, 1) = "M" Then
            If Mid$(InFile$, I + 1, 3) = "Thd" Then
                X = I
                W = 1
            End If
        End If
    End If
Next I

End Sub
Public Sub RipFsh()
On Error Resume Next

Extension$ = "Fsh"

For I = 1 To Len(InFile$)
        If Mid$(InFile$, I, 1) = "S" Then
            If Mid$(InFile$, I + 1, 3) = "HPI" Then
                X = I
                    ConvertHeight
                Y = I + Height1&
                RippIt
            End If
        End If
Next I
End Sub


Public Sub ConvertHeight()
Atemp2$ = Atemp2$ + Space(2 - Len(Hex$(Asc(Mid$(InFile$, I + 7, 1))))) + Hex$(Asc(Mid$(InFile$, I + 7, 1)))
Atemp2$ = Atemp2$ + Space(2 - Len(Hex$(Asc(Mid$(InFile$, I + 6, 1))))) + Hex$(Asc(Mid$(InFile$, I + 6, 1)))
Atemp2$ = Atemp2$ + Space(2 - Len(Hex$(Asc(Mid$(InFile$, I + 5, 1))))) + Hex$(Asc(Mid$(InFile$, I + 5, 1)))
Atemp2$ = Atemp2$ + Space(2 - Len(Hex$(Asc(Mid$(InFile$, I + 4, 1))))) + Hex$(Asc(Mid$(InFile$, I + 4, 1)))

ConvHexToDec2 (UCase$(Atemp2$))
End Sub
Public Sub RipZip()
On Error Resume Next

Extension$ = "Zip"

For I = 1 To Len(InFile$)
    If W = 1 Then
        If Mid$(InFile$, I, 1) = "P" Then
            If (Mid$(InFile$, I + 1, 5) = "K" + Chr$(5) + Chr$(6) + Chr$(0) + Chr$(0)) Then
                Y = I + 18 + 4
                RippIt
            End If
        End If
    Else
        If Mid$(InFile$, I, 1) = "P" Then
            If Mid$(InFile$, I + 1, 1) = "K" Then
                X = I
                W = 1
            End If
        End If
    End If
Next I

End Sub

Public Sub RipHmp()
On Error Resume Next

Extension$ = "Hmp"

For I = 1 To Len(InFile$)
    If W = 1 Then
        If Mid$(InFile$, I, 1) = Chr$(1) Or Mid$(InFile$, I, 1) = Chr$(255) Then
            If (Mid$(InFile$, I + 1, 5) = Chr$(0) + Chr$(&H41) + Chr$(0) + Chr$(&H40) + Chr$(0)) Then
                Y = I + 5
                RippIt
            End If
            If (Mid$(InFile$, I + 1, 8) = Chr$(&H2F) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0)) Then
                Y = I + 17
                RippIt
            End If
        End If
    Else
        If Mid$(InFile$, I, 1) = "H" Then
            If Mid$(InFile$, I + 1, 6) = "MIMIDI" Then
                X = I
                W = 1
            End If
        End If
    End If
Next I
End Sub

Public Sub RipPsp()
On Error Resume Next

Extension$ = "Psp"

For I = 1 To Len(InFile$)
    If W = 1 Then
        If Mid$(InFile$, I, 1) = Chr$(0) Then
            If (Mid$(InFile$, I + 1, 5) = Chr$(0) + Chr$(255) + Chr$(255) + Chr$(3) + Chr$(0)) And (Mid$(InFile$, I + 10, 1) <> Chr$(126)) Then
                Y = I + 10
                RippIt
            End If
        End If
    Else
        If Mid$(InFile$, I, 1) = "P" Then
            If Mid$(InFile$, I + 1, 4) = "aint" Then
                X = I
                W = 1
            End If
        End If
    End If
Next I

End Sub

Public Sub RipXmi()
On Error Resume Next

Extension$ = "Xmi"

For I = 1 To Len(InFile$)
    If W = 1 Then
        If Mid$(InFile$, I, 1) = Chr$(255) Then
            If Mid$(InFile$, I + 1, 2) = Chr$(&H2F) + Chr$(0) Then
                Y = I + 2
                RippIt
            End If
        End If
    Else
        If Mid$(InFile$, I, 1) = "F" Then
            If Mid$(InFile$, I + 1, 3) = "ORM" Then
                X = I
                W = 1
            End If
        End If
    End If
Next I

End Sub

Private Sub Image1_DragDrop(Source As Control, X As Single, Y As Single)
If X > 5220 Then X = 5200
If Y > 3600 Then Y = 3600
If X < 4830 Then X = 4900
If Y < 720 Then Y = 720
Picture1.Left = X
Picture1.Top = Y
End Sub

Private Sub Mus_Click()
On Error GoTo 10
If Mus.Value = 1 Then
Start:
    USer& = InputBox("What chance value would you like to add in? (40-Max 2-Min 6-Def)", "Value", "6")
    If USer& <= 1 Then GoTo Start
    If USer& >= 41 Then GoTo Start
    tmrep = MsgBox("Do you want extra bytes for increased success rates with winamp, or do you want less bytes for increased success rates with the Mus converter? (Yes='Prefer Winamp' No='Prefer MusConverter')", vbYesNo, "Success Rate")
    If tmrep = vbNo Then Extra& = 20
    If tmrep = vbYes Then Extra& = 100000
End If
GoTo 20
10
Mus.Value = 0
20
End Sub

Public Sub RipMus()
On Error Resume Next

Extension$ = "Mus"

For I = 1 To Len(InFile$)
    If W = 1 Then
         If Mid$(InFile$, I + 20, 1) = Chr$(&H60) Then
            If Mid$(InFile$, I, 1) = Chr$(3) Then Chance& = Chance& + 2
            If Mid$(InFile$, I + 1, 1) = Chr$(3) Then Chance& = Chance& + 2
            If Mid$(InFile$, I + 2, 1) = Chr$(0) Then Chance& = Chance& + 2
            If Mid$(InFile$, I + 2, 1) = Chr$(8) Then Chance& = Chance& + 2
            If Mid$(InFile$, I + 3, 1) = Chr$(&HF) Then Chance& = Chance& + 2
            If Mid$(InFile$, I + 4, 1) = Chr$(3) Then Chance& = Chance& + 4
            If Mid$(InFile$, I + 5, 1) = Chr$(0) Then Chance& = Chance& + 2
            If Mid$(InFile$, I + 7, 1) = Chr$(3) Then Chance& = Chance& + 2
            If Mid$(InFile$, I + 8, 1) = Chr$(0) Then Chance& = Chance& + 2
            If Mid$(InFile$, I + 9, 1) = Chr$(1) Then Chance& = Chance& + 2
            If Mid$(InFile$, I + 10, 1) = Chr$(3) Then Chance& = Chance& + 2
            If Mid$(InFile$, I + 10, 1) = Chr$(&H28) Then Chance& = Chance& + 2
            If Mid$(InFile$, I + 11, 1) = Chr$(0) Then Chance& = Chance& + 2
            If Mid$(InFile$, I + 11, 1) = Chr$(&H8F) Then Chance& = Chance& + 2
            If Mid$(InFile$, I + 12, 1) = Chr$(0) Then Chance& = Chance& + 2
            If Mid$(InFile$, I + 12, 1) = Chr$(1) Then Chance& = Chance& + 2
            If Mid$(InFile$, I + 12, 1) = Chr$(3) Then Chance& = Chance& + 2
            If Mid$(InFile$, I + 13, 1) = Chr$(&H26) Then Chance& = Chance& + 2
            If Mid$(InFile$, I + 13, 1) = Chr$(&H39) Then Chance& = Chance& + 2
            If Mid$(InFile$, I + 14, 1) = Chr$(4) Then Chance& = Chance& + 2
            If Mid$(InFile$, I + 14, 1) = Chr$(2) Then Chance& = Chance& + 2
            If Mid$(InFile$, I + 15, 1) = Chr$(&H32) Then Chance& = Chance& + 2
            If Mid$(InFile$, I + 16, 1) = Chr$(0) Then Chance& = Chance& + 3
            If Mid$(InFile$, I + 16, 1) = Chr$(1) Then Chance& = Chance& + 2
            If Mid$(InFile$, I + 16, 1) = Chr$(6) Then Chance& = Chance& + 3
            If Mid$(InFile$, I + 16, 1) = Chr$(3) Then Chance& = Chance& + 3
            If Mid$(InFile$, I + 17, 1) = Chr$(&H35) Then Chance& = Chance& + 2
            If Mid$(InFile$, I + 18, 1) = Chr$(1) Then Chance& = Chance& + 3
            If Mid$(InFile$, I + 18, 1) = Chr$(7) Then Chance& = Chance& + 3
            If Mid$(InFile$, I + 18, 1) = Chr$(5) Then Chance& = Chance& + 3
            
            If Chance& >= USer& Then
                RippIt
                Y = (I + Extra&)
            End If
            
        End If
    Else
        If Mid$(InFile$, I, 1) = "M" Then
            If Mid$(InFile$, I + 1, 2) = "US" Then
                X = I
                W = 1
            End If
        End If
    End If
    Chance& = 0
Next I

End Sub

Public Sub RipTga()
On Error Resume Next

Extension$ = "Tga"

For I = 1 To Len(InFile$)
        If W = 1 Then
            If Mid$(InFile$, I, 1) = "T" Then
                If Mid$(InFile$, I + 1, 17) = "RUEVISION-XFILE." + Chr$(0) Then
                    Y = I + 18
                    RippIt
                End If
        End If
    Else
        If Mid$(InFile$, I, 1) = Chr$(10) Then
            If Mid$(InFile$, I + 1, 9) = Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) Then
                X = I - 2
                W = 1
            End If
        End If
    End If
Next I

End Sub

Private Sub Selectall_Click()
For H = 0 To (Result1.Count - 1)
    Result1(H).Caption = "0"
Next H
If Selectall.Value = 1 Then
    Label17.Caption = "Select none"
    Bitmap.Value = 1
    Gif.Value = 1
    Jpeg.Value = 1
    Midi.Value = 1
    Zip.Value = 1
    Hmp.Value = 1
    Paint.Value = 1
    Xmi.Value = 1
    Mus.Value = 1
    Bmp256.Value = 1
    Bmp16.Value = 1
    Tga.Value = 1
    Wave.Value = 1
    Acrobat.Value = 1
    Html.Value = 1
    Bink.Value = 1
    Avi.Value = 1
    Swf.Value = 1
    Tiff.Value = 1
Else
    Label17.Caption = "Select all"
    Bitmap.Value = 0
    Gif.Value = 0
    Jpeg.Value = 0
    Midi.Value = 0
    Mus.Value = 0
    Zip.Value = 0
    Hmp.Value = 0
    Paint.Value = 0
    Xmi.Value = 0
    Bmp256.Value = 0
    Bmp16.Value = 0
    Tga.Value = 0
    Wave.Value = 0
    Acrobat.Value = 0
    Html.Value = 0
    Bink.Value = 0
    Avi.Value = 0
    Swf.Value = 0
    Tiff.Value = 0
End If
End Sub

Private Sub Timer1_Timer()
Picture1.Left = Picture1.Left + X1%
If Picture1.Left >= 5220 Then X1% = X1% * -1
If Picture1.Left <= 4830 Then X1% = X1% * -1
Picture1.Top = Picture1.Top + X2%
If Picture1.Top >= 3600 Then X2% = X2% * -1
If Picture1.Top <= 720 Then X2% = X2% * -1
X3% = X3% + 1
If X3% = 360 Then X3% = 0
Picture1.Top = Picture1.Top + (Sin(X3%) * 100)
End Sub

Private Sub ConvHexToDec2(NeededString$)
Height1& = 0
NeededString$ = UCase$(NeededString$)

For Y = 0 To (Len(NeededString$) - 1)
    atemp1$ = Mid$(NeededString$, Len(NeededString$) - Y, 1)
    If Asc(atemp1$) = 32 Then atemp1$ = "0"
    If Asc(atemp1$) > 64 Then atemp1$ = Str$(Asc(atemp1$) - 55)
    ATempval& = 16 ^ Y
    Height1& = Height1& + (Val(atemp1$) * ATempval&)
Next Y

End Sub


Private Sub ConvHexToDec(NeededString$)
Height1& = 0

For Y = 0 To (Len(NeededString$) - 1)
    atemp1$ = Mid$(NeededString$, Len(NeededString$) - Y, 1)
    If Asc(atemp1$) > 64 Then atemp1$ = Str$(Asc(atemp1$) - 55)
    ATempval& = 16 ^ Y
    Height1& = Height1& + (Val(atemp1$) * ATempval&)
Next Y

End Sub

Public Sub RipWave()
On Error Resume Next

Extension$ = "Wav"

For I = 1 To Len(InFile$)
        If Mid$(InFile$, I, 1) = "R" Then
            If Mid$(InFile$, I + 1, 3) = "IFF" Then
                X = I
                    ConvertHeight
                Y = I + Height1& + 8
                RippIt
            End If
        End If
Next I
End Sub

Public Sub RipPdf()
On Error Resume Next

Extension$ = "Pdf"

For I = 1 To Len(InFile$)
    If W = 1 Then
            If Mid$(InFile$, I, 1) = "%" Then
                If Mid$(InFile$, I + 1, 4) = "%EOF" Then
                    If Firsttime% = Times% Then
                        Firsttime% = 1
                    Else
                        Y = I + 6
                        RippIt
                    End If
                End If
            End If
        
    Else
        If Mid$(InFile$, I, 1) = "%" Then
            If Mid$(InFile$, I + 1, 3) = "PDF" Then
                Firsttime% = 0
                X = I
                W = 1
            End If
        End If
    End If
Next I
End Sub

Public Sub RipHtml()
On Error Resume Next

Extension$ = "Htm"

For I = 1 To Len(InFile$)
    If W = 1 Then
            If Mid$(InFile$, I, 1) = "<" Then
                If UCase$(Mid$(InFile$, I + 1, 6)) = "/HTML>" Then
                    Y = I + 7
                    RippIt
                End If
            End If
    Else
        If Mid$(InFile$, I, 1) = "<" Then
            If Mid$(InFile$, I + 1, 5) = "HTML>" Then
                X = I
                W = 1
            End If
        End If
    End If
Next I

End Sub

Public Sub RipBink()
On Error Resume Next

Extension$ = "Bik"

For I = 1 To Len(InFile$)
        If Mid$(InFile$, I, 1) = "B" Then
            If Mid$(InFile$, I + 1, 3) = "IKf" Then
                X = I
                    ConvertHeight
                Y = I + Height1& + 8
                RippIt
            End If
        End If
Next I

End Sub

Public Sub RipAvi()
On Error Resume Next

Extension$ = "avi"

For I = 1 To Len(InFile$)
        If Mid$(InFile$, I, 1) = "R" Then
            If Mid$(InFile$, I + 1, 3) = "IFF" And Mid$(InFile$, I + 8, 3) = "AVI" Then
                X = I
                    ConvertHeight
                Y = I + Height1& + 8
                RippIt
            End If
        End If
Next I

End Sub

Private Sub Timer2_Timer()
Me.Height = Me.Height + R%
If Me.Height >= 6960 Then Timer2.Enabled = False
If Me.Height <= 5160 Then Timer2.Enabled = False
End Sub

Public Sub RipSWF()
On Error Resume Next

Extension$ = "Swf"

For I = 1 To Len(InFile$)
        If Mid$(InFile$, I, 1) = "F" Then
            If Mid$(InFile$, I + 1, 2) = "WS" Then
                X = I
                    ConvertHeight
                Y = I + Height1&
                RippIt
            End If
        End If
Next I

End Sub

Public Sub RipTif()
On Error Resume Next

Extension$ = "Tif"

For I = 1 To Len(InFile$)
        If W = 1 Then
            If Mid$(InFile$, I, 1) = Chr$(1) Then
                If Mid$(InFile$, I + 1, 6) = Chr$(3) + Chr$(0) + Chr$(1) + Chr$(0) + Chr$(0) + Chr$(0) And Mid$(InFile$, I + 8, 3) = Chr$(0) + Chr$(0) + Chr$(0) And Mid$(InFile$, I + 12, 3) = Chr$(1) + Chr$(3) + Chr$(0) And Mid$(InFile$, I + 17, 2) = Chr$(0) + Chr$(0) And Mid$(InFile$, I + 19, 1) <> Chr$(0) And Mid$(InFile$, I + 20, 7) = Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) Then
                    Y = I + 27
                    RippIt
                End If
        End If
    Else
        If Mid$(InFile$, I, 1) = Chr$(&H49) Then
            If Mid$(InFile$, I + 1, 3) = Chr$(&H49) + Chr$(&H2A) + Chr$(0) Then
                X = I
                W = 1
            End If
        End If
    End If
Next I

End Sub


'Encryption
Private Sub Encrypt(WorkingSTring$, Password$)
If Len(WorkingSTring$) = 0 Then Exit Sub
For IOY = 1 To Len(WorkingSTring$)

Y2% = Y2% + 1

If Y2% > Len(Password$) Then Y2% = 1
StrinG1$ = Mid$(WorkingSTring$, IOY, 1)
StrinG2$ = Mid$(Password$, Y2%, 1)

AscIIValuE% = Asc(StrinG1$) + (Asc(StrinG2$) - 1)
If AscIIValuE% > 255 Then AscIIValuE% = AscIIValuE% - 256

Tempyu$ = Tempyu$ + Chr$(AscIIValuE%)

Next IOY

InFile2$ = InFile2$ + Tempyu$
End Sub

Private Sub Decrypt(WorkingSTring$, Password$)
If Len(WorkingSTring$) = 0 Then Exit Sub
For IOY = 1 To Len(WorkingSTring$)

Y2% = Y2% + 1

If Y2% > Len(Password$) Then Y2% = 1
StrinG1$ = Mid$(WorkingSTring$, IOY, 1)
StrinG2$ = Mid$(Password$, Y2%, 1)

AscIIValuE% = Asc(StrinG1$) - (Asc(StrinG2$) - 1)
If AscIIValuE% < 0 Then AscIIValuE% = AscIIValuE% + 256

Tempyu$ = Tempyu$ + Chr$(AscIIValuE%)

Next IOY

InFile2$ = InFile2$ + Tempyu$

End Sub

Private Sub RecordTime()
TheOldTime$ = Time$
End Sub

Private Sub Command6_Click()
On Error GoTo 10
Password1$ = InputBox("Type your password (Case sensitive).", "Password", "password")
If CommonDialog1.FileName = "" Then CommonDialog1.ShowOpen
Do_Open

RecordTime
Label1.Caption = "Encryption Started..."
Me.Refresh
ENCR& = 1
InFile2$ = ""
Y2% = 0
ENCR2& = 1
InFile$ = Chr$(0) + Chr$(12) + Chr$(19) + InFile$

While ENCR2& < Len(InFile$)
    EncrypTTemp$ = Mid$(InFile$, ENCR2&, 300000)
    ENCR& = 1
    For HGT = 0 To 10
        Call Encrypt(Mid$(EncrypTTemp$, ENCR&, 30000), Password1$)
        ENCR& = ENCR& + 30000
    Next HGT
    ENCR2& = ENCR2& + 300000
    If ENCR2& = 300001 Then
        DoTime (ENCR2&)
    End If
Wend

Label1.Caption = "Saving File..."
TimeLabel.Caption = ""
Me.Refresh

Open CommonDialog1.FileName For Output Access Write Lock Write As #6
    Print #6, InFile2$;
Close #6

Mess ("Encryption Complete!!")
GoTo 20
10
Mess ("There was an error!!")
20
TimeLabel.Caption = ""
Label1.Caption = "Search For:-"
Me.Refresh
CommonDialog1.FileName = ""
End Sub

Private Sub DoTime(Position As Long)
TempTimeSec& = Val(Mid$(Time$, 7, 2)) - Val(Mid$(TheOldTime$, 7, 2))
TempTimeMin& = Val(Mid$(Time$, 4, 2)) - Val(Mid$(TheOldTime$, 4, 2))
TempTimeHour& = Val(Mid$(Time$, 1, 2)) - Val(Mid$(TheOldTime$, 1, 2))
If TempTimeSec& < 0 Then
    TempTimeSec& = TempTimeSec& + 60
    TempTimeMin& = TempTimeMin& - 1
End If
If TempTimeMin& < 0 Then
    TempTimeMin& = TempTimeMin& + 60
    TempTimeHour& = TempTimeHour& - 1
End If

TempTimeSec& = TempTimeSec& * (TheFileLength& / Position)
TempTimeMin& = TempTimeMin& * (TheFileLength& / Position)
TempTimeHour& = TempTimeHour& * (TheFileLength& / Position)

While TempTimeSec& >= 60
    If TempTimeSec& >= 60 Then
        TempTimeSec& = TempTimeSec& - 60
        TempTimeMin& = TempTimeMin& + 1
    End If
Wend

While TempTimeMin& >= 60
    If TempTimeMin& >= 60 Then
        TempTimeMin& = TempTimeMin& - 60
        TempTimeHour& = TempTimeHour& + 1
    End If
Wend

TimeLabel.Caption = "Est Time" + Str$(TempTimeHour&) + ":" + Mid$(Str$(TempTimeMin&), 2, 200) + ":" + Mid$(Str$(TempTimeSec&), 2, 200)
TimeLabel.Refresh

End Sub

Private Sub Command7_Click()
On Error GoTo 10
Password1$ = InputBox("Type your password (Case sensitive).", "Password", "password")

InFile2$ = ""
Y2% = 0
Call Encrypt(Chr$(0) + Chr$(12) + Chr$(19), Password1$)
TempInFile$ = InFile2$

If CommonDialog1.FileName = "" Then CommonDialog1.ShowOpen
Do_Open

If TempInFile$ <> Mid$(InFile$, 1, 3) Then
      MsgBox ("The password is incorrect!!")
      CommonDialog1.FileName = ""
      Exit Sub
End If

RecordTime
Label1.Caption = "Decryption Started..."
Me.Refresh
ENCR& = 1
InFile2$ = ""
Y2% = 0

ENCR2& = 1

While ENCR2& < Len(InFile$)
    EncrypTTemp$ = Mid$(InFile$, ENCR2&, 300000)
    ENCR& = 1
    For HGT = 0 To 10
        Call Decrypt(Mid$(EncrypTTemp$, ENCR&, 30000), Password1$)
        ENCR& = ENCR& + 30000
    Next HGT
    ENCR2& = ENCR2& + 300000
    If ENCR2& = 300001 Then
        DoTime (ENCR2&)
    End If
Wend

Label1.Caption = "Saving File..."
TimeLabel.Caption = ""
Me.Refresh

InFile2$ = Mid$(InFile2$, 4, Len(InFile2$))

Open CommonDialog1.FileName For Output Access Write Lock Write As #6
    Print #6, InFile2$;
Close #6

Mess ("Decryption Complete!!")
GoTo 20
10
Mess ("There was an error!!")
20

Label1.Caption = "Search For:-"
TimeLabel.Caption = ""
CommonDialog1.FileName = ""
Me.Refresh
End Sub

'--------------------------------------------------------
'Display tool tips on each button and what the button
'does
'--------------------------------------------------------
Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label23.Caption = "This button will ripp the files that you select from a file. The suppress messages check box stops annoying messages during ripping. 'Select using all options' trys every option with the files that you select. the 'delete corrupted files' checkbox deletes bitmap that are invalid."
End Sub
Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label23.Caption = "You can replace a file for another one in a exe or other file. e.g. You can change Paint Shop Pro's splash screen to an image of your choice. The file(s) has to be of the same type and prefferably of the same size."
End Sub
Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label23.Caption = "Newly added Encryption can protect your files from other people. Also, Encryption and decryption speed has been improved"
End Sub
Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label23.Caption = "Newly added Encryption can protect your files from other people. Also, Encryption and decryption speed has been improved"
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label23.Caption = "Try dragging the ripper image."
End Sub
Private Sub Result1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label23.Caption = "Select the filetypes that you wish to search for."
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label23.Caption = "Select the filetypes that you wish to search for, and check the options that you want and then select 'Search file' or 'Replace file'"
End Sub
Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label23.Caption = "The newly added text finder can find passwords in compiled VB programs. You can also select remove junk to narrow your searching efforts."
End Sub
Private Sub Command9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label23.Caption = "New Replace text in exe, replaces captions and labels in exe files for some mischeivous fun"
End Sub
Private Sub Command8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label23.Caption = "Exit Program."
End Sub
Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label23.Caption = "You can search for passwords in compiled VB programs. The text you are looking for may have a extra letter on the front or mixed in with different text e.g. 'vPassword*.*' may be 'Password'. To help you search for this, there is a Filter that removes unprintable Characters."
End Sub
Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label23.Caption = "Exit Program."
End Sub
Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label23.Caption = "Click for help and About. Created by Alex Murray"
End Sub
'--------------------------------------------------------

