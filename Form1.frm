VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "   "
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16320
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   16320
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   17400
      MaskColor       =   &H00000000&
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   8640
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   12000
      TabIndex        =   28
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   12000
      TabIndex        =   27
      Top             =   5745
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   12000
      TabIndex        =   26
      Top             =   4920
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   13200
      TabIndex        =   22
      Top             =   1320
      Width           =   7095
      Begin VB.Label Label25 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "STAND BY"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label26 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "OFF MODE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2280
         TabIndex        =   24
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label27 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "OPERATIONAL "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   4560
         TabIndex        =   23
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Timer Timer14 
      Interval        =   1000
      Left            =   2160
      Top             =   1080
   End
   Begin VB.Timer Timer12 
      Interval        =   1000
      Left            =   7560
      Top             =   1080
   End
   Begin VB.Timer Timer11 
      Interval        =   100
      Left            =   1080
      Top             =   1080
   End
   Begin VB.Timer Timer9 
      Interval        =   1000
      Left            =   9480
      Top             =   1080
   End
   Begin VB.Timer Timer8 
      Interval        =   1000
      Left            =   9000
      Top             =   1080
   End
   Begin VB.Timer Timer7 
      Interval        =   1000
      Left            =   8520
      Top             =   1080
   End
   Begin VB.Timer Timer6 
      Interval        =   1000
      Left            =   8040
      Top             =   1080
   End
   Begin VB.Timer Timer5 
      Interval        =   1000
      Left            =   3600
      Top             =   1080
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   3120
      Top             =   1080
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   2640
      Top             =   1080
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   600
      Top             =   1080
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1560
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   1080
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Text            =   " "
      Top             =   8040
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   8280
      TabIndex        =   4
      Text            =   " "
      Top             =   9120
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Text            =   " "
      Top             =   9120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1560
      TabIndex        =   2
      Text            =   " "
      Top             =   9120
      Width           =   975
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      X1              =   4440
      X2              =   5640
      Y1              =   3720
      Y2              =   3360
   End
   Begin VB.Line Line39 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      X1              =   5640
      X2              =   5520
      Y1              =   3360
      Y2              =   3600
   End
   Begin VB.Line Line40 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      Index           =   0
      X1              =   5640
      X2              =   5400
      Y1              =   3360
      Y2              =   3240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      X1              =   8760
      X2              =   8520
      Y1              =   2400
      Y2              =   2280
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      X1              =   8760
      X2              =   8640
      Y1              =   2400
      Y2              =   2640
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      X1              =   7560
      X2              =   8760
      Y1              =   2760
      Y2              =   2400
   End
   Begin VB.Line Line22 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      X1              =   11520
      X2              =   12720
      Y1              =   2520
      Y2              =   2880
   End
   Begin VB.Line Line45 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      X1              =   12720
      X2              =   12600
      Y1              =   2880
      Y2              =   2640
   End
   Begin VB.Line Line46 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      X1              =   12720
      X2              =   12480
      Y1              =   2880
      Y2              =   3000
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      X1              =   16200
      X2              =   17400
      Y1              =   3960
      Y2              =   4320
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      X1              =   14640
      X2              =   15840
      Y1              =   3480
      Y2              =   3840
   End
   Begin VB.Line Line47 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      X1              =   15840
      X2              =   15720
      Y1              =   3840
      Y2              =   3600
   End
   Begin VB.Line Line48 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      X1              =   15840
      X2              =   15600
      Y1              =   3840
      Y2              =   3960
   End
   Begin VB.Line Line49 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      X1              =   17400
      X2              =   17280
      Y1              =   4320
      Y2              =   4080
   End
   Begin VB.Line Line50 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      X1              =   17400
      X2              =   17160
      Y1              =   4320
      Y2              =   4440
   End
   Begin VB.Image Image3 
      Height          =   5625
      Left            =   16560
      Picture         =   "Form1.frx":0E20
      Top             =   2400
      Width           =   3750
   End
   Begin VB.Line Line37 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      X1              =   4200
      X2              =   4080
      Y1              =   3840
      Y2              =   4080
   End
   Begin VB.Line Line38 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      X1              =   4200
      X2              =   3960
      Y1              =   3840
      Y2              =   3720
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      X1              =   3000
      X2              =   4200
      Y1              =   4200
      Y2              =   3840
   End
   Begin VB.Line Line25 
      BorderColor     =   &H00FFFF80&
      BorderWidth     =   5
      X1              =   9720
      X2              =   9720
      Y1              =   2520
      Y2              =   2880
   End
   Begin VB.Line Line26 
      BorderColor     =   &H0000C000&
      BorderWidth     =   5
      X1              =   9600
      X2              =   9600
      Y1              =   2400
      Y2              =   3000
   End
   Begin VB.Line Line27 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   5
      X1              =   9480
      X2              =   9480
      Y1              =   2280
      Y2              =   3120
   End
   Begin VB.Line Line28 
      BorderColor     =   &H008080FF&
      BorderWidth     =   5
      X1              =   9360
      X2              =   9360
      Y1              =   2160
      Y2              =   3240
   End
   Begin VB.Line Line29 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      X1              =   9240
      X2              =   9240
      Y1              =   2040
      Y2              =   3360
   End
   Begin VB.Line Line30 
      BorderColor     =   &H000000C0&
      BorderWidth     =   5
      X1              =   9120
      X2              =   9120
      Y1              =   1920
      Y2              =   3600
   End
   Begin VB.Line Line31 
      BorderColor     =   &H00FFFF80&
      BorderWidth     =   5
      X1              =   10800
      X2              =   10800
      Y1              =   2520
      Y2              =   2880
   End
   Begin VB.Line Line32 
      BorderColor     =   &H0000C000&
      BorderWidth     =   5
      X1              =   10920
      X2              =   10920
      Y1              =   2400
      Y2              =   3000
   End
   Begin VB.Line Line33 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   5
      X1              =   11040
      X2              =   11040
      Y1              =   2280
      Y2              =   3120
   End
   Begin VB.Line Line34 
      BorderColor     =   &H008080FF&
      BorderWidth     =   5
      X1              =   11160
      X2              =   11160
      Y1              =   2160
      Y2              =   3240
   End
   Begin VB.Line Line35 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      X1              =   11280
      X2              =   11280
      Y1              =   2040
      Y2              =   3360
   End
   Begin VB.Line Line36 
      BorderColor     =   &H000000C0&
      BorderWidth     =   5
      X1              =   11400
      X2              =   11400
      Y1              =   1920
      Y2              =   3480
   End
   Begin VB.Image Image2 
      Height          =   5625
      Left            =   480
      Picture         =   "Form1.frx":6656
      Top             =   2160
      Width           =   3330
   End
   Begin VB.Image Image1 
      Height          =   5625
      Left            =   8520
      Picture         =   "Form1.frx":9CFE
      Top             =   1440
      Width           =   3405
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      X1              =   6000
      X2              =   7200
      Y1              =   3240
      Y2              =   2880
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GREEN COMMUNICATION INTERFACE FOR CELLULAR TOWERS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   2280
      TabIndex        =   34
      Top             =   120
      Width           =   15975
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   480
      TabIndex        =   33
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label19 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Nos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   13200
      TabIndex        =   31
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label Label13 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "mA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   12960
      TabIndex        =   30
      Top             =   6000
      Width           =   735
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   12960
      TabIndex        =   29
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Label24 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "lux"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   9480
      TabIndex        =   21
      Top             =   9240
      Width           =   615
   End
   Begin VB.Label Label23 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   6240
      TabIndex        =   20
      Top             =   9240
      Width           =   255
   End
   Begin VB.Label Label22 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2760
      TabIndex        =   19
      Top             =   9240
      Width           =   255
   End
   Begin VB.Label Label21 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2640
      TabIndex        =   18
      Top             =   9000
      Width           =   255
   End
   Begin VB.Label Label18 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Hz"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3840
      TabIndex        =   17
      Top             =   8160
      Width           =   495
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "TIME"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   18240
      TabIndex        =   16
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "TEMP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   9240
      Width           =   1215
   End
   Begin VB.Line Line44 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      X1              =   14280
      X2              =   14160
      Y1              =   3360
      Y2              =   3120
   End
   Begin VB.Line Line43 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      X1              =   14280
      X2              =   14040
      Y1              =   3360
      Y2              =   3480
   End
   Begin VB.Line Line42 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      X1              =   7200
      X2              =   6960
      Y1              =   2880
      Y2              =   2760
   End
   Begin VB.Line Line41 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      X1              =   7200
      X2              =   7080
      Y1              =   2880
      Y2              =   3120
   End
   Begin VB.Line Line19 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      X1              =   13080
      X2              =   14280
      Y1              =   3000
      Y2              =   3360
   End
   Begin VB.Label Label14 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "RESPONDER FREQUENCY "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   720
      TabIndex        =   14
      Top             =   7920
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "LIMIT 700 USERS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8880
      TabIndex        =   13
      Top             =   7560
      Width           =   2775
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "AMP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   15960
      TabIndex        =   12
      Top             =   9720
      Width           =   735
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "LIGHT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   14040
      TabIndex        =   11
      Top             =   9720
      Width           =   975
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "R.T"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   12600
      TabIndex        =   10
      Top             =   9720
      Width           =   615
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "A/C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   10920
      TabIndex        =   9
      Top             =   9720
      Width           =   615
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   15960
      Top             =   9120
      Width           =   735
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   14160
      Top             =   9120
      Width           =   615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   12600
      Top             =   9120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   10920
      Top             =   9120
      Width           =   615
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "VOLTAGE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6960
      TabIndex        =   8
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "CURRENT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6960
      TabIndex        =   7
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "USER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7200
      TabIndex        =   6
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "LIGHT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   6840
      TabIndex        =   1
      Top             =   9240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "HUMIDITY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   9240
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim buf As String, Out As Integer, OUT1 As Integer, a As Integer
Dim Val1 As Integer, Val2 As Integer, Val3 As Integer, Val4 As Integer, VAL0, VAL5, VAL6 As Integer

Private Sub Command1_Click()
Form1.Hide
Form2.Show
End Sub

Private Sub Form_Load()
'Form2.WebBrowser1.Navigate "about:<html><body scroll='no' bgcolor=green><FONT SIZE='15' FACE='cooper black' COLOR=white><center><MARQUEE WIDTH=70% BEHAVIOR=ALTERNATE BGColor=green>GREEN RADIO TECHNOLOGY</MARQUEE></center><img src ='D:\vb2017\B-59-GREEN-RADIO\new1.jpg' width=1340 height=100><br><center><MARQUEE WIDTH=55% BEHAVIOR=scroll>Wind and Solar Hybrid Generation</marquee></img></br></FONT></body></html>"
FindComPort
MSComm1.Output = "{27}"
    Sleep 100
MSComm1.Output = "{1D00}"
Sleep (100)

MSComm1.Output = "{1AFF}"
Sleep (100)

'Form3.Data1.DatabaseName = App.Path & "\hybrid.mdb"
'    Form3.Data1.RecordSource = "hybrid"
Line5.Visible = False
Line37.Visible = False
Line38.Visible = False
Line10.Visible = False
Line39.Visible = False
Line40(0).Visible = False
Line1.Visible = False
Line41.Visible = False
Line42.Visible = False
Line4.Visible = False
Line3.Visible = False
Line2.Visible = False

Line22.Visible = False
Line45.Visible = False
Line46.Visible = False
Line19.Visible = False
Line43.Visible = False
Line44.Visible = False
Line16.Visible = False
Line47.Visible = False
Line48.Visible = False
Line13.Visible = False
Line50.Visible = False
Line49.Visible = False

Line25.Visible = False
Line31.Visible = False
Line26.Visible = False
Line32.Visible = False
Line27.Visible = False
Line33.Visible = False
Line28.Visible = False
Line34.Visible = False
Line29.Visible = False
Line35.Visible = False
Line30.Visible = False
Line36.Visible = False
'    Timer14.Enabled = False
'    Timer3.Enabled = False
'    Timer4.Enabled = False
'    Timer5.Enabled = False
'    Timer12.Enabled = False
    Label25.Visible = False
    Label26.Visible = False
    Label27.Visible = False


Label2.Visible = True
Text2.Visible = True
Label23.Visible = False

End Sub
Private Sub Form_Unload(Cancel As Integer)
MSComm1.Output = "{5D00}"
Sleep (100)
End
End Sub




Private Sub Timer1_Timer()
    Label15.Caption = Date
    Label16.Caption = Time
    Text5.Text = Analog(0)
    Text5.Text = Round(Val(Text5.Text) / 3.9)

    Text6.Text = Round(Analog(1) * 8.2)
    Text6.Text = Text6.Text - 250
    If Text6.Text < 0 Then
    Text6.Text = 0
    End If
    Text1.Text = Analog(3) / 7.8
    Text1.Text = Round(Val(Text1.Text))
    Text2.Text = Analog(4) / 10
    Text2.Text = Round(Val(Text2.Text))
''
    Text2.Text = Round((Val(Text2.Text) / Val(Text1.Text)) * 100)
    Text3.Text = Analog(5)
  
  
    Text4.Text = Analog(2)
    
    If Val(Text4.Text) < 1 Then
    Text4.Text = 0
    End If

    Text7.Text = Round(Val(Text5.Text) * (Val(Text6.Text) / 50) / 2.5)
    If Text5.Text < 1 Then
    Text5.Text = 0
    End If

    Form2.Text3.Text = Round(Analog(7) / 36)
    Text3.Text = Analog(5)
    Form2.Text7.Text = Round(Analog(6) / 20)
    Form2.Text7.Text = Form2.Text7.Text
    If Form2.Text7.Text < 0 Then
    Form2.Text7.Text = 0
    End If
    Form1.Text4.Text = Analog(2)
'    Form1.Text1.Text = Round(Val(Form2.Text7.Text) * Val(Form1.Text4.Text) / 100)
'   If Val(Form2.Text7.Text) <= 10 Then
'   Form1.Text3.Text = 120
'   ElseIf Val(Form2.Text7.Text) > 10 And Val(Form2.Text7.Text) <= 20 Then
'   Form1.Text3.Text = 150
'   ElseIf Val(Form2.Text7.Text) > 20 And Val(Form2.Text7.Text) <= 30 Then
'   Form2.Text3.Text = 250
'   ElseIf Val(Form2.Text7.Text) > 30 And Val(Form2.Text7.Text) <= 40 Then
'   Form2.Text3.Text = 500
'   ElseIf Val(Form2.Text7.Text) > 40 And Val(Form2.Text7.Text) <= 50 Then
'   Form2.Text3.Text = 700
'   End If
End Sub
Function Analog(no As Integer)
    MSComm1.Output = "{4" & CStr(no) & "}"
    Sleep 100
    buf = MSComm1.Input
    If (buf <> "") Then
    Analog = CInt(Mid$(buf, 2, 4))
    Else
        Analog = 0
    End If
End Function
Public Sub FindComPort()

    On Error GoTo Port2

    If (MSComm1.PortOpen = True) Then MSComm1.PortOpen = False

    MSComm1.CommPort = 1
    MSComm1.PortOpen = True
    MSComm1.Output = "{3D}"
    Sleep (100)
    buf = MSComm1.Input
    If (Len(buf) > 2) Then Exit Sub

Port2:

    On Error GoTo Port3

    If (MSComm1.PortOpen = True) Then MSComm1.PortOpen = False
    MSComm1.CommPort = 2
    MSComm1.PortOpen = True
    MSComm1.Output = "{3D}"
    Sleep (100)
    buf = MSComm1.Input
    If (Len(buf) > 2) Then Exit Sub

Port3:

    On Error GoTo Port4

    If (MSComm1.PortOpen = True) Then MSComm1.PortOpen = False
    MSComm1.CommPort = 3
    MSComm1.PortOpen = True
    MSComm1.Output = "{3D}"
    Sleep (100)
    buf = MSComm1.Input
    If (Len(buf) > 2) Then Exit Sub

Port4:

    On Error GoTo InvalidPort

    If (MSComm1.PortOpen = True) Then MSComm1.PortOpen = False
    MSComm1.CommPort = 4
    MSComm1.PortOpen = True
    MSComm1.Output = "{3D}"
    Sleep (100)
    buf = MSComm1.Input
    If (Len(buf) > 2) Then Exit Sub

InvalidPort:

    MsgBox "SUbstation couldn't find in Com1 - Com4", vbCritical, "Communication Error."
    End

End Sub

Private Sub Timer11_Timer()
If Val(Text4.Text) > 40 Then
Line5.Visible = True
Line37.Visible = True
Line38.Visible = True
Line10.Visible = True
Line39.Visible = True
Line40(0).Visible = True
Line1.Visible = True
Line41.Visible = True
Line42.Visible = True
Line4.Visible = True
Line3.Visible = True
Line2.Visible = True
ElseIf Val(Text4.Text) <= 40 Then
Line5.Visible = False
Line37.Visible = False
Line38.Visible = False
Line10.Visible = False
Line39.Visible = False
Line40(0).Visible = False
Line1.Visible = False
Line41.Visible = False
Line42.Visible = False
Line4.Visible = False
Line3.Visible = False
Line2.Visible = False


End If
End Sub
Private Sub Timer12_Timer()
If Val(Text7.Text) > 50 And Val(Text7.Text) < 100 Then
Line25.Visible = True
Line31.Visible = True
Else
Line25.Visible = False
Line31.Visible = False
End If

If Val(Text7.Text) > 100 And Val(Text7.Text) < 200 Then
Line25.Visible = True
Line31.Visible = True
Line26.Visible = True
Line32.Visible = True
Else

Line26.Visible = False
Line32.Visible = False
End If
If Val(Text7.Text) > 200 And Val(Text7.Text) < 300 Then
Line25.Visible = True
Line31.Visible = True
Line26.Visible = True
Line32.Visible = True
Line27.Visible = True
Line33.Visible = True
Else

Line27.Visible = False
Line33.Visible = False
End If

If Val(Text7.Text) > 300 And Val(Text7.Text) < 400 Then
Line25.Visible = True
Line31.Visible = True
Line26.Visible = True
Line32.Visible = True
Line27.Visible = True
Line33.Visible = True
Line28.Visible = True
Line34.Visible = True
Else

Line28.Visible = False
Line34.Visible = False
End If

If Val(Text7.Text) > 400 Then
Line25.Visible = True
Line31.Visible = True
Line26.Visible = True
Line32.Visible = True
Line27.Visible = True
Line33.Visible = True
Line28.Visible = True
Line34.Visible = True
Line29.Visible = True
Line35.Visible = True
Else
Line29.Visible = False
Line35.Visible = False

End If

If Val(Text7.Text) > 700 Then
Line25.Visible = True
Line31.Visible = True
Line26.Visible = True
Line32.Visible = True
Line27.Visible = True
Line33.Visible = True
Line28.Visible = True
Line34.Visible = True
Line29.Visible = True
Line35.Visible = True
Line30.Visible = True
Line36.Visible = True

Line22.Visible = True
Line45.Visible = True
Line46.Visible = True
Line19.Visible = True
Line43.Visible = True
Line44.Visible = True
Line16.Visible = True
Line47.Visible = True
Line48.Visible = True
Line13.Visible = True
Line50.Visible = True
Line49.Visible = True
Shape2.BackColor = vbGreen
Else
Shape2.BackColor = vbRed
Line30.Visible = False
Line36.Visible = False

Line22.Visible = False
Line45.Visible = False
Line46.Visible = False
Line19.Visible = False
Line43.Visible = False
Line44.Visible = False
Line16.Visible = False
Line47.Visible = False
Line48.Visible = False
Line13.Visible = False
Line50.Visible = False
Line49.Visible = False
End If
End Sub

'
'Private Sub Timer14_Timer()
'Line6.Visible = True
'Line2.Visible = True
'Line7.Visible = True
'Line37.Visible = True
'Line38.Visible = True
'End Sub

Private Sub Timer2_Timer()
If Val(Text1.Text) > 40 Then
 Out = Out Or &H1
 Shape1.BackColor = vbGreen
Else
  Out = Out And &HFE
  Shape1.BackColor = vbRed
End If


If Val(Text7.Text) > 800 Then
  Out = Out Or &H2
' MSComm1.Output = "{5B10}"
  Shape2.BackColor = vbGreen

Else
Out = Out And &HFD

  Shape2.BackColor = vbRed

End If

 If Val(Form1.Text3.Text) < 250 Then
 Out = Out Or &H4
   Shape3.BackColor = vbGreen
Else
    Out = Out And &HFB
      Shape3.BackColor = vbRed
End If

 If Val(Text4.Text) > 40 Then

Out = Out Or &H8
     Shape4.BackColor = vbGreen
    Else
   Out = Out And &HF7

    Shape4.BackColor = vbRed
 End If

 If Val(Text7.Text) > 50 Then
Label25.Visible = True
Label25.BackColor = &HFF00FF
Label26.BackColor = &HE0E0E0
Label26.Visible = False

ElseIf Val(Text7.Text) < 50 Then
Label26.BackColor = vbGreen
Label25.BackColor = &HE0E0E0
Label26.Visible = True
Label25.Visible = False
End If

If Val(Text4.Text) > 40 Then
Label27.Visible = True
Label25.Visible = False
'Line22.Visible = True
'Line19.Visible = True
'Line16.Visible = True
'Line13.Visible = True
'Line50.Visible = True
'Line49.Visible = True
'Line48.Visible = True
'Line47.Visible = True
'Line43.Visible = True
'Line44.Visible = True
'Line45.Visible = True
'Line46.Visible = True


Label27.BackColor = &HC0&
Label25.Visible = False
Label26.Visible = False
Else
'Line22.Visible = False
'Line19.Visible = False
'Line16.Visible = False
'Line13.Visible = False
'Line50.Visible = False
'Line49.Visible = False
'Line48.Visible = False
'Line47.Visible = False
'Line43.Visible = False
'Line44.Visible = False
'Line45.Visible = False
'Line46.Visible = False

Label27.Visible = False
End If
     If Len(CStr(Hex(Out))) <> 2 Then
    MSComm1.Output = "{5D0" & CStr(Hex(Out)) & "}"
    Sleep (100)
Else
    MSComm1.Output = "{5D" & CStr(Hex(Out)) & "}"
    Sleep (100)
End If

End Sub
'
'Private Sub Timer3_Timer()
'Line11.Visible = True
'Line12.Visible = True
'Line10.Visible = True
'Line39.Visible = True
'Line40.Visible = True
'End Sub
'
'Private Sub Timer4_Timer()
'Line3.Visible = True
'Line4.Visible = True
'Line1.Visible = True
'Line41.Visible = True
'Line42.Visible = True
'End Sub
'
'Private Sub Timer5_Timer()
'Line8.Visible = True
'Line9.Visible = True
'Line5.Visible = True
'Shape4.BackColor = vbGreen
'End Sub



