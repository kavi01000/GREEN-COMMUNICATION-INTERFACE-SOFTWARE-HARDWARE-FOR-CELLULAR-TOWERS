VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00008000&
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   FillColor       =   &H00008000&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Height          =   2655
      Left            =   5040
      Picture         =   "�.frx":0000
      ScaleHeight     =   2595
      ScaleWidth      =   2715
      TabIndex        =   11
      Top             =   720
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   840
      Picture         =   "�.frx":6D59
      ScaleHeight     =   2595
      ScaleWidth      =   2715
      TabIndex        =   10
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   9
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   14880
      Top             =   3120
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   14520
      Top             =   3120
   End
   Begin VB.CommandButton Command2 
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
      Left            =   18240
      Picture         =   "�.frx":C14B
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
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
      Left            =   15840
      Picture         =   "�.frx":10F16
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9480
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   360
      Top             =   7080
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   615
      Left            =   13200
      TabIndex        =   1
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   8
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " WIND VOLTAGE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   7
      Top             =   5520
      Width           =   2655
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "V O  L T A G E"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   5880
      TabIndex        =   4
      Top             =   7560
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "TIME"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   3
      Top             =   10200
      Width           =   1695
   End
   Begin VB.Line Line2 
      X1              =   6360
      X2              =   10080
      Y1              =   10080
      Y2              =   10080
   End
   Begin VB.Line Line1 
      X1              =   6360
      X2              =   6360
      Y1              =   7320
      Y2              =   10080
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14880
      TabIndex        =   2
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SOLAR VOLTAGE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   0
      Top             =   4320
      Width           =   2535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sx1, sx2, sx3, ex1, ex2, ex3, sy1, sy2, sy3, ey1, ey2, ey3, a As Integer
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Command1_Click()
Form2.Hide
Form1.Show
End Sub

Private Sub Command2_Click()
Form2.Show
Form3.Show
End Sub

Private Sub Form_Load()
Form2.Text7.Text = 10
'Form2.Text2.Text = 10
Form2.Text3.Text = 10
'WebBrowser1.Navigate "about:<html><body scroll='no' bgcolor=green><FONT SIZE='25' FACE='Colonna MT' COLOR=white><center><MARQUEE STYLE=overflow WIDTH=70%  BEHAVIOR=scroll scrollamount=10 BGColor=green>RENEWABLE ENERGY</MARQUEE></center><img src = 'D:\vb2017\B-59-GREEN-RADIO\new1.jpg' width=1340 height=100><br><center><MARQUEE WIDTH=90% BEHAVIOR=alternate>Wind and Solar Hybrid Generation</marquee></img></br></FONT></body></html>"
'If (Form2.Text7.Text = 0) Then
' WebBrowser2.Navigate "about:<html><body scroll='no' bgcolor=green><center><FONT SIZE='15' FACE='cooper black' COLOR=white> Wind Energy<font color=green><br><img src ='D:\PROJECTS\GREENRADIO JEPPIAR\vertical.jpg'  width=200 height=200><br><img src ='D:\vb2017\B-59-GREEN-RADIO\windmill1.jpg'  width=200 height=200></center></img></br></FONT></body></html>"
' Else
'  WebBrowser2.Navigate "about:<html><body scroll='no' bgcolor=green><center><FONT SIZE='15' FACE='cooper black' COLOR=white> Wind Energy<font color=green><br><img src ='D:\PROJECTS\GREENRADIO hybrid JEPPIAR\vertical.gif'  width=200 height=200><br><img src ='D:\vb2017\B-59-GREEN-RADIO\windmill1.gif'  width=200 height=200></center></img></br></FONT></body></html>"
'  End If
'  WebBrowser3.Navigate "about:<html><body scroll='no' bgcolor=green><FONT SIZE='15' FACE='cooper black' COLOR=white> <center>Solar Energy<br><font color=green><img src ='D:\vb2017\B-59-GREEN-RADIO\solaar.gif'  width=300 height=300></center></img></br></FONT></body></html>"
sx1 = Line2.X1
sy1 = Line1.Y2
'sx2 = Line4.X1
'sy2 = Line3.Y2

End Sub


Private Sub Form_Unload(Cancel As Integer)

End
End Sub

Private Sub Timer1_Timer()
a = Val(Form2.Text7.Text) + Val(Form2.Text3.Text) + Val(Form1.Text3.Text)

ex1 = sx1 + 100
ey1 = Line1.Y2 - (a / 1200) * (Line1.Y2 - Line1.Y1)
Line (sx1, sy1)-(ex1, ey1), vbBlue
sx1 = ex1
sy1 = ey1
If (sx1 > Line2.X2 - 50) Then
Line (Line1.X1, Line1.Y1)-(Line2.X2, Line2.Y2), Me.BackColor, BF
sx1 = Line2.X1
sy1 = ey1
Line1.Refresh
Line2.Refresh
End If

If Text7.Text > 7 Then
Picture1.Visible = True
Picture2.Visible = False
Else
Picture2.Visible = True
Picture1.Visible = False
End If

'ex2 = sx2 + 100
'ey2 = Line3.Y2 - (Form2.Text4.Text / 1000) * (Line3.Y2 - Line3.Y1)
'Line (sx2, sy2)-(ex2, ey2), vbBlue
'sx2 = ex2
'sy2 = ey2
'If (sx2 > Line4.X2 - 50) Then
'Line (Line3.X1, Line3.Y1)-(Line4.X2, Line4.Y2), Me.BackColor, BF
'sx2 = Line4.X1
'sy2 = ey2
'Line3.Refresh
'Line4.Refresh
'End If
'a = Val(Form2.Text7.Text) + Val(Form2.Text2.Text) + Val(Form2.Text3.Text)
'Form2.Text1.Text = Round(a * Val(Form2.Text4.Text) / 1000)

End Sub

'Private Sub Timer2_Timer()
'Form3.Data1.Recordset.AddNew
'Form3.Data1.Recordset.Fields(0) = Form1.Label15.Caption
'Form3.Data1.Recordset.Fields(1) = Form1.Label16.Caption
'Form3.Data1.Recordset.Fields(2) = Form2.Text7.Text
'Form3.Data1.Recordset.Fields(3) = Form2.Text2.Text
'Form3.Data1.Recordset.Fields(4) = Form2.Text3.Text
'Form3.Data1.Recordset.Fields(5) = Form2.Text4.Text
'Form3.Data1.Recordset.Fields(6) = Form2.Text1.Text
'Form3.Data1.Recordset.Update
'
'End Sub



''Private Sub Timer3_Timer()
''WebBrowser1.Navigate "about:<html><body scroll='no' bgcolor=green><FONT SIZE='25' FACE='Colonna MT' COLOR=white><center><MARQUEE STYLE=overflow WIDTH=70%  BEHAVIOR=scroll scrollamount=10 BGColor=green>RENEWABLE ENERGY</MARQUEE></center><img src = 'D:\vb2017\B-59-GREEN-RADIO\new1.jpg' width=1340 height=100><br><center><MARQUEE WIDTH=90% BEHAVIOR=alternate>Wind and Solar Hybrid Generation</marquee></img></br></FONT></body></html>"
''If (Form2.Text7.Text) < 5 Then
'' WebBrowser3.Navigate "about:<html><body scroll='no' bgcolor=green><FONT SIZE='15' FACE='cooper black' COLOR=white> <center>Solar Energy<br><font color=green><img src ='D:\vb2017\B-59-GREEN-RADIO\no sun.jpg'  width=350 height=350></center></img></br></FONT></body></html>"
'' Else
'' WebBrowser3.Navigate "about:<html><body scroll=no bgcolor=green><FONT SIZE='15' FACE='cooper black' COLOR=white> <center>Solar Energy<br><font color=green><img src ='D:\vb2017\B-59-GREEN-RADIO\sun.jpg'  width=350 height=350></center></img></br></FONT></body></html>"
'' End If
'' If (Form2.Text2.Text) < 5 Then
'' WebBrowser2.Navigate "about:<html><body scroll='no' bgcolor=green><center><FONT SIZE='5' FACE='cooper black' COLOR=white> Vertical Turbine<br><img src ='D:\vb2017\B-59-GREEN-RADIO\vertical.jpg'  width=200 height=200></center></img></br></FONT></body></html>"
'' Else
'' If (Form2.Text2.Text) > 5 Then
'' WebBrowser2.Navigate "about:<html><body scroll='no' bgcolor=green><center><FONT SIZE='5' FACE='cooper black' COLOR=white> Vertical Turbine<br><img src ='D:\vb2017\B-59-GREEN-RADIO\vertical.gif'  width=200 height=200></center></img></br></FONT></body></html>"
'' End If
'' End If
'' If (Form2.Text3.Text) > 5 Then
'' WebBrowser4.Navigate "about:<html><body scroll='no' bgcolor=green><center><h6><FONT size='5' fACE='cooper black' COLOR=white> <center>Horizontal Turbine<font color=green><br><img src ='D:\vb2017\B-59-GREEN-RADIO\windmill1.gif'  width=200 height=180></center></img></br></FONT></body></html>"
'' Else
'' If (Form2.Text3.Text) < 5 Then
'' WebBrowser4.Navigate "about:<html><body scroll='no' bgcolor=green><center><FONT SIZE='5' FACE='cooper black' COLOR=white><center>Horizontal Turbine<font color=green><br><img src ='D:\vb2017\B-59-GREEN-RADIO\windmill1.JPG'  width=200 height=180></center></img></br></FONT></body></html>"
'' End If
'' End If
''    If (Form2.Text2.Text) = 0 Then
''   WebBrowser2.Navigate "about:<html><body scroll='no' bgcolor=green><center><FONT SIZE='5' FACE='cooper black' COLOR=white>vertical turbine<font color=green><br><img src ='D:\vb2017\B-59-GREEN-RADIO\vertical.jpg'  width=200 height=200><br><img src ='D:\PROJECTS\GREENRADIO hybrid JEPPIAR\windmill1.jpg'  width=200 height=200></center></img></br></FONT></body></html>"
''   End If
''
''End Sub
