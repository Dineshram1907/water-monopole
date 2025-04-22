VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   FillColor       =   &H00008000&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   18
      Top             =   6480
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   17
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   16
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   15
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   8
      Top             =   6480
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   7
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   6
      Top             =   4320
      Width           =   1335
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   11280
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   9120
      Top             =   5400
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   0
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label16 
      Caption         =   "mA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   23
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label15 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label14 
      Caption         =   " DEPARTMENT OF IT PAAVAI ENGINEERING COLLEGE"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   21
      Top             =   9600
      Width           =   16935
   End
   Begin VB.Label Label13 
      Caption         =   "ENGINE TEMP GRAPH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   10680
      TabIndex        =   20
      Top             =   1560
      Width           =   5415
   End
   Begin VB.Line Line24 
      X1              =   360
      X2              =   8520
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line23 
      X1              =   360
      X2              =   8520
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line22 
      X1              =   8520
      X2              =   8520
      Y1              =   1440
      Y2              =   7920
   End
   Begin VB.Line Line21 
      X1              =   360
      X2              =   8520
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Line Line20 
      X1              =   360
      X2              =   360
      Y1              =   1440
      Y2              =   7920
   End
   Begin VB.Label Label12 
      Caption         =   "PARAMETERS AND SETPOINTS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   480
      TabIndex        =   19
      Top             =   1560
      Width           =   7695
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   6960
      Shape           =   3  'Circle
      Top             =   6480
      Width           =   375
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   6960
      Shape           =   3  'Circle
      Top             =   5400
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   6960
      Shape           =   3  'Circle
      Top             =   4320
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   6960
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label11 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   14
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "VAC"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   13
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "Turbine Voltage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   600
      TabIndex        =   12
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "Tempereture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   11
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "Fire sensor "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   10
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Turbine Temperature"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   9
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label5 
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
      Left            =   11880
      TabIndex        =   5
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "V A L U E"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   9840
      TabIndex        =   4
      Top             =   3360
      Width           =   255
   End
   Begin VB.Line Line19 
      X1              =   10320
      X2              =   10440
      Y1              =   3240
      Y2              =   3480
   End
   Begin VB.Line Line18 
      X1              =   10320
      X2              =   10200
      Y1              =   3240
      Y2              =   3480
   End
   Begin VB.Line Line17 
      X1              =   10320
      X2              =   10320
      Y1              =   3240
      Y2              =   4320
   End
   Begin VB.Line Line16 
      X1              =   15000
      X2              =   14880
      Y1              =   8640
      Y2              =   8880
   End
   Begin VB.Line Line15 
      X1              =   14880
      X2              =   15000
      Y1              =   8400
      Y2              =   8640
   End
   Begin VB.Line Line14 
      X1              =   13560
      X2              =   15000
      Y1              =   8640
      Y2              =   8640
   End
   Begin VB.Label Label3 
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
      Height          =   615
      Left            =   18360
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "HIGH EFFICIENCY SEA WATER MONOPOLE ANTENNA"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   360
      Width           =   15015
   End
   Begin VB.Label Label1 
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.Line Line13 
      X1              =   2160
      X2              =   2160
      Y1              =   120
      Y2              =   1320
   End
   Begin VB.Line Line12 
      X1              =   18120
      X2              =   18120
      Y1              =   120
      Y2              =   1320
   End
   Begin VB.Line Line11 
      X1              =   120
      X2              =   20160
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line10 
      X1              =   20160
      X2              =   20160
      Y1              =   120
      Y2              =   10920
   End
   Begin VB.Line Line9 
      X1              =   120
      X2              =   20160
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line8 
      X1              =   120
      X2              =   20160
      Y1              =   10920
      Y2              =   10920
   End
   Begin VB.Line Line7 
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   10920
   End
   Begin VB.Line Line6 
      X1              =   16440
      X2              =   16320
      Y1              =   7800
      Y2              =   8040
   End
   Begin VB.Line Line5 
      X1              =   16320
      X2              =   16440
      Y1              =   7560
      Y2              =   7800
   End
   Begin VB.Line Line4 
      X1              =   10800
      X2              =   11040
      Y1              =   2520
      Y2              =   2760
   End
   Begin VB.Line Line3 
      X1              =   10800
      X2              =   10560
      Y1              =   2520
      Y2              =   2760
   End
   Begin VB.Line Line2 
      X1              =   10800
      X2              =   16440
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Line Line1 
      X1              =   10800
      X2              =   10800
      Y1              =   2520
      Y2              =   7800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer, sx As Integer, sy As Integer
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim XX, YY, ZZ As Integer
Dim Buf As String, Out As Integer
Private Sub Form_Load()

MSComm1.PortOpen = True
MSComm1.Output = "{1B00}"
Sleep 100
MSComm1.Output = "{5B00}"
Sleep 100
MSComm1.Output = "{1D00}"
Sleep 100
MSComm1.Output = "{5D00}"
Sleep 100

MSComm1.Output = "{1C80}"
Sleep 100
MSComm1.Output = "{5C80}"
Sleep 100
MSComm1.Output = "{27}"
Sleep 300
sx = Line2.X1
sy = Line2.Y1

Out = &H0
End Sub

Private Sub Timer1_Timer()
''Text1.Text = Round(Rnd * 100)
CH1 = Analog(3)
XX = Analog(0) / 6
YY = Analog(1) / 6
ZZ = Analog(2)

Text1.Text = Round(CH1 / 16)
Text2.Text = Round(XX)
Text3.Text = Round(YY)
Text4.Text = Round(ZZ)


'//////////////////SETPOINTS///////////////////////////////////////////////
Text5.Text = "16"
Text6.Text = "50"
Text7.Text = "50"
Text8.Text = "800"
'//////////////////SETPOINTS///////////////////////////////////////////////
If Val(Text1.Text) < Val(Text5.Text) Then
 Out = Out Or &H1
Shape1.BackColor = vbRed
Else
Out = Out And &HFE
Shape1.BackColor = vbGreen

End If

If XX > Val(Text6.Text) Then
Out = Out Or &H2
Shape2.BackColor = vbRed
Else
 Out = Out And &HFD
Shape2.BackColor = vbGreen

End If

If YY > Val(Text7.Text) Then
Out = Out Or &H4
Shape3.BackColor = vbRed
Else
Out = Out And &HFB
Shape3.BackColor = vbGreen
End If
'
If ZZ > Val(Text8.Text) Then
Out = Out Or &H8
Shape4.BackColor = vbRed
Else
Out = Out And &H7
Shape4.BackColor = vbGreen
End If


MSComm1.Output = "{5D" & CStr(Hex(Out)) & "}"
    Sleep 100
    Form1.Caption = Out




ex = sx + 25
ey = Line1.Y2 - (ZZ / 1023) * (Line1.Y2 - Line1.Y1)
Line (sx, sy)-(ex, ey), vbRed
sx = ex
sy = ey
If (sx > Line2.X2 - 70) Then
Line (Line1.X1, Line1.Y1)-(Line2.X2, Line2.Y2), Me.BackColor, BF
sx = Line2.X1
sy = ey
Line1.Refresh
Line2.Refresh
Line5.Refresh
Line4.Refresh
End If
Label1.Caption = Date
Label3.Caption = Time
End Sub

Function Analog(no As Integer)
    MSComm1.Output = "{4" & CStr(no) & "}"
    Sleep 100
    Buf = MSComm1.Input
    If (Buf <> "") Then
        Analog = CInt(Mid$(Buf, 2, 4))
    Else
        Analog = 0
    End If
End Function

