VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FF80&
      Caption         =   "FIRE"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7200
      Width           =   3615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FF80&
      Caption         =   "VOLTAGE"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5760
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "TURBINE TEMP"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "ENGINE TEMP"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      Width           =   3615
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   360
      Top             =   7680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   840
      Top             =   4440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WATER MONOPOLE ANTENNA RECEIVER"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4080
      TabIndex        =   2
      Top             =   360
      Width           =   9225
   End
   Begin VB.Line Line7 
      X1              =   3120
      X2              =   3120
      Y1              =   240
      Y2              =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   14520
      TabIndex        =   0
      Top             =   360
      Width           =   945
   End
   Begin VB.Line Line6 
      X1              =   14400
      X2              =   14400
      Y1              =   240
      Y2              =   960
   End
   Begin VB.Line Line5 
      X1              =   16560
      X2              =   16560
      Y1              =   240
      Y2              =   8880
   End
   Begin VB.Line Line4 
      X1              =   1320
      X2              =   16560
      Y1              =   8880
      Y2              =   8880
   End
   Begin VB.Line Line3 
      X1              =   1320
      X2              =   1320
      Y1              =   240
      Y2              =   8880
   End
   Begin VB.Line Line2 
      X1              =   1320
      X2              =   16560
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      X1              =   1320
      X2              =   16560
      Y1              =   240
      Y2              =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim Val1 As Integer, Val2 As Integer, Val3 As Integer, Val4 As Integer
Dim Buf As String, Out As Integer

Private Sub Form_Load()
    MSComm1.PortOpen = True
    MSComm1.Output = "{24}"
    Sleep 100
    MSComm1.Output = "{1C80}"
    Sleep 100
    MSComm1.Output = "{1D00}"
    Sleep 100
    MSComm1.Output = "{5DFF}"
    Sleep 100
    MSComm1.Output = "{5C80}"
    Sleep 100
    Out = &H0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MSComm1.Output = "{5C80}"
    Sleep 100
    MSComm1.Output = "{5DFF}"
    Sleep 100
End Sub

Private Sub Timer1_Timer()

Label2.Caption = Date
Label1.Caption = Time
    Val1 = Analog(4)
    Val2 = Analog(5)
    Val3 = Analog(6)
    Val4 = Analog(7)
    
    
'    Label2.Caption = Val1
'    Label5.Caption = Val2
'    Label6.Caption = Val3
'    Label8.Caption = Val4
    
    If Val(Val1) < 200 Then
    Command1.BackColor = vbRed
    End If
    
     If Val(Val2) < 200 Then
    Command2.BackColor = vbRed
    End If
    
     If Val(Val3) < 200 Then
    Command3.BackColor = vbRed
    End If
    
     If Val(Val4) < 200 Then
    Command4.BackColor = vbRed
    End If
    
    
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


