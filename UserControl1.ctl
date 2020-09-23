VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   3165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3090
   ScaleHeight     =   3165
   ScaleWidth      =   3090
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2040
      TabIndex        =   12
      Text            =   "10"
      Top             =   2580
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Top             =   2640
      Width           =   195
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1080
      Top             =   540
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Int."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      TabIndex        =   11
      Top             =   2640
      Width           =   300
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Off"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   300
      TabIndex        =   10
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape Shape5 
      Height          =   2895
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   315
      Left            =   0
      Top             =   2580
      Width           =   2595
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   160
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   1190
      Width           =   160
   End
   Begin VB.Line Line13 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      X1              =   1275
      X2              =   2325
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "135"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   540
      TabIndex        =   8
      Top             =   1800
      Width           =   330
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "90"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   1185
      TabIndex        =   7
      Top             =   2040
      Width           =   240
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "45"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   1725
      TabIndex        =   6
      Top             =   1800
      Width           =   240
   End
   Begin VB.Line Line6 
      X1              =   420
      X2              =   540
      Y1              =   2220
      Y2              =   2100
   End
   Begin VB.Line Line5 
      X1              =   2040
      X2              =   2175
      Y1              =   2040
      Y2              =   2175
   End
   Begin VB.Line Line4 
      X1              =   1275
      X2              =   1275
      Y1              =   2535
      Y2              =   2385
   End
   Begin VB.Line Line3 
      X1              =   2520
      X2              =   2400
      Y1              =   1275
      Y2              =   1275
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "315"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   1755
      TabIndex        =   5
      Top             =   525
      Width           =   345
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "225"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   525
      Width           =   330
   End
   Begin VB.Line Line16 
      X1              =   450
      X2              =   375
      Y1              =   450
      Y2              =   375
   End
   Begin VB.Line Line2 
      X1              =   2175
      X2              =   2100
      Y1              =   375
      Y2              =   480
   End
   Begin VB.Line Line15 
      X1              =   1275
      X2              =   1275
      Y1              =   150
      Y2              =   0
   End
   Begin VB.Line Line14 
      X1              =   150
      X2              =   0
      Y1              =   1275
      Y2              =   1275
   End
   Begin VB.Line Line1 
      X1              =   2475
      X2              =   2400
      Y1              =   1275
      Y2              =   1275
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   2145
      TabIndex        =   3
      Top             =   1200
      Width           =   120
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "180"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   1200
      Width           =   330
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "270"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   1110
      TabIndex        =   1
      Top             =   225
      Width           =   345
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00808080&
      Height          =   2400
      Left            =   150
      Shape           =   3  'Circle
      Top             =   75
      Width           =   2265
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   660
      TabIndex        =   0
      Top             =   2580
      Width           =   1125
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   2565
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   2565
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Const PIE = 3.14159265 / 96
Private Th As Single
Private R1 As Single
Private R2 As Single
Private Lx1 As Single
Private Ly1 As Single
Private Lx2 As Single
Private Ly2 As Single
Private Lx3 As Single
Private Ly3 As Single

Dim m_Enabled As Boolean
Dim m_Interval As Long
Public Property Get Interval() As Long
Interval = m_Interval
End Property

Public Property Let Interval(ByVal new_value As Long)
    m_Interval = new_value
    PropertyChanged "Interval"
    Timer1.Interval = m_Interval
    Refresh
End Property
Public Property Get Enabled() As Boolean
Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal new_value As Boolean)
    m_Enabled = new_value
    PropertyChanged "Enabled"
    Timer1.Enabled = m_Enabled
    Refresh
End Property



Private Sub Check1_Click()
m_Interval = Val(Text1)
Timer1.Interval = Val(Text1)
If Check1.Value = 0 Then
    Timer1.Enabled = False
    Label10.Caption = "Off"
ElseIf Check1.Value = 1 Then
    Timer1.Enabled = True
    Label10.Caption = "On"
End If
End Sub

Private Sub Timer1_Timer()
    Th = Th + PIE
    Label1.Caption = Format(Th, "##.##")
    Line13.X2 = Lx1 + Cos(Th) * R1
    Line13.Y2 = Ly1 + Sin(Th) * R1
End Sub


Private Sub UserControl_Initialize()
    Dim dx As Single
    Dim dy As Single
    dx = Line13.X2 - Line13.X1
    dy = Line13.Y2 - Line13.Y1
    R1 = Sqr(dx * dx + dy * dy)
    Lx1 = Line13.X1
    Ly1 = Line13.Y1
End Sub

Private Sub UserControl_InitProperties()
m_Enabled = False
m_Interval = 0
End Sub

Private Sub UserControl_Resize()
UserControl.Width = Shape5.Width
UserControl.Height = Shape5.Height
'UserControl.Height = Shape1.Height / 2 + (Shape1.Height / 50)
End Sub
