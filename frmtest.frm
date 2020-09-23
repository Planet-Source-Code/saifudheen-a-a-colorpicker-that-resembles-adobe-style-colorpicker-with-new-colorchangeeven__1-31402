VERSION 5.00
Object = "*\AcpDialogprj.vbp"
Begin VB.Form Form1 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6300
   ClientLeft      =   4590
   ClientTop       =   2490
   ClientWidth     =   8085
   DrawMode        =   7  'Invert
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmtest.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   420
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   539
   WindowState     =   2  'Maximized
   Begin ColorBox_ColorPicker.ColorBox ColorBox1 
      Left            =   3840
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   503
      AdjustPosition  =   0   'False
      DialogStartUpPosition=   2
      Mode            =   1
   End
   Begin VB.Frame Frame3 
      Height          =   675
      Left            =   240
      TabIndex        =   10
      Top             =   4350
      Width           =   3195
      Begin VB.CheckBox Check2 
         Caption         =   "Full Screen"
         Height          =   285
         Left            =   480
         TabIndex        =   11
         Top             =   270
         Value           =   1  'Checked
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dialog Position"
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   1770
      Width           =   3195
      Begin VB.CommandButton Command1 
         Caption         =   "E&xit"
         Height          =   315
         Left            =   2160
         TabIndex        =   12
         Top             =   450
         Width           =   915
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Adjust Position."
         Height          =   255
         Left            =   330
         TabIndex        =   8
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1785
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   345
         Left            =   2460
         TabIndex        =   5
         Text            =   "-10"
         Top             =   1440
         Width           =   585
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   330
         Left            =   2460
         TabIndex        =   4
         Text            =   "800"
         Top             =   1080
         Width           =   585
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Manual"
         Height          =   285
         Left            =   300
         TabIndex        =   3
         Top             =   1110
         Width           =   1125
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Centre Screen"
         Height          =   255
         Left            =   300
         TabIndex        =   2
         Top             =   765
         Width           =   2175
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Centre Owner"
         Height          =   255
         Left            =   300
         TabIndex        =   1
         Top             =   420
         Value           =   -1  'True
         Width           =   2265
      End
      Begin VB.Label Label3 
         Caption         =   "Right Click overrides settings"
         Height          =   315
         Left            =   180
         TabIndex        =   9
         Top             =   2040
         Width           =   2805
      End
      Begin VB.Label Label2 
         Caption         =   "Y :"
         Enabled         =   0   'False
         Height          =   225
         Left            =   2160
         TabIndex        =   7
         Top             =   1500
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "X :"
         Enabled         =   0   'False
         Height          =   225
         Left            =   2160
         TabIndex        =   6
         Top             =   1140
         Width           =   255
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   -30
      Top             =   5490
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Dim BkColor As Long
Dim FkColor As Long

Private Sub Check1_Click()
    If Check1.Value = vbChecked Then
        ColorBox1.AdjustPosition = True
    Else
        ColorBox1.AdjustPosition = False
        
    End If
End Sub

Private Sub Check2_Click()
    With Me
    If Check2.Value = vbChecked Then
        Me.BorderStyle = vbBSNone
        Me.WindowState = vbMaximized
        Me.FontSize = 45
    Else
        Me.BorderStyle = vbSizable
        Me.WindowState = vbNormal
        Me.FontSize = 25
    End If
    End With
End Sub



Private Sub ColorBox1_ColorChange(ByVal NewColor As Long)
    Me.BackColor = NewColor
    
End Sub

Private Sub Command1_Click()
        Unload Form1
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
            Unload Form1
    End If
End Sub

Private Sub Form_Load()
    BkColor = vbWhite
    Me.FontSize = 45
    Me.FontBold = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        If Button = 1 Then
            With ColorBox1
            Select Case True
            Case Option3.Value
                .DialogStartUpPosition = CentreOwner
            Case Option4.Value
                .DialogStartUpPosition = CentreScreen
            Case Option5.Value
                .DialogStartUpPosition = Manual
            End Select
            .DialogLeft = Text1.Text
            .DialogTop = Text2.Text
            .Color = Me.BackColor
            .Show
            End With
            Me.BackColor = ColorBox1.Color
        End If
        If Button = 2 Then
            ColorBox1.DialogStartUpPosition = Manual
            ColorBox1.DialogLeft = x
            ColorBox1.DialogTop = y
            ColorBox1.Color = Me.ForeColor
            ColorBox1.Mode = PickerSaturation
            ColorBox1.Show
            l = ColorBox1.Mode
            Me.ForeColor = ColorBox1.Color
        End If
    
End Sub

Private Sub Form_Paint()
    Dim tm As String
    tm = CStr(Time)
    Me.CurrentX = ((Me.Width + 15 * Frame2.Width) / 30 - Me.TextWidth(tm) / 2)
    Me.CurrentY = (Me.Height / 30 - Me.TextHeight(tm) / 2)
    Me.Print Time
End Sub

Private Sub Option3_Click()
    ColorBox1.DialogStartUpPosition = CentreOwner
    Text1.Enabled = False
    Text2.Enabled = False
    Label1.Enabled = False
    Label2.Enabled = False

End Sub

Private Sub Option4_Click()
    ColorBox1.DialogStartUpPosition = CentreScreen
    Text1.Enabled = False
    Text2.Enabled = False
    Label1.Enabled = False
    Label2.Enabled = False

End Sub

Private Sub Option5_Click()
    ColorBox1.DialogStartUpPosition = Manual
    Text1.Enabled = True
    Text2.Enabled = True
    Label1.Enabled = True
    Label2.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Me.Cls
    Form_Paint
End Sub
