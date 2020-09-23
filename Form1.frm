VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ColorBox"
   ClientHeight    =   4140
   ClientLeft      =   1890
   ClientTop       =   1740
   ClientWidth     =   6510
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":0CCA
   ScaleHeight     =   276
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   434
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   -2670
      Top             =   1830
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O&K"
      Height          =   405
      Left            =   5550
      MousePointer    =   1  'Arrow
      TabIndex        =   6
      Top             =   2670
      Width           =   825
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   4530
      TabIndex        =   26
      Top             =   3060
      Width           =   975
      Begin VB.Label lblADDColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   30
         TabIndex        =   27
         Top             =   120
         Visible         =   0   'False
         Width           =   825
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Preset"
      Height          =   405
      Left            =   5550
      TabIndex        =   7
      Top             =   3135
      Width           =   825
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   " << Add "
      Height          =   405
      Left            =   4560
      TabIndex        =   25
      Top             =   3600
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.ComboBox cmbPreset 
      Height          =   315
      Left            =   4650
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   2280
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Frame Frame1 
      Height          =   1185
      Left            =   4620
      MousePointer    =   1  'Arrow
      TabIndex        =   9
      Top             =   60
      Width           =   1725
      Begin VB.Shape Shape1 
         Height          =   735
         Left            =   90
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         Height          =   735
         Left            =   855
         TabIndex        =   12
         Top             =   360
         Width           =   780
      End
      Begin VB.Label lblSelColor 
         BackColor       =   &H000000FF&
         Height          =   735
         Left            =   90
         TabIndex        =   10
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label3 
         Caption         =   "Old"
         Height          =   195
         Left            =   1050
         TabIndex        =   13
         Top             =   150
         Width           =   285
      End
      Begin VB.Label Label1 
         Caption         =   "New"
         Height          =   195
         Left            =   300
         TabIndex        =   11
         Top             =   150
         Width           =   465
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   2235
      Left            =   4590
      TabIndex        =   14
      Top             =   960
      Width           =   2055
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "0"
         ToolTipText     =   "Blue"
         Top             =   1200
         Width           =   405
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "0"
         ToolTipText     =   "Green"
         Top             =   840
         Width           =   405
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "255"
         ToolTipText     =   "Red"
         Top             =   480
         Width           =   405
      End
      Begin VB.TextBox txtH 
         Height          =   285
         Left            =   450
         MaxLength       =   3
         MousePointer    =   4  'Icon
         TabIndex        =   3
         Text            =   "0"
         ToolTipText     =   "Hue"
         Top             =   480
         Width           =   405
      End
      Begin VB.TextBox txtS 
         Height          =   285
         Left            =   450
         MaxLength       =   3
         MousePointer    =   4  'Icon
         TabIndex        =   4
         Text            =   "100"
         ToolTipText     =   "Saturation"
         Top             =   840
         Width           =   405
      End
      Begin VB.TextBox txtB 
         Height          =   285
         Left            =   450
         MaxLength       =   3
         MousePointer    =   4  'Icon
         TabIndex        =   5
         Text            =   "100"
         ToolTipText     =   "Brightness"
         Top             =   1200
         Width           =   405
      End
      Begin VB.OptionButton optH 
         Caption         =   "H:"
         Height          =   255
         Left            =   0
         TabIndex        =   0
         ToolTipText     =   "Hue"
         Top             =   540
         Value           =   -1  'True
         Width           =   465
      End
      Begin VB.OptionButton optS 
         Caption         =   "S:"
         Height          =   255
         Left            =   0
         TabIndex        =   1
         ToolTipText     =   "Saturation"
         Top             =   885
         Width           =   465
      End
      Begin VB.OptionButton optB 
         Caption         =   "B:"
         Height          =   255
         Left            =   0
         TabIndex        =   2
         ToolTipText     =   "Brightness"
         Top             =   1230
         Width           =   495
      End
      Begin VB.Label lblR 
         Caption         =   "R:"
         Height          =   225
         Left            =   1140
         TabIndex        =   23
         ToolTipText     =   "Red"
         Top             =   540
         Width           =   195
      End
      Begin VB.Label lblG 
         Caption         =   "G:"
         Height          =   225
         Left            =   1140
         TabIndex        =   22
         ToolTipText     =   "Green"
         Top             =   885
         Width           =   225
      End
      Begin VB.Label lblB 
         Caption         =   "B:"
         Height          =   225
         Left            =   1140
         TabIndex        =   21
         ToolTipText     =   "Blue"
         Top             =   1230
         Width           =   195
      End
      Begin VB.Label lblS 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   900
         TabIndex        =   20
         Top             =   870
         Width           =   165
      End
      Begin VB.Label lblBB 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   900
         TabIndex        =   19
         Top             =   1230
         Width           =   195
      End
      Begin VB.Label lblH 
         Caption         =   "°"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   900
         TabIndex        =   18
         Top             =   510
         Width           =   105
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   5550
      MousePointer    =   1  'Arrow
      TabIndex        =   8
      Top             =   3600
      Width           =   825
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'************************ColorBox OCX Vertion 2.0************************
'Main dialog
'Author : Saifudheen A.A.
'E-mail : keraleeyan@msn.com.
'This ColorPicker was developed for my Paint programme
' Suggestions, Votes all are welcome.
'********************************************************************
Enum sMode
    Picker = 0
    About = 1
    Custom = 2
    SafeColor = 3
End Enum

Dim MainBoxHit  As Boolean
Dim SelectBoxHit As Boolean
Public Mode As sMode
Public UserMode As Integer 'usercontrol Mode variable
Dim HueEntering As Boolean
Dim SaturationEntering As Boolean
Dim BrightnessEntering As Boolean
Dim OldHue As Integer, OldSaturation As Integer, OldBrightness As Integer

Private Sub LoadHueShades()
    Dim OldP As POINTAPI
    Dim V As Integer
    On Error Resume Next
    Dim H As Single, M As Single
    Dim a As Single, b As Single, C As Single, D As Single, E As Single, F As Single
    Dim Ratio As Single
    Dim sDc As Long
    
    H = SelectBox.Bottom - SelectBox.Top
    M = H / 6
    a = M
    b = 2 * M
    C = 3 * M
    D = 4 * M
    E = 5 * M
    F = 6 * M
    Dim sBitmap(0 To 16 * 256 * 3) As Byte            '256 ^ 2 * 3
    Dim cpos  As Long
    With lpBI.bmiHeader
        .biHeight = 256
        .biWidth = 15
    End With
    
    cpos = 0
    With Me '1
        .DrawWidth = 1
        .DrawMode = 13
        sDc = .hdc
        For y = 0 To Int(a)
            For j = 1 To 16
                sBitmap(cpos + 2) = 255
                sBitmap(cpos + 1) = 0
                sBitmap(cpos + 0) = y * 6
                cpos = cpos + 3
            Next j
        Next y
    '2
                
        For y = Int(a) + 1 To Int(b)
            V = 255 - (y - a) * 6
            For j = 1 To 16
                sBitmap(cpos + 2) = V
                sBitmap(cpos + 1) = 0
                sBitmap(cpos + 0) = 255
                cpos = cpos + 3
            Next j
            
        Next y
     '3
         
        For y = Int(b) + 1 To Int(C)
            V = (y - b) * 6
            For j = 1 To 16
                sBitmap(cpos + 2) = 0
                sBitmap(cpos + 1) = V
                sBitmap(cpos + 0) = 255
                cpos = cpos + 3
            Next j
            
        Next y
     '4
        For y = Int(C) + 1 To Int(D)
            V = 255 - (y - C) * 6
            For j = 1 To 16
                sBitmap(cpos + 2) = 0
                sBitmap(cpos + 1) = 255
                sBitmap(cpos + 0) = V
                cpos = cpos + 3
            Next j
        Next y
    '5
        For y = Int(D) + 1 To Int(E)
            V = (y - D) * 6
            For j = 1 To 16
                sBitmap(cpos + 2) = V
                sBitmap(cpos + 1) = 255
                sBitmap(cpos + 0) = 0
                cpos = cpos + 3
            Next j
            
        Next y
    '6
        For y = Int(E) + 1 To Int(F)
            V = 255 - (y - E) * 6
            For j = 1 To 16
                sBitmap(cpos + 2) = 255
                sBitmap(cpos + 1) = V
                sBitmap(cpos + 0) = 0
                cpos = cpos + 3
            Next j
        Next y
    End With
            BltBitmap Me.hdc, sBitmap, MainBox.Left, MainBox.Bottom, 15, -256, True
            DrawMainFrame

End Sub


Private Sub cmbPreset_Click()
Me.Cls
On Error GoTo r:
Select Case cmbPreset.ListIndex
Case 0
    lblADDColor.Visible = False
    cmdADD.Visible = False
    Me.Refresh
    Mode = About
    PrintAbout Form1.hdc
    
Case 1
    pMode = 3
    Me.DrawStyle = 0
    LoadCustomColors
    cmdADD.Visible = True
    lblADDColor.Visible = True
    Mode = Custom
    
Case 2
    pMode = 4
    Me.DrawStyle = 0
    LoadSafePalette
    cmdADD.Visible = False
    lblADDColor.Visible = False
    Mode = SafeColor
End Select
Exit Sub
r:
Exit Sub
'MsgBox Error
End Sub

Private Sub cmdADD_Click()
    svdColor(cPaletteIndex) = lblADDColor.BackColor
    DrawSafePicker cPaletteIndex, True 'Erase
    Form1.DrawWidth = 1
    Form1.DrawMode = 13
    Form1.FillStyle = 0
    Form1.ForeColor = 0
    Form1.FillColor = svdColor(cPaletteIndex)
    Rectangle Form1.hdc, Preset(cPaletteIndex).Left, Preset(cPaletteIndex).Top, Preset(cPaletteIndex).Right, Preset(cPaletteIndex).Bottom
    SaveCustomColors
    DrawSafePicker cPaletteIndex, False
    lblSelColor.BackColor = svdColor(cPaletteIndex)
    Form1.Refresh
End Sub

Private Sub Command1_Click()
    If Mode = About Then cmbPreset.ListIndex = 2
    m_Color = lblSelColor.BackColor
    Me.Hide
End Sub



Private Sub Command2_Click()
    Me.lblSelColor.BackColor = mOldColor
    Form1.Hide
End Sub

Sub Command3_Click()
Me.Cls
Dim r As Integer, g As Integer, b As Integer
Dim HexV As String
If Mode = Picker Then
    Frame2.Visible = False
    Mode = Custom
    LoadCustomColors
    Command3.Caption = "&Picker"
    cmbPreset.Visible = True
    cmbPreset.ListIndex = 1
    lblADDColor.BackColor = lblSelColor.BackColor
    cmdADD.Visible = True
    lblADDColor.Visible = True
    GetRGB lblSelColor.BackColor, r, g, b
    GetHexVal r, g, b, HexV
    PrintRGBHEX r, g, b, HexV
    
Else
    Me.Cls
    cmbPreset.Visible = False
    Command3.Caption = "&Preset"
    DrawPicker
    Select Case True
    Case optH
         optH_Click
    Case optS
        optS_Click
    Case optB
        optB_Click
    End Select
   
    DrawSlider SelectedMainPos
    Mode = Picker
    Frame2.Visible = True
    cmdADD.Visible = False
    lblADDColor.Visible = False
End If


End Sub



Function GetSafeColor(Index As Integer, r As Integer, g As Integer, b As Integer, Optional HexVal As String) As Long
Dim i As Long, j As Long, k As Long
Dim count As Integer
Dim strR As String, strG As String, strB As String

For i = 0 To &HFF Step &H33
    For j = 0 To &HFF Step &H33
        For k = 0 To &HFF Step &H33
            count = count + 1
            If count = Index Then
                r = i: g = j: b = k
                GetSafeColor = RGB(i, j, k)
                GetHexVal r, g, b, HexVal
                Exit Function
            End If
           
        Next k
    Next j
Next i

End Function

Sub GetHexVal(Red As Integer, Green As Integer, Blue As Integer, strHex As String)
    Dim strR As String, strG As String, strB As String
    strR = Trim(Hex(Red))
        If Len(strR) = 1 Then strR = "0" & strR
    strG = Trim(Hex(Green))
        If Len(strG) = 1 Then strR = "0" & strR
    strB = Trim(Hex(Blue))
        If Len(strB) = 1 Then strR = "0" & strR
        strHex = strR & strG & strB

End Sub






Private Sub Form_Load()
    
    ' Set these Parameters on Basis of where  should Select Box and  Main Box Should appear
    Dim OldP As POINTAPI
    ReDim Preset(1 To 224)  'RECT structure
    Preset(1).Left = 10
    Preset(1).Top = 10
    Preset(1).Right = 25
    Preset(1).Bottom = 25
    Mode = Picker   'Initialize Mode as  normal Picker
    With lpBI.bmiHeader
        .biBitCount = 24
        .biCompression = 0&
        .biPlanes = 1
        .biSize = Len(lpBI.bmiHeader)
    End With
    
    '// Setting the position of RECTS for Safe Colors
    For i = 2 To 224
        If i Mod 16 = 0 Then
            Preset(i).Top = Preset(i - 1).Top
            Preset(i).Left = Preset(i - 1).Right + 3
            Preset(i).Bottom = Preset(i).Top + 15
            Preset(i).Right = Preset(i).Left + 15
            If i = 224 Then GoTo Jump
            i = i + 1
            Preset(i).Top = Preset(i - 1).Bottom + 3
            Preset(i).Left = Preset(1).Left
        Else
            Preset(i).Top = Preset(i - 1).Top
            Preset(i).Left = Preset(i - 1).Right + 3
        End If
Jump:
        Preset(i).Bottom = Preset(i).Top + 15
        Preset(i).Right = Preset(i).Left + 15
    Next i

    
    k = 255
    SelectBox.Left = 10
    SelectBox.Top = 10
    SelectBox.Right = SelectBox.Left + k
    SelectBox.Bottom = SelectBox.Top + k
    MainBox.Left = SelectBox.Right + 12
    MainBox.Top = SelectBox.Top
    MainBox.Right = MainBox.Left + 15
    MainBox.Bottom = SelectBox.Bottom
     
    With cmbPreset
    .AddItem "About ColorBox"
    .AddItem "Custom..."
    .AddItem "Safe Palette (216)"
    End With

     
    LoadHueShades
    SelectedMainPos = MainBox.Top
    SelectedPos.x = SelectBox.Right
    SelectedPos.y = SelectBox.Top
    Call DrawSlider(SelectedMainPos)
    DrawPicker
    LoadVariantsHue 255, 0, 0
    cPaletteIndex = 1
    Me.ForeColor = vbBlack

    txtH.Text = 360
    txtB.Text = 100
    txtS.Text = 100
    
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If Mode = Picker Then
            If x >= SelectBox.Left And x <= SelectBox.Right And y >= SelectBox.Top And y <= SelectBox.Bottom Then
                '// In SelectBox Boundary
                SelectBoxHit = True
                Me.MousePointer = vbCustom
                'Me.MousePointer = vbCustom
               Call MouseOnSelectBox(x, y)
            End If
            If x >= MainBox.Left And x < MainBox.Right + 11 And y >= MainBox.Top - 2 And y < MainBox.Bottom + 3 Then
                '// In MainBox Boundary
                MainBoxHit = True
                If y > MainBox.Bottom Then y = MainBox.Bottom
                If y < MainBox.Top Then y = MainBox.Top
                Call MouseOnMainBox(x, y)
            End If
        Else
            If Mode <> About Then
                HandlePresetValues x, y, Mode
            End If
        End If
        
        
    End If
End Sub

Sub DrawSafePicker(Index As Integer, Clear As Boolean)
    Dim r As RECT
    Dim l As Long
    Me.FillStyle = 1
    Me.DrawMode = 13
    Me.DrawWidth = 3
    If Clear Then
        Me.ForeColor = Me.BackColor
        Rectangle Me.hdc, Preset(Index).Left - 2, Preset(Index).Top - 2, Preset(Index).Right + 2, Preset(Index).Bottom + 2
    Else
        r.Left = Preset(Index).Left - 3
        r.Top = Preset(Index).Top - 3
        r.Right = Preset(Index).Right + 3
        r.Bottom = Preset(Index).Bottom + 3
        Call DrawEdge(Form1.hdc, r, BDR_SUNKENINNER Or BDR_SUNKENOUTER, BF_RECT Or BF_SOFT)
    End If
    
End Sub

Private Sub MouseOnSelectBox(x As Single, y As Single)
            Dim cl As Long
            Dim r As Integer
            Dim g As Integer
            Dim b As Integer
            DrawPicker
            SelectedPos.x = x
            SelectedPos.y = y
            DrawPicker
            cl = GetPixel(Me.hdc, x, y)
            GetRGB cl, r, g, b
            If optS.Value Then
                LoadMainSaturation Form1.hdc, r, g, b
                txtB.Text = Int((255 - SelectedPos.y + SelectBox.Top) * 100 / 255) 'Brightness Level
                txtH.Text = Int((SelectedPos.x - SelectBox.Left) * 360 / 255) 'Hue Level
            End If
            
            If optB.Value Then
                LoadMainBrightness Form1.hdc, r, g, b
                txtS.Text = Int((SelectBox.Bottom - SelectedPos.y) * 100 / 255)    ' Saturation Level
                txtH.Text = Int((SelectedPos.x - SelectBox.Left) * 360 / 255)   'Hue Level
            End If

            If optH.Value Then
                txtS.Text = Int((SelectedPos.x - SelectBox.Left) * 100 / 255)   ' Saturation Level
                txtB.Text = Int((255 - SelectedPos.y + SelectBox.Top) * 100 / 255) 'Brightness Level
            Else
                cl = GetPixel(Me.hdc, MainBox.Left + 5, SelectedMainPos)
                GetRGB cl, r, g, b
            End If
            Text1.Text = r
            Text2.Text = g
            Text3.Text = b
            lblSelColor.BackColor = cl

End Sub

Private Sub MouseOnMainBox(x As Single, y As Single)
        Dim cl As Long
        Dim r As Integer
        Dim g As Integer
        Dim b As Integer

            DrawSlider SelectedMainPos
            DrawSlider y
            SelectedMainPos = y
            cl = Me.Point(MainBox.Left + 5, SelectedMainPos)
            GetRGB cl, r, g, b
            
            If optH.Value Then
                txtH.Text = Int((255 - y + SelectBox.Top) * 360 / 255)
                GetColorFromHSB Val(txtS.Text), Val(txtB.Text)
            Else
                Text1.Text = r
                Text2.Text = g
                Text3.Text = b
                lblSelColor.BackColor = cl
            End If
            If optS.Value Then
                txtS.Text = Int((255 - y + SelectBox.Top) * 100 / 255)
                
            End If
            If optB.Value Then
                txtB.Text = Int((255 - y + SelectBox.Top) * 100 / 255)
            End If

End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim pt As POINTAPI
    Dim ClipRect As RECT
    pt.x = x
    pt.y = y
    
    If Button = 1 Then
        If Mode = Picker Then
            If SelectBoxHit Then
                
                If x < SelectBox.Left Then x = SelectBox.Left
                If x > SelectBox.Right Then x = SelectBox.Right
                If y < SelectBox.Top Then y = SelectBox.Top
                If y > SelectBox.Bottom Then y = SelectBox.Bottom
                '// In SelectBox Region
                
                pt.x = 10
                pt.y = 10
                ClientToScreen Me.hwnd, pt
                ClipRect.Left = pt.x
                ClipRect.Top = pt.y
                
                pt.x = 266
                pt.y = 266
                ClientToScreen Me.hwnd, pt
                ClipRect.Right = pt.x
                ClipRect.Bottom = pt.y
                ClipCursor ClipRect
                Call MouseOnSelectBox(x, y)
            End If
            
            If MainBoxHit Then
                '// In MainBox region
                x = MainBox.Left + 2
                If y > MainBox.Bottom Then y = MainBox.Bottom
                If y < MainBox.Top Then y = MainBox.Top
                Call MouseOnMainBox(x, y)
            End If
        Else
            If Mode <> About Then
                HandlePresetValues x, y, Mode
            End If
        End If
    End If
End Sub

Private Sub HandlePresetValues(x As Single, y As Single, pMode As sMode)
            Dim i As Integer
            Dim r As Integer, g As Integer, b As Integer
            Dim HexV As String
            For i = 1 To 224
                If x > Preset(i).Left And x < Preset(i).Right And y > Preset(i).Top And y < Preset(i).Bottom Then
                    DrawSafePicker cPaletteIndex, True
                    DrawSafePicker i, False
                    Me.Refresh
                    cPaletteIndex = i
                    Select Case pMode
                    Case 1
                        
                    Case 2
                        lblSelColor.BackColor = svdColor(i)
                        GetRGB svdColor(i), r, g, b
                        GetHexVal r, g, b, HexV
                    Case 3
                        lblSelColor.BackColor = GetSafeColor(i, r, g, b, HexV)
                        
                    End Select
                    
                    PrintRGBHEX r, g, b, HexV
                    Exit Sub
                End If
            Next i

End Sub

Private Sub PrintRGBHEX(r As Integer, g As Integer, b As Integer, HexV As String)
    Me.FontSize = 8
    Me.FillColor = Me.BackColor
    Me.ForeColor = Me.BackColor
    Me.DrawMode = 13
    Me.Line (MainBox.Right + 20, 90)-(MainBox.Right + 20 + 190, 90 + 50), Me.BackColor, BF
    Me.CurrentX = MainBox.Right + 20
    Me.CurrentY = 90
    Me.ForeColor = 0
    Me.Print "R: " & r
    Me.CurrentX = MainBox.Right + 20
    Me.Print "G: " & g
    Me.CurrentX = MainBox.Right + 20
    Me.Print "B: " & b
    Me.CurrentX = MainBox.Right + 20
    Me.Print "Hex: #" & HexV

End Sub

Private Sub GetColorFromHSB(ByVal Sat As Integer, ByVal br As Integer)
    '//
    '// This Function Evaluates the Resulting Color Value While Sliding the Hue Shades From Insisting Brightness and Saturation Values
    Dim r As Integer, g As Integer, b As Integer
    Dim Red As Integer, Green As Integer, Blue As Integer
    Dim cl As Long
    Dim x As Long, y As Long
    On Error Resume Next
    x = Sat * (255 / 100)
    cl = Me.Point(MainBox.Left + 5, SelectedMainPos)
    GetRGB cl, Red, Green, Blue
    r = (Red + ((255 - Red) * (100 - Sat)) / 100) * br / 100
    g = (Green + ((255 - Green) * (100 - Sat)) / 100) * br / 100
    b = (Blue + ((255 - Blue) * (100 - Sat)) / 100) * br / 100
    lblSelColor.BackColor = RGB(r, g, b)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim em As RECT
    em.Right = Screen.Width / 15
    em.Bottom = Screen.Height / 15
    ClipCursor em
    SelectBoxHit = False
    Me.MousePointer = vbDefault
    If MainBoxHit And optH.Value Then
        Dim r As Integer, g As Integer, b As Integer
        Dim cl As Long
        cl = GetPixel(Me.hdc, MainBox.Left + 3, SelectedMainPos)
        GetRGB cl, r, g, b
        LoadVariantsHue r, g, b
        cl = Me.Point(SelectedPos.x, SelectedPos.y)
        GetRGB cl, r, g, b
        Text1.Text = r
        Text2.Text = g
        Text3.Text = b
        lblSelColor.BackColor = Me.Point(SelectedPos.x, SelectedPos.y)
        
    End If
    MainBoxHit = False
End Sub





Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        If UnloadMode = vbFormControlMenu Then
            Cancel = True
            Me.Hide
        End If
End Sub

Private Sub Label2_Click()
    Dim r As Integer, g As Integer, b As Integer
    Dim HexV As String
    GetRGB Label2.BackColor, r, g, b
    GetHexVal r, g, b, HexV
    PrintRGBHEX r, g, b, HexV
End Sub

Private Sub Label2_DblClick()
    lblSelColor.BackColor = Label2.BackColor

End Sub

Private Sub lblADDColor_Click()
    Dim r As Integer, g As Integer, b As Integer
    Dim HexV As String
    GetRGB lblADDColor.BackColor, r, g, b
    GetHexVal r, g, b, HexV
    PrintRGBHEX r, g, b, HexV
End Sub

Private Sub lblSelColor_Click()
    Dim r As Integer, g As Integer, b As Integer
    Dim HexV As String
    GetRGB lblSelColor.BackColor, r, g, b
    GetHexVal r, g, b, HexV
    PrintRGBHEX r, g, b, HexV

End Sub

Private Sub optB_Click()
    Dim r As Integer, g As Integer, b As Integer
    Dim C As Long
    pMode = 2
    LoadVariantsBrightness  '// Loading the Shades at SelectBox
    DrawPicker 'Erase picker
    SelectedPos.x = SelectBox.Left + Val(txtH.Text) * 255 / 360  'Depending on Hue
    SelectedPos.y = SelectBox.Bottom - Val(txtB.Text) * 255 / 100 'Depending on Saturation
    DrawPicker
    C = Me.Point(SelectedPos.x, SelectedPos.y)
    GetRGB C, r, g, b
    LoadMainBrightness Form1.hdc, r, g, b

End Sub

Private Sub optH_Click()
    Dim cl As Long
    Dim r As Integer, g As Integer, b As Integer
    pMode = 0
    LoadHueShades
    DrawSlider SelectedMainPos
    SelectedMainPos = Int(MainBox.Bottom - Val(txtH.Text) * 255 / 360)
    DrawSlider SelectedMainPos
    cl = Me.Point(MainBox.Left + 3, SelectedMainPos)
    GetRGB cl, r, g, b
    LoadVariantsHue r, g, b
    DrawPicker
    SelectedPos.x = SelectBox.Left + Val(txtS.Text) * 255 / 100
    SelectedPos.y = SelectBox.Bottom - Val(txtB.Text) * 255 / 100
    DrawPicker
    cl = Me.Point(SelectedPos.x, SelectedPos.y)
    lblSelColor.BackColor = cl
    
End Sub

Private Sub optS_Click()
    Dim r As Integer, g As Integer, b As Integer
    Dim C As Long
    pMode = 1
    LoadVariantsSaturation  '// Loading the Shades at SelectBox
    DrawPicker 'Erase picker
    SelectedPos.x = SelectBox.Left + Val(txtH.Text) * 255 / 360  'Depending on Hue
    SelectedPos.y = SelectBox.Bottom - Val(txtB.Text) * 255 / 100 'Depending on Brightness
    DrawPicker
    C = Me.Point(SelectedPos.x, SelectedPos.y)
    GetRGB C, r, g, b
    LoadMainSaturation Form1.hdc, r, g, b
    Me.Refresh
End Sub


 
Sub PrintAbout(hdc As Long)
    Dim qrc As RECT, brdr As RECT
    qrc.Left = 20
    qrc.Top = 90
    qrc.Right = 280
    qrc.Bottom = 180
    brdr.Left = 10
    brdr.Top = 10
    brdr.Right = 290
    brdr.Bottom = 265
    'Me.Line (13, 13)-(287, 262), RGB(32, 100, 100), B
    Me.ForeColor = 0
    Me.CurrentX = 70
    Me.CurrentY = 150
    Me.FontSize = 10
    Me.Print "ColorBox ver " & Vertion
    Me.CurrentX = 70
    Me.Print "Author: Saifudheen. A. A. "
    Me.CurrentX = 70
    Me.Print "keraleeyan@msn.com"
    Me.CurrentX = 70
    Me.Print "Copyright © (2001) SaifSoft inc."
    'DrawEdge hdc, qrc, BDR_RAISEDINNER, BF_RECT
    DrawEdge hdc, brdr, BDR_SUNKENOUTER Or BDR_RAISEDINNER, BF_RECT
End Sub



Private Sub Timer1_Timer()
    HueEntering = False
    SaturationEntering = False
    BrightnessEntering = False
    Timer1.Enabled = False
End Sub

Private Sub txtB_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not BrightnessEntering Then
        OldBrightness = Val(txtB.Text)
    End If
    Timer1.Enabled = True
    BrightnessEntering = True
End Sub

Private Sub txtB_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrB
    AdjustBrightness txtB.Text
    Exit Sub
ErrB:
    MsgBox "Enter an Integer Value Between 0 and 100", , "ColorBox"
    txtB.Text = OldBrightness
    txtB.SetFocus

End Sub

Sub DrawPicker()
    Me.DrawMode = 6
    Me.FillStyle = 1
    Me.DrawWidth = 1
    Me.DrawStyle = 0
    Me.Circle (SelectedPos.x, SelectedPos.y), 5
End Sub

Sub DrawSelFrame()
    Dim SelFrame As RECT
    SelFrame.Left = SelectBox.Left - 1
    SelFrame.Top = SelectBox.Top - 1
    SelFrame.Right = SelectBox.Right + 3
    SelFrame.Bottom = SelectBox.Bottom + 3
    DrawEdge Me.hdc, SelFrame, BDR_SUNKENINNER, BF_RECT
    Me.Refresh
End Sub

Sub DrawMainFrame()
    Dim MainFrame As RECT
     
    MainFrame.Left = MainBox.Left - 1
    MainFrame.Top = MainBox.Top - 1
    MainFrame.Right = MainBox.Right + 1
    MainFrame.Bottom = MainBox.Bottom + 3
    DrawEdge Me.hdc, MainFrame, BDR_SUNKENINNER, BF_RECT
    Me.Refresh
End Sub

Private Sub UpdateColorValues()
        Dim r As Integer, g As Integer, b As Integer
        Dim cl As Long
        cl = Me.Point(SelectedPos.x, SelectedPos.y)
        lblSelColor.BackColor = cl
        GetRGB cl, r, g, b
        Text1.Text = r
        Text2.Text = g
        Text3.Text = b

End Sub



Private Sub txtB_LostFocus()
    BrightnessEntering = False
End Sub

Private Sub txtH_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not HueEntering Then
        OldHue = Val(txtH.Text)
    End If
    HueEntering = True
    Timer1.Enabled = True
End Sub

Private Sub txtH_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrH
AdjustHue txtH.Text
Exit Sub
ErrH:
    MsgBox "Enter an Integer Value Between 0 and 360", , "ColorBox"
    txtH.SetFocus
    txtH.Text = OldHue
End Sub

Private Sub txtH_LostFocus()
    HueEntering = False
End Sub




Private Sub txtS_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not SaturationEntering Then
        OldSaturation = Val(txtS.Text)
    End If
    SaturationEntering = True
    Timer1.Enabled = True

End Sub

Private Sub txtS_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrS
    AdjustSaturation txtS.Text
    Exit Sub
ErrS:
    MsgBox "Enter an Integer Value Between 0 and 100", , "ColorBox"
    txtS.Text = OldSaturation
    txtS.SetFocus

End Sub

Private Sub AdjustHue(ByVal Hue As Single)
        Dim cl As Long
        Dim r As Integer, g As Integer, b As Integer
        If Hue > 360 Or Hue < 0 Or Int(Hue) <> Hue Then
            MsgBox "Enter an Integer Value Between 0 and 360", , "ColorBox"
            txtH.SetFocus
            txtH.Text = OldHue
            AdjustHue OldHue
            Exit Sub
        End If

        DrawPicker
        DrawSlider SelectedMainPos

        Select Case True
        Case optH.Value
            SelectedMainPos = MainBox.Bottom - Hue * (255 / 360)
            cl = Me.Point(MainBox.Left + 5, SelectedMainPos)
            GetRGB cl, r, g, b
            LoadVariantsHue r, g, b
            UpdateColorValues
            DrawPicker
        Case optS.Value
            SelectedPos.x = 10 + Hue * (255 / 360)
            cl = Me.Point(SelectedPos.x, SelectedPos.y)
            GetRGB cl, r, g, b
            LoadMainSaturation Me.hdc, r, g, b
            cl = Me.Point(MainBox.Left + 5, SelectedMainPos)
            lblSelColor.BackColor = cl
        Case optB.Value
            SelectedPos.x = 10 + Hue * (255 / 360)
            cl = Me.Point(SelectedPos.x, SelectedPos.y)
            GetRGB cl, r, g, b
            LoadMainSaturation Me.hdc, r, g, b
            cl = Me.Point(MainBox.Left + 5, SelectedMainPos)
            lblSelColor.BackColor = cl
        End Select
        DrawPicker
        DrawSlider SelectedMainPos
End Sub

Private Sub AdjustSaturation(ByVal Saturation As Single)
        Dim cl As Long
        Dim r As Integer, g As Integer, b As Integer
        If Saturation > 100 Or Saturation < 0 Or Int(Saturation) <> Saturation Then
            MsgBox "Enter an Integer Value Between 0 and 100", , "ColorBox"
            txtS.SetFocus
            txtS.Text = OldSaturation
            AdjustSaturation OldSaturation
            Exit Sub
        End If

        DrawPicker
        DrawSlider SelectedMainPos
        Select Case True
        Case optH.Value
            SelectedPos.x = SelectBox.Left + Saturation * (255 / 100)
            cl = Me.Point(SelectedPos.x, SelectedPos.y)
            lblSelColor.BackColor = cl
        Case optS.Value
            SelectedMainPos = MainBox.Bottom - Saturation * (255 / 100)
            cl = Me.Point(MainBox.Left + 5, SelectedMainPos)
            lblSelColor.BackColor = cl
        Case optB.Value
            SelectedPos.y = SelectBox.Bottom - Saturation * (255 / 100)
            cl = Me.Point(SelectedPos.x, SelectedPos.y)
            GetRGB cl, r, g, b
            LoadMainBrightness Me.hdc, r, g, b
            cl = Me.Point(MainBox.Left + 5, SelectedMainPos)
            lblSelColor.BackColor = cl
        End Select
        DrawPicker
        DrawSlider SelectedMainPos
End Sub

Private Sub AdjustBrightness(ByVal Brightness As Single)
        Dim r As Integer, g As Integer, b As Integer
        Dim cl As Long
        If Brightness > 100 Or Brightness < 0 Or Int(Brightness) <> Brightness Then
            MsgBox "Enter an Integer Value Between 0 and 100", , "ColorBox"
            txtB.Text = OldBrightness
            txtB.SetFocus
            AdjustBrightness OldBrightness
            Exit Sub
        End If

        DrawPicker
        DrawSlider SelectedMainPos

        Select Case True
        Case optH.Value
            SelectedPos.y = SelectBox.Bottom - Brightness * (255 / 100)
            cl = Me.Point(SelectedPos.x, SelectedPos.y)
            lblSelColor.BackColor = cl
        Case optS.Value
            SelectedPos.y = SelectBox.Bottom - Brightness * (255 / 100)
            SelectedPos.x = 10 + Brightness * (255 / 360)
            cl = Me.Point(SelectedPos.x, SelectedPos.y)
            GetRGB cl, r, g, b
            LoadMainSaturation Me.hdc, r, g, b
            cl = Me.Point(MainBox.Left + 5, SelectedMainPos)
            lblSelColor.BackColor = cl
        Case optB.Value
            SelectedMainPos = MainBox.Bottom - Brightness * (255 / 100)
            SelectedPos.y = SelectBox.Bottom - Brightness * (255 / 100)
            SelectedPos.x = 10 + Brightness * (255 / 360)
            cl = Me.Point(SelectedPos.x, SelectedPos.y)
            GetRGB cl, r, g, b
            LoadMainBrightness Me.hdc, r, g, b
            cl = Me.Point(MainBox.Left + 5, SelectedMainPos)
            lblSelColor.BackColor = cl
        End Select
        DrawPicker
        DrawSlider SelectedMainPos

End Sub

Private Sub txtS_LostFocus()
    SaturationEntering = False
End Sub
