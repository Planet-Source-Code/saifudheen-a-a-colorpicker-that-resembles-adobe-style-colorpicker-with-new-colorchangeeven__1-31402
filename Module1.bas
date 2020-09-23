Attribute VB_Name = "Module1"


'************************ColorBox Vertion 2.0************************
'Functions module; Color algorithms
'Author : Saifudheen A.A.
'E-mail : keraleeyan@msn.com.
'This ColorPicker was developed for my Paint programme
' Suggestions, Votes all are welcome.
'********************************************************************
'Type Declerations
Public Type RGBTRIPLE
    rgbtBlue As Byte
    rgbtGreen As Byte
    rgbtRed As Byte
End Type

Type POINTAPI
    x As Long
    y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type BITMAPINFOHEADER       ' 40 bytes
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As Long
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type

Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmicolors(15) As Long
End Type

Public Type PALETTEENTRY
        peRed As Byte
        peGreen As Byte
        peBlue As Byte
        peFlags As Byte
End Type

Public Type LOGPALETTE
        palVersion As Integer
        palNumEntries As Integer
        palPalEntry() As PALETTEENTRY
End Type
'API Function Declerations
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SetICMMode Lib "gdi32" (ByVal hdc As Long, ByVal n As Long) As Long
Public Declare Function CheckColorsInGamut Lib "gdi32" (ByVal hdc As Long, lpv As Any, lpv2 As Any, ByVal dw As Long) As Long

Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long

Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, pt As POINTAPI) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Declare Function CreatehalfTonePalette Lib "gdi32" Alias "CreateHalftonePalette" (ByVal hdc As Long) As Long
Public Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long

Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long



' Constants
Public Vertion   As Single   'Product vertion
 
Public Const BitsPixel = 12
Public Const Planes = 14

Public Const BDR_RAISEDINNER = &H4

Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2

Public Const BF_RIGHT = &H4
Public Const BF_TOP = &H2
Public Const BF_LEFT = &H1
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_SOFT = &H1000

Public Const ICM_ON = 2
Public Const DIB_RGB_COLORS = 0 '  color table in RGBs
Public Const DIB_PAL_COLORS = 1 '  color table in palette indices

'Variables
Public pMode As Integer
Dim lpbmINFO As BITMAPINFO
Public lpBI As BITMAPINFO
Public m_Color As Long
Public mOldColor As Long
Public SelectBox As RECT
Public MainBox As RECT
Public Preset() As RECT
Public SelectedPos As POINTAPI
Public SelectedMainPos As Single
Public cPaletteIndex As Integer
Public svdColor() As Long






Sub LoadVariantsHue(Red As Integer, Green As Integer, Blue As Integer)
  
    'On Error Resume Next
    Dim x As Integer, y As Integer
    Dim sDc As Long
    Dim K1 As Double, K2 As Double, K3 As Double
    Form1.DrawMode = 6
    Form1.Circle (SelectedPos.x, SelectedPos.y), 5  'Erases Previous Circle
    K1 = Red / 255
    K2 = Green / 255
    K3 = Blue / 255
    With Form1
        .DrawWidth = 1
        .DrawMode = 13
        
        Dim M1    As Double, M2     As Double, M3     As Double
        Dim J1    As Double, J2     As Double, J3     As Double
        Dim YMax As Byte
        Dim shdBitmap(0 To 196608) As Byte  '256 ^ 2 * 3
        Dim l As Long
        Dim bpos As Long
        Dim count As Long
        bpos = 0
        count = 0
        
        With lpBI.bmiHeader
            .biHeight = 256
            .biWidth = 256
        End With
        
        On Error Resume Next
        For y = 255 To 0 Step -1
                 M1 = Red - y * K1
                 M2 = Green - y * K2
                 M3 = Blue - y * K3
                 YMax = 255 - y
                 J1 = (YMax - M1) / 255
                 J2 = (YMax - M2) / 255
                 J3 = (YMax - M3) / 255
            For x = 255 To 0 Step -1
                shdBitmap(bpos) = M3 + x * J3    'Blue
                shdBitmap(bpos + 1) = M2 + x * J2    'Green
                shdBitmap(bpos + 2) = M1 + x * J1     'Red
                bpos = bpos + 3
            Next x
        Next y
        
    BltBitmap Form1.hdc, shdBitmap, 10, 10, 256, 256, True
       
        
    End With
    Form1.DrawSelFrame
    Form1.DrawMode = 6
    Form1.Circle (SelectedPos.x, SelectedPos.y), 5, 1       'Refresh Circle
End Sub




Sub LoadVariantsBrightness()
Dim OldP As POINTAPI
Dim V As Integer
On Error Resume Next
Dim H, M As Single
Dim a As Integer, b As Integer, C As Integer, D As Integer, E As Integer, F As Integer
Dim sDc As Long
Dim Color As Long
Dim Red As Integer, Green As Integer, Blue As Integer
H = SelectBox.Bottom - SelectBox.Top
M = H / 6
a = M
b = 2 * M
C = 3 * M
D = 4 * M
E = 5 * M
F = 6 * M

Dim bBitmap(0 To 256 ^ 2 * 3) As Byte '256 ^ 2 * 3


Form1.DrawMode = 6
Form1.DrawWidth = 1
Form1.Circle (SelectedPos.x, SelectedPos.y), 5  'Erases Previous Circle
With Form1
    .DrawMode = 13
    sDc = .hdc
End With

With lpBI.bmiHeader
    .biHeight = 256
    .biWidth = 256
End With

    Dim Maa As Double, Mcc As Double, Mee As Double
    Dim MV As Double
    Dim Kc As Integer
    Dim YPos As Long
    Maa = 255 + 6 * a  'These are the common Terms taken out from the For Loop for efficciency
    Mcc = 255 + 6 * C  ' ""
    Mee = 255 + 6 * E  ' ""
    Dim pos As Long
    pos = 0
    
Dim x  As Integer, y As Integer
For y = 255 To 0 Step -1
        MV = 1 - y / 255 ' ""
    '1
        For x = 0 To a
            V = x * 6
            Kc = V * MV + y
            bBitmap(pos + 2) = 255
            bBitmap(pos + 1) = y
            bBitmap(pos + 0) = Kc
            pos = pos + 3
        Next x
    '2
        For x = a + 1 To b
            V = Maa - 6 * x ' 255 - (X - A) * 6
            Kc = V * MV + y
            bBitmap(pos + 2) = Kc
            bBitmap(pos + 1) = y
            bBitmap(pos + 0) = 255
            pos = pos + 3
        Next x
     '3
        For x = b + 1 To C
            V = (x - b - 1) * 6
            Kc = V * MV + y
            bBitmap(pos + 2) = y
            bBitmap(pos + 1) = Kc
            bBitmap(pos + 0) = 255
            pos = pos + 3
        Next x
     '4
        For x = C + 1 To D
            V = Mcc - 6 * x
            Kc = V * MV + y
            bBitmap(pos + 2) = y
            bBitmap(pos + 1) = 255
            bBitmap(pos + 0) = Kc
            pos = pos + 3
        Next x
    '5
        For x = D + 1 To E
            V = (x - D - 1) * 6
            Kc = V * MV + y
            bBitmap(pos + 2) = Kc
            bBitmap(pos + 1) = 255
            bBitmap(pos + 0) = y
            pos = pos + 3
        Next x
    '6
        For x = E + 1 To F
            V = Mee - 6 * x
            Kc = V * MV + y
            bBitmap(pos + 2) = 255
            bBitmap(pos + 1) = Kc
            bBitmap(pos + 0) = y
            pos = pos + 3
        Next x
       
Next y

    BltBitmap Form1.hdc, bBitmap, 265, 10, -256, 256, True
    Form1.DrawSelFrame
    Form1.DrawMode = 6
    Form1.Circle (SelectedPos.x, SelectedPos.y), 5 'Refresh Circle
    
End Sub


Sub LoadVariantsSaturation()
Dim OldP As POINTAPI
Dim V As Integer
On Error Resume Next
Dim H, M As Single
Dim x As Integer, y As Integer
Dim a As Integer, b As Integer, C As Integer, D As Integer, E As Integer, F As Integer
Dim sDc As Long
Dim Color As Long
Dim Red As Integer, Green As Integer, Blue As Integer
Dim bBitmap(0 To 256 ^ 2 * 3) As Byte '256 ^ 2 * 3
Dim cpos As Long
cpos = 0
H = SelectBox.Bottom - SelectBox.Top
M = H / 6
a = M
b = 2 * M
C = 3 * M
D = 4 * M
E = 5 * M
F = 6 * M

Form1.DrawMode = 6
Form1.Circle (SelectedPos.x, SelectedPos.y), 5  'Erases Previous Circle

With Form1
    .DrawWidth = 1
    .DrawMode = 13
    sDc = .hdc
End With
    Dim Maa As Double, Mcc As Double, Mee As Double
    Dim MV As Double
    Dim Kc As Integer
    Dim YPos As Long
    Maa = 255 + 6 * a  'These are the common Terms taken out from the For Loop for efficiency
    Mcc = 255 + 6 * C  ' ""
    Mee = 255 + 6 * E  ' ""
    
For y = 255 To 0 Step -1
        MV = 1 - y / 255 ' ""
        YPos = SelectBox.Top + y
    '1
        For x = 0 To a
            V = x * 6
            Kc = V * MV
            bBitmap(pos + 2) = 255 - y
            bBitmap(pos + 1) = 0
            bBitmap(pos + 0) = Kc
            pos = pos + 3
        Next x
    '2
        For x = a + 1 To b
            V = Maa - 6 * x
            Kc = V * MV
            bBitmap(pos + 2) = Kc
            bBitmap(pos + 1) = 0
            bBitmap(pos + 0) = 255 - y
            pos = pos + 3
        Next x
     '3
        For x = b + 1 To C
            V = (x - b - 1) * 6
            Kc = V * MV
            bBitmap(pos + 2) = 0
            bBitmap(pos + 1) = Kc
            bBitmap(pos + 0) = 255 - y
            pos = pos + 3
        Next x
     '4
        For x = C + 1 To D
            V = Mcc - 6 * x
            Kc = V * MV
            bBitmap(pos + 2) = 0
            bBitmap(pos + 1) = 255 - y
            bBitmap(pos + 0) = Kc
            pos = pos + 3
        Next x
    '5
        For x = D + 1 To E
            V = (x - D - 1) * 6
            Kc = V * MV
            bBitmap(pos + 2) = Kc
            bBitmap(pos + 1) = 255 - y
            bBitmap(pos + 0) = 0
            pos = pos + 3
        Next x
    '6
        For x = E + 1 To F
            V = Mee - 6 * x
            Kc = V * MV
            bBitmap(pos + 2) = 255 - y
            bBitmap(pos + 1) = Kc
            bBitmap(pos + 0) = 0
            pos = pos + 3
        Next x
       
Next y
    BltBitmap Form1.hdc, bBitmap, 265, 10, -256, 256, True
    Form1.DrawSelFrame
    Form1.DrawMode = 6
    Form1.Circle (SelectedPos.x, SelectedPos.y), 5 'Refresh Circle

End Sub

Sub GetRGB(ByRef cl As Long, ByRef Red As Integer, ByRef Green As Integer, ByRef Blue As Integer)
    Dim C As Long
    C = cl
    Red = C Mod &H100
    C = C \ &H100
    Green = C Mod &H100
    C = C \ &H100
    Blue = C Mod &H100
End Sub

Sub DrawSlider(ByVal Position As Integer)
    Form1.DrawMode = 6
    Form1.DrawWidth = 2
    Form1.Line (MainBox.Right + 2, Position)-(MainBox.Right + 5, Position)
    Form1.Line (MainBox.Left - 2, Position)-(MainBox.Left - 5, Position)
    Form1.DrawWidth = 1
End Sub

Sub LoadSafePalette()
Form1.FillStyle = 0
Form1.DrawMode = 13
Form1.DrawWidth = 1
On Error Resume Next
Dim i, j, k As Integer
Dim l As Long
Dim count As Integer
Dim Plt As Long
Dim ret As Long
Dim br As Long
Dim pal As Long, oldpal As Long
pal = CreatehalfTonePalette(Form1.hdc)
oldpal = SelectPalette(Form1.hdc, pal, 0)
RealizePalette (Form1.hdc)

For i = 0 To &HFF Step &H33
    For j = 0 To &HFF Step &H33
        For k = 0 To &HFF Step &H33
            count = count + 1
            DrawSafeColor Preset(count), i, j, k
        Next k
    Next j
Next i


For i = 217 To 224
    Form1.FillColor = 0
    Rectangle Form1.hdc, Preset(i).Left, Preset(i).Top, Preset(i).Right, Preset(i).Bottom
Next i

SelectPalette Form1.hdc, oldpal, 0
DeleteObject pal

Form1.DrawSafePicker cPaletteIndex, False
Dim r As Integer, g As Integer, b As Integer
Form1.GetSafeColor cPaletteIndex, r, g, b
Form1.lblSelColor.BackColor = RGB(r, g, b)
End Sub

Sub LoadCustomColors()
    Dim FileHandle As Integer
    Dim i As Integer
    Dim strColor As String
    On Error Resume Next
    FileHandle = FreeFile()
    ReDim svdColor(0 To 224)
    Open App.Path & "/usercolors.cps" For Input As #FileHandle
    i = 0
    Form1.Cls
    Form1.FillStyle = 0
    Form1.DrawMode = 13
    Form1.DrawWidth = 1
    For i = 0 To 224
        Line Input #FileHandle, strColor
        svdColor(i) = Val(strColor)
        Form1.ForeColor = vbBlack 'svdColor(i)
        Form1.FillColor = svdColor(i)
        Rectangle Form1.hdc, Preset(i).Left, Preset(i).Top, Preset(i).Right, Preset(i).Bottom
    Next i
    Close #FileHandle
    Form1.DrawSafePicker cPaletteIndex, False
    Form1.Refresh
End Sub

Sub SaveCustomColors()
    Dim FileHandle As Integer
    Dim i As Integer
    On Error Resume Next
    FileHandle = FreeFile()
    Open App.Path & "/usercolors.cps" For Output As #FileHandle
    For i = 0 To 224
        Print #FileHandle, svdColor(i)
    Next i
    Close #FileHandle
  
End Sub



Public Sub LoadMainSaturation(ByVal hdc As Long, ByVal Red As Integer, ByVal Green As Integer, ByVal Blue As Integer)

    Dim r As Integer, g As Integer, b As Integer
    Dim bBitmap(17 * 256 * 3) As Byte
    Dim pos As Integer
    Dim x As Long, y As Long
    Dim F As Single
    For y = 0 To 255
        F = 1 - (y / 255)
        r = Red * F + y
        g = Green * F + y
        b = Blue * F + y

        For x = 0 To 15

            bBitmap(pos) = b
            bBitmap(pos + 1) = g
            bBitmap(pos + 2) = r
            pos = pos + 3
        Next x
    Next y

    BltBitmap hdc, bBitmap, 277, 265, 15, -256, True
    Form1.DrawMainFrame
End Sub

Public Sub LoadMainBrightness(ByVal hdc As Long, ByVal Red As Integer, ByVal Green As Integer, ByVal Blue As Integer)
    Dim r As Integer, g As Integer, b As Integer
    Dim bBitmap(17 * 256 * 3) As Byte
    Dim pos As Integer
    Dim x As Long, y As Long
    For y = 0 To 255
        r = Red - Red * y / 255
        g = Green - Green * y / 255
        b = Blue - Blue * y / 255
        For x = 0 To 15
            bBitmap(pos) = b
            bBitmap(pos + 1) = g
            bBitmap(pos + 2) = r
            pos = pos + 3
        Next x
    Next y
    
    BltBitmap hdc, bBitmap, 277, 265, 15, -256, True
    Form1.DrawMainFrame
End Sub
   

Public Sub DrawSafeColor(sPos As RECT, ByVal Red As Integer, ByVal Green As Integer, ByVal Blue As Integer)

    Dim bBitmap(3 * 16 * 16) As Byte
    Dim pos As Integer
    Dim x As Long, y As Long
    Dim Width As Integer
    Dim Height As Integer
    Width = sPos.Bottom - sPos.Top
    Height = sPos.Bottom - sPos.Top
    For y = 0 To Height
        For x = 0 To Width
                bBitmap(pos) = Blue
                bBitmap(pos + 1) = Green
                bBitmap(pos + 2) = Red
                pos = pos + 3
        Next x
    Next y
    BltBitmap Form1.hdc, bBitmap, sPos.Left, sPos.Top, Width, Height, False


End Sub

Public Sub BltBitmap(ByVal hdc As Long, bmptr() As Byte, ByVal x As Integer, ByVal y As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal CreatehfTPalette As Boolean)
        Dim lpBI As BITMAPINFO
        lpBI.bmiHeader.biBitCount = 24
        lpBI.bmiHeader.biCompression = BI_RGB
        lpBI.bmiHeader.biWidth = Abs(Width)
        lpBI.bmiHeader.biHeight = Abs(Height)
        lpBI.bmiHeader.biPlanes = 1
        lpBI.bmiHeader.biSize = 40
        If CreatehfTPalette Then
            Dim pal As Long, oldpal As Long
            pal = CreatehalfTonePalette(hdc)
            oldpal = SelectPalette(hdc, pal, 0)
            RealizePalette (hdc)
        End If
        StretchDIBits hdc, x, y, Width, Height, 0, 0, Abs(Width), Abs(Height), bmptr(0), lpBI, DIB_RGB_COLORS, vbSrcCopy
        If CreatehfTPalette Then
            SelectPalette hdc, oldpal, 0
            DeleteObject pal
        End If
End Sub
