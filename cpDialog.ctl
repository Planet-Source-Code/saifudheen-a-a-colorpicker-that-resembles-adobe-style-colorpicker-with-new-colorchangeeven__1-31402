VERSION 5.00
Begin VB.UserControl ColorBox 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1035
   InvisibleAtRuntime=   -1  'True
   Picture         =   "cpDialog.ctx":0000
   PropertyPages   =   "cpDialog.ctx":0BE2
   ScaleHeight     =   53
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   69
   ToolboxBitmap   =   "cpDialog.ctx":0C0A
End
Attribute VB_Name = "ColorBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Enum Position
    CentreOwner = 0
    CentreScreen = 1
    Manual = 2
End Enum
Enum Public_Mode
    PickerHue = 0
    PickerSaturation = 1
    PickerBrightness = 2
    CustomColors = 3
    SafeColor = 4
End Enum

'Default Property Values:
Const m_def_Mode = 0
Const m_def_About = Null
Const m_def_DialogStartUpPosition = 0
Const m_def_AdjustPosition = True
Const m_def_hWnd = 0
Const m_def_PositionLeft = 0
Const m_def_PositionTop = 0
Const m_def_Color = 0

'Property Variables:
Dim m_Mode As Public_Mode
Dim m_About As Variant
Dim m_DialogStartUpPosition As Position
Dim m_AdjustPosition As Boolean
Dim m_hWnd As Long
Dim m_PositionLeft As Long
Dim m_PositionTop As Long
Dim m_Icon As Picture
Dim ownerHWND As Long
Public Event ColorChange(ByVal NewColor As Long)
Private WithEvents F   As Form  'This will "point" to Form1
Attribute F.VB_VarHelpID = -1



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function Show() As Boolean
    On Error Resume Next
    
    Form1.Hide
    m_hWnd = Form1.hwnd
    Form1.lblSelColor.BackColor = m_Color
    Form1.Label2.BackColor = m_Color
    Dim ownerRect As RECT
    Dim ret As Long
    Dim x As Long, y As Long
    ownerHWND = UserControl.Parent.hwnd
    Select Case m_DialogStartUpPosition
    Case 0 'CentreOwner
            ret = GetWindowRect(ownerHWND, ownerRect)
            x = 15 * (ownerRect.Left + (ownerRect.Right - ownerRect.Left) / 2 - Form1.Width / (15 * 2))
            y = 15 * (ownerRect.Top + (ownerRect.Bottom - ownerRect.Top) / 2 - Form1.Height / (15 * 2))
    Case 1 ' CentreScreen
            x = Screen.Width / 2 - Form1.Width / 2
            y = Screen.Height / 2 - Form1.Height / 2
    Case 2 'Manual
            x = m_PositionLeft * 15
            y = m_PositionTop * 15
    End Select
    
    If m_AdjustPosition Then
        If x < 0 Then x = 0
        If y < 0 Then y = 0
        If x > (Screen.Width - Form1.Width) Then x = (Screen.Width - Form1.Width)
        If y > (Screen.Height - Form1.Height) Then y = (Screen.Height - Form1.Height)
    End If
    Form1.Move x, y
    
    Select Case m_Mode
    Case 0
        If Form1.Mode <> Picker Then Form1.Command3_Click
        Form1.optH.Value = True
    Case 1
        If Form1.Mode <> Picker Then Form1.Command3_Click
        Form1.optS.Value = True
    Case 2
        If Form1.Mode <> Picker Then Form1.Command3_Click

        Form1.optB.Value = True
    Case 3
        If Form1.Mode = Picker Then Form1.Command3_Click
        Form1.cmbPreset.ListIndex = 1
    Case 4
        If Form1.Mode = Picker Then Form1.Command3_Click
        Form1.cmbPreset.ListIndex = 2
    End Select
    
    
    If Not Form1.Visible Then
        Form1.Icon = m_Icon
        Form1.Show vbModal
    End If
    m_Mode = pMode 'If mode changed by user
    PropertyChanged "Mode"
    
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get Color() As OLE_COLOR
Attribute Color.VB_Description = "Stores the Selected color"
    Color = m_Color
End Property

Public Property Let Color(ByVal New_Color As OLE_COLOR)
    If New_Color < 0 Then 'System Color
        New_Color = New_Color + 2147483648# '( &h80,000,000 )
        New_Color = GetSysColor(New_Color)
    End If
    m_Color = New_Color
    PropertyChanged "Color"
End Property
'

Private Sub F_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    F_MouseMove Button, Shift, x, y
End Sub

Private Sub F_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then RaiseEvent ColorChange(Form1.lblSelColor.BackColor)
End Sub

Private Sub UserControl_Initialize()
    Vertion = 2.1 'App vertion
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Color = m_def_Color
    Set m_Icon = LoadPicture("")
    m_PositionLeft = m_def_PositionLeft
    m_PositionTop = m_def_PositionTop
    m_hWnd = m_def_hWnd
    m_AdjustPosition = m_def_AdjustPosition
    m_DialogStartUpPosition = m_def_DialogStartUpPosition
    m_About = m_def_About
    m_Mode = m_def_Mode
End Sub

        
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Color = PropBag.ReadProperty("Color", m_def_Color)
    Set m_Icon = PropBag.ReadProperty("Icon", Nothing)
    m_PositionLeft = PropBag.ReadProperty("DialogLeft", m_def_PositionLeft)
    m_PositionTop = PropBag.ReadProperty("DialogTop", m_def_PositionTop)

    m_hWnd = PropBag.ReadProperty("hWnd", m_def_hWnd)
    m_AdjustPosition = PropBag.ReadProperty("AdjustPosition", m_def_AdjustPosition)
    m_DialogStartUpPosition = PropBag.ReadProperty("DialogStartUpPosition", m_def_DialogStartUpPosition)
    m_Mode = PropBag.ReadProperty("Mode", m_def_Mode)
    
    If Ambient.UserMode Then Set F = Form1 'Point to Form1 to get events
End Sub

Private Sub UserControl_Resize()
    If UserControl.Width <> 480 Then UserControl.Width = 480
    If UserControl.Height <> 280 Then UserControl.Height = 280

End Sub




Private Sub UserControl_Terminate()
    Set F = Nothing
    Set Form1 = Nothing

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Color", m_Color, m_def_Color)
    Call PropBag.WriteProperty("Icon", m_Icon, Nothing)
    Call PropBag.WriteProperty("DialogLeft", m_PositionLeft, m_def_PositionLeft)
    Call PropBag.WriteProperty("DialogTop", m_PositionTop, m_def_PositionTop)
    Call PropBag.WriteProperty("hWnd", m_hWnd, m_def_hWnd)
    Call PropBag.WriteProperty("AdjustPosition", m_AdjustPosition, m_def_AdjustPosition)
    Call PropBag.WriteProperty("DialogStartUpPosition", m_DialogStartUpPosition, m_def_DialogStartUpPosition)
    Call PropBag.WriteProperty("Mode", m_Mode, m_def_Mode)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get Icon() As Picture
Attribute Icon.VB_Description = "Sets the Icon of Color  DialogBox"
    Set Icon = m_Icon
End Property

Public Property Set Icon(ByVal New_Icon As Picture)
    
    If New_Icon.Type = vbPicTypeIcon Then
        Set m_Icon = New_Icon
        PropertyChanged "Icon"
    Else
        MsgBox "Not an Icon", vbOKOnly + vbInformation, "ColorBox"
    End If

End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get DialogLeft() As Long
Attribute DialogLeft.VB_Description = "Sets the initial Left position of Color Dialog"
Attribute DialogLeft.VB_ProcData.VB_Invoke_Property = "General"
    DialogLeft = m_PositionLeft
End Property

Public Property Let DialogLeft(ByVal New_PositionLeft As Long)
    m_PositionLeft = New_PositionLeft
    PropertyChanged "DialogLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get DialogTop() As Long
Attribute DialogTop.VB_Description = "Sets the initial Top position of Color Dialog"
Attribute DialogTop.VB_ProcData.VB_Invoke_Property = "General"
    DialogTop = m_PositionTop
End Property

Public Property Let DialogTop(ByVal New_PositionTop As Long)
    m_PositionTop = New_PositionTop
    PropertyChanged "DialogTop"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,1,2,0
Public Property Get DialoghWnd() As Long
Attribute DialoghWnd.VB_Description = "Handle of Color Dialog"
Attribute DialoghWnd.VB_MemberFlags = "400"
    
    DialoghWnd = m_hWnd
End Property

Public Property Let DialoghWnd(ByVal New_hWnd As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_hWnd = New_hWnd
    PropertyChanged "hWnd"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,TRUE
Public Property Get AdjustPosition() As Boolean
Attribute AdjustPosition.VB_Description = "Adjusts the position of colorDialog such that it remains within the screen "
    AdjustPosition = m_AdjustPosition
End Property

Public Property Let AdjustPosition(ByVal New_AdjustPosition As Boolean)
    m_AdjustPosition = New_AdjustPosition
    PropertyChanged "AdjustPosition"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get DialogStartUpPosition() As Position
Attribute DialogStartUpPosition.VB_Description = "Sets Iniitial position of ColorDialog"
    DialogStartUpPosition = m_DialogStartUpPosition
End Property

Public Property Let DialogStartUpPosition(ByVal New_DialogStartUpPosition As Position)
    m_DialogStartUpPosition = New_DialogStartUpPosition
    PropertyChanged "DialogStartUpPosition"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function Refresh() As Variant
Attribute Refresh.VB_Description = "Refreshes ColorDialog"
    Dim Sho As Boolean
    If Form1.Visible Then Sho = True
    Unload Form1
    If Sho Then Show
End Function


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,1,1,
Public Property Get About() As Variant
Attribute About.VB_ProcData.VB_Invoke_Property = "About"
    On Error Resume Next
  
End Property

Public Property Let About(New_About As Variant)
    On Error Resume Next
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Mode() As Public_Mode
Attribute Mode.VB_Description = "Sets the Mode of startup"
    Mode = m_Mode
End Property

Public Property Let Mode(ByVal New_Mode As Public_Mode)
    m_Mode = New_Mode
    PropertyChanged "Mode"
End Property

