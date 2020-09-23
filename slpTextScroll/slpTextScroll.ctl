VERSION 5.00
Begin VB.UserControl slpTextScroll 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2040
   PropertyPages   =   "slpTextScroll.ctx":0000
   ScaleHeight     =   48
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   136
   ToolboxBitmap   =   "slpTextScroll.ctx":001D
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   240
      Top             =   180
   End
   Begin VB.PictureBox PicText 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   420
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "slpTextScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Default Property Values:
Const m_def_WordWrap = False
Const m_def_TextAlign = 1
Const m_def_Text = ""
Const m_def_FontColor = &H0
Const m_def_ShadowColor = &H808080
Const m_def_ShadowSize = 3
Const m_def_ShadowDirection = slpDropShadowSouthEast
Const m_def_ScrollDirection = slpVertical
Const m_def_ScrollRepeatCount = 1
'Property Variables:
Dim m_WordWrap As Boolean
Dim m_TextAlign As slpScrollText_TextJustify
Dim m_Text As String
Dim m_FontColor As OLE_COLOR
Dim m_ShadowColor As OLE_COLOR
Dim m_ShadowSize As Integer
Dim m_ShadowDirection As slpScrollText_DropShadowDirection
Dim m_ScrollDirection As slpScrollText_ScrollDirection
Dim m_ScrollRepeatCount As Single
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event ScrollFinished()

'API Declares
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'Private Variables
Private aText() As String
Private mVarTextScrollTop As Single
Private mVarTextScrollLeft As Single
Private mVarScrollCounter As Single

'Private Constants
Private Const slpHorizontalBreakSpace = "    "

Sub About()
Attribute About.VB_UserMemId = -552
    frmAbout.Show vbModal
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PicText,PicText,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    On Error GoTo BSS_ErrorHandler
    BackColor = PicText.BackColor
Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Get BackColor"
    Resume Next
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    On Error GoTo BSS_ErrorHandler
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    Call DrawText

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Let BackColor"
    Resume Next
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "GeneralSettings"
    On Error GoTo BSS_ErrorHandler
    Enabled = UserControl.Enabled

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Get Enabled"
    Resume Next
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    On Error GoTo BSS_ErrorHandler
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Let Enabled"
    Resume Next
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PicText,PicText,-1,Font
Public Property Get Font() As Font
    On Error GoTo BSS_ErrorHandler
    Set Font = PicText.Font

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Get Font"
    Resume Next
End Property

Public Property Set Font(ByVal New_Font As Font)
    On Error GoTo BSS_ErrorHandler
    Set PicText.Font = New_Font
    Call DrawText
    PropertyChanged "Font"

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Set Font"
    Resume Next
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As slpBackStyle
    On Error GoTo BSS_ErrorHandler
    BackStyle = UserControl.BackStyle

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Get BackStyle"
    Resume Next
End Property

Public Property Let BackStyle(ByVal New_BackStyle As slpBackStyle)
    On Error GoTo BSS_ErrorHandler
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Let BackStyle"
    Resume Next
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As slpBorderStyles
    On Error GoTo BSS_ErrorHandler
    BorderStyle = UserControl.BorderStyle

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Get BorderStyle"
    Resume Next
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As slpBorderStyles)
    On Error GoTo BSS_ErrorHandler
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Let BorderStyle"
    Resume Next
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
    On Error GoTo BSS_ErrorHandler
    UserControl.Refresh

Exit Sub

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Sub Refresh"
    Resume Next
End Sub

Private Sub PicText_Click()
    On Error GoTo BSS_ErrorHandler
    RaiseEvent Click

Exit Sub

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Sub PicText_Click"
    Resume Next
End Sub

Private Sub PicText_DblClick()
    On Error GoTo BSS_ErrorHandler
    RaiseEvent DblClick

Exit Sub

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Sub PicText_DblClick"
    Resume Next
End Sub

Private Sub Timer1_Timer()
    On Error GoTo BSS_ErrorHandler
    Dim lResult As Long
    Select Case m_ScrollDirection
        Case slpVertical, slpDefault
            lResult = BitBlt(UserControl.hDC, PicText.left, mVarTextScrollTop, _
                             PicText.ScaleWidth, PicText.ScaleHeight, _
                             PicText.hDC, 0, 0, &HCC0020)
            mVarTextScrollTop = mVarTextScrollTop - 1
            If mVarTextScrollTop < (-PicText.ScaleHeight + UserControl.ScaleHeight) Then
                Call ScrollDone
            End If
        Case slpHorizontal
            lResult = BitBlt(UserControl.hDC, mVarTextScrollLeft, PicText.top, _
                             PicText.ScaleWidth, PicText.ScaleHeight, _
                             PicText.hDC, 0, 0, &HCC0020)
            mVarTextScrollLeft = mVarTextScrollLeft - 1
            If mVarTextScrollLeft < (-PicText.ScaleWidth + UserControl.ScaleWidth) Then
                Call ScrollDone
            End If
    End Select

Exit Sub

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Sub Timer1_Timer"
    Resume Next
End Sub

Private Sub UserControl_Click()
    On Error GoTo BSS_ErrorHandler
    RaiseEvent Click

Exit Sub

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Sub UserControl_Click"
    Resume Next
End Sub

Private Sub UserControl_DblClick()
    On Error GoTo BSS_ErrorHandler
    RaiseEvent DblClick

Exit Sub

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Sub UserControl_DblClick"
    Resume Next
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo BSS_ErrorHandler
    RaiseEvent KeyDown(KeyCode, Shift)

Exit Sub

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Sub UserControl_KeyDown"
    Resume Next
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    On Error GoTo BSS_ErrorHandler
    RaiseEvent KeyPress(KeyAscii)

Exit Sub

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Sub UserControl_KeyPress"
    Resume Next
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo BSS_ErrorHandler
    RaiseEvent KeyUp(KeyCode, Shift)

Exit Sub

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Sub UserControl_KeyUp"
    Resume Next
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo BSS_ErrorHandler
    RaiseEvent MouseDown(Button, Shift, X, Y)

Exit Sub

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Sub UserControl_MouseDown"
    Resume Next
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo BSS_ErrorHandler
    RaiseEvent MouseMove(Button, Shift, X, Y)

Exit Sub

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Sub UserControl_MouseMove"
    Resume Next
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo BSS_ErrorHandler
    RaiseEvent MouseUp(Button, Shift, X, Y)

Exit Sub

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Sub UserControl_MouseUp"
    Resume Next
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    On Error GoTo BSS_ErrorHandler
    PicText.Visible = False
    m_Text = m_def_Text
    m_FontColor = m_def_FontColor
    m_ShadowColor = m_def_ShadowColor
    m_ShadowSize = m_def_ShadowSize
    m_ShadowDirection = m_def_ShadowDirection
    m_ScrollDirection = m_def_ScrollDirection
    m_ScrollRepeatCount = m_def_ScrollRepeatCount
    m_TextAlign = m_def_TextAlign
    m_WordWrap = m_def_WordWrap

Exit Sub

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Sub UserControl_InitProperties"
    Resume Next
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error GoTo BSS_ErrorHandler
    UserControl.BackColor = PropBag.ReadProperty("BackColor", Ambient.BackColor)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set PicText.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    m_Text = PropBag.ReadProperty("Text", m_def_Text)
    m_FontColor = PropBag.ReadProperty("FontColor", m_def_FontColor)
    m_ShadowColor = PropBag.ReadProperty("ShadowColor", m_def_ShadowColor)
    m_ShadowSize = PropBag.ReadProperty("ShadowSize", m_def_ShadowSize)
    m_ShadowDirection = PropBag.ReadProperty("ShadowDirection", m_def_ShadowDirection)
    m_ScrollDirection = PropBag.ReadProperty("ScrollDirection", m_def_ScrollDirection)
    Timer1.Interval = PropBag.ReadProperty("ScrollSpeed", 10)
    m_ScrollRepeatCount = PropBag.ReadProperty("ScrollRepeatCount", m_def_ScrollRepeatCount)
    m_TextAlign = PropBag.ReadProperty("TextAlign", m_def_TextAlign)
    m_WordWrap = PropBag.ReadProperty("WordWrap", m_def_WordWrap)

Exit Sub

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Sub UserControl_ReadProperties"
    Resume Next
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error GoTo BSS_ErrorHandler
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled)
    Call PropBag.WriteProperty("Font", PicText.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle)
    Call PropBag.WriteProperty("Text", m_Text)
    Call PropBag.WriteProperty("FontColor", m_FontColor)
    Call PropBag.WriteProperty("ShadowColor", m_ShadowColor)
    Call PropBag.WriteProperty("ShadowSize", m_ShadowSize)
    Call PropBag.WriteProperty("ShadowDirection", m_ShadowDirection)
    Call PropBag.WriteProperty("ScrollDirection", m_ScrollDirection)
    Call PropBag.WriteProperty("ScrollSpeed", Timer1.Interval)
    Call PropBag.WriteProperty("ScrollRepeatCount", m_ScrollRepeatCount)
    Call PropBag.WriteProperty("TextAlign", m_TextAlign)
    Call PropBag.WriteProperty("WordWrap", m_WordWrap)

Exit Sub

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Sub UserControl_WriteProperties"
    Resume Next
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,""
Public Property Get Text() As String
Attribute Text.VB_ProcData.VB_Invoke_Property = "Options"
    On Error GoTo BSS_ErrorHandler
    Text = m_Text

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Get Text"
    Resume Next
End Property

Public Property Let Text(ByVal New_Text As String)
    On Error GoTo BSS_ErrorHandler
    m_Text = New_Text
    Call DrawText
    PropertyChanged "Text"

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Let Text"
    Resume Next
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H0
Public Property Get FontColor() As OLE_COLOR
    On Error GoTo BSS_ErrorHandler
    FontColor = m_FontColor

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Get FontColor"
    Resume Next
End Property

Public Property Let FontColor(ByVal New_FontColor As OLE_COLOR)
    On Error GoTo BSS_ErrorHandler
    m_FontColor = New_FontColor
    Call DrawText
    PropertyChanged "FontColor"

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Let FontColor"
    Resume Next
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H0
Public Property Get ShadowColor() As OLE_COLOR
    On Error GoTo BSS_ErrorHandler
    ShadowColor = m_ShadowColor

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Get ShadowColor"
    Resume Next
End Property

Public Property Let ShadowColor(ByVal New_ShadowColor As OLE_COLOR)
    On Error GoTo BSS_ErrorHandler
    m_ShadowColor = New_ShadowColor
    Call DrawText
    PropertyChanged "ShadowColor"

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Let ShadowColor"
    Resume Next
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,3
Public Property Get ShadowSize() As Integer
Attribute ShadowSize.VB_ProcData.VB_Invoke_Property = "Options"
    On Error GoTo BSS_ErrorHandler
    ShadowSize = m_ShadowSize

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Get ShadowSize"
    Resume Next
End Property

Public Property Let ShadowSize(ByVal New_ShadowSize As Integer)
    On Error GoTo BSS_ErrorHandler
    m_ShadowSize = New_ShadowSize
    Call DrawText
    PropertyChanged "ShadowSize"

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Let ShadowSize"
    Resume Next
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get ShadowDirection() As slpScrollText_DropShadowDirection
    On Error GoTo BSS_ErrorHandler
    ShadowDirection = m_ShadowDirection

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Get ShadowDirection"
    Resume Next
End Property

Public Property Let ShadowDirection(ByVal New_ShadowDirection As slpScrollText_DropShadowDirection)
    On Error GoTo BSS_ErrorHandler
    m_ShadowDirection = New_ShadowDirection
    Call DrawText
    PropertyChanged "ShadowDirection"

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Let ShadowDirection"
    Resume Next
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get ScrollDirection() As slpScrollText_ScrollDirection
    On Error GoTo BSS_ErrorHandler
    ScrollDirection = m_ScrollDirection

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Get ScrollDirection"
    Resume Next
End Property

Public Property Let ScrollDirection(ByVal New_ScrollDirection As slpScrollText_ScrollDirection)
    On Error GoTo BSS_ErrorHandler
    m_ScrollDirection = New_ScrollDirection
    Call DrawText
    PropertyChanged "ScrollDirection"

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Let ScrollDirection"
    Resume Next
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Timer1,Timer1,-1,Interval
Public Property Get ScrollSpeed() As Long
Attribute ScrollSpeed.VB_ProcData.VB_Invoke_Property = "Options"
    On Error GoTo BSS_ErrorHandler
    ScrollSpeed = Timer1.Interval

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Get ScrollSpeed"
    Resume Next
End Property

Public Property Let ScrollSpeed(ByVal New_ScrollSpeed As Long)
    On Error GoTo BSS_ErrorHandler
    Timer1.Interval() = New_ScrollSpeed
    PropertyChanged "ScrollSpeed"

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Let ScrollSpeed"
    Resume Next
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function StartScroll() As Boolean
    On Error GoTo BSS_ErrorHandler
    Call DrawText
    Timer1.Enabled = True

Exit Function

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Function StartScroll"
    Resume Next
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function StopScroll() As Boolean
    On Error GoTo BSS_ErrorHandler
    Timer1.Enabled = False
    Call DrawText
    mVarScrollCounter = 0

Exit Function

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Function StopScroll"
    Resume Next
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function LoadFromFile(strFilePath As String) As Boolean
    On Error GoTo BSS_ErrorHandler
    Dim iFile As Integer
    Dim strResult As String
    Dim strText As String
    ' First make sure we find the file
    strResult = Dir(strFilePath)
    If Len(strResult) = 0 Then
        LoadFromFile = False
    Else
        ' file exists.. let's load it..
        iFile = FreeFile
        Open strFilePath For Input As #iFile
            strText = Input(LOF(iFile), iFile)
        Close #iFile
        ' Replace all of the CR/LF with Pipes
        strText = Replace(strText, vbCrLf, "|")
        ' Now set the text property of the usercontrol..
        ' That will take care of all the redraws and such..
        Text = strText
        LoadFromFile = True
    End If

Exit Function

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Function LoadFromFile"
    Resume Next
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,1
Public Property Get ScrollRepeatCount() As Single
Attribute ScrollRepeatCount.VB_ProcData.VB_Invoke_Property = "Options"
    On Error GoTo BSS_ErrorHandler
    ScrollRepeatCount = m_ScrollRepeatCount

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Get ScrollRepeatCount"
    Resume Next
End Property

Public Property Let ScrollRepeatCount(ByVal New_ScrollRepeatCount As Single)
    On Error GoTo BSS_ErrorHandler
    m_ScrollRepeatCount = New_ScrollRepeatCount
    PropertyChanged "ScrollRepeatCount"

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Let ScrollRepeatCount"
    Resume Next
End Property

Private Function DrawText() As Boolean
    On Error GoTo BSS_ErrorHandler
    Dim lCounter As Long
    Dim bTimerState As Boolean
    'First we have to store the state of the timer and set it to
    'false.. we don't want it trying to scroll while we are building
    'the picText screen.
    bTimerState = Timer1.Enabled
    Timer1.Enabled = False
    
    'Now determine if we even have any data to work with
    If Len(m_Text) = 0 Then
        DrawText = False
    Else
        Call LoadTextArray
        PicText.Cls
        UserControl.Refresh
        PicText.AutoRedraw = True
        PicText.Visible = False
        PicText.BackColor = UserControl.BackColor
        PicText.ScaleMode = vbPixels
        UserControl.ScaleMode = vbPixels
        ' Establish the height and width of the PicText control based on scroll direction
        Select Case m_ScrollDirection
            Case slpVertical, slpDefault
                PicText.Height = (UBound(aText) * (PicText.TextHeight("A") + m_ShadowSize)) + _
                                  UserControl.ScaleHeight + 20
                For lCounter = 0 To UBound(aText)
                    If PicText.Width < (PicText.TextWidth(aText(lCounter)) + (m_ShadowSize * 2) + 20) Then
                        PicText.Width = PicText.TextWidth(aText(lCounter)) + (m_ShadowSize * 2) + 20
                    End If
                Next
            Case slpHorizontal
                PicText.Height = PicText.TextHeight("A") + m_ShadowSize + 1
                PicText.Width = PicText.TextWidth(aText(0)) + m_ShadowSize + UserControl.ScaleWidth + 20
        End Select
        ' Now to draw the text on the picturebox
        For lCounter = 0 To UBound(aText)
            PrintText aText(lCounter)
        Next
        
        ' Now to setup our drawing Variables
        mVarTextScrollTop = UserControl.ScaleHeight
        mVarTextScrollLeft = UserControl.ScaleWidth
        ' Now to position our PicText based on how we are scrolling
        Select Case m_ScrollDirection
            Case slpVertical, slpDefault
                'Now position based on our alignment..
                Select Case m_TextAlign
                    Case slpJustifyDefault, slpJustifyLeft
                        PicText.left = 1
                    Case slpJustifyRight
                        PicText.left = UserControl.ScaleWidth - PicText.Width
                    Case slpJustifyCenter
                        If PicText.ScaleWidth > UserControl.ScaleWidth Then
                            PicText.left = 0
                        Else
                            PicText.left = (UserControl.ScaleWidth \ 2) - (PicText.Width \ 2)
                        End If
                End Select
            Case slpHorizontal
                'alignment doesn't matter when we are scrolling horizontal..
                If PicText.ScaleHeight > UserControl.ScaleHeight Then
                    ' Must be using a big 'ole font..
                    PicText.top = 0
                Else
                    ' Plenty of room.. let's center the text in the parent control.
                    PicText.top = (UserControl.ScaleHeight \ 2) - (PicText.ScaleHeight \ 2)
                End If
        End Select
    End If
    ' Restore the origional timer state
    Timer1.Enabled = bTimerState

Exit Function

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Function DrawText"
    Resume Next
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get TextAlign() As slpScrollText_TextJustify
    On Error GoTo BSS_ErrorHandler
    TextAlign = m_TextAlign

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Get TextAlign"
    Resume Next
End Property

Public Property Let TextAlign(ByVal New_TextAlign As slpScrollText_TextJustify)
    On Error GoTo BSS_ErrorHandler
    m_TextAlign = New_TextAlign
    Call DrawText
    PropertyChanged "TextAlign"

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Let TextAlign"
    Resume Next
End Property

Private Sub PrintText(Text As String)
    On Error GoTo BSS_ErrorHandler
    Dim lCounter As Long
    Dim X As Single
    Dim Y As Single
    ' First handle the alignment of the text..
    Select Case m_TextAlign
        Case slpJustifyDefault, slpJustifyLeft
            PicText.CurrentX = 1 + m_ShadowSize
        Case slpJustifyRight
            PicText.CurrentX = PicText.ScaleWidth - PicText.TextWidth(Text) - (m_ShadowSize * 2)
        Case slpJustifyCenter
            PicText.CurrentX = (PicText.ScaleWidth \ 2) - (PicText.TextWidth(Text) \ 2) - m_ShadowSize
    End Select
    PicText.ForeColor = m_ShadowColor
    X = PicText.CurrentX
    Y = PicText.CurrentY
    ' Now to deal with the Shadow..
    If m_ShadowSize > 0 Then
        For lCounter = 1 To m_ShadowSize
            PicText.Print Text
            ' Here we work the drop shadow direction
            Select Case m_ShadowDirection
                Case slpDropShadowDefault, slpDropShadowSouthEast
                    X = X - 1
                    Y = Y - 1
                Case slpDropShadowSouthWest
                    X = X + 1
                    Y = Y - 1
                Case slpDropShadowNorthEast
                    X = X - 1
                    Y = Y + 1
                Case slpDropShadowNorthWest
                    X = X + 1
                    Y = Y + 1
            End Select
            PicText.CurrentX = X
            PicText.CurrentY = Y
        Next
    End If
    ' Now to finally print our real text
    PicText.ForeColor = m_FontColor
    PicText.Print Text

Exit Sub

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Sub PrintText"
    Resume Next
End Sub

Private Sub LoadTextArray()
    On Error GoTo BSS_ErrorHandler
    ReDim aText(0)
    If Len(m_Text) > 0 Then
        ' First determine how it is we are going to scroll
        Select Case m_ScrollDirection
            Case slpVertical, slpDefault
                ' going vertical.. we gotta split on the delim..
                aText = Split(m_Text, "|")
                ' Are we going to wordwrap this thing?
                If m_WordWrap Then
                    Call DoWordWrap
                End If
            Case slpHorizontal
                ' Horizontal.. gotta bust out all the delims..
                aText(0) = Replace(m_Text, "|", slpHorizontalBreakSpace)
        End Select
    End If

Exit Sub

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Sub LoadTextArray"
    Resume Next
End Sub

Private Sub ScrollDone()
    On Error GoTo BSS_ErrorHandler
    If m_ScrollRepeatCount > 0 Then
        mVarScrollCounter = mVarScrollCounter + 1
        If mVarScrollCounter >= m_ScrollRepeatCount Then
            Timer1.Enabled = False
            mVarScrollCounter = 0
        End If
    End If
    Call DrawText
    RaiseEvent ScrollFinished

Exit Sub

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Sub ScrollDone"
    Resume Next
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,1,2,false
Public Property Get IsScrolling() As Boolean
    On Error GoTo BSS_ErrorHandler
    IsScrolling = Timer1.Enabled

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Get IsScrolling"
    Resume Next
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,false
Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_ProcData.VB_Invoke_Property = "Options"
    On Error GoTo BSS_ErrorHandler
    WordWrap = m_WordWrap

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Get WordWrap"
    Resume Next
End Property

Public Property Let WordWrap(ByVal New_WordWrap As Boolean)
    On Error GoTo BSS_ErrorHandler
    m_WordWrap = New_WordWrap
    Call DrawText
    PropertyChanged "WordWrap"

Exit Property

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Property Let WordWrap"
    Resume Next
End Property

Private Sub DoWordWrap()
    On Error GoTo BSS_ErrorHandler
    Dim tmpArray() As String
    Dim iLastSpace As Integer
    Dim iMaxWidth As Integer
    Dim sTempStr As String
    Dim iPosition As Integer
    Dim lSrcArrayCounter As Long
    Dim lTargetArrayCounter As Long
    Dim iLastPosition As Integer
    
    iMaxWidth = UserControl.ScaleWidth
    lTargetArrayCounter = 0
    
    For lSrcArrayCounter = LBound(aText) To UBound(aText)
        sTempStr = aText(lSrcArrayCounter)
        If PicText.TextWidth(sTempStr) + (m_ShadowSize * 2) <= iMaxWidth Then
            ' The source text is not wider then the usercontrol.. all is good.
            ReDim Preserve tmpArray(lTargetArrayCounter)
            tmpArray(lTargetArrayCounter) = sTempStr
            lTargetArrayCounter = lTargetArrayCounter + 1
        Else
            ' The source text is wider.. must do some more work.
            iLastPosition = 1
            For iPosition = 1 To Len(sTempStr)
                If PicText.TextWidth(Mid(sTempStr, iLastPosition, iPosition - iLastPosition + 1)) + _
                                    (m_ShadowSize * 2) > iMaxWidth Then
                    ' We have reached our boundry.. time to write out what we got..
                    If iLastSpace > iLastPosition Then
                        ' Ok.. we can wrap on a space..
                        ReDim Preserve tmpArray(lTargetArrayCounter)
                        tmpArray(lTargetArrayCounter) = Mid(sTempStr, iLastPosition, iLastSpace - iLastPosition)
                        ' Inc our target array counter
                        lTargetArrayCounter = lTargetArrayCounter + 1
                        ' Store our position for later use
                        iLastPosition = iLastSpace
                    Else
                        ' There is no space in the last string..
                        ' so we are gonna havta split up the text
                        ReDim Preserve tmpArray(lTargetArrayCounter)
                        tmpArray(lTargetArrayCounter) = Mid(sTempStr, iLastPosition, iPosition - iLastPosition)
                        ' Inc our target array counter
                        lTargetArrayCounter = lTargetArrayCounter + 1
                        ' store our position for later use
                        iLastPosition = iPosition
                    End If
                End If
                ' We are not at our max width, so is this character a space?
                If Mid(sTempStr, iPosition, 1) = " " Then
                    iLastSpace = iPosition
                End If
            Next
            ' Write the last of the wrap string
            If Len(Mid(sTempStr, iLastPosition, iPosition - iLastPosition + 1)) > 0 Then
                ReDim Preserve tmpArray(lTargetArrayCounter)
                tmpArray(lTargetArrayCounter) = Mid(sTempStr, iLastPosition, iPosition - iLastPosition + 1)
                lTargetArrayCounter = lTargetArrayCounter + 1
            End If
        End If
    Next
    ' Now to reset the aText() with the new data
    ReDim aText(UBound(tmpArray))
    aText = tmpArray

Exit Sub

BSS_ErrorHandler:

    If Err.Number > 0 Then ProjectErrorHandler "(User Control) slpTextScroll::Sub DoWordWrap"
    Resume Next
End Sub
