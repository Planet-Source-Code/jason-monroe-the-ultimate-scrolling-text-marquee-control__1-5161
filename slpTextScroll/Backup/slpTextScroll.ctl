VERSION 5.00
Begin VB.UserControl slpTextScroll 
   ClientHeight    =   720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2040
   ScaleHeight     =   48
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   136
   ToolboxBitmap   =   "slpTextScroll.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
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
'Default Property Values:
Const m_def_WordWrap = False
Const m_def_TextAlign = 1
Const m_def_Text = ""
Const m_def_FontColor = &H0
Const m_def_ShadowColor = &H0
Const m_def_ShadowSize = 3
Const m_def_ShadowDirection = 1
Const m_def_ScrollDirection = 0
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
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event ScrollFinished()

'API Declares
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'Private Variables
Private aText() As String
Private mVarTextScrollTop As Single
Private mVarTextScrollLeft As Single
Private mVarScrollCounter As Single

'Private Constants
Private Const slpHorizontalBreakSpace = "    "

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PicText,PicText,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = PicText.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    Call DrawText
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PicText,PicText,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = PicText.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set PicText.Font = New_Font
    Call DrawText
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As slpBackStyle
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As slpBackStyle)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As slpBorderStyles
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As slpBorderStyles)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

Private Sub PicText_Click()
    RaiseEvent Click
End Sub

Private Sub Timer1_Timer()
    Dim lResult As Long
    Select Case m_ScrollDirection
        Case slpVertical
            lResult = BitBlt(UserControl.hDC, PicText.Left, mVarTextScrollTop, _
                             PicText.ScaleWidth, PicText.ScaleHeight, _
                             PicText.hDC, 0, 0, &HCC0020)
            mVarTextScrollTop = mVarTextScrollTop - 1
            If mVarTextScrollTop < (-PicText.ScaleHeight + UserControl.ScaleHeight) Then
                Call ScrollDone
            End If
        Case slpHorizontal
            lResult = BitBlt(UserControl.hDC, mVarTextScrollLeft, PicText.Top, _
                             PicText.ScaleWidth, PicText.ScaleHeight, _
                             PicText.hDC, 0, 0, &HCC0020)
            mVarTextScrollLeft = mVarTextScrollLeft - 1
            If mVarTextScrollLeft < (-PicText.ScaleWidth + UserControl.ScaleWidth) Then
                Call ScrollDone
            End If
    End Select
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
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
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set PicText.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    m_Text = PropBag.ReadProperty("Text", m_def_Text)
    m_FontColor = PropBag.ReadProperty("FontColor", m_def_FontColor)
    m_ShadowColor = PropBag.ReadProperty("ShadowColor", m_def_ShadowColor)
    m_ShadowSize = PropBag.ReadProperty("ShadowSize", m_def_ShadowSize)
    m_ShadowDirection = PropBag.ReadProperty("ShadowDirection", m_def_ShadowDirection)
    m_ScrollDirection = PropBag.ReadProperty("ScrollDirection", m_def_ScrollDirection)
    Timer1.Interval = PropBag.ReadProperty("ScrollSpeed", 0)
    m_ScrollRepeatCount = PropBag.ReadProperty("ScrollRepeatCount", m_def_ScrollRepeatCount)
    m_TextAlign = PropBag.ReadProperty("TextAlign", m_def_TextAlign)
    m_WordWrap = PropBag.ReadProperty("WordWrap", m_def_WordWrap)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", PicText.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    Call PropBag.WriteProperty("FontColor", m_FontColor, m_def_FontColor)
    Call PropBag.WriteProperty("ShadowColor", m_ShadowColor, m_def_ShadowColor)
    Call PropBag.WriteProperty("ShadowSize", m_ShadowSize, m_def_ShadowSize)
    Call PropBag.WriteProperty("ShadowDirection", m_ShadowDirection, m_def_ShadowDirection)
    Call PropBag.WriteProperty("ScrollDirection", m_ScrollDirection, m_def_ScrollDirection)
    Call PropBag.WriteProperty("ScrollSpeed", Timer1.Interval, 0)
    Call PropBag.WriteProperty("ScrollRepeatCount", m_ScrollRepeatCount, m_def_ScrollRepeatCount)
    Call PropBag.WriteProperty("TextAlign", m_TextAlign, m_def_TextAlign)
    Call PropBag.WriteProperty("WordWrap", m_WordWrap, m_def_WordWrap)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,""
Public Property Get Text() As String
Attribute Text.VB_Description = "Sets/Returns the text to be scrolled. PIPE ""|"" is used for the Line Break when scrolling vertical.  PIPE is ignored when scrolling horizontal"
    Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As String)
    m_Text = New_Text
    Call DrawText
    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H0
Public Property Get FontColor() As OLE_COLOR
Attribute FontColor.VB_Description = "Sets/Return the Scrolling Text Font Color"
    FontColor = m_FontColor
End Property

Public Property Let FontColor(ByVal New_FontColor As OLE_COLOR)
    m_FontColor = New_FontColor
    Call DrawText
    PropertyChanged "FontColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H0
Public Property Get ShadowColor() As OLE_COLOR
Attribute ShadowColor.VB_Description = "Sets/Returns the Drop Shadow color of the scrolling text"
    ShadowColor = m_ShadowColor
End Property

Public Property Let ShadowColor(ByVal New_ShadowColor As OLE_COLOR)
    m_ShadowColor = New_ShadowColor
    Call DrawText
    PropertyChanged "ShadowColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,3
Public Property Get ShadowSize() As Integer
Attribute ShadowSize.VB_Description = "Sets/Returns the size(depth) of the Drop Shadow for the Scrolling Text"
    ShadowSize = m_ShadowSize
End Property

Public Property Let ShadowSize(ByVal New_ShadowSize As Integer)
    m_ShadowSize = New_ShadowSize
    Call DrawText
    PropertyChanged "ShadowSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get ShadowDirection() As slpScrollText_DropShadowDirection
Attribute ShadowDirection.VB_Description = "Sets/Returns the direction of the Drop Shadow of the Scrolling Text"
    ShadowDirection = m_ShadowDirection
End Property

Public Property Let ShadowDirection(ByVal New_ShadowDirection As slpScrollText_DropShadowDirection)
    m_ShadowDirection = New_ShadowDirection
    Call DrawText
    PropertyChanged "ShadowDirection"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get ScrollDirection() As slpScrollText_ScrollDirection
Attribute ScrollDirection.VB_Description = "Determins the direction of the text scroll"
    ScrollDirection = m_ScrollDirection
End Property

Public Property Let ScrollDirection(ByVal New_ScrollDirection As slpScrollText_ScrollDirection)
    m_ScrollDirection = New_ScrollDirection
    Call DrawText
    PropertyChanged "ScrollDirection"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Timer1,Timer1,-1,Interval
Public Property Get ScrollSpeed() As Long
Attribute ScrollSpeed.VB_Description = "Returns/sets the number of milliseconds between calls to a Timer control's Timer event."
    ScrollSpeed = Timer1.Interval
End Property

Public Property Let ScrollSpeed(ByVal New_ScrollSpeed As Long)
    Timer1.Interval() = New_ScrollSpeed
    PropertyChanged "ScrollSpeed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function StartScroll() As Boolean
    Call DrawText
    Timer1.Enabled = True
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function StopScroll() As Boolean
    Timer1.Enabled = False
    Call DrawText
    mVarScrollCounter = 0
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function LoadFromFile(strFilePath As String) As Boolean
Attribute LoadFromFile.VB_Description = "Loads the Text Property from a text file"
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
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,1
Public Property Get ScrollRepeatCount() As Single
Attribute ScrollRepeatCount.VB_Description = "Sets/Returns the number of times the scrolling should repeat.  Set to 0 for continious scroll"
    ScrollRepeatCount = m_ScrollRepeatCount
End Property

Public Property Let ScrollRepeatCount(ByVal New_ScrollRepeatCount As Single)
    m_ScrollRepeatCount = New_ScrollRepeatCount
    PropertyChanged "ScrollRepeatCount"
End Property

Private Function DrawText() As Boolean
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
            Case slpVertical
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
            Case slpVertical
                'Now position based on our alignment..
                Select Case m_TextAlign
                    Case slpJustifyDefault, slpJustifyLeft
                        PicText.Left = 1
                    Case slpJustifyRight
                        PicText.Left = UserControl.ScaleWidth - PicText.Width
                    Case slpJustifyCenter
                        If PicText.ScaleWidth > UserControl.ScaleWidth Then
                            PicText.Left = 0
                        Else
                            PicText.Left = (UserControl.ScaleWidth \ 2) - (PicText.Width \ 2)
                        End If
                End Select
            Case slpHorizontal
                'alignment doesn't matter when we are scrolling horizontal..
                If PicText.ScaleHeight > UserControl.ScaleHeight Then
                    ' Must be using a big 'ole font..
                    PicText.Top = 0
                Else
                    ' Plenty of room.. let's center the text in the parent control.
                    PicText.Top = (UserControl.ScaleHeight \ 2) - (PicText.ScaleHeight \ 2)
                End If
        End Select
    End If
    ' Restore the origional timer state
    Timer1.Enabled = bTimerState
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get TextAlign() As slpScrollText_TextJustify
    TextAlign = m_TextAlign
End Property

Public Property Let TextAlign(ByVal New_TextAlign As slpScrollText_TextJustify)
    m_TextAlign = New_TextAlign
    Call DrawText
    PropertyChanged "TextAlign"
End Property

Private Sub PrintText(Text As String)
    Dim lCounter As Long
    Dim x As Single
    Dim y As Single
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
    x = PicText.CurrentX
    y = PicText.CurrentY
    ' Now to deal with the Shadow..
    If m_ShadowSize > 0 Then
        For lCounter = 1 To m_ShadowSize
            PicText.Print Text
            ' Here we work the drop shadow direction
            Select Case m_ShadowDirection
                Case slpDropShadowDefault, slpDropShadowSouthEast
                    x = x - 1
                    y = y - 1
                Case slpDropShadowSouthWest
                    x = x + 1
                    y = y - 1
                Case slpDropShadowNorthEast
                    x = x - 1
                    y = y + 1
                Case slpDropShadowNorthWest
                    x = x + 1
                    y = y + 1
            End Select
            PicText.CurrentX = x
            PicText.CurrentY = y
        Next
    End If
    ' Now to finally print our real text
    PicText.ForeColor = m_FontColor
    PicText.Print Text
End Sub

Private Sub LoadTextArray()
    ReDim aText(0)
    If Len(m_Text) > 0 Then
        ' First determine how it is we are going to scroll
        Select Case m_ScrollDirection
            Case slpVertical
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
End Sub

Private Sub ScrollDone()
    If m_ScrollRepeatCount > 0 Then
        mVarScrollCounter = mVarScrollCounter + 1
        If mVarScrollCounter > m_ScrollRepeatCount Then
            Timer1.Enabled = False
            mVarScrollCounter = 0
        End If
    End If
    Call DrawText
    RaiseEvent ScrollFinished
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,1,2,false
Public Property Get IsScrolling() As Boolean
Attribute IsScrolling.VB_MemberFlags = "400"
    IsScrolling = Timer1.Enabled
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,false
Public Property Get WordWrap() As Boolean
    WordWrap = m_WordWrap
End Property

Public Property Let WordWrap(ByVal New_WordWrap As Boolean)
    m_WordWrap = New_WordWrap
    Call DrawText
    PropertyChanged "WordWrap"
End Property

Private Sub DoWordWrap()
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
End Sub
