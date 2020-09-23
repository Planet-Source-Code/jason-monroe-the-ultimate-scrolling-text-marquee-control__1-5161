Attribute VB_Name = "ShowFontDialog"

'API's
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONTS) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetCurrentThreadId Lib "KERNEL32" () As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Const LF_FACESIZE = 32
Private Const GWL_HINSTANCE = (-6)
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOACTIVATE = &H10
Private Const HCBT_ACTIVATE = 5
Private Const WH_CBT = 5
Private Const CF_SCREENFONTS = &H1
Private Const CF_PRINTERFONTS = &H2
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_SHOWHELP = &H4&
Private Const CF_ENABLEHOOK = &H8&
Private Const CF_ENABLETEMPLATE = &H10&
Private Const CF_ENABLETEMPLATEHANDLE = &H20&
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_USESTYLE = &H80&
Private Const CF_EFFECTS = &H100&
Private Const CF_APPLY = &H200&
Private Const CF_ANSIONLY = &H400&
Private Const CF_SCRIPTSONLY = CF_ANSIONLY
Private Const CF_NOVECTORFONTS = &H800&
Private Const CF_NOOEMFONTS = CF_NOVECTORFONTS
Private Const CF_NOSIMULATIONS = &H1000&
Private Const CF_LIMITSIZE = &H2000&
Private Const CF_FIXEDPITCHONLY = &H4000&
Private Const CF_WYSIWYG = &H8000 '  must also have CF_SCREENFONTS CF_PRINTERFONTS
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_SCALABLEONLY = &H20000
Private Const CF_TTONLY = &H40000
Private Const CF_NOFACESEL = &H80000
Private Const CF_NOSTYLESEL = &H100000
Private Const CF_NOSIZESEL = &H200000
Private Const CF_SELECTSCRIPT = &H400000
Private Const CF_NOSCRIPTSEL = &H800000
Private Const CF_NOVERTFONTS = &H1000000

Private Const SIMULATED_FONTTYPE = &H8000
Private Const PRINTER_FONTTYPE = &H4000
Private Const SCREEN_FONTTYPE = &H2000
Private Const BOLD_FONTTYPE = &H100
Private Const ITALIC_FONTTYPE = &H200
Private Const REGULAR_FONTTYPE = &H400

'Types
Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Type RECT
    left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type

Type CHOOSEFONTS
    lStructSize As Long
    hwndOwner As Long          '  caller's window handle
    hDC As Long                '  printer DC/IC or NULL
    lpLogFont As Long          '  ptr. to a LOGFONT struct
    iPointSize As Long         '  10 * size in points of selected font
    flags As Long              '  enum. type flags
    rgbColors As Long          '  returned text color
    lCustData As Long          '  data passed to hook fn.
    lpfnHook As Long           '  ptr. to hook function
    lpTemplateName As String     '  custom template name
    hInstance As Long          '  instance handle of.EXE that
    lpszStyle As String          '  return the style field here
    nFontType As Integer          '  same value reported to the EnumFonts
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long           '  minimum pt size allowed &
    nSizeMax As Long           '  max pt size allowed if
End Type

Public Type SelectedFont
    sSelectedFont As String
    bCanceled As Boolean
    bBold As Boolean
    bItalic As Boolean
    nSize As Integer
    bUnderline As Boolean
    bStrikeOut As Boolean
    lColor As Long
    sFaceName As String
End Type

Private hHook As Long
Private FontDialog As CHOOSEFONTS

Public Function ShowFont(ByVal hWnd As Long, ByRef startingFontName As String, ByRef startingFontSize As Long) As SelectedFont
Dim ret As Long
Dim lfLogFont As LOGFONT
Dim hInst As Long
Dim Thread As Long
Dim i As Integer
    
    ParenthWnd = hWnd
    FontDialog.nSizeMax = 0
    FontDialog.nSizeMin = 0
    FontDialog.nFontType = Screen.FontCount
    FontDialog.hwndOwner = hWnd
    FontDialog.hDC = 0
    FontDialog.lpfnHook = 0
    FontDialog.lCustData = 0
    FontDialog.lpLogFont = VarPtr(lfLogFont)
    FontDialog.iPointSize = startingFontSize * 10
    FontDialog.lpTemplateName = Space$(2048)
    FontDialog.rgbColors = RGB(0, 255, 255)
    FontDialog.lStructSize = Len(FontDialog)
    
    If FontDialog.flags = 0 Then
        'FontDialog.flags = CF_SCREENFONTS Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT 'Or CF_EFFECTS
        FontDialog.flags = CF_SCREENFONTS Or CF_NOSCRIPTSEL Or CF_INITTOLOGFONTSTRUCT
    End If
    
    For i = 0 To Len(startingFontName) - 1
        lfLogFont.lfFaceName(i) = Asc(Mid(startingFontName, i + 1, 1))
    Next
    
    'Set up the CBT hook
    hInst = GetWindowLong(hWnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    
    hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcAlignForm, hInst, Thread)
    
    ret = ChooseFont(FontDialog)
        
    If ret Then
        ShowFont.bCanceled = False
        ShowFont.bBold = IIf(lfLogFont.lfWeight > 400, 1, 0)
        ShowFont.bItalic = lfLogFont.lfItalic
        ShowFont.bStrikeOut = lfLogFont.lfStrikeOut
        ShowFont.bUnderline = lfLogFont.lfUnderline
        ShowFont.lColor = FontDialog.rgbColors
        ShowFont.nSize = FontDialog.iPointSize / 10
        For i = 0 To 31
            ShowFont.sSelectedFont = ShowFont.sSelectedFont + Chr(lfLogFont.lfFaceName(i))
        Next
    
        ShowFont.sSelectedFont = Mid(ShowFont.sSelectedFont, 1, InStr(1, ShowFont.sSelectedFont, Chr(0)) - 1)
        Exit Function
    Else
        ShowFont.bCanceled = True
        Exit Function
    End If
End Function

Private Function WinProcAlignForm(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim rectForm As RECT, rectMsg As RECT
    Dim x As Long, y As Long
    'On HCBT_ACTIVATE, show the MsgBox centered over Form1
    If lMsg = HCBT_ACTIVATE Then
        'Get the coordinates of the form and the message box so that
        'you can determine where the center of the form is located
        'GetWindowRect ParenthWnd, rectForm
        GetWindowRect wParam, rectMsg
        x = rectMsg.left + 10
        y = rectMsg.top + 10
        'Position the msgbox
        SetWindowPos wParam, 0, x, y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
        'Release the CBT hook
        UnhookWindowsHookEx hHook
     End If
End Function

