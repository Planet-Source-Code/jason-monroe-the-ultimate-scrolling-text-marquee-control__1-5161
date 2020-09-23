Attribute VB_Name = "ColorPickerCode"

Public Type CHOOSECOLORS
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Type RECT
    left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type

Public Type SelectedColor
    oSelectedColor As OLE_COLOR
    bCanceled As Boolean
End Type

Public Const CC_RGBINIT = &H1
Public Const CC_FULLOPEN = &H2
Public Const CC_PREVENTFULLOPEN = &H4
Public Const CC_SHOWHELP = &H8
Public Const CC_ENABLEHOOK = &H10
Public Const CC_ENABLETEMPLATE = &H20
Public Const CC_ENABLETEMPLATEHANDLE = &H40
Public Const CC_SOLIDCOLOR = &H80
Public Const CC_ANYCOLOR = &H100
Public Const COLOR_FLAGS = CC_FULLOPEN Or CC_ANYCOLOR Or CC_RGBINIT
Public Const GWL_HINSTANCE = (-6)
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOACTIVATE = &H10
Public Const HCBT_ACTIVATE = 5
Public Const WH_CBT = 5

Public Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLORS) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetCurrentThreadId Lib "KERNEL32" () As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private hHook As Long


Public Function ShowColor(ByVal hWnd As Long) As SelectedColor
    Dim ColorDialog As CHOOSECOLORS
    Dim customcolors() As Byte  ' dynamic (resizable) array
    Dim i As Integer
    Dim ret As Long
    Dim hInst As Long
    Dim Thread As Long
    
    'ParenthWnd = hWnd
    If ColorDialog.lpCustColors = "" Then
        ReDim customcolors(0 To 16 * 4 - 1) As Byte  'resize the array
    
        For i = LBound(customcolors) To UBound(customcolors)
          customcolors(i) = 254 ' sets all custom colors to white
        Next i
        
        ColorDialog.lpCustColors = StrConv(customcolors, vbUnicode)  ' convert array
    End If
    
    ColorDialog.hwndOwner = hWnd
    ColorDialog.lStructSize = Len(ColorDialog)
    ColorDialog.flags = COLOR_FLAGS
    
    'Set up the CBT hook
    hInst = GetWindowLong(hWnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcAlignForm, hInst, Thread)
    
    ret = ChooseColor(ColorDialog)
    If ret Then
        ShowColor.bCanceled = False
        ShowColor.oSelectedColor = ColorDialog.rgbResult
        Exit Function
    Else
        ShowColor.bCanceled = True
        ShowColor.oSelectedColor = &H0&
        Exit Function
    End If
End Function

Public Function WinProcAlignForm(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
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

