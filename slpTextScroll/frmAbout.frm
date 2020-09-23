VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2010
   ClientLeft      =   6225
   ClientTop       =   6165
   ClientWidth     =   4800
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4800
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   600
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   450
      TabIndex        =   2
      Top             =   60
      Width           =   450
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1230
      Left            =   60
      Picture         =   "frmAbout.frx":0842
      ScaleHeight     =   1170
      ScaleWidth      =   945
      TabIndex        =   1
      Top             =   720
      Width           =   1005
   End
   Begin SLPTxtScrll.slpTextScroll slpTextScroll1 
      Height          =   1875
      Left            =   1140
      TabIndex        =   0
      Top             =   60
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   3307
      BackColor       =   12648447
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   1
      BorderStyle     =   1
      Text            =   ""
      FontColor       =   255
      ShadowColor     =   8421504
      ShadowSize      =   3
      ShadowDirection =   4
      ScrollDirection =   1
      ScrollSpeed     =   3
      ScrollRepeatCount=   0
      TextAlign       =   1
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strScrollText As String
    Me.left = (Screen.Width - Me.Width) \ 2
    Me.top = (Screen.Height - Me.Height) \ 2
    Me.Caption = App.Title
    strScrollText = "slpTextScroll|" & _
                    "v" & App.Major & "." & App.Minor & "." & App.Revision & "||" & _
                    "by|" & _
                    "Silverlance|" & _
                    "Productions||" & _
                    App.FileDescription & "||" & _
                    App.Comments
    slpTextScroll1.Text = strScrollText
    slpTextScroll1.StartScroll
End Sub

Private Sub Picture1_Click()
    Unload Me
End Sub

Private Sub Picture2_Click()
    Unload Me
End Sub

Private Sub slpTextScroll1_Click()
    Unload Me
End Sub
