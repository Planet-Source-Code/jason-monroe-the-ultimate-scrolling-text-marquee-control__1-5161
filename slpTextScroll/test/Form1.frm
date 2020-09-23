VERSION 5.00
Object = "{3DF22404-BBF9-11D3-9FD9-0050DA088718}#1.0#0"; "SLPTools.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3000
   ClientLeft      =   4170
   ClientTop       =   2715
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   5190
   Begin SLPTxtScrll.slpTextScroll slpTextScroll1 
      Height          =   1875
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3307
      BackColor       =   8454143
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   1
      BorderStyle     =   1
      Text            =   "Hello World!||Welcome to the slpTextScroll Demo!||IHope that you enjoy this control"
      FontColor       =   255
      ShadowColor     =   8421504
      ShadowSize      =   3
      ShadowDirection =   4
      ScrollDirection =   1
      ScrollSpeed     =   10
      ScrollRepeatCount=   1
      TextAlign       =   1
      WordWrap        =   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click me!"
      Height          =   495
      Left            =   1860
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If slpTextScroll1.IsScrolling Then
        slpTextScroll1.StopScroll
    Else
        slpTextScroll1.StartScroll
    End If
End Sub

Private Sub Form_Load()
    Me.Move (Screen.Width \ 2) - (Me.Width \ 2), (Screen.Height \ 2) - (Me.Height \ 2)
End Sub

Private Sub slpTextScroll1_DblClick()
    MsgBox "double cool"

End Sub
