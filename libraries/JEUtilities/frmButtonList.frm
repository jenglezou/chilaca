VERSION 5.00
Begin VB.Form frmButtonList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dialog Caption"
   ClientHeight    =   2385
   ClientLeft      =   7140
   ClientTop       =   5460
   ClientWidth     =   2985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   2985
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2520
      Top             =   1440
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   2775
   End
   Begin VB.CommandButton Button 
      Caption         =   "Button"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label txtDescription 
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmButtonList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Orientation_High = 0
Const Orientation_Wide = 1

Public sSelectedButton
Public sOrientation

'sOrientation = Orientation_High

Private Sub CancelButton_Click()
    sSelectedButton = ""
    Unload Me
End Sub

Private Sub Form_Load()
    sSelectedButton = ""
    'Me.ScaleMode = 3
    SetAlwaysOnTopMode Me, True
End Sub

Private Sub Button_Click(Index As Integer)
    sSelectedButton = Button(Index).Caption
    Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        MsgBox "RightClicked"
        sOrientation = sOrientation Xor 1
    End If
End Sub

Private Sub Timer1_Timer()
    'FlashWindow Me.hWnd, 1
End Sub

'Private Sub Form_Load()
'End Sub
'
'Private Sub Form_Resize()
'    GlassifyForm Me
'End Sub

