VERSION 5.00
Begin VB.Form frmItemList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Item"
   ClientHeight    =   3645
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   2925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   2925
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   2655
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2655
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
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
      TabIndex        =   4
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmItemList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sSelectedItem
Public vbButtonPressed

Private Sub CancelButton_Click()
    sSelectedItem = ""
    vbButtonPressed = vbCancel
    Unload Me
End Sub

Private Sub Form_Load()
    sSelectedItem = ""
    vbButtonPressed = -1
    SetAlwaysOnTopMode Me, True
End Sub

Private Sub List1_Click()
    Text1.Text = List1.Text
End Sub

Private Sub OKButton_Click()
    sSelectedItem = Text1.Text
    vbButtonPressed = vbOK
    Unload Me
End Sub
