VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPopUp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dialog Caption"
   ClientHeight    =   1590
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   120
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Left            =   4680
      Top             =   1080
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bShow
Private iCurrentProgressBar

Private Sub Form_Load()
    'iCurrentProgressBar = ProgressBar1.Max
    iCurrentProgressBar = 0
    bShow = True
    Me.ScaleMode = 3
    SetAlwaysOnTopMode Me, True
End Sub

Private Sub OKButton_Click()
    bShow = False
    Timer1.Interval = 0
    Unload Me
End Sub

Private Sub Timer1_Timer()
    bShow = False
    Unload Me
End Sub

Private Sub Timer2_Timer()
    Label2.Visible = False
    Label2 = Timer1.Interval
    If Timer1.Interval = 0 Then
        'Use 3 minutes by default
        'iCurrentProgressBar = Abs(iCurrentProgressBar - iCurrentProgressBar / 180)
        iCurrentProgressBar = (iCurrentProgressBar + ProgressBar1.Max / (Timer2.Interval + 1)) Mod ProgressBar1.Max
        Label2 = iCurrentProgressBar
    Else
        'iCurrentProgressBar = Abs(iCurrentProgressBar - Timer2.Interval * ProgressBar1.Max / Timer1.Interval)
        iCurrentProgressBar = iCurrentProgressBar + Timer2.Interval * ProgressBar1.Max / Timer1.Interval
        Label2 = iCurrentProgressBar
    End If
    ProgressBar1.Value = iCurrentProgressBar
End Sub

Private Sub Form_Resize()
    'GlassifyForm Me
End Sub

