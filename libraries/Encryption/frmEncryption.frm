VERSION 5.00
Begin VB.Form frmEncryption 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encrypt String"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   840
      Top             =   1920
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   8280
      TabIndex        =   8
      ToolTipText     =   "Click here to clear the Decrypted field"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtDecrypted 
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      ToolTipText     =   "Privacy Warning. If you click Decrypt then the decrypted string will appear here."
      Top             =   1440
      Width           =   6855
   End
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "Decrypt"
      Height          =   375
      Left            =   8280
      TabIndex        =   5
      ToolTipText     =   "Privacy Warning. If you click Decrypt then the decrypted string will appear below but it will disappear after 1 second."
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtEncrypted 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      ToolTipText     =   "Type your string here and press Decrypt"
      Top             =   840
      Width           =   6855
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "Encrypt"
      Height          =   375
      Left            =   8280
      TabIndex        =   2
      ToolTipText     =   "Click here to encrypt the Source text.  The encrypted string will appear in the Encrypted field."
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox txtSource 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Type your string here and press Encrypt"
      Top             =   240
      Width           =   6855
   End
   Begin VB.Label Label3 
      Caption         =   "Decrypted"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Encrypted"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Source"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmEncryption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const sKey = "encryption"

Private Sub cmdClear_Click()
    txtDecrypted = ""
End Sub

Private Sub cmdDecrypt_Click()
    txtDecrypted = Decrypt(txtEncrypted, sKey)
End Sub

Private Sub cmdEncrypt_Click()
    txtEncrypted = Encrypt(txtSource, sKey)
End Sub

Function Encrypt(sStr, sKey)
    Dim iLenKey As Integer
    Dim iKeyPos As Integer
    Dim iLenStr As Integer
    Dim i As Integer
    Dim sNewstr As String
    
    sNewstr = ""
    iLenKey = Len(sKey)
    iKeyPos = 1
    iLenStr = Len(sStr)
    sStr = StrReverse(sStr)
    For i = 1 To iLenStr
        sNewstr = sNewstr & Chr(Asc(Mid(sStr, i, 1)) + Asc(Mid(sKey, iKeyPos, 1)))
        iKeyPos = iKeyPos + 1
        If iKeyPos > iLenKey Then iKeyPos = 1
    Next
    
    Encrypt = sNewstr
End Function

Function Decrypt(sStr, sKey)
    Dim iLenKey As Integer
    Dim iKeyPos As Integer
    Dim iLenStr As Integer
    Dim i As Integer
    Dim sNewstr As String

    sNewstr = ""
    iLenKey = Len(sKey)
    iKeyPos = 1
    iLenStr = Len(sStr)

    sStr = StrReverse(sStr)
    For i = iLenStr To 1 Step -1
      sNewstr = sNewstr & Chr(Asc(Mid(sStr, i, 1)) - Asc(Mid(sKey, iKeyPos, 1)))
      iKeyPos = iKeyPos + 1
      If iKeyPos > iLenKey Then iKeyPos = 1
    Next

    sNewstr = StrReverse(sNewstr)
    Decrypt = sNewstr
End Function


Private Sub Timer1_Timer()
    txtDecrypted = ""
End Sub
