VERSION 5.00
Begin VB.Form AirlinesLogin 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "AirlinesLogin.frx":0000
   ScaleHeight     =   8940
   ScaleWidth      =   11955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H0080C0FF&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0080C0FF&
      Caption         =   "L&ogin"
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox txtpwd 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   5280
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4680
      Width           =   2655
   End
   Begin VB.TextBox txtid 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   5280
      TabIndex        =   0
      Top             =   4080
      Width           =   2655
   End
End
Attribute VB_Name = "AirlinesLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LoginSucceeded As Boolean
Public MsgString As String
Public MsgString1 As String
Public MsgString2 As String

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
    MsgBox "Thank you for using this Software", vbInformation
    End
End Sub

Private Sub cmdOK_Click()
    'check for correct password
    If txtid = "airlines" Or txtpwd = "bangalore" Then
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
        LoginSucceeded = True
        Me.Hide
        AirlinesPanel.Show
    Else
        MsgBox "Invalid , try again!", , "Login"
        txtpwd.SetFocus
        SendKeys "{Home}+{End}"
    End If
If txtpwd = "" Then
MsgBox "Enter password", , "password"
AirlinesLogin.Show
End If
End Sub

Private Sub cmdOK_KeyPress(KeyAscii As Integer)
If (KeyAscii < 9) And Not (KeyAscii = 46) Then
MsgBox "Enter correct id", , "ID"
End If
End Sub
