VERSION 5.00
Begin VB.Form AirlinesAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About the Project"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "AirlinesAbout.frx":0000
   ScaleHeight     =   7905
   ScaleWidth      =   12000
   Begin VB.Label suhasweb 
      BackStyle       =   0  'Transparent
      Caption         =   "Click Here to See More Projects"
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   7320
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   765
      Left            =   4560
      Picture         =   "AirlinesAbout.frx":1A06C4
      Top             =   7080
      Width           =   750
   End
   Begin VB.Label suhas 
      BackStyle       =   0  'Transparent
      Caption         =   "Suhas Manjunath"
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   9720
      TabIndex        =   0
      Top             =   7560
      Width           =   2415
   End
End
Attribute VB_Name = "AirlinesAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Dim verify As Boolean
verify = ChkAuthor(CStr(suhas.Caption))
If verify <> False Then
    MsgBox ("This project cannot be used by Illega Persons")
    End
Else
    AirlinesAbout.Show
End If
End Sub

Private Sub Image1_Click()
Dim strlink As String
strlink = "http://www.suhasmanjunath.co.nr"
Shell "C:/program files/Mozilla Firefox/firefox.exe " & strlink
End Sub

Private Sub suhas_Click()

End Sub
