VERSION 5.00
Begin VB.Form AirlinesFlightDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Airlines Flight Details"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "AirlinesFlightDetails.frx":0000
   ScaleHeight     =   7920
   ScaleWidth      =   12000
   Begin VB.Timer Timer1 
      Left            =   10920
      Top             =   240
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0C0C0&
      Caption         =   "C&lose"
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&View"
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox t4 
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2640
      Width           =   3435
   End
   Begin VB.TextBox t6 
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   4920
      Width           =   3435
   End
   Begin VB.TextBox t5 
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   4320
      Width           =   3435
   End
   Begin VB.TextBox t2 
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2040
      Width           =   3435
   End
   Begin VB.TextBox t3 
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3240
      Width           =   3435
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5760
      TabIndex        =   1
      Top             =   1440
      Width           =   3495
   End
   Begin VB.TextBox t1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   3720
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
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
      Height          =   375
      Left            =   9840
      TabIndex        =   10
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
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
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   7440
      Width           =   1695
   End
End
Attribute VB_Name = "AirlinesFlightDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmdclose_Click()
Unload Me
AirlinesPanel.Show
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
 If (KeyAscii <> 77) And (KeyAscii <> 70) Then
        MsgBox "Select from the given list", vbExclamation, "List"
        KeyAscii = 0
    End If
End Sub

Private Sub Command1_Click()
Set rs = conn.Execute("select * from flight where route='" + Combo1 + "'")
If Combo1 = "" Then
MsgBox "Please select the route", vbOKOnly, "flight"
Exit Sub
End If
Do While rs.EOF = False
If Combo1 = rs.Fields(2) Then
T1 = rs.Fields(0)
T2 = rs.Fields(1)
T3 = rs.Fields(3)
T4 = rs.Fields(4)
T5 = rs.Fields(5)
T6 = rs.Fields(6)
Exit Sub
End If
Loop
MsgBox "This route does not exist", vbOKOnly, "flight"
Call clear
Combo1.SetFocus
Exit Sub
End Sub

Private Sub Form_Activate()
Call clear
Set rs.ActiveConnection = conn
Set rs = conn.Execute("select route from flight")
Do While rs.EOF = False
Combo1.AddItem rs.Fields(0)
rs.MoveNext
Loop
rs.Close
End Sub



Private Sub Form_Load()
Label2.Caption = Format(Date, "dd-mmm-yy")
Label3.Caption = Time$
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\DVD Writting\DVD_2\VB Projects\RDBMS Project\Microsoft Visual Basic Projects\Airlines Project\Database\airline_db.mdb;Persist Security Info=False"
conn.Open
End Sub

Sub clear()
Combo1.Text = ""
T1.Text = ""
T2.Text = ""
T3.Text = ""
T4.Text = ""
T6.Text = ""
End Sub

Private Sub Timer1_Timer()
If flag = True Then
rs.Close
flag = False
End If
End Sub

