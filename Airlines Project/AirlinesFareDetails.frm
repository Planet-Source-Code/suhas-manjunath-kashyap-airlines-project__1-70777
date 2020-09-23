VERSION 5.00
Begin VB.Form AirlinesFareDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fare Details"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "AirlinesFareDetails.frx":0000
   ScaleHeight     =   7935
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdclose 
      BackColor       =   &H00FFC0FF&
      Caption         =   "&Back"
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton command2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "&Reservation"
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0FF&
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
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   11160
      Top             =   120
   End
   Begin VB.TextBox T2 
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   3720
      Width           =   3495
   End
   Begin VB.TextBox T4 
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   4320
      Width           =   3495
   End
   Begin VB.TextBox T3 
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   2640
      Width           =   3495
   End
   Begin VB.TextBox T1 
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   3240
      Width           =   3495
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   5760
      TabIndex        =   1
      Top             =   1440
      Width           =   3495
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5760
      TabIndex        =   0
      Top             =   2040
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
      Left            =   9120
      TabIndex        =   7
      Top             =   7560
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
      Left            =   240
      TabIndex        =   6
      Top             =   7560
      Width           =   1695
   End
End
Attribute VB_Name = "AirlinesFareDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AMT As String
 
Private Sub Cmdclose_Click()
Unload Me
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
 If (KeyAscii <> 77) And (KeyAscii <> 70) Then
        MsgBox "Select from the given list", vbExclamation, "List"
        KeyAscii = 0
 End If

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)

 If (KeyAscii <> 77) And (KeyAscii <> 70) Then
        MsgBox "Select from the given list", vbExclamation, "List"
        KeyAscii = 0
 End If

End Sub

Private Sub Command1_Click()
Set rs = conn.Execute("select * from FARE where flight_no='" + Combo1 + "' and class='" + Combo2 + "'")
If Combo1 = "" And Combo2 = "" Then
MsgBox "Please select the class and flight number", vbOKOnly, "fare"
Exit Sub
End If
If Combo1 = "" Then
MsgBox "Please select the flight number", vbOKOnly, "fare"
Exit Sub
End If
If Combo2 = "" Then
MsgBox "Please select the class", vbOKOnly, "fare"
Exit Sub
End If
T1 = rs.Fields(1)
T3 = rs.Fields(3)
T2 = rs.Fields(4)
T4 = rs.Fields(5)
End Sub

Private Sub Command2_Click()
SeatReservation.Show
AirlinesFareDetails.Hide
End Sub

Private Sub Form_Activate()
Call clear
Set rs.ActiveConnection = conn
Set rs = conn.Execute("select flight_no from flight")
Do While rs.EOF = False
Combo1.AddItem rs.Fields(0)
rs.MoveNext
Loop
rs.Close
End Sub

Private Sub Form_Load()
Label2.Caption = Format(Date, "DD-MMMM-YYYY")
Label3.Caption = Time$
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\DVD Writting\DVD_2\VB Projects\RDBMS Project\Microsoft Visual Basic Projects\Airlines Project\Database\airline_db.mdb;Persist Security Info=False"
conn.Open
 Combo2.AddItem ("Executive")
 Combo2.AddItem ("Economy")
 Combo2.AddItem ("Business")
End Sub

Sub clear()
Combo1 = ""
Combo2 = ""
T1 = ""
T2 = ""
T3 = ""
T4 = ""
End Sub



Private Sub T4_Change()
AMT = T2 * 0.1
T4 = AMT
End Sub

Private Sub Timer1_Timer()
If flag = True Then
rs.Close
flag = False
End If
End Sub

