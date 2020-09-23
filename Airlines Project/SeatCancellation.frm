VERSION 5.00
Begin VB.Form SeatCancellation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancellation Details"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "SeatCancellation.frx":0000
   ScaleHeight     =   7950
   ScaleWidth      =   12000
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Cancel"
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton cmdclose 
      BackColor       =   &H00C0E0FF&
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "V&iew"
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7320
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5760
      TabIndex        =   9
      Top             =   6720
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   5760
      TabIndex        =   8
      Top             =   6120
      Width           =   3375
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   5520
      Width           =   3375
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   5760
      TabIndex        =   6
      Top             =   4920
      Width           =   3375
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   4320
      Width           =   3375
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   3720
      Width           =   3375
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   2640
      Width           =   3375
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   2040
      Width           =   3375
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5760
      TabIndex        =   0
      Top             =   1440
      Width           =   3375
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
      Left            =   9000
      TabIndex        =   14
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
      Left            =   120
      TabIndex        =   13
      Top             =   7560
      Width           =   1695
   End
End
Attribute VB_Name = "SeatCancellation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AMT As String
Dim ap As String
Dim ad As String
Dim ba As String
Dim i As Integer

Private Sub Cmdclose_Click()
Unload Me
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
 If (KeyAscii <> 77) And (KeyAscii <> 70) Then
        MsgBox "Select from the given list", vbExclamation, "List"
        KeyAscii = 0
    End If
End Sub

Private Sub Command1_Click()
Set rs = conn.Execute("select * from reserve_passenger where ticket_no='" + Combo1 + "'")
If Combo1 = "" Then
MsgBox "please specify the ticket number", vbOKOnly, "cancel"
Exit Sub
End If
Do While rs.EOF = False
If Combo1 = rs.Fields(8) Then
Text1.Text = rs.Fields(10)
Text3.Text = rs.Fields(9)
Text5.Text = rs.Fields(0)
Text10 = Format(Date, "dd-mmm-yyyy")
Text8.Text = Format(rs.Fields(15), "dd-mmm-yyyy")
Text11.Text = Format(rs.Fields(14), "dd-mmm-yyyy")
Text7.Text = rs.Fields(16)
If Text7 = "" Then
MsgBox " ENTER THE AMOUNT"
Else
AMT = Text7 * 0.1
Text6 = AMT
End If
ap = Text7
ad = Text6
Text4 = ap - ad
rs.MoveNext
Exit Sub
End If
Loop
MsgBox "This ticket number does not exist", vbOKOnly, "cancel"
Call clear
Exit Sub
End Sub

Private Sub Command2_Click()
i = MsgBox("Do you really want to delete the record ", vbExclamation + TxtProdid + vbYesNo, "DELETE REQUEST")
If i = 6 Then
Set rs = conn.Execute(" DELETE FROM reserve_passenger WHERE ticket_no='" + Combo1 + "'")
Unload Me
MsgBox "Your Seat is Successfully Cancelled"
AirlinesPanel.Show
End If
End Sub

Private Sub Form_Activate()
Set rs.ActiveConnection = conn
Set rs = conn.Execute("select TICKET_NO from reserve_passenger")
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

Private Sub Text10_GotFocus()

Text6.SetFocus
End Sub

Sub clear()
Combo1.Text = ""
Text1.Text = ""
Text3.Text = ""
Text5.Text = ""
Text8.Text = ""
Text11.Text = ""
Text7.Text = ""
Text10.Text = ""
Text6.Text = ""
Text4.Text = ""
End Sub

Private Sub Timer1_Timer()
If flag = True Then
rs.Close
flag = False
End If
End Sub

