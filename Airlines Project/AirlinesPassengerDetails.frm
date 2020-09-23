VERSION 5.00
Begin VB.Form AirlinesPassengerDetails 
   Caption         =   "Airlines Passenger Details"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11970
   BeginProperty Font 
      Name            =   "Trajan Pro"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "AirlinesPassengerDetails.frx":0000
   ScaleHeight     =   7935
   ScaleWidth      =   11970
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Back"
      Height          =   615
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "V&iew"
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Left            =   10920
      Top             =   240
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3240
      Width           =   3495
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3720
      Width           =   3495
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   4320
      Width           =   3495
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2640
      Width           =   3495
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   4800
      Width           =   3495
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   5400
      Width           =   3495
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   5880
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2040
      Width           =   3495
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   5760
      TabIndex        =   0
      Top             =   1440
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
      Left            =   9000
      TabIndex        =   10
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
      TabIndex        =   9
      Top             =   7560
      Width           =   1695
   End
End
Attribute VB_Name = "AirlinesPassengerDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim uni As Integer

Private Sub Combo1_KeyPress(KeyAscii As Integer)
 If (KeyAscii <> 77) And (KeyAscii <> 70) Then
        MsgBox "Select from the given list", vbExclamation, "List"
        KeyAscii = 0
    End If
End Sub

Private Sub Command1_Click()
Set rs = conn.Execute("select * from reserve_passenger where passport_no='" & Combo1 & "'")
If Combo1 = "" Then
MsgBox "Please select the passport number", vbOKOnly, "reserve_passenger"
Exit Sub
End If
Do While rs.EOF = False
If Combo1 = rs.Fields(0) Then
Text1.Text = rs.Fields(11)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(8)
Text4.Text = rs.Fields(9)
Text5.Text = rs.Fields(10)
Text6.Text = rs.Fields(13)
Text7.Text = Format(rs.Fields(14), "dd-mmm-yy")
Text8.Text = Format(rs.Fields(15), "dd-mmm-yy")
rs.MoveNext
Exit Sub
End If
Loop
MsgBox "This passenger does not exist", vbOKOnly, "reserve_passenger"
Call clear
Exit Sub
End Sub

Private Sub Command2_Click()
Unload Me
AirlinesPanel.Show
End Sub

Private Sub Form_Activate()
Call clear
Set rs.ActiveConnection = conn
Set rs = conn.Execute("select passport_no from reserve_passenger")
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
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
End Sub

Private Sub Timer1_Timer()
If flag = True Then
rs.Close
flag = False
End If
End Sub

