VERSION 5.00
Begin VB.Form AirportDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Airport Details"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "AirportDetails.frx":0000
   ScaleHeight     =   7980
   ScaleWidth      =   11985
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   7560
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   4440
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   3720
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   3120
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   2640
      Width           =   3375
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label5 
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
      Left            =   9720
      TabIndex        =   6
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Label Label4 
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
      Left            =   7920
      TabIndex        =   5
      Top             =   7560
      Width           =   1695
   End
End
Attribute VB_Name = "AirportDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
AirlinesPanel.Show
End Sub

Private Sub Form_Activate()
Set rs = conn.Execute("select * from airport")
Text1 = rs.Fields(0)
Text2 = rs.Fields(1)
Text3 = rs.Fields(2)
Text4 = rs.Fields(3)
End Sub

Private Sub Form_Load()
Label4.Caption = Format(Date, "DD-MMMM-YYYY")
Label5.Caption = Time$
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\DVD Writting\DVD_2\VB Projects\RDBMS Project\Microsoft Visual Basic Projects\Airlines Project\Database\airline_db.mdb;Persist Security Info=False"
conn.Open
End Sub


Private Sub Timer1_Timer()
If flag = True Then
rs.Close
flag = False
End If
End Sub
