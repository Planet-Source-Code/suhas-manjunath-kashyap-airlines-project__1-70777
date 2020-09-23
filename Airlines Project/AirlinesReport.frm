VERSION 5.00
Begin VB.Form AirlinesReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Airlines Report"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "AirlinesReport.frx":0000
   ScaleHeight     =   7950
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3120
      TabIndex        =   5
      Top             =   5400
      Width           =   5415
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
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
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
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
      Height          =   405
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3120
      TabIndex        =   2
      Top             =   3840
      Width           =   5415
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   3120
      TabIndex        =   1
      Top             =   2160
      Width           =   5415
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF8080&
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
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
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
      Left            =   9360
      TabIndex        =   7
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
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
      Left            =   600
      TabIndex        =   6
      Top             =   7080
      Width           =   2055
   End
End
Attribute VB_Name = "AirlinesReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dflag As Boolean
Dim flag As Boolean

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

Private Sub Combo3_KeyPress(KeyAscii As Integer)
 If (KeyAscii <> 77) And (KeyAscii <> 70) Then
 MsgBox "Select from the given list", vbExclamation, "List"
        KeyAscii = 0
    End If
End Sub

Private Sub Command1_Click()
flag = True
If (dflag = True) Then
 DataEnvironment1.Connection1.Close
 dflag = False
 End If
DataEnvironment1.Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\DVD Writting\DVD_2\VB Projects\RDBMS Project\Microsoft Visual Basic Projects\Airlines Project\Database\airline_db.mdb;Persist Security Info=False"
DataEnvironment1.Connection1.Open
DataEnvironment1.Command3 Combo1
If Combo1.Text = "" Then
MsgBox "Please select the fare code", vbOKOnly, "faresreport"
dflag = True
Exit Sub
End If
DataReport1.Show
dflag = True
Unload Me
End Sub

Private Sub Command2_Click()
flag = True
If (dflag = True) Then
 DataEnvironment1.Connection1.Close
 dflag = False
 End If
DataEnvironment1.Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\DVD Writting\DVD_2\VB Projects\RDBMS Project\Microsoft Visual Basic Projects\Airlines Project\Database\airline_db.mdb;Persist Security Info=False"
DataEnvironment1.Connection1.Open
DataEnvironment1.Command4 Combo2
If Combo2.Text = "" Then
MsgBox "Please select the flight number code", vbOKOnly, "faresreport"
dflag = True
Exit Sub
End If
DataReport3.Show
dflag = True
Unload Me
End Sub

Private Sub Command3_Click()
flag = True
If (dflag = True) Then
 DataEnvironment1.Connection1.Close
 dflag = False
 End If
DataEnvironment1.Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\DVD Writting\DVD_2\VB Projects\RDBMS Project\Microsoft Visual Basic Projects\Airlines Project\Database\airline_db.mdb;Persist Security Info=False"
DataEnvironment1.Connection1.Open
DataEnvironment1.Command2 Combo3
If Combo3.Text = "" Then
MsgBox "Please select the passport number code", vbOKOnly, "faresreport"
dflag = True
Exit Sub
End If
DataReport2.Show
dflag = True
Unload Me
End Sub

Private Sub Command4_Click()
flag = True
If (dflag = True) Then
 DataEnvironment1.Connection1.Close
 dflag = False
 End If
DataEnvironment1.Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\DVD Writting\DVD_2\VB Projects\RDBMS Project\Microsoft Visual Basic Projects\Airlines Project\Database\airline_db.mdb;Persist Security Info=False"
DataEnvironment1.Connection1.Open
DataEnvironment1.Command5 Combo4
If Combo4.Text = "" Then
MsgBox "Please select the Ticket number code", vbOKOnly, "faresreport"
dflag = True
Exit Sub
End If
DataReport5.Show
dflag = True
Unload Me
End Sub

Private Sub Command5_Click()
flag = True
If (dflag = True) Then
 DataEnvironment1.Connection1.Close
 dflag = False
 End If
DataEnvironment1.Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\DVD Writting\DVD_2\VB Projects\RDBMS Project\Microsoft Visual Basic Projects\Airlines Project\Database\airline_db.mdb;Persist Security Info=False"
DataEnvironment1.Connection1.Open
DataEnvironment1.Command1
AirReport.Show
dflag = True
Unload Me
End Sub

Private Sub Timer1_Timer()
If flag = True Then
DataEnvironment1.Connection1.Open
End Sub

Private Sub Form_Activate()
Set rs.ActiveConnection = conn
Set rs = conn.Execute("select fare_code from fare")
Do While rs.EOF = False
Combo1.AddItem rs.Fields(0)
rs.MoveNext
Loop
rs.Close
Set rs1.ActiveConnection = conn
Set rs1 = conn.Execute("select flight_no from flight")
Do While rs1.EOF = False
Combo2.AddItem rs1.Fields(0)
rs1.MoveNext
Loop
rs1.Close
Set rs2.ActiveConnection = conn
Set rs2 = conn.Execute("select passport_no from reserve_passenger")
Do While rs2.EOF = False
Combo3.AddItem rs2.Fields(0)
rs2.MoveNext
Loop
rs2.Close
Set rs3.ActiveConnection = conn
Set rs3 = conn.Execute("select ticket_no from reserve_passenger")
Do While rs3.EOF = False
'Combo4.AddItem rs3.Fields(0)
rs3.MoveNext
Loop
rs3.Close
End Sub

Private Sub Form_Load()
Label5.Caption = Format(Date, "dd-mmm-yy")
Label6.Caption = Time$
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set rs3 = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\DVD Writting\DVD_2\VB Projects\RDBMS Project\Microsoft Visual Basic Projects\Airlines Project\Database\airline_db.mdb;Persist Security Info=False"
conn.Open
End Sub




