VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form SeatReservation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seat Reservation Details"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "SeatReservation.frx":0000
   ScaleHeight     =   7965
   ScaleWidth      =   12000
   Begin VB.TextBox T11 
      BackColor       =   &H00FFFFFF&
      DataField       =   "AIRLINES"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   4560
      Width           =   2475
   End
   Begin VB.TextBox T12 
      BackColor       =   &H00FFFFFF&
      DataField       =   "RDATE"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   5475
      Width           =   1995
   End
   Begin VB.TextBox T10 
      BackColor       =   &H00FFFFFF&
      DataField       =   "FLIGHT_NO"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   6270
      Width           =   2000
   End
   Begin VB.TextBox T9 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   " "
      Top             =   6720
      Width           =   615
   End
   Begin VB.TextBox T8 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   5880
      Width           =   2415
   End
   Begin VB.TextBox T14 
      BackColor       =   &H00FFFFFF&
      DataField       =   "AMOUNT"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   5040
      Width           =   2000
   End
   Begin VB.TextBox T15 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   10800
      TabIndex        =   14
      Top             =   6720
      Width           =   975
   End
   Begin VB.TextBox T1 
      DataField       =   "PASSENGER_NAME"
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   3000
      TabIndex        =   13
      Top             =   1560
      Width           =   3435
   End
   Begin VB.TextBox T13 
      DataField       =   "JDATE"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   3000
      TabIndex        =   12
      Top             =   6600
      Width           =   3435
   End
   Begin VB.TextBox T2 
      DataField       =   "PASSENGER_NAME"
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   3000
      TabIndex        =   11
      Top             =   2040
      Width           =   3435
   End
   Begin VB.ComboBox Com3 
      Height          =   315
      Left            =   3000
      TabIndex        =   10
      Top             =   5400
      Width           =   3495
   End
   Begin VB.ComboBox Com2 
      Height          =   315
      ItemData        =   "SeatReservation.frx":1A06C4
      Left            =   3000
      List            =   "SeatReservation.frx":1A06D1
      TabIndex        =   9
      Top             =   6000
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Reserve"
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7080
      Width           =   1455
   End
   Begin VB.TextBox T3 
      Height          =   405
      Left            =   3000
      MaxLength       =   3
      TabIndex        =   7
      Top             =   4320
      Width           =   1215
   End
   Begin VB.ComboBox Com1 
      Height          =   315
      ItemData        =   "SeatReservation.frx":1A06F3
      Left            =   5160
      List            =   "SeatReservation.frx":1A06FD
      TabIndex        =   6
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox T6 
      Height          =   315
      Left            =   3000
      TabIndex        =   5
      Top             =   2640
      Width           =   3435
   End
   Begin VB.TextBox T4 
      Height          =   315
      Left            =   3000
      TabIndex        =   4
      Top             =   3240
      Width           =   3435
   End
   Begin VB.TextBox T7 
      Height          =   285
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   3
      Top             =   4920
      Width           =   3435
   End
   Begin VB.TextBox T5 
      Height          =   285
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   2
      Top             =   3840
      Width           =   3435
   End
   Begin VB.CommandButton cmdclose 
      BackColor       =   &H00FFC0C0&
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
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7080
      Width           =   1455
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2655
      Left            =   6960
      TabIndex        =   1
      Top             =   1440
      Width           =   4815
      _Version        =   524288
      _ExtentX        =   8493
      _ExtentY        =   4683
      _StockProps     =   1
      BackColor       =   16761024
      Year            =   2007
      Month           =   3
      Day             =   26
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trajan Pro"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trajan Pro"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trajan Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      TabIndex        =   22
      Top             =   7560
      Width           =   1695
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
      TabIndex        =   21
      Top             =   7560
      Width           =   1695
   End
End
Attribute VB_Name = "SeatReservation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim max As String
Dim DD As Date
Dim DD1 As Date



Private Sub Calendar1_Click()
T13.Text = Format(Calendar1.Value, "dd-mmm-yyyy")
End Sub

Private Sub Cmdclose_Click()
Unload Me
End Sub



Private Sub Com1_KeyPress(KeyAscii As Integer)
 If (KeyAscii <> 77) And (KeyAscii <> 70) Then
        MsgBox "Select from the given list", vbExclamation, "List"
        KeyAscii = 0
    End If
End Sub

Private Sub Com2_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 77) And (KeyAscii <> 70) Then
        MsgBox "Select from the given list", vbExclamation, "List"
        KeyAscii = 0
    End If
End Sub

Private Sub Com3_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 77) And (KeyAscii <> 70) Then
        MsgBox "Select from the given list", vbExclamation, "List"
        KeyAscii = 0
    End If
End Sub

Private Sub Command1_Click()
'DETAILS.Visible = False
If T1.Text = "" Or T2.Text = "" Or T3.Text = "" Or Com1.Text = "" Or T4.Text = "" Or T5.Text = "" Or T6.Text = "" Or T7.Text = "" Or Com2.Text = "" Or Com3.Text = "" Or T13.Text = "" Then
 MsgBox "Details are incomplete,please enter appropriate details", vbOKOnly, "res"
Exit Sub
End If
If IsNumeric(T2.Text) Then
 MsgBox " Enter a valid Passenger", vbOKOnly, "RES"
 T2.SetFocus
 Exit Sub
 End If
T12.Text = Format(Date, "dd-mmm-yy")
DD = Format(T12, "dd-mmm-yy")
DD1 = Format(T13, "dd-mmm-yy")
If DD1 < DD Then
MsgBox " PLEASE ENTER THE DATE WHICH IS MORE THEN RESERVATION DATE", vbOKOnly, "RES"
Exit Sub
 End If
 
Call ticketno
Call seatno
Set rs1 = conn.Execute("select * from FLIGHT where ROUTE='" + Com3 + "'")
T10 = rs1.Fields(0)
T11 = rs1.Fields(1)
Set rs2 = conn.Execute("select * from fare where class='" + Com2 + "' ")
T14 = rs2.Fields(4)
T15 = rs2.Fields(5)
Set rs3 = conn.Execute("select passport_no from reserve_passenger")
Do While rs3.EOF = False
If T1 = rs3.Fields(0) Then
MsgBox "Passport number Invalid", vbOKOnly, "res"
Call clear
Exit Sub
End If
rs3.MoveNext
Loop
conn.Execute "insert into reserve_passenger values('" + T1 + "','" + T2 + "'," + T3 + ",'" + Com1 + "','" + T4 + "','" + T5 + "','" + T6 + "'," + T7 + ",'" + T8 + "'," + T9 + ",'" + T10 + "','" + T11 + "','" + Com2 + "','" + Com3 + "','" + T12 + "','" + T13 + "','" + T14 + "','" + T15 + "')"
MsgBox "Seat reserved succesfully"
'DETAILS.Visible = True
End Sub

Private Sub Form_Activate()
Set rs.ActiveConnection = conn
Set rs = conn.Execute("select route from flight")
Do While rs.EOF = False
Com3.AddItem rs.Fields(0)
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

Private Sub ticketno()
Set rs.ActiveConnection = conn
rs.Open ("select max(ticket_no) from reserve_passenger")
If IsNull(rs.Fields(0)) Then
T8 = "1000"
Else
T8 = rs.Fields(0) + 1
End If
rs.Close
End Sub
Private Sub seatno()
Set rs.ActiveConnection = conn
rs.Open ("select max(seat_no) from reserve_passenger")
If IsNull(rs.Fields(0)) Then
T9 = "1"
Else
T9 = rs.Fields(0) + 1
End If
rs.Close
End Sub

Private Sub clear()
T1 = ""
T2 = ""
T3 = ""
T4 = ""
T5 = ""
T6 = ""
T7 = ""
T13 = ""
Com1 = ""
Com3 = ""
Com2 = ""
End Sub

Private Sub T1_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And Not (KeyAscii = 8) And Not (KeyAscii = 46) Then
KeyAscii = 0
End If
If (KeyAscii < 8) And Not (KeyAscii = 46) Then
MsgBox "Please enter only numbers", , "Number"
End If
End Sub

Private Sub T13_KeyPress(KeyAscii As Integer)
 If (KeyAscii <> 77) And (KeyAscii <> 70) Then
        MsgBox "Select from the given calender", vbexlamation, "Calender"
        KeyAscii = 0
    End If
End Sub

Private Sub T2_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) Or (KeyAscii >= 33 And KeyAscii <= 64) Or (KeyAscii >= 91 And KeyAscii <= 96) Or (KeyAscii >= 123 And KeyAscii <= 126) Then
KeyAscii = 0
MsgBox "Only letters shud b entered", vbOKOnly, " Wats this!"
T2 = " "
T2.SetFocus
End If
End Sub

Private Sub T3_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 58) And KeyAscii <> 8 Then
KeyAscii = 0
MsgBox "Only numbers are allowed to enter", vbOKOnly, "Wats this!"
End If
If Len(T3) > 1 Then
KeyAscii = 0
MsgBox " max digits allowed to enter are 2", vbOKOnly
End If
End Sub

Private Sub T5_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii = 47) Then
KeyAscii = 0
End If
If (KeyAscii < 8) And Not (KeyAscii = 46) Then
MsgBox "Please enter only Characters", , "Character"
End If
End Sub

Private Sub T6_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) Or (KeyAscii >= 33 And KeyAscii <= 64) Or (KeyAscii >= 91 And KeyAscii <= 96) Or (KeyAscii >= 123 And KeyAscii <= 126) Then
KeyAscii = 0
MsgBox "Only letters shud b entered", vbOKOnly, " Wats this!"
T6 = " "
T6.SetFocus
End If
End Sub

Private Sub T7_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 58) And KeyAscii <> 8 Then
KeyAscii = 0
MsgBox "Only numbers are allowed to enter", vbOKOnly, "wats this!"
End If
If Len(T7) > 7 Then
KeyAscii = 0
MsgBox " Maximum digits allowed to enter are 8"
End If
End Sub

Private Sub Timer1_Timer()
If flag = True Then
rs.Close
flag = False
End If
End Sub

