VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form SeatStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seat Status - Airlines"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "SeatStatus.frx":0000
   ScaleHeight     =   7965
   ScaleWidth      =   12000
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3960
      Width           =   3435
   End
   Begin VB.TextBox txtaseat 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   4680
      Width           =   3435
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   8280
      TabIndex        =   7
      Top             =   3360
      Width           =   3435
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2760
      Width           =   3435
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   8280
      TabIndex        =   5
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   7440
   End
   Begin VB.CommandButton Cmdclose 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "C&lose"
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "&View"
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   1815
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   3615
      Left            =   0
      TabIndex        =   4
      Top             =   2640
      Width           =   5175
      _Version        =   524288
      _ExtentX        =   9128
      _ExtentY        =   6376
      _StockProps     =   1
      BackColor       =   12632319
      Year            =   2008
      Month           =   6
      Day             =   30
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   0
      GridLinesColor  =   -2147483640
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
      Left            =   840
      TabIndex        =   3
      Top             =   7320
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
      Left            =   9720
      TabIndex        =   2
      Top             =   7320
      Width           =   1695
   End
End
Attribute VB_Name = "SeatStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sum As Integer
Dim A As Integer
Dim B As String
Dim intYear As Integer
Dim intMonth As Integer
Dim intDay As Integer
Dim lngSerial As Long
Dim strNewDate As String
Dim datDate As Date
Dim jd As Date
Dim intSeatRet As Integer
Option Explicit

Private Sub Calendar1_Click()
Text4.Text = Format(Calendar1.Value, "dd-mmm-yy")
End Sub

Private Sub Cmdclose_Click()
conn.Close
Unload Me
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
 If (KeyAscii <> 77) And (KeyAscii <> 70) Then
        MsgBox "Select from the given list", vbExclamation, "List"
        KeyAscii = 0
    End If
End Sub

Private Sub Command1_Click()
Set rs = conn.Execute("select * from flight where flight_no='" + Combo1 + "'")

If Combo1 = "" And Text4 = "" Then
MsgBox "Please select the Flight number and jouney date", vbOKOnly, "SEAT"
Exit Sub
End If
Do While rs.EOF = False
If Combo1 = rs.Fields(0) Then
If Combo1 = "" Then
MsgBox "Please select the Flight number", vbOKOnly, "SEAT"
Exit Sub
Else
Text1.Text = rs.Fields(4)
Text5.Text = rs.Fields(1)
End If
If Text4.Text = "" Then
MsgBox "Please select the journey date from the calender", vbOKOnly, "seat"
Exit Sub
End If
jd = Format(Text4.Text, "dd-mmm-yy")
If jd < strNewDate Then
MsgBox "Please select a valid date", vbOKOnly, "seat"
Call clear
Exit Sub
End If
Call Avail
'rs.MoveNext
Exit Sub
End If
Loop
MsgBox "Please select a valid flight number", vbOKOnly, "seat"
Exit Sub
End Sub

Private Sub Form_Activate()
Set rs.ActiveConnection = conn
Set rs1.ActiveConnection = conn
Set rs = conn.Execute("select flight_no from flight")
Do While rs.EOF = False
Combo1.AddItem rs.Fields(0)
rs.MoveNext
Loop
rs.Close
End Sub

Private Sub Form_Load()
datDate = CDate("01-jul-08")
intYear = Year(datDate)
intMonth = Month(datDate)
intDay = Day(datDate)
lngSerial = DateSerial(intYear, intMonth, intDay)
strNewDate = Format$(lngSerial, "dd-mmm-yy")
Label2.Caption = strNewDate
Label3.Caption = Time$

Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\DVD Writting\DVD_2\VB Projects\RDBMS Project\Microsoft Visual Basic Projects\Airlines Project\Database\airline_db.mdb;Persist Security Info=False"
conn.Open
End Sub
Private Sub Avail()
Set rs = conn.Execute("select * from reserve_passenger where jdate = '" + Text4 + "' and flight_no = '" + Combo1 + "'")
A = 0
Do While rs.EOF = False
A = A + 1
rs.MoveNext
Loop
If A = Text1 Then
MsgBox "No Seats Available"
Else
sum = Text1 - A
txtaseat = sum
End If
End Sub
Sub clear()
Text4 = ""
Text1 = ""
Text5 = ""
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
 If (KeyAscii <> 77) And (KeyAscii <> 70) Then
        MsgBox "Select from the given calender", vbExclamation, "Calender"
        KeyAscii = 0
    End If
End Sub

Private Sub Timer1_Timer()
If flag = True Then
rs.Close
flag = False
End If
End Sub

