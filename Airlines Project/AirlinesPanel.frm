VERSION 5.00
Begin VB.MDIForm AirlinesPanel 
   BackColor       =   &H8000000C&
   Caption         =   "Airlines Panel - Bangalore Airport"
   ClientHeight    =   7470
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11565
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu airdet 
      Caption         =   "&Airport Details"
   End
   Begin VB.Menu Enq 
      Caption         =   "&Enquire"
      Begin VB.Menu Flig 
         Caption         =   "&Flight Details"
      End
      Begin VB.Menu Sea 
         Caption         =   "&Seat Status"
      End
      Begin VB.Menu passdet 
         Caption         =   "&Passenger Details"
      End
      Begin VB.Menu far 
         Caption         =   "&Fares Details"
      End
   End
   Begin VB.Menu oper 
      Caption         =   "&Operations"
      Begin VB.Menu res5 
         Caption         =   "&Reservation"
      End
      Begin VB.Menu can 
         Caption         =   "&Cancellation"
      End
   End
   Begin VB.Menu Enquire 
      Caption         =   "En&quire Report"
      Begin VB.Menu Rep 
         Caption         =   "&Reports"
      End
   End
   Begin VB.Menu pro 
      Caption         =   "&Project"
      Begin VB.Menu about 
         Caption         =   "A&bout"
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu win 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "AirlinesPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click()
AirlinesAbout.Show
End Sub

Private Sub airdet_Click()
AirportDetails.Show
End Sub

Private Sub can_Click()
SeatCancellation.Show
End Sub

Private Sub exit_Click()
Dim msg, style, title, help, ctxt, response, mystring

msg = "Do you want to exit?"
style = vbYesNo
title = "Airline"
help = "DEMO.HLP"
ctxt = 1000
response = MsgBox(msg, vbYesNo, title, help, context)
If response = vbYes Then
mystring = "yes"
MsgBox "Thank You for using this Software", , "Exit"
End
Else
mystring = "no"
AirlinesPanel.Show
End If
End Sub

Private Sub far_Click()
AirlinesFareDetails.Show
End Sub

Private Sub Flig_Click()
AirlinesFlightDetails.Show
End Sub

Private Sub passdet_Click()
AirlinesPassengerDetails.Show
End Sub

Private Sub Rep_Click()
AirlinesReport.Show
End Sub

Private Sub res5_Click()
SeatReservation.Show
End Sub

Private Sub Sea_Click()
SeatStatus.Show
End Sub
