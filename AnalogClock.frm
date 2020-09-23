VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   180
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private hourHand As Single, _
        minuteHand As Single, _
        secondHand As Single

Private sizeX As Long, _
        sizeY As Long

Private Sub Timer1_Timer()
' Draws the clock hands when seconds have changed
Dim sec As Single, cir As Single

  ' FYI:
  ' pi = Atn(1) * 4
  ' date = Int(Now)
  ' time = Now - Int(Now)
  
  ' hour hand makes 2 revolutions per day (pi * 4)
  cir = Atn(1) * 16
  hourHand = (Now - Int(Now)) * cir
  ' and it needs to go clockwise (radians increase counter-clockwise)
  hourHand = cir - hourHand
  ' minute hand is 12 times as fast as the hour hand
  minuteHand = hourHand * 12
  ' second hand is 60 times as fast as the minute hand
  sec = minuteHand * 60
  
  ' only re-draw if the seconds have changed
  If secondHand <> sec Then
    secondHand = sec
    Cls
    DrawWidth = 7
    Line (sizeX, sizeY)-Step(-Sin(hourHand) * (sizeX * 0.5), _
                             -Cos(hourHand) * (sizeY * 0.5)), _
                             vbBlack
    DrawWidth = 3
    Line (sizeX, sizeY)-Step(-Sin(minuteHand) * (sizeX * 0.85), _
                             -Cos(minuteHand) * (sizeY * 0.85)), _
                             vbBlack
    DrawWidth = 1
    Line (sizeX, sizeY)-Step(-Sin(secondHand) * sizeX, _
                             -Cos(secondHand) * sizeY), _
                             vbRed
  End If
End Sub

Private Sub Form_Load()
  ' Initialize form and timer
  Move 2000, 2000, 4000, 4240
  Font.Size = 18
  Font.Bold = True
  Timer1.Interval = 250
  Timer1.Enabled = True
End Sub

Private Sub Form_Resize()
' catch down-sizing
  Form_Paint
End Sub

Private Sub Form_Paint()
' Draws the clock face when needed
Dim num As String
Dim angle As Single, sinX As Single, cosY As Single
Dim tick As Long, fontX As Long, fontY As Long
Static Busy As Boolean

  ' Avoid re-entry problems (resizing form can cause many paints)
  If Busy Then Exit Sub
  Busy = True
  
  ' Fit to current form size
  sizeX = ScaleWidth / 2
  sizeY = ScaleHeight / 2
  
  ' Start with a blank screen
  AutoRedraw = True
  Cls
  DrawWidth = 3
  
  ' Loop through clock circle (starts at 1 o'clock)
  For angle = 8.9 To 2.7 Step -0.1047198
    ' Draw tick marks
    tick = tick + 1
    sinX = Sin(angle)
    cosY = Cos(angle)
    Line (sizeX + sinX * (sizeX * 0.9), _
          sizeY + cosY * (sizeY * 0.9))- _
         (sizeX + sinX * sizeX, _
          sizeY + cosY * sizeY), _
          vbBlack
    ' Make every 5th tick darker
    Select Case tick Mod 5
    Case 0
      DrawWidth = 3
    Case 1
      ' Center number where it belongs
      num = CStr(tick \ 5 + 1)
      fontX = TextWidth(num) / 2
      fontY = TextHeight(num) / 2
      PSet (sizeX + sinX * (sizeX * 0.75) - fontX, _
            sizeY + cosY * (sizeY * 0.75) - fontY), _
            Me.BackColor
      Print num
      DrawWidth = 1
    Case Else
      DrawWidth = 1
    End Select
  Next
  
  ' Save image to memory
  AutoRedraw = False
  ' Force redraw of hands
  secondHand = -1
  
  Busy = False
End Sub

