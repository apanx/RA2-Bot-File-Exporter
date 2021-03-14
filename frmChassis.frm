VERSION 5.00
Begin VB.Form frmChassis 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Make Chassis"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   Icon            =   "frmChassis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5280
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   4800
      Width           =   1815
   End
   Begin VB.ListBox lstPoints 
      Height          =   4545
      Left            =   5280
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Line lneChassisLength 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   16  'Merge Pen
      Index           =   0
      X1              =   1320
      X2              =   1320
      Y1              =   3960
      Y2              =   5040
   End
   Begin VB.Shape shpChassisPoint 
      BorderColor     =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   5160
      Width           =   195
   End
   Begin VB.Line Line22 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   5120
      Y1              =   4620
      Y2              =   4620
   End
   Begin VB.Line Line21 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   5120
      Y1              =   4120
      Y2              =   4120
   End
   Begin VB.Line Line20 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   5120
      Y1              =   3620
      Y2              =   3620
   End
   Begin VB.Line Line19 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   5120
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   5120
      Y1              =   2120
      Y2              =   2120
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   5120
      Y1              =   1620
      Y2              =   1620
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   5120
      Y1              =   1120
      Y2              =   1120
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   5120
      Y1              =   620
      Y2              =   620
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00808080&
      X1              =   5120
      X2              =   5120
      Y1              =   120
      Y2              =   5120
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00808080&
      X1              =   4620
      X2              =   4620
      Y1              =   120
      Y2              =   5120
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00808080&
      X1              =   4120
      X2              =   4120
      Y1              =   120
      Y2              =   5120
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00808080&
      X1              =   3615
      X2              =   3615
      Y1              =   120
      Y2              =   5120
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00808080&
      X1              =   2620
      X2              =   2620
      Y1              =   120
      Y2              =   5120
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00808080&
      X1              =   2120
      X2              =   2120
      Y1              =   120
      Y2              =   5120
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      X1              =   1620
      X2              =   1620
      Y1              =   120
      Y2              =   5120
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   1120
      X2              =   1120
      Y1              =   120
      Y2              =   5120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   620
      X2              =   620
      Y1              =   120
      Y2              =   5120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   5120
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   5120
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   5120
      Y1              =   5115
      Y2              =   5115
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   5120
      Y1              =   2620
      Y2              =   2620
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00808080&
      X1              =   3120
      X2              =   3120
      Y1              =   120
      Y2              =   5120
   End
End
Attribute VB_Name = "frmChassis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Dim i As Long
Dim temp() As String
For i = 0 To Robot.Chassis.Geometry.BottomPointCount - 1
    lstPoints.AddItem Robot.Chassis.Geometry.BottomPointPos(i)
    temp() = Split(Robot.Chassis.Geometry.BottomPointPos(i), " ")
    If Not i = 0 Then Load shpChassisPoint(i)
    shpChassisPoint(i).Visible = True
    shpChassisPoint(i).Left = temp(0) * 5 + 20
    shpChassisPoint(i).Top = temp(1) * 5 + 20
    shpChassisPoint(i).ZOrder
Next
For i = 0 To Robot.Chassis.Geometry.BottomPointCount - 1
    If Not i = 0 Then Load lneChassisLength(i)
    If i = Robot.Chassis.Geometry.BottomPointCount - 1 Then Exit For
    lneChassisLength(i).Visible = True
    lneChassisLength(i).X1 = shpChassisPoint(i).Left + 100
    lneChassisLength(i).X2 = shpChassisPoint(i + 1).Left + 100
    lneChassisLength(i).Y1 = shpChassisPoint(i).Top + 100
    lneChassisLength(i).Y2 = shpChassisPoint(i + 1).Top + 100
    lneChassisLength(i).ZOrder
Next
i = Robot.Chassis.Geometry.BottomPointCount - 1
lneChassisLength(i).Visible = True
lneChassisLength(i).X1 = shpChassisPoint(i).Left + 100
lneChassisLength(i).X2 = shpChassisPoint(0).Left + 100
lneChassisLength(i).Y1 = shpChassisPoint(i).Top + 100
lneChassisLength(i).Y2 = shpChassisPoint(0).Top + 100
lneChassisLength(i).ZOrder
End Sub
