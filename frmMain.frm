VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exporter"
   ClientHeight    =   6990
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6645
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtWeaponString 
      Height          =   390
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   6480
      Width           =   6615
   End
   Begin MSComctlLib.TreeView TvwCompList 
      Height          =   5895
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   10398
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   6
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblBotname 
      AutoSize        =   -1  'True
      Caption         =   "No Bot Loaded"
      Height          =   270
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblConnectTo 
      Caption         =   "Label1"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   3375
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open..."
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Begin VB.Menu mnuCompID 
            Caption         =   "Component ID data"
         End
         Begin VB.Menu mnuSnapshot 
            Caption         =   "Snapshot"
         End
         Begin VB.Menu mnuTex 
            Caption         =   "Texture"
         End
      End
      Begin VB.Menu mnuhyp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuChassis 
         Caption         =   "Chassis"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
 frmChassis.Show
End Sub
Private Sub mnuAbout_Click()
 MsgBox "No fancy stuff here", vbInformation, "The About Thingy"
 MsgBox "This thing was made by apanx", vbInformation, "The About Thingy"
End Sub

Private Sub mnuChassis_Click()
frmChassis.Show
End Sub

Private Sub mnuClose_Click()
End
End Sub

Private Sub mnuOpen_Click()
On Error GoTo ErrHandler
CommonDialog1.CancelError = True
CommonDialog1.Filter = "Bot File (*.bot)|*.bot"
CommonDialog1.ShowOpen
Close #1
Open CommonDialog1.FileName For Binary As #1
If LOF(1) = 0 Then Exit Sub
Dim Data As String
Dim Success As Boolean
Data = Space$(LOF(1))
Get #1, , Data
Dim i As Long
i = 1
Dim run As Long
For run = 1 To 5
    Do
        If Mid$(Data, i, 1) = vbLf Then
            Select Case run
                Case 1
                    Robot.Header.Version = Left$(Data, i - 1)
                    If Not Robot.Header.Version = "1.12" Then
                        MsgBox "Incompatible version of bot file. Please send in the bot file to apanx@apanx.net", vbCritical, "Error"
                        Exit Sub
                    End If
                    Data = Right$(Data, Len(Data) - i)
                    i = 1
                    Exit Do
                Case 2
                    Robot.Header.Name = Left$(Data, i - 1)
                    Robot.Header.Name = Right(Robot.Header.Name, Len(Robot.Header.Name) - 6)
                    Data = Right$(Data, Len(Data) - i)
                    i = 1
                    Exit Do
                Case 3
                    Robot.Header.Class = Right$(Left$(Data, i - 1), Len(Left$(Data, i - 1)) - 6)
                    Data = Right$(Data, Len(Data) - i)
                    i = 1
                    Exit Do
                Case 4
                    Robot.Header.DefaultTexture = Left$(Data, i - 1)
                    Data = Right$(Data, Len(Data) - i)
                    i = 1
                    Exit Do
                Case 5
                    Robot.Header.Snapshot = Left$(Data, i - 1)
                    Data = Right$(Data, Len(Data) - i)
                    i = 1
                    Exit Do
            End Select
        End If
        i = i + 1
    Loop
Next
run = 1
If Robot.Header.Snapshot Then
    Robot.Header.BotShot = Left$(Data, 65554)
    Data = Right$(Data, Len(Data) - 65554)
End If
Do
    If Mid$(Data, i, 1) = vbLf Then
    Robot.ComponentNumber = Left$(Data, i - 1)
    Data = Right$(Data, Len(Data) - i)
    i = 1
    Exit Do
    End If
    i = i + 1
Loop
ReDim Robot.Component(Robot.ComponentNumber - 1)
i = 1
For run = 1 To 8
    Do
        If Mid$(Data, i, 1) = vbLf Then
            Select Case run
                Case 1
                    Robot.Component(0).ParentID = Left$(Data, i - 1)
                    Data = Right$(Data, Len(Data) - i)
                    i = 1
                    Exit Do
                Case 2
                    Robot.Component(0).SelfMount = Left$(Data, i - 1)
                    Data = Right$(Data, Len(Data) - i)
                    i = 1
                    Exit Do
                Case 3
                    Robot.Component(0).OtherMount = Left$(Data, i - 1)
                    Data = Right$(Data, Len(Data) - i)
                    i = 1
                    Exit Do
                Case 4
                    Robot.Component(0).Translation = Left$(Data, i - 1)
                    Data = Right$(Data, Len(Data) - i)
                    i = 1
                    Exit Do
                Case 5
                    Robot.Component(0).Rotation = Left$(Data, i - 1)
                    Data = Right$(Data, Len(Data) - i)
                    i = 1
                    Exit Do
                Case 6
                    Robot.Component(0).AngleHeight = Left$(Data, i - 1)
                    Data = Right$(Data, Len(Data) - i)
                    i = 1
                    Exit Do
                Case 7
                    Robot.Component(0).Class = Left$(Data, i - 1)
                    Data = Right$(Data, Len(Data) - i)
                    i = 1
                    Exit Do
                Case 8
                    Robot.Component(0).Path = Left$(Data, i - 1)
                    Data = Right$(Data, Len(Data) - i)
                    i = 1
                    Exit Do
            End Select
        End If
        i = i + 1
    Loop
Next
Do
    If Mid$(Data, i, 1) = vbLf Then
        Robot.Chassis.Geometry.BottomPointCount = Left$(Data, i - 1)
        Data = Right$(Data, Len(Data) - i)
        i = 1
        Exit Do
    End If
    i = i + 1
Loop
If Robot.Chassis.Geometry.BottomPointCount = 0 Then GoTo TopPointCount
ReDim Robot.Chassis.Geometry.BottomPointPos(Robot.Chassis.Geometry.BottomPointCount - 1)
For run = 1 To Robot.Chassis.Geometry.BottomPointCount
    Do
        If Mid$(Data, i, 1) = vbLf Then
            Robot.Chassis.Geometry.BottomPointPos(run - 1) = Left$(Data, i - 1)
            Data = Right$(Data, Len(Data) - i)
            i = 1
            Exit Do
        End If
        i = i + 1
    Loop
Next
TopPointCount:
Do
    If Mid$(Data, i, 1) = vbLf Then
        Robot.Chassis.Geometry.TopPointCount = Left$(Data, i - 1)
        Data = Right$(Data, Len(Data) - i)
        i = 1
        Exit Do
    End If
    i = i + 1
Loop
If Robot.Chassis.Geometry.TopPointCount = 0 Then GoTo HeightFaceCount
ReDim Robot.Chassis.Geometry.TopPointPos(Robot.Chassis.Geometry.TopPointCount - 1)
For run = 1 To Robot.Chassis.Geometry.TopPointCount
    Do
        If Mid$(Data, i, 1) = vbLf Then
            Robot.Chassis.Geometry.TopPointPos(run - 1) = Left$(Data, i - 1)
            Data = Right$(Data, Len(Data) - i)
            i = 1
            Exit Do
        End If
        i = i + 1
    Loop
Next
HeightFaceCount:
Do
    If Mid$(Data, i, 1) = vbLf Then
        Robot.Chassis.Geometry.HeightFaceCount = Left$(Data, i - 1)
        Data = Right$(Data, Len(Data) - i)
        i = 1
        Exit Do
    End If
    i = i + 1
Loop
Do
    If Mid$(Data, i, 1) = vbLf Then
        Robot.Chassis.ModelPresent = Left$(Data, i - 1)
        Data = Right$(Data, Len(Data) - i)
        i = 1
        Exit Do
    End If
    i = i + 1
Loop
If Robot.Chassis.ModelPresent Then
    Do
        If Mid$(Data, i, 1) = vbLf Then
            Robot.Chassis.TDModel.PointIndexFaceCount = Left$(Data, i - 1)
            Data = Right$(Data, Len(Data) - i)
            i = 1
            Exit Do
        End If
        i = i + 1
    Loop
    Do
        If Mid$(Data, i, 1) = vbLf Then
            Robot.Chassis.TDModel.ModelCount = Left$(Data, i - 1)
            Data = Right$(Data, Len(Data) - i)
            i = 1
            Exit Do
        End If
        i = i + 1
    Loop
    Dim temp() As String
    Dim Numbers(1 To 3) As Long
    temp() = Split(Robot.Chassis.TDModel.PointIndexFaceCount, " ")
    For i = 1 To UBound(temp)
        Numbers(i) = CLng(temp(i))
    Next
    i = 1
    Do
        If Mid$(Data, i, 1) = vbLf Then
            Robot.Chassis.TDModel.IndexFacePointCount = Left$(Data, i - 1)
            Data = Right$(Data, Len(Data) - i)
            i = 1
            Exit Do
        End If
        i = i + 1
    Loop
    Do
        If Mid$(Data, i, 1) = vbLf Then
            Robot.Chassis.TDModel.RAW = Left$(Data, i - 1)
            Data = Right$(Data, Len(Data) - i)
            i = 1
            Exit Do
        End If
        i = i + 1
    Loop
    ReDim Robot.Chassis.TDModel.PointNormal(1 To Numbers(1))
    ReDim Robot.Chassis.TDModel.TexCoord(1 To Numbers(1))
    ReDim Robot.Chassis.TDModel.PointPos(1 To Numbers(1))
    Dim j As Long
    For j = 1 To 3
        Select Case j
            Case 1
                For run = 1 To Numbers(1)
                    Do
                        If Mid$(Data, i, 1) = vbLf Then
                            Robot.Chassis.TDModel.PointNormal(run) = Left$(Data, i - 1)
                            Data = Right$(Data, Len(Data) - i)
                            i = 1
                            Exit Do
                        End If
                        i = i + 1
                    Loop
                Next
            Case 2
                For run = 1 To Numbers(1)
                    Do
                        If Mid$(Data, i, 1) = vbLf Then
                            Robot.Chassis.TDModel.TexCoord(run) = Left$(Data, i - 1)
                            Data = Right$(Data, Len(Data) - i)
                            i = 1
                            Exit Do
                        End If
                        i = i + 1
                    Loop
                Next
        Case 3
                For run = 1 To Numbers(1)
                    Do
                        If Mid$(Data, i, 1) = vbLf Then
                            Robot.Chassis.TDModel.PointPos(run) = Left$(Data, i - 1)
                            Data = Right$(Data, Len(Data) - i)
                            i = 1
                            Exit Do
                        End If
                        i = i + 1
                    Loop
                Next
        End Select
    Next
    ReDim Robot.Chassis.TDModel.FaceIndex(1 To Numbers(2))
    For run = 1 To Numbers(2)
        Do
            If Mid$(Data, i, 1) = vbLf Then
                Robot.Chassis.TDModel.FaceIndex(run) = Left$(Data, i - 1)
                Data = Right$(Data, Len(Data) - i)
                i = 1
                Exit Do
            End If
            i = i + 1
        Loop
    Next
    Do
        If Mid$(Data, i, 1) = vbLf Then
            Robot.Chassis.TexturePresent = Left$(Data, i - 1)
            Data = Right$(Data, Len(Data) - i)
            i = 1
            Exit Do
        End If
        i = i + 1
    Loop
    If Robot.Chassis.TexturePresent Then
        Dim tempstorage As String
        Do
            If LCase$(Mid$(Data, i, 4)) = "alse" Or LCase$(Mid$(Data, i, 4)) = "true" Then
                Robot.Chassis.Texture = Left$(Data, i + 3 - 1)
                Data = Right$(Data, Len(Data) - i - 4)
                i = 1
                Exit Do
            End If
            i = i + 1
        Loop
        tempstorage = LCase$(Right$(Robot.Chassis.Texture, 4))
        If tempstorage = "alse" Then
            Robot.Chassis.Unknown = False
            Robot.Chassis.Texture = Left$(Robot.Chassis.Texture, Len(Robot.Chassis.Texture) - 5)
        End If
        If tempstorage = "true" Then
            Robot.Chassis.Unknown = True
            Robot.Chassis.Texture = Left$(Robot.Chassis.Texture, Len(Robot.Chassis.Texture) - 4)
        End If
        'Robot.Chassis.Texture = Left$(Data, 262162)
        'Data = Right$(Data, Len(Data) - 262162)
        GoTo Unknown
    End If
    Do
        If Mid$(Data, i, 1) = vbLf Then
            Robot.Chassis.Unknown = Left$(Data, i - 1)
            Data = Right$(Data, Len(Data) - i)
            i = 1
            Exit Do
        End If
        i = i + 1
    Loop
Unknown:
    If Robot.Chassis.Unknown Then
        MsgBox "Your Robot has an unknown section set to True. This program does not support .bot files of this type. You could help the development of this program greatly by sending the .bot file to apanx@apanx.net", vbInformation, "Message"
        Exit Sub
    End If
    Do
        If Mid$(Data, i, 1) = vbLf Then
            Robot.Chassis.FacePointNumber = Left$(Data, i - 1)
            Data = Right$(Data, Len(Data) - i)
            i = 1
            Exit Do
        End If
        i = i + 1
    Loop
    ReDim Robot.Chassis.FacePoint(1 To Robot.Chassis.FacePointNumber)
    For j = 1 To Robot.Chassis.FacePointNumber
        Do
            If Mid$(Data, i, 1) = vbLf Then
                Robot.Chassis.FacePoint(j).PointNumber = Left$(Data, i - 1)
                Data = Right$(Data, Len(Data) - i)
                i = 1
                Exit Do
            End If
            i = i + 1
        Loop
        ReDim Robot.Chassis.FacePoint(j).PointPosition(1 To Robot.Chassis.FacePoint(j).PointNumber)
        For run = 1 To Robot.Chassis.FacePoint(j).PointNumber
            Do
                If Mid$(Data, i, 1) = vbLf Then
                    Robot.Chassis.FacePoint(j).PointPosition(run) = Left$(Data, i - 1)
                    Data = Right$(Data, Len(Data) - i)
                    i = 1
                    Exit Do
                End If
                i = i + 1
            Loop
        Next
    Next
    Do
        If Mid$(Data, i, 1) = vbLf Then
            Robot.Chassis.FaceGroupNumber = Left$(Data, i - 1)
            Data = Right$(Data, Len(Data) - i)
            i = 1
            Exit Do
        End If
        i = i + 1
    Loop
    ReDim Robot.Chassis.FaceGroup(1 To Robot.Chassis.FaceGroupNumber)
    For run = 1 To Robot.Chassis.FaceGroupNumber
        Do
            If Mid$(Data, i, 1) = vbLf Then
                Robot.Chassis.FaceGroup(run) = Left$(Data, i - 1)
                Data = Right$(Data, Len(Data) - i)
                i = 1
                Exit Do
            End If
            i = i + 1
        Loop
    Next
    Do
        If Mid$(Data, i, 1) = vbLf Then
            Robot.Chassis.FaceEntryNumber = Left$(Data, i - 1)
            Data = Right$(Data, Len(Data) - i)
            i = 1
            Exit Do
        End If
        i = i + 1
    Loop
    ReDim Robot.Chassis.FaceEntry(1 To Robot.Chassis.FaceEntryNumber)
    For run = 1 To Robot.Chassis.FaceEntryNumber
        Do
            If Mid$(Data, i, 1) = vbLf Then
                Robot.Chassis.FaceEntry(run) = Left$(Data, i - 1)
                Data = Right$(Data, Len(Data) - i)
                i = 1
                Exit Do
            End If
            i = i + 1
        Loop
    Next
    Do
        If Mid$(Data, i, 1) = vbLf Then
            Robot.Chassis.CornerEntryNumber = Left$(Data, i - 1)
            Data = Right$(Data, Len(Data) - i)
            i = 1
            Exit Do
        End If
        i = i + 1
    Loop
    ReDim Robot.Chassis.CornerEntry(1 To Robot.Chassis.CornerEntryNumber)
    For run = 1 To Robot.Chassis.CornerEntryNumber
        Do
            If Mid$(Data, i, 1) = vbLf Then
                Robot.Chassis.CornerEntry(run) = Left$(Data, i - 1)
                Data = Right$(Data, Len(Data) - i)
                i = 1
                Exit Do
            End If
            i = i + 1
        Loop
    Next
End If
Do
    If Mid$(Data, i, 1) = vbLf Then
        Robot.ArmourType = Left$(Data, i - 1)
        Data = Right$(Data, Len(Data) - i)
        i = 1
        Exit Do
    End If
    i = i + 1
Loop
Do
    If Mid$(Data, i, 1) = vbLf Then
        Robot.ArmourData = Left$(Data, i - 1)
        Data = Right$(Data, Len(Data) - i)
        i = 1
        Exit Do
    End If
    i = i + 1
Loop
If Robot.Chassis.ModelPresent Then
    Do
        If Mid$(Data, i, 1) = vbLf Then
            Robot.BucklingPointCount = Left$(Data, i - 1)
            Data = Right$(Data, Len(Data) - i)
            i = 1
            Exit Do
        End If
        i = i + 1
    Loop
End If
If Robot.BucklingPointCount = 0 Then GoTo Components
ReDim Robot.BucklingPoint(1 To Robot.BucklingPointCount)
For j = 1 To Robot.BucklingPointCount
    Do
        If Mid$(Data, i, 1) = vbLf Then
            Robot.BucklingPoint(j).PointPosition = Left$(Data, i - 1)
            Data = Right$(Data, Len(Data) - i)
            i = 1
            Exit Do
        End If
        i = i + 1
    Loop
    Do
        If Mid$(Data, i, 1) = vbLf Then
            Robot.BucklingPoint(j).IndexCount = Left$(Data, i - 1)
            Data = Right$(Data, Len(Data) - i)
            i = 1
            Exit Do
        End If
        i = i + 1
    Loop
    ReDim Robot.BucklingPoint(j).Index(1 To Robot.BucklingPoint(j).IndexCount)
    For run = 1 To Robot.BucklingPoint(j).IndexCount
        Do
            If Mid$(Data, i, 1) = vbLf Then
                Robot.BucklingPoint(j).Index(run) = Left$(Data, i - 1)
                Data = Right$(Data, Len(Data) - i)
                i = 1
                Exit Do
            End If
            i = i + 1
        Loop
    Next
Next
Components:
If Robot.ComponentNumber = 1 Then GoTo BuckleModel
For j = 1 To Robot.ComponentNumber - 1
    For run = 1 To 8
        Do
            If Mid$(Data, i, 1) = vbLf Then
                Select Case run
                    Case 1
                        Robot.Component(j).ParentID = Left$(Data, i - 1)
                        Data = Right$(Data, Len(Data) - i)
                        i = 1
                        Exit Do
                    Case 2
                        Robot.Component(j).SelfMount = Left$(Data, i - 1)
                        Data = Right$(Data, Len(Data) - i)
                        i = 1
                        Exit Do
                    Case 3
                        Robot.Component(j).OtherMount = Left$(Data, i - 1)
                        Data = Right$(Data, Len(Data) - i)
                        i = 1
                        Exit Do
                    Case 4
                        Robot.Component(j).Translation = Left$(Data, i - 1)
                        Data = Right$(Data, Len(Data) - i)
                        i = 1
                        Exit Do
                    Case 5
                        Robot.Component(j).Rotation = Left$(Data, i - 1)
                        Data = Right$(Data, Len(Data) - i)
                        i = 1
                        Exit Do
                    Case 6
                        Robot.Component(j).AngleHeight = Left$(Data, i - 1)
                        Data = Right$(Data, Len(Data) - i)
                        i = 1
                        Exit Do
                    Case 7
                        Robot.Component(j).Class = Left$(Data, i - 1)
                        Data = Right$(Data, Len(Data) - i)
                        i = 1
                        Exit Do
                    Case 8
                        Robot.Component(j).Path = Left$(Data, i - 1)
                        Data = Right$(Data, Len(Data) - i)
                        i = 1
                        Exit Do
                End Select
            End If
            i = i + 1
        Loop
    Next
    If Robot.Component(j).Class = "BurstMotor" Then
        Do
            If Mid$(Data, i, 1) = vbLf Then
                Robot.Component(j).BurstSettings = Left$(Data, i - 1)
                Data = Right$(Data, Len(Data) - i)
                i = 1
                Exit Do
            End If
            i = i + 1
        Loop
    End If
    If Robot.Component(j).Class = "SmartZone" Then
        Do
            If Mid$(Data, i, 1) = vbLf Then
                Robot.Component(j).SmartzoneName = Left$(Data, i - 1)
                Data = Right$(Data, Len(Data) - i)
                i = 1
                Exit Do
            End If
            i = i + 1
        Loop
    End If
Next
BuckleModel:
Do
    If Mid$(Data, i, 1) = vbLf Then
        Robot.BuckleModelPresent = Left$(Data, i - 1)
        Data = Right$(Data, Len(Data) - i)
        i = 1
        Exit Do
    End If
    i = i + 1
Loop
If Robot.BuckleModelPresent Then
    Do
        If Mid$(Data, i, 1) = vbLf Then
            Robot.BuckleModel.PointIndexFaceCount = Left$(Data, i - 1)
            Data = Right$(Data, Len(Data) - i)
            i = 1
            Exit Do
        End If
        i = i + 1
    Loop
    Do
        If Mid$(Data, i, 1) = vbLf Then
            Robot.BuckleModel.ModelCount = Left$(Data, i - 1)
            Data = Right$(Data, Len(Data) - i)
            i = 1
            Exit Do
        End If
        i = i + 1
    Loop
    Dim Numbers2(1 To 3) As Long
    temp() = Split(Robot.BuckleModel.PointIndexFaceCount, " ")
    For i = 1 To UBound(temp)
        Numbers2(i) = CLng(temp(i))
    Next
    i = 1
    Do
        If Mid$(Data, i, 1) = vbLf Then
            Robot.BuckleModel.IndexFacePointCount = Left$(Data, i - 1)
            Data = Right$(Data, Len(Data) - i)
            i = 1
            Exit Do
        End If
        i = i + 1
    Loop
    Do
        If Mid$(Data, i, 1) = vbLf Then
            Robot.BuckleModel.RAW = Left$(Data, i - 1)
            Data = Right$(Data, Len(Data) - i)
            i = 1
            Exit Do
        End If
        i = i + 1
    Loop
    ReDim Robot.BuckleModel.PointNormal(1 To Numbers2(1))
    ReDim Robot.BuckleModel.TexCoord(1 To Numbers2(1))
    ReDim Robot.BuckleModel.PointPos(1 To Numbers2(1))
    For j = 1 To 3
        Select Case j
            Case 1
                For run = 1 To Numbers2(1)
                    Do
                        If Mid$(Data, i, 1) = vbLf Then
                            Robot.BuckleModel.PointNormal(run) = Left$(Data, i - 1)
                            Data = Right$(Data, Len(Data) - i)
                            i = 1
                            Exit Do
                        End If
                        i = i + 1
                    Loop
                Next
            Case 2
                For run = 1 To Numbers2(1)
                    Do
                        If Mid$(Data, i, 1) = vbLf Then
                            Robot.BuckleModel.TexCoord(run) = Left$(Data, i - 1)
                            Data = Right$(Data, Len(Data) - i)
                            i = 1
                            Exit Do
                        End If
                        i = i + 1
                    Loop
                Next
        Case 3
                For run = 1 To Numbers2(1)
                    Do
                        If Mid$(Data, i, 1) = vbLf Then
                            Robot.BuckleModel.PointPos(run) = Left$(Data, i - 1)
                            Data = Right$(Data, Len(Data) - i)
                            i = 1
                            Exit Do
                        End If
                        i = i + 1
                    Loop
                Next
        End Select
    Next
    ReDim Robot.BuckleModel.FaceIndex(1 To Numbers2(2))
    For run = 1 To Numbers2(2)
        Do
            If Mid$(Data, i, 1) = vbLf Then
                Robot.BuckleModel.FaceIndex(run) = Left$(Data, i - 1)
                Data = Right$(Data, Len(Data) - i)
                i = 1
                Exit Do
            End If
            i = i + 1
        Loop
    Next
End If
Do
    If Mid$(Data, i, 1) = vbLf Then
        Robot.ForwardHeading = Val(Left$(Data, i - 1))
        Data = Right$(Data, Len(Data) - i)
        i = 1
        Exit Do
    End If
    i = i + 1
Loop
Do
    If Mid$(Data, i, 1) = vbLf Then
        Robot.ControllerNumber = Left$(Data, i - 1)
        Data = Right$(Data, Len(Data) - i)
        i = 1
        Exit Do
    End If
    i = i + 1
Loop
If Robot.ControllerNumber = 0 Then GoTo ControllerBind
ReDim Robot.Controller(1 To Robot.ControllerNumber)
For j = 1 To Robot.ControllerNumber
    Do
        If Mid$(Data, i, 1) = vbLf Then
            Robot.Controller(j).Name = Left$(Data, i - 1)
            Robot.Controller(j).Name = Right(Robot.Controller(j).Name, Len(Robot.Controller(j).Name) - 6)
            Data = Right$(Data, Len(Data) - i)
            i = 1
            Exit Do
        End If
        i = i + 1
    Loop
    Do
        If Mid$(Data, i, 1) = vbLf Then
            Robot.Controller(j).Position = Left$(Data, i - 1)
            Data = Right$(Data, Len(Data) - i)
            i = 1
            Exit Do
        End If
        i = i + 1
    Loop
    Do
        If Mid$(Data, i, 1) = vbLf Then
            Robot.Controller(j).KeyBindNumber = Left$(Data, i - 1)
            Data = Right$(Data, Len(Data) - i)
            i = 1
            Exit Do
        End If
        i = i + 1
    Loop
    ReDim Robot.Controller(j).KeyBind(1 To Robot.Controller(j).KeyBindNumber)
    For run = 1 To Robot.Controller(j).KeyBindNumber
        Do
            If Mid$(Data, i, 1) = vbLf Then
                Robot.Controller(j).KeyBind(run) = Left$(Data, i - 1)
                Data = Right$(Data, Len(Data) - i)
                i = 1
                Exit Do
            End If
            i = i + 1
        Loop
    Next
Next
ControllerBind:
Do
    If Mid$(Data, i, 1) = vbLf Then
        Robot.ControllerComponentBindNumber = Left$(Data, i - 1)
        Data = Right$(Data, Len(Data) - i)
        i = 1
        Exit Do
    End If
    i = i + 1
Loop
If Robot.ControllerComponentBindNumber = 0 Then GoTo loaded:
ReDim Robot.ControllerComponentBind(1 To Robot.ControllerComponentBindNumber)
For j = 1 To Robot.ControllerComponentBindNumber
    Do
        If Mid$(Data, i, 1) = vbLf Then
            Robot.ControllerComponentBind(j).Component = Left$(Data, i - 1)
            Data = Right$(Data, Len(Data) - i)
            i = 1
            Exit Do
        End If
        i = i + 1
    Loop
    Do
        If Mid$(Data, i, 1) = vbLf Then
            Robot.ControllerComponentBind(j).Action = Left$(Data, i - 1)
            Data = Right$(Data, Len(Data) - i)
            i = 1
            Exit Do
        End If
        i = i + 1
    Loop
    Do
        If Mid$(Data, i, 1) = vbLf Then
            Robot.ControllerComponentBind(j).Controller = Left$(Data, i - 1)
            Data = Right$(Data, Len(Data) - i)
            i = 1
            Exit Do
        End If
        i = i + 1
    Loop
Next
loaded:
lblBotname.Caption = Robot.Header.Name
'Load data into UI
TvwCompList.Nodes.Clear
TvwCompList.Nodes.Add , , , "Chassis"
For run = 1 To Robot.ComponentNumber - 1
    Robot.Component(run).Name = Left$(Robot.Component(run).Path, Len(Robot.Component(run).Path) - 4)
    ReDim temp(0 To UBound(Split(Robot.Component(run).Name, "\")))
    temp() = Split(Robot.Component(run).Name, "\")
    Robot.Component(run).Name = temp(UBound(temp()))
    Robot.Component(run).Name = UCase(Left$(Robot.Component(run).Name, 1)) & Right$(Robot.Component(run).Name, Len(Robot.Component(run).Name) - 1)
    TvwCompList.Nodes.Add Robot.Component(run).ParentID + 1, tvwChild, , Robot.Component(run).Name & " (" & run & ")"
    TvwCompList.Nodes(1).Expanded = True
    DoEvents
Next
Close #1
Exit Sub
ErrHandler:
If Err = 32755 Then Exit Sub
MsgBox "An error occured during load. Please send in the offending .bot file to apanx@apanx.net in order to aid development", vbCritical, "Error"
End Sub
Private Sub mnuCompID_Click()
On Error GoTo ErrHandler
Dim i As Long
Dim Data As String
Data = Robot.Header.Name + vbCrLf
For i = 1 To TvwCompList.Nodes.Count
    If TvwCompList.Nodes(i).Children = 0 Then
        Data = Data + TvwCompList.Nodes(i).FullPath & vbCrLf
    End If
Next
CommonDialog1.FileName = Robot.Header.Name + ".txt"
CommonDialog1.Filter = "Text (*.txt)|*.txt"
CommonDialog1.ShowSave
Open CommonDialog1.FileName For Binary As #1
Put #1, , Data
Close #1
Exit Sub
ErrHandler:
If Err = 32755 Then Exit Sub
MsgBox "An error occured during save", vbCritical, "Error"
End Sub
Private Sub mnuSnapshot_Click()
On Error GoTo ErrHandler
If Robot.Header.Snapshot Then
Dim i As Long
Dim Data As String
Data = Robot.Header.BotShot
CommonDialog1.Filter = "Targa (*.tga)|*.tga"
CommonDialog1.FileName = Robot.Header.Name + " Snapshot" + ".tga"
CommonDialog1.ShowSave
Open CommonDialog1.FileName For Binary As #1
Put #1, , Data
Close #1
Exit Sub
Else
MsgBox "No snapshot in this botfile", vbInformation, "No snapshot"
Exit Sub
End If
ErrHandler:
If Err = 32755 Then Exit Sub
MsgBox "An error occured during save", vbCritical, "Error"
End Sub

Private Sub mnuTex_Click()
On Error GoTo ErrHandler
If Robot.Chassis.TexturePresent Then
Dim i As Long
Dim Data As String
Data = Robot.Chassis.Texture
CommonDialog1.Filter = "Targa (*.tga)|*.tga"
CommonDialog1.FileName = Robot.Header.Name + " Texture" + ".tga"
CommonDialog1.ShowSave
Open CommonDialog1.FileName For Binary As #1
Put #1, , Data
Close #1
Exit Sub
Else
MsgBox "No texture in this botfile", vbInformation, "No texture"
Exit Sub
End If
ErrHandler:
If Err = 32755 Then Exit Sub
MsgBox "An error occured during save", vbCritical, "Error"
End Sub

Private Sub TvwCompList_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim aNode As MSComctlLib.Node
Dim weapons As String
weapons = "'weapons': ("
   For Each aNode In TvwCompList.Nodes
      If aNode.Checked Then
        weapons = weapons + Str(aNode.Index - 1) + ","
        End If
   Next
Mid$(weapons, Len(weapons), 1) = " "
weapons = weapons + ")"
txtWeaponString = weapons
End Sub
Private Sub txtWeaponString_DblClick()
txtWeaponString.SelStart = 0
txtWeaponString.SelLength = Len(txtWeaponString)
End Sub
