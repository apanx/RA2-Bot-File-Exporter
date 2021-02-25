Attribute VB_Name = "ModRobot"
Option Explicit
Private Type Header
   Version As String
   Name As String
   Class As Byte
   DefaultTexture As Byte
   Snapshot As Boolean
   BotShot As String
End Type
Private Type Geometry
   BottomPointCount As Long
   BottomPointPos() As String
   TopPointCount As Long
   TopPointPos() As String
   HeightFaceCount As String
End Type
Private Type TDModel
   PointIndexFaceCount As String
   ModelCount As Long
   IndexFacePointCount As String
   RAW As String
   PointNormal() As String
   TexCoord() As String
   PointPos() As String
   FaceIndex() As Long
End Type
Private Type BucklingPoint
   PointPosition As String
   IndexCount As Long
   Index() As Long
End Type
Private Type FacePoint
   PointNumber As Long
   PointPosition() As String
End Type
Private Type Chassis
   Geometry As Geometry
   ModelPresent As Boolean
   TDModel As TDModel
   TexturePresent As Boolean
   Texture As String
   Unknown As Boolean
   FacePointNumber As Long
   FacePoint() As FacePoint
   FaceGroupNumber As Long
   FaceGroup() As String
   FaceEntryNumber As Long
   FaceEntry() As String
   CornerEntryNumber As Long
   CornerEntry() As String
End Type
Private Type Component
   Name As String
   ParentID As Long
   SelfMount As Long
   OtherMount As Long
   Translation As String
   Rotation As String
   AngleHeight As String
   Class As String
   Path As String
   SmartzoneName As String
   BurstSettings As String
End Type
Private Type Controller
   Name As String
   Position As String
   KeyBindNumber As Long
   KeyBind() As String
End Type
Private Type ControllerComponentBind
   Component As Long
   Action As String
   Controller As String
End Type
Private Type Robot
   Header As Header
   ComponentNumber As Long
   Chassis As Chassis
   ArmourType As String
   ArmourData As String
   BucklingPointCount As Long
   BucklingPoint() As BucklingPoint
   Component() As Component
   BuckleModelPresent As Boolean
   BuckleModel As TDModel
   ForwardHeading As Double
   ControllerNumber As Long
   Controller() As Controller
   ControllerComponentBindNumber As Long
   ControllerComponentBind() As ControllerComponentBind
End Type
Public Robot As Robot
