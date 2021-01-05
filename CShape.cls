VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ======= Interface Class ========

Option Explicit

Public Enum ShapeTypeConstants
    shpRect
    shpEllipse
    shpPolygon
End Enum

Private mcol_Coords As Collection
Private m_CombineMode As CombineModeConstants
'=====================================================

Public Property Get CombineMode() As CombineModeConstants
    CombineMode = m_CombineMode
End Property
Public Property Let CombineMode(Mode As CombineModeConstants)
    m_CombineMode = Mode
End Property

Friend Property Get Coords() As Collection
    Set Coords = mcol_Coords
End Property

Public Property Get CoordCount() As Long
    CoordCount = mcol_Coords.Count
End Property

Public Function getRegion() As Long
' You must override this method to return a valid window region.
End Function

Public Sub addCoords(x As Long, y As Long)
    Dim pt As POINTAPI
    
    With pt
        .x = x
        .y = y
    End With
    
    mcol_Coords.Add pt
End Sub

Public Function getSaveString() As String
' Override this method to return a string to write to file.
End Function

Public Sub LoadFromString(strShape As String)
' Override this method to create a shape from values in strShape.
End Sub

'=====================================================
Private Sub Class_Initialize()
    Set mcol_Coords = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcol_Coords = Nothing
End Sub
