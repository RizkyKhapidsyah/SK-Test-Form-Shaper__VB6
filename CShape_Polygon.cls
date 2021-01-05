VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CShape_Polygon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements CShape
Private m_Shape As CShape
Private m_pt() As POINTAPI

Private Sub Class_Initialize()
    Set m_Shape = New CShape
End Sub

Private Sub Class_Terminate()
    Set m_Shape = Nothing
End Sub


Private Sub CShape_addCoords(x As Long, y As Long)
    m_Shape.addCoords x, y
    
    ReDim Preserve m_pt(0 To m_Shape.CoordCount - 1) As POINTAPI
    With m_pt(m_Shape.CoordCount - 1)
        .x = x
        .y = y
    End With
End Sub

Private Property Let CShape_CombineMode(RHS As CombineModeConstants)
    m_Shape.CombineMode = RHS
End Property
Private Property Get CShape_CombineMode() As CombineModeConstants
    CShape_CombineMode = m_Shape.CombineMode
End Property

Private Property Get CShape_CoordCount() As Long
    CShape_CoordCount = m_Shape.CoordCount
End Property

Private Function CShape_getRegion() As Long
    Dim hRgn As Long

    hRgn = CreatePolygonRgn(m_pt(0), m_Shape.CoordCount, 1)
        
    CShape_getRegion = hRgn
End Function

Private Function CShape_getSaveString() As String
    Dim strSave As String
    Dim i As Integer
    
    strSave = "Polygon/" & CStr(m_Shape.CombineMode) & "/" & CStr(m_Shape.CoordCount)
    
    For i = 1 To m_Shape.CoordCount
        strSave = strSave & "/" & CStr(m_Shape.Coords(i).x) & "," & CStr(m_Shape.Coords(i).y)
    Next i
    
    CShape_getSaveString = strSave
End Function

Private Sub CShape_LoadFromString(strShape As String)
    Dim arValue() As String
    Dim i As Integer, j As Integer
    Dim nPoints As Integer
    
    arValue = Split(strShape, "/")
    
    ' Get combine mode.
    m_Shape.CombineMode = Val(arValue(1))
    
    ' Get polygon point count.
    nPoints = Val(arValue(2))
    
    For j = 3 To nPoints + 2
        ' Get top-left coords.
        i = InStr(1, arValue(j), ",")
        CShape_addCoords Val(Left$(arValue(j), i - 1)), Val(Mid$(arValue(j), i + 1))
    Next j
End Sub
