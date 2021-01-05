VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CShape_RoundRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements CShape
Private m_Shape As CShape

Private Sub Class_Initialize()
    Set m_Shape = New CShape
End Sub

Private Sub Class_Terminate()
    Set m_Shape = Nothing
End Sub


Private Sub CShape_addCoords(x As Long, y As Long)
    With m_Shape
        If .CoordCount > 3 Then
            Err.Raise ERR_OVER_MAX_COORD_COUNT, "CShape_Rect::AddCoords()", "You cannot add more than two coordinates"
        Else
            .addCoords x, y
        End If
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
    
    With m_Shape
        hRgn = CreateRoundRectRgn(.Coords(1).x, .Coords(1).y, .Coords(2).x, .Coords(2).y, .Coords(3).x, .Coords(3).y)
    End With
    
    CShape_getRegion = hRgn
End Function

Private Function CShape_getSaveString() As String
    Dim strSave As String
    Dim i As Integer
    
    strSave = "RoundRect/" & CStr(m_Shape.CombineMode)
    
    For i = 1 To m_Shape.CoordCount
        strSave = strSave & "/" & CStr(m_Shape.Coords(i).x) & "," & CStr(m_Shape.Coords(i).y)
    Next i
    
    CShape_getSaveString = strSave
End Function

Private Sub CShape_LoadFromString(strShape As String)
    Dim arValue() As String
    Dim i As Integer
    
    arValue = Split(strShape, "/")
    
    ' Get combine mode.
    m_Shape.CombineMode = Val(arValue(1))
    
    ' Get top-left coords.
    i = InStr(1, arValue(2), ",")
    CShape_addCoords Val(Left$(arValue(2), i - 1)), Val(Mid$(arValue(2), i + 1))
    
    ' Get bottom-right coords.
    i = InStr(1, arValue(3), ",")
    CShape_addCoords Val(Left$(arValue(3), i - 1)), Val(Mid$(arValue(3), i + 1))
    
    ' Get corner bits.
    i = InStr(1, arValue(4), ",")
    CShape_addCoords Val(Left$(arValue(4), i - 1)), Val(Mid$(arValue(4), i + 1))

End Sub
