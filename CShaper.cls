VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CShaper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Enum CombineModeConstants
    RGN_AND = 1
    RGN_OR = 2
    RGN_XOR = 3
    RGN_DIFF = 4
    RGN_COPY = 5
End Enum

Private mcol_Shape As Collection

Public Sub ShapeForm(hWnd As Long)
    Dim shp As CShape
    Static hRgn As Long
    Dim l As Long
    
    ' Delete the old region.
    If hRgn <> 0 Then DeleteObject hRgn
    
    hRgn = CreateRectRgn(0, 0, 1, 1)
    
    For Each shp In mcol_Shape
        l = CombineRgn(hRgn, shp.getRegion, hRgn, shp.CombineMode)
    Next shp
    
    SetWindowRgn hWnd, hRgn, True
End Sub

Public Sub AddRect(x1 As Long, y1 As Long, x2 As Long, y2 As Long, Optional CombineMode As CombineModeConstants = RGN_OR)
    Dim rct As CShape
    Set rct = New CShape_Rect
    
    With rct
        .addCoords x1, y1
        .addCoords x2, y2
        .CombineMode = CombineMode
    End With
    
    mcol_Shape.Add rct
End Sub

Public Sub AddEllipse(x1 As Long, y1 As Long, x2 As Long, y2 As Long, Optional CombineMode As CombineModeConstants = RGN_OR)
    Dim elp As CShape
    Set elp = New CShape_Ellipse
    
    With elp
        .addCoords x1, y1
        .addCoords x2, y2
        .CombineMode = CombineMode
    End With
    
    mcol_Shape.Add elp
End Sub

Public Sub AddPolygon(pt() As POINTAPI, Optional CombineMode As CombineModeConstants = RGN_OR)
    Dim ply As CShape
    Dim i As Integer
    Set ply = New CShape_Polygon
    
    ply.CombineMode = CombineMode
    
    For i = 0 To UBound(pt)
        ply.addCoords pt(i).x, pt(i).y
    Next i
    
    mcol_Shape.Add ply
End Sub

Public Sub AddRoundRect(x1 As Long, y1 As Long, x2 As Long, y2 As Long, x3 As Long, y3 As Long, Optional CombineMode As CombineModeConstants = RGN_OR)
    Dim rrt As CShape
    Set rrt = New CShape_RoundRect
    
    With rrt
        .addCoords x1, y1
        .addCoords x2, y2
        .addCoords x3, y3
        .CombineMode = CombineMode
    End With
    
    mcol_Shape.Add rrt
End Sub

Public Sub SaveShape(strFileName As String)
    Dim shp As CShape
    
    On Error GoTo ErrHandler
    
    Open strFileName For Output As #1
    
        For Each shp In mcol_Shape
            Print #1, shp.getSaveString
        Next shp
    
    Close #1
    
    Exit Sub
    
ErrHandler:
    Select Case Err.Number
    Case 75
        MsgBox "Invalid file name passed to SaveShape!", vbExclamation
    Case Else
        MsgBox Err.Number & vbCrLf & Err.Description
        Close
    End Select
End Sub

Public Sub LoadShape(strFileName As String)
    On Error GoTo ErrHandler
    Dim shp As CShape
    Dim strLine As String
    Dim strType As String
    
    Open strFileName For Input As #1
        Do Until EOF(1)
            Line Input #1, strLine
            strType = Left$(strLine, InStr(1, strLine, "/") - 1)
            
            Select Case strType
            Case "Rect": Set shp = New CShape_Rect
            Case "RoundRect": Set shp = New CShape_RoundRect
            Case "Ellipse": Set shp = New CShape_Ellipse
            Case "Polygon": Set shp = New CShape_Polygon
            End Select
            
            shp.LoadFromString strLine
            mcol_Shape.Add shp
        Loop
    Close #1
    
    Exit Sub
ErrHandler:
    Select Case Err.Number
    Case 75
        MsgBox "File does not exist!", vbExclamation
    Case Else
        MsgBox Err.Number & vbCrLf & Err.Description
        Close
    End Select
End Sub

Private Sub Class_Initialize()
    Set mcol_Shape = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcol_Shape = Nothing
End Sub
