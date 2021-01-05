VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form Shaper Test"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   4050
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboShape 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   120
      List            =   "frmMain.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   2535
   End
   Begin VB.CommandButton cmdShowShape 
      Caption         =   "Show Shape"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "by Andrew Davey"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select a shape and press Show Shape."
      ForeColor       =   &H80000017&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Form Shaper"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdShowShape_Click()
    Dim myShaper As CShaper
    Set myShaper = New CShaper
    
    Load frmShape
    
    Select Case cboShape.ListIndex
    Case -1
        MsgBox "Please select a shape!", vbExclamation
        Unload frmShape
        Exit Sub
        
    Case 0
        ' rect
        myShaper.AddRect 0, 0, 100, 100
    Case 1
        ' round rect
        myShaper.AddRoundRect 0, 0, 100, 100, 25, 25
    Case 2
        ' ellipse
        myShaper.AddEllipse 0, 0, 100, 100
    Case 3
        ' polygon
        Dim pt(2) As POINTAPI
        pt(0).X = 50: pt(0).Y = 0
        pt(1).X = 100: pt(1).Y = 100
        pt(2).X = 0: pt(2).Y = 100
        
        myShaper.AddPolygon pt()
    Case 4
        myShaper.LoadShape App.Path & "\test.shp"
    End Select
    
    myShaper.ShapeForm frmShape.hWnd
    
    frmShape.Show
End Sub
