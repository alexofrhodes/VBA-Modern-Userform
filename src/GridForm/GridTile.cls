VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GridTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private WithEvents pTile As MSForms.Label
Attribute pTile.VB_VarHelpID = -1
Private WithEvents pCaption As MSForms.Label
Attribute pCaption.VB_VarHelpID = -1
Public Parent As GridBox
Attribute Parent.VB_VarHelpID = -1

'******************************************************************************************
' PROPERTIES
'******************************************************************************************
Private Property Get Top() As Double
    Top = pTile.Top
End Property
Private Property Let Top(Value As Double)
    pTile.Top = Value
End Property

Private Property Get Left() As Double
    Left = pTile.Left
End Property
Private Property Let Left(Value As Double)
    pTile.Left = Value
End Property

Private Property Get Width() As Double
    Width = pTile.Width
End Property
Private Property Let Width(Value As Double)
    pTile.Width = Value
End Property

Private Property Get Height() As Double
    Height = pTile.Height
End Property
Private Property Let Height(Value As Double)
    pTile.Height = Value
End Property

Public Property Get Name() As String
    Name = pTile.Name
End Property

Public Property Get BackColor() As Double
    BackColor = pTile.BackColor
End Property
Public Property Let BackColor(Value As Double)
    pTile.BackColor = Value
End Property

Public Property Get BorderColor() As Double
    BorderColor = pTile.BorderColo
End Property

Public Property Let BorderColor(Value As Double)
    pTile.BorderColor = Value
    pTile.BorderStyle = fmBorderStyleNone
    pTile.BorderStyle = fmBorderStyleSingle
End Property

Public Property Get ForeColor() As Double
    ForeColor = pCaption.ForeColor
End Property
Public Property Let ForeColor(Value As Double)
    pCaption.ForeColor = Value
End Property


'******************************************************************************************
' EVENTS
'******************************************************************************************
Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()
    pTile.Parent.Remove pTile.Name
    pCaption.Parent.Remove pCaption.Name
End Sub

'******************************************************************************************
' PUBLIC METHODS/FUNCTIONS
'******************************************************************************************
Public Sub Initialize(GridSource As GridBox, Value As String, TileTop As Double, TileLeft As Double, TileHeight As Double, TileWidth As Double, TileBackColor As Double, TileForeColor As Double)
    
    Set Parent = GridSource
    Set pTile = Parent.GridFrame.Controls.Add("Forms.Label.1", Value)

    With pTile
        .BackColor = TileBackColor
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = TileBackColor
        .ForeColor = TileForeColor
        .BackStyle = fmBackStyleOpaque
    End With
    
    Top = TileTop
    Left = TileLeft
    Height = TileHeight
    Width = TileWidth
    CreateCaption
    
    HandCursor pTile
    
End Sub


'******************************************************************************************
' PRIVATE METHODS/FUNCTIONS
'******************************************************************************************
Private Sub CreateCaption()
    
    Set pCaption = Parent.GridFrame.Controls.Add("Forms.Label.1", "Caption" & Me.Name)
    
    With pCaption
        .Caption = Me.Name
        .AutoSize = True
        .AutoSize = False
        .Width = pTile.Width - 2
        .BackStyle = fmBackStyleTransparent
        .ForeColor = pTile.ForeColor
        .TextAlign = fmTextAlignCenter
        
        .Top = pTile.Top + (pTile.Height * 0.5) - (.Height * 0.5)
        .Left = pTile.Left + (pTile.Width * 0.5) - (.Width * 0.5)
        
    End With
    HandCursor pCaption
    
End Sub

'JUST FOR A CURSOR HAND 'NOT NEEDED
Private Sub HandCursor(Control As MSForms.Control)
    
    On Error GoTo Catch
    With Control
        .MouseIcon = LoadPicture(ThisWorkbook.Path & "\HandCursor.ico")
        .MousePointer = fmMousePointerCustom
    End With
Catch:
    
End Sub

'CALLS TO GRIDBOX TO TRIGGER EVENTS.
Private Sub pTile_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'@alexofrhodes
'error 5 occurs when switching sidebar item and clicking on a main grid button
'temp solution to tryagain, dunno
tryagain:
    On Error GoTo tryagain
    Parent.RaiseTileClicked Me.Name
End Sub
Private Sub pTile_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Parent.RaiseTileHovered Me.Name
End Sub
Private Sub pCaption_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'@alexofrhodes
'error 5 occurs when switching sidebar item and clicking on a main grid button
'temp solution to tryagain, dunno
tryagain:
    On Error GoTo tryagain
    Parent.RaiseTileClicked Me.Name
End Sub
Private Sub pCaption_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Parent.RaiseTileHovered Me.Name
End Sub
