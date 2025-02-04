VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GridBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'******************************************************************************************
' PRIVATE PROPERTIES\VARIABLES
'******************************************************************************************

'COLLECTION OF GRIDTILES
Private pGridTiles As Collection

'THE ACTUAL GRID, AS WELL AS AN ARRAY TO KEEP TRACK OF WHERE THE TILES ARE
Private WithEvents pGrid As MSForms.Frame
Attribute pGrid.VB_VarHelpID = -1
Private pGridArray As Variant

'NUMBER OF ROWS AND COLUMNS, SET'S MAX NUMBER OF TILES
Private pColumnSpan As Double
Private pRowSpan As Double

'GAPS BETWEEN TILES
Private pColumnGap As Double
Private pRowGap As Double

'TILE PROPERTIES THAT SHOULD BE SET FOR ALL TILES
Private pTileHeights As Double
Private pTileWidths As Double
Private NameOfHoveredTile As String

'******************************************************************************************
' PUBLIC PROPERTIES
'******************************************************************************************
' - ActiveTileName
' - ColumnGap
' - RowGap
' - Top
' - Left
' - GridFrame
' - TileHeights
' - TileWidths
' - TileCount
' - Tile
' - Tiles

'CUSTOM EVENTS
Public Event TileClicked(Tile As GridTile)
Public Event TileHovered(Tile As GridTile)
Public Event TileBlured(Tile As GridTile)

Public ActiveTileName As String

Public Property Get ColumnGap() As Double
    ColumnGap = pColumnGap
End Property
Public Property Let ColumnGap(Value As Double)
    pColumnGap = Value
    ResizeGrid
End Property

Public Property Get RowGap() As Double
    RowGap = pRowGap
End Property
Public Property Let RowGap(Value As Double)
    pRowGap = Value
    ResizeGrid
End Property

Public Property Get Top() As Double
    Top = pGrid.Top
End Property
Public Property Let Top(Value As Double)
    pGrid.Top = Value
End Property

Public Property Get Left() As Double
    Left = pGrid.Left
End Property
Public Property Let Left(Value As Double)
    pGrid.Left = Value
End Property

Public Property Get GridFrame() As MSForms.Control
    Set GridFrame = pGrid
End Property

Public Property Let TileHeights(Value As Double)
    pTileHeights = Value
End Property
Public Property Get TileHeights() As Double
    TileHeights = pTileHeights
End Property

Public Property Let TileWidths(Value As Double)
    pTileWidths = Value
    ResizeGrid
End Property
Public Property Get TileWidths() As Double
    TileWidths = pTileWidths
    ResizeGrid
End Property

'COLLECTION
Property Get TileCount() As Long
    TileCount = pGridTiles.Count
End Property
Property Get Tile(NameOrIndex As Variant) As GridTile
    Set Tile = pGridTiles(NameOrIndex)
End Property
Property Get Tiles() As Collection
    Set Tiles = pGridTiles
End Property


'******************************************************************************************
' EVENTS
'******************************************************************************************
Private Sub Class_Initialize()
    Set pGridTiles = New Collection
    pTileHeights = 50
    pTileWidths = 50
    pRowGap = 1.5
    pColumnGap = 1.5
End Sub

'******************************************************************************************
' PUBLIC METHODS
'******************************************************************************************

'THESE THREE METHODS THROW EVENTS. CALLED FROM GRIDTILE CLASS.
Public Sub RaiseTileClicked(TileName As String)
    
    Dim Temp As String
    Temp = ActiveTileName
    ActiveTileName = TileName
    
    If Temp <> TileName And Temp <> "" Then
        RaiseEvent TileBlured(Tile(Temp))
    End If
    
    RaiseEvent TileClicked(Tile(TileName))
    RaiseEvent TileHovered(Tile(TileName))
End Sub
Public Sub RaiseTileHovered(TileName As String)
    
    If TileName <> NameOfHoveredTile Then
        
        If NameOfHoveredTile <> "" Then
            RaiseEvent TileBlured(Tile(NameOfHoveredTile))
        End If
        
        If TileName <> "" Then
            RaiseEvent TileHovered(Tile(TileName))
        End If
        
        NameOfHoveredTile = TileName
        
    End If
End Sub
Private Sub pGrid_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'CLEARS OUT ALL TILES HOVER FORMAT
    RaiseTileHovered ""
End Sub

'NEED TO SEPERATE CREATING IT WITH THE FORMATING OF IT
Public Sub CreateGrid(SourceContainer As Object, ColumnSpan As Double, RowSpan As Double, Optional GridTop As Double, Optional GridLeft As Double)
    
    Set pGrid = SourceContainer.Controls.Add("Forms.Frame.1", "Grid")
    pColumnSpan = ColumnSpan
    pRowSpan = RowSpan
    Top = GridTop
    Left = GridLeft
    
    ReDim pGridArray(0 To RowSpan - 1, 0 To ColumnSpan - 1)
    Dim RowIndex As Integer
    Dim ColIndex As Integer
    For RowIndex = LBound(pGridArray, 1) To UBound(pGridArray, 1)
        For ColIndex = LBound(pGridArray, 2) To UBound(pGridArray, 2)
            pGridArray(RowIndex, ColIndex) = -1
        Next ColIndex
    Next RowIndex
    
    ResizeGrid
    
End Sub

Public Sub AddGridTile(ByVal Value As String, ByVal ColumnSpan As Double, ByVal RowSpan As Double, ByVal BackColor As Double, Optional ByVal ForeColor As Double)
    
    Dim RowIndex As Integer
    Dim ColIndex As Integer
'    Dim Location As IndexReference
'    Location = FindArrayGroup(pGridArray, -1, RowSpan, ColumnSpan)

    Dim ColumnsFreeCount As Integer

    For RowIndex = LBound(pGridArray, 1) To UBound(pGridArray, 1)

        ColumnsFreeCount = 0

        For ColIndex = LBound(pGridArray, 2) To UBound(pGridArray, 2)

            If pGridArray(RowIndex, ColIndex) = -1 Then

                pGridArray(RowIndex, ColIndex) = 0

                'INCREMENT THE COUNT OF COLUMNS FREE
                ColumnsFreeCount = ColumnsFreeCount + 1

                'CHECK TO SEE IF THERE ARE ENOUGH FREE COLUMNS
                If ColumnSpan = ColumnsFreeCount Then

                    Dim index As Integer
                    For index = (ColIndex - (ColumnsFreeCount - 1)) To ColIndex
                        pGridArray(RowIndex, index) = ColumnSpan
                    Next index

                    ColIndex = ColIndex - (ColumnsFreeCount - 1)

                    GoTo Complete

                End If
            Else
                ColumnsFreeCount = 0
            End If

        Next ColIndex

        'CLEAR OUT EMPTY ARRAYS
         For ColIndex = LBound(pGridArray, 2) To UBound(pGridArray, 2)
            If pGridArray(RowIndex, index) = Empty Then
                pGridArray(RowIndex, index) = 1
            End If
         Next ColIndex

    Next RowIndex
    
Complete:

    Dim Tile As New GridTile
    Tile.Initialize Me, Value _
      , (RowIndex * TileHeights) + (RowIndex * RowGap) _
      , (ColIndex * TileWidths) + (ColIndex * ColumnGap) _
      , (TileHeights * RowSpan) + ((RowSpan - 1) * RowGap) _
      , (TileWidths * ColumnSpan) + ((ColumnSpan - 1) * ColumnGap) _
      , BackColor _
      , ForeColor _

    'ADD TO COLLECTION
    pGridTiles.Add Tile, Value
    
End Sub

Public Sub Remove(NameOrNumber As Variant)
    pGridTiles.Remove (NameOrNumber)
End Sub

Public Sub RemoveAll()
    
    Set pGridTiles = Nothing
    Set pGridTiles = New Collection
    
    Dim RowIndex As Integer
    Dim ColIndex As Integer
    For RowIndex = LBound(pGridArray, 1) To UBound(pGridArray, 1)
        For ColIndex = LBound(pGridArray, 2) To UBound(pGridArray, 2)
            pGridArray(RowIndex, ColIndex) = -1
        Next ColIndex
    Next RowIndex
    
End Sub


'******************************************************************************************
' PRIVATE METHODS/FUNCTIONS
'******************************************************************************************
Public Sub ResizeGrid()
     With pGrid
        .Height = (pRowSpan * pTileHeights) + ((pRowSpan - 1) * pRowGap)
        .Width = (pColumnSpan * pTileWidths) + ((pColumnSpan - 1) * pColumnGap)
        .BorderStyle = fmBorderStyleSingle
        .BorderStyle = fmBorderStyleNone
    End With
End Sub


'NOT IN USE YET, WILL HELP THOUGH
Private Function FindArrayGroup(SourceArray As Variant, SearchValue As Variant, RowSpan As Double, ColSpan As Double) As Variant
    
    'EXAMPLE OF WHAT ARRAY LOOKS LIKE
    '0 0 1 1
    '1 0 1 1
    '1 0 0 1
    '1 0 0 1
    
    'TRYING TO FIND FIRST GROUP WHERE IT SPANS 2 ROWS AND 2 COLUMNS
    '0 0
    '0 0
    
    'LOOP ROWS
    Dim RowIndex As Integer
    For RowIndex = LBound(SourceArray, 1) To UBound(SourceArray, 1)
        
        'LOOP COLUMNS
        Dim ColIndex As Integer
        For ColIndex = LBound(SourceArray, 2) To UBound(SourceArray, 2)
            
            'COLUMN MATCHES SEARCH VALUE
            If SourceArray(RowIndex, ColIndex) = SearchValue Then
                
                
                'SKIP IF NOT ENOUGH RANGE IN ARRAY
                If RowIndex + (RowSpan - 1) > UBound(SourceArray, 1) Or ColIndex + (ColSpan - 1) > UBound(SourceArray, 2) Then
                    GoTo NotFound
                End If
                
                
                'INNER LOOP CHECKING INNER ARRAY (SKIP ERRORS)
                Dim InnerRowIndex As Integer
                For InnerRowIndex = RowIndex To RowIndex + (RowSpan - 1)
                    
                    Dim InnerColIndex As Integer
                    For InnerColIndex = ColIndex To ColIndex + (ColSpan - 1)
                        If Not SourceArray(InnerRowIndex, InnerColIndex) = SearchValue Then
                            GoTo NotFound
                        End If
                    Next InnerColIndex
                    
                Next InnerRowIndex
                
                'RETURN VALUE
'                Dim Temp As IndexReference
'                Temp.RowIndex = RowIndex
'                Temp.ColIndex = ColIndex
'                FindArrayGroup = Temp
                Exit Function
    
            End If
            
NotFound:

        Next ColIndex
    Next RowIndex

End Function



