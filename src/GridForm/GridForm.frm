VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GridForm 
   Caption         =   "Grid Form"
   ClientHeight    =   9732.001
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   9768.001
   OleObjectBlob   =   "GridForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GridForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private WithEvents Grid         As GridBox
Attribute Grid.VB_VarHelpID = -1
Private WithEvents SidebarGrid  As GridBox
Attribute SidebarGrid.VB_VarHelpID = -1

Dim DataTable As ListObject

Private ThemeColor          As Double
Private PrimaryColor        As Double
Private SecondaryColor      As Double
Private FontColor           As Double

Private TileBluredBackcolor
Private TileBluredBordercolor
Private TileHoveredActive
Private TileHoveredNotActive
Private TileHoveredBordercolor
    
Private TileWidht           As Double
Private TileHeight          As Double

Private SidebarColumnSpan   As Double
Private SidebarRowSpan      As Double
Private SidebarTileHeights  As Double

Private GridColumnSpan      As Double
Private GridRowSpan         As Double
Private GridColumnGap       As Double
Private GridRowGap          As Double
Private GridTileHeights     As Double
Private GridTileWidths      As Double

'USERFORM SECTION

Private Sub UserForm_Initialize()
    Set DataTable = Sheet1.ListObjects("DataTable")
    Call OriginalTheme
    Set Grid = New GridBox
    Set SidebarGrid = New GridBox
    CreateSidebar
    AddSidebarMenuList
    CreateMainGrid
    AddList DataTable.HeaderRowRange.Cells(1, 1).Value
End Sub

Private Sub CreateSidebar()
    SidebarGrid.CreateGrid Sidebar, SidebarColumnSpan, SidebarRowSpan, 65, 0
    SidebarGrid.TileHeights = SidebarTileHeights
    SidebarGrid.TileWidths = Sidebar.InsideWidth
    SidebarGrid.RowGap = 0
End Sub

Private Sub AddSidebarMenuList()
    Dim Tiles As Variant
        Tiles = DataTable.HeaderRowRange.Value
    SidebarGrid.ActiveTileName = Tiles(1, 1)
    Dim index As Integer
    For index = LBound(Tiles, 2) To UBound(Tiles, 2)
        SidebarGrid.AddGridTile Tiles(1, index), 1, 1, Sidebar.BackColor, FontColor
    Next index
End Sub

Private Sub CreateMainGrid()
    Grid.CreateGrid Me, GridColumnSpan, GridRowSpan, 65, 171
    Grid.ColumnGap = GridColumnGap
    Grid.RowGap = GridRowGap
    Grid.TileHeights = GridTileHeights
    Grid.TileWidths = GridTileWidths
End Sub

Private Sub AddList(Optional TileName As String)
    SidebarGrid.RaiseTileHovered TileName
    Grid.RemoveAll
    Dim Tile As Range
    For Each Tile In DataTable.ListColumns(TileName).DataBodyRange.Cells.SpecialCells(xlCellTypeConstants)
        Grid.AddGridTile Tile, 2, 1, SecondaryColor, FontColor
    Next
End Sub

'THEME SECTION

Private Sub OriginalTheme()
    ThemeColor = 8435998        'GREEN
    PrimaryColor = 4208182      'BACKGROUND DARK GRAY
    SecondaryColor = 5457992    'TILE COLORS LIGHTER GRAY
    FontColor = vbWhite
    
    TileBluredBackcolor = 4668733
    TileBluredBordercolor = 5457992  'TILE COLORS LIGHTER GRAY
    TileHoveredBordercolor = 8435998 'GREEN
    TileHoveredActive = 8435998      'GREEN
    TileHoveredNotActive = 5457992   'TILE COLORS LIGHTER GRAY
    
    SidebarColumnSpan = 1
    SidebarRowSpan = 10
    SidebarTileHeights = 30
    
    GridColumnSpan = 4
    GridRowSpan = 10
    GridColumnGap = 1.5
    GridRowGap = 1.5
    GridTileHeights = 50
    GridTileWidths = 65
End Sub

'SIDEBAR SECTION

Private Sub SidebarGrid_TileBlured(Tile As GridTile)
    If Not SidebarGrid.ActiveTileName = Tile.Name Then
        Tile.BackColor = TileBluredBackcolor
    End If
End Sub

Private Sub SidebarGrid_TileHovered(Tile As GridTile)
    If SidebarGrid.ActiveTileName = Tile.Name Then
        Tile.BackColor = TileHoveredActive
    Else
        Tile.BackColor = TileHoveredNotActive
    End If
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'@alexofrhodes
'error 5 occurs when switching sidebar item and clicking on a main grid button
'temp solution to tryagain, dunno
tryagain:
On Error GoTo tryagain
    Grid.RaiseTileHovered ""
    SidebarGrid.RaiseTileHovered ""
End Sub

'GRID SECTION

Private Sub Grid_TileBlured(Tile As GridTile)
    Tile.BorderColor = TileBluredBordercolor
End Sub
Private Sub Grid_TileHovered(Tile As GridTile)
    Tile.BorderColor = TileHoveredBordercolor
End Sub

Private Sub SidebarGrid_TileClicked(Tile As GridTile)
    AddList Tile.Name
End Sub

Private Sub Grid_TileClicked(Tile As GridTile)
    Select Case Tile.Name
    Case "aaa": 'sample in case you don't want the Tile to show a macro name but other text
    Case Else: Application.Run Tile.Name 'have the macros be public in a normal module
    End Select
End Sub


