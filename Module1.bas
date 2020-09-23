Attribute VB_Name = "Module1"
Option Explicit

Public ObjectIndex As Integer
Public RelationIndex As Integer
Public GridSize As Integer
Public SizeFormat As Integer
Public i As Integer
Public j As Integer
Public k As Integer
Public ColumnMarge As Single
Public FileContent As String
Public CopyIndex As Integer

Public tempTable1 As Integer
Public tempTable2 As Integer
Public tempColumn1 As Integer
Public tempColumn2 As Integer

Dim LeftRightMarge As Integer
Dim LineTopValue As Integer
Dim LineLeftValue As Integer
Dim CreateId As Integer
Dim OldTableIndex As Integer
Dim OldTableCaption As String

'prevent flickering
Declare Function LockWindowUpdate Lib "user32" _
    (ByVal hWnd As Long) As Long

Public Sub LockWindow(hWnd As Long, blnValue As Boolean)
    If blnValue Then
        LockWindowUpdate hWnd
    Else
        LockWindowUpdate 0
    End If
End Sub

'create new empty project
Function NewEmptyProject()

    If MsgBox("Are you sure?", vbQuestion + vbYesNo, "Create new project") = vbYes Then
        
        CreateEmptyProject
        
    End If

End Function

Function DeleteRelation()

    If MsgBox("Are you sure?", vbQuestion + vbYesNo, "Delete relation") = vbYes Then
        Main.RelationLabel(RelationIndex).Tag = ""
        Main.RelationLabel(RelationIndex).Visible = False
        Main.RelationLeft(RelationIndex).Visible = False
        Main.RelationLeftType(RelationIndex).Visible = False
        Main.RelationRight(RelationIndex).Visible = False
        Main.RelationRightType(RelationIndex).Visible = False
        Main.RelationVertical(RelationIndex).Visible = False
        Main.Combo2.RemoveItem Main.Combo2.ListIndex
        Main.Combo2.Text = "-- None --"
        DeselectRelation
    End If

End Function

Function DeselectRelation()

    For i = 1 To Main.RelationLabel.UBound
        Main.RelationLabel(i).FontBold = False
    Next i
    
    Main.mnuEditDeselect.Enabled = False
    Main.Toolbar1.Buttons(7).Enabled = False
    Main.mnuEditDelete.Enabled = False
    Main.Combo2.Text = "-- None --"
    Main.Combo2.Enabled = True
    RelationIndex = 0

End Function

Function SelectRelation()

    For i = 1 To Main.RelationLabel.UBound
        Main.RelationLabel(i).FontBold = False
    Next i

    Main.mnuEditDelete.Enabled = True
    Main.mnuEditDeselect.Enabled = True
    Main.Toolbar1.Buttons(7).Enabled = True
    Main.RelationLabel(RelationIndex).FontBold = True
    Main.Combo2.Text = Main.RelationLabel(RelationIndex).Caption
    Main.Combo3.Text = Main.RelationLabel(RelationIndex).Tag
    Main.Combo2.Enabled = False

End Function

Function CreateEmptyProject()
        
    Main.StatusBar1.Panels(1).Text = "Untitled.sdd"
    Main.mnuFileSave.Enabled = False
    Main.mnuFileSaveAs.Caption = "Save Untitled &As..."
    
    'unload table
    For i = 1 To Main.Table.UBound
        Unload Main.Table(i)
        Unload Main.TableShadow(i)
        Unload Main.TableLabel(i)
        Unload Main.TableSelector(i)
    Next i
    
    'unload columns
    For i = 1 To Main.TableColumnName.UBound
        Unload Main.TableColumnName(i)
        Unload Main.TableColumnType(i)
    Next i
    
    'unload relations
    For i = 1 To Main.RelationLabel.UBound
        Unload Main.RelationLabel(i)
        Unload Main.RelationLeft(i)
        Unload Main.RelationRight(i)
        Unload Main.RelationVertical(i)
        Unload Main.RelationLeftType(i)
        Unload Main.RelationRightType(i)
    Next i
    
    'clear combo boxes
    Main.Combo1.Clear
    Main.Combo2.Clear
    Main.Combo3.Clear
    
    'add default combo boxes data
    Main.Combo1.AddItem "-- None --"
    Main.Combo2.AddItem "-- None --"
    Main.Combo3.AddItem "Default"
    
    'select default combo boxes data
    Main.Combo1.Text = "-- None --"
    Main.Combo2.Text = "-- None --"
    Main.Combo3.Text = "Default"

End Function

'resize all elements
Function ResizeElements()

    On Error Resume Next
    
    Main.WorkSpace.Height = Main.Height - Main.StatusBar1.Height - Main.Properties.Height - 1000
    
    Main.BarLeft.Left = 0
    Main.BarLeft.Top = 0 + Main.BarTop.Height
    Main.BarLeft.Height = Main.WorkSpace.Height - Main.BarTop.Height
    
    Main.BarTop.Top = 0
    Main.BarTop.Left = 0 + Main.BarLeft.Width
    Main.BarTop.Width = Main.WorkSpace.Width - Main.BarLeft.Width
    
    Main.BarEmptyTop.Left = 0
    Main.BarEmptyTop.Top = 0
    Main.BarEmptyTop.Width = Main.BarLeft.Width
    Main.BarEmptyTop.Height = Main.BarTop.Height
    
    Main.BarEmptyBottom.Left = 0 + Main.WorkSpace.ScaleWidth - Main.BarEmptyBottom.Width
    Main.BarEmptyBottom.Top = 0 + Main.WorkSpace.ScaleHeight - Main.BarEmptyBottom.Height - 0.9
    
    Main.EditorShadow.Height = Main.Editor.Height
    Main.EditorShadow.Width = Main.Editor.Width
    Main.EditorShadow.Left = Main.Editor.Left + 0.5
    Main.EditorShadow.Top = Main.Editor.Top + 0.5
    
    Main.LineTop.Width = Main.Editor.ScaleWidth
    Main.LineTop.Left = Main.Editor.Left - Main.BarLeft.ScaleWidth
    Main.LineLeft.Top = Main.Editor.Top - Main.BarTop.ScaleHeight
    Main.LineLeft.Height = Main.Editor.ScaleHeight
    
    Main.HScroll.Left = 0 + Main.BarLeft.ScaleWidth
    Main.HScroll.Top = Main.WorkSpace.ScaleHeight - Main.HScroll.Height - 0.9
    Main.HScroll.Width = Main.WorkSpace.ScaleWidth - Main.BarLeft.ScaleWidth - 4.5
    
    Main.VScroll.Top = 0 + Main.BarTop.ScaleHeight
    Main.VScroll.Height = Main.WorkSpace.ScaleHeight - Main.BarTop.ScaleHeight - 5.2
    Main.VScroll.Left = Main.WorkSpace.ScaleWidth - Main.VScroll.Width
    
    GenerateScrollbars
    GenerateLineNumbers
    GenerateGrid

End Function

'generate scrollbars
Function GenerateScrollbars()

    LeftRightMarge = 10

    If Main.Editor.ScaleWidth + LeftRightMarge > Main.WorkSpace.ScaleWidth Then
        Main.HScroll.Min = 1
        Main.HScroll.Max = Main.Editor.Width
        Main.HScroll.Enabled = True
    Else
        Main.HScroll.Enabled = False
    End If
    
    If Main.Editor.ScaleHeight > Main.WorkSpace.ScaleHeight Then
        Main.VScroll.Min = 1
        Main.VScroll.Max = Main.Editor.Height
        Main.VScroll.Enabled = True
    Else
        Main.VScroll.Enabled = False
    End If

End Function

'generate grid
Function GenerateGrid()

    Main.Editor.Cls

    If Main.mnuViewGrid.Checked = True Then

        LineTopValue = (Round(Main.LineTop.ScaleWidth, 0)) / GridSize
        LineLeftValue = (Round(Main.LineLeft.ScaleHeight, 0)) / GridSize

        'vertical grid
        For i = 1 To LineTopValue
            Main.Editor.Line ((i * GridSize), Main.WorkSpace.Height)-((i * GridSize), 0)
        Next i

        'horizontal grid
        For i = 1 To LineLeftValue
            Main.Editor.Line (0, (i * GridSize))-(Main.WorkSpace.Width, (i * GridSize))
        Next i

    End If

End Function

'generate line numbers (vertical and horizontal)
Function GenerateLineNumbers()

    Dim tmpI As Integer

    Main.LineTop.Cls
    Main.LineLeft.Cls
    
    'horizontal numbers
    LineTopValue = (Round(Main.LineTop.ScaleWidth, 0)) / SizeFormat
    For i = 1 To LineTopValue - 1
        tmpI = i * SizeFormat
        Main.LineTop.CurrentX = (i * SizeFormat) - (Len(tmpI) - 0.3)
        Main.LineTop.CurrentY = 1
        Main.LineTop.Print i
    Next i
    
    'vertical numbers
    LineLeftValue = (Round(Main.LineLeft.ScaleHeight, 0)) / SizeFormat
    For i = 1 To LineLeftValue - 1
        tmpI = i * SizeFormat
        Main.LineLeft.CurrentX = 0
        Main.LineLeft.CurrentY = (i * SizeFormat) - 0.3
        Main.LineLeft.Print i
    Next i

End Function

'create new table
Function CreateTable(TableName As String, Width As Integer, Height As Integer, Left As Integer, Top As Integer, Layer As String, Color As String)

    Load Main.Table(Main.Table.Count)
    Load Main.TableShadow(Main.TableShadow.Count)
    Load Main.TableLabel(Main.TableLabel.Count)
    Load Main.TableSelector(Main.TableSelector.Count)

    'shadow (set position to back of current table)
    With Main.TableShadow(Main.TableShadow.UBound)
        .ZOrder (0)
    End With

    'table
    With Main.Table(Main.Table.UBound)
        .Left = Left
        .Top = Top
        .FillColor = Color
        .Width = Width
        .Height = Height
        .Tag = Layer
        .ZOrder (0)
        .Visible = True
    End With
    
    'shadow
    With Main.TableShadow(Main.TableShadow.UBound)
        .Left = Main.Table(Main.Table.UBound).Left + 0.3
        .Top = Main.Table(Main.Table.UBound).Top + 0.3
        .Width = Main.Table(Main.Table.UBound).Width
        .Height = Main.Table(Main.Table.UBound).Height
        .Visible = True
    End With
    
    'label
    With Main.TableLabel(Main.TableLabel.UBound)
        .Caption = TableName
        .ForeColor = ReadIniValue(App.Path & "\Setup.ini", "Properties", "HeaderFontColor")
        .BackColor = ReadIniValue(App.Path & "\Setup.ini", "Properties", "HeaderColor")
        .FontSize = ReadIniValue(App.Path & "\Setup.ini", "Properties", "HeaderSize")
        .Left = Main.Table(Main.Table.UBound).Left
        .Top = Main.Table(Main.Table.UBound).Top
        .Width = Main.Table(Main.Table.UBound).Width
        .ZOrder (0)
        .Visible = True
    End With
    
    'selector
    With Main.TableSelector(Main.TableSelector.UBound)
        .Left = Main.Table(Main.Table.UBound).Left + 0.3
        .Top = Main.Table(Main.Table.UBound).Top + 0.3
        .Width = Main.Table(Main.Table.UBound).Width
        .Height = Main.Table(Main.Table.UBound).Height
        .ZOrder (0)
        .Visible = True
    End With
    
    Main.Combo1.AddItem TableName

End Function

'deselect tables
Function DeselectTables()

    For i = 1 To Main.Table.UBound
        If Main.Table(i).Visible = True Then
            Main.TableSelector(i).DragMode = 0
            Main.TableSelector(i).Visible = True
            Main.TableLabel(i).Enabled = 1
        End If
    Next i

    For i = 1 To Main.TableColumnName.UBound
        Main.TableColumnName(i).Enabled = True
        Main.TableColumnType(i).Enabled = True
    Next i

    Main.Selector.Visible = False
    Main.Combo1.Text = "-- None --"
    
    Main.Combo1.Enabled = True
    
    Main.Toolbar1.Buttons(11).Enabled = False
    Main.Toolbar1.Buttons(12).Enabled = False
    Main.mnuEditProperties.Enabled = False
    Main.Toolbar1.Buttons(18).Enabled = False
    Main.Toolbar1.Buttons(5).Enabled = False
    Main.Toolbar1.Buttons(7).Enabled = False

End Function

'select table
Function SelectTable()

    For i = 1 To Main.Table.UBound
        Main.TableSelector(i).Visible = False
    Next i
    
    Main.mnuEditProperties.Enabled = True
    Main.Toolbar1.Buttons(5).Enabled = True
    Main.Toolbar1.Buttons(11).Enabled = True
    Main.Toolbar1.Buttons(12).Enabled = True
    Main.Toolbar1.Buttons(18).Enabled = True
    Main.Toolbar1.Buttons(7).Enabled = True

    Main.TableSelector(ObjectIndex).Visible = True
    
    Main.TableSelector(ObjectIndex).DragMode = 1
    Main.TableLabel(ObjectIndex).Enabled = False
    
    'disable columns
    For i = 1 To Main.TableColumnName.UBound
        If Main.TableColumnName(i).Tag = ObjectIndex Then
            Main.TableColumnName(i).Enabled = False
            Main.TableColumnType(i).Enabled = False
        End If
    Next i
    
    Main.Selector.Visible = True
    Main.Selector.Left = Main.Table(ObjectIndex).Left - 0.5
    Main.Selector.Top = Main.Table(ObjectIndex).Top - 0.5
    Main.Selector.Width = Main.Table(ObjectIndex).Width + 1.3
    Main.Selector.Height = Main.Table(ObjectIndex).Height + 1.3
    Main.Selector.ZOrder (0)

    Main.Combo1.Enabled = False

    'select table in combo
    Main.Combo1.Text = Main.TableLabel(ObjectIndex).Caption
    
    'select layer in combo
    Main.Combo3.Text = Main.Table(ObjectIndex).Tag

End Function

'delete table
Function DeleteTable()

    If MsgBox("Are you sure?", vbQuestion + vbYesNo, "Delete table") = vbYes Then
        HideTable (ObjectIndex)
        Main.Table(ObjectIndex).Tag = ""
        Main.Combo1.RemoveItem Main.Combo1.ListIndex
        Main.Combo1.Text = "-- None --"
        DeselectTables
    End If

End Function

'hide table
Function HideTable(TableIndex As Integer)

    Main.Table(TableIndex).Visible = False
    Main.TableShadow(TableIndex).Visible = False
    Main.TableLabel(TableIndex).Visible = False
    Main.TableSelector(TableIndex).Visible = False
    
    'hide columns
    For i = 1 To Main.TableColumnName.UBound
        If Main.TableColumnName(i).Tag = TableIndex Then
            Main.TableColumnName(i).Visible = False
            Main.TableColumnType(i).Visible = False
        End If
    Next i

End Function


'show table
Function ShowTable(Index As Integer)

    Main.Table(Index).Visible = True
    Main.TableShadow(Index).Visible = True
    Main.TableLabel(Index).Visible = True
    Main.TableSelector(Index).Visible = True

    'show columns
    For i = 1 To Main.TableColumnName.UBound
        If Main.TableColumnName(i).Tag = Index Then
            Main.TableColumnName(i).Visible = True
            Main.TableColumnType(i).Visible = True
        End If
    Next i

End Function

'show relation
Function ShowRelation(Index As Integer)

    Main.RelationLabel(Index).Visible = True
    Main.RelationLeft(Index).Visible = True
    Main.RelationRight(Index).Visible = True
    Main.RelationVertical(Index).Visible = True
    Main.RelationLeftType(Index).Visible = True
    Main.RelationRightType(Index).Visible = True

End Function

'hide relation
Function HideRelation(Index As Integer)

    Main.RelationLabel(Index).Visible = False
    Main.RelationLeft(Index).Visible = False
    Main.RelationRight(Index).Visible = False
    Main.RelationVertical(Index).Visible = False
    Main.RelationLeftType(Index).Visible = False
    Main.RelationRightType(Index).Visible = False

End Function

'create new columns
Function AddColumn(Name As Variant, ColumnType As Variant, TableId As Integer, CreateId As Integer)

    Dim NewWidth As Integer

    Load Main.TableColumnName(Main.TableColumnName.Count)
    Load Main.TableColumnType(Main.TableColumnType.Count)

    With Main.TableColumnName(Main.TableColumnName.UBound)
        .Left = Main.Table(TableId).Left + (ColumnMarge * 2)
        If CreateId = 0 Then
            .Top = Main.Table(TableId).Top + Main.TableLabel(TableId).Height + ColumnMarge
        Else
            .Top = Main.TableColumnName(Main.TableColumnName.UBound - 1).Top + Main.TableColumnName(Main.TableColumnName.UBound - 1).Height + ColumnMarge
        End If
        .Width = Main.Table(TableId).Width - Main.TableColumnType(0).Width - (ColumnMarge * 4)
        .Tag = TableId
        .Caption = Name
        .ZOrder (0)
        .Visible = True
    End With

    With Main.TableColumnType(Main.TableColumnType.UBound)
        .Left = Main.Table(TableId).Left + Main.Table(TableId).Width - Main.TableColumnType(Main.TableColumnType.UBound).Width - (ColumnMarge * 2)
        .Top = Main.TableColumnName(Main.TableColumnName.UBound).Top
        .Tag = TableId
        .Caption = ColumnType
        .ZOrder (0)
        .Visible = True
    End With
    
    Main.TableSelector(TableId).ZOrder (0)

End Function

'send table to back
Function SendTableBack()

    Main.TableSelector(ObjectIndex).ZOrder (1)

    'bring all columns to front
    For i = 1 To Main.TableColumnName.UBound
        'if column is related to current table
        If Main.TableColumnName(i).Tag = ObjectIndex Then
            Main.TableColumnName(i).ZOrder (1)
            Main.TableColumnType(i).ZOrder (1)
        End If
    Next i
    
    OldTableIndex = Main.Combo1.ListIndex
    OldTableCaption = Main.Combo1.Text
    Main.Combo1.RemoveItem OldTableIndex
    Main.Combo1.AddItem OldTableCaption, 1
    
    Main.TableLabel(ObjectIndex).ZOrder (1)
    Main.Table(ObjectIndex).ZOrder (1)
    Main.TableShadow(ObjectIndex).ZOrder (1)

End Function

'bring table to front
Function BringTableFront()

    Main.TableShadow(ObjectIndex).ZOrder (0)
    Main.Table(ObjectIndex).ZOrder (0)
    Main.TableLabel(ObjectIndex).ZOrder (0)
    
    'bring all columns to front
    For i = 1 To Main.TableColumnName.UBound
        'if column is related to current table
        If Main.TableColumnName(i).Tag = ObjectIndex Then
            Main.TableColumnName(i).ZOrder (0)
            Main.TableColumnType(i).ZOrder (0)
        End If
    Next i
    
    OldTableIndex = Main.Combo1.ListIndex
    OldTableCaption = Main.Combo1.Text
    Main.Combo1.RemoveItem OldTableIndex
    Main.Combo1.AddItem OldTableCaption, Main.Combo1.ListCount
    
    Main.TableSelector(ObjectIndex).ZOrder (0)

End Function

Function CreateRelation(Table1 As String, Table2 As String, Column1 As String, Column2 As String, LeftType As Integer, RightType As Integer, LayerName As String, Name As String)

    Main.Combo2.AddItem Name

    'get tempColumn1 value
    For j = 1 To Main.TableLabel.UBound
        If Main.TableLabel(j).Caption = Table1 Then
            For i = 1 To Main.TableColumnName.UBound
                If Main.TableColumnName(i).Caption = Column1 And Main.TableColumnName(i).Tag = Main.TableLabel(j).Index Then
                    tempColumn1 = i
                End If
            Next i
        End If
    Next j
    
    'get tempColumn2 value
    For j = 1 To Main.TableLabel.UBound
        If Main.TableLabel(j).Caption = Table2 Then
            For i = 1 To Main.TableColumnName.UBound
                If Main.TableColumnName(i).Caption = Column2 And Main.TableColumnName(i).Tag = Main.TableLabel(j).Index Then
                    tempColumn2 = i
                End If
            Next i
        End If
    Next j

    'load new objects
    Load Main.RelationLabel(Main.RelationLabel.Count)
    Load Main.RelationLeft(Main.RelationLeft.Count)
    Load Main.RelationRight(Main.RelationRight.Count)
    Load Main.RelationVertical(Main.RelationVertical.Count)
    Load Main.RelationLeftType(Main.RelationLeftType.Count)
    Load Main.RelationRightType(Main.RelationRightType.Count)
    
    'setup RelationLeft properties
    With Main.RelationLeft(Main.RelationLeft.UBound)
        .X1 = Main.TableColumnType(tempColumn1).Left + Main.TableColumnType(tempColumn1).Width
        .X2 = Main.TableColumnType(tempColumn1).Left + Main.TableColumnType(tempColumn1).Width + 3
        .Y1 = Main.TableColumnType(tempColumn1).Top + (Main.TableColumnType(tempColumn1).Height / 2)
        .Y2 = Main.TableColumnType(tempColumn1).Top + (Main.TableColumnType(tempColumn1).Height / 2)
        .Tag = Table1 & "|" & Column1
        .BorderStyle = ReadIniValue(App.Path & "\Setup.ini", "Properties", "RelationLineStyle") + 1
        .BorderColor = ReadIniValue(App.Path & "\Setup.ini", "Properties", "RelationLineColor")
        .ZOrder (0)
        .Visible = True
    End With
    
    'setup RelationRight properties
    With Main.RelationRight(Main.RelationRight.UBound)
        .X1 = Main.TableColumnType(tempColumn2).Left + Main.TableColumnType(tempColumn2).Width
        .X2 = Main.TableColumnType(tempColumn2).Left + Main.TableColumnType(tempColumn2).Width + 3
        .Y1 = Main.TableColumnType(tempColumn2).Top + (Main.TableColumnType(tempColumn2).Height / 2)
        .Y2 = Main.TableColumnType(tempColumn2).Top + (Main.TableColumnType(tempColumn2).Height / 2)
        .Tag = Table2 & "|" & Column2
        .BorderStyle = ReadIniValue(App.Path & "\Setup.ini", "Properties", "RelationLineStyle") + 1
        .BorderColor = ReadIniValue(App.Path & "\Setup.ini", "Properties", "RelationLineColor")
        .ZOrder (0)
        .Visible = True
    End With
    
    'setup RelationVertical properties
    With Main.RelationVertical(Main.RelationVertical.UBound)
        .X1 = Main.RelationLeft(Main.RelationLeft.UBound).X2
        .Y1 = Main.RelationLeft(Main.RelationLeft.UBound).Y1
        .X2 = Main.RelationRight(Main.RelationRight.UBound).X2
        .Y2 = Main.RelationRight(Main.RelationRight.UBound).Y1
        .BorderStyle = ReadIniValue(App.Path & "\Setup.ini", "Properties", "RelationLineStyle") + 1
        .BorderColor = ReadIniValue(App.Path & "\Setup.ini", "Properties", "RelationLineColor")
        .ZOrder (0)
        .Visible = True
    End With
    
    'setup RelationLabel properties
    With Main.RelationLabel(Main.RelationLabel.UBound)
        .Caption = Name
        .Tag = LayerName
        .Left = Main.RelationVertical(Main.RelationVertical.UBound).X1 + ((Main.RelationVertical(Main.RelationVertical.UBound).X2 - Main.RelationVertical(Main.RelationVertical.UBound).X1) / 2) - (.Width / 2)
        .Top = Main.RelationVertical(Main.RelationVertical.UBound).Y1 + ((Main.RelationVertical(Main.RelationVertical.UBound).Y2 - Main.RelationVertical(Main.RelationVertical.UBound).Y1) / 2) - (.Height / 2)
        .BackColor = ReadIniValue(App.Path & "\Setup.ini", "Properties", "RelationLabelColor")
        .ForeColor = ReadIniValue(App.Path & "\Setup.ini", "Properties", "RelationLabelFontColor")
        .ZOrder (0)
        .Visible = True
    End With
    
    'setup RelationLeftType properties
    With Main.RelationLeftType(Main.RelationLeftType.UBound)
        .Caption = LeftType
        .Top = Main.RelationLeft(Main.RelationLeft.UBound).Y1 - .Height
        .Left = Main.RelationLeft(Main.RelationLeft.UBound).X1 + 1.5
        .ZOrder (0)
        .Visible = True
    End With
    
    'setup RelationLeftType properties
    With Main.RelationRightType(Main.RelationRightType.UBound)
        .Caption = RightType
        .Top = Main.RelationRight(Main.RelationRight.UBound).Y1 - .Height
        .Left = Main.RelationRight(Main.RelationRight.UBound).X1 + 1.5
        .ZOrder (0)
        .Visible = True
    End With

End Function

'resize table
Function ResizeTable(Width As Integer, Height As Integer)

    LockWindow Main.hWnd, True
    Main.Table(ObjectIndex).Width = Width
    Main.Table(ObjectIndex).Height = Height
    Main.TableSelector(ObjectIndex).Width = Width
    Main.TableSelector(ObjectIndex).Height = Height
    Main.TableLabel(ObjectIndex).Width = Width
    Main.TableShadow(ObjectIndex).Width = Width
    Main.TableShadow(ObjectIndex).Height = Height
    Main.TableShadow(ObjectIndex).Left = Main.Table(ObjectIndex).Left + 0.3
    Main.TableShadow(ObjectIndex).Top = Main.Table(ObjectIndex).Top + 0.3
    
    'resize columns
    For i = 1 To Main.TableColumnName.UBound
        If Main.TableColumnName(i).Tag = ObjectIndex Then
            Main.TableColumnType(i).Left = Main.Table(ObjectIndex).Left + Main.Table(ObjectIndex).Width - Main.TableColumnType(i).Width - ColumnMarge
            Main.TableColumnName(i).Width = Main.Table(ObjectIndex).Width - Main.TableColumnType(i).Width - (ColumnMarge * 4)
        End If
    Next i
    
    LockWindow Main.hWnd, False

End Function

'create drawing
Function CreateDrawing(ScaleSize As Integer, ShowColumns As Boolean, ShowRelations As Boolean, ShowRelationLabels As Boolean)

    'remove old picture
    Main.Drawing.Picture = LoadPicture("")
    
    'set picture size
    Main.Drawing.Width = Main.Editor.Width * ScaleSize
    Main.Drawing.Height = Main.Editor.Height * ScaleSize

    'set background color
    Main.Drawing.BackColor = Main.Editor.BackColor

    'draw tables
    For k = 0 To Main.Combo1.ListCount - 1
        For i = 1 To Main.Table.UBound
            If Main.Table(i).Visible = True Then
                If Main.TableLabel(i).Caption = Main.Combo1.List(k) Then
                    'draw shadow
                    Main.Drawing.Line (Main.TableShadow(i).Left * ScaleSize, Main.TableShadow(i).Top * ScaleSize)-((Main.TableShadow(i).Width * ScaleSize) + (Main.TableShadow(i).Left * ScaleSize), (Main.TableShadow(i).Height * ScaleSize) + (Main.TableShadow(i).Top * ScaleSize)), vbBlack, BF
                    'draw fillcolor
                    Main.Drawing.Line (Main.Table(i).Left * ScaleSize, Main.Table(i).Top * ScaleSize)-((Main.Table(i).Width * ScaleSize) + (Main.Table(i).Left * ScaleSize), (Main.Table(i).Height * ScaleSize) + (Main.Table(i).Top * ScaleSize)), Main.Table(i).FillColor, BF
                    'draw table title box
                    Main.Drawing.Line (Main.Table(i).Left * ScaleSize, Main.Table(i).Top * ScaleSize)-((Main.Table(i).Width * ScaleSize) + (Main.Table(i).Left * ScaleSize), (Main.TableLabel(i).Height * ScaleSize) + (Main.TableShadow(i).Top * ScaleSize)), Main.TableLabel(i).BackColor, BF
                    'draw line around table
                    Main.Drawing.Line (Main.Table(i).Left * ScaleSize, Main.Table(i).Top * ScaleSize)-((Main.Table(i).Width * ScaleSize) + (Main.Table(i).Left * ScaleSize), (Main.Table(i).Height * ScaleSize) + (Main.Table(i).Top * ScaleSize)), vbBlack, B
                    'print tablename
                    Main.Drawing.FontName = Main.TableLabel(i).FontName
                    Main.Drawing.ForeColor = Main.TableLabel(i).ForeColor
                    Main.Drawing.FontSize = Main.TableLabel(i).FontSize * ScaleSize
                    Main.Drawing.CurrentX = (((Main.Table(i).Left * ScaleSize) + ((Main.Table(i).Width * ScaleSize)) / 2)) - ((Len(Main.TableLabel(i).Caption) * ScaleSize) / 2)
                    Main.Drawing.CurrentY = ((Main.Table(i).Top * ScaleSize) + (Main.TableLabel(i).Height * ScaleSize)) - ((Main.TableLabel(i).FontSize * ScaleSize) / 2)
                    Main.Drawing.Print Main.TableLabel(i).Caption
                    'create table fields
                    If ShowColumns = True Then
                        For j = 1 To Main.TableColumnName.UBound
                            If Main.TableColumnName(j).Tag = Main.Table(i).Index Then
                                'draw field if field is related to current table
                                Main.Drawing.Line ((Main.TableColumnName(j).Left * ScaleSize), (Main.TableColumnName(j).Top * ScaleSize))-(((Main.TableColumnName(j).Left * ScaleSize) + (Main.TableColumnName(j).Width * ScaleSize)), ((Main.TableColumnName(j).Top * ScaleSize) + (Main.TableColumnName(j).Height * ScaleSize))), vbWhite, BF
                                Main.Drawing.Line ((Main.TableColumnName(j).Left * ScaleSize), (Main.TableColumnName(j).Top * ScaleSize))-(((Main.TableColumnName(j).Left * ScaleSize) + (Main.TableColumnName(j).Width * ScaleSize)), ((Main.TableColumnName(j).Top * ScaleSize) + (Main.TableColumnName(j).Height * ScaleSize))), vbWhite, B
                                Main.Drawing.Line ((Main.TableColumnType(j).Left * ScaleSize), (Main.TableColumnType(j).Top * ScaleSize))-(((Main.TableColumnType(j).Left * ScaleSize) + (Main.TableColumnType(j).Width * ScaleSize)), ((Main.TableColumnType(j).Top * ScaleSize) + (Main.TableColumnType(j).Height * ScaleSize))), vbWhite, BF
                                Main.Drawing.Line ((Main.TableColumnType(j).Left * ScaleSize), (Main.TableColumnType(j).Top * ScaleSize))-(((Main.TableColumnType(j).Left * ScaleSize) + (Main.TableColumnType(j).Width * ScaleSize)), ((Main.TableColumnType(j).Top * ScaleSize) + (Main.TableColumnType(j).Height * ScaleSize))), vbWhite, B
                                'draw field name content
                                Main.Drawing.FontName = Main.TableColumnName(j).FontName
                                Main.Drawing.ForeColor = Main.TableColumnName(j).ForeColor
                                Main.Drawing.FontSize = Main.TableColumnName(j).FontSize * ScaleSize
                                Main.Drawing.CurrentX = Main.TableColumnName(j).Left * ScaleSize
                                Main.Drawing.CurrentY = Main.TableColumnName(j).Top * ScaleSize
                                Main.Drawing.Print Main.TableColumnName(j).Caption
                                'draw field type content
                                Main.Drawing.FontName = Main.TableColumnType(j).FontName
                                Main.Drawing.ForeColor = Main.TableColumnType(j).ForeColor
                                Main.Drawing.FontSize = Main.TableColumnType(j).FontSize * ScaleSize
                                Main.Drawing.CurrentX = Main.TableColumnType(j).Left * ScaleSize
                                Main.Drawing.CurrentY = Main.TableColumnType(j).Top * ScaleSize
                                Main.Drawing.Print Main.TableColumnType(j).Caption
                            End If
                        Next j
                    End If
                End If
            End If
        Next i
    Next k
    
    'draw relations
    If ShowRelations = True Then
        For i = 1 To Main.RelationLabel.UBound
            If Main.RelationLabel(i).Visible = True Then
                'draw relation left line
                Main.Drawing.DrawStyle = Main.RelationLeft(i).BorderStyle - 1
                Main.Drawing.DrawWidth = Main.RelationLeft(i).BorderWidth * ScaleSize
                Main.Drawing.Line ((Main.RelationLeft(i).X1 * ScaleSize), (Main.RelationLeft(i).Y1 * ScaleSize))-((Main.RelationLeft(i).X2 * ScaleSize), (Main.RelationLeft(i).Y2 * ScaleSize)), Main.RelationLeft(i).BorderColor
                'draw relation right line
                Main.Drawing.DrawStyle = Main.RelationRight(i).BorderStyle - 1
                Main.Drawing.DrawWidth = Main.RelationRight(i).BorderWidth * ScaleSize
                Main.Drawing.Line ((Main.RelationRight(i).X1 * ScaleSize), (Main.RelationRight(i).Y1 * ScaleSize))-((Main.RelationRight(i).X2 * ScaleSize), (Main.RelationRight(i).Y2 * ScaleSize)), Main.RelationRight(i).BorderColor
                'draw relation vertical line
                Main.Drawing.DrawStyle = Main.RelationVertical(i).BorderStyle - 1
                Main.Drawing.DrawWidth = Main.RelationVertical(i).BorderWidth * ScaleSize
                Main.Drawing.Line ((Main.RelationVertical(i).X1 * ScaleSize), (Main.RelationVertical(i).Y1 * ScaleSize))-((Main.RelationVertical(i).X2 * ScaleSize), (Main.RelationVertical(i).Y2 * ScaleSize)), Main.RelationVertical(i).BorderColor
    
                If ShowRelationLabels = True Then
                    'draw label
                    Main.Drawing.DrawStyle = 0
                    Main.Drawing.DrawWidth = 1
                    Main.Drawing.Line ((Main.RelationLabel(i).Left * ScaleSize), (Main.RelationLabel(i).Top * ScaleSize))-(((Main.RelationLabel(i).Left * ScaleSize) + (Main.RelationLabel(i).Width * ScaleSize)), ((Main.RelationLabel(i).Top * ScaleSize) + (Main.RelationLabel(i).Height * ScaleSize))), Main.RelationLabel(i).BackColor, BF
                    Main.Drawing.Line ((Main.RelationLabel(i).Left * ScaleSize), (Main.RelationLabel(i).Top * ScaleSize))-(((Main.RelationLabel(i).Left * ScaleSize) + (Main.RelationLabel(i).Width * ScaleSize)), ((Main.RelationLabel(i).Top * ScaleSize) + (Main.RelationLabel(i).Height * ScaleSize))), vbBlack, B
                    'draw relation name
                    Main.Drawing.FontName = Main.RelationLabel(i).FontName
                    Main.Drawing.ForeColor = Main.RelationLabel(i).ForeColor
                    Main.Drawing.FontSize = Main.RelationLabel(i).FontSize * ScaleSize
                    Main.Drawing.CurrentX = Main.RelationLabel(i).Left * ScaleSize
                    Main.Drawing.CurrentY = Main.RelationLabel(i).Top * ScaleSize
                    Main.Drawing.Print Main.RelationLabel(i).Caption
                End If
            End If
        Next i
    End If

    Set Main.Drawing.Picture = Main.Drawing.Image

End Function

Function PasteTable()

    CreateTable Main.TableLabel(CopyIndex).Caption & "_" & Main.Table.Count, Main.Table(CopyIndex).Width, Main.Table(CopyIndex).Height, Main.Table(CopyIndex).Left + GridSize, Main.Table(CopyIndex).Top + GridSize, Main.Combo3.Text, Main.Table(CopyIndex).FillColor
    
    'check if table contains columns
    CreateId = 0
    For i = 1 To Main.TableColumnName.UBound
        If Main.TableColumnName(i).Tag = CopyIndex Then
            AddColumn Main.TableColumnName(i).Caption, Main.TableColumnType(i).Caption, Main.Table.UBound, CreateId
            CreateId = CreateId + 1
        End If
    Next i

End Function
