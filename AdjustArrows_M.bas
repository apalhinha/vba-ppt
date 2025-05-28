Attribute VB_Name = "AdjustArrows_M"
Option Explicit

Public Sub AdjustArrows()
Dim sld As PowerPoint.Slide
Dim shp As PowerPoint.Shape
    Set sld = Application.ActiveWindow.View.Slide
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        For Each shp In ActiveWindow.Selection.ShapeRange
            If shp.Type = msoFreeform Then
                If doAdjustArrows(sld, shp) Then
                    shp.Delete
                Else
                    Exit For
                End If
            End If
        Next shp
    End If
End Sub

Private Function doAdjustArrows(sld As PowerPoint.Slide, shp As PowerPoint.Shape) As Boolean
Dim sn As PowerPoint.ShapeNodes, node As PowerPoint.ShapeNode
Dim nodeX As Single, nodeY As Single
Dim iNode As Integer, iSerie As Integer
Dim aSeries() As Variant ' Array of 2 dimentions:
                         '[1=X, 2=Y, 3=Position-within-Nodes] + index.

' SegmentType=0 is msoSegmentLine. SegmentType=1 is msoSegmentCurve.
    Set sn = shp.Nodes
    ReDim aSeries(1 To 3, 0 To 0)
    iNode = 1: iSerie = 0
    'ListShapeNodesInfo shp
    While iNode <= sn.Count
        Set node = shp.Nodes(iNode)
        iSerie = iSerie + 1
        ReDim Preserve aSeries(1 To 3, 0 To iSerie)
        aSeries(1, iSerie) = CDbl(node.Points(1, 1))
        aSeries(2, iSerie) = CDbl(node.Points(1, 2))
        aSeries(3, iSerie) = iNode
        nodeX = node.Points(1, 1)
        nodeY = node.Points(1, 2)
        iNode = iNode + 1
        If iNode <= sn.Count Then
            Set node = shp.Nodes(iNode)
            If node.SegmentType = msoSegmentCurve Then iNode = iNode + 2
        End If
    Wend
    For iSerie = 1 To UBound(aSeries, 2)
        'Debug.Print iSerie & ": " & aSeries(1, iSerie) & ", " & aSeries(2, iSerie) & " (idx " & aSeries(3, iSerie) & ")"
    Next iSerie
    doAdjustArrows = CreateBezierLine(sld, aSeries)
End Function

Function CreateBezierLine(sld As PowerPoint.Slide, aSeries() As Variant) As Boolean
Dim i As Long
Dim dx As Double, dy As Double
Dim ffb As PowerPoint.FreeformBuilder, shp As PowerPoint.Shape
Dim xcur As Double, ycur As Double
Dim xdelta As Double, ydelta As Double
Const DEFDELTA = 10
Dim errmsg As String, previous As String

    Set ffb = sld.Shapes.BuildFreeform(msoEditingCorner, aSeries(1, 1), aSeries(2, 1))
    xcur = aSeries(1, 1): ycur = aSeries(2, 1)
    previous = "none"
    With ffb
        For i = 1 To UBound(aSeries, 2) - 2
            dx = aSeries(1, i + 1) - aSeries(1, i)
            dy = aSeries(2, i + 1) - aSeries(2, i)
            If Abs(dx) > Abs(dy) Then
                'Debug.Print "horizontal"
                If previous = "horizontal" Then errmsg = "horizontal"
                previous = "horizontal"
                xdelta = DEFDELTA
                If dx < 0 Then xdelta = -xdelta
                ydelta = DEFDELTA ' Assume next vertical goes DOWN
                If aSeries(2, i + 2) < aSeries(2, i + 1) Then ydelta = -ydelta  ' or UP
                xcur = aSeries(1, i + 1)
                .AddNodes msoSegmentLine, msoEditingCorner, xcur - xdelta, ycur   'msoEditingCorner, xcur - 2 * xdelta, ycur, xcur, ycur, xcur - xdelta, ycur
                .AddNodes msoSegmentCurve, msoEditingCorner, xcur, ycur, xcur, ycur, xcur, ycur + ydelta
            Else
                'Debug.Print "vertical"
                If previous = "vertical" Then errmsg = "vertical"
                previous = "vertical"
                ydelta = DEFDELTA
                If dy < 0 Then ydelta = -ydelta
                xdelta = DEFDELTA
                If aSeries(1, i + 2) < aSeries(1, i + 1) Then xdelta = -xdelta
                ycur = aSeries(2, i + 1)
                .AddNodes msoSegmentLine, msoEditingCorner, xcur, ycur - ydelta  ' msoEditingCorner, xcur, ycur - 2 * ydelta, xcur, ycur, xcur, ycur - ydelta
                .AddNodes msoSegmentCurve, msoEditingCorner, xcur, ycur, xcur, ycur, xcur + xdelta, ycur
            End If
        Next i
        dx = aSeries(1, i + 1) - aSeries(1, i)
        dy = aSeries(2, i + 1) - aSeries(2, i)
        If Abs(dx) > Abs(dy) Then
            'Debug.Print "horizontal"
            xcur = aSeries(1, i + 1)
            .AddNodes msoSegmentLine, msoEditingCorner, xcur, ycur
        Else
            'Debug.Print "vertical"
            ycur = aSeries(2, i + 1)
            .AddNodes msoSegmentLine, msoEditingCorner, xcur, ycur
        End If
    End With
    
    If errmsg <> "" Then
        MsgBox "Multiple consecutive " & errmsg & " lines." & vbCrLf & "Exiting...", vbCritical
        CreateBezierLine = False
    Else
        Set shp = ffb.ConvertToShape
        With shp.Line
            .ForeColor.RGB = RGB(166, 166, 166)
            .Weight = 1.5
            .BeginArrowheadLength = msoArrowheadShort
            .BeginArrowheadStyle = msoArrowheadOval
            .BeginArrowheadWidth = msoArrowheadNarrow
            .EndArrowheadLength = msoArrowheadShort 'msoArrowheadLengthMedium
            .EndArrowheadStyle = msoArrowheadTriangle
            .EndArrowheadWidth = msoArrowheadNarrow 'msoArrowheadWidthMedium
        End With
        CreateBezierLine = True
    End If
    Set ffb = Nothing
    Exit Function
End Function
