Sub MuiliPlot()
    Dim rChart As Range
    Dim cht1 As Chart
    Dim StrtRow As Long
    Dim TitleCell As Range
    Dim HFLCell As Range

    With Worksheets("DataSheet")
        StrtRow = 2
        For Each X In .Range(.Cells(StrtRow, "F"), .Cells(.Range("F:F").SpecialCells(xlCellTypeLastCell).Row, "F"))
            If X.Value = "" Then
                Set TitleCell = .Cells(StrtRow, "E")
                Set HFLCell = .Cells(StrtRow, "H")

                Set rChart = .Range(.Cells(StrtRow, "F"), .Cells(X.Offset(-1, 0).Row, "H"))

                Set cht1 = .ChartObjects.Add(Left:=Cells(StrtRow, 14).Left, Width:=Worksheets("DataSheet").Range("K4").Value * 72, Top:=Cells(StrtRow, 14).Top, Height:=Worksheets("DataSheet").Range("K5").Value * 72).Chart

                With cht1
                    .Parent.Top = Cells(StrtRow, 14).Top
                    .Parent.Left = Cells(StrtRow, 14).Left
                    .ChartType = xlXYScatterSmooth
                    .HasLegend = True
                    .Legend.Position = xlLegendPositionTop
                    .SetSourceData rChart
                    .HasTitle = True
                    .ChartTitle.Text = Worksheets("DataSheet").Range("K1").Value & " " & TitleCell.Value
                    .ChartTitle.Format.TextFrame2.TextRange.Font.Name = "Times New Roman"
                    .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 14
                    .ChartTitle.Format.TextFrame2.TextRange.Font.Bold = msoTrue

                    For Each s In .SeriesCollection
                        s.MarkerStyle = xlNone
                    Next s

                    .Axes(xlCategory, xlPrimary).HasTitle = True
                    .Axes(xlCategory, xlPrimary).AxisTitle.Text = Worksheets("DataSheet").Range("K2").Value
                    .Axes(xlCategory, xlPrimary).AxisTitle.Format.TextFrame2.TextRange.Font.Name = "Times New Roman"
                    .Axes(xlCategory, xlPrimary).AxisTitle.Format.TextFrame2.TextRange.Font.Size = 10
                    .Axes(xlCategory, xlPrimary).AxisTitle.Format.TextFrame2.TextRange.Font.Bold = msoFalse

                    ' Set the minimum scale of X-axis to zero
                    .Axes(xlCategory, xlPrimary).MinimumScaleIsAuto = False
                    .Axes(xlCategory, xlPrimary).MinimumScale = 0
                    ' Hide negative values on the x-axis
                    .Axes(xlCategory, xlPrimary).TickLabelPosition = xlTickLabelPositionLow

                    .Axes(xlValue, xlPrimary).HasTitle = True
                    .Axes(xlValue, xlPrimary).AxisTitle.Text = Worksheets("DataSheet").Range("K3").Value
                    .Axes(xlValue, xlPrimary).AxisTitle.Format.TextFrame2.TextRange.Font.Name = "Times New Roman"
                    .Axes(xlValue, xlPrimary).AxisTitle.Format.TextFrame2.TextRange.Font.Size = 10
                    .Axes(xlValue, xlPrimary).AxisTitle.Format.TextFrame2.TextRange.Font.Bold = msoFalse

                    ' Set the "Dash Type" for gridlines to "Round Dot"
                    .Axes(xlValue, xlPrimary).MajorGridlines.Format.Line.DashStyle = msoLineRoundDot

                End With

                ' Rename series in the first chart with custom names
                On Error Resume Next
                cht1.SeriesCollection(1).Name = "Bed Level"
                cht1.SeriesCollection(2).Name = "HFL " & HFLCell.Value & " (mMSL)"
                On Error GoTo 0

                If Err.Number <> 0 Then
                    MsgBox "Error renaming series: " & Err.Description, vbExclamation, "Error"
                    Err.Clear
                End If

                StrtRow = X.Offset(1, 0).Row
            End If
        Next X
    End With
End Sub

