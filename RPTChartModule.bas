Option Explicit

'图表尺寸和位置相关常量
Private Const CHART_WIDTH As Long = 450    '图表总宽度（磅）
Private Const CHART_HEIGHT As Long = 300   '图表总高度（磅）

'绘图区尺寸和位置相关常量（相对于图表的百分比）
Private Const PLOT_WIDTH As Long = 370     '绘图区宽度（约为图表宽度的82%）
Private Const PLOT_HEIGHT As Long = 215    '绘图区高度（约为图表高度的72%）
Private Const PLOT_LEFT As Long = 55       '绘图区左边距（约为图表宽度的12%）
Private Const PLOT_TOP As Long = 30        '绘图区顶部边距（约为图表高度的10%）

'图表颜色常量（使用RGB值）
Private Const COLOR_GRIDLINE As Long = &HBFBFBF '网格线颜色（浅灰色，RGB: 191,191,191）

' =====================
' 中检恢复率和DCR增长率散点图创建函数
' =====================
' 功能：创建中检恢复率和DCR增长率散点图
' 参数：
'   ws - 目标工作表对象
'   xRng - X轴数据范围（循环圈数）
'   rptRngs - 中检恢复率数据范围集合
'   dcrRngs - DCR增长率数据范围集合
'   yAxisTitle - Y轴标题
'   batteriesInfoCollection - 电池信息集合
'   leftParam - 左边距
'   topParam - 顶部边距
' 返回值：创建的图表对象，失败返回Nothing
' =====================
Public Function CreateRPTRetentionChart(ByVal ws As Worksheet, _
                                      ByVal xRng As Range, _
                                      ByVal rptRngs As Collection, _
                                      ByVal dcrRngs As Collection, _
                                      ByVal yAxisTitle As String, _
                                      ByVal batteriesInfoCollection As Collection, _
                                      ByVal leftParam As Long, _
                                      ByVal topParam As Long) As ChartObject
    On Error GoTo ErrorHandler
    
    ' 创建散点图
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=leftParam, Top:=topParam, Width:=CHART_WIDTH, Height:=CHART_HEIGHT)
    
    With chartObj.Chart
        .ChartType = xlXYScatterSmooth
        
        ' 添加中检恢复率数据系列（主坐标轴）
        Dim i As Long
        For i = 1 To rptRngs.Count
            With .SeriesCollection.NewSeries
                .XValues = xRng
                .Values = rptRngs(i)
                .Name = IIf(Len(Trim(batteriesInfoCollection(i).BatteryName)) = 0, "Cell #" & i, batteriesInfoCollection(i).BatteryName)
                .MarkerStyle = xlMarkerStyleCircle
                .MarkerSize = 5
                .Format.Line.Weight = 1.25
                .MarkerForegroundColor = batteriesInfoCollection(i).BatteryColor
                .MarkerBackgroundColor = vbWhite
                .Format.Line.ForeColor.RGB = batteriesInfoCollection(i).BatteryColor
            End With
        Next i
        
        ' 添加DCR增长率数据系列（次坐标轴）
        For i = 1 To dcrRngs.Count
            With .SeriesCollection.NewSeries
                .XValues = xRng
                .Values = dcrRngs(i)
                .Name = IIf(Len(Trim(batteriesInfoCollection(i).BatteryName)) = 0, "Cell #" & i, batteriesInfoCollection(i).BatteryName)
                .MarkerStyle = xlMarkerStyleCircle
                .MarkerSize = 5
                .Format.Line.DashStyle = msoLineDash
                .Format.Line.Weight = 1.25
                .MarkerForegroundColor = batteriesInfoCollection(i).BatteryColor
                .MarkerBackgroundColor = vbWhite
                .Format.Line.ForeColor.RGB = batteriesInfoCollection(i).BatteryColor
                .AxisGroup = xlSecondary
            End With
        Next i

        ' 设置网格线
        SetupRPTChartGridlines chartObj.Chart

        ' 设置坐标轴属性
        SetupRPTChartAxes chartObj.Chart, yAxisTitle

        ' 设置图表标题
        SetupRPTChartTitle chartObj.Chart, yAxisTitle

        ' 设置图例属性
        SetupRPTChartLegend .Legend

        ' 设置绘图区属性
        SetupRPTPlotArea .PlotArea
        
    End With
    
    Set CreateRPTRetentionChart = chartObj
    Exit Function

ErrorHandler:
    MsgBox "创建图表时发生错误: " & Err.Description, vbCritical
    Set CreateRPTRetentionChart = Nothing
End Function

'******************************************
' 函数: SetupRPTChartGridlines
' 用途: 设置中检恢复率图表的网格线格式
' 参数:
'   - cht: 图表对象
'******************************************
Private Sub SetupRPTChartGridlines(ByVal cht As Chart)
    With cht.Axes(xlValue).MajorGridlines.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = COLOR_GRIDLINE
        .Weight = 0.25
    End With
End Sub

'******************************************
' 函数: SetupRPTChartAxes
' 用途: 设置中检恢复率图表的X轴和Y轴属性
' 参数:
'   - cht: 图表对象
'******************************************
Private Sub SetupRPTChartAxes(ByVal cht As Chart, ByVal yAxisTitle As String)
    '设置X轴属性
    With cht.Axes(xlCategory, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = "Cycle Number(N)"
        .AxisTitle.Font.Name = "Times New Roman"
        .AxisTitle.Font.Size = 10
        .AxisTitle.Font.Bold = True
        .MinimumScale = 0        '从0开始
        .MaximumScale = 1000     '最大1000圈
        .MajorUnit = 100         '主刻度间隔100圈
        .TickLabels.Font.Name = "Times New Roman"
        .TickLabels.Font.Bold = True
        .TickMarkSpacing = 1     '设置刻度间隔
        .MajorTickMark = xlTickMarkInside  '主刻度线向内
        .MinorTickMark = xlTickMarkNone    '不显示次刻度线
        '设置主网格线
        .MajorGridlines.Format.Line.Visible = msoTrue
        .MajorGridlines.Format.Line.ForeColor.RGB = COLOR_GRIDLINE
        .MajorGridlines.Format.Line.Weight = 0.25
    End With
    
    '设置主Y轴属性（中检恢复率）
    With cht.Axes(xlValue)
        .HasTitle = True
        dim yAxisTitleValue as string
        yAxisTitleValue = IIf(yAxisTitle = "容量保持率/%", "Residual Capacity", "Residual Energy")
        .AxisTitle.Text = yAxisTitleValue
        .AxisTitle.Font.Name = "Times New Roman"
        .AxisTitle.Font.Size = 10
        .AxisTitle.Font.Bold = True
        .MinimumScale = 70       '最小70%
        .MaximumScale = 100      '最大100%
        .MajorUnit = 5           '主刻度间隔5%
        .TickLabels.Font.Name = "Times New Roman"
        .TickLabels.Font.Bold = True
        .TickLabels.NumberFormat = "0""%""" '显示为整数加百分号
        .MajorTickMark = xlTickMarkInside  '主刻度线向内
        .MinorTickMark = xlTickMarkNone    '不显示次刻度线
    End With
    
    '设置次Y轴属性（DCR增长率）
    With cht.Axes(xlValue, xlSecondary)
        .HasTitle = True
        .AxisTitle.Text = "DCIR increase rate"
        .AxisTitle.Font.Name = "Times New Roman"
        .AxisTitle.Font.Size = 10
        .AxisTitle.Font.Bold = True
        .MinimumScale = -10        '最小-10%
        .MaximumScale = 150      '最大150%
        .MajorUnit = 20          '主刻度间隔20%
        .TickLabels.Font.Name = "Times New Roman"
        .TickLabels.Font.Bold = True
        .TickLabels.NumberFormat = "0""%""" '显示为整数加百分号
        .MajorTickMark = xlTickMarkInside  '主刻度线向内
        .MinorTickMark = xlTickMarkNone    '不显示次刻度线
        .HasMajorGridlines = False
    End With
End Sub

'******************************************
' 函数: SetupRPTChartLegend
' 用途: 设置中检恢复率图表的图例属性
' 参数:
'   - legend: 图例对象
'******************************************
Private Sub SetupRPTChartLegend(ByVal legend As Legend)
    With legend
        '设置图例位置
        .Position = xlLegendPositionRight
        .Left = PLOT_LEFT + PLOT_WIDTH - .Width - 30  '设置左侧位置（紧贴绘图区右边）
        .Top = PLOT_TOP + PLOT_HEIGHT * 0.01    '设置顶部位置（略微偏离绘图区顶部）
        
        '设置字体属性
        .Font.Name = "Times New Roman"
        .Font.Size = 10
        .Format.TextFrame2.TextRange.Font.Size = 9
        
        '设置图例背景和边框
        With .Format.Fill
            .Visible = msoFalse  '不显示背景填充
        End With
        .Format.Line.Visible = msoFalse  '不显示边框
    End With
End Sub

'******************************************
' 函数: SetupRPTPlotArea
' 用途: 设置中检恢复率图表的绘图区属性
' 参数:
'   - plotArea: 绘图区对象
'******************************************
Private Sub SetupRPTPlotArea(ByVal plotArea As PlotArea)
    With plotArea
        .Format.Line.Visible = msoTrue
        .Format.Line.ForeColor.RGB = COLOR_GRIDLINE
        .Format.Line.Weight = 0.25
        .InsideWidth = PLOT_WIDTH - 30      '设置绘图区内部宽度
        .InsideHeight = PLOT_HEIGHT    '设置绘图区内部高度
        .InsideLeft = PLOT_LEFT        '设置绘图区左边距
        .InsideTop = PLOT_TOP          '设置绘图区顶部边距
    End With
End Sub

'******************************************
' 函数: SetupRPTChartTitle
' 用途: 设置中检恢复率图表标题
' 参数:
'   - cht: 图表对象
'   - yAxisTitle: Y轴标题文本
'******************************************
Private Sub SetupRPTChartTitle(ByVal cht As Chart, ByVal yAxisTitle As String)
    With cht
        .HasTitle = True
        .ChartTitle.Text = "ZP of Cycles"
        .ChartTitle.Font.Size = 14
        .ChartTitle.Font.Name = "Times New Roman"
    End With
End Sub