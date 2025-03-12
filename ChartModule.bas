Option Explicit

'图表尺寸和位置相关常量
Public Const CHART_WIDTH As Long = 450    '图表总宽度（磅）
Public Const CHART_HEIGHT As Long = 300   '图表总高度（磅）

'绘图区尺寸和位置相关常量（相对于图表的百分比）
Public Const PLOT_WIDTH As Long = 370     '绘图区宽度（约为图表宽度的82%）
Public Const PLOT_HEIGHT As Long = 215    '绘图区高度（约为图表高度的72%）
Public Const PLOT_LEFT As Long = 55       '绘图区左边距（约为图表宽度的12%）
Public Const PLOT_TOP As Long = 30        '绘图区顶部边距（约为图表高度的10%）

'图表颜色常量（使用RGB值）
Public Const COLOR_435 As Long = &HC07000     '435系列电池曲线颜色（蓝色，RGB: 0,112,192）
Public Const COLOR_450 As Long = &HC0FF&      '450系列电池曲线颜色（黄色，RGB: 255,192,0）
Public Const COLOR_GRIDLINE As Long = &HBFBFBF '网格线颜色（浅灰色，RGB: 191,191,191）

' =====================
' 散点图创建函数
' =====================
' 功能：创建电芯容量保持率散点图
' 参数：
'   ws - 目标工作表对象
'   xRng - X轴数据范围（循环圈数）
'   yRngs - Y轴数据范围集合（容量保持率）
' 返回值：创建的图表对象，失败返回Nothing
' =====================
Public Function CreateCapacityRetentionChart(ByVal ws As Worksheet, ByVal xRng As Range, ByVal yRngs As Collection, ByVal yAxisTitle As String, ByVal reportName As String, ByVal batteriesInfoCollection As Collection, ByVal leftParam As Long, ByVal topParam As Long, Optional ByVal dcrYRngs As Collection = Nothing) As ChartObject
    On Error GoTo ErrorHandler
    
    ' 创建散点图
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=leftParam, Top:=topParam, Width:=CHART_WIDTH, Height:=CHART_HEIGHT)
    
    With chartObj.Chart
        .ChartType = xlXYScatterSmooth
        
        ' 添加数据系列
        Dim i As Long
        For i = 1 To yRngs.Count
            With .SeriesCollection.NewSeries
                .XValues = xRng
                .Values = yRngs(i)
                .Name = IIf(Len(Trim(batteriesInfoCollection(i).BatteryName)) = 0, "Cell #" & i, batteriesInfoCollection(i).BatteryName)
                .MarkerStyle = xlMarkerStyleNone
                .Format.Line.ForeColor.RGB = batteriesInfoCollection(i).BatteryColor
            End With
        Next i
        
        ' 添加DCR增长率数据系列（如果有）
        If Not dcrYRngs Is Nothing Then
            For i = 1 To dcrYRngs.Count
                With .SeriesCollection.NewSeries
                    .XValues = xRng
                    .Values = dcrYRngs(i)
                    .Name = "DCR #" & i
                    .MarkerStyle = xlMarkerStyleNone
                    .Format.Line.ForeColor.RGB = batteriesInfoCollection(i).BatteryColor
                    .AxisGroup = xlSecondary
                End With
            Next i
            
            ' 设置次坐标轴属性
            With .Axes(xlValue, xlSecondary)
                .HasTitle = True
                .AxisTitle.Text = "DCR增长率/%"
                .HasMajorGridlines = False
                .HasMinorGridlines = False
            End With
        End If

        SetupChartGridlines chartObj.Chart

        '设置坐标轴属性
        SetupChartAxes chartObj.chart, yAxisTitle
        SetupChartTitle chartObj.chart, reportName

        '设置图例属性
        SetupChartLegend .legend

        '设置绘图区属性
        SetupPlotArea .plotArea
        
    End With
    
    Set CreateCapacityRetentionChart = chartObj
    Exit Function

ErrorHandler:
    MsgBox "创建图表时发生错误: " & Err.Description, vbCritical
    Set CreateCapacityRetentionChart = Nothing
End Function

'******************************************
' 函数: SetupChartGridlines
' 用途: 设置图表的网格线格式
' 参数:
'   - cht: 图表对象
'******************************************
Public Sub SetupChartGridlines(ByVal cht As chart)
    With cht.Axes(xlValue).MajorGridlines.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = COLOR_GRIDLINE
        .Weight = 0.25
    End With
End Sub

'******************************************
' 函数: SetupChartAxes
' 用途: 设置图表的X轴和Y轴属性
' 参数:
'   - cht: 图表对象
'******************************************
Public Sub SetupChartAxes(ByVal cht As chart, ByVal yAxisTitle As String)
    '设置X轴属性
    With cht.Axes(xlCategory, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = "Cycle Number(N)"
        .AxisTitle.Font.Name = "Times New Roman"
        .AxisTitle.Font.Size = 10
        .AxisTitle.Font.Bold = True
        .MinimumScale = 0        '从0开始
        .MaximumScale = 1000     '最大1000圈
        .majorUnit = 100         '主刻度间隔100圈
        .TickLabels.Font.Name = "Times New Roman"
        .TickLabels.Font.Bold = True
        .TickMarkSpacing = 1     '设置刻度间隔
        .MajorTickMark = xlTickMarkInside  '主刻度线向内
        .MinorTickMark = xlTickMarkNone    '不显示次刻度线
        '设置主网格线的可见性为可见
        .MajorGridlines.Format.Line.Visible = msoTrue
        '设置主网格线的颜色为预定义的浅灰色（COLOR_GRIDLINE）
        .MajorGridlines.Format.Line.ForeColor.RGB = COLOR_GRIDLINE
        '设置主网格线的线条粗细为0.25磅
        .MajorGridlines.Format.Line.Weight = 0.25
    End With
    
    '设置Y轴属性
    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = yAxisTitle
        .AxisTitle.Font.Name = "Times New Roman"
        .AxisTitle.Font.Size = 10
        .AxisTitle.Font.Bold = True
        .MinimumScale = 70       '最小70%
        .MaximumScale = 100      '最大100%
        .majorUnit = 5           '主刻度间隔5%
        .TickLabels.Font.Name = "Times New Roman"
        .TickLabels.Font.Bold = True
        .TickLabels.NumberFormat = "0""%""" '显示为整数加百分号
        .MajorTickMark = xlTickMarkInside  '主刻度线向内
        .MinorTickMark = xlTickMarkNone    '不显示次刻度线
        
    End With
End Sub
        
'******************************************
' 函数: SetupChartLegend
' 用途: 设置图表的图例属性
' 参数:
'   - legend: 图例对象
'   - Optional plotWidth: 绘图区宽度，默认为 PLOT_WIDTH
' 说明: 此函数设置图例的位置、字体和外观
'       - 位置：右侧，靠近绘图区右上角
'       - 字体：Times New Roman，10号字
'       - 外观：无边框，无背景填充
'******************************************
Public Sub SetupChartLegend(ByVal legend As legend, _
                           Optional ByVal plotWidth As Long = PLOT_WIDTH)
    With legend
        '设置图例位置
        .position = xlLegendPositionRight
        .Left = PLOT_LEFT + plotWidth - .width  '设置左侧位置（紧贴绘图区右边）
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
' 函数: SetupPlotArea
' 用途: 设置图表的绘图区属性
' 参数:
'   - plotArea: 绘图区对象
'   - Optional plotWidth: 绘图区宽度，默认为 PLOT_WIDTH
'******************************************
Public Sub SetupPlotArea(ByVal plotArea As plotArea, _
                         Optional ByVal plotWidth As Long = PLOT_WIDTH)
    With plotArea
        .Format.Line.Visible = msoTrue
        .Format.Line.ForeColor.RGB = COLOR_GRIDLINE
        .Format.Line.Weight = 0.25
        .InsideWidth = plotWidth     '设置绘图区内部宽度
        .InsideHeight = PLOT_HEIGHT   '设置绘图区内部高度
        .InsideLeft = PLOT_LEFT       '设置绘图区左边距
        .InsideTop = PLOT_TOP         '设置绘图区顶部边距
    End With
End Sub

'******************************************
' 函数: SetupChartTitle
' 用途: 设置图表标题
' 参数:
'   - cht: 图表对象
'   - titleText: 标题文本
'******************************************
Public Sub SetupChartTitle(ByVal cht As chart, ByVal titleText As String)
    With cht
        .HasTitle = True
        .ChartTitle.Text = titleText
        .ChartTitle.Font.Size = 14
        .ChartTitle.Font.Name = "Times New Roman"
    End With
End Sub