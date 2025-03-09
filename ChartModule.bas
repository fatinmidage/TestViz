'******************************************
' 模块: ChartModule
' 用途: 处理和绘制数据图表相关的功能
' 说明: 本模块主要负责绘制两种类型的图表：
'      1. 容量保持率随循环圈数变化的散点图
'      2. DCIR和DCIR Rise随循环圈数变化的散点图
'******************************************
Option Explicit

'图表尺寸和位置相关常量
Private Const CHART_WIDTH As Long = 450    '图表总宽度（磅）
Private Const CHART_HEIGHT As Long = 300   '图表总高度（磅）
Private Const CHART_GAP As Long = 20       '图表之间的垂直间距（磅）
Private Const CHART_TOTAL_SPACING As Long = 40  '图表总间距（CHART_HEIGHT + CHART_GAP）

'绘图区尺寸和位置相关常量（相对于图表的百分比）
Private Const PLOT_WIDTH As Long = 370     '绘图区宽度（约为图表宽度的82%）
Private Const PLOT_HEIGHT As Long = 215    '绘图区高度（约为图表高度的72%）
Private Const PLOT_LEFT As Long = 55       '绘图区左边距（约为图表宽度的12%）
Private Const PLOT_TOP As Long = 30        '绘图区顶部边距（约为图表高度的10%）

'图表颜色常量（使用RGB值）
Private Const COLOR_435 As Long = &HC07000     '435系列电池曲线颜色（蓝色，RGB: 0,112,192）
Private Const COLOR_450 As Long = &HC0FF&      '450系列电池曲线颜色（黄色，RGB: 255,192,0）
Private Const COLOR_GRIDLINE As Long = &HBFBFBF '网格线颜色（浅灰色，RGB: 191,191,191）

'******************************************
' 函数: CreateDataCharts
' 用途: 创建所有数据图表的主函数
' 参数:
'   - ws: 目标工作表对象
'   - nextRow: 图表开始绘制的行号
'   - reportName: 报告标题
'   - commonConfig: 公共配置信息，包含电池名称等数据
'   - zpTables: 中检数据表格集合，包含所有电池的中检数据
'   - cycleDataTables: 循环数据表格集合，包含所有电池的循环数据
' 返回: Long，最后一个图表底部的行号
' 说明: 此函数负责创建所有图表，并返回最后图表之后的行号
'******************************************
Public Function CreateDataCharts(ByVal ws As Worksheet, _
                               ByVal nextRow As Long, _
                               ByVal reportName As String, _
                               ByVal commonConfig As Collection, _
                               ByVal zpTables As Collection, _
                               ByVal cycleDataTables As Collection) As Long
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
        
    '创建图表标题行
    With ws.Cells(nextRow, 2)
        .value = "2.测试数据图表:"
        .Font.Bold = True
        .Font.Name = "微软雅黑"
        .Font.Size = 10
    End With
    
    '修改 CreateDataCharts 中的调用
    nextRow = nextRow + 2  '标题行后空一行
    
    '创建容量和能量保持率图表
    CreateCapacityEnergyChart ws, nextRow, reportName, cycleDataTables, commonConfig
    
    '计算下一个图表的起始位置
    nextRow = nextRow + CHART_TOTAL_SPACING
    
    '创建DCR增长率图表
    CreateDCRRiseChart ws, nextRow, zpTables, commonConfig(3), "容量保持率"

    '计算下一个图表的起始位置
    nextRow = nextRow + CHART_TOTAL_SPACING/2
    
    '创建DCR增长率图表
    CreateDCRRiseChart ws, nextRow, zpTables, commonConfig(3), "能量保持率"

    
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    CreateDataCharts = nextRow
    Exit Function
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    LogError "CreateDataCharts", Err.Description
    MsgBox "创建图表时出错: " & Err.Description, vbExclamation
    CreateDataCharts = nextRow
End Function

'******************************************
' 函数: AddDataSeriesToChart
' 用途: 为图表添加电池数据系列
' 参数:
'   - cht: 图表对象
'   - ws: 工作表对象
'   - cycleDataTables: 循环数据表格集合
'******************************************
Private Sub AddDataSeriesToChart(ByVal cht As chart, _
                                ByVal ws As Worksheet, _
                                ByVal cycleDataTables As Collection, _
                                ByVal dataColumnName As String, _
                                ByVal commonConfig As Collection)
    
    '为每个电池添加数据系列
    Dim batteryIndex As Long
    For batteryIndex = 1 To cycleDataTables.count
        '获取当前电池的数据表格
        Dim cycleDataTable As ListObject
        Set cycleDataTable = cycleDataTables(batteryIndex)
        
        '获取电池名称（从表格上方的单元格）
        Dim batteryName As String
        batteryName = ws.Range(cycleDataTable.Range.Cells(1, 1).Address).End(xlUp).value
        
        '添加数据系列并设置格式
        With cht.SeriesCollection.NewSeries
            .XValues = cycleDataTable.ListColumns("循环圈数").DataBodyRange
            .Values = cycleDataTable.ListColumns(dataColumnName).DataBodyRange
            .Name = batteryName
            .markerStyle = xlMarkerStyleNone  '不显示数据点标记
            .Format.Line.Weight = 1.5
            
            '使用commonConfig中的颜色值设置曲线颜色
            .Format.Line.ForeColor.RGB = commonConfig(3)(batteryIndex)
        End With
    Next batteryIndex
End Sub
  
'******************************************
' 函数: SetupChartAxes
' 用途: 设置图表的X轴和Y轴属性
' 参数:
'   - cht: 图表对象
'******************************************
Private Sub SetupChartAxes(ByVal cht As chart, ByVal yAxisTitle As String)
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
        .MajorGridlines.Format.Line.Visible = msoTrue
        .MajorGridlines.Format.Line.ForeColor.RGB = COLOR_GRIDLINE
        .MajorGridlines.Format.Line.Weight = 0.25
    End With
    
    '设置Y轴属性
    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = yAxisTitle
        .AxisTitle.Font.Name = "Times New Roman"
        .AxisTitle.Font.Size = 10
        .AxisTitle.Font.Bold = True
        .MinimumScale = 0.7      '最小70%
        .MaximumScale = 1        '最大100%
        .majorUnit = 0.05        '主刻度间隔5%
        .TickLabels.Font.Name = "Times New Roman"
        .TickLabels.Font.Bold = True
        .TickLabels.numberFormat = "0%"
        .MajorTickMark = xlTickMarkInside
        .MajorGridlines.Format.Line.Visible = msoTrue
        .MajorGridlines.Format.Line.ForeColor.RGB = COLOR_GRIDLINE
        .MajorGridlines.Format.Line.Weight = 0.25
    End With
End Sub
        
'******************************************
' 函数: CreateCapacityRetentionChart
' 用途: 创建容量保持率图表对象并设置基本属性
' 参数:
'   - ws: 工作表对象
'   - topRow: 图表顶部所在行号
'   - reportName: 报告标题
'   - cycleDataTables: 循环数据表格集合
'   - dataColumnName: 数据列名称
'   - yAxisTitle: Y轴标题
' 返回: ChartObject，创建好的图表对象
'******************************************
Private Function CreateCapacityRetentionChart(ByVal ws As Worksheet, _
                                            ByVal topRow As Long, _
                                            ByVal reportName As String, _
                                            ByVal cycleDataTables As Collection, _
                                            ByVal dataColumnName As String, _
                                            ByVal yAxisTitle As String, _
                                            ByVal commonConfig As Collection) As chartObject
    
    '创建图表对象并设置基本属性
    Dim chartObj As chartObject
    Set chartObj = ws.ChartObjects.Add(Left:=ws.Cells(topRow, 3).Left, _
                                     width:=CHART_WIDTH, _
                                     Top:=ws.Cells(topRow, 2).Top, _
                                     height:=CHART_HEIGHT)
    
    With chartObj.chart
        .chartType = xlXYScatterLines  '设置为散点图（带平滑线）
        
        '添加数据系列
        AddDataSeriesToChart chartObj.chart, ws, cycleDataTables, dataColumnName, commonConfig
        
        '设置网格线和标题
        SetupChartGridlines chartObj.chart
        SetupChartTitle chartObj.chart, reportName
        
        '设置坐标轴属性
        SetupChartAxes chartObj.chart, yAxisTitle
        
        '设置图例属性
        SetupChartLegend .legend
        
        '设置绘图区属性
        SetupPlotArea .plotArea
    End With
    
    Set CreateCapacityRetentionChart = chartObj
End Function

'******************************************
' 函数: CreateCapacityEnergyChart
' 用途: 创建容量保持率随循环圈数变化的散点图
' 参数:
'   - ws: 目标工作表对象
'   - topRow: 图表顶部所在行号
'   - reportName: 报告标题，用于图表标题
'   - cycleDataTables: 包含所有电池循环数据的表格集合
' 说明: 此函数创建一个散点图，显示所有电池的容量保持率变化趋势
'       - X轴显示循环圈数（0-1000）
'       - Y轴显示容量保持率（70%-100%）
'       - 不同型号电池使用不同颜色曲线
'       - 图例位于图表右上角，半透明显示
'******************************************
Private Sub CreateCapacityEnergyChart(ByVal ws As Worksheet, _
                                    ByVal topRow As Long, _
                                    ByVal reportName As String, _
                                    ByVal cycleDataTables As Collection, _
                                    ByVal commonConfig As Collection)
    
    '创建容量保持率图表
    CreateCapacityRetentionChart ws, topRow, reportName, cycleDataTables, "容量保持率", "Capacity Retention", commonConfig

    '创建能量保持率图表
    CreateCapacityRetentionChart ws, topRow + 20, reportName, cycleDataTables, "能量保持率", "Energy Retention", commonConfig
End Sub

'******************************************
' 函数: SetupPlotArea
' 用途: 设置图表的绘图区属性
' 参数:
'   - plotArea: 绘图区对象
'   - Optional plotWidth: 绘图区宽度，默认为 PLOT_WIDTH
'******************************************
Private Sub SetupPlotArea(ByVal plotArea As plotArea, _
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
' 函数: SetupChartGridlines
' 用途: 设置图表的网格线格式
' 参数:
'   - cht: 图表对象
'******************************************
Private Sub SetupChartGridlines(ByVal cht As chart)
    With cht.Axes(xlValue).MajorGridlines.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = COLOR_GRIDLINE
        .Weight = 0.25
    End With
End Sub

'******************************************
' 函数: SetupChartTitle
' 用途: 设置图表标题
' 参数:
'   - cht: 图表对象
'   - titleText: 标题文本
'******************************************
Private Sub SetupChartTitle(ByVal cht As chart, ByVal titleText As String)
    With cht
        .HasTitle = True
        .ChartTitle.Text = titleText
        .ChartTitle.Font.Size = 14
        .ChartTitle.Font.Name = "Times New Roman"
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
Private Sub SetupChartLegend(ByVal legend As legend, _
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
' 过程: LogError
' 用途: 记录错误信息到即时窗口（Debug Window）
' 参数:
'   - functionName: 发生错误的函数名
'   - errorDescription: 错误描述信息
' 说明: 用于调试和错误追踪，输出格式：
'       当前时间 - 函数名 error: 错误描述
'******************************************
Private Sub LogError(ByVal functionName As String, ByVal errorDescription As String)
    Debug.Print Now & " - " & functionName & " error: " & errorDescription
End Sub

'******************************************
' 函数: AddZPDataSeriesToChart
' 用途: 为中检图表添加电池数据系列
' 参数:
'   - cht: 图表对象
'   - ws: 工作表对象
'   - zpTables: 中检数据表格集合
'   - colorCollection: 颜色集合
'   - fieldName: 数据列名称
'******************************************
Private Sub AddZPDataSeriesToChart(ByVal cht As chart, _
                                 ByVal ws As Worksheet, _
                                 ByVal zpTables As Collection, _
                                 ByVal colorCollection As Collection, _
                                 ByVal fieldName As String)
    
    '遍历每个电池的中检数据表
    Dim batteryIndex As Long
    
    For batteryIndex = 1 To zpTables.count
        Dim shouldSkip As Boolean
        shouldSkip = False  ' 每次循环重置标志
        
        Dim batteryTables As Collection
        Set batteryTables = zpTables(batteryIndex)
        
        '获取中检容量保持率表
        Dim zpCapacityTable As ListObject
        Set zpCapacityTable = batteryTables(1)
        
        '判断zpCapacityTable是否为空
        If zpCapacityTable.ListRows.count = 0 Then
            ' 如果为空，跳过当前电池
            shouldSkip = True
        End If

        If Not shouldSkip Then
            '添加数据系列
            With cht.SeriesCollection.NewSeries
                .XValues = zpCapacityTable.ListColumns("循环圈数").DataBodyRange
                '根据字段名称选择不同的数据列
                If fieldName = "容量保持率" Then
                    .Values = zpCapacityTable.ListColumns("容量保持率").DataBodyRange
                ElseIf fieldName = "能量保持率" Then
                    .Values = zpCapacityTable.ListColumns("能量保持率").DataBodyRange
                End If
                .Name = ws.Cells(zpCapacityTable.Range.row - 1, zpCapacityTable.Range.column).value
                .markerStyle = xlMarkerStyleCircle  ' 设置圆形标记
                .Format.Line.Weight = 1.5
                .Format.Line.ForeColor.RGB = colorCollection(batteryIndex)
                .MarkerSize = 4                     ' 设置标记大小
                .MarkerForegroundColor = colorCollection(batteryIndex)  ' 标记边框颜色
                .MarkerBackgroundColor = RGB(255, 255, 255)  ' 标记填充白色
            End With
        End If
    Next batteryIndex
End Sub

'******************************************
' 函数: CreateDCRRiseChart
' 用途: 创建DCR增长率随循环圈数变化的散点图
' 参数:
'   - ws: 目标工作表对象
'   - topRow: 图表顶部所在行号
'   - zpTables: 中检数据表格集合
'   - colorCollection: 颜色集合
'   - fieldName: 数据列名称（"容量保持率"或"能量保持率"）
'******************************************
Private Sub CreateDCRRiseChart(ByVal ws As Worksheet, _
                             ByVal topRow As Long, _
                             ByVal zpTables As Collection, _
                             ByVal colorCollection As Collection, _
                             ByVal fieldName As String)

    '创建图表对象并设置基本属性
    Dim chartObj As chartObject
    Set chartObj = ws.ChartObjects.Add(Left:=ws.Cells(topRow, 3).Left, _
                                     width:=CHART_WIDTH, _
                                     Top:=ws.Cells(topRow, 2).Top, _
                                     height:=CHART_HEIGHT)
    With chartObj.chart
        .chartType = xlXYScatterLines
        
        '添加主坐标轴数据系列
        AddZPDataSeriesToChart chartObj.chart, ws, zpTables, colorCollection, fieldName
        
        '添加次坐标轴数据系列
        AddDCIRDataSeriesToChart chartObj.chart, ws, zpTables, colorCollection
        
        '设置次坐标轴属性
        With .Axes(xlValue, xlSecondary)
            .HasTitle = True
            .AxisTitle.Text = "DCIR increase rate"
            .AxisTitle.Font.Name = "Times New Roman"
            .AxisTitle.Font.Size = 10
            .AxisTitle.Font.Bold = True
            .MinimumScale = -0.1        '从0%开始
            .MaximumScale = 1.5       '最大100%
            .majorUnit = 0.2         '主刻度间隔10%
            .TickLabels.Font.Name = "Times New Roman"
            .TickLabels.Font.Bold = True
            .TickLabels.numberFormat = "0%"
            .MajorTickMark = xlTickMarkInside
        End With
        
        '设置网格线和标题
        SetupChartGridlines chartObj.chart
        SetupChartTitle chartObj.chart, "ZP of Cycle"
        
        '设置坐标轴属性
        SetupChartAxes chartObj.chart, IIf(fieldName = "容量保持率", "Residual Capacity", "Residual Energy")
        
        '设置图例属性
        SetupChartLegend .legend, 330
        
        '设置绘图区属性
        SetupPlotArea .plotArea, 330
    End With
End Sub

'******************************************
' 函数: AddDCIRDataSeriesToChart
' 用途: 为图表添加DCIR数据系列（次坐标轴）
' 参数:
'   - cht: 图表对象
'   - ws: 工作表对象
'   - zpTables: 中检数据表格集合
'   - colorCollection: 颜色集合
'******************************************
Private Sub AddDCIRDataSeriesToChart(ByVal cht As chart, _
                                    ByVal ws As Worksheet, _
                                    ByVal zpTables As Collection, _
                                    ByVal colorCollection As Collection)
    
    '遍历每个电池的中检数据表
    Dim batteryIndex As Long
    For batteryIndex = 1 To zpTables.count
        Dim batteryTables As Collection
        Set batteryTables = zpTables(batteryIndex)
        
        '获取DCIR表格
        Dim dcirTable As ListObject
        Set dcirTable = batteryTables(3)
        '判断dcirTable是否为空
        If dcirTable.ListRows.count = 0 Then
           '如果为空，跳过当前电池
            GoTo NextBattery
        End If
        
        '添加数据系列
        With cht.SeriesCollection.NewSeries
            .XValues = batteryTables(1).ListColumns("循环圈数").DataBodyRange
            .Values = dcirTable.ListColumns("50%").DataBodyRange
            .Name = ws.Cells(batteryTables(1).Range.row - 1, batteryTables(1).Range.column).value
            .markerStyle = xlMarkerStyleNone
            .Format.Line.Weight = 1.5
            .markerStyle = xlMarkerStyleCircle
            .MarkerSize = 4
            .MarkerForegroundColor = colorCollection(batteryIndex)
            .MarkerBackgroundColor = RGB(255, 255, 255)
            .Format.Line.ForeColor.RGB = colorCollection(batteryIndex)
            .Format.Line.DashStyle = msoLineDash  '设置为标准虚线
            .AxisGroup = xlSecondary
        End With
        
NextBattery: ' 跳转标签
    Next batteryIndex
End Sub



