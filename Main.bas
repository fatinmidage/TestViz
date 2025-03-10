Option Explicit

' 在模块顶部添加常量定义
Private Const SHEET_NAME_CYCLE_LIFE As String = "Cycle Life"
Private Const SHEET_NAME_RPT_CYCLE_LIFE As String = "RPT of Cycle Life"
Private Const TABLE_NAME_FILES As String = "文件名表"
Private Const COL_NAME_FILENAME As String = "文件名"
Private Const COL_NAME_CAPACITY_RETENTION As String = "容量保持率/%"
Private Const COL_NAME_ENERGY_RETENTION As String = "能量保持率/%"
Private Const ERR_TABLE_NOT_FOUND As Long = 1002
Private Const ERR_INVALID_DATA_FORMAT As Long = 1007

Public Sub Main()
    On Error GoTo ErrorHandler
    SetPerformanceMode True
    
    ' 声明变量
    Dim originDataWorkbook As Workbook
    Dim CycleLifeSheet As Worksheet
    Dim capacityRetentionCol As Long
    Dim cellCount As Long
    
    ' 打开数据工作簿
    Set originDataWorkbook = OpenWorkbookFromTable()
    If originDataWorkbook Is Nothing Then
        GoTo ExitSub
    End If
    
    ' 获取并验证Cycle Life工作表
    Set CycleLifeSheet = originDataWorkbook.Worksheets(SHEET_NAME_CYCLE_LIFE)
    If CycleLifeSheet Is Nothing Then
        Err.Raise ERR_INVALID_DATA_FORMAT, "Main", "无法获取" & SHEET_NAME_CYCLE_LIFE & "工作表"
    End If

    ' 获取容量保持率列
    capacityRetentionCol = FindColumnByTitle(CycleLifeSheet, COL_NAME_CAPACITY_RETENTION)
    If capacityRetentionCol = 0 Then
        Err.Raise ERR_INVALID_DATA_FORMAT, "Main", "无法找到" & COL_NAME_CAPACITY_RETENTION & "列"
    End If
    
    ' 获取电芯数量
    cellCount = GetCellCount(CycleLifeSheet, capacityRetentionCol)
    If cellCount = 0 Then
        Err.Raise ERR_INVALID_DATA_FORMAT, "Main", "无法获取电芯数量"
    End If
    
    Dim newWorksheet As Worksheet
    Set newWorksheet = CreateWorksheet(originDataWorkbook)

    Dim maxCyclesCount As Long
    maxCyclesCount = 0
    
    ' 遍历每一列，找出最大的有效数据行数
    Dim currentCol As Long
    For currentCol = capacityRetentionCol To capacityRetentionCol + cellCount - 1
        Dim currentLastRow As Long
        currentLastRow = CycleLifeSheet.Cells(CycleLifeSheet.Rows.Count, currentCol).End(xlUp).Row - 3 ' 减去表头行数
        If currentLastRow > maxCyclesCount Then
            maxCyclesCount = currentLastRow
        End If
    Next currentCol

    Dim xRng As Range
    Set xRng = CycleLifeSheet.Range(CycleLifeSheet.Cells(4, 1), CycleLifeSheet.Cells(maxCyclesCount + 3, 1))
    Dim yRngs As collection
    Set yRngs = New collection

    ' 遍历每一列，添加容量保持率数据范围
    For currentCol = capacityRetentionCol To capacityRetentionCol + cellCount - 1
        Dim yRng As Range
        Set yRng = CycleLifeSheet.Range(CycleLifeSheet.Cells(4, currentCol), CycleLifeSheet.Cells(maxCyclesCount + 3, currentCol))
        yRngs.Add yRng
    Next currentCol

    Dim chartObj As ChartObject
    Set chartObj = CreateCapacityRetentionChart(newWorksheet, xRng, yRngs)
    If chartObj Is Nothing Then
        GoTo ExitSub
    End If

    ' 正常退出
ExitSub:
    SetPerformanceMode False
    Exit Sub

ErrorHandler:
    MsgBox "错误: " & Err.Description, vbCritical
    Resume ExitSub
End Sub

' =====================
' 性能优化设置函数
' =====================
' 功能：控制Excel应用程序的性能优化设置
' 参数：
'   enable - True表示启用性能优化，False表示关闭性能优化
' =====================
Private Sub SetPerformanceMode(ByVal enable As Boolean)
    With Application
        .ScreenUpdating = Not enable
        .DisplayAlerts = Not enable
        .Calculation = IIf(enable, xlCalculationManual, xlCalculationAutomatic)
        .EnableEvents = Not enable
        If enable Then
            .StatusBar = "正在处理数据..."
        Else
            .StatusBar = False
        End If
    End With
End Sub

' =====================
' 工作簿打开函数
' =====================
' 功能：从指定的表格中获取文件名并打开对应的工作簿
' 参数：无
' 返回值：
'   - 成功时返回打开的工作簿对象
'   - 失败时返回Nothing
' 错误处理：
'   - 如果未找到文件名表，抛出ERR_TABLE_NOT_FOUND错误
'   - 如果文件不存在，显示错误消息并返回Nothing
' =====================
Private Function OpenWorkbookFromTable() As Workbook
    On Error GoTo ErrorHandler
    
    ' 获取文件名表
    Dim tblFileNames As ListObject
    Dim atws As Worksheet
    Set atws = ActiveSheet
    
    ' 参数验证
    ' 检查当前工作表是否可用
    If atws Is Nothing Then
        Err.Raise ERR_INVALID_DATA_FORMAT, "OpenWorkbookFromTable", "无法获取当前工作表"
    End If
    
    ' 获取并验证文件名表是否存在
    Set tblFileNames = atws.ListObjects(TABLE_NAME_FILES)
    If tblFileNames Is Nothing Then
        Err.Raise ERR_TABLE_NOT_FOUND, "OpenWorkbookFromTable", "未找到'" & TABLE_NAME_FILES & "'表格"
    End If
    
    ' 验证文件名列是否存在且有数据
    ' 获取文件名列对象
    Dim fileNameColumn As ListColumn
    Set fileNameColumn = tblFileNames.ListColumns(COL_NAME_FILENAME)
    If fileNameColumn Is Nothing Then
        Err.Raise ERR_INVALID_DATA_FORMAT, "OpenWorkbookFromTable", "未找到'" & COL_NAME_FILENAME & "'列"
    End If
    
    ' 检查文件名列是否包含数据
    If fileNameColumn.DataBodyRange Is Nothing Then
        Err.Raise ERR_INVALID_DATA_FORMAT, "OpenWorkbookFromTable", "文件名列没有数据"
    End If
    
    ' 获取文件名并处理文件路径
    ' 从第一行获取文件名
    Dim fileName As String
    fileName = fileNameColumn.DataBodyRange(1).Value
    
    ' 验证文件名是否为空
    ' 移除文件名前后的空格并检查长度
    If Len(Trim(fileName)) = 0 Then
        Err.Raise ERR_INVALID_DATA_FORMAT, "OpenWorkbookFromTable", "文件名不能为空"
    End If
    
    ' 处理文件扩展名
    ' 如果文件名不包含.xlsx或.xls扩展名，则添加.xlsx
    If InStr(1, fileName, ".xlsx", vbTextCompare) = 0 And InStr(1, fileName, ".xls", vbTextCompare) = 0 Then
        fileName = fileName & ".xlsx"
    End If
    
    ' 构建完整的文件路径
    ' 使用当前工作簿所在目录作为基准路径
    Dim filePath As String
    filePath = ThisWorkbook.Path & Application.PathSeparator & fileName
    
    ' 验证文件是否存在
    ' 如果文件不存在，显示错误消息并退出
    If Not FileExists(filePath) Then
        MsgBox "找不到文件: " & fileName, vbExclamation
        Exit Function
    End If
    
    ' 打开文件
    ' 使用Workbooks.Open方法打开工作簿
    Set OpenWorkbookFromTable = Workbooks.Open(filePath)
    Exit Function

ErrorHandler:
    ' 错误处理
    ' 记录错误信息到即时窗口并显示错误消息
    Debug.Print "错误发生在OpenWorkbookFromTable函数中: " & Err.Description
    MsgBox "错误: " & Err.Description, vbCritical
    Set OpenWorkbookFromTable = Nothing
End Function

' =====================
' 列标题查找函数
' =====================
' 功能：在工作表中查找指定列标题的列号
' 参数：
'   ws - 目标工作表对象
'   columnTitle - 要查找的列标题
' 返回值：找到的列号，如果未找到则返回0
' 错误处理：
'   - 如果工作表对象为空，抛出ERR_INVALID_DATA_FORMAT错误
'   - 如果列标题为空，抛出ERR_INVALID_DATA_FORMAT错误
' =====================
Private Function FindColumnByTitle(ws As Worksheet, ByVal columnTitle As String) As Long
    On Error GoTo ErrorHandler
    
    ' 参数验证
    ' 检查工作表对象是否为空，确保传入了有效的工作表引用
    If ws Is Nothing Then
        Err.Raise ERR_INVALID_DATA_FORMAT, "FindColumnByTitle", "工作表对象不能为空"
    End If
    ' 检查列标题是否为空字符串，移除前后空格后判断
    If Len(Trim(columnTitle)) = 0 Then
        Err.Raise ERR_INVALID_DATA_FORMAT, "FindColumnByTitle", "列标题不能为空"
    End If
    
    ' 使用Find函数快速查找列标题
    ' Find函数参数说明：
    ' What - 要查找的文本（列标题）
    ' LookIn - 在单元格值中查找（xlValues）
    ' LookAt - 完全匹配（xlWhole）
    ' SearchOrder - 按列搜索（xlByColumns）
    ' MatchCase - 不区分大小写（False）
    Dim foundCell As Range
    Set foundCell = ws.Rows(1).Find(What:=columnTitle, _
                                  LookIn:=xlValues, _
                                  LookAt:=xlWhole, _
                                  SearchOrder:=xlByColumns, _
                                  MatchCase:=False)
    
    ' 返回值处理：
    ' - 如果找到匹配的单元格，返回其列号
    ' - 如果未找到匹配，返回0表示未找到
    FindColumnByTitle = IIf(Not foundCell Is Nothing, foundCell.Column, 0)
    Exit Function

ErrorHandler:
    ' 错误处理：
    ' - 将返回值设置为0，表示查找失败
    ' - 显示错误消息，包含具体的错误描述
    FindColumnByTitle = 0
    MsgBox "查找列标题时发生错误: " & Err.Description, vbExclamation
End Function

' =====================
' 电芯数量获取函数
' =====================
' 功能：获取指定列中电芯的数量
' 参数：
'   ws - 目标工作表对象
'   targetCol - 目标列号
' 返回值：
'   - 如果是合并单元格，返回合并区域的列数
'   - 如果是单个单元格，返回1
'   - 发生错误时返回0
' 错误处理：
'   - 如果工作表对象为空，抛出ERR_INVALID_DATA_FORMAT错误
'   - 如果列号无效，抛出ERR_INVALID_DATA_FORMAT错误
' =====================
Private Function GetCellCount(ws As Worksheet, ByVal targetCol As Long) As Long
    On Error GoTo ErrorHandler
    
    ' 验证工作表对象是否为空
    If ws Is Nothing Then
        Err.Raise ERR_INVALID_DATA_FORMAT, "GetCellCount", "工作表对象为空"
    End If
    
    ' 验证列号是否在有效范围内（大于0且不超过工作表最大列数）
    If targetCol <= 0 Or targetCol > ws.Columns.Count Then
        Err.Raise ERR_INVALID_DATA_FORMAT, "GetCellCount", "无效的列号"
    End If
    
    ' 使用With语句优化对目标单元格的访问
    With ws.Cells(1, targetCol)
        ' 判断目标单元格是否为合并单元格
        If .MergeCells Then
            ' 如果是合并单元格，返回合并区域的列数
            GetCellCount = .MergeArea.Columns.Count
        Else
            ' 如果不是合并单元格，返回1表示单个单元格
            GetCellCount = 1
        End If
    End With
    ' 正常退出函数
    Exit Function
    
ErrorHandler:
    GetCellCount = 0 ' 发生错误时返回0
    Err.Raise Err.Number, "GetCellCount", "获取电芯数量失败: " & Err.Description
End Function

' =====================
' 工作表创建函数
' =====================
' 功能：创建新的工作表并关闭网格线
' 参数：
'   wb - 工作簿对象
' 返回值：新创建的工作表对象，失败返回Nothing
Private Function CreateWorksheet(ByVal wb As Workbook) As Worksheet
    On Error GoTo ErrorHandler
    
    ' 添加新工作表
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.activate
    
    ' 关闭网格线
    Application.ActiveWindow.DisplayGridlines = False
    
    ' 返回工作表对象
    Set CreateWorksheet = ws
    Exit Function

ErrorHandler:
    Call HandleError(Err.Number, "创建工作表时发生错误")
    Set CreateWorksheet = Nothing
End Function

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
Private Function CreateCapacityRetentionChart(ByVal ws As Worksheet, ByVal xRng As Range, ByVal yRngs As Collection) As ChartObject
    On Error GoTo ErrorHandler
    
    ' 创建散点图
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=50, Top:=50, Width:=600, Height:=400)
    
    With chartObj.Chart
        .ChartType = xlXYScatterSmooth
        
        ' 添加数据系列
        Dim i As Long
        For i = 1 To yRngs.Count
            With .SeriesCollection.NewSeries
                .XValues = xRng
                .Values = yRngs(i)
                .Name = "电芯" & i
            End With
        Next i
        
        ' 设置图表标题
        .HasTitle = True
        .ChartTitle.Text = "电芯容量保持率变化趋势"
        
        ' 设置X轴标题
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "循环圈数"
        End With
        
        ' 设置Y轴标题
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "容量保持率 (%)"
            .MinimumScale = 0
            .MaximumScale = 100
        End With
        
        ' 设置图例
        With .Legend
            .Position = xlLegendPositionRight
            .Format.Fill.Transparency = 0.2
        End With
    End With
    
    Set CreateCapacityRetentionChart = chartObj
    Exit Function

ErrorHandler:
    MsgBox "创建图表时发生错误: " & Err.Description, vbCritical
    Set CreateCapacityRetentionChart = Nothing
End Function

