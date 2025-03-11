Option Explicit

' =====================
' 容量保持率数据处理函数
' =====================
' 功能：处理容量保持率数据并创建图表
' 参数：
'   CycleLifeSheet - Cycle Life工作表对象
'   newWorksheet - 用于创建图表的新工作表对象
'   columnTitle - 容量保持率列标题
'   reportTitle - 报告标题
' 返回值：处理成功返回True，失败返回False
' =====================
Public Function ProcessCapacityRetentionData(ByVal CycleLifeSheet As Worksheet, ByVal newWorksheet As Worksheet, ByVal columnTitle As String, ByVal reportTitle As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' 获取容量保持率列
    Dim capacityRetentionCol As Long
    capacityRetentionCol = FindColumnByTitle(CycleLifeSheet, columnTitle)
    If capacityRetentionCol = 0 Then
        Err.Raise ERR_INVALID_DATA_FORMAT, "ProcessCapacityRetentionData", "无法找到" & columnTitle & "列"
    End If
    
    ' 获取电芯数量
    Dim cellCount As Long
    cellCount = GetCellCount(CycleLifeSheet, capacityRetentionCol)
    If cellCount = 0 Then
        Err.Raise ERR_INVALID_DATA_FORMAT, "ProcessCapacityRetentionData", "无法获取电芯数量"
    End If
    
    ' 计算最大循环次数
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

    ' 准备数据范围
    Dim xRng As Range
    Set xRng = CycleLifeSheet.Range(CycleLifeSheet.Cells(4, 1), CycleLifeSheet.Cells(maxCyclesCount + 3, 1))
    Dim yRngs As Collection
    Set yRngs = New Collection

    ' 遍历每一列，添加容量保持率数据范围
    For currentCol = capacityRetentionCol To capacityRetentionCol + cellCount - 1
        Dim yRng As Range
        Set yRng = CycleLifeSheet.Range(CycleLifeSheet.Cells(4, currentCol), CycleLifeSheet.Cells(maxCyclesCount + 3, currentCol))
        yRngs.Add yRng
    Next currentCol

    ' 创建图表
    Dim chartObj As ChartObject
    Dim chartLeft As Long
    Dim chartTop As Long
    
    ' 根据列标题设置图表位置
    If columnTitle = COL_NAME_CAPACITY_RETENTION Then
        chartLeft = 50
        chartTop = 50
    ElseIf columnTitle = COL_NAME_ENERGY_RETENTION Then
        chartLeft = 550
        chartTop = 50
    Else
        chartLeft = 50
        chartTop = 50
    End If
    
    Dim axisTitle As String
    If columnTitle = COL_NAME_CAPACITY_RETENTION Then
        axisTitle = "Capacity Retention"
    ElseIf columnTitle = COL_NAME_ENERGY_RETENTION Then
        axisTitle = "Energy Retention"
    Else
        axisTitle = columnTitle
    End If
    
    Set chartObj = CreateCapacityRetentionChart(newWorksheet, xRng, yRngs, axisTitle, reportTitle, chartLeft, chartTop)
    If chartObj Is Nothing Then
        ProcessCapacityRetentionData = False
        Exit Function
    End If
    
    ProcessCapacityRetentionData = True
    Exit Function

ErrorHandler:
    Debug.Print "错误发生在ProcessCapacityRetentionData函数中: " & Err.Description
    ProcessCapacityRetentionData = False
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
Public Function FindColumnByTitle(ws As Worksheet, ByVal columnTitle As String) As Long
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
Public Function GetCellCount(ws As Worksheet, ByVal targetCol As Long) As Long
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
Public Function CreateWorksheet(ByVal wb As Workbook) As Worksheet
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