' ======================
' 标准模块：DataProcessor
' ======================
' 平台兼容性：本模块支持 Windows 和 macOS 平台
' 通过使用 Application.PathSeparator 来处理文件路径分隔符
' Windows 使用 "\" 而 macOS 使用 "/"
' ======================
Option Explicit

' 在模块顶部添加常量定义
Private Const SHEET_NAME_HOME As String = "首页"
Private Const SHEET_NAME_CYCLE_LIFE As String = "Cycle Life"
Private Const SHEET_NAME_RPT_CYCLE_LIFE As String = "RPT of Cycle Life"
Private Const TABLE_NAME_FILES As String = "文件名表"
Private Const COL_NAME_FILENAME As String = "文件名"
Private Const COL_NAME_CAPACITY_RETENTION As String = "容量保持率/%"
Private Const COL_NAME_ENERGY_RETENTION As String = "能量保持率/%"
Private Const ERR_SHEET_NOT_FOUND As Long = 1001
Private Const ERR_TABLE_NOT_FOUND As Long = 1002
Private Const ERR_NO_DATA As Long = 1003
Private Const ERR_FILE_NOT_FOUND As Long = 1004
Private Const ERR_INVALID_FILE_EXT As Long = 1005
Private Const ERR_WORKBOOK_OPEN_FAILED As Long = 1006
Private Const ERR_INVALID_DATA_FORMAT As Long = 1007

' =====================
' 主入口函数
' =====================
' 功能：处理电池测试数据，提取和分析循环寿命数据
' 输入：无
' 输出：无
' 说明：
'   1. 该函数是数据处理的主入口，负责协调整个数据处理流程
'   2. 从Cycle Life工作表中提取容量保持率和能量保持率数据
'   3. 包含性能优化设置，以提高大量数据处理时的效率
'   4. 实现了完整的错误处理机制，确保异常情况下的安全退出
' =====================
Public Sub ProcessTestData()
    On Error GoTo ErrorHandler
    
    ' 声明工作簿对象，用于保存原始数据
    Dim sourceWorkbook As Workbook
     ' 创建新的工作表用于数据分析
    Dim wsNewChart As Worksheet
   
    ' 性能优化设置部分
    ' 以下设置用于提高大量数据处理时的性能表现
    With Application
        ' 关闭屏幕更新以减少视觉刷新开销，可显著提升处理速度
        .ScreenUpdating = False
        ' 关闭警告提示以避免用户交互导致的中断
        .DisplayAlerts = False
        ' 设置为手动计算模式，避免每次单元格变更都触发公式重算
        .Calculation = xlCalculationManual
        ' 关闭事件触发以避免不必要的事件处理开销
        .EnableEvents = False
        ' 在状态栏显示当前处理状态，提供用户反馈
        .StatusBar = "正在处理数据..."
    End With
    
    ' 数据获取部分
    ' 从Cycle Life工作表中提取两类关键性能指标数据
    Dim capacityRetentionRates As Collection
    Set capacityRetentionRates = GetWorksheetData(SHEET_NAME_CYCLE_LIFE, COL_NAME_CAPACITY_RETENTION, sourceWorkbook)
    
    ' 获取新的工作表对象
    Set wsNewChart = FileUtils.CreateWorksheet(sourceWorkbook)
    If wsNewChart Is Nothing Then
        Err.Raise ERR_WORKBOOK_OPEN_FAILED, "ProcessTestData", "创建数据分析工作表失败"
    End If
    
    Dim energyRetentionRates As Collection
    Set energyRetentionRates = GetWorksheetData(SHEET_NAME_CYCLE_LIFE, COL_NAME_ENERGY_RETENTION, sourceWorkbook)
    
ExitSub:
    ' 恢复所有设置
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .StatusBar = False
    End With
    Exit Sub

ErrorHandler:
    Call FileUtils.HandleError(Err.Number, Err.Description)
    Resume ExitSub
End Sub


' =====================
' 工作表数据获取函数
' =====================
' =====================
' 工作表数据获取函数
' =====================
' 功能：从指定工作表中获取特定列的数据
' 参数：
'   sheetName - 目标工作表名称，通常为"Cycle Life"工作表
'   columnTitle - 要提取数据的列标题（如容量保持率或能量保持率）
'   wb - 引用参数，返回包含目标工作表的工作簿对象
' 返回值：
'   Collection类型，包含从指定列提取的数据
' 错误处理：
'   - 如果工作表不存在，抛出ERR_SHEET_NOT_FOUND错误
'   - 如果数据获取失败，调用FileUtils.HandleError处理错误
' 调用关系：
'   - 调用FileUtils.OpenCycleLifeWorksheet打开目标工作表
'   - 调用GetCycleData获取具体的数据内容
' =====================
Private Function GetWorksheetData(ByVal sheetName As String, ByVal columnTitle As String, ByRef wb As Workbook) As Collection
    On Error GoTo ErrorHandler
    
    ' 打开目标工作表
    Dim ws As Worksheet
    Set ws = FileUtils.OpenCycleLifeWorksheet(SHEET_NAME_HOME, TABLE_NAME_FILES, sheetName)
    
    ' 验证工作表是否成功打开
    If ws Is Nothing Then
        Err.Raise ERR_SHEET_NOT_FOUND, "GetWorksheetData", "无法打开" & sheetName & "工作表"
    End If

    ' 获取工作表所属的工作簿
    Set wb = ws.Parent
    
    ' 从工作表中提取指定列的数据
    Dim result As Collection
    Set result = GetCycleData(ws, columnTitle)
    
    ' 返回提取的数据集合
    Set GetWorksheetData = result
    Exit Function
    
ErrorHandler:
    Call FileUtils.HandleError(Err.Number, "获取" & sheetName & "数据时发生错误")
    Set GetWorksheetData = Nothing
End Function


' =====================
' 数据处理核心函数
' =====================
Private Function GetCycleData(ws As Worksheet, ByVal columnTitle As String) As Collection
    On Error GoTo ErrorHandler
    
    ' 参数验证部分：确保输入参数的有效性
    If ws Is Nothing Then
        Err.Raise ERR_INVALID_DATA_FORMAT, "GetCycleData", "工作表对象不能为空"
    End If
    If Len(Trim(columnTitle)) = 0 Then
        Err.Raise ERR_INVALID_DATA_FORMAT, "GetCycleData", "列标题不能为空"
    End If
    
    ' 数据定位和结构分析部分
    Dim cellCount As Long
    Dim targetCol As Long
    
    Dim result As Collection
    
    targetCol = FindColumnByTitle(ws, columnTitle)
    If targetCol = 0 Then
        Err.Raise ERR_INVALID_DATA_FORMAT, "GetCycleData", "未找到'" & columnTitle & "'列"
    End If
    
    cellCount = GetCellCount(ws, targetCol)
    If cellCount <= 0 Then
        Err.Raise ERR_INVALID_DATA_FORMAT, "GetCycleData", "无效的电芯数量"
    End If
    
    Set result = New Collection
    
    Dim currentCol As Long
    Dim cellName As String
    dim battery As Battery '声明电池对象

    '循环处理每一列
    For currentCol = targetCol To targetCol + cellCount - 1
        '提取电芯名称
        cellName = Right(ws.Cells(2, currentCol).Value, 4)
        '创建新的电池对象
        Set battery = New Battery
        battery.CellName = cellName
        battery.Cycles = ExtractColumnData(ws, currentCol)
        result.Add battery
    Next currentCol
    
    Set GetCycleData = result
    Exit Function

ErrorHandler:
    Dim errMsg As String
    errMsg = "处理循环数据时发生错误: " & Err.Description
    Call FileUtils.HandleError(Err.Number, errMsg)
    Set GetCycleData = Nothing
End Function

Private Function FindColumnByTitle(ws As Worksheet, ByVal columnTitle As String) As Long
    ' 在第一行查找指定列
    Dim lastCol As Long
    Dim i As Long
    Dim targetCol As Long
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    targetCol = 0
    
    For i = 1 To lastCol
        If ws.Cells(1, i).Value = columnTitle Then
            targetCol = i
            Exit For
        End If
    Next i
    
    FindColumnByTitle = targetCol
End Function

' =====================
' 电芯数量获取函数
' =====================
Private Function GetCellCount(ws As Worksheet, ByVal targetCol As Long) As Long
    ' 检查目标列是否有效
    If targetCol = 0 Then
        Err.Raise ERR_INVALID_DATA_FORMAT, "GetCellCount", "无效的列号"
    End If
    
    ' 检查目标列的单元格是否为合并单元格
    Dim mergeArea As Range
    Set mergeArea = ws.Cells(1, targetCol).MergeArea
    Dim mergeWidth As Long
    mergeWidth = mergeArea.Columns.Count
    
    ' 返回电芯数量
    GetCellCount = mergeWidth
End Function

' =====================
' 列数据提取函数
' =====================
Private Function ExtractColumnData(ByVal ws As Worksheet, ByVal targetCol As Long, Optional ByVal startRow As Long = 4) As Collection
    Dim resultData As New Collection
    ' 获取数据的最后一行
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, targetCol).End(xlUp).Row
        
    ' 数据验证：确保存在有效数据
    If lastRow < startRow Then
        Err.Raise ERR_NO_DATA, "ExtractColumnData", "数据为空"
        Exit Function
    End If
    
    ' 一次性读取整列数据到数组
    Dim targetDataArray As Variant
    Dim targetDataRange As Range
    Dim cycleDataIndexRange As Range
    Dim cycleDataIndexArray As Variant
    Dim cycleData As CycleData

    Set targetDataRange = ws.Range(ws.Cells(startRow, targetCol), ws.Cells(lastRow, targetCol))
    targetDataArray = targetDataRange.Value2

    Set cycleDataIndexRange = ws.Range(ws.Cells(startRow, 1), ws.Cells(lastRow, 1))
    cycleDataIndexArray = cycleDataIndexRange.Value2
    
    ' 数据转换：将Variant数组转换为Double数组
    Dim i As Long
    For i = 1 To UBound(targetDataArray, 1)
        set cycleData = New CycleData
        cycleData.CycleNumber = cycleDataIndexArray(i, 1)
        cycleData.CycleData = targetDataArray(i, 1)
        resultData.Add cycleData
    Next i
    
    ' 返回处理后的数据数组
    Set ExtractColumnData = resultData
End Function

' =====================
' 错误处理和日志函数
' =====================
Private Sub HandleError(errNumber As Long, Optional customMessage As String = "")
    Dim errMsg As String
    Select Case errNumber
        Case ERR_SHEET_NOT_FOUND
            errMsg = "未找到指定的工作表"
        Case ERR_TABLE_NOT_FOUND
            errMsg = "未找到指定的表格"
        Case ERR_NO_DATA
            errMsg = "表格中没有数据"
        Case ERR_FILE_NOT_FOUND
            errMsg = "找不到文件"
        Case ERR_INVALID_FILE_EXT
            errMsg = "文件扩展名无效"
        Case ERR_WORKBOOK_OPEN_FAILED
            errMsg = "无法打开工作簿"
        Case ERR_INVALID_DATA_FORMAT
            errMsg = "数据格式无效"
        Case Else
            errMsg = "未知错误 (错误代码: " & errNumber & ")"
    End Select
    
    errMsg = "错误 " & errNumber & ": " & errMsg & vbNewLine & _
             IIf(Len(customMessage) > 0, "详细信息: " & customMessage & vbNewLine, "") & _
             "发生在第 " & Erl & " 行" & vbNewLine & _
             "时间: " & Now
    
    ' 记录错误到日志
    LogError errMsg
    
    ' 显示错误消息
    MsgBox errMsg, vbCritical, "错误"
End Sub

Private Sub LogError(ByVal errMsg As String)
    Dim fso As Object
    Dim logFile As Object
    Dim logPath As String
    
    On Error Resume Next
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    logPath = ThisWorkbook.Path & Application.PathSeparator & "error_log.txt"
    
    Set logFile = fso.OpenTextFile(logPath, 8, True) ' 8 = ForAppending, Create if doesn't exist
    logFile.WriteLine errMsg & vbNewLine & String(50, "-")
    logFile.Close
    
    Set logFile = Nothing
    Set fso = Nothing
    
    On Error GoTo 0
End Sub

' =====================
' 进度显示函数
' =====================
Private Sub UpdateProgress(ByVal progressText As String, Optional ByVal progressPercent As Long = -1)
    If progressPercent >= 0 Then
        Application.StatusBar = progressText & " (" & progressPercent & "%)"
    Else
        Application.StatusBar = progressText
    End If
End Sub

            
