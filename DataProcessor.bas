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
Public Sub ProcessTestData()
    On Error GoTo ErrorHandler
    
    Dim wsCycleLife As Worksheet
    Dim wsRptCycleLife As Worksheet
    
    ' 添加性能优化设置
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .StatusBar = "正在处理数据..."
    End With
    
    ' 获取Cycle Life工作表
    Set wsCycleLife = FileUtils.OpenCycleLifeWorksheet(SHEET_NAME_HOME, TABLE_NAME_FILES, SHEET_NAME_CYCLE_LIFE)
    
    If wsCycleLife Is Nothing Then
        Err.Raise ERR_SHEET_NOT_FOUND, "ProcessTestData", "无法打开Cycle Life工作表"
    End If

    ' 获取RPT of Cycle Life工作表
    Set wsRptCycleLife = FileUtils.OpenCycleLifeWorksheet(SHEET_NAME_HOME, TABLE_NAME_FILES, SHEET_NAME_RPT_CYCLE_LIFE)

    If wsRptCycleLife Is Nothing Then
        Err.Raise ERR_SHEET_NOT_FOUND, "ProcessTestData", "无法打开RPT of Cycle Life工作表"
    End If
    
    ' 处理数据
    Dim capacityRetentionRates As Collection
    Set capacityRetentionRates = GetCycleData(wsCycleLife, "容量保持率/%")
    dim energyRetentionRates As Collection
    Set energyRetentionRates = GetCycleData(wsCycleLife, "能量保持率/%")
    
    ' 完成后关闭工作簿
    If Not wsCycleLife Is Nothing Then
        wsCycleLife.Parent.Close False
    End If
    If Not wsRptCycleLife Is Nothing Then
        wsRptCycleLife.Parent.Close False
    End If
    
ExitSub:
    ' 恢复所有设置
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .StatusBar = False
    End With
    Set wsCycleLife = Nothing
    Set wsRptCycleLife = Nothing
    Exit Sub

ErrorHandler:
    Call FileUtils.HandleError(Err.Number, Err.Description)
    Resume ExitSub
End Sub

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
    For currentCol = targetCol To targetCol + cellCount - 1
        result.Add ExtractColumnData(ws, currentCol)
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
Private Function ExtractColumnData(ByVal ws As Worksheet, ByVal targetCol As Long, Optional ByVal startRow As Long = 4) As Double()
    ' 获取数据的最后一行
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, targetCol).End(xlUp).Row
        
    ' 数据验证：确保存在有效数据
    If lastRow < startRow Then
        Err.Raise ERR_NO_DATA, "ExtractColumnData", "数据为空"
        Exit Function
    End If
    
    ' 一次性读取整列数据到数组
    Dim dataArray As Variant
    Dim dataRange As Range
    Set dataRange = ws.Range(ws.Cells(startRow, targetCol), ws.Cells(lastRow, targetCol))
    dataArray = dataRange.Value2

    ' 初始化结果数组
    Dim resultData() As Double
    ReDim resultData(1 To UBound(dataArray, 1))
    
    ' 数据转换：将Variant数组转换为Double数组
    Dim i As Long
    For i = 1 To UBound(dataArray, 1)
        resultData(i) = dataArray(i, 1)
    Next i
    
    ' 返回处理后的数据数组
    ExtractColumnData = resultData
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

            
