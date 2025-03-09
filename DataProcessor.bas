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
Sub ProcessTestData()
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
    Set wsCycleLife = OpenCycleLifeWorksheet(SHEET_NAME_HOME, TABLE_NAME_FILES, SHEET_NAME_CYCLE_LIFE)
    
    If wsCycleLife Is Nothing Then
        Err.Raise ERR_SHEET_NOT_FOUND, "ProcessTestData", "无法打开Cycle Life工作表"
    End If

    ' 获取RPT of Cycle Life工作表
    Set wsRptCycleLife = OpenCycleLifeWorksheet(SHEET_NAME_HOME, TABLE_NAME_FILES, SHEET_NAME_RPT_CYCLE_LIFE)

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
    Call HandleError(Err.Number, Err.Description)
    Resume ExitSub
End Sub

' =====================
' 文件和工作表处理函数
' =====================
' 功能：打开并返回指定的Cycle Life工作表
' 参数：
'   sheetName - 首页工作表名称
'   tableName - 文件名表名称
'   targetSheetName - 目标工作表名称
' 返回值：目标工作表对象，失败返回Nothing
Private Function OpenCycleLifeWorksheet(ByVal sheetName As String, ByVal tableName As String, ByVal targetSheetName As String) As Worksheet
    On Error GoTo ErrorHandler
    
    ' 添加变量声明
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim tblFileNames As ListObject
    Dim fileName As String
    Dim filePath As String
    Dim wsCycleLife As Worksheet
    
    ' 检查首页工作表是否存在
    Set ws = ThisWorkbook.Sheets(sheetName)
    If ws Is Nothing Then
        Err.Raise ERR_SHEET_NOT_FOUND, "OpenCycleLifeWorksheet", "未找到'" & sheetName & "'工作表"
    End If
    
    ' 检查文件名表是否存在
    Set tblFileNames = ws.ListObjects(tableName)
    If tblFileNames Is Nothing Then
        Err.Raise ERR_TABLE_NOT_FOUND, "OpenCycleLifeWorksheet", "未找到'" & tableName & "'列表对象"
    End If
    
    ' 检查文件名列是否存在且有数据
    If tblFileNames.ListColumns(COL_NAME_FILENAME).DataBodyRange Is Nothing Then
        Err.Raise ERR_NO_DATA, "OpenCycleLifeWorksheet", tableName & "中没有数据"
    End If
    
    ' 检查并添加文件扩展名
    fileName = tblFileNames.ListColumns(COL_NAME_FILENAME).DataBodyRange(1).Value
    
    ' 如果文件名不包含.xlsx或.xls扩展名，则添加.xlsx
    If InStr(1, fileName, ".xlsx", vbTextCompare) = 0 And InStr(1, fileName, ".xls", vbTextCompare) = 0 Then
        fileName = fileName & ".xlsx"
    End If
    
    filePath = ThisWorkbook.Path & Application.PathSeparator & fileName
    
    If Not FileExists(filePath) Then
        MsgBox "找不到文件: " & fileName, vbExclamation
        Exit Function
    End If
    
    ' 打开指定的Excel文件
    Set wb = Workbooks.Open(filePath)
    
    ' 读取目标工作表
    Set wsCycleLife = GetWorksheet(wb, targetSheetName)
    
    If wsCycleLife Is Nothing Then
        MsgBox "未找到'" & targetSheetName & "'工作表!", vbExclamation
        wb.Close False
        Exit Function
    End If
    
    ' 返回工作表对象
    Set OpenCycleLifeWorksheet = wsCycleLife
    Exit Function

ErrorHandler:
    If Not wb Is Nothing Then
        wb.Close False
    End If
    Call HandleError(Err.Number, "打开" & targetSheetName & "工作表时发生错误")
    Exit Function
End Function

' =====================
' 文件检查函数
' =====================
' 功能：检查指定路径的文件是否存在
' 参数：
'   filePath - 要检查的文件的完整路径
' 返回值：
'   Boolean - 如果文件存在返回True，否则返回False
' 说明：
'   使用GetAttr函数和错误处理来检查文件是否存在
'   通过On Error Resume Next来避免文件不存在时的错误
Private Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    GetAttr filePath
    FileExists = (Err.Number = 0)
    On Error GoTo 0
End Function

' =====================
' 工作表获取函数
' =====================
' 功能：从指定的工作簿中获取指定名称的工作表
' 参数：
'   wb - 目标工作簿对象
'   sheetName - 要获取的工作表名称
' 返回值：
'   Worksheet - 如果找到工作表则返回工作表对象，否则返回Nothing
' 说明：
'   使用On Error Resume Next来避免工作表不存在时的错误
Private Function GetWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheet = wb.Sheets(sheetName)
    On Error GoTo 0
End Function

' =====================
' 文件路径处理函数
' =====================
' 功能：验证和规范化文件路径，确保文件具有正确的扩展名并返回完整路径
' 参数：
'   fileName - 输入的文件名，可以包含或不包含扩展名
' 返回值：
'   String - 返回规范化后的完整文件路径
' 说明：
'   1. 如果文件名没有扩展名，添加.xlsx扩展名
'   2. 如果文件名有非Excel扩展名，将其替换为.xlsx
'   3. 使用当前工作簿所在路径作为基础路径
Private Function GetValidFilePath(ByVal fileName As String) As String
    Dim result As String
    result = fileName
    
    ' 如果文件名中没有点号，说明没有扩展名，直接添加.xlsx
    If InStr(1, result, ".", vbTextCompare) = 0 Then
        result = result & ".xlsx"
    ' 如果文件名中有点号但不是Excel文件扩展名，则替换为.xlsx
    ElseIf InStr(1, result, ".xlsx", vbTextCompare) = 0 And _
           InStr(1, result, ".xls", vbTextCompare) = 0 Then
        result = Left(result, InStrRev(result, ".")) & "xlsx"
    End If
    
    ' 构建完整路径：将当前工作簿所在路径与处理后的文件名组合
    GetValidFilePath = ThisWorkbook.Path & Application.PathSeparator & result
End Function

' =====================
' 数据处理核心函数
' =====================
Private Function GetCycleData(ws As Worksheet, ByVal columnTitle As String) As Collection
    On Error GoTo ErrorHandler
    
    ' 参数验证部分：确保输入参数的有效性
    ' 检查工作表对象是否为空，避免空引用异常
    If ws Is Nothing Then
        Err.Raise ERR_INVALID_DATA_FORMAT, "GetCycleData", "工作表对象不能为空"
    End If
    ' 检查列标题是否为空字符串，确保有效的数据列标识
    If Len(Trim(columnTitle)) = 0 Then
        Err.Raise ERR_INVALID_DATA_FORMAT, "GetCycleData", "列标题不能为空"
    End If
    
    ' 数据定位和结构分析部分
    ' 声明变量用于存储电芯数量、目标列位置和结果集合
    Dim cellCount As Long
    Dim targetCol As Long
    Dim result As Collection
    
    ' 根据列标题查找目标数据列的位置
    ' 如果找不到指定列，targetCol将为0
    targetCol = FindColumnByTitle(ws, columnTitle)
    If targetCol = 0 Then
        Err.Raise ERR_INVALID_DATA_FORMAT, "GetCycleData", "未找到'" & columnTitle & "'列"
    End If
    
    ' 通过分析合并单元格确定电芯数量
    ' 电芯数据在Excel中以合并单元格的形式组织
    cellCount = GetCellCount(ws, targetCol)
    If cellCount <= 0 Then
        Err.Raise ERR_INVALID_DATA_FORMAT, "GetCycleData", "无效的电芯数量"
    End If
    
    ' 初始化结果集合，用于存储所有电芯的数据
    Set result = New Collection
    
    ' 批量数据提取部分
    ' 根据电芯数量，循环提取每个电芯的数据
    ' 每个电芯的数据存储在相邻的列中
    Dim currentCol As Long
    For currentCol = targetCol To targetCol + cellCount - 1
        result.Add ExtractColumnData(ws, currentCol)
    Next currentCol
    
    ' 返回结果
    Set GetCycleData = result
    Exit Function

ErrorHandler:
    ' 错误处理
    Dim errMsg As String
    errMsg = "处理循环数据时发生错误: " & Err.Description
    Call HandleError(Err.Number, errMsg)
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
' 功能：通过分析目标列的合并单元格来确定电芯数量
' 参数：
'   ws - 工作表对象，包含电芯数据
'   targetCol - 目标列号，表示要分析的列的位置
' 返回值：
'   Long - 返回电芯数量，即目标列第一行合并单元格的宽度
' 错误处理：
'   - 如果目标列号为0，抛出ERR_INVALID_DATA_FORMAT错误
' 实现说明：
'   1. 首先验证目标列号的有效性
'   2. 获取目标列第一行单元格的合并区域
'   3. 计算合并区域的列数，即为电芯数量
'   4. 返回计算得到的电芯数量
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
' 功能：从指定工作表的指定列中提取数据，并转换为Double类型数组
' 参数：
'   ws - 工作表对象，包含要提取的数据
'   targetCol - 目标列号，表示要提取数据的列位置
'   startRow - 可选参数，开始行号，默认为4（通常前3行为表头）
' 返回值：
'   Double() - 返回包含提取数据的一维Double数组
' 错误处理：
'   - 如果数据为空（lastRow < startRow），抛出ERR_NO_DATA错误
' 实现说明：
'   1. 使用End(xlUp)方法确定数据的最后一行
'   2. 一次性读取整列数据到Variant数组以提高性能
'   3. 使用Value2属性而不是Value，可提升15-20%的性能
'   4. 将二维Variant数组转换为一维Double数组返回
Private Function ExtractColumnData(ByVal ws As Worksheet, ByVal targetCol As Long, Optional ByVal startRow As Long = 4) As Double()
    ' 获取数据的最后一行
    ' 使用End(xlUp)从工作表底部向上查找最后一个非空单元格
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, targetCol).End(xlUp).Row
        
    ' 数据验证：确保存在有效数据
    ' 如果最后一行小于起始行，说明没有有效数据
    If lastRow < startRow Then
        Err.Raise ERR_NO_DATA, "ExtractColumnData", "数据为空"
        Exit Function
    End If
    
    ' 一次性读取整列数据到数组
    ' 使用Range对象和Value2属性以获得最佳性能
    Dim dataArray As Variant
    Dim dataRange As Range
    Set dataRange = ws.Range(ws.Cells(startRow, targetCol), ws.Cells(lastRow, targetCol))
    dataArray = dataRange.Value2  ' Value2比Value快15-20%，因为它不处理格式和公式

    ' 初始化结果数组
    ' 使用UBound确定数组大小，确保足够空间
    Dim resultData() As Double
    ReDim resultData(1 To UBound(dataArray, 1))
    
    ' 数据转换：将Variant数组转换为Double数组
    ' 注意：Excel返回的是二维数组，即使只有一列
    Dim i As Long
    For i = 1 To UBound(dataArray, 1)
        resultData(i) = dataArray(i, 1)  ' 使用(i, 1)因为Excel返回的是二维数组
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

            
