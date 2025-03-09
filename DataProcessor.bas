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
Private Const TABLE_NAME_FILES As String = "文件名表"
Private Const COL_NAME_FILENAME As String = "文件名"
Private Const ERR_SHEET_NOT_FOUND As Long = 1001
Private Const ERR_TABLE_NOT_FOUND As Long = 1002
Private Const ERR_NO_DATA As Long = 1003
Private Const ERR_FILE_NOT_FOUND As Long = 1004
Private Const ERR_INVALID_FILE_EXT As Long = 1005
Private Const ERR_WORKBOOK_OPEN_FAILED As Long = 1006
Private Const ERR_INVALID_DATA_FORMAT As Long = 1007

Sub ProcessTestData()
    On Error GoTo ErrorHandler
    
    Dim wsCycleLife As Worksheet
    
    ' 添加性能优化设置
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .StatusBar = "正在处理数据..."
    End With
    
    ' 获取Cycle Life工作表
    Set wsCycleLife = OpenCycleLifeWorksheet()
    
    If wsCycleLife Is Nothing Then
        GoTo ExitSub
    End If
    
    ' TODO: 在这里添加对wsCycleLife的数据处理逻辑
    
    ' 完成后关闭工作簿
    wsCycleLife.Parent.Close False
    
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
    Exit Sub

ErrorHandler:
    Call HandleError(Err.Number, "错误处理代码")
    Resume ExitSub
End Sub

Private Function OpenCycleLifeWorksheet() As Worksheet
    On Error GoTo ErrorHandler
    
    ' 添加变量声明
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim tblFileNames As ListObject
    Dim fileName As String
    Dim filePath As String
    Dim wsCycleLife As Worksheet
    
    ' 检查首页工作表是否存在
    Set ws = ThisWorkbook.Sheets(SHEET_NAME_HOME)
    If ws Is Nothing Then
        MsgBox "未找到'" & SHEET_NAME_HOME & "'工作表!", vbExclamation
        Exit Function
    End If
    
    ' 检查文件名表是否存在
    Set tblFileNames = ws.ListObjects(TABLE_NAME_FILES)
    If tblFileNames Is Nothing Then
        MsgBox "未找到'" & TABLE_NAME_FILES & "'列表对象!", vbExclamation
        Exit Function
    End If
    
    ' 检查文件名列是否存在且有数据
    If tblFileNames.ListColumns(COL_NAME_FILENAME).DataBodyRange Is Nothing Then
        MsgBox TABLE_NAME_FILES & "中没有数据!", vbExclamation
        Exit Function
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
    
    ' 读取Cycle Life工作表
    Set wsCycleLife = GetWorksheet(wb, SHEET_NAME_CYCLE_LIFE)
    
    If wsCycleLife Is Nothing Then
        MsgBox "未找到'" & SHEET_NAME_CYCLE_LIFE & "'工作表!", vbExclamation
        wb.Close False
        Exit Function
    End If
    
    ' 返回工作表对象
    Set OpenCycleLifeWorksheet = wsCycleLife
    Exit Function

ErrorHandler:
    Call HandleError(Err.Number, "打开Cycle Life工作表时发生错误")
    Exit Function
End Function


' 添加新的辅助函数在模块末尾
Private Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    GetAttr filePath
    FileExists = (Err.Number = 0)
    On Error GoTo 0
End Function

' 优化工作表检查逻辑
Private Function GetWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheet = wb.Sheets(sheetName)
    On Error GoTo 0
End Function

' 增强错误处理函数
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

' 添加错误日志功能
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

Private Sub UpdateProgress(ByVal progressText As String, Optional ByVal progressPercent As Long = -1)
    If progressPercent >= 0 Then
        Application.StatusBar = progressText & " (" & progressPercent & "%)"
    Else
        Application.StatusBar = progressText
    End If
End Sub

Private Function GetValidFilePath(ByVal fileName As String) As String
    Dim result As String
    result = fileName
    
    ' 添加文件扩展名检查
    If InStr(1, result, ".", vbTextCompare) = 0 Then
        result = result & ".xlsx"
    ElseIf InStr(1, result, ".xlsx", vbTextCompare) = 0 And _
           InStr(1, result, ".xls", vbTextCompare) = 0 Then
        result = Left(result, InStrRev(result, ".")) & "xlsx"
    End If
    
    ' 构建完整路径
    GetValidFilePath = ThisWorkbook.Path & Application.PathSeparator & result
End Function

            
