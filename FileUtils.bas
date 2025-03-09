' ======================
' 标准模块：FileUtils
' ======================
' 平台兼容性：本模块支持 Windows 和 macOS 平台
' 通过使用 Application.PathSeparator 来处理文件路径分隔符
' Windows 使用 "\" 而 macOS 使用 "/"
' ======================
Option Explicit

' 错误常量定义
Private Const ERR_SHEET_NOT_FOUND As Long = 1001
Private Const ERR_TABLE_NOT_FOUND As Long = 1002
Private Const ERR_NO_DATA As Long = 1003
Private Const ERR_FILE_NOT_FOUND As Long = 1004
Private Const ERR_INVALID_FILE_EXT As Long = 1005
Private Const ERR_WORKBOOK_OPEN_FAILED As Long = 1006
Private Const ERR_INVALID_DATA_FORMAT As Long = 1007

' =====================
' 文件和工作表处理函数
' =====================
' 功能：打开并返回指定的Cycle Life工作表
' 参数：
'   sheetName - 首页工作表名称
'   tableName - 文件名表名称
'   targetSheetName - 目标工作表名称
' 返回值：目标工作表对象，失败返回Nothing
Public Function OpenCycleLifeWorksheet(ByVal sheetName As String, ByVal tableName As String, ByVal targetSheetName As String) As Worksheet
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
    If tblFileNames.ListColumns("文件名").DataBodyRange Is Nothing Then
        Err.Raise ERR_NO_DATA, "OpenCycleLifeWorksheet", tableName & "中没有数据"
    End If
    
    ' 检查并添加文件扩展名
    fileName = tblFileNames.ListColumns("文件名").DataBodyRange(1).Value
    
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
Public Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    GetAttr filePath
    FileExists = (Err.Number = 0)
    On Error GoTo 0
End Function

' =====================
' 工作表获取函数
' =====================
Public Function GetWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheet = wb.Sheets(sheetName)
    On Error GoTo 0
End Function

' =====================
' 文件路径处理函数
' =====================
Public Function GetValidFilePath(ByVal fileName As String) As String
    Dim result As String
    result = fileName
    
    If InStr(1, result, ".", vbTextCompare) = 0 Then
        result = result & ".xlsx"
    ElseIf InStr(1, result, ".xlsx", vbTextCompare) = 0 And _
           InStr(1, result, ".xls", vbTextCompare) = 0 Then
        result = Left(result, InStrRev(result, ".")) & "xlsx"
    End If
    
    GetValidFilePath = ThisWorkbook.Path & Application.PathSeparator & result
End Function

' =====================
' 错误处理和日志函数
' =====================
Public Sub HandleError(errNumber As Long, Optional customMessage As String = "")
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
Public Sub UpdateProgress(ByVal progressText As String, Optional ByVal progressPercent As Long = -1)
    If progressPercent >= 0 Then
        Application.StatusBar = progressText & " (" & progressPercent & "%)"
    Else
        Application.StatusBar = progressText
    End If
End Sub 