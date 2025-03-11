Option Explicit

'******************************************
' 函数: GetFileNameFromTable
' 用途: 从指定的表格中获取文件名
' 参数:
'   columnName - 文件名列的标题名称
' 返回值: 成功返回文件名，失败抛出错误
'******************************************
Public Function GetFileNameFromTable(ByVal columnName As String) As String
    On Error GoTo ErrorHandler
    
    ' 获取文件名表
    Dim tblFileNames As ListObject
    Dim atws As Worksheet
    Set atws = ActiveSheet
    
    ' 参数验证
    ' 检查当前工作表是否可用
    If atws Is Nothing Then
        Err.Raise ERR_INVALID_DATA_FORMAT, "GetFileNameFromTable", "无法获取当前工作表"
    End If
    
    ' 获取并验证文件名表是否存在
    Set tblFileNames = atws.ListObjects(TABLE_NAME_FILES)
    If tblFileNames Is Nothing Then
        Err.Raise ERR_TABLE_NOT_FOUND, "GetFileNameFromTable", "未找到'" & TABLE_NAME_FILES & "'表格"
    End If
    
    ' 验证文件名列是否存在且有数据
    ' 获取文件名列对象
    Dim fileNameColumn As ListColumn
    Set fileNameColumn = tblFileNames.ListColumns(columnName)
    If fileNameColumn Is Nothing Then
        Err.Raise ERR_INVALID_DATA_FORMAT, "GetFileNameFromTable", "未找到'" & columnName & "'列"
    End If
    
    ' 检查文件名列是否包含数据
    If fileNameColumn.DataBodyRange Is Nothing Then
        Err.Raise ERR_INVALID_DATA_FORMAT, "GetFileNameFromTable", "文件名列没有数据"
    End If
    
    ' 获取文件名并处理文件路径
    ' 从第一行获取文件名
    Dim fileName As String
    fileName = fileNameColumn.DataBodyRange(1).Value
    
    ' 验证文件名是否为空
    ' 移除文件名前后的空格并检查长度
    If Len(Trim(fileName)) = 0 Then
        Err.Raise ERR_INVALID_DATA_FORMAT, "GetFileNameFromTable", "文件名不能为空"
    End If
    
    
    GetFileNameFromTable = fileName
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "GetFileNameFromTable", Err.Description
End Function

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
Public Function OpenWorkbookFromTable() As Workbook
    On Error GoTo ErrorHandler
    
    ' 获取文件名
    Dim fileName As String
    fileName = GetFileNameFromTable(COL_NAME_FILENAME)
    
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

'******************************************
' 函数: HandleError
' 用途: 统一处理错误信息
' 参数:
'   - errNumber: 错误号
'   - errDescription: 错误描述
'******************************************
Public Sub HandleError(ByVal errNumber As Long, ByVal errDescription As String)
    Debug.Print "错误号: " & errNumber & ", 描述: " & errDescription
    MsgBox errDescription, vbCritical
End Sub

'******************************************
' 函数: FileExists
' 用途: 检查指定文件是否存在
' 参数:
'   - filePath: 要检查的文件路径
' 返回值: 文件存在返回True，否则返回False
'******************************************
Public Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(filePath) <> "")
    On Error GoTo 0
End Function 