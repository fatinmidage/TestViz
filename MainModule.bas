Option Explicit

' 在模块顶部添加常量定义
Public Const SHEET_NAME_CYCLE_LIFE As String = "Cycle Life"
Public Const SHEET_NAME_RPT_CYCLE_LIFE As String = "RPT of Cycle Life"
Public Const TABLE_NAME_FILES As String = "文件名表"
Public Const COL_NAME_FILENAME As String = "文件名"
Public Const COL_NAME_REPORTNAME As String = "报告标题"
Public Const COL_NAME_CAPACITY_RETENTION As String = "容量保持率/%"
Public Const COL_NAME_ENERGY_RETENTION As String = "能量保持率/%"
Public Const ERR_TABLE_NOT_FOUND As Long = 1002
Public Const ERR_INVALID_DATA_FORMAT As Long = 1007

' =====================
' 主函数
' =====================
Public Sub Main()
    On Error GoTo ErrorHandler
    SetPerformanceMode True
    
    ' 声明变量
    Dim originDataWorkbook As Workbook
    Dim CycleLifeSheet As Worksheet
    Dim newWorksheet As Worksheet
    Dim reportTitle As String
    reportTitle = GetFileNameFromTable(COL_NAME_REPORTNAME)
    
    
    ' 打开数据工作簿
    Set originDataWorkbook = OpenWorkbookFromTable()
    If originDataWorkbook Is Nothing Then
        GoTo ExitSub
    End If
    
    ' 获取并验证Cycle Life工作表
    Set CycleLifeSheet = originDataWorkbook.Worksheets(SHEET_NAME_CYCLE_LIFE)
    If CycleLifeSheet Is Nothing Then
        MsgBox "无法获取" & SHEET_NAME_CYCLE_LIFE & "工作表", vbCritical
        GoTo ExitSub
    End If
    
    ' 创建新工作表
    Set newWorksheet = CreateWorksheet(CycleLifeSheet.Parent)
    If newWorksheet Is Nothing Then
        GoTo ExitSub
    End If
    
    ' 处理容量保持率数据并创建图表
    If Not ProcessCapacityRetentionData(CycleLifeSheet, newWorksheet, COL_NAME_CAPACITY_RETENTION, reportTitle) Then
        GoTo ExitSub
    End If
    
    ' 处理能量保持率数据并创建图表
    If Not ProcessCapacityRetentionData(CycleLifeSheet, newWorksheet, COL_NAME_ENERGY_RETENTION, reportTitle) Then
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
Public Sub SetPerformanceMode(ByVal enable As Boolean)
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