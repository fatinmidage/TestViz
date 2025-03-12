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
    Dim RPTCycleLifeSheet As Worksheet
    Dim newWorksheet As Worksheet
    Dim reportTitle As String
    reportTitle = GetFileNameFromTable(COL_NAME_REPORTNAME)
    
    ' 获取电池信息集合
    Dim batteriesInfoCollection As Collection
    Set batteriesInfoCollection = GetBatteriesInfo("电池名字颜色")
    
    ' 打开数据工作簿
    Set originDataWorkbook = OpenWorkbookFromTable()
    If originDataWorkbook Is Nothing Then
        GoTo ExitSub
    End If
    
    ' 获取并验证Cycle Life工作表
    Set CycleLifeSheet = originDataWorkbook.Worksheets(SHEET_NAME_CYCLE_LIFE)
    If CycleLifeSheet Is Nothing Then
        MsgBox "无法获取" & SHEET_NAME_CYCLE_LIFE & "工作表", vbCritical
    End If

    '获取并验证RPT of Cycle Life工作表
    Set RPTCycleLifeSheet = originDataWorkbook.Worksheets(SHEET_NAME_RPT_CYCLE_LIFE)
    If RPTCycleLifeSheet Is Nothing Then
        MsgBox "无法获取" & SHEET_NAME_RPT_CYCLE_LIFE & "工作表", vbCritical
    End If
    
    If CycleLifeSheet Is Nothing And RPTCycleLifeSheet Is Nothing Then
        GoTo ExitSub
    End If
    
    ' 创建新工作表
    Set newWorksheet = CreateWorksheet(CycleLifeSheet.Parent)
    If newWorksheet Is Nothing Then
        GoTo ExitSub
    End If
    
    ' 处理Cycle Life工作表数据
    If Not CycleLifeSheet Is Nothing Then
        ' 处理容量保持率数据并创建图表
        If Not ProcessRetentionData(CycleLifeSheet, newWorksheet, COL_NAME_CAPACITY_RETENTION, reportTitle, batteriesInfoCollection) Then
            MsgBox "处理容量保持率数据时发生错误", vbExclamation
        End If
        
        ' 处理能量保持率数据并创建图表
        If Not ProcessRetentionData(CycleLifeSheet, newWorksheet, COL_NAME_ENERGY_RETENTION, reportTitle, batteriesInfoCollection) Then
            MsgBox "处理能量保持率数据时发生错误", vbExclamation
        End If
    End If
    
    ' 处理RPT of Cycle Life工作表数据
    If Not RPTCycleLifeSheet Is Nothing Then
       '处理容量保持率数据并创建图表
        If Not ProcessRPTData(RPTCycleLifeSheet, newWorksheet, COL_NAME_CAPACITY_RETENTION, "DCIR增长率/%", batteriesInfoCollection) Then
            MsgBox "处理容量保持率数据时发生错误", vbExclamation
        End If

       '处理能量保持率数据并创建图表
        If Not ProcessRPTData(RPTCycleLifeSheet, newWorksheet, COL_NAME_ENERGY_RETENTION, "DCIR增长率/%", batteriesInfoCollection) Then
            MsgBox "处理能量保持率数据时发生错误", vbExclamation
        End If
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
        If enable Then
            .Calculation = xlCalculationManual
        Else
            .Calculation = xlCalculationAutomatic
        End If
        .EnableEvents = Not enable
        If enable Then
            .StatusBar = "正在处理数据..."
        Else
            .StatusBar = False
        End If
    End With
End Sub

'******************************************
' 函数: GetBatteriesInfo
' 用途: 从指定表格中读取电池信息
' 参数:
'   tableName - 电池信息表的名称
' 返回值: 包含BatteryInfo对象的Collection集合
'******************************************
Private Function GetBatteriesInfo(ByVal tableName As String) As Collection
    On Error GoTo ErrorHandler
    
    ' 初始化返回值
    Dim batteriesInfoCollection As Collection
    Set batteriesInfoCollection = New Collection
    
    ' 获取电池信息表
    Dim batteriesInfoTable As ListObject
    Set batteriesInfoTable = GetTableByName(tableName)
    
    ' 验证表格是否存在且有数据
    If Not batteriesInfoTable Is Nothing Then
        If Not batteriesInfoTable.DataBodyRange Is Nothing Then
            ' 遍历表格中的每一行
            Dim row As ListRow
            Dim batteryInfo As BatteryInfo
            
            For Each row In batteriesInfoTable.ListRows
                ' 创建新的BatteryInfo对象
                Set batteryInfo = New BatteryInfo
                
                ' 获取电池名称和颜色
                batteryInfo.BatteryName = row.Range.Cells(1, row.Parent.ListColumns("名字").Index).Value
                batteryInfo.BatteryColor = row.Range.Cells(1, row.Parent.ListColumns("颜色").Index).Interior.Color
                
                ' 将BatteryInfo对象添加到集合中
                batteriesInfoCollection.Add batteryInfo
            Next row
        Else
            Debug.Print "电池名字颜色表没有数据"
        End If
    Else
        Debug.Print "未找到电池名字颜色表"
    End If
    
    Set GetBatteriesInfo = batteriesInfoCollection
    Exit Function

ErrorHandler:
    Set GetBatteriesInfo = New Collection
    Debug.Print "获取电池信息时发生错误: " & Err.Description
End Function