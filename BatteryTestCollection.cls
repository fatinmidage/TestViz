' 类模块名称: BatteryTestData
Option Explicit

' 基础信息
Private pCycleNumber As Long          ' 循环号
Private pSampleID As String           ' 样本ID (如04QCE27421101JEBS0003183)
Private pTestTime As Date             ' 测试时间

' 容量相关
Private pDischargeCapacity As Double  ' 放电容量(Ah)
Private pCapacityRetention As Double  ' 容量保持率(%)
Private pChargeCapacity As Double     ' 充电容量(Ah)

' 能量相关
Private pDischargeEnergy As Double    ' 放电能量(Wh)
Private pEnergyRetention As Double    ' 能量保持率(%)
Private pChargeEnergy As Double       ' 充电能量(Wh)

' 直流内阻(DCIR)
Private pDCIR As Double               ' 直流内阻(mΩ)
Private pDCIRGrowthRate As Double     ' DCIR增长率(%)

' 电压特性
Private pOCV As Double                ' 开路电压(V)
Private pLoadVoltage As Double        ' 负载电压(V)
Private pCurrent As Double            ' 测试电流(A)

' 数据校验标志
Private ValidationEnabled As Boolean

' 初始化设置
Private Sub Class_Initialize()
    ValidationEnabled = True
    pTestTime = Now
End Sub

' ========== 属性访问器 ==========
' 循环号
Public Property Get CycleNumber() As Long
    CycleNumber = pCycleNumber
End Property
Public Property Let CycleNumber(value As Long)
    If value < 0 Then Err.Raise 5, , "循环号不能为负数"
    pCycleNumber = value
End Property

' 样本ID
Public Property Get SampleID() As String
    SampleID = pSampleID
End Property
Public Property Let SampleID(value As String)
    If Len(value) < 10 Then Err.Raise 5, , "无效样本ID"
    pSampleID = value
End Property

' 放电容量
Public Property Get DischargeCapacity() As Double
    DischargeCapacity = pDischargeCapacity
End Property
Public Property Let DischargeCapacity(value As Double)
    If ValidationEnabled Then
        If value < 0 Or value > 500 Then Err.Raise 5, , "放电容量范围错误"
    End If
    pDischargeCapacity = value
End Property

' 容量保持率
Public Property Get CapacityRetention() As Double
    CapacityRetention = pCapacityRetention
End Property
Public Property Let CapacityRetention(value As Double)
    If ValidationEnabled Then
        If value < 0 Or value > 100 Then Err.Raise 5, , "容量保持率需在0-100%之间"
    End If
    pCapacityRetention = value
End Property

' [...] 其他属性类似，此处省略重复代码

' ========== 高级功能 ==========
' 关闭数据校验（用于导入历史数据）
Public Sub DisableValidation()
    ValidationEnabled = False
End Sub

' 生成数据摘要
Public Function Summary() As String
    Summary = "循环号: " & pCycleNumber & vbCrLf & _
              "样本: " & pSampleID & vbCrLf & _
              "容量保持: " & Format(pCapacityRetention, "0.00") & "%"
End Function

' 数据导出为数组（用于写入工作表）
Public Function ToArray() As Variant
    ToArray = Array(pCycleNumber, pSampleID, pDischargeCapacity, _
                    pCapacityRetention, pChargeCapacity, pDischargeEnergy, _
                    pEnergyRetention, pChargeEnergy, pDCIR, pDCIRGrowthRate)
End Function

' 电压差计算（自定义扩展点）
Public Property Get VoltageDelta() As Double
    VoltageDelta = pOCV - pLoadVoltage
End Property
