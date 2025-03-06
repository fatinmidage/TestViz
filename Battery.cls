' 存储电池样本及其所有循环数据
Private pBatteryID As String
Private pCycles As Collection

' 初始化时创建集合
Private Sub Class_Initialize()
    Set pCycles = New Collection
End Sub

' 电池编号属性
Public Property Get BatteryID() As String
    BatteryID = pBatteryID
End Property
Public Property Let BatteryID(Value As String)
    pBatteryID = Value
End Property

' 添加循环数据
Public Sub AddCycleData(cd As CycleData)
    pCycles.Add cd
End Sub

' 获取循环数据集合
Public Property Get Cycles() As Collection
    Set Cycles = pCycles
End Property
