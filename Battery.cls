' ======================
' 类模块：Battery
' ======================
' 用途：表示单个电池对象，包含电池的基本信息和循环数据
' 属性：
'   - CellName: 电池名称/编号
'   - Cycles: 电池的循环数据集合
' ======================
Option Explicit

' 私有变量声明
Private pCellName As String        ' 存储电池名称/编号
Private pCycles As Collection     ' 存储电池的循环数据集合

' 构造函数
' 用途：初始化电池对象，创建循环数据集合
Public Sub Class_Initialize()
    Set pCycles = New Collection
End Sub
' 属性：CellName
' 用途：获取或设置电池名称/编号
' 参数：
'   - value: 要设置的电池名称/编号
Public Property Get CellName() As String
    CellName = pCellName
End Property
Public Property Let CellName(ByVal value As String)
    pCellName = value
End Property
' 属性：Cycles
' 用途：获取或设置电池的循环数据集合
' 参数：
'   - value: 要设置的循环数据集合
Public Property Get Cycles() As Collection
    Set Cycles = pCycles
End Property
Public Property Let Cycles(ByVal value As Collection)
    Set pCycles = value
End Property