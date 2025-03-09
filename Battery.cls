' ======================
' 类模块：Battery
' ======================
Option Explicit

Private pTestOrderID As String    ' 测试单号
Private pCellSerial As String     ' 电芯序号
Private pBarcode As String        ' 电芯条码
Private pCycles As Collection     ' 循环数据集合

Private Sub Class_Initialize()
    Set pCycles = New Collection
End Sub

' 属性访问器
Public Property Get TestOrderID() As String
    TestOrderID = pTestOrderID
End Property
Public Property Let TestOrderID(Value As String)
    pTestOrderID = Value
End Property

Public Property Get CellSerial() As String
    CellSerial = pCellSerial
End Property
Public Property Let CellSerial(Value As String)
    pCellSerial = Value
End Property

Public Property Get Barcode() As String
    Barcode = pBarcode
End Property
Public Property Let Barcode(Value As String)
    pBarcode = Value
End Property

Public Sub AddCycle(cd As CycleData)
    pCycles.Add cd
End Property

Public Property Get Cycles() As Collection
    Set Cycles = pCycles
End Property