' =====================
' BatteryInfo 类
' =====================
' 功能：管理电池的基本信息，包括电池名称和颜色属性
' 用途：用于在图表中标识和区分不同的电池
' =====================
Option Explicit

'私有成员变量
' m_batteryName - 存储电池的名称标识
' m_batteryColor - 存储电池在图表中的显示颜色（使用Long类型表示RGB值）
Private m_batteryName As String
Private m_batteryColor As Long

' =====================
' 构造函数
' =====================
' 功能：初始化BatteryInfo对象的成员变量
' 说明：设置默认值 - 空名称和默认颜色值0
' =====================
Private Sub Class_Initialize()
    m_batteryName = ""
    m_batteryColor = 0
End Sub

' =====================
' 电池名称属性
' =====================
' 功能：提供对电池名称的读写访问
' 属性：
'   Get - 获取电池名称
'   Let - 设置电池名称
' =====================
Public Property Get BatteryName() As String
    BatteryName = m_batteryName
End Property

Public Property Let BatteryName(ByVal value As String)
    m_batteryName = value
End Property

' =====================
' 电池颜色属性
' =====================
' 功能：提供对电池显示颜色的读写访问
' 属性：
'   Get - 获取电池颜色值
'   Let - 设置电池颜色值
' 说明：颜色值使用Long类型存储RGB值
' =====================
Public Property Get BatteryColor() As Long
    BatteryColor = m_batteryColor
End Property

Public Property Let BatteryColor(ByVal value As Long)
    m_batteryColor = value
End Property