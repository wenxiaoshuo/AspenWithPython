|名称与参数|成员类型|只读|说明|
|-|-|-|-|
|Activate()|Sub||激活应用程序|
|Application As HappLS|Property|是|返回对象的所属应用程序|
|FullName As String|Property|是|返回应用程序的完整路径名|
|Name As String|Property|是|返回应用程序的名称|
|Parent As Happ|Property|是|返回对象的创建者|
|Visible As Boolean|Property||返回/设置应用程序的可见状态|
|UpdateIdleStatus()|Sub||更新主窗口右下角的模拟状态，并释放已删除对象占用的内存。应在重写 OnIdle() 的函数中调用，否则状态不会更新且内存不会被释放。|

