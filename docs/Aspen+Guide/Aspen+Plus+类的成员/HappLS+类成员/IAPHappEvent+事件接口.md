|名称与参数|成员类型|只读|说明|
|-|-|-|-|
|**OnDialogSuppressed(msg As String, result As String)**|Event||当警告信息被抑制未显示，或本应弹窗时被触发。
**msg**：被抑制的警告文本或描述信息。
**result**：消息框的默认按钮/选项值。|
|**OnControlPanelMessage(clear As Boolean, msg As String)**|Event||控制面板产生消息时触发。
**clear**=True 时表示清空面板；否则 **msg** 为要显示的新行文本。|
|**OnGUIClosing()**|Event||用户交互关闭主窗口时触发。可用于通知自动化客户端断开连接；**不**表示服务器已退出，仅 GUI 关闭。|
|**OnDataChanged(pObj As Object)**|Event||数据被修改时触发，可用来检查输入完整性。
**pObj**：被修改的对象（通常为 IHNode）。|

**与事件接口相关的属性**

|名称与参数|成员类型|只读|说明|
|-|-|-|-|
|**SuppressDialogs As Integer**|Property||是否抑制消息/对话框。
自动化启动时默认为 **True（非零）**；若随后供交互使用，应设为 **False（零）**。|

