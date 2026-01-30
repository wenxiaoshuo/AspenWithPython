Aspen Plus 应用对象 **HappLS** 支持一个**出向事件接口** **IAPHappEvent**，包含 HappLS 类的各类事件。
有关每个事件的详情，参见 **IAPHappEvent 事件接口**。

在 **Visual Basic 或 VBA** 中，只需在**窗体或类模块**里用 **WithEvents** 声明 HappLS 变量，然后将该变量指向模拟问题即可接收事件；
务必在适当时候释放对象。

在 **C++** 中，需实现一个支持 **IAPHappEvent**（dispinterface）的对象，并通过连接点接口把出向接口连接到应用对象。

下方代码片段展示了一个 **VB 自动化客户端**，用于在窗体文本框中显示控制面板消息：

```VBScript
Dim WithEvents sim As Happ.IHapp
Dim fCP As frmControlPanel   ' 带文本框的窗体

Private Sub LoadSimulation(fn As String)
  On Error GoTo ErrorHandler
  
  Set fCP = New frmControlPanel
  LabelProblem.Caption = "loading " & fn
  MousePointer = vbHourglass
  
  ' 避免“服务器忙”提示
  App.OleRequestPendingTimeout = 100000
  
  Set sim = CreateObject("Apwn.Document")
  Call sim.InitFromArchive2(fn, 0)
  sim.Visible = False
  
  App.OleRequestPendingTimeout = 10000
  MousePointer = vbDefault
  Exit Sub
  
ErrorHandler:
  Debug.Print Err.Description
  Err.Clear
  LabelProblem.Caption = "failed to load problem"
  Set sim = Nothing
  MousePointer = vbDefault
  App.OleRequestPendingTimeout = 10000
End Sub

' 事件：控制面板消息
Private Sub Sim_OnControlPanelMessage(ByVal Clear As Long, ByVal msg As String)
  If Clear Then
    fCP.txtMsg.Text = ""
  Else
    fCP.txtMsg = fCP.txtMsg.Text & vbCrLf & msg
  End If
  On Error GoTo ErrHandler
  Exit Sub
  
ErrHandler:
  Debug.Print "Error " & Err.Description
  Err.Clear
End Sub
```

Python 没有内建“WithEvents”语法，但可以用 `win32com.client.WithEvents` 把 Aspen Plus 事件绑定到自定义类。
下面给出完整可运行示例：加载模拟 → 实时把控制面板消息打印到终端（或你自己的 GUI）。

```VBScript
import win32com.client as win32
import pythoncom, sys, os

# ---------- 事件处理类 ----------
class AspenEvents:
    """处理 Aspen Plus 出向事件 IAPHappEvent"""
    def OnControlPanelMessage(self, clear, msg):
        """每次控制面板输出时被调用"""
        if clear:
            os.system('cls' if os.name == 'nt' else 'clear')
        else:
            print(msg)          # 也可写入 QTextEdit / tkinter Text 等

# ---------- 主程序 ----------
def main():
    pythoncom.CoInitialize()     # 多线程公寓
    ap = win32.DispatchWithEvents("AspenPlus.Document.34.0", AspenEvents)
    ap.Visible = False

    # 加载现有模拟（路径自定）
    bkp_path = r"D:\Demo\flash.bkp"
    ap.InitFromArchive2(bkp_path, 0)

    # 运行模拟，控制面板消息会自动触发 OnControlPanelMessage
    ap.Engine.Run2()

    input("按 Enter 退出...\n")   # 保持运行，可看实时消息
    ap.Quit()
    pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()
```

使用要点

1. `DispatchWithEvents` 把 COM 出向接口绑定到 `AspenEvents` 类；

2. 类方法名必须与 COM 事件名 **完全一致**（大小写敏感）：`OnControlPanelMessage`；

3. 运行模拟后，所有控制面板输出会实时触发你的 Python 函数，可用于日志、进度条或前端展示。

