Run2 接口方法 **IHapp.Run2** 与 **IHapp.Engine.Run2** 均接受一个可选参数，用于将模拟设为**异步运行**。这样，自动化客户端在等待模拟完成的同时可以继续执行其他任务。
当应用程序可见时，**应始终使用异步方式**运行 Aspen Plus。

下方示例展示了在类模块事件过程中以异步方式启动引擎的代码。

AspenPlus.cls

```VBScript
VERSION 1.0 CLASS

BEGIN

MultiUse = -1 'True 

Persistable = 0 'NotPersistable 

DataBindingBehavior = 0 'vbNone 

DataSourceBehavior = 0 'vbNone 

MTSTransactionMode = 0 'NotAnMTSObject 

END

Attribute VB_Name = "AspenPlus"

Attribute VB_GlobalNameSpace = False

Attribute VB_Creatable = True

Attribute VB_PredeclaredId = False

Attribute VB_Exposed = False

Option Explicit

 

'An example class module to illustrate the use of the Aspen

'Plus automation interface.

'

'This example is intended for illustration purposes only.

'It is not intended as production code.

 

 

' This module requires a reference to the Aspen Plus

' type library

 

 

' The interface to the Aspen Plus Automation Server

Private WithEvents simulationObject As Happ.HappLS

Attribute simulationObject.VB_VarHelpID = -1

Private isRunning As Boolean

 

Public Event RunFinished()

 

 

 

Private Sub Class_Initialize()

On Error Resume Next 

' Initialize data members 

isRunning = False 

End Sub

 

Private Sub Class_Terminate()

On Error Resume Next 

' Release automation server 

Set simulationObject = Nothing 

End Sub

 

Public Property Get simulation() As Happ.IHapp

On Error GoTo GetSimulationErr 

Set simulation = simulationObject 

Exit Property 

 

' simple error handler

GetSimulationErr:

Set simulationObject = Nothing 

 

End Property

 

Public Property Let simulation(ByVal vNewValue As Happ.IHapp)

On Error GoTo SetSimulationErr 

Set simulationObject = vNewValue 

 

Exit Property 

' simple error handler

SetSimulationErr:

Set simulationObject = Nothing 

 

End Property

 

Public Sub Run()

On Error GoTo RunError 

If simulationObject Is Nothing Then 

' Throw user defined exception 

Exit Sub 

End If 

If isRunning = True Then 

' Throw user defined exception 

Exit Sub 

End If 

If InputComplete = False Then 

' Throw user defined exception 

Exit Sub 

End If 

 

' Do asynchronous run 

isRunning = True 

simulationObject.Run2 True 

Exit Sub 

 

RunError:

isRunning = False 

End Sub

 

' Event recieved when the run is completed

Private Sub simulationObject_OnCalculationCompleted()

On Error Resume Next 

If isRunning = True Then 

isRunning = False 

RaiseEvent RunFinished 

End If 

End Sub

 

 

 

Public Property Get InputComplete() As Boolean

On Error GoTo InputCompleteErr 

Dim problemData As Happ.IHNode 

Dim completionMask As Long 

 

InputComplete = False 

If simulationObject Is Nothing Then 

' Throw user defined exception 

 

Exit Property 

End If 

 

Set problemData = simulationObject.Tree.Data 

completionMask = _
problemData.AttributeValue(Happ.HAP_COMPSTATUS) 

If completionMask And Happ.HAP_INPUT_COMPLETE Then 

InputComplete = True 

End If 

 

Exit Property 

InputCompleteErr:

InputComplete = False 

End Property

 

 

 

Public Sub LoadBkp(ByVal problem As String)

Dim newSimulation As New HappLS 

On Error GoTo LoadBkpErr 

' load bkp with local engine 

Call newSimulation.InitFromArchive2(problem, 0) 

simulation = newSimulation 

Set newSimulation = Nothing 

Exit Sub 

 

LoadBkpErr:

Err.Clear 

' log error or throw user defined exception

End Sub
```

Mainform.frm

```VBScript
VERSION 5.00

Begin VB.Form Form1

Caption = "Form1" 

ClientHeight = 3195 

ClientLeft = 60 

ClientTop = 345 

ClientWidth = 4680 

LinkTopic = "Form1" 

ScaleHeight = 3195 

ScaleWidth = 4680 

StartUpPosition = 3 'Windows Default 

Begin VB.CommandButton Load  

Caption = "Load" 

Height = 375 

Left = 720 

TabIndex = 1 

Top = 480 

Width = 1095 

End 

Begin VB.CommandButton Run  

Caption = "Run" 

Height = 375 

Left = 720 

TabIndex = 0 

Top = 1200 

Width = 1095 

End 

End

Attribute VB_Name = "Form1"

Attribute VB_GlobalNameSpace = False

Attribute VB_Creatable = False

Attribute VB_PredeclaredId = True

Attribute VB_Exposed = False

Dim WithEvents problem As AspenPlus

 

 

Private Sub Form_Load()

On Error Resume Next

' Initialize variables and set state 

Set problem = New AspenPlus 

Load.Enabled = True 

Run.Enabled = False 

End Sub

 

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next

Set problem = Nothing 

End Sub

 

 

Private Sub Load_Click()

On Error Resume Next

MousePointer = vbHourglass 

Load.Enabled = False 

' load file into class 

problem.LoadBkp ("C:\Aspen Plus V14\GUI\Examples\pfdtut.bkp") 

Run.Enabled = True 

MousePointer = vbDefault 

End Sub

 

Private Sub problem_RunFinished()

On Error Resume Next 

MousePointer = vbDefault 

' Re-enable user controls 

Run.Enabled = True 

End Sub

 

Private Sub Run_Click()

On Error Resume Next 

' Disable user controls 

Run.Enabled = False 

MousePointer = vbHourglass 

problem.Run 

End Sub


```

Python翻译

AspenPlus.cls
一个示例类模块，演示如何调用 Aspen Plus 自动化接口。
仅用于教学，非生产代码。

- 使用 WithEvents 声明 simulationObject，可接收计算完成等事件

- 提供 LoadBkp、Run 等公共方法，内部调用异步 Run2

- 若输入未完成或正在运行，将提前退出

- 计算完成后触发 RunFinished 事件

Mainform.frm
简单窗体：两个按钮（Load / Run）

- Load 按钮调用 LoadBkp 加载 bkp 文件

- Run 按钮调用异步 Run2，计算完成后通过事件重新启用按钮

（以下保留 VB 原码，仅注释与说明已中文化）

【Python 完整直译版】

aspen_plus.py

```VBScript
import win32com.client as win32
import pythoncom

class AspenPlus:
    """与 VB AspenPlus 类功能一致：封装加载、运行、事件"""
    _reg_clsid_ = pythoncom.CreateGuid()
    _reg_desc_ = "Aspen Plus 自动化封装（异步 Run2）"
    _reg_progid_ = "AspenPlus.Wrapper"

    def __init__(self):
        pythoncom.CoInitialize()
        self.simulation_object = None
        self.is_running = False

    def __del__(self):
        self.simulation_object = None
        pythoncom.CoUninitialize()

    @property
    def simulation(self):
        return self.simulation_object

    @simulation.setter
    def simulation(self, v):
        self.simulation_object = v

    def load_bkp(self, problem_path):
        new_sim = win32.Dispatch("AspenPlus.Document.34.0")
        new_sim.InitFromArchive2(problem_path, 0)
        self.simulation = new_sim

    @property
    def input_complete(self):
        if self.simulation is None: return False
        comp_mask = self.simulation.Tree.Data.AttributeValue(12)  # HAP_COMPSTATUS
        return bool(comp_mask & 0x20)                             # HAP_INPUT_COMPLETE

    def run(self):
        if self.simulation is None or self.is_running or not self.input_complete:
            return
        self.is_running = True
        self.simulation.Run2(True)          # True = 异步

    def OnCalculationCompleted(self):
        if self.is_running:
            self.is_running = False
            self.run_finished()

    def run_finished(self):
        """供外部窗体重写"""
        pass
```

main_gui.py（tkinter 等效窗体）

```VBScript
import tkinter as tk
from tkinter import ttk, messagebox
from aspen_plus import AspenPlus

class MainForm(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Aspen Plus 异步运行示例")
        self.geometry("240x160")
        self.problem = AspenPlus()
        self.problem.run_finished = self.on_run_finished   # 绑定事件

        self.load_btn = ttk.Button(self, text="Load", command=self.load_click)
        self.run_btn  = ttk.Button(self, text="Run",  command=self.run_click, state=tk.DISABLED)
        self.load_btn.pack(pady=10)
        self.run_btn.pack(pady=10)

    def load_click(self):
        self.config(cursor="watch")
        self.load_btn.config(state=tk.DISABLED)
        try:
            self.problem.load_bkp(r"C:\Aspen Plus V14\GUI\Examples\pfdtut.bkp")
            self.run_btn.config(state=tk.NORMAL)
        except Exception as e:
            messagebox.showerror("Load Error", str(e))
        finally:
            self.config(cursor="")
            self.load_btn.config(state=tk.NORMAL)

    def run_click(self):
        self.config(cursor="watch")
        self.run_btn.config(state=tk.DISABLED)
        try:
            self.problem.run()
        except Exception as e:
            messagebox.showerror("Run Error", str(e))

    def on_run_finished(self):
        """事件：计算完成"""
        self.config(cursor="")
        self.run_btn.config(state=tk.NORMAL)

if __name__ == "__main__":
    MainForm().mainloop()
```

