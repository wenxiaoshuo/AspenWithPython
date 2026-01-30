**效果区别取决于节点是“输入”还是“输出”：**

- **输出值**
改 `HAP_UNITCOL` 会把 **取出的结果** 自动换算成新单位，相当于在 GUI 的 **Results** 表单上换单位——数值会变，但底层储存不变。

- **输入值**
改 `HAP_UNITCOL` 只会改 **输入框的计量单位**，不会帮你把已有数值换算过去，相当于在 GUI 的 **Input** 表单上换单位——原数字保留，含义随之改变。





以下子程序演示如何 **直接修改 Aspen Plus 运行中的单位**：

```VBScript
Sub UnitsChangeExample(ByVal ihAPsim As Happ.IHapp)
Dim ihPres As Happ.IHNode
On Error GoTo ErrorHandler
    ' 取得 B3 模块输出压力节点
    Set ihPres = ihAPsim.Tree.Data.Blocks.B3.Output.B_PRES
    
    ' 显示默认单位下的压力
    MsgBox "Pressure in default units: " & ihPres.Value & _
           Chr(9) & ihPres.UnitString
    
    ' 把单位改为 bar（第 5 列）
    ihPres.AttributeValue(Happ.HAPAttributeNumber.HAP_UNITCOL, True) = 5
    
    ' 再次读取，已自动换算成 bar
    MsgBox "Pressure in selected units: " & ihPres.Value & _
           Chr(9) & ihPres.UnitString
Exit Sub
ErrorHandler:
    MsgBox "UnitsChangeExample raised error " & Err.Description
End Sub
```

要点：

- 第二个参数 `True` 表示 **允许就地修改**（可写）。

- 修改后 **`Value` 会自动换算**成新单位，等同于在 Results 表单换单位。

