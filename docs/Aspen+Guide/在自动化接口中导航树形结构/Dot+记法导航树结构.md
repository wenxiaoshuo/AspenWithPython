可以把 `Item`（或 `Elements`）一路链下去，例如：

vb

复制

```VBScript
Set ihNStageNode = ihAPsim.Elements("Data") _
                  .Elements("Blocks") _
                  .Elements("B6") _
                  .Elements("Input") _
                  .Elements("NSTAGE")
```

更简洁的写法是直接使用“点分”路径：

vb

复制

```VBScript
Set ihNStageNode = ihAPsim.Tree.Data.Blocks.B6.Input.NSTAGE
```

注意：

- 只有节点名符合编程语言标识符规则（无空格、无特殊字符、非保留字）时才可用这种简写。
例：名字里带空格 `"Unit Table"` 或 VB 保留字 `"Value"` 时，必须退回 `Elements("Unit Table")` 或 `Elements("VALUE")` 形式。

- 以下节点类型**不支持**点分简写：
connection、port、setting table、route、label、unit table。
访问这些节点仍需显式使用 `Elements("...")` 或 `FindNode`。

