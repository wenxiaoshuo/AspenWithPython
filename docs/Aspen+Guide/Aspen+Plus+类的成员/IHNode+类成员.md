### 一、标准 VB 属性（只读）

表格

复制

|名称|Python 用法|说明|
|-|-|-|
|**Application**|`node.Application`|返回所属 HappLS 对象|
|**Parent**|`node.Parent`|返回父节点|

---

### 二、访问子结构

表格

复制

|名称|Python 用法|说明|
|-|-|-|
|**Dimension**|`node.Dimension`|维度数；0=标量|
|**Elements**|`node.Elements`|子节点集合（IHNodeCol）|

---

### 三、数据值/属性/单位（读写）

表格

复制

|名称|Python 用法|说明|
|-|-|-|
|**Value**|`node.Value = 3.14`|当前值；类型见 ValueType|
|**ValueType**|`node.ValueType`|0 未定义 1=int 2=real 3=string 4=node|
|**ValueForUnit(row, col)**|`node.ValueForUnit(2, 3)`|按指定单位取值|
|**UnitString**|`node.UnitString`|当前单位符号串|
|**AttributeValue(n, [force])**|`node.AttributeValue(HAP_UNITROW)`|取属性值（HAP_* 枚举）|
|**AttributeType(n)**|`node.AttributeType(HAP_UNITROW)`|返回该属性类型（1-5）|
|**HasAttribute(n)**|`node.HasAttribute(HAP_UNITROW)`|节点是否支持该属性|
|**SetValueAndUnit(val, col)**|`node.SetValueAndUnit(3.14, 3)`|同时写值与单位列|
|**SetValueUnitAndBasis(v, c, b)**|`node.SetValueUnitAndBasis(3.14, 3, "MOLE")`|值+单位+基准一次写|

---

### 四、导航/助手

表格

复制

|名称|Python 用法|说明|
|-|-|-|
|**FindNode(path)**|`ap.Tree.FindNode(r"\Data\Blocks\B6\Input\NSTAGE")`|按 Variable Explorer 路径跳转|
|**Name**|`node.Name`|节点名称|

---

### 五、增删改

表格

复制

|名称|Python 用法|说明|
|-|-|-|
|**Clear()**|`node.Clear()`|清空节点内容|
|**Delete()**|`node.Delete()`|删除该元素|
|**RemoveAll()**|`collection.RemoveAll()`|清空集合|

---

### 六、批量校核（Reconcile）

**用途**：让 Aspen 根据输入自动补齐/修正流股或模块的 T/P/组成、总流量等。
**适用节点**：单个流股、模块、Streams 集合、整个 Tree.Data。

#### 1. 常用“对象范围”标志

表格

复制

|常量（hex）|含义|
|-|-|
|`0x2`|仅更新有输入规格的流股|
|`0x400000`|仅更新进料流股|
|`0x1000000`|仅更新流股（层次/全流表）|
|`0x2000000`|仅非流股（模块）|
|`0x4000000`|仅撕裂流股|

#### 2. 常用“变量策略”标志

表格

复制

|常量（hex）|含义|
|-|-|
|`0x1`|按输入规格|
|`0x4`|T+P 作状态变量|
|`0x8`|T+V/F 作状态变量|
|`0x10`|P+V/F 作状态变量|
|`0x20`|按组分流量更新流量/组成|
|`0x40`|按总流+摩尔分数更新|
|`0x80`|MIXED 子物流 摩尔基准|
|`0x100`|MIXED 子物流 质量基准|
|`0x200`|MIXED 子物流 标准体积基准|

#### 3. 对话框抑制

表格

复制

|常量（hex）|含义|
|-|-|
|`0x100000`|静默（不弹任何窗）|
|`0x200000`|不弹警告/错误窗|

#### 4. Python 一次性调用示例

Python

复制

```Python
HAPP_RECONCILE_INPUT   = 0x1
HAPP_RECONCILE_ONLY    = 0x2
HAPP_RECONCILE_QUIET   = 0x100000

# 让 Aspen 仅根据已有输入补齐所有流股
ap.Tree.Data.Streams.Reconcile(HAPP_RECONCILE_INPUT |
                               HAPP_RECONCILE_ONLY |
                               HAPP_RECONCILE_QUIET)
```

即可在后台完成 T/P/流量/组成的自动推算，无需手动填写。

