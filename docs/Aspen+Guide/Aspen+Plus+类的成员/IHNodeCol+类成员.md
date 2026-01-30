### 一、标准 VB 属性（只读）

表格

复制

|名称|Python 用法|说明|
|-|-|-|
|**Application**|`col.Application`|返回所属 HappLS 对象|
|**Parent**|`col.Parent`|返回父节点|

---

### 二、主导航：按位置/标签取子节点

表格

复制

|名称|Python 用法|说明|
|-|-|-|
|**Item(*loc_or_name*)**|`col.Item(0)` 或 `col.Item("B6")`|按序号或标签取子节点；多维时可传多个参数|

---

### 三、增删改行/元素

表格

复制

|名称|Python 用法|说明|
|-|-|-|
|**Add(...) **|`col.Add("B6!RADFRAC")`|新建并返回子节点；格式 `"名!类型"` 或序号|
|**Insert(node, ...)**|`col.Insert(new_node, 2)`|把已有节点插入指定位置|
|**Remove(...)**|`col.Remove("B6")`|删除子节点；返回被删节点|
|**InsertRow(dim, loc)**|`col.InsertRow(0, 3)`|在第 dim 维第 loc 行插入空行|
|**RemoveRow(dim, loc)**|`col.RemoveRow(0, 3)`|删除第 dim 维第 loc 行|

---

### 四、关键维度/标签/行数

表格

复制

|名称|Python 用法|说明|
|-|-|-|
|**Dimension**|`col.Dimension`|维度数（0=标量列表）|
|**RowCount(dim)**|`col.RowCount(0)`|指定维的行数|
|**DimensionName(dim)**|`col.DimensionName(0)`|维度的显示名称（如 "Stage"）|
|**Label(dim, loc)**|`col.Label(0, 5)`|指定维指定行的标签字符串|
|**LabelLocation(label, dim)**|`col.LabelLocation("IC4", 1)`|根据标签返回行号|
|**IsNamedDimension(dim)**|`col.IsNamedDimension(0)`|该维行是否用标签命名|
|**ItemName(loc, [dim])**|`col.ItemName(3, 0)`|返回指定位置元素的名称/行名|

---

### 五、标签节点与标签属性

表格

复制

|名称|Python 用法|说明|
|-|-|-|
|**LabelNode(dim, loc)**|`col.LabelNode(0, 5)`|返回标签本身对应的 IHNode（可改标签文本）|
|**LabelAttribute(dim, loc, attr)**|`col.LabelAttribute(0, 5, HAP_UNITROW)`|取标签行的属性值|
|**LabelAttributeType(dim, loc, attr)**|`col.LabelAttributeType(0, 5, HAP_UNITROW)`|取标签行的属性类型|

---

### 六、其他

表格

复制

|名称|Python 用法|说明|
|-|-|-|
|**Count**|`col.Count`|总元素个数（所有维乘积）|

---

### 七、Python 速用模板

Python

复制

```Python
HAP_UNITROW = 3
col = ap.Tree.Data.Blocks.B6.Output.B_TEMP.Elements   # 以塔温为例

# 1. 遍历塔板
for loc in range(col.RowCount(0)):
    stage_lbl = col.Label(0, loc)          # "1", "2", ...
    temp_node = col.Item(loc)              # IHNode
    print(stage_lbl, temp_node.Value, temp_node.UnitString)

# 2. 插入空塔板（行）
col.InsertRow(0, 3)   # 在第 3 行前插入新空行

# 3. 删除塔板（行）
col.RemoveRow(0, 15)  # 删除第 15 行
```

