下面把 **IHAPLibRef** 所有“库管理 + 模型选项板分类”成员一次性列完，并给出 **Python 速查+用法**。

---

### 一、库（Library）操纵

表格

复制

|名称|Python 用法|说明|
|-|-|-|
|**CountLibs**|`n = libref.CountLibs`|当前引用库数量|
|**LibraryName(index)**|`name = libref.LibraryName(0)`|第 index 个库的名字（0 起）|
|**LibraryPath(index)**|`path = libref.LibraryPath(0)`|返回完整路径字符串|
|**InsertLibrary(path, loc)**|`disp = libref.InsertLibrary(r"f:\fcc.apm", 0)`|在指定位置插入库，返回显示名|
|**RemoveLibrary(loc)**|`libref.RemoveLibrary(2)`|删除第 loc 个引用|
|**MoveLibrary(from, to)**|`libref.MoveLibrary(2, 0)`|把库从位置 2 移到 0|

#### Python 快速遍历示例

Python

复制

```Python
libref = ap.LibRef
for i in range(libref.CountLibs):
    print(i, libref.LibraryName(i), "→", libref.LibraryPath(i))
```

---

### 二、模型选项板分类（Category）操纵

表格

复制

|名称|Python 用法|说明|
|-|-|-|
|**CategoryName(index)**|`cat = libref.CategoryName(0)`|第 index 个分类名称|
|**CategorySelected(name)**|`sel = libref.CategorySelected("Reactors")`|1=已勾选显示，0=未显示|
|**CategoryLocSelected(index)**|`sel = libref.CategoryLocSelected(0)`|同上，按序号|
|**MoveCategory(fr, to)**|`libref.MoveCategory(3, 0)`|把分类从位置 3 移到 0|

#### Python 勾选/取消分类示例

Python

复制

```Python
# 让“Reactors”分类在模型选项板可见
libref.CategorySelected("Reactors") = 1        # 勾选
# 或按序号
# libref.CategoryLocSelected(5) = 1
```

---

### 三、完整模板：插入自定义库并置顶

Python

复制

```Python
import win32com.client as win32

ap = win32.Dispatch("AspenPlus.Document.34.0")
ap.InitNew2()                      # 新建空白模拟

libref = ap.LibRef

# 1. 把自定义库插入到首位（索引 0）
new_lib = r"D:\MyLibs\fcc.apm"
disp_name = libref.InsertLibrary(new_lib, 0)
print("已插入库：", disp_name)

# 2. 把“Reactors”分类移到选项板最前
react_idx = None
for i in range(libref.CountCategories):   # 遍历找 Reactors
    if libref.CategoryName(i) == "Reactors":
        react_idx = i
        break
if react_idx is not None and react_idx != 0:
    libref.MoveCategory(react_idx, 0)     # 移到首位

# 3. 保存带库引用的新模拟
ap.SaveAs(r"D:\MyLibs\with_fcc.bkp")
ap.Quit()
```

运行后打开 `with_fcc.bkp`，可在模型选项板最前端看到自定义库，且“Reactors”分类已被置顶。

