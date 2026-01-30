### 一、基本运行控制

表格

复制

|名称|Python 用法|说明|
|-|-|-|
|**Run2([async])**|`ap.Engine.Run2(True)`|运行模拟；True=异步（GUI 可见时必须异步）|
|**Stop()**|`ap.Engine.Stop()`|立即停止计算|
|**Reinit([obj_type, obj_id])**|`ap.Engine.Reinit()`|重初始化；无参会弹窗，有参则定向重初始化|
|**ReinitializeEO([hierarchy])**|`ap.Engine.ReinitializeEO()`|重建 EO 模型并用当前结果重初始化|
|**SynchronizeEO([hierarchy])**|`ap.Engine.SynchronizeEO()`|仅把配置变更同步到 EO 侧|
|**MoveTo(obj_type, obj_id)**|`ap.Engine.MoveTo(IAP_MOVETO_BLOCK, "B6")`|把运行断点移到指定对象（调试/分步）|

---

### 二、断点（Stop Points）操纵

表格

复制

|名称|Python 用法|说明|
|-|-|-|
|**AddStopPoint(type, id, before_or_after)**|`ap.Engine.AddStopPoint(IAP_STOPPOINT_BLOCK, "B6", 1)`|1=before 2=after|
|**ClearStopPoints()**|`ap.Engine.ClearStopPoints()`|清空所有断点|
|**DeleteStopPoint(index)**|`ap.Engine.DeleteStopPoint(1)`|1-based 索引|
|**StopPointCount**|`cnt = ap.Engine.StopPointCount`|当前断点个数|
|**GetStopPoint(index, type, id, before_or_after)**|`ap.Engine.GetStopPoint(1, type, id, before)`|读断点信息（出参）|

---

### 三、运行选项/引擎文件设置

表格

复制

|名称|Python 用法|说明|
|-|-|-|
|**OptionSettings(type)**|`ap.Engine.OptionSettings(IAP_RUN_OPTION_XXX)`|读运行选项布尔值|
|**EngineFilesSettings(file)**|`ap.Engine.EngineFilesSettings(IAP_ENGINEFILES_LOG)`|取引擎文件路径/设置|

---

### 四、常用枚举速抄（happ.tlb 抄）

Python

复制

```Python
IAP_REINIT_SIMULATION = 0      # 整案重初始化
IAP_MOVETO_BLOCK      = 3      # MoveTo 目标类型
IAP_STOPPOINT_BLOCK   = 0      # 断点类型
IAP_STOPPOINT_BEFORE  = 1      # before
IAP_STOPPOINT_AFTER   = 2      # after
```

---

### 五、Python 完整模板

Python

复制

```Python
import win32com.client as win32
import time

ap = win32.Dispatch("AspenPlus.Document.34.0")
ap.InitFromArchive2(r"pfdtut.bkp", 0)

# 1. 异步运行 + 实时断点
ap.Engine.AddStopPoint(0, "B6", 1)   # 在 B6 前停
ap.Engine.Run2(True)                 # True = 异步

while ap.Engine.IsRunning:
    time.sleep(5)
    ap.Engine.UpdateIdleStatus()     # 必须定期调，释放内存+更新状态条

# 2. 从断点继续
ap.Engine.MoveTo(3, "B6")            # 移到 B6 开始处
ap.Engine.Run2(True)                 # 继续跑

# 3. 结果
ap.SaveAs(r"pfdtut_new.bkp")
ap.Quit()
```

**要点小结**

1. **Run2(True)** 必须配 **UpdateIdleStatus()** 循环，否则内存不释放 + 状态条卡死。

2. **Reinit / ReinitializeEO / SynchronizeEO** 三套重初始化按需求选。

3. 断点 API 与 GUI 的“分步运行”完全对应，可编程实现调试流程。

