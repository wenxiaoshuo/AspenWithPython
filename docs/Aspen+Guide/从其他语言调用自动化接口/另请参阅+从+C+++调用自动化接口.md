您可以在 C++ 程序中使用 Aspen Plus 的自动化功能。
像对待任何 COM 对象一样，通过 **vtable 方式** 链接到 Aspen Plus 类型库即可。

代码中应定期调用 **void UpdateIdleStatus()** 函数：

- 更新 GUI 右下角仿真状态消息

- 释放已删除对象占用的内存

建议频率：

- 计算第 1 分钟内每 10 秒调用 1 次

- 之后每 30 秒调用 1 次
更频繁调用不会带来副作用，仅增加极小开销。

**对话框行为**

- 若从未调用 **SuppressDialogs**，则 **Visible 属性** 决定是否弹窗；

- 一旦调用 **SuppressDialogs**，**Visible 对后续运行失效**，由 SuppressDialogs 参数决定处理方式。



