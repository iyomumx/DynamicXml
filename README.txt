这是一个寂寞的人用VB写的动态对象模型XML库

核心思想是继承DynamicObject，对各种成员调用提供支持。
目前只有一个类：XDynamic
C#看到两个类，其中的另一个是Module……
XDynamic原本是对XElement的动态包装，现在改为对List(Of XElement)的包装
对TryGetIndex/TrySetIndex调用List，对Member则调用第一个元素。

欢迎提意见