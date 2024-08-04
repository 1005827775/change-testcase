#背景：
平时写用例习惯用xmind书写，效率快一些，但用例需要以excel模板的格式导入到devops平台中统一管理，故写一个小工具实现xmind用例向excel用例转换

#文件结构：
ChangeXmind.py #实现xmind文件转换成excel文件
Template.py #excel文件样式
ChangeCase.py #使用tkinter库生成可视化组件
ChangeExcel.py #excel用例转换成xmind(未完成）

#使用说明：
见用例模板.xmind
标签1、2、3代表执行优先级

#调整代码后执行：
pyinstaller --onefile --windowed D:\Project\xmind转excel用例\pythonProject\pythonProject\ChangeCase.py
生成.exe文件，位置：./xmind转excel用例/pythonProject/pythonProject/dist/
