作者:`lovemefan`
时间:`2017/10/18`
# Excel 宏命令处理
&nbsp;&nbsp;&nbsp;&nbsp;在我们工作的时候,常常会遇到大量的数据处理.人工操作费时费力,于是笔者开始百度看看有没有什么自动化的方法,最后发现了excel可以使用宏指令来实现编程式处理.
## Excel VBA
> 
要熟悉VBA，那么就不得不先了解宏和它们之间的关系。
VBA被称为宏语言，但是它决不能和宏划上等号，宏是一条指令或者几条指令的集合，控制WORD或EXCEL等执行一连串的操作。而VBA则是不折不扣的高级语言，通过面向对象的方法来实现不能实现的功能。在编辑一个宏的时候，visual basic 编辑器会将这个宏记录为一个VBA过程
VBA使应用程序具有生命的特征，以适应不同的环境，不同的应用，其主要表现为定制、自动化、协作化。
VBA的特点是将VB语言与应用对象模型结合起来，处理各种应用需求。WORD VBA则是将VB应用于Word对象模型，或者说是用VB语言来操控这些Word对象模型，以达到各种应用的要求。所以，如果你想通过VBA控制WORD或EXCEL，必须同时熟悉VB语言和Word对象模型。

## 如何使用VBA
视图-宏
![VBA](http://oskhhyaq3.bkt.clouddn.com/blog/171018/hHHBK2h4eh.png?imageslim)
![mark](http://oskhhyaq3.bkt.clouddn.com/blog/171018/F7CBBaFhH3.png?imageslim)

如果不能使用宏,可能你的excel禁用了.启动即可
启用方式:
`文件-选项-信任中心-信任中心设置-宏设置-启用所有宏`
配合录制宏就可以试着写代码了
