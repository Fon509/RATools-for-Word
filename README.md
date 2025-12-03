<div align="center">
    <img src="_image\logo.png" alt="Logo" width="400"/>
    <p>为药品注册（RA）人开发的Word插件 - 提升资料编写效率，更好地完成注册申报</p>
</div>


本插件采用 .dotm 文件定义功能，并通过 .dotx 文件指定样式模板。目前已在 Windows 10 及 Windows 11 系统下的 Microsoft 365 - Word 中测试通过。

功能主要来自于开发者注册申报过程中的长期实践，后续也会添加更多在资料编写过程中觉得方便的功能。

<div align="center">
    <img src="_image\0.png" width=100%/>
</div>

# 配置方法

下载[release](https://github.com/Fon509/RATools/releases)中的.dotm和.dotx文件

在D盘中创建「RAtools」文件夹，并将下载完成的master-template-cn.dotx放在该路径下

建议在其中新建名为「Startup」的文件夹，将RAtools.dotm放入「Startup」文件夹中（也可以自行选择其他路径，后续步骤需相应调整）

<div align="center">
    <img src="_image\folder.png" width=75%/>
</div>


打开Word，点击左上角「 文件」

<div align="center">
    <img src="_image\1.png" width=25%/>
</div>


点击「选项」

<div align="center">
    <img src="_image\2.png" width=25%/>
</div>


点击「高级」，右侧滚动找到「常规」，点击「文件位置」

<div align="center">
    <img src="_image\3.png" width=100%/>
</div>


选择「启动」，双击修改为放入RAtools.dotm的文件夹路径（D:\RATools\Startup）

<div align="center">
    <img src="_image\4.png" width=100%/>
</div>


在Word选项对话框左侧，选择「信任中心」，点击「信任中心设置」

<div align="center">
    <img src="_image\5.png" width=100%/>
</div>


在信任中心对话框左侧，选择「受信任位置」，在右侧点击「添加新位置」。

<div align="center">
    <img src="_image\6.png" width=100%/>
</div>


点击「浏览」，选择放入RAtools.dotm的文件夹路径（D:\RATools\Startup），点击「确定」

<div align="center">
    <img src="_image\7.png" width=50%/>
</div>


确认无误后，点击「确定」

<div align="center">
    <img src="_image\8.png" width=100%/>
</div>

# 使用方法

## 加载模版

在「RA工具栏」选项卡中点击「点击加载」按钮加载主模板

<div align="center">
    <img src="_image\9.png" width=100%/>
</div>

点击后会提示「主模板已附加」

<div align="center">
    <img src="_image\10.png" width=25%/>
</div>

如master-template-cn.dotx未放在D:\RATools\路径下，点击后则会提示默认位置找不到主模板，会提示手动选择。点击「是」后选择master-template-cn.dotx即可

<div align="center">
    <img src="_image\11.png" width=50%/>
</div>

## 功能介绍

当前版本主要功能分为三部分

1. 快捷应用预设样式：通过.dotx文件定义样式模板，点击按钮进行

   具体细节暂不介绍，请自行摸索

2. RA常用功能集合：来源于Word中不同选项卡下的功能按钮，将其集合到同一选项卡中，避免来回切换选项卡，提高效率

3. 通过宏实现更多功能扩展（宏后续会调整为通过统一对话框调用，目前比较散乱）

   目前内置以下3个宏

   1. 保护引用域和页码引用域格式：主要用于交叉引用刷新后设置的蓝字字体颜色还原为黑色
   2. 将超链接和域设置为蓝色：点击后将会自动找到所有超链接和域（交叉引用等）设置为蓝色，排除了页码（有bug请反馈）
   3. 批量将Word转为PDF，并通过标题创建目录

## 进阶用法

### 修改.dotx文件实现样式自定义，满足自己独特的样式偏好



### 创建属于自己的宏

