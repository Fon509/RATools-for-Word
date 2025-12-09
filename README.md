<div align="center"> <img src="_image\logo.png" alt="Logo" width="400"/> <h3>RATools for Word - 专为药品注册（RA）打造的 Word 效率插件</h3> <p>基于实战经验开发，提升申报资料编写效率，助力注册申报工作更上一层楼。</p> </div>



## 📖 项目简介

本插件基于 `.dotm`（启用宏的模板）定义功能核心，并通过 `.dotx` 文件管理样式模板。旨在解决药品RA人在文档编写中频繁切换选项卡、格式调整繁琐等痛点。

**主要特性：**

- **实战导向**：功能源于开发者在注册申报过程中的长期实践，精准解决痛点。
- **兼容性**：已在 Windows 10 / 11 系统下的 Microsoft 365 Word 环境中测试通过。
- **持续迭代**：后续将根据实际需求，持续添加更多便捷功能。

<div align="center"> <img src="_image\0.png" width=100%/> </div>

## ⚙️ 安装与配置

为了确保插件正常运行，请严格按照以下步骤进行配置。

### 1. 下载与文件准备

- 前往 [Releases](https://github.com/Fon509/RATools/releases) 下载最新的 `.dotm` 和 `.dotx` 文件。
- **创建目录**：建议在 D 盘根目录创建文件夹，结构如下（推荐）：
  - 将 `master-template-cn.dotx` 放入 `D:\RATools\`
  - 将 `RAtools.dotm` 放入 `D:\RATools\Startup\`

> **注意**：您也可以自定义 `.dotx` 文件路径，但后续步骤需对应修改路径配置。本文档以推荐路径 `D:\RATools\Startup` 为例。

<div align="center"> <img src="_image\folder.png" width=75%/> </div>

### 2. 配置 Word 启动路径

1. 打开 Word，点击左上角 **「文件」** -> **「选项」**。
2. 在弹出的对话框中点击 **「高级」**，向下滑动至“常规”栏目，点击 **「文件位置」**。
3. 选中 **「启动」** 项，点击“修改”，选择存放 `RAtools.dotm` 的文件夹路径（例如：`D:\RATools\Startup`）。

<div align="center"> <img src="_image\1.png" width=25%/> <img src="_image\2.png" width=25%/> </div> <div align="center"> <img src="_image\3.png" width=100%/> </div> <div align="center"> <img src="_image\4.png" width=100%/> </div>

### 3. 添加受信任位置

为防止宏被系统安全策略拦截，需将插件目录设为受信任位置：

1. 在 Word 选项对话框左侧，选择 **「信任中心」** -> **「信任中心设置」**。
2. 选择 **「受信任位置」**，点击 **「添加新位置」**。
3. 点击 **「浏览」**，选择插件文件夹路径（`D:\RATools\Startup`），确认无误后点击 **「确定」** 保存所有设置。

<div align="center"> <img src="_image\5.png" width=100%/> </div> 
<div align="center"> <img src="_image\6.png" width=100%/> </div> 
<div align="center"> <img src="_image\7.png" width=50%/> </div>
<div align="center">  <img src="_image\8.png" width=100%/> </div>

## 🚀 使用指南

### 加载模板

安装成功后，Word 顶部会出现 **「RA工具栏」** 选项卡。

1. 点击 **「点击加载」** 按钮即可挂载主样式模板。
2. 成功加载后将提示 **「主模板已附加」**。

> **提示**：如果未将 `master-template-cn.dotx` 放在默认路径（`D:\RATools\`），插件会提示找不到文件。此时点击“是”并手动选择文件位置即可。

<div align="center"> <img src="_image\9.png" width=100%/> </div> 
<div align="center"> <img src="_image\10.png" width=25%/> </div>
<div align="center"> <img src="_image\11.png" width=50%/> </div>



### 功能模块详解

当前版本集成了三大核心模块：

#### 1. 样式快速应用

基于 `.dotx` 定义的标准样式模板，提供一键应用预设样式功能，统一文档格式标准。

#### 2. RA 常用选项

将分散在 Word 不同选项卡中的高频功能（如页面设置、视图切换等）聚合至同一面板，减少鼠标点击与页面切换，显著提升操作流。

#### 3. 增强型宏工具

内置宏列表对话框，比 Word 原生界面更清晰直观。目前包含以下实用宏：

| **功能名称**            | **说明**                                                     |
| ----------------------- | ------------------------------------------------------------ |
| **保护引用域/页码格式** | 用于解决交叉引用刷新后已经调整蓝色字体变黑的问题。           |
| **超链接与域蓝字化**    | 自动查找文中所有超链接和域（排除页码），将其统一设置为蓝色，符合电子申报规范。（如有 Bug 欢迎反馈） |
| **批量转 PDF 与书签**   | 批量将 Word 文档转换为 PDF，并自动根据标题大纲生成 PDF 书签。 |

## ⬆️进阶用法

### 修改.dotx文件实现样式自定义，满足自己独特的样式偏好



### 创建属于自己的宏并添加至宏列表中



## 📝 交流与反馈

如果您在使用过程中遇到问题或有新的功能建议，欢迎提交 Issue 或联系开发者。

## 📅 更新日志

查看版本更新历史，请参阅 [CHANGELOG](https://github.com/Fon509/RATools-for-Word/blob/main/CHANGELOG.md)。
