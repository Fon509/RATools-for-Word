Attribute VB_Name = "mRibbon"
Option Explicit

'=====================  模 块 级 变 量  =====================
Private mRibbon       As IRibbonUI     '缓存 Ribbon

'=====================  配 置 区 域  =====================
' 请根据您的实际环境修改此路径
Private Const DEF_STYLE_FILE As String = "D:\RAtools\master-template-cn.dotx"
Private Const DEF_RIBBON_TMPL As String = "RAtools.dotm"

' 定义要匹配的自定义后缀
Private Const TARGET_SUFFIX As String = "-F"

'=====================  Ribbon 必 要 回 调  =====================
'Ribbon OnLoad
Public Sub Onload(ribbon As IRibbonUI)
    Set mRibbon = ribbon
End Sub

'=====================  导 入 样 式 (核 心 逻 辑)  =====================
' 说明：双重导入策略，支持 -F 后缀及目录样式
Public Sub AttachTemplate(ByVal control As IRibbonControl)
    Dim tmplPath As String
    Dim sourceDoc As Document
    Dim currentDoc As Document
    Dim sty As Style
    Dim stylesList As New Collection ' 使用集合暂存样式名
    Dim vStyleName As Variant
    Dim importCount As Integer
    Dim pass As Integer
    Dim sName As String
    
    ' 1. 获取路径 (复用 mRibbon 原有的路径获取函数)
    tmplPath = GetStyleFilePath
    If tmplPath = "" Then Exit Sub
    
    Set currentDoc = ActiveDocument
    
    Application.ScreenUpdating = False ' 关闭屏幕刷新
    
    ' 2. 后台打开模版并筛选样式
    ' 以只读、不可见的方式打开模版文件
    Set sourceDoc = Documents.Open(fileName:=tmplPath, ReadOnly:=True, Visible:=False)
    
    On Error Resume Next
    
    ' 遍历模版中的样式，建立“待导入名单”
    For Each sty In sourceDoc.Styles
        sName = sty.NameLocal
        
        ' 判断逻辑：
        ' 1. 名字以 "-F" 结尾
        ' 2. 或者 名字以 "TOC" 开头 (兼容 TOC 1, TOC 2...)
        ' 3. 或者 名字包含 "图表目录"或 "Table of Figures"
        If (UCase(Right(sName, Len(TARGET_SUFFIX))) = UCase(TARGET_SUFFIX)) Or _
           (UCase(Left(sName, 3)) = "TOC") Or _
           (InStr(sName, "图表目录") > 0) Or _
           (InStr(sName, "Table of Figures") > 0) Then
            
            ' 将符合条件的样式名加入集合
            stylesList.Add sName
        End If
    Next sty
    
    ' 如果没有找到任何样式，直接退出
    If stylesList.count = 0 Then
        sourceDoc.Close SaveChanges:=wdDoNotSaveChanges
        Application.ScreenUpdating = True
        MsgBox "模版中没有找到符合条件（-F 或 TOC）的样式。", vbExclamation
        Exit Sub
    End If
    
    ' 3. 执行“双重导入”策略
    ' 第一遍：创建样式实体（此时基于 BaseOn 的链接可能会断裂）
    ' 第二遍：覆盖样式定义（修复链接关系）
    For pass = 1 To 2
        For Each vStyleName In stylesList
            Application.OrganizerCopy _
                Source:=sourceDoc.FullName, _
                Destination:=currentDoc.FullName, _
                Name:=vStyleName, _
                Object:=wdOrganizerObjectStyles
        Next vStyleName
    Next pass
    
    ' 4. 清理工作
    sourceDoc.Close SaveChanges:=wdDoNotSaveChanges
    Set sourceDoc = Nothing
    
    Application.ScreenUpdating = True
    
    ' 5. 反馈结果
    MsgBox "操作完成！" & vbCrLf & _
           "已成功导入 " & stylesList.count & " 个样式。", vbInformation, "导入成功"
End Sub

'=====================  应 用 样 式  =====================
Private Sub ApplyStyle(ByVal styleName As String)
    Selection.Style = ActiveDocument.Styles(styleName)
End Sub

'=====================  段 落 样 式  =====================
Public Sub btnStyle_Click(ByVal control As IRibbonControl)
    On Error GoTo ErrH
    ApplyStyle control.Tag
    Exit Sub
ErrH:
    HandleStyleErr
End Sub

'=====================  字 符 样 式  =====================
Public Sub btnChar_Click(ByVal control As IRibbonControl)
    On Error GoTo ErrH
    Dim s As String: targetStyle = control.Tag
    If Selection.Style = targetStyle Then targetStyle = "正文-F" ' 重复点击后撤销样式
    ApplyStyle targetStyle
    Exit Sub
ErrH:
    HandleStyleErr
End Sub

'=====================  一 键 大 写  =====================
Public Sub btnCap_Click(ByVal control As IRibbonControl)
    On Error Resume Next
    Selection.Range.Case = wdUpperCase
End Sub

'=====================  私 有 过 程  =====================
'确保 mMainTemplate 已指向 RATools.dotm
Private Function EnsureMainTemplate() As Boolean
    If mMainTemplate Is Nothing Then
        Dim t As Template
        For Each t In Templates
            If StrComp(t.Name, DEF_RIBBON_TMPL, vbTextCompare) = 0 Then
                Set mMainTemplate = t
                Exit For
            End If
        Next
    End If
    EnsureMainTemplate = Not mMainTemplate Is Nothing
    If Not EnsureMainTemplate Then _
        MsgBox "请先加载 " & DEF_RIBBON_TMPL, vbCritical
End Function

'为选区内 REF/PAGEREF 加 \* MERGEFORMAT,保护域格式
'包含智能判断(全文/选区) + 结果弹窗
Private Sub AddMergeFormat()
    Dim fld As field, rng As Range
    Dim targetFields As Fields ' 目标域集合
    Dim msgTip As String       ' 提示信息

    ' 判断：如果是光标插入点(wdSelectionIP)则处理全文，否则处理选区
    If Selection.Type = wdSelectionIP Then
        Set targetFields = ActiveDocument.Fields
        msgTip = "未选中文字，已对【全文】域代码进行格式保护。"
    Else
        Set targetFields = Selection.Fields
        msgTip = "已对【选中区域】域代码进行格式保护。"
    End If

    ' 遍历处理
    For Each fld In targetFields
        If fld.Type = wdFieldRef Or fld.Type = wdFieldPageRef Then
            Set rng = fld.Code
            If InStr(1, rng.Text, "mergeformat", vbTextCompare) = 0 Then
                rng.Text = rng.Text & " \* MERGEFORMAT "
                fld.Update
            End If
        End If
    Next fld

    ' 操作完成后弹出提示
    MsgBox msgTip, vbInformation, "操作完成"
End Sub

' 功能区按钮回调：调用上面的处理过程
Sub RunAddMergeFormat(control As IRibbonControl)
    AddMergeFormat
End Sub

'样式错误统一提示
Private Sub HandleStyleErr()
    If Err.Number = 5941 Or Err.Number = 91 Then
        MsgBox "请先加载主模板 dotx！", vbExclamation
    Else
        MsgBox "样式应用失败：" & Err.Description, vbCritical
    End If
End Sub

'取主模板路径（默认/浏览）
Private Function GetStyleFilePath() As String
    If Dir(DEF_STYLE_FILE) <> "" Then
        GetStyleFilePath = DEF_STYLE_FILE
        Exit Function
    End If
    
    If MsgBox("默认位置找不到主模板，是否手动选择？", vbYesNo + vbQuestion) = vbNo Then Exit Function
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Word 模板", "*.dot;*.dotx;*.dotm"
        If .Show = -1 Then GetStyleFilePath = .SelectedItems(1)
    End With
End Function

'================  下拉选择对齐方式  ================
'================  下拉菜单：左对齐  ================
Public Sub AlignLeft_Click(control As IRibbonControl)
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
End Sub

'================  顶部大按钮：直接设为居中  ================
Public Sub AlignCenter_Click(control As IRibbonControl)
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
End Sub

'================  下拉菜单：右对齐  ================
Public Sub AlignRight_Click(control As IRibbonControl)
    Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
End Sub

'================  下拉菜单：两端对齐  ================
Public Sub AlignJustify_Click(control As IRibbonControl)
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
End Sub


' 显示/隐藏样式管理窗格
Public Sub ShowStylePane(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    ' 尝试使用内置命令打开样式窗格
    Application.CommandBars.ExecuteMso "StylesPane"
    
    Exit Sub
    
ErrorHandler:
    ' 如果内置命令失败，使用快捷键
    SendKeys "%^{+}s", True
End Sub

'=====================  宏 列 表 管 理  =====================

' 1. Ribbon 回调：点击按钮弹出窗体
' 在 Ribbon XML 中，将按钮的 onAction 指向这个 Sub
Public Sub ShowMacroListWindow(control As IRibbonControl)
    frmMacroList.Show
End Sub


' 2. 供窗体调用的数据源函数
' 返回值：Variant 数组
Public Function GetMyMacroRegistry() As Variant
    Dim items As New Collection
    Dim vArr() As Variant
    Dim i As Long
    
    ' ================= 配置区域：Array(英文代码, 简短名称, 详细描述) =================
    
    ' 格式：items.Add Array("宏代码名", "列表显示的名称", "下方显示的详细介绍")
    
    ' 第1个
    items.Add Array("SetHyperlinksAndFieldsToBlue", _
                    "超链接和域批量设置为蓝色", _
                    "智能遍历文档，将所有超链接和域（REF/PAGEREF等）的颜色设置为蓝色，但在处理过程中会自动排除图表题注和页码。")
      
    ' 第2个
    items.Add Array("Wrapper_RunAddMergeFormat", _
                    "域格式保护", _
                    "扫描全文或选区内的引用域，自动添加 \* MERGEFORMAT 开关，防止更新域后格式丢失。")
                    
    ' 第3个
    items.Add Array("BatchConvertWordToPDF", _
                    "Word批量转PDF", _
                    "批量将Word转为PDF，并通过Word标题创建PDF书签。")
    
    ' 第4个
    items.Add Array("BatchRenameFiles", _
                    "批量修改文件名", _
                    "批量修改文件名" & vbCrLf & _
                    "1. 仅保留汉字、小写字母、数字、中划线和下划线" & vbCrLf & _
                    "2. 空格将被直接删除，大写字母会替换为小写字母，其他非法字符替换为中划线 ""-""" & vbCrLf & _
                    "3. 支持“文件夹模式”和“多文件选择模式”" & vbCrLf & _
                    "4. 如果文件被占用无法重命名，自动创建改名后的副本")
    
    ' 第5个
    items.Add Array("ConvertHeadingNumbers", _
                    "标题自动编号转文本", _
                    "将文档中所有标题（大纲 1-9 级）的自动编号转换为固定的静态文本。")
    
    ' 第6个
    items.Add Array("RenameCurrentDocument", _
                    "重命名当前文件", _
                    "无需关闭文件，直接重命名当前文件。")
    
    ' 第7个
    items.Add Array("BatchSetMargins", _
                    "一键设置页边距", _
                    "一键将单个或多个文件页面上、下、左、右的页边距设置为 2.54厘米（即标准的 1 英寸）。")
    
    ' 第8个
    items.Add Array("BatchAutoFitTablesToWindow", _
                    "批量表格自动调整表格", _
                    "将文档中所有表格批量设置为“根据窗口自动调整”")
                    
    ' 如果以后要加新宏，直接复制粘贴即可，无需修改其他地方
    ' 如果需要control参数的宏，需要下面做一个Wrapper，见下面Wrapper包装器下的内容，同时需要在上面添加
    
    ' ================= 配置结束 =================
    
    If items.count > 0 Then
        ReDim vArr(0 To items.count - 1)
        For i = 1 To items.count
            vArr(i - 1) = items(i)
        Next i
        GetMyMacroRegistry = vArr
    Else
        GetMyMacroRegistry = Empty
    End If
End Function

'=====================  Wrapper 包装器  =====================
' 解释：因为很多宏是 Ribbon 回调 (带 control 参数)，
' Application.Run 无法自动提供 control 参数，直接运行会报错。
' 所以我们需要一些不带参数的“外壳”过程。

Public Sub Wrapper_RunAddMergeFormat()
    ' 调用原有的逻辑
    ' 注意：因为原 Sub 需要 control 参数，我们传 Nothing 进去
    ' 只要原 Sub 内部没用到 control.ID 或 control.Tag，这样写就是安全的
    RunAddMergeFormat Nothing
End Sub

