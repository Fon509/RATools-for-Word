Attribute VB_Name = "mRibbon"
Option Explicit

'=====================  模 块 级 变 量  =====================
Private mRibbon       As IRibbonUI     '缓存 Ribbon
Private mMainTemplate As Template      '缓存本模板（RAtools.dotm）

'常 量
Private Const DEF_STYLE_FILE As String = "D:\RAtools\master-template-cn.dotx"
Private Const DEF_RIBBON_TMPL As String = "RAtools.dotm"

'=====================  Ribbon 必 要 回 调  =====================
'Ribbon OnLoad
Public Sub Onload(ribbon As IRibbonUI)
    Set mRibbon = ribbon
End Sub

'=====================  附 加 主 模 板  =====================
Public Sub AttachTemplate(ByVal control As IRibbonControl)
    Dim tmplPath As String
    tmplPath = GetStyleFilePath
    If tmplPath = "" Then Exit Sub
    
    On Error GoTo ErrH
    With ActiveDocument
        .UpdateStylesOnOpen = True
        .AttachedTemplate = tmplPath
    End With
    MsgBox "主模板已附加！", vbInformation
    Exit Sub
ErrH:
    MsgBox "模板附加失败！", vbCritical
End Sub

'=====================  段 落 样 式  =====================
Public Sub btnStyle_Click(ByVal control As IRibbonControl)
    On Error GoTo ErrH
    ApplyStyle control.tag
    Exit Sub
ErrH:
    HandleStyleErr
End Sub

'=====================  字 符 样 式  =====================
Public Sub btnChar_Click(ByVal control As IRibbonControl)
    On Error GoTo ErrH
    Dim s As String: targetStyle = control.tag
    If Selection.Style = targetStyle Then targetStyle = "正文-F"
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

'================  对齐方式  ================
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
                    "超链接一键蓝字", _
                    "智能遍历文档，将所有超链接和域（REF/PAGEREF等）的颜色设置为蓝色，但在处理过程中会自动排除图表题注。")
      
    ' 第2个
    items.Add Array("Wrapper_RunAddMergeFormat", _
                    "域格式保护", _
                    "扫描选区内的引用域，自动添加 \* MERGEFORMAT 开关，防止更新域后格式丢失。")
                    
    ' 第3个
    items.Add Array("BatchConvertWordToPDF", _
                    "Word批量转PDF", _
                    "批量将Word转为PDF，并通过Word标题创建PDF书签")

                    
    ' 如果以后要加新宏，直接复制粘贴即可，无需修改其他地方
    ' 如果需要 control 参数的宏，需要下面做一个 Wrapper，见下面Wrapper包装器下的内容，同时需要在上面添加
    
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
