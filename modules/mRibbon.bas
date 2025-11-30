Attribute VB_Name = "mRibbon"
Option Explicit

'=====================  模 块 级 变 量  =====================
Private mRibbon       As IRibbonUI     '缓存 Ribbon
Private mMainTemplate As Template      '缓存本模板（RA工具栏.dotm）

'常 量
Private Const DEF_STYLE_FILE As String = "D:\RAtools\主模板.dotx"
Private Const DEF_RIBBON_TMPL As String = "RA工具栏.dotm"

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
'确保 mMainTemplate 已指向 F工具栏.dotm
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


'应用样式 + 补 MERGEFORMAT
Private Sub ApplyStyle(ByVal styleName As String)
    Selection.Style = ActiveDocument.Styles(styleName)
    AddMergeFormat
End Sub

'为选区内 REF/PAGEREF 加 \* MERGEFORMAT,保护域格式
Private Sub AddMergeFormat()
    Dim fld As field, rng As Range
    For Each fld In Selection.Fields
        If fld.Type = wdFieldRef Or fld.Type = wdFieldPageRef Then
            Set rng = fld.Code
            If InStr(1, rng.Text, "mergeformat", vbTextCompare) = 0 Then
                rng.Text = rng.Text & " \* MERGEFORMAT "
                fld.Update
            End If
        End If
    Next fld
End Sub

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

'=== 智能设置超链接和域为蓝色（避免不必要的格式更改）===
Sub SetHyperlinksAndFieldsToBlue()
    Dim hyperlink As hyperlink
    Dim field As field
    Dim storyRange As Range
    Dim countChanged As Integer
    
    Application.ScreenUpdating = False
    countChanged = 0
    
    ' 处理超链接
    For Each hyperlink In ActiveDocument.Hyperlinks
        If hyperlink.Range.Font.Color <> RGB(0, 0, 255) Then
            hyperlink.Range.Font.Color = RGB(0, 0, 255)
            countChanged = countChanged + 1
        End If
    Next hyperlink
    
    ' 处理域，但排除题注
    For Each storyRange In ActiveDocument.StoryRanges
        countChanged = countChanged + ProcessFieldsExcludeCaptions(storyRange)
    Next storyRange
    
    Application.ScreenUpdating = True
    
    If countChanged > 0 Then
        MsgBox "已将 " & countChanged & " 个超链接和域设置为蓝色", vbInformation
    Else
        MsgBox "所有超链接和域已经是蓝色", vbInformation
    End If
End Sub

Private Function ProcessFieldsExcludeCaptions(rng As Range) As Integer
    Dim field As field
    Dim count As Integer
    Dim fieldCode As String
    
    count = 0
    For Each field In rng.Fields
        ' 获取域代码并检查是否是题注
        fieldCode = LCase(field.Code.Text)
        
        ' 排除题注相关的域（SEQ域和包含"图"、"表"等题注关键词的域）
        If Not IsCaptionField(fieldCode) Then
            If field.Result.Font.Color <> RGB(0, 0, 255) Then
                field.Result.Font.Color = RGB(0, 0, 255)
                count = count + 1
            End If
        End If
    Next field
    
    ' 处理链接的范围
    Do While Not (rng.NextStoryRange Is Nothing)
        Set rng = rng.NextStoryRange
        For Each field In rng.Fields
            fieldCode = LCase(field.Code.Text)
            If Not IsCaptionField(fieldCode) Then
                If field.Result.Font.Color <> RGB(0, 0, 255) Then
                    field.Result.Font.Color = RGB(0, 0, 255)
                    count = count + 1
                End If
            End If
        Next field
    Loop
    
    ProcessFieldsExcludeCaptions = count
End Function

' 判断是否为题注域的函数
Private Function IsCaptionField(fieldCode As String) As Boolean
    ' 常见的题注域标识
    Dim captionIndicators As Variant
    captionIndicators = Array("seq", "图", "表", "chart", "figure", "table", "caption")
    
    Dim indicator As Variant
    For Each indicator In captionIndicators
        If InStr(1, fieldCode, indicator, vbTextCompare) > 0 Then
            IsCaptionField = True
            Exit Function
        End If
    Next indicator
    
    IsCaptionField = False
End Function

Public Sub SetHyperlinksAndFieldsToBlueRibbon(control As IRibbonControl)
    SetHyperlinksAndFieldsToBlue
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
