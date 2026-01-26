Attribute VB_Name = "Mod_BatchDetectHighlights"
Option Explicit

' ==========================================
' 宏名称：BatchDetectHighlights
' 功能：批量检测文档是否包含高亮（突出显示颜色）
' ==========================================

' 定义全局 FSO 对象，避免重复创建
Private fso As Object

Sub BatchDetectHighlights()
    Dim userChoice As String
    Dim reportDoc As Document
    Dim doc As Document
    Dim targetFiles As Collection
    Dim filePath As Variant
    Dim fDialog As FileDialog
    Dim folderPath As String
    Dim startTime As Double
    
    ' 初始化 FSO
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 1. 获取用户选择的模式 (默认值为 1)
    userChoice = InputBox("请输入数字选择检测模式：" & vbCrLf & vbCrLf & _
                          "1 - 检测【当前打开】的文件" & vbCrLf & _
                          "2 - 选择【多个文件】进行批量检测" & vbCrLf & _
                          "3 - 选择【文件夹】（包含子文件夹）检测所有 Word 文件", _
                          "高亮检测工具", "1")
    
    If userChoice = "" Then Exit Sub ' 用户取消
    
    ' 初始化文件集合
    Set targetFiles = New Collection
    
    ' 2. 根据选择处理逻辑
    Select Case userChoice
        Case "1" ' 当前文件
            If Documents.count = 0 Then
                MsgBox "当前没有打开的文档！", vbExclamation
                Exit Sub
            End If
            
            If CheckForHighlight(ActiveDocument) Then
                MsgBox "当前文档【包含】突出显示颜色。", vbInformation
            Else
                MsgBox "当前文档【不包含】突出显示颜色。", vbInformation
            End If
            Exit Sub ' 模式1不需要生成报告，直接结束
            
        Case "2" ' 选择多个文件
            Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
            With fDialog
                .Title = "请选择要检测的 Word 文件"
                .Filters.Clear
                .Filters.Add "Word 文档", "*.docx; *.doc; *.docm"
                .AllowMultiSelect = True
                If .Show = -1 Then
                    For Each filePath In .SelectedItems
                        targetFiles.Add filePath
                    Next filePath
                Else
                    Exit Sub
                End If
            End With
            
        Case "3" ' 选择文件夹（包含子文件夹）
            Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
            With fDialog
                .Title = "请选择包含 Word 文件的文件夹"
                If .Show = -1 Then
                    folderPath = .SelectedItems(1)
                    ' 调用递归函数扫描文件夹
                    Call RecursiveScan(folderPath, targetFiles)
                Else
                    Exit Sub
                End If
            End With
            
        Case Else
            MsgBox "输入无效，请输入 1、2 或 3。", vbCritical
            Exit Sub
    End Select
    
    ' 3. 如果没有找到文件
    If targetFiles.count = 0 Then
        MsgBox "未找到需要处理的文件。", vbExclamation
        Exit Sub
    End If
    
    ' 4. 开始批量处理
    Application.ScreenUpdating = False ' 关闭屏幕更新（防止界面刷新）
    startTime = Timer
    
    ' 创建报告文档
    Set reportDoc = Documents.Add
    With reportDoc.Content
        .InsertAfter "突出显示颜色检测报告" & vbCrLf
        .InsertAfter "检测时间: " & Now & vbCrLf
        .InsertAfter "总计文件: " & targetFiles.count & vbCrLf & vbCrLf
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Font.Bold = True
        .Font.Size = 14
        .InsertParagraphAfter
    End With
    
    ' 插入表格头
    Dim tbl As Table
    Set tbl = reportDoc.Tables.Add(Range:=reportDoc.Characters.Last, NumRows:=1, NumColumns:=2)
    tbl.Borders.Enable = True
    tbl.Cell(1, 1).Range.Text = "文件路径"
    tbl.Cell(1, 2).Range.Text = "检测结果"
    
    ' 循环处理文件
    Dim hasHighlight As Boolean
    
    For Each filePath In targetFiles
        ' 容错处理：如果文件无法打开（如损坏或正在使用），跳过并记录
        On Error Resume Next
        
        ' Visible:=False 彻底静默打开，不再闪烁
        Set doc = Documents.Open(fileName:=filePath, Visible:=False, ReadOnly:=True, AddToRecentFiles:=False)
        
        If Err.Number <> 0 Then
            tbl.Rows.Add
            ' 错误时也显示完整路径，方便排查
            tbl.Cell(tbl.Rows.count, 1).Range.Text = filePath
            tbl.Cell(tbl.Rows.count, 2).Range.Text = "无法打开文件"
            tbl.Cell(tbl.Rows.count, 2).Range.Font.Color = wdColorGray50
            Err.Clear
        Else
            ' 检测
            hasHighlight = CheckForHighlight(doc)
            
            ' 写入报告表格
            tbl.Rows.Add
            
            If hasHighlight Then
                ' 有高亮：添加超链接并显示红色结果
                reportDoc.Hyperlinks.Add Anchor:=tbl.Cell(tbl.Rows.count, 1).Range, _
                                         Address:=filePath, _
                                         TextToDisplay:=filePath
                                         
                tbl.Cell(tbl.Rows.count, 2).Range.Text = "包含高亮"
                tbl.Cell(tbl.Rows.count, 2).Range.Font.Color = wdColorRed
                tbl.Cell(tbl.Rows.count, 2).Range.Font.Bold = True
            Else
                ' 无高亮：仅显示纯文本路径
                tbl.Cell(tbl.Rows.count, 1).Range.Text = filePath
                
                tbl.Cell(tbl.Rows.count, 2).Range.Text = "无"
                tbl.Cell(tbl.Rows.count, 2).Range.Font.Color = wdColorGreen
            End If
            
            doc.Close SaveChanges:=wdDoNotSaveChanges
        End If
        On Error GoTo 0
    Next filePath
    
    ' 清理与结束
    Set fso = Nothing
    Application.ScreenUpdating = True
    tbl.AutoFitBehavior (wdAutoFitWindow)
    
    MsgBox "检测完成！共扫描 " & targetFiles.count & " 个文件。" & vbCrLf & "耗时 " & Format(Timer - startTime, "0.00") & " 秒。", vbInformation
End Sub

' ==========================================
' 辅助过程：递归扫描文件夹
' ==========================================
Sub RecursiveScan(ByVal folderPath As String, ByRef fileCollection As Collection)
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object
    Dim ext As String
    
    On Error Resume Next ' 防止权限错误导致中断
    Set folder = fso.GetFolder(folderPath)
    
    ' 遍历当前文件夹下的文件
    For Each file In folder.Files
        ext = LCase(fso.GetExtensionName(file.Path))
        ' 检查扩展名，并排除临时文件（~$开头）
        If (ext = "docx" Or ext = "doc" Or ext = "docm") And Left(file.Name, 2) <> "~$" Then
            fileCollection.Add file.Path
        End If
    Next file
    
    ' 递归遍历子文件夹
    For Each subFolder In folder.SubFolders
        Call RecursiveScan(subFolder.Path, fileCollection)
    Next subFolder
    On Error GoTo 0
End Sub

' ==========================================
' 辅助函数：检测单个文档是否包含高亮
' ==========================================
Function CheckForHighlight(doc As Document) As Boolean
    Dim rng As Range
    
    CheckForHighlight = False
    
    ' 遍历文档的所有“故事类型”（StoryRanges）
    ' 这包括正文、页眉、页脚、文本框、脚注、尾注、批注等
    For Each rng In doc.StoryRanges
        Do

            ' 1. wdNoHighlight (0) = 无高亮
            ' 2. wdUndefined (9999999) = 混合状态（说明部分有高亮，部分没有）-> 判定为包含
            ' 3. 其他颜色值 (如 wdYellow) = 全区域都是该颜色 -> 判定为包含
            
            If rng.HighlightColorIndex <> wdNoHighlight Then
                CheckForHighlight = True
                Exit Function
            End If
            
            ' 如果当前区域有链接的部分（例如长文本框或多个页眉），继续检测下一部分
            Set rng = rng.NextStoryRange
        Loop While Not rng Is Nothing
    Next rng
End Function

