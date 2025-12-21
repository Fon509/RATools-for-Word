Attribute VB_Name = "Mod_BatchConvertWordToPDF"
Option Explicit

' 定义全局变量用于记录日志
Dim processLog As String
Dim successCount As Integer
Dim failCount As Integer

Sub BatchConvertWordToPDF()
    Dim targetType As Integer
    Dim folderPath As String
    Dim selectedFiles As Variant
    Dim i As Integer
    Dim reportDoc As Document
    Dim viewReport As Integer ' 新增变量用于存储用户选择
    
    ' 初始化日志变量
    processLog = "【批量转PDF处理报告】" & vbCrLf & "时间：" & Now & vbCrLf & String(50, "-") & vbCrLf
    successCount = 0
    failCount = 0
    
    ' 优化性能：关闭屏幕更新和警告
    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone
    
    On Error GoTo ErrorHandler

    ' 让用户选择转换类型
    targetType = MsgBox("请选择处理模式：" & vbCrLf & vbCrLf & _
                      "【是 (Yes)】- 转换文件夹（包括所有子文件夹）" & vbCrLf & _
                      "【否 (No)】 - 转换单个或多个Word文件" & vbCrLf & _
                      "【取消 (Cancel)】- 退出宏" & vbCrLf & vbCrLf & _
                      "注意：转换前会自动刷新目录和图表目录。", _
                      vbYesNoCancel + vbQuestion, "增强版批量Word转PDF")
    
    If targetType = vbCancel Then GoTo Cleanup
    
    If targetType = vbYes Then
        ' --- 文件夹模式 ---
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "请选择包含Word文档的文件夹"
            If .Show <> -1 Then GoTo Cleanup
            folderPath = .SelectedItems(1)
        End With
        
        If folderPath <> "" Then
            ProcessFolderWithSubfolders folderPath
        End If
    Else
        ' --- 多文件模式 ---
        With Application.FileDialog(msoFileDialogFilePicker)
            .Title = "请选择一个或多个Word文档"
            .Filters.Clear
            .Filters.Add "Word文档", "*.doc;*.docx;*.docm"
            .AllowMultiSelect = True
            If .Show <> -1 Then GoTo Cleanup
            
            ' 遍历处理选中的文件
            For i = 1 To .SelectedItems.count
                ConvertOneFile .SelectedItems(i)
            Next i
        End With
    End If
    
    ' --- 处理完成，询问是否查看报告 ---
    Application.ScreenUpdating = True ' 恢复屏幕更新以便显示对话框
    
    viewReport = MsgBox("处理完成！" & vbCrLf & _
                        "成功: " & successCount & " 个" & vbCrLf & _
                        "失败: " & failCount & " 个" & vbCrLf & vbCrLf & _
                        "是否生成并查看详细处理报告？", vbYesNo + vbQuestion, "批量转换完成")
    
    If viewReport = vbYes Then
        Set reportDoc = Documents.Add
        With reportDoc.Content
            .Text = processLog & vbCrLf & String(50, "=") & vbCrLf & _
                    "处理完成！" & vbCrLf & _
                    "成功：" & successCount & " 个" & vbCrLf & _
                    "失败：" & failCount & " 个"
            .Font.Name = "微软雅黑"
            .Font.Size = 10
        End With
    End If

Cleanup:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll
    Exit Sub

ErrorHandler:
    MsgBox "发生意外错误: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

' 递归处理文件夹及其所有子文件夹
Sub ProcessFolderWithSubfolders(folderPath As String)
    Dim fso As Object
    Dim mainFolder As Object
    Dim subFolder As Object
    Dim file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set mainFolder = fso.GetFolder(folderPath)
    
    ' 处理当前文件夹中的Word文件
    For Each file In mainFolder.Files
        If IsWordDocument(file.Path) Then
            ConvertOneFile file.Path
        End If
    Next
    
    ' 递归处理所有子文件夹
    For Each subFolder In mainFolder.SubFolders
        ProcessFolderWithSubfolders subFolder.Path
    Next
    
    Set fso = Nothing
    Set mainFolder = Nothing
End Sub

' 【核心功能】处理单个文件：打开 -> 刷新目录 -> 导出 -> 关闭
Sub ConvertOneFile(filePath As String)
    Dim doc As Document
    Dim fso As Object
    Dim pdfFileName As String
    Dim toc As TableOfContents
    Dim tof As TableOfFigures
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 错误捕捉：防止单个文件损坏导致整个宏停止
    On Error GoTo FileError
    
    ' 以不可见方式打开，避免屏幕闪烁，ReadOnly防止修改原文件属性
    Set doc = Documents.Open(fileName:=filePath, Visible:=False, ReadOnly:=True, AddToRecentFiles:=False)
    
    ' --- 刷新目录核心代码 ---
    ' 1. 刷新普通目录
    If doc.TablesOfContents.count > 0 Then
        For Each toc In doc.TablesOfContents
            toc.Update ' 更新整个目录（页码和标题）
        Next toc
    End If
    
    ' 2. 刷新图表目录
    If doc.TablesOfFigures.count > 0 Then
        For Each tof In doc.TablesOfFigures
            tof.Update
        Next tof
    End If
    ' -----------------------
    
    ' 构建PDF路径
    pdfFileName = fso.BuildPath(fso.GetParentFolderName(filePath), _
                               fso.GetBaseName(filePath) & ".pdf")
    
    ' 导出PDF
    doc.ExportAsFixedFormat _
        OutputFileName:=pdfFileName, _
        ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=False, _
        OptimizeFor:=wdExportOptimizeForPrint, _
        CreateBookmarks:=wdExportCreateHeadingBookmarks, _
        DocStructureTags:=True
    
    ' 关闭文档（不保存对原Word的更改，因为我们只是为了生成PDF才更新了目录）
    doc.Close SaveChanges:=wdDoNotSaveChanges
    
    ' 记录成功日志
    successCount = successCount + 1
    processLog = processLog & "[成功] " & fso.GetFileName(filePath) & vbCrLf
    
    GoTo Finally

FileError:
    ' 记录失败日志
    failCount = failCount + 1
    processLog = processLog & "[失败] " & fso.GetFileName(filePath) & " - 原因: " & Err.Description & vbCrLf
    If Not doc Is Nothing Then doc.Close SaveChanges:=wdDoNotSaveChanges

Finally:
    Set doc = Nothing
    Set fso = Nothing
End Sub

' 检查文件是否为Word文档
Function IsWordDocument(filePath As String) As Boolean
    Dim fso As Object
    Dim ext As String
    Dim fileName As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    ext = LCase(fso.GetExtensionName(filePath))
    fileName = fso.GetFileName(filePath)
    
    ' 排除临时文件(~$开头) 并检查扩展名
    If (ext = "doc" Or ext = "docx" Or ext = "docm") And Left(fileName, 2) <> "~$" Then
        IsWordDocument = True
    Else
        IsWordDocument = False
    End If
    
    Set fso = Nothing
End Function

