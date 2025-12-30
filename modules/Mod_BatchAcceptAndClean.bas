Attribute VB_Name = "Mod_BatchAcceptAndClean"
Option Explicit

' =============================================
' 宏名称：BatchAcceptAndClean
' 功能：将tracking版转换为clean版，接受所有修订并停止修订同时删除文档中的所有批注
' =============================================

' === 主程序入口 ===
Sub BatchAcceptAndClean()
    Dim strMode As String
    Dim folderPath As String
    Dim fileCollection As New Collection
    Dim fileItem As Variant
    Dim i As Long
    Dim processedCount As Integer
    
    ' 1. 模式选择
    strMode = InputBox("请输入模式编号：" & vbCrLf & vbCrLf & _
                       "1 - 【当前文档】处理当前打开的文档" & vbCrLf & _
                       "2 - 【文件模式】选择单个或多个文件" & vbCrLf & _
                       "3 - 【文件夹模式】递归处理文件夹", _
                       "WordCleaner", "1")
    
    If StrPtr(strMode) = 0 Or strMode = "" Then Exit Sub
    
    ' 2. 性能设置
    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone
    
    On Error GoTo ErrorHandler
    processedCount = 0
    
    Select Case strMode
        Case "1"
            If Documents.count > 0 Then
                Call DeepCleanDocument(ActiveDocument)
                processedCount = 1
                MsgBox "当前文档处理完成！", vbInformation
            End If
            
        Case "2"
            With Application.FileDialog(msoFileDialogFilePicker)
                .Title = "选择Word文档"
                .AllowMultiSelect = True
                .Filters.Clear
                .Filters.Add "Word文档", "*.doc; *.docx; *.docm", 1
                If .Show = -1 Then
                    For Each fileItem In .SelectedItems
                        Call ProcessFile(CStr(fileItem))
                        processedCount = processedCount + 1
                    Next
                End If
            End With
            
        Case "3"
            With Application.FileDialog(msoFileDialogFolderPicker)
                .Title = "选择根文件夹"
                If .Show = -1 Then
                    folderPath = .SelectedItems(1)
                    Application.StatusBar = "正在扫描文件..."
                    Call RecursiveFindFiles(folderPath, fileCollection)
                    
                    If fileCollection.count > 0 Then
                        For i = 1 To fileCollection.count
                            Application.StatusBar = "正在处理 [" & i & "/" & fileCollection.count & "]"
                            Call ProcessFile(fileCollection(i))
                            processedCount = processedCount + 1
                        Next
                    Else
                        MsgBox "未找到Word文档", vbExclamation
                    End If
                End If
            End With
            
        Case Else
            MsgBox "无效输入", vbExclamation
    End Select

ExitHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll
    Application.StatusBar = False
    If processedCount > 0 And strMode <> "1" Then
        MsgBox "处理完成！共处理 " & processedCount & " 个文件。", vbInformation
    End If
    Exit Sub

ErrorHandler:
    MsgBox "错误: " & Err.Description, vbCritical
    Resume ExitHandler
End Sub

' === 模块：处理单个文件 ===
Sub ProcessFile(filePath As String)
    Dim doc As Document
    On Error Resume Next
    
    ' 后台模式打开
    Set doc = Documents.Open(fileName:=filePath, Visible:=False, AddToRecentFiles:=False)
    
    If Err.Number = 0 Then
        ' 关键修复：确保文档完全加载后再处理
        DoEvents
        Call DeepCleanDocument(doc)
        doc.Save
        doc.Close
    Else
        Debug.Print "打开失败: " & filePath
    End If
    On Error GoTo 0
End Sub

' === 模块：深度清理逻辑（核心修复） ===
Sub DeepCleanDocument(doc As Document)
    Dim rng As Range
    Dim storyIndex As Variant
    Dim commentObj As Comment  ' 新增：用于遍历批注对象
    Dim shp As Shape
    Dim frm As Frame
    
    ' 仅遍历核心文本区域
    Dim targetStories As Variant
    targetStories = Array(1, 2, 3, 6, 7, 8, 9, 10, 11)
    
    With doc
        ' 1. 解除文档保护
        If .ProtectionType <> wdNoProtection Then
            On Error Resume Next
            .Unprotect
            If Err.Number <> 0 Then
                Debug.Print "无法解除保护: " & .Name
                Err.Clear
            End If
            On Error GoTo 0
        End If

        ' 2. 接受所有修订
        On Error Resume Next
        .Revisions.AcceptAll
        On Error GoTo 0

        ' 3. 清理核心Story区域
        For Each storyIndex In targetStories
            On Error Resume Next
            Set rng = .StoryRanges(storyIndex)
            On Error GoTo 0
            
            If Not rng Is Nothing Then
                Do While Not rng Is Nothing
                    Call CleanSingleRange(rng)
                    Set rng = rng.NextStoryRange
                Loop
            End If
        Next storyIndex
        
        ' 4. 直接遍历Comments集合删除批注
        ' 这是唯一能确保所有批注被删除的方法
        On Error Resume Next
        If .Comments.count > 0 Then
            ' 方法A：逐个删除（最可靠）
            For Each commentObj In .Comments
                commentObj.Delete
            Next commentObj
            
            ' 方法B：一次性删除（备选）
            ' .DeleteAllComments
        End If
        On Error GoTo 0
        
        ' 5. 结束设置
        .TrackRevisions = False
    End With
End Sub

' === 模块：清理单个 Range ===
Sub CleanSingleRange(rng As Range)
    Dim shp As Shape
    Dim frm As Frame
    
    On Error Resume Next
    
    ' A. 解锁域
    If rng.Fields.count > 0 Then rng.Fields.Locked = False
    
    ' B. 接受修订
    rng.Revisions.AcceptAll
    
    ' C. 清理 Shapes
    If rng.ShapeRange.count > 0 Then
        For Each shp In rng.ShapeRange
            Call ProcessShapeRecursively(shp)
        Next
    End If
    
    ' D. 清理 Frames
    If rng.Frames.count > 0 Then
        For Each frm In rng.Frames
            If frm.Range.Fields.count > 0 Then frm.Range.Fields.Locked = False
            frm.Range.Revisions.AcceptAll
        Next
    End If
    
    On Error GoTo 0
End Sub

' === 模块：递归处理图形对象 ===
Sub ProcessShapeRecursively(shp As Shape)
    Dim subShp As Shape
    
    On Error Resume Next
    
    ' 1. 文本框内容
    If Not shp.TextFrame Is Nothing Then
        If shp.TextFrame.HasText Then
            If shp.TextFrame.TextRange.Fields.count > 0 Then _
                shp.TextFrame.TextRange.Fields.Locked = False
            shp.TextFrame.TextRange.Revisions.AcceptAll
        End If
    End If
    
    ' 2. 组合图形
    If shp.Type = msoGroup Then
        For Each subShp In shp.GroupItems
            Call ProcessShapeRecursively(subShp)
        Next
    End If
    
    ' 3. 画布
    If shp.Type = msoCanvas Then
        For Each subShp In shp.CanvasItems
            Call ProcessShapeRecursively(subShp)
        Next
    End If
    
    On Error GoTo 0
End Sub

' === 模块：文件递归搜索 ===
Sub RecursiveFindFiles(ByVal sPath As String, ByRef fCollection As Collection)
    Dim FSO As Object, Folder As Object, SubFolder As Object, File As Object
    Dim ext As String
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set Folder = FSO.GetFolder(sPath)
    If Err.Number <> 0 Then Exit Sub
    
    ' 遍历文件
    For Each File In Folder.Files
        ext = LCase(FSO.GetExtensionName(File.Name))
        If (ext = "doc" Or ext = "docx" Or ext = "docm") Then
            If Left(File.Name, 2) <> "~$" Then fCollection.Add File.Path
        End If
    Next
    
    ' 递归子文件夹
    For Each SubFolder In Folder.SubFolders
        RecursiveFindFiles SubFolder.Path, fCollection
    Next
    Set FSO = Nothing
End Sub

