VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMacroList 
   Caption         =   "UserForm1"
   ClientHeight    =   5436
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7800
   OleObjectBlob   =   "frmMacroList.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "frmMacroList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 模块级变量，用于缓存所有宏数据，供搜索筛选使用
Private m_AllMacros As Variant

Private Sub UserForm_Initialize()
    ' --- 1. 界面美化 ---
    Me.Caption = "宏管理器"
    
    With Me.lstMacros.Font
        .Name = "微软雅黑"
        .Size = 10
    End With
    
    With Me.lblDetail.Font
        .Name = "微软雅黑"
        .Size = 9
    End With
    Me.lblDetail.ForeColor = RGB(80, 80, 80)
    
    ' --- 2. 关键设置：3列模式 ---
    ' Col 0: 英文代码 (隐藏)
    ' Col 1: 中文名称 (显示)
    ' Col 2: 详细描述 (隐藏)
    Me.lstMacros.ColumnCount = 3
    Me.lstMacros.ColumnWidths = "0;150;0"
    
    ' 初始化时加载数据
    LoadDataToMemory
End Sub

' 输入框变化事件：实现实时搜索
Private Sub txtSearch_Change()
    Dim sKeyword As String
    sKeyword = Trim(Me.txtSearch.Text)
    RefreshList sKeyword
End Sub

' 选中列表项时触发
Private Sub lstMacros_Change()
    If Me.lstMacros.ListIndex <> -1 Then
        Dim engName As String
        Dim cnDesc As String
        
        ' List(行, 列) -> 列索引从0开始
        engName = Me.lstMacros.List(Me.lstMacros.ListIndex, 0) '第1列：代码
        cnDesc = Me.lstMacros.List(Me.lstMacros.ListIndex, 2)  '第3列：描述
        
        ' --- 修改点：将 .Caption 改回 .Value ---
        Me.lblDetail.Value = "【功能说明】" & vbCrLf & cnDesc & vbCrLf & vbCrLf & _
                             "【宏名称】" & engName
    Else
        ' --- 修改点：将 .Caption 改回 .Value ---
        Me.lblDetail.Value = ""
    End If
End Sub

Private Sub lstMacros_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btnRun_Click
End Sub

' 第一步：将数据加载到内存变量 m_AllMacros，只执行一次
Private Sub LoadDataToMemory()
    On Error Resume Next
    m_AllMacros = Application.Run("GetMyMacroRegistry")
    On Error GoTo 0
    
    ' 加载完数据后，进行初次显示（无关键词）
    RefreshList ""
End Sub

' 第二步：根据关键词刷新列表
Private Sub RefreshList(keyword As String)
    Dim i As Integer
    Dim bMatch As Boolean
    Dim sEng As String, sCn As String, sDesc As String
    
    Me.lstMacros.Clear
    
    ' 如果数据源为空，直接退出
    If IsEmpty(m_AllMacros) Then Exit Sub
    
    For i = LBound(m_AllMacros) To UBound(m_AllMacros)
        ' 获取原始数据
        sEng = m_AllMacros(i)(0)
        sCn = m_AllMacros(i)(1)
        sDesc = m_AllMacros(i)(2)
        
        ' 判断是否匹配：如果关键词为空，或 英文/中文/描述 中包含关键词
        ' vbTextCompare 表示不区分大小写
        If keyword = "" Then
            bMatch = True
        ElseIf InStr(1, sEng, keyword, vbTextCompare) > 0 Or _
               InStr(1, sCn, keyword, vbTextCompare) > 0 Or _
               InStr(1, sDesc, keyword, vbTextCompare) > 0 Then
            bMatch = True
        Else
            bMatch = False
        End If
        
        ' 如果匹配，则添加到列表
        If bMatch Then
            Me.lstMacros.AddItem
            Me.lstMacros.List(Me.lstMacros.ListCount - 1, 0) = sEng
            Me.lstMacros.List(Me.lstMacros.ListCount - 1, 1) = sCn
            Me.lstMacros.List(Me.lstMacros.ListCount - 1, 2) = sDesc
        End If
    Next i
    
    ' 默认选中第一项（如果有数据）
    If Me.lstMacros.ListCount > 0 Then Me.lstMacros.Selected(0) = True
End Sub

Private Sub btnRun_Click()
    ' 1. 检查是否选择了功能
    If Me.lstMacros.ListIndex = -1 Then
        MsgBox "请先选择一个功能！", vbExclamation
        Exit Sub
    End If
    
    Dim macroName As String
    ' 读取第1列(隐藏的英文名)
    macroName = Me.lstMacros.List(Me.lstMacros.ListIndex, 0)
       
    ' 2. 容错运行
    On Error GoTo ErrH
    
    ' 核心只有这一句：运行宏
    Application.Run macroName
    
    Exit Sub

ErrH:
    MsgBox "运行出错：" & Err.Description, vbCritical
    ' 出错也不用 Me.Show 了，因为窗体本来就没关
End Sub

Private Sub btnRefresh_Click()
    ' 刷新按钮现在重新从注册表获取数据
    LoadDataToMemory
    ' 清空搜索框
    Me.txtSearch.Text = ""
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

