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

' ==============================================================================
' Windows API 声明 (实现窗体拖拽调整大小)
' ==============================================================================
#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If

Private Const GWL_STYLE As Long = -16
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_MINIMIZEBOX As Long = &H20000

' ==============================================================================
' 变量声明
' ==============================================================================
Private m_AllMacros As Variant
Private m_IsInitialized As Boolean

' --- 初始尺寸记录 ---
' 用于计算增量 (Delta)
Private m_InitFormW As Double
Private m_InitFormH As Double

' 控件初始状态缓存
Private m_Init_txtSearch_W As Double

Private m_Init_lstMacros_W As Double
Private m_Init_lstMacros_H As Double

Private m_Init_lblDetail_L As Double
Private m_Init_lblDetail_W As Double
Private m_Init_lblDetail_H As Double

Private m_Init_btnRun_T As Double

Private m_Init_btnClose_L As Double
Private m_Init_btnClose_T As Double


' ==============================================================================
' 初始化与布局计算
' ==============================================================================
Private Sub UserForm_Initialize()
    ' 1. 设置 API 样式
    Me.Caption = "宏管理器"
    Call MakeFormResizable
    
    ' 2. 记录初始状态 (Baseline)
    m_InitFormW = Me.InsideWidth
    m_InitFormH = Me.InsideHeight
    
    ' Label_1 不需要记录，因为它是完全静态的
    
    ' txtSearch: 记录初始宽
    m_Init_txtSearch_W = Me.txtSearch.Width
    
    ' lstMacros: 记录初始宽、高
    m_Init_lstMacros_W = Me.lstMacros.Width
    m_Init_lstMacros_H = Me.lstMacros.Height
    
    ' lblDetail: 记录初始左、宽、高 (注意：Top是固定的，不记录)
    m_Init_lblDetail_L = Me.lblDetail.Left
    m_Init_lblDetail_W = Me.lblDetail.Width
    m_Init_lblDetail_H = Me.lblDetail.Height
    
    ' btnRun: 记录初始Top (Left是固定的)
    m_Init_btnRun_T = Me.btnRun.Top
    
    ' btnClose: 记录初始Left, Top
    m_Init_btnClose_L = Me.btnClose.Left
    m_Init_btnClose_T = Me.btnClose.Top
    
    m_IsInitialized = True
    
    ' 3. 界面美化
    With Me.lstMacros.Font
        .Name = "微软雅黑"
        .Size = 10
    End With
    With Me.lblDetail.Font
        .Name = "微软雅黑"
        .Size = 9
    End With
    Me.lblDetail.ForeColor = RGB(80, 80, 80)
    
    ' 4. 列表设置
    Me.lstMacros.ColumnCount = 3
    Call ResizeColumnWidths ' 初始化列宽
    
    ' 5. 加载数据
    LoadDataToMemory
End Sub

Private Sub MakeFormResizable()
    Dim hWnd As LongPtr
    Dim iStyle As Long
    hWnd = FindWindow("ThunderDFrame", Me.Caption)
    If hWnd = 0 Then Exit Sub
    iStyle = GetWindowLong(hWnd, GWL_STYLE)
    iStyle = iStyle Or WS_THICKFRAME Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX
    SetWindowLong hWnd, GWL_STYLE, iStyle
End Sub

' ==============================================================================
' 窗体调整大小逻辑
' ==============================================================================
Private Sub UserForm_Resize()
    If Not m_IsInitialized Then Exit Sub
    On Error Resume Next
    
    Dim currW As Double, currH As Double
    Dim dW As Double, dH As Double
    
    currW = Me.InsideWidth
    currH = Me.InsideHeight
    
    ' 避免缩得太小
    If currW < 200 Then currW = 200
    If currH < 200 Then currH = 200
    
    ' 计算增量 (当前尺寸 - 初始尺寸)
    dW = currW - m_InitFormW
    dH = currH - m_InitFormH
    
    ' --- 1. Label_1 ---
    ' 左边距上边距保持不变，高和宽也不随拖拽变化 -> 不做任何操作
    
    ' --- 2. txtSearch ---
    ' 左边距上边距保持不变 (默认不改Left/Top)
    ' 高不随拖拽变化 (默认不改Height)
    ' 宽随拖拽变化 -> 增加 dW
    Me.txtSearch.Width = m_Init_txtSearch_W + dW
    
    ' --- 3. lstMacros ---
    ' 左边距上边距保持不变
    ' 高和宽随拖拽变化 -> 宽加 dW, 高加 dH
    Me.lstMacros.Width = m_Init_lstMacros_W + dW
    Me.lstMacros.Height = m_Init_lstMacros_H + dH
    
    ' --- 4. lblDetail ---
    ' 上边距保持不变 (默认不改Top)
    ' 左边距随拖拽变化 -> Left 加 dW (向右移动)
    ' 高随拖拽变化 -> Height 加 dH
    Me.lblDetail.Left = m_Init_lblDetail_L + dW
    Me.lblDetail.Height = m_Init_lblDetail_H + dH
    
    ' --- 5. btnRun ---
    ' 左边距保持不变
    ' 上边距随拖拽变化 -> Top 加 dH (向下移动)
    ' 高和宽不随拖拽变化
    Me.btnRun.Top = m_Init_btnRun_T + dH
    
    ' --- 6. btnClose ---
    ' 左边距随拖拽变化 -> Left 加 dW (向右移动)
    ' 上边距随拖拽变化 -> Top 加 dH (向下移动)
    ' 高和宽不随拖拽变化
    Me.btnClose.Left = m_Init_btnClose_L + dW
    Me.btnClose.Top = m_Init_btnClose_T + dH
    
    ' 刷新列宽
    Call ResizeColumnWidths
    
    On Error GoTo 0
End Sub

Private Sub ResizeColumnWidths()
    On Error Resume Next
    ' 让第2列(中文名)填满，留出滚动条空间
    Dim w2 As Double
    w2 = Me.lstMacros.Width - 20 ' 减去大概的滚动条宽度
    If w2 < 0 Then w2 = 0
    Me.lstMacros.ColumnWidths = "0;" & w2 & ";0"
End Sub

' ==============================================================================
' 功能逻辑
' ==============================================================================
Private Sub txtSearch_Change()
    RefreshList Trim(Me.txtSearch.Text)
End Sub

Private Sub lstMacros_Change()
    If Me.lstMacros.ListIndex <> -1 Then
        Dim engName As String, cnDesc As String
        engName = Me.lstMacros.List(Me.lstMacros.ListIndex, 0)
        cnDesc = Me.lstMacros.List(Me.lstMacros.ListIndex, 2)
        Me.lblDetail.Value = "【功能说明】" & vbCrLf & cnDesc & vbCrLf & vbCrLf & "【宏名称】" & engName
    Else
        Me.lblDetail.Value = ""
    End If
End Sub

Private Sub lstMacros_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btnRun_Click
End Sub

Private Sub LoadDataToMemory()
    On Error Resume Next
    m_AllMacros = Application.Run("GetMyMacroRegistry")
    On Error GoTo 0
    RefreshList ""
End Sub

Private Sub RefreshList(keyword As String)
    Dim i As Integer, bMatch As Boolean
    Dim sEng As String, sCn As String, sDesc As String
    
    Me.lstMacros.Clear
    If IsEmpty(m_AllMacros) Then Exit Sub
    
    For i = LBound(m_AllMacros) To UBound(m_AllMacros)
        sEng = m_AllMacros(i)(0)
        sCn = m_AllMacros(i)(1)
        sDesc = m_AllMacros(i)(2)
        
        If keyword = "" Then
            bMatch = True
        ElseIf InStr(1, sEng, keyword, vbTextCompare) > 0 Or _
               InStr(1, sCn, keyword, vbTextCompare) > 0 Or _
               InStr(1, sDesc, keyword, vbTextCompare) > 0 Then
            bMatch = True
        Else
            bMatch = False
        End If
        
        If bMatch Then
            Me.lstMacros.AddItem
            Me.lstMacros.List(Me.lstMacros.ListCount - 1, 0) = sEng
            Me.lstMacros.List(Me.lstMacros.ListCount - 1, 1) = sCn
            Me.lstMacros.List(Me.lstMacros.ListCount - 1, 2) = sDesc
        End If
    Next i
    
    If Me.lstMacros.ListCount > 0 Then Me.lstMacros.Selected(0) = True
End Sub

Private Sub btnRun_Click()
    If Me.lstMacros.ListIndex = -1 Then
        MsgBox "请先选择一个功能！", vbExclamation
        Exit Sub
    End If
    On Error GoTo ErrH
    Application.Run Me.lstMacros.List(Me.lstMacros.ListIndex, 0)
    Exit Sub
ErrH:
    MsgBox "运行出错：" & Err.Description, vbCritical
End Sub

Private Sub btnRefresh_Click()
    LoadDataToMemory
    Me.txtSearch.Text = ""
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

