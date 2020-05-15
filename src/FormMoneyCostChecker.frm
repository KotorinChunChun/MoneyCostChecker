VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormMoneyCostChecker 
   Caption         =   "UserForm1"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   OleObjectBlob   =   "FormMoneyCostChecker.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "FormMoneyCostChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        FormMoneyCostChecker
Rem
Rem  @description   Time is money.
Rem                 会議等にかけているコストをカウントするツール
Rem
Rem  @update        2020/05/15
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Rem
Rem  @note
Rem     問題点：編集機能、保存機能、拡大バグ、計測時間が適当すぎ
Rem
Rem --------------------------------------------------------------------------------

Option Explicit

Const WIDTH_DEFAULT = 400
Const HEIGHT_MIN = 160
Const HEIGHT_MAX = 320

Const SPIN_DEFAULT_VALUE = 50

Private IsClosing As Boolean
Private TotalCost_ As Double

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Sub Update予定人件費()
    Dim 会議予定時間 As Date
    会議予定時間 = 終了日時 - 開始日時
    ToggleButton_StartEnd.Enabled = (会議予定時間 > 0)
    
    Dim 予定人件費 As Long
    予定人件費 = 秒単価 * (会議予定時間 * 24 * 60 * 60)
    
    Label_PlanPrice.Caption = _
        "　会議予定時間：" & Format(会議予定時間, "hh:mm") & _
        "　予定人件費：" & Format(予定人件費, "0,000") & "円"
End Sub

Private Sub TextBox_EndDate_Change(): Call Update予定人件費: End Sub
Private Sub TextBox_StartDate_Change(): Call Update予定人件費: End Sub
Private Sub TextBox_EndTime_Change(): Call Update予定人件費: End Sub
Private Sub TextBox_StartTime_Change(): Call Update予定人件費: End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call ListItemEdit
End Sub

Private Function Get役職単価(ByRef 役職, ByRef 単価) As Boolean
    役職 = InputBox("役職", "", 役職)
    If 役職 = "" Then Exit Function
    
    単価 = InputBox("人件費", "", 単価)
    If 単価 = "" Then Exit Function
    If Not IsNumStr(単価) Then Exit Function
    
    Get役職単価 = True
End Function

Private Sub ListItemEdit()
    Dim colItems As Object 'Dictionary
    Set colItems = ListBox_GetSelectedItemsDictionary(ListBox1)
    
    Dim rowIndex
    For Each rowIndex In colItems.Keys
        Dim 役職: 役職 = colItems(rowIndex)(0)
        Dim 単価: 単価 = colItems(rowIndex)(1)
        If Get役職単価(役職, 単価) Then
            ListBox1.List(rowIndex, 0) = 役職
            ListBox1.List(rowIndex, 1) = 単価
        End If
    Next
End Sub

Private Sub CommandButton_Add_Click()
    Dim 役職: 役職 = "ほにゃららさん"
    Dim 単価: 単価 = "1000"
    Dim 人数: 人数 = "0"
    If Not Get役職単価(役職, 単価) Then Exit Sub
    
    Dim item: item = Array(役職, 単価, 人数)
    Call ListBox_AddItem(ListBox1, item)
End Sub

Private Sub ListItemCountUp(add As Long)
    Dim colItems As Object 'Dictionary
    Set colItems = ListBox_GetSelectedItemsDictionary(ListBox1)
    
    Dim rowIndex
    For Each rowIndex In colItems.Keys
        ListBox1.List(rowIndex, 2) = WorksheetFunction.Max(0, colItems(rowIndex)(2) + add)
    Next
    
    Call Update予定人件費
End Sub

Private Sub CommandButton_Minus_Click(): Call ListItemCountUp(-1): End Sub
Private Sub CommandButton_Plus_Click(): Call ListItemCountUp(1): End Sub

Private Sub SpinButton_MemberCount_Change()
    Call ListItemCountUp(SpinButton_MemberCount.Value - SPIN_DEFAULT_VALUE)
    SpinButton_MemberCount.Value = SPIN_DEFAULT_VALUE
End Sub

Private Property Get TotalCost() As Double: TotalCost = TotalCost_: End Property
Private Property Let TotalCost(set_cost As Double)
    TotalCost_ = set_cost
    CommandButton_Reset.Visible = (TotalCost_ > 0)
    Label_人件費.Caption = Format(TotalCost_, "#,##0") & "円"
End Property
Private Sub TotalCost_Add(add_cost As Double)
    TotalCost = TotalCost + add_cost
End Sub

Private Sub CommandButton_Reset_Click()
    TotalCost = 0
End Sub

Private Property Get 秒単価() As Double
    Dim ret As Double
    Dim data: data = ListBox1.List
    Dim i As Long
    For i = LBound(data) To UBound(data)
        ret = ret + (data(i, 1) / 60 / 60 * data(i, 2))
    Next
    秒単価 = ret
End Property

Private Sub ToggleButton_StartEnd_Click()
    ConfigCtrlEnabled = False
    Dim 会議予定時間 As Date
    会議予定時間 = 終了日時 - 開始日時
    
    Dim 前景色初期値 As Long: 前景色初期値 = Label_人件費.ForeColor
    Dim 背景色初期値 As Long: 背景色初期値 = Label_人件費.BackColor
    
    '開始待ち
    Do
        If 開始日時 < Now() Then Exit Do
        
        Dim 開始まで残り As Date
        開始まで残り = 開始日時 - Now()

        Label_メッセージ.Caption = "会議開始まで残り : " & 開始まで残り
        
        Call Sleep(100)
        DoEvents
    Loop
    
    '開始後
    Do
        If Not ToggleButton_StartEnd.Value Then Exit Do
        If IsClosing Then Exit Do
        
        Dim 経過時間 As Date
        経過時間 = Now() - 開始日時
'        経過時間 = CDate("1:00:00")

        Dim 残り時間 As Date
        残り時間 = 終了日時 - Now()
        
        Label_メッセージ.Caption = "経過時間 :  " & Format(経過時間, "hh:mm:ss") & _
                                    "　毎秒 : " & Format(秒単価, "0.00") & "円" & _
                                    "　会議終了まで残り : " & 残り時間
        
        '10分を割ったら色を切り替える例
'        Label_人件費.ForeColor = IIf(残り時間 < (1 / 24 / 60 * 10), vbRed, vbBlack)
        '進捗率に応じて色をフェードインさせる例
        Dim 進捗率 As Long
        進捗率 = CLng(CDbl(経過時間) / CDbl(会議予定時間) * 100)
        If 進捗率 <= 100 Then
            Label_人件費.ForeColor = GetFadeColor(前景色初期値, vbRed, 進捗率)
        Else
            If Label_人件費.BackColor = 背景色初期値 Then
                Label_人件費.ForeColor = vbWhite
                Label_人件費.BackColor = vbRed
            Else
                Label_人件費.ForeColor = vbRed
                Label_人件費.BackColor = 背景色初期値
            End If
        End If
        
        Call TotalCost_Add(秒単価 / 5)
        Call Sleep(200)
        DoEvents
    Loop
    Label_人件費.ForeColor = 前景色初期値
    Label_人件費.BackColor = 背景色初期値
    ConfigCtrlEnabled = True
End Sub

Private Property Let ConfigCtrlEnabled(set_enabled As Boolean)
    
    CommandButton_Reset.Enabled = set_enabled

    TextBox_EndDate.Enabled = set_enabled
    TextBox_EndTime.Enabled = set_enabled

    TextBox_StartDate.Enabled = set_enabled
    TextBox_StartTime.Enabled = set_enabled
    
    ListBox1.Enabled = set_enabled
    CommandButton_Add.Enabled = set_enabled
    SpinButton_MemberCount.Enabled = set_enabled
    
End Property

Private Property Get 開始日時() As Date
    Static Last開始日時
    On Error Resume Next
    開始日時 = CDate(TextBox_StartDate.Text & " " & TextBox_StartTime.Text & ":00")
    On Error GoTo 0
    If 開始日時 = 0 Then
        開始日時 = Last開始日時
        TextBox_StartDate.BackColor = vbRed
        TextBox_StartTime.BackColor = vbRed
    Else
        Last開始日時 = 開始日時
        TextBox_StartDate.BackColor = vbWhite
        TextBox_StartTime.BackColor = vbWhite
    End If
End Property
Private Property Let 開始日(dt As Date)
    TextBox_StartDate.Text = Format(dt, "yyyy/mm/dd")
End Property
Private Property Let 開始時間(dt As Date)
    TextBox_StartTime.Text = Format(dt, "hh:mm")
End Property

Private Property Get 終了日時() As Date
    Static Last終了日時
    On Error Resume Next
    TextBox_EndTime.Text = StrConv(TextBox_EndTime.Text, vbNarrow)
    終了日時 = CDate(TextBox_EndDate.Text & " " & TextBox_EndTime.Text & ":00")
    On Error GoTo 0
    If 終了日時 = 0 Then
        終了日時 = Last終了日時
        TextBox_EndDate.BackColor = vbRed
        TextBox_EndTime.BackColor = vbRed
    Else
        Last終了日時 = 終了日時
        TextBox_EndDate.BackColor = vbWhite
        TextBox_EndTime.BackColor = vbWhite
    End If
End Property
Private Property Let 終了日(dt As Date)
    TextBox_EndDate.Text = Format(dt, "yyyy/mm/dd")
End Property
Private Property Let 終了時間(dt As Date)
    TextBox_EndTime.Text = Format(dt, "hh:mm")
End Property

Private Sub ToggleButton_最前面_Click()
    UserForm_TopMost Me, ToggleButton_最前面.Value
End Sub

Private Sub ToggleButton_設定表示_Click(): UserForm_画面サイズ設定: End Sub
Private Sub SpinButton_Zoom_Change(): Call UserForm_画面サイズ設定: End Sub
Private Property Get 画面倍率() As Double
    画面倍率 = SpinButton_Zoom.Value / SPIN_DEFAULT_VALUE
End Property
Private Sub UserForm_画面サイズ設定()
    Me.Width = WIDTH_DEFAULT * 画面倍率
    Me.Height = IIf(ToggleButton_設定表示.Value, HEIGHT_MAX * 画面倍率, HEIGHT_MIN * 画面倍率)
    Me.Zoom = 100 * 画面倍率
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "会議をダラダラ続けるのを止めさせよう！"
    Label_メッセージ.Caption = ""

    CommandButton_Reset.Visible = False
    ToggleButton_最前面.Value = False
    ToggleButton_設定表示.Value = True
    With ListBox1
        .ColumnCount = 3
        .ColumnWidths = "50;30;30"
    End With
    
    ListBox_AddItem ListBox1, Array("部長", "6000", "0")
    ListBox_AddItem ListBox1, Array("課長", "5000", "0")
    ListBox_AddItem ListBox1, Array("係長", "4000", "0")
    ListBox_AddItem ListBox1, Array("社員", "3000", "0")
    
    SpinButton_MemberCount.Value = SPIN_DEFAULT_VALUE
    SpinButton_Zoom.Value = SPIN_DEFAULT_VALUE
    
    開始日 = Now()
    終了日 = DateAdd("h", 1, Now())
    開始時間 = Now()
    終了時間 = DateAdd("h", 1, Now())
    
#If DEF_EXCEL Then
    Excel.Application.Visible = False
#ElseIf DEF_WORD Then
    Word.Application.Visible = False
#End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    IsClosing = True
#If DEF_EXCEL Then
    Excel.Application.Visible = True
#ElseIf DEF_WORD Then
    Word.Application.Visible = True
#End If
End Sub


